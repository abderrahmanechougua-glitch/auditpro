"""
Zero-dependency REST API for AuditPro's Ollama agent.

Use this server when FastAPI/Uvicorn are unavailable.

Run:
  python -m agent.api_server_stdlib --host 127.0.0.1 --port 8008
"""
from __future__ import annotations

import argparse
import importlib.util
import json
import os
import sys
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from threading import Lock, Thread
import time
from urllib.parse import urlparse


APP_ROOT = Path(__file__).resolve().parent.parent


def _bootstrap_app_path() -> None:
    """Force imports to resolve against AuditPro_SHARE, not sibling workspaces."""
    app_root_str = str(APP_ROOT)
    if app_root_str not in sys.path:
        sys.path.insert(0, app_root_str)
    os.chdir(APP_ROOT)


def _ensure_base_site_packages() -> None:
    """Allow venv runtime to reuse base interpreter packages when offline."""
    if importlib.util.find_spec("pandas"):
        return

    base_site = Path(sys.base_prefix) / "Lib" / "site-packages"
    if base_site.exists():
        base_site_str = str(base_site)
        if base_site_str not in sys.path:
            sys.path.append(base_site_str)


_bootstrap_app_path()
_ensure_base_site_packages()


from core.module_registry import ModuleRegistry
from core.debug_utils import build_diagnostic
from agent.llama_agent import LlamaAgent
from agent.tools import execute_tool


_registry_lock = Lock()
_agent_lock = Lock()
_registry: ModuleRegistry | None = None
_agent: LlamaAgent | None = None
_registry_init_error: str | None = None
_registry_loading = False


def _registry_worker() -> None:
    global _registry, _registry_init_error, _registry_loading
    try:
        reg = ModuleRegistry()
        with _registry_lock:
            _registry = reg
            _registry_init_error = None
    except Exception as exc:
        with _registry_lock:
            _registry_init_error = str(exc)
    finally:
        with _registry_lock:
            _registry_loading = False


def _ensure_registry_loading() -> None:
    global _registry_loading
    with _registry_lock:
        if _registry is not None or _registry_loading:
            return
        _registry_loading = True
    Thread(target=_registry_worker, daemon=True).start()


def _get_registry(wait_seconds: float = 0.0) -> ModuleRegistry | None:
    _ensure_registry_loading()
    if wait_seconds > 0:
        deadline = time.time() + wait_seconds
        while time.time() < deadline:
            with _registry_lock:
                if _registry is not None:
                    return _registry
                if not _registry_loading:
                    break
            time.sleep(0.1)
    with _registry_lock:
        return _registry


def _get_agent() -> LlamaAgent:
    global _agent
    with _agent_lock:
        if _agent is not None:
            return _agent
        reg = _get_registry(wait_seconds=1.0)
        if reg is None:
            raise RuntimeError("Les modules AuditPro sont en cours de chargement, réessayez dans quelques secondes.")
        _agent = LlamaAgent(reg)
        return _agent


class ApiHandler(BaseHTTPRequestHandler):
    server_version = "AuditProStdlibAPI/1.0"

    def _send_json(self, status: int, payload: dict) -> None:
        body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Headers", "Content-Type")
        self.send_header("Access-Control-Allow-Methods", "GET, POST, OPTIONS")
        self.end_headers()
        self.wfile.write(body)

    def _send_exception(self, status: int, exc: Exception, context: dict | None = None) -> None:
        diagnostic = build_diagnostic(exc, context)
        self._send_json(status, {"error": str(exc), "diagnostic": diagnostic})

    def _read_json_body(self) -> dict:
        length = int(self.headers.get("Content-Length", "0"))
        if length <= 0:
            return {}
        raw = self.rfile.read(length)
        if not raw:
            return {}
        data = json.loads(raw.decode("utf-8"))
        if not isinstance(data, dict):
            raise ValueError("Body must be a JSON object")
        return data

    def do_OPTIONS(self) -> None:
        self._send_json(200, {"ok": True})

    def do_GET(self) -> None:
        path = urlparse(self.path).path

        if path == "/health":
            model = "llama3.2:3b"
            ollama_available = False
            modules_count = None
            _ensure_registry_loading()
            registry_loaded = _registry is not None
            registry_loading = _registry_loading

            if _agent is not None:
                model = _agent.model
                ollama_available = _agent.available

            if _registry is not None:
                modules_count = _registry.count()

            self._send_json(
                200,
                {
                    "ok": True,
                    "ollama_client_installed": ollama_available,
                    "modules_count": modules_count,
                    "registry_loaded": registry_loaded,
                    "registry_loading": registry_loading,
                    "registry_init_error": _registry_init_error,
                    "model": model,
                },
            )
            return

        if path == "/models":
            try:
                agent = _get_agent()
                self._send_json(200, {"models": agent.list_models()})
            except Exception as exc:
                self._send_exception(500, exc, {"endpoint": "/models"})
            return

        if path == "/modules":
            try:
                registry = _get_registry(wait_seconds=0.2)
                if registry is None:
                    self._send_json(
                        202,
                        {
                            "loading": True,
                            "error": _registry_init_error,
                            "modules": [],
                        },
                    )
                    return
                modules = []
                for name, module in registry.get_all().items():
                    modules.append(
                        {
                            "name": name,
                            "category": getattr(module, "category", "General"),
                            "description": getattr(module, "description", ""),
                            "required_inputs": [
                                {
                                    "key": inp.key,
                                    "label": inp.label,
                                    "type": inp.input_type,
                                    "required": inp.required,
                                    "multiple": inp.multiple,
                                }
                                for inp in module.get_required_inputs()
                            ],
                        }
                    )
                self._send_json(200, {"modules": modules})
            except Exception as exc:
                diagnostic = build_diagnostic(exc, {"endpoint": "/modules"})
                self._send_json(500, {"error": str(exc), "diagnostic": diagnostic, "modules": []})
            return

        self._send_json(404, {"error": "Not found"})

    def do_POST(self) -> None:
        path = urlparse(self.path).path

        try:
            payload = self._read_json_body()
        except json.JSONDecodeError:
            self._send_json(400, {"error": "Invalid JSON body"})
            return
        except ValueError as exc:
            self._send_json(400, {"error": str(exc)})
            return

        if path == "/tool/run":
            module_name = payload.get("module_name")
            arguments = payload.get("arguments", {})
            if not isinstance(module_name, str) or not module_name.strip():
                self._send_json(400, {"error": "Field 'module_name' is required"})
                return
            if not isinstance(arguments, dict):
                self._send_json(400, {"error": "Field 'arguments' must be an object"})
                return

            try:
                registry = _get_registry(wait_seconds=2.0)
                if registry is None:
                    self._send_json(503, {"error": "Les modules AuditPro sont en cours de chargement. Réessayez dans quelques secondes."})
                    return
                result = execute_tool(registry, module_name, dict(arguments))
            except Exception as exc:
                self._send_exception(500, exc, {"endpoint": "/tool/run", "module": module_name})
                return

            self._send_json(200, {"result": result})
            return

        if path == "/chat":
            message = payload.get("message")
            model = payload.get("model")

            if not isinstance(message, str) or not message.strip():
                self._send_json(400, {"error": "Field 'message' is required"})
                return
            if model is not None and not isinstance(model, str):
                self._send_json(400, {"error": "Field 'model' must be a string"})
                return

            try:
                agent = _get_agent()
                if not agent.available:
                    self._send_json(503, {"error": "Ollama n'est pas accessible (démarrez Ollama Desktop ou ollama serve)."})
                    return

                with _agent_lock:
                    if model:
                        agent.model = model
                    answer = agent.chat(message)
            except Exception as exc:
                self._send_exception(500, exc, {"endpoint": "/chat", "model": model or "default"})
                return

            self._send_json(200, {"model": agent.model, "answer": answer})
            return

        if path == "/chat/reset":
            with _agent_lock:
                if _agent is not None:
                    _agent.reset()
            self._send_json(200, {"ok": True})
            return

        self._send_json(404, {"error": "Not found"})

    def log_message(self, fmt: str, *args) -> None:
        # Keep logs minimal while still useful when run in terminal.
        print(f"[{self.log_date_time_string()}] {self.address_string()} - {fmt % args}")


def main() -> None:
    parser = argparse.ArgumentParser(description="AuditPro stdlib REST API")
    parser.add_argument("--host", default="127.0.0.1", help="Bind host")
    parser.add_argument("--port", type=int, default=8008, help="Bind port")
    args = parser.parse_args()

    server = ThreadingHTTPServer((args.host, args.port), ApiHandler)
    print(f"AuditPro stdlib API running on http://{args.host}:{args.port}")
    print("Endpoints: GET /health /models /modules | POST /chat /chat/reset /tool/run")
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        pass
    finally:
        server.server_close()


if __name__ == "__main__":
    main()
