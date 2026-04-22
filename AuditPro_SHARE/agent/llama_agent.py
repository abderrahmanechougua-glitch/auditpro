"""
Llama Agent — uses Ollama's OpenAI-compatible API with tool calling.

Requires: pip install ollama
Requires: Ollama running locally with a model that supports tool use
          (llama3.1, llama3.2, mistral-nemo, etc.)
"""
from __future__ import annotations
import json
import os
from pathlib import Path
import sys
import threading
import urllib.error
import urllib.request
from typing import Callable

from core.debug_utils import build_diagnostic

APP_ROOT = Path(__file__).resolve().parent.parent
if str(APP_ROOT) not in sys.path:
    sys.path.insert(0, str(APP_ROOT))
os.chdir(APP_ROOT)

try:
    import ollama
    OLLAMA_AVAILABLE = True
except ImportError:
    OLLAMA_AVAILABLE = False

OLLAMA_BASE_URL = "http://127.0.0.1:11434"
OLLAMA_CHAT_TIMEOUT = 300
OLLAMA_META_TIMEOUT = 15

from agent.tools import build_tools, get_tool_name_map, execute_tool


SYSTEM_PROMPT = """Tu es un assistant expert en audit financier intégré dans AuditPro.
Tu aides les auditeurs marocains à automatiser leurs travaux d'audit.

Tu as accès aux outils suivants pour exécuter des traitements :
- Centralisation TVA : extraire les déclarations TVA depuis des PDFs DGI
- Centralisation CNSS : extraire les cotisations sociales depuis des bordereaux PDF
- Extraction Factures : extraire des données de factures PDF par OCR
- Extraction IR : extraire l'impôt sur le revenu depuis des déclarations PDF
- Lettrage GL : rapprocher automatiquement les écritures comptables
- Retraitement Comptable : appliquer des ajustements comptables sur un GL
- SRM Generator : générer un tableau de synthèse des anomalies (Summary of Recorded Misstatements)
- Circularisation des Tiers : générer des lettres de confirmation pour clients/fournisseurs

RÈGLES :
- Demande toujours les chemins de fichiers manquants avant d'exécuter un outil.
- Confirme avec l'utilisateur avant d'exécuter un outil si les paramètres sont ambigus.
- Réponds en français sauf si l'utilisateur écrit en anglais.
- Après exécution d'un outil, explique les résultats de façon claire et professionnelle.
- Si un outil échoue, explique l'erreur et propose des solutions.
"""


class LlamaAgent:
    """
    Stateful agent that maintains conversation history and calls AuditPro tools.
    """

    def __init__(self, registry, model: str = "llama3.2:3b"):
        self.registry = registry
        self.model = model
        self.tools = build_tools(registry)
        self.tool_name_map = get_tool_name_map(self.tools)
        # Strip internal key before sending to Ollama
        self._ollama_tools = [
            {k: v for k, v in t.items() if k != "_module_name"}
            for t in self.tools
        ]
        self.history: list[dict] = []
        self._lock = threading.Lock()

    def _post_ollama_json(self, endpoint: str, payload: dict) -> dict:
        url = f"{OLLAMA_BASE_URL}{endpoint}"
        data = json.dumps(payload).encode("utf-8")
        req = urllib.request.Request(
            url,
            data=data,
            headers={"Content-Type": "application/json"},
            method="POST",
        )
        with urllib.request.urlopen(req, timeout=OLLAMA_CHAT_TIMEOUT) as response:
            raw = response.read().decode("utf-8")
        return json.loads(raw) if raw else {}

    def _get_ollama_json(self, endpoint: str) -> dict:
        url = f"{OLLAMA_BASE_URL}{endpoint}"
        req = urllib.request.Request(url, method="GET")
        with urllib.request.urlopen(req, timeout=OLLAMA_META_TIMEOUT) as response:
            raw = response.read().decode("utf-8")
        return json.loads(raw) if raw else {}

    @property
    def available(self) -> bool:
        if OLLAMA_AVAILABLE:
            return True
        try:
            self._get_ollama_json("/api/tags")
            return True
        except Exception:
            return False

    def reset(self):
        with self._lock:
            self.history.clear()

    def chat(
        self,
        user_message: str,
        on_token: Callable[[str], None] | None = None,
        on_tool_call: Callable[[str, dict], None] | None = None,
        on_tool_result: Callable[[str, str], None] | None = None,
    ) -> str:
        """
        Send a message and return the final assistant response.

        Callbacks (all optional, called from the same thread):
          on_token(text)                  — streaming token
          on_tool_call(tool_name, args)   — before executing a tool
          on_tool_result(tool_name, json) — after executing a tool
        """
        if not self.available:
            return "⚠️ Ollama n'est pas accessible. Lancez Ollama Desktop (ou `ollama serve`) avant d'utiliser l'agent."

        with self._lock:
            self.history.append({"role": "user", "content": user_message})

            messages = [{"role": "system", "content": SYSTEM_PROMPT}] + self.history

            # Agentic loop — keeps running until no more tool calls
            while True:
                if OLLAMA_AVAILABLE:
                    try:
                        response = ollama.chat(
                            model=self.model,
                            messages=messages,
                            tools=self._ollama_tools,
                            stream=False,
                        )
                    except Exception as exc:
                        diagnostic = build_diagnostic(exc, {"stage": "ollama.chat", "model": self.model})
                        return f"Erreur agent: {diagnostic['type']}: {diagnostic['message']}"
                else:
                    try:
                        response = self._post_ollama_json(
                            "/api/chat",
                            {
                                "model": self.model,
                                "messages": messages,
                                "tools": self._ollama_tools,
                                "stream": False,
                            },
                        )
                    except Exception as exc:
                        diagnostic = build_diagnostic(exc, {"stage": "http_api_chat", "model": self.model})
                        return f"Erreur agent: {diagnostic['type']}: {diagnostic['message']}"

                msg = response["message"]

                # ── Tool calls ────────────────────────────────────
                if msg.get("tool_calls"):
                    messages.append({"role": "assistant", "content": msg.get("content", ""), "tool_calls": msg["tool_calls"]})
                    self.history.append({"role": "assistant", "content": msg.get("content", ""), "tool_calls": msg["tool_calls"]})

                    for tc in msg["tool_calls"]:
                        fn_name = tc["function"]["name"]
                        args = tc["function"].get("arguments", {})
                        if isinstance(args, str):
                            try:
                                args = json.loads(args)
                            except Exception:
                                args = {}

                        module_name = self.tool_name_map.get(fn_name, fn_name)

                        if on_tool_call:
                            on_tool_call(module_name, args)

                        tool_result = execute_tool(self.registry, module_name, dict(args))

                        if on_tool_result:
                            on_tool_result(module_name, tool_result)

                        tool_msg = {
                            "role": "tool",
                            "content": tool_result,
                            "name": fn_name
                        }
                        messages.append(tool_msg)
                        self.history.append(tool_msg)

                    # Continue loop to get final response
                    continue

                # ── Final text response ───────────────────────────
                content = msg.get("content", "")
                self.history.append({"role": "assistant", "content": content})
                return content

    def list_models(self) -> list[str]:
        """Returns locally available Ollama models."""
        if not self.available:
            return []
        try:
            if OLLAMA_AVAILABLE:
                result = ollama.list()
            else:
                result = self._get_ollama_json("/api/tags")
            return [m["name"] for m in result.get("models", [])]
        except Exception:
            return []


    if __name__ == "__main__":
        from agent.chat_cli import main as chat_cli_main

        raise SystemExit(chat_cli_main())
