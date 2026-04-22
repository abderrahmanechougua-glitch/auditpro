"""
Smoke tests for AuditPro's stdlib API.

Run after starting the server:
  python -m agent.api_server_stdlib --host 127.0.0.1 --port 8008
  python agent/test_stdlib_api.py --base-url http://127.0.0.1:8008
"""
from __future__ import annotations

import argparse
import json
import sys
import urllib.error
import urllib.request


def http_get_json(base_url: str, path: str) -> tuple[int, dict]:
    req = urllib.request.Request(f"{base_url}{path}", method="GET")
    try:
        with urllib.request.urlopen(req, timeout=15) as response:
            return response.status, json.loads(response.read().decode("utf-8"))
    except urllib.error.HTTPError as exc:
        body = exc.read().decode("utf-8") if exc.fp else "{}"
        return exc.code, json.loads(body or "{}")


def http_post_json(base_url: str, path: str, payload: dict) -> tuple[int, dict]:
    data = json.dumps(payload).encode("utf-8")
    req = urllib.request.Request(
        f"{base_url}{path}",
        data=data,
        headers={"Content-Type": "application/json"},
        method="POST",
    )
    try:
        with urllib.request.urlopen(req, timeout=30) as response:
            return response.status, json.loads(response.read().decode("utf-8"))
    except urllib.error.HTTPError as exc:
        body = exc.read().decode("utf-8") if exc.fp else "{}"
        return exc.code, json.loads(body or "{}")


def print_test(name: str, ok: bool, detail: str) -> None:
    status = "PASS" if ok else "FAIL"
    print(f"[{status}] {name}: {detail}")


def run_tests(base_url: str) -> int:
    failures = 0

    status, payload = http_get_json(base_url, "/health")
    ok = status == 200 and payload.get("ok") is True
    detail = f"status={status}, registry_loaded={payload.get('registry_loaded')}, registry_error={payload.get('registry_init_error')}"
    print_test("GET /health", ok, detail)
    failures += 0 if ok else 1

    status, payload = http_get_json(base_url, "/modules")
    ok = status in (200, 202) and isinstance(payload.get("modules"), list)
    detail = f"status={status}, modules={len(payload.get('modules', []))}, loading={payload.get('loading', False)}"
    print_test("GET /modules", ok, detail)
    failures += 0 if ok else 1

    status, payload = http_get_json(base_url, "/models")
    ok = status in (200, 500)
    if status == 200:
        detail = f"status={status}, models={len(payload.get('models', []))}"
    else:
        diagnostic = payload.get("diagnostic", {})
        detail = f"status={status}, error={payload.get('error')}, diag_type={diagnostic.get('type')}"
    print_test("GET /models", ok, detail)
    failures += 0 if ok else 1

    status, payload = http_post_json(base_url, "/chat", {"message": "bonjour"})
    ok = status in (200, 503, 500)
    if status == 200:
        detail = f"status={status}, model={payload.get('model')}, answer_len={len(payload.get('answer', ''))}"
    else:
        diagnostic = payload.get("diagnostic", {})
        detail = f"status={status}, error={payload.get('error')}, diag_type={diagnostic.get('type')}"
    print_test("POST /chat", ok, detail)
    failures += 0 if ok else 1

    status, payload = http_post_json(base_url, "/tool/run", {"module_name": "Lettrage Grand Livre", "arguments": {}})
    ok = status in (200, 503)
    if status == 200:
        raw_result = payload.get("result", "")
        try:
            result = json.loads(raw_result) if isinstance(raw_result, str) else raw_result
        except Exception:
            result = {"parse_error": True}
        detail = f"status={status}, success={result.get('success')}, errors={result.get('errors') or result.get('error')}"
    else:
        detail = f"status={status}, error={payload.get('error')}"
    print_test("POST /tool/run", ok, detail)
    failures += 0 if ok else 1

    return failures


def main() -> int:
    parser = argparse.ArgumentParser(description="Smoke tests for AuditPro stdlib API")
    parser.add_argument("--base-url", default="http://127.0.0.1:8008", help="Base URL of the stdlib API")
    args = parser.parse_args()

    try:
        failures = run_tests(args.base_url.rstrip("/"))
    except urllib.error.URLError as exc:
        print(f"[FAIL] API unreachable: {exc}")
        return 1

    print(f"\nTotal failures: {failures}")
    return 0 if failures == 0 else 1


if __name__ == "__main__":
    raise SystemExit(main())