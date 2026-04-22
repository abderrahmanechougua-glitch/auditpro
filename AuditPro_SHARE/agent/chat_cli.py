"""
Terminal chat client for AuditPro + Ollama.

Run:
  python -m agent.chat_cli --model llama3.2:3b
"""
from __future__ import annotations

import argparse
import json
import os
import sys
from pathlib import Path


APP_ROOT = Path(__file__).resolve().parent.parent
os.chdir(APP_ROOT)
if str(APP_ROOT) not in sys.path:
    sys.path.insert(0, str(APP_ROOT))

from core.module_registry import ModuleRegistry
from agent.llama_agent import LlamaAgent


def _tool_call_printer(module_name: str, args: dict) -> None:
    print(f"\n[TOOL CALL] {module_name}")
    print(json.dumps(args, ensure_ascii=False, indent=2))


def _tool_result_printer(module_name: str, payload: str) -> None:
    print(f"\n[TOOL RESULT] {module_name}")
    try:
        data = json.loads(payload)
        print(json.dumps(data, ensure_ascii=False, indent=2))
    except Exception:
        print(payload)


def main() -> int:
    parser = argparse.ArgumentParser(description="AuditPro Ollama CLI")
    parser.add_argument("--model", default="llama3.2:3b", help="Ollama model name")
    args = parser.parse_args()

    registry = ModuleRegistry()
    agent = LlamaAgent(registry, model=args.model)

    if not agent.available:
        print("La librairie ollama n'est pas installée. Installez-la avec: pip install ollama")
        return 1

    print("AuditPro AI CLI (Ollama)")
    print(f"Modele: {agent.model}")
    print("Tapez 'exit' pour quitter, '/reset' pour vider l'historique.\n")

    while True:
        try:
            user_text = input("Vous > ").strip()
        except (EOFError, KeyboardInterrupt):
            print("\nFin de session.")
            return 0

        if not user_text:
            continue
        if user_text.lower() in {"exit", "quit"}:
            print("Fin de session.")
            return 0
        if user_text == "/reset":
            agent.reset()
            print("Historique reinitialise.")
            continue

        try:
            reply = agent.chat(
                user_text,
                on_tool_call=_tool_call_printer,
                on_tool_result=_tool_result_printer,
            )
            print(f"\nAgent > {reply}\n")
        except Exception as exc:
            print(f"Erreur: {exc}")


if __name__ == "__main__":
    raise SystemExit(main())
