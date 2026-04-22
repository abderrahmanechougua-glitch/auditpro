"""
AuditPro tool wrappers — exposes each module as a callable tool for the Llama agent.
"""
from __future__ import annotations
import json
from pathlib import Path
from typing import Any

from core.debug_utils import build_diagnostic


def build_tools(registry) -> list[dict]:
    """
    Returns a list of Ollama-compatible tool definitions, one per module.
    """
    tools = []
    for name, module in registry.get_all().items():
        inputs = module.get_required_inputs()
        properties = {}
        required_keys = []

        for inp in inputs:
            if inp.input_type in ("file", "folder"):
                prop = {
                    "type": "string",
                    "description": f"Chemin vers {inp.label}"
                }
                if inp.multiple:
                    prop = {
                        "type": "array",
                        "items": {"type": "string"},
                        "description": f"Liste de chemins pour {inp.label}"
                    }
            elif inp.input_type == "number":
                prop = {"type": "number", "description": inp.label}
            elif inp.input_type == "combo":
                prop = {
                    "type": "string",
                    "description": inp.label,
                    "enum": inp.options if inp.options else []
                }
            elif inp.input_type == "date":
                prop = {
                    "type": "string",
                    "description": f"{inp.label} (format YYYY-MM-DD)"
                }
            else:
                prop = {"type": "string", "description": inp.label}

            if inp.default is not None:
                prop["default"] = inp.default
            if inp.tooltip:
                prop["description"] += f" — {inp.tooltip}"

            properties[inp.key] = prop
            if inp.required:
                required_keys.append(inp.key)

        # Always add output_dir
        properties["output_dir"] = {
            "type": "string",
            "description": "Dossier de sortie pour les fichiers générés (optionnel)"
        }

        safe_name = name.lower().replace(" ", "_").replace("é", "e").replace("è", "e").replace("ê", "e")
        tools.append({
            "type": "function",
            "function": {
                "name": f"auditpro_{safe_name}",
                "description": module.description,
                "parameters": {
                    "type": "object",
                    "properties": properties,
                    "required": required_keys
                }
            },
            "_module_name": name  # internal mapping
        })

    return tools


def get_tool_name_map(tools: list[dict]) -> dict[str, str]:
    """Maps tool function name → module name."""
    return {t["function"]["name"]: t["_module_name"] for t in tools}


def execute_tool(registry, module_name: str, arguments: dict) -> str:
    """
    Runs an AuditPro module and returns a JSON-serialisable result string.
    """
    module = registry.get(module_name)
    if module is None:
        return json.dumps({"success": False, "error": f"Module '{module_name}' introuvable."})

    output_dir = arguments.pop("output_dir", None)
    if not output_dir:
        output_dir = str(Path.home() / "Documents" / "AuditPro_Output")

    ok, errors = module.validate(arguments)
    if not ok:
        return json.dumps({"success": False, "errors": errors})

    try:
        result = module.execute(arguments, output_dir)
        return json.dumps({
            "success": result.success,
            "output_path": result.output_path,
            "message": result.message,
            "stats": result.stats,
            "warnings": result.warnings,
            "errors": result.errors
        }, ensure_ascii=False)
    except Exception as exc:
        diagnostic = build_diagnostic(
            exc,
            {
                "module": module_name,
                "output_dir": output_dir,
                "input_keys": ",".join(sorted(arguments.keys())),
            },
        )
        return json.dumps({"success": False, "error": str(exc), "diagnostic": diagnostic}, ensure_ascii=False)
