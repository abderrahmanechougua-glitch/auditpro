"""
REST API for AuditPro's Ollama agent.

Run:
  uvicorn agent.api_server:app --host 127.0.0.1 --port 8008 --reload
"""
from __future__ import annotations

from pathlib import Path
from threading import Lock

from fastapi import FastAPI, HTTPException
from pydantic import BaseModel, Field

from core.module_registry import ModuleRegistry
from agent.llama_agent import LlamaAgent
from agent.tools import execute_tool


class ChatRequest(BaseModel):
    message: str = Field(..., description="User message for the agent")
    model: str | None = Field(default=None, description="Optional Ollama model override")


class ToolRunRequest(BaseModel):
    module_name: str = Field(..., description="Exact AuditPro module name")
    arguments: dict = Field(default_factory=dict, description="Tool/module arguments")


app = FastAPI(title="AuditPro Agent API", version="1.0.0")

_registry_lock = Lock()
_agent_lock = Lock()
_registry = ModuleRegistry()
_agent = LlamaAgent(_registry)


@app.get("/health")
def health() -> dict:
    return {
        "ok": True,
        "ollama_client_installed": _agent.available,
        "modules_count": _registry.count(),
        "model": _agent.model,
    }


@app.get("/models")
def models() -> dict:
    return {"models": _agent.list_models()}


@app.get("/modules")
def modules() -> dict:
    with _registry_lock:
        data = []
        for name, module in _registry.get_all().items():
            data.append(
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
    return {"modules": data}


@app.post("/tool/run")
def run_tool(req: ToolRunRequest) -> dict:
    with _registry_lock:
        result = execute_tool(_registry, req.module_name, dict(req.arguments))
    return {"result": result}


@app.post("/chat")
def chat(req: ChatRequest) -> dict:
    if not _agent.available:
        raise HTTPException(status_code=503, detail="Python package 'ollama' is not installed")

    with _agent_lock:
        if req.model:
            _agent.model = req.model
        try:
            answer = _agent.chat(req.message)
        except Exception as exc:
            raise HTTPException(status_code=500, detail=str(exc)) from exc

    return {
        "model": _agent.model,
        "answer": answer,
    }


@app.post("/chat/reset")
def reset_chat() -> dict:
    with _agent_lock:
        _agent.reset()
    return {"ok": True}
