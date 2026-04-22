"""
AuditPro AI Agent — API Server
Intègre Ollama pour router les requêtes en langage naturel vers les modules AuditPro.
"""
import os
import sys
import logging
from pathlib import Path
from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.staticfiles import StaticFiles
from fastapi.responses import HTMLResponse
from pydantic import BaseModel, Field
import requests
import json
import shutil
import tempfile
from typing import Optional, List, Dict, Any

# ── Configuration ─────────────────────────────────────────
AUDITPRO_DIR = Path(__file__).parent / "AuditPro"
if not AUDITPRO_DIR.exists():
    print(f"[ERROR] AuditPro not found at {AUDITPRO_DIR}")
    print("Make sure the symlink/junction is created: mklink /J AuditPro ..\\AuditPro")
    sys.exit(1)

sys.path.insert(0, str(AUDITPRO_DIR))

# Logging setup
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger("auditpro-agent")

OLLAMA_BASE_URL = os.getenv("OLLAMA_URL", "http://localhost:11434")
OLLAMA_MODEL = os.getenv("OLLAMA_MODEL", "llama3.2")

# CORS — restrict origins via env var; default to localhost only (not open wildcard)
_CORS_ORIGINS = [o.strip() for o in os.getenv("CORS_ORIGINS", "http://localhost:8000,http://127.0.0.1:8000").split(",") if o.strip()]

app = FastAPI(title="AuditPro AI Agent", version="1.0", description="AI-powered audit automation with Ollama")

app.add_middleware(
    CORSMiddleware,
    allow_origins=_CORS_ORIGINS,
    allow_credentials=True,
    allow_methods=["GET", "POST"],
    allow_headers=["Content-Type", "Authorization"],
)

# ── Import AuditPro modules ──────────────────────────────
try:
    from modules.base_module import BaseModule, ModuleResult
    from core.module_registry import ModuleRegistry
    registry = ModuleRegistry()
    logger.info(f"Loaded {len(registry.names())} AuditPro modules: {registry.names()}")
except Exception as e:
    logger.error(f"Failed to load AuditPro modules: {e}")
    registry = None

# ── Modèles Pydantic ─────────────────────────────────────
class ChatRequest(BaseModel):
    message: str = Field(..., description="Message en langage naturel")
    conversation: List[Dict[str, str]] = Field(default_factory=list, description="Historique de conversation")


class ModuleExecuteRequest(BaseModel):
    module_name: str = Field(..., description="Nom du module à exécuter")
    inputs: Dict[str, Any] = Field(default_factory=dict, description="Inputs du module")
    params: Dict[str, Any] = Field(default_factory=dict, description="Paramètres optionnels")


class AIResponse(BaseModel):
    response: str = Field(..., description="Réponse de l'IA")
    module_used: Optional[str] = Field(None, description="Module utilisé si applicable")
    output_path: Optional[str] = Field(None, description="Chemin vers le fichier de sortie")
    data: Optional[Dict[str, Any]] = Field(None, description="Données de résultat/stats")
    error: Optional[str] = Field(None, description="Message d'erreur si échec")


class ModuleInfo(BaseModel):
    name: str
    description: str
    category: str
    inputs: List[Dict[str, Any]]


class HealthResponse(BaseModel):
    status: str
    ollama: bool
    modules_loaded: int
    version: str


# ── Helper Ollama ────────────────────────────────────────
def check_ollama_available() -> bool:
    """Check if Ollama is available and responding."""
    try:
        resp = requests.get(f"{OLLAMA_BASE_URL}/api/version", timeout=5)
        return resp.status_code == 200
    except:
        return False

def call_ollama(messages: list, tools: list | None = None, temperature: float = 0.7) -> dict:
    """Call Ollama API. Raises HTTPException if Ollama is unavailable."""
    payload = {
        "model": OLLAMA_MODEL,
        "messages": messages,
        "stream": False,
        "options": {"temperature": temperature}
    }
    if tools:
        payload["tools"] = tools

    try:
        resp = requests.post(f"{OLLAMA_BASE_URL}/api/chat", json=payload, timeout=120)
        resp.raise_for_status()
        return resp.json()
    except requests.exceptions.Timeout:
        raise HTTPException(status_code=504, detail="Ollama timeout - le modèle est peut-être trop lent")
    except requests.exceptions.ConnectionError:
        raise HTTPException(status_code=503, detail="Ollama non disponible. Vérifiez qu'il tourne sur localhost:11434")
    except Exception as e:
        logger.error(f"Ollama error: {e}")
        raise HTTPException(status_code=500, detail=f"Erreur Ollama: {str(e)}")


def get_module_tools() -> list:
    """Génère la liste des outils (modules AuditPro) pour Ollama."""
    if not registry:
        return []

    tools = []
    for name, module in registry.get_all().items():
        try:
            inputs_desc = []
            for inp in module.get_required_inputs():
                inputs_desc.append(f"- {inp.key} ({inp.input_type}): {inp.label}")

            tool_name = f"execute_{name.lower().replace(' ', '_').replace('-', '_')}"

            properties = {
                "reasoning": {
                    "type": "string",
                    "description": "Expliquez pourquoi ce module est approprié pour cette requête"
                }
            }
            required = ["reasoning"]

            for inp in module.get_required_inputs():
                if inp.required:
                    properties[inp.key] = {
                        "type": "string",
                        "description": inp.label,
                        "examples": inp.default if inp.default else []
                    }
                    required.append(inp.key)

            tools.append({
                "type": "function",
                "function": {
                    "name": tool_name,
                    "description": f"{module.description}\n\nInputs requis:\n" + "\n".join(inputs_desc),
                    "parameters": {
                        "type": "object",
                        "properties": properties,
                        "required": required,
                        "additionalProperties": False
                    }
                }
            })
        except Exception as e:
            logger.warning(f"Failed to create tool for module {name}: {e}")

    return tools


# ── Routes API ───────────────────────────────────────────
@app.get("/", response_model=Dict[str, Any])
def root():
    """Endpoint de bienvenue avec état du service."""
    return {
        "service": "AuditPro AI Agent",
        "version": "1.0",
        "description": "API d'automatisation d'audit pilotée par IA",
        "modules_disponibles": registry.names() if registry else [],
        "ollama_model": OLLAMA_MODEL,
        "ollama_url": OLLAMA_BASE_URL,
        "endpoints": {
            "chat": "/chat",
            "modules": "/modules",
            "execute": "/execute/{module_name}",
            "upload": "/upload",
            "analyze": "/analyze",
            "health": "/health",
            "ui": "/ui",
            "docs": "/docs"
        }
    }


@app.get("/health", response_model=HealthResponse)
def health_check():
    """Vérifie l'état de santé du service."""
    ollama_ok = check_ollama_available()
    return {
        "status": "healthy" if ollama_ok and registry else "degraded",
        "ollama": ollama_ok,
        "modules_loaded": len(registry.names()) if registry else 0,
        "version": "1.0"
    }


@app.get("/modules", response_model=List[ModuleInfo])
def list_modules():
    """Liste tous les modules AuditPro disponibles."""
    if not registry:
        raise HTTPException(status_code=503, detail="AuditPro modules not loaded")

    return [
        ModuleInfo(
            name=m.name,
            description=m.description,
            category=m.category,
            inputs=[
                {"key": i.key, "label": i.label, "type": i.input_type, "required": i.required, "tooltip": i.tooltip}
                for i in m.get_required_inputs()
            ]
        )
        for m in registry.get_all().values()
    ]


@app.post("/chat", response_model=AIResponse)
def chat(request: ChatRequest):
    """
    Envoie un message en langage naturel.
    L'IA détermine le module approprié et l'exécute si nécessaire.
    """
    if not registry:
        return AIResponse(
            response="AuditPro modules not loaded. Check server logs.",
            error="Modules not available"
        )

    # System prompt for better routing
    system_prompt = """Tu es un assistant expert en audit financier et comptable.
Tu as accès à des modules AuditPro pour automatiser des tâches.
Quand l'utilisateur demande une tâche d'audit, identifie le module approprié et utilise-le.

Modules disponibles:
"""
    for name, module in registry.get_all().items():
        system_prompt += f"- {name}: {module.description}\n"

    system_prompt += "\nSi la demande ne nécessite pas de module, réponds directement."

    messages = [{"role": "system", "content": system_prompt}]
    messages.extend(request.conversation)
    messages.append({"role": "user", "content": request.message})

    tools = get_module_tools()
    response = call_ollama(messages, tools if tools else None)
    message = response.get("message", {})

    if message.get("tool_calls"):
        for tool_call in message["tool_calls"]:
            func_name = tool_call["function"]["name"]
            args = tool_call["function"].get("arguments", {})

            # Trouver le module correspondant
            module_name = func_name.replace("execute_", "").replace("_", " ").title()
            module = registry.get(module_name)

            if not module:
                # Try alternative name formats
                for alt_name in registry.names():
                    if alt_name.lower() in func_name.lower():
                        module = registry.get(alt_name)
                        break

            if module:
                try:
                    # Exécuter le module
                    with tempfile.TemporaryDirectory() as tmpdir:
                        inputs = {k: v for k, v in args.items() if k != "reasoning"}
                        logger.info(f"Executing module {module.name} with inputs: {list(inputs.keys())}")

                        result = module.execute(inputs, tmpdir)

                        if result.success:
                            return AIResponse(
                                response=f"Module {module.name} exécuté avec succès. {result.message}",
                                module_used=module.name,
                                output_path=result.output_path,
                                data=result.stats
                            )
                        else:
                            return AIResponse(
                                response=f"Échec de l'exécution: {result.message}",
                                module_used=module.name,
                                error=result.message
                            )
                except Exception as e:
                    logger.error(f"Module execution error: {e}")
                    return AIResponse(
                        response=f"Erreur lors de l'exécution: {str(e)}",
                        module_used=module.name,
                        error=str(e)
                    )
            else:
                logger.error(f"Module not found for: {func_name}")
                return AIResponse(
                    response=f"Module '{module_name}' non trouvé",
                    error=f"Module not found: {module_name}"
                )

    return AIResponse(
        response=message.get("content", ""),
        module_used=None,
        output_path=None,
        data=None
    )


@app.post("/execute/{module_name}")
def execute_module(module_name: str, request: ModuleExecuteRequest):
    """Exécute un module AuditPro spécifique avec les inputs fournis."""
    if not registry:
        raise HTTPException(status_code=503, detail="AuditPro modules not loaded")

    # Try case-insensitive lookup
    module = None
    for name in registry.names():
        if name.lower() == module_name.lower():
            module = registry.get(name)
            break

    if not module:
        raise HTTPException(
            status_code=404,
            detail=f"Module '{module_name}' non trouvé. Modules disponibles: {registry.names()}"
        )

    # Validation des inputs
    try:
        valid, errors = module.validate(request.inputs)
        if not valid:
            raise HTTPException(status_code=400, detail="; ".join(errors))
    except NotImplementedError:
        logger.warning(f"Module {module_name} has no validation implemented")

    with tempfile.TemporaryDirectory() as tmpdir:
        try:
            result = module.execute(request.inputs, tmpdir, request.params)

            if result.success:
                return {
                    "success": True,
                    "message": result.message,
                    "output_path": result.output_path,
                    "stats": result.stats,
                    "warnings": result.warnings
                }
            else:
                return {
                    "success": False,
                    "message": result.message,
                    "errors": result.errors
                }
        except Exception as e:
            logger.error(f"Module execution error: {e}")
            raise HTTPException(status_code=500, detail=f"Erreur d'exécution: {str(e)}")


@app.post("/upload")
async def upload_file(file: UploadFile = File(...)):
    """Upload un fichier pour analyse et détection automatique du module."""
    if not registry:
        raise HTTPException(status_code=503, detail="AuditPro modules not loaded")

    upload_dir = AUDITPRO_DIR / "uploads"
    upload_dir.mkdir(exist_ok=True)

    # Sanitize filename to prevent path traversal attacks
    safe_name = Path(file.filename).name  # strip any directory components
    if not safe_name or safe_name.startswith("."):
        raise HTTPException(status_code=400, detail="Nom de fichier invalide")
    file_path = (upload_dir / safe_name).resolve()
    try:
        file_path.relative_to(upload_dir.resolve())
    except ValueError:
        raise HTTPException(status_code=400, detail="Chemin de fichier non autorisé")

    with open(file_path, "wb") as f:
        shutil.copyfileobj(file.file, f)

    # Analyse automatique du fichier
    try:
        from core.file_detector import FileDetector
        detector = FileDetector()
        all_modules = registry.get_all()
        detections = detector.detect(str(file_path), all_modules)

        modules_suggérés = []
        for item in detections:
            mod = item.get("module")
            if mod is None:
                continue
            score = item.get("score", 0)
            if score >= getattr(mod, "detection_threshold", 0.5):
                modules_suggérés.append({
                    "name": mod.name,
                    "score": round(score * 100, 1),
                    "description": mod.description,
                    "matched_keywords": item.get("matched_keywords", []),
                })

        file_info = detector.get_file_info(str(file_path))

        return {
            "file_path": str(file_path),
            "filename": safe_name,
            "size_kb": file_info.get("size_kb", 0),
            "total_rows": file_info.get("total_rows", 0),
            "sheets": file_info.get("sheets", []),
            "detected_modules": modules_suggérés,
            "recommendation": modules_suggérés[0]["name"] if modules_suggérés else "Aucun module détecté"
        }
    except Exception as e:
        logger.error(f"File detection error: {e}")
        return {
            "file_path": str(file_path),
            "filename": safe_name,
            "detected_modules": [],
            "error": str(e)
        }


@app.post("/analyze")
def analyze_with_ai(file_path: str, question: str = "Que contient ce fichier ?"):
    """Analyse un fichier (déjà uploadé) avec Ollama et répond à une question."""
    if not registry:
        raise HTTPException(status_code=503, detail="AuditPro modules not loaded")

    # Restrict access to the uploads directory only — use relative_to() for reliable containment check
    upload_dir = (AUDITPRO_DIR / "uploads").resolve()
    resolved = Path(file_path).resolve()
    try:
        resolved.relative_to(upload_dir)
    except ValueError:
        raise HTTPException(
            status_code=403,
            detail="Accès refusé : seuls les fichiers dans le répertoire uploads sont analysables"
        )

    # Lire le fichier (support Excel/CSV)
    import pandas as pd

    path = resolved
    if not path.exists():
        raise HTTPException(status_code=404, detail="Fichier non trouvé")

    try:
        if path.suffix in [".xlsx", ".xls"]:
            df = pd.read_excel(path, nrows=100)
        elif path.suffix == ".csv":
            df = pd.read_csv(path, nrows=100)
        else:
            raise HTTPException(status_code=400, detail="Format non supporté")
    except Exception as e:
        raise HTTPException(status_code=400, detail=f"Erreur lecture fichier: {e}")

    # Préparer un résumé pour Ollama
    summary = f"""Fichier: {path.name}
Colonnes: {', '.join(df.columns.tolist())}
Lignes: {len(df)}
Aperçu (5 premières lignes):
{df.head().to_string()}

Question: {question}"""

    messages = [{"role": "user", "content": summary}]
    response = call_ollama(messages)

    return {
        "file": path.name,
        "rows": len(df),
        "columns": list(df.columns),
        "ai_analysis": response["message"]["content"]
    }


# ── Static Files (Web UI) ─────────────────────────────────
static_dir = Path(__file__).parent / "static"
static_dir.mkdir(exist_ok=True)

@app.get("/ui", response_class=HTMLResponse)
def serve_ui():
    """Sert l'interface web de l'agent."""
    index_path = static_dir / "index.html"
    if index_path.exists():
        return index_path.read_text()
    return HTMLResponse("<h1>AuditPro AI Agent</h1><p>UI not found. Check static/index.html</p>")


# ── Lancement ────────────────────────────────────────────
if __name__ == "__main__":
    import uvicorn
    print()
    print("╔═══════════════════════════════════════════════════════════╗")
    print("║            AuditPro AI Agent — Serveur API               ║")
    print("╠═══════════════════════════════════════════════════════════╣")
    print(f"║ Modules chargés : {len(registry.names()) if registry else 0}")
    print(f"║ Ollama : {OLLAMA_MODEL} @ {OLLAMA_BASE_URL}")
    print("╠═══════════════════════════════════════════════════════════╣")
    print("║ Endpoints:                                                ║")
    print("║   API:  http://localhost:8000                             ║")
    print("║   Docs: http://localhost:8000/docs                        ║")
    print("║   UI:   http://localhost:8000                             ║")
    print("╚═══════════════════════════════════════════════════════════╝")
    print()
    uvicorn.run(app, host="0.0.0.0", port=8000)
