# AuditPro AI Agent

API server qui intègre **Ollama** avec les modules **AuditPro** pour permettre :
- Requêtes en langage naturel (français/anglais)
- Sélection automatique du module approprié
- Analyse conversationnelle de fichiers Excel/CSV

## Prérequis

1. **Ollama** installé et en cours d'exécution :
   ```bash
   ollama serve
   ```

2. **Modèle Llama** téléchargé :
   ```bash
   ollama pull llama3.2
   ```

3. **AuditPro** présent dans le dossier parent

## Installation

```bash
cd AuditPro_Agent
pip install -r requirements.txt
```

## Démarrage

```bash
python server.py
```

L'API sera disponible sur `http://localhost:8000`

## Utilisation

### 1. Via le client Python

```python
from client import AuditProAgent

agent = AuditProAgent()

# Chat en langage naturel
response = agent.chat("Je veux centraliser la TVA de 2025")
print(response["response"])

# Upload et analyse d'un fichier
detection = agent.upload_file("declarations.xlsx")
print(detection["detected_modules"])

# Analyse conversationnelle
analysis = agent.analyze_file("grand_livre.csv", "Quels sont les écarts ?")
print(analysis["ai_analysis"])
```

### 2. Via l'API REST

```bash
# Lister les modules
curl http://localhost:8000/modules

# Chat avec l'IA
curl -X POST http://localhost:8000/chat \
  -H "Content-Type: application/json" \
  -d '{"message": "Je veux faire un lettrage du grand livre"}'

# Upload de fichier
curl -X POST http://localhost:8000/upload \
  -F "file=@declarations.xlsx"

# Analyse AI d'un fichier
curl "http://localhost:8000/analyze?file_path=declarations.xlsx&question=Que contient ce fichier ?"
```

### 3. Documentation interactive

Ouvrez `http://localhost:8000/docs` pour accéder à l'interface Swagger UI.

## Architecture

```
AuditPro_Agent/
├── server.py       # API FastAPI + intégration Ollama
├── client.py       # Client Python exemple
├── requirements.txt
└── README.md
```

## Variables d'environnement

| Variable | Défaut | Description |
|----------|--------|-------------|
| `OLLAMA_URL` | `http://localhost:11434` | URL de l'API Ollama |
| `OLLAMA_MODEL` | `llama3.2` | Modèle LLM à utiliser |

## Modules AuditPro supportés

- **TVA** — Centralisation des déclarations TVA
- **CNSS** — Centralisation des bordereaux CNSS
- **Lettrage** — Lettrage automatique du grand livre
- **Retraitement** — Retraitement comptable
- **Circularisation** — Génération des lettres de circularisation
- **SRM Generator** — Génération de rapports SRM
- **Extraction Factures** — Extraction des données de factures
- **Extraction IR** — Extraction des données IR

## Exemple de workflow complet

```python
agent = AuditProAgent()

# 1. Upload du fichier
result = agent.upload_file("balance_aux.xlsx")
print(f"Modules suggérés: {result['detected_modules']}")

# 2. Demande en langage naturel
response = agent.chat(
    "Génère la circularisation pour la société XYZ "
    "avec une couverture de 90%, exercice 2025"
)

if response.get('output_path'):
    print(f"Résultat généré: {response['output_path']}")
    print(f"Stats: {response['data']}")
```
