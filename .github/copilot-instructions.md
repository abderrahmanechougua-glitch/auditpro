# GitHub Copilot Instructions — AuditPro

This repository uses **three custom Copilot agents** that work together to
create a continuous improvement loop. Invoke them in VS Code Copilot Chat
using the `@agent-name` syntax.

---

## Agent Overview

| Agent | Trigger | Role |
|---|---|---|
| Activity Logger | `@activity-logger` | Records every problem, Copilot suggestion outcome, bug, and manual intervention |
| Memoriser | `@memoriser` | Maintains the living project knowledge base (architecture, patterns, decisions) |
| Copilot Optimization Agent | `@copilot-optimization-agent` | Analyses logs + memory → generates Knowledge, Templates, Workflows, Reports |

---

## How They Work Together

```
You (Developer facing a problem)
          │
          ▼
┌─────────────────────┐     ┌─────────────────────┐
│  @activity-logger   │────▶│    @memoriser       │
│  (What happened)    │     │  (Project context)  │
└──────────┬──────────┘     └──────────┬──────────┘
           │         Logs + Context    │
           └─────────────┬─────────────┘
                         │
                         ▼
          ┌──────────────────────────────┐
          │  @copilot-optimization-agent │
          │  Produces:                   │
          │  · Knowledge Syntheses       │
          │  · Instruction Templates     │
          │  · Workflow Definitions      │
          │  · Quarterly Reports         │
          └──────────────────────────────┘
```

---

## Project Context (Always Active)

- **Application**: AuditPro — a PyQt6 desktop app for Moroccan accounting audits.
- **AI Layer**: FastAPI + Ollama (llama3.2) in `AuditPro_Agent/`.
- **Modules**: TVA, CNSS, Lettrage, Retraitement, Circularisation, SRM Generator, Extraction Factures, Extraction IR.
- **Data formats**: Excel (`.xlsx`) and CSV via pandas / openpyxl.
- **Language rule**: UI strings in French; code identifiers in English.
- **Packaging**: PyInstaller → Inno Setup installer.

---

## Quick-Start Commands

```
# Log a problem
@activity-logger Log: I asked Copilot to add a new TVA module method and it used the wrong base class.

# Store a decision
@memoriser Remember: All modules must inherit from BaseModule in core/base.py.

# Analyse logs
@copilot-optimization-agent Analyse logs: <paste activity log entries>

# Get a workflow
@copilot-optimization-agent Generate workflow for: adding a new AuditPro module

# Quarterly report
@copilot-optimization-agent Quarterly report
```
