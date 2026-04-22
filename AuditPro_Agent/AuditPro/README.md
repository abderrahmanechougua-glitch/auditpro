# AuditPro v1.1 — Assistant d'audit intelligent

**FIDAROC GRANT THORNTON** | 8 modules | PyQt6 | Windows

## Installation

```bash
pip install -r requirements.txt
python main.py
```

## Activation des modules — copiez vos scripts ici :

| Module | Fichier à copier | → Dossier |
|--------|-----------------|-----------|
| SRM Generator | `srmgen.py` | `modules/srm_generator/` |
| Centralisation TVA | `extract_tva_v3.py` | `modules/tva/` |
| Centralisation CNSS | `extract_cnss.py` | `modules/cnss/` |
| Extraction Factures | `factext_v2.py` | `modules/extraction_factures/` |
| Extraction IR | `extract_ir.py` | `modules/extraction_ir/` |
| Retraitement Comptable | `main.py` (v7) | `modules/retraitement/` |
| Lettrage Grand Livre | `lettrage_gl.py` | `modules/lettrage/` |
| Circularisation | `script.py` | `modules/circularisation/` |

## Packaging Windows

```bat
build.bat
```
