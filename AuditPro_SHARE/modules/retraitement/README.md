# Retraitement Comptable - Module v2.0

Module de retraitement intelligent des documents comptables. **Production-grade**, **audit-safe**, **vectorisé**.

## Caractéristiques

✅ **Vectorisé** : O(n) performance, pas de boucles Python  
✅ **Audit-safe** : Traçabilité complète, zéro suppression silencieuse  
✅ **Configurable** : Règles externalisées, profils ERP (SAP, Ciel, Sage)  
✅ **Modulaire** : Chaque étape testable indépendamment  
✅ **Robuste** : Gestion complète des erreurs, logging structuré  
✅ **Production-ready** : Rapports multi-feuilles Excel, audit trail  

## Types de documents supportés

| Type | Description | Colonnes clés |
|------|-------------|---------------|
| **GL** | Grand Livre | N°Compte, Date, Journal, Pièce, Libellé, Débit/Crédit |
| **BG** | Balance Générale | N°Compte, Intitulé, Solde Débit/Crédit |
| **AUX** | Balance Auxiliaire | Tiers, Solde Initial/Final, Mouvements Débit/Crédit |

## Installation rapide

```bash
# Automatique (requirements.txt)
pip install -r requirements.txt

# Manuel si nécessaire
pip install pandas openpyxl numpy
```

## Utilisation

### Simple (2 lignes)

```python
from modules.retraitement import process_file

result = process_file("data/gl.xlsx")
print(result["report_file"])  # Chemin du rapport Excel généré
```

### Avancée (configuration personnalisée)

```python
from modules.retraitement import IntelligentRetraitement, Config

# Configurer
config = Config(
    tolerance=0.005,           # Seuil d'équilibre débit/crédit
    erp_profile="sap",         # Profil ERP (default/sap/ciel/sage)
    remove_empty_rows=True,    # Supprimer les lignes vides
    remove_total_rows=False,   # Conserver les totaux
)

# Processor
processor = IntelligentRetraitement(
    config=config,
    output_dir="./reports"
)

# Traiter
result = processor.process_file(
    "data/bg.xlsx",
    doc_type_hint="BG",        # Forcer le type (optionnel)
    erp_profile="sap",         # Profil pour ce fichier
    keep_flagged_rows=True,    # Conserver les lignes flaggées
)

# Résultat
if result["success"]:
    df = result["dataframe"]
    report = result["report_file"]
    print(f"✓ Rapport : {report}")
    print(f"✓ Lignes : {len(df)}")
    print(f"✓ Violations : {len(result['validation_report'])}")
else:
    print(f"✗ Erreur : {result['error']}")
```

### Batch (plusieurs fichiers)

```python
from modules.retraitement import process_files
from pathlib import Path

files = list(Path("data").glob("*.xlsx"))
results = process_files(files, output_dir="./output")

for result in results:
    status = "✓" if result["success"] else "✗"
    print(f"{status} {result['metadata']['file_path']}")
```

## Structure des rapports Excel

Chaque fichier traité génère un rapport avec **4 feuilles** :

### 1. **Données_Normalisées**
- Les données après standardisation et normalisation
- Colonne `_flag` : État de chaque ligne (None = OK)
- Colonnes standardisées selon le type de document

**Exemple :**
| n°compte | intitulé | date | libellé | débit | crédit | _flag |
|----------|----------|------|---------|-------|--------|-------|
| 401 | Fournisseur X | 2024-01-15 | Facture | 500.00 | 0.00 | None |
| 512 | Banque | 2024-01-15 | Paiement | 0.00 | 500.00 | None |

### 2. **Validation**
- Violations des règles métier
- Sévérité : INFO, WARNING, ERROR
- Message détaillé + ligne affectée + valeur

**Exemple :**
| rule | severity | message | value | row |
|------|----------|---------|-------|-----|
| unbalanced_balance | ERROR | Débit/Crédit non équilibrés | 12.50 | 5 |
| invalid_date | WARNING | Date invalide | "2024-13-01" | 8 |

### 3. **Transformation_Log**
- Audit trail complet des transformations
- Phase : input, detection, normalization, validation, cleaning
- Action + timestamp + détails

**Exemple :**
| timestamp | phase | action | details | value |
|-----------|-------|--------|---------|-------|
| 2024-01-15T10:30:45 | input | file_loaded | Fichier : data/gl.xlsx \| Feuille : GL | 1250 |
| 2024-01-15T10:30:45 | detection | doc_type_detected | Type : GL \| Confiance : 92.5% \| Raison : ... | 0.925 |
| 2024-01-15T10:30:45 | normalization | column_renamed | 'N°Compte' → 'n°compte' | None |

### 4. **Métadonnées**
- Infos fichier source (chemin, feuille, en-tête détecté)
- Infos détection (type, confiance, raison)
- Infos traitement (lignes brutes/finales, colonnes, mappages)
- Timestamp du rapport

---

## Architecture modulaire

```
processor.py          → Orchestrateur principal (API publique)
├─ loader.py          → Chargement et détection d'en-tête
├─ detector.py        → Détection automatique du type
├─ cleaner.py         → Flagging des lignes problématiques
├─ normalizer.py      → Standardisation + conversion de types
├─ validator.py       → Validation des règles métier
├─ reporter.py        → Génération rapports Excel
└─ config.py          → Configuration externalisée
```

Chaque module :
- ✅ Testable indépendamment
- ✅ Vectorisé (pas de boucles Python sur les données)
- ✅ Documenté avec docstrings
- ✅ Gère ses propres erreurs
- ✅ Loggé

---

## Configuration

Voir [config.py](config.py) pour :
- Paramètres de tolérance (débit/crédit, dates)
- Profils ERP (SAP, Ciel, Sage)
- Règles de flagging personnalisées
- Mappages de colonnes

```python
from modules.retraitement import Config

config = Config(
    tolerance=0.005,                    # ±0.5% sur débit/crédit
    remove_empty_rows=True,             # Supprimer les lignes vides
    remove_total_rows=False,            # Conserver les totaux
    erp_profile="default",              # Profil à appliquer
    validation_rules={"min_rows": 10}   # Règles custom
)
```

---

## Exemples de code

### Exemple 1 : Détecter automatiquement le type

```python
from modules.retraitement import process_file

result = process_file("mystery_file.xlsx")
print(f"Type détecté : {result['doc_type']}")
print(f"Confiance : {result['doc_type_explanation']['confidence']:.1%}")
```

### Exemple 2 : Forcer un type et filtrer les flagged rows

```python
from modules.retraitement import process_file

result = process_file("data/bg.xlsx", doc_type_hint="BG")

# Données sans les flagged rows
df_clean = result["dataframe"][result["dataframe"]["_flag"].isna()]
print(f"Lignes valides : {len(df_clean)}")
```

### Exemple 3 : Accéder aux violations de validation

```python
result = process_file("data/gl.xlsx")

violations = result["validation_report"]
errors = violations[violations["severity"] == "ERROR"]

for _, error in errors.iterrows():
    print(f"Line {error['row']} : {error['message']}")
```

### Exemple 4 : Intégration dans une application

```python
from modules.retraitement import IntelligentRetraitement, Config

class AuditEngine:
    def __init__(self):
        config = Config(erp_profile="sap", tolerance=0.01)
        self.processor = IntelligentRetraitement(
            config=config,
            output_dir="./audit_reports"
        )
    
    def process_invoice_file(self, path):
        result = self.processor.process_file(
            path,
            doc_type_hint="GL",
            keep_flagged_rows=False
        )
        return result["dataframe"], result["report_file"]
```

---

## Logging

Activez le logging pour déboguer :

```python
import logging

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s"
)

# Tous les logs retraitement apparaîtront
```

**Niveaux :**
- `DEBUG` : Détails bas niveau
- `INFO` : Progression du traitement
- `WARNING` : Anomalies (colonnes manquantes, dates invalides)
- `ERROR` : Échecs (fichier introuvable, format invalide)

---

## Migration depuis v1

Les utilisateurs de v1 peuvent continuer avec le shim de compatibilité dans main.py.  
Voir [MIGRATION.md](MIGRATION.md) pour détails.

```python
# v1 (ancien, encore compatible)
from modules.retraitement.main import IntelligentRetraitement

# v2.0 (recommandé)
from modules.retraitement import IntelligentRetraitement
```

---

## Support

- **Bugs** : Consultez [MIGRATION.md](MIGRATION.md#troubleshooting)
- **Docstrings** : Lisez le docstring de chaque fonction
- **Tests** : Voir si des tests unitaires existent

---

**Version :** 2.0.0  
**Dernière mise à jour :** 2024-01-15  
**Auteur :** AuditPro Engine  
**Support :** Production-ready, audit-safe, vectorisé
