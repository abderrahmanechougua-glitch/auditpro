# Migration du Module Retraitement v1 → v2.0

## Résumé des changements

Le module `retraitement` a été complètement refactorisé pour être **production-grade**, **audit-safe** et **configurable**.

### Avant (v1) vs Après (v2.0)

| Aspect | v1 | v2.0 |
|--------|-----|------|
| **Performance** | Boucles Python (apply axis=1) | Vectorisé NumPy/Pandas (O(n)) |
| **Audit trail** | Suppression silencieuse de lignes | Flagging complet, traçabilité Excel |
| **Configuration** | Hardcodée dans le code | Externalisée (config.py) |
| **Architecture** | Monolithique (main.py) | Modulaire et testable |
| **Rapports** | Basique | Multi-feuilles avec audit complet |
| **Gestion d'erreurs** | Partielle | Robuste avec logging |

---

## Nouvelles structures

### 1. Modules refactorisés

```
modules/retraitement/
├── __init__.py                 # API publique
├── processor.py                # ⭐ Orchestrateur principal
├── config.py                   # Configuration externalisée
├── loader.py                   # Chargement Excel
├── detector.py                 # Détection de type de document
├── cleaner.py                  # Marquage des lignes (flagging)
├── normalizer.py               # Standardisation et conversion types
├── validator.py                # Validation des règles métier
├── reporter.py                 # Génération rapports Excel
└── main.py                     # ⚠️ Shim de compatibilité (à supprimer)
```

### 2. API Publique (inchangée pour les utilisateurs)

```python
from modules.retraitement import (
    IntelligentRetraitement,  # Classe orchestratrice
    Config,                    # Configuration
    process_file,              # Fonction raccourcie
    process_files,             # Batch processing
)

# Usage simple
result = process_file("data/gl.xlsx")
if result["success"]:
    df = result["dataframe"]
    print(result["report_file"])

# Usage avancé
config = Config(tolerance=0.005, erp_profile="sap")
processor = IntelligentRetraitement(config=config, output_dir="./reports")
result = processor.process_file("data/bg.xlsx", doc_type_hint="BG")
```

---

## Guide de migration

### Phase 1 : Tests (Vous êtes ici)

✅ Tous les modules refactorisés sont en place  
✅ API publique stable et rétro-compatible  
✅ Logs structurés pour debugging  

**À tester :**
- [ ] Charger un fichier GL
- [ ] Charger un fichier BG
- [ ] Charger un fichier AUX
- [ ] Vérifier les rapports Excel générés
- [ ] Vérifier l'audit trail

### Phase 2 : Intégration UI (À faire)

Fichiers à mettre à jour :

#### [ui/main_window.py](../ui/main_window.py)

**Avant :**
```python
from modules.retraitement.main import IntelligentRetraitement
```

**Après :**
```python
from modules.retraitement import IntelligentRetraitement, Config
```

#### [ui/module_panel.py](../ui/module_panel.py)

Vérifier que les callbacks utilisent la nouvelle API :
```python
# Si usage : processor = IntelligentRetraitement()
# Ajouter : config = Config(tolerance=0.005)
# processor = IntelligentRetraitement(config=config)
```

#### [INTEGRATION_GUIDE.py](../../INTEGRATION_GUIDE.py)

Mettre à jour les exemples d'intégration.

### Phase 3 : Nettoyage (Optionnel)

Une fois testée et intégrée :

```bash
# Supprimer le shim de compatibilité
rm modules/retraitement/main.py

# Archiver l'ancien code
git tag v1.0-legacy HEAD  # ou git branch legacy
```

---

## Changements de comportement

### 1. Marquage des lignes (Audit trail)

**v1 :** Supprimait silencieusement les lignes problématiques  
**v2.0 :** Les marque avec une colonne `_flag`

```python
# Nouvelle colonne dans tous les DataFrames :
# '_flag' : None (OK), 'empty_row', 'duplicate', 'no_numerics', etc.
```

**Impact :** Permet traçabilité complète, aucune donnée perdue silencieusement

### 2. Rapports Excel

**v1 :** Deux feuilles (Données + Contrôle)  
**v2.0 :** Quatre feuilles

1. **Données_Normalisées** : Données traitées + colonne `_flag`
2. **Validation** : Violations des règles métier avec sévérité
3. **Transformation_Log** : Audit trail détaillé de chaque étape
4. **Métadonnées** : Infos fichier source + détails traitement

### 3. Configuration

**v1 :** Hardcodée  
**v2.0 :** Externalisée (modifiable sans code)

```python
# config.py
config = Config(
    tolerance=0.005,              # Seuil débit/crédit
    erp_profile="sap",            # Profil SAP/Ciel/Sage
    column_mapping={...},         # Mappages personnalisés
    validation_rules={...},       # Règles métier
)
```

### 4. Logging

**v1 :** Minimal  
**v2.0 :** Structuré et verbeux

```
2024-01-15 10:30:45 - modules.retraitement.processor - INFO - Traitement du fichier : data/gl.xlsx
2024-01-15 10:30:45 - modules.retraitement.processor - INFO - Fichier chargé : 1250 lignes, 8 colonnes
2024-01-15 10:30:45 - modules.retraitement.processor - INFO - Type détecté : GL (confiance: 92.5%)
...
```

---

## Checklist de migration

- [ ] **Tests unitaires** : Vérifier chaque module indépendamment
- [ ] **Tests d'intégration** : Processor + UI
- [ ] **Tests de régression** : Comparer v1 vs v2.0 sur fichiers anciens
- [ ] **Performance** : Benchmarker sur gros fichiers (10K+ lignes)
- [ ] **Audit trail** : Vérifier que les rapports sont générés
- [ ] **Documentation** : Mettre à jour les docs API
- [ ] **Support** : Former les utilisateurs aux nouveaux rapports
- [ ] **Cleanup** : Supprimer v1 après validation complète

---

## Troubleshooting

### Q: Mon code l'utilise encore `from modules.retraitement.main import ...` ?
**R:** C'est OK, le shim dans main.py assure la compatibilité. Mais mettez à jour dès que possible.

### Q: Les rapports Excel ont plus de feuilles, ça casse mon intégration ?
**R:** Les trois premières feuilles (Données, Validation, Métadonnées) sont les principales. Ignorez Transformation_Log si vous n'en avez pas besoin.

### Q: Pourquoi les flagged rows ne sont pas supprimées ?
**R:** C'est l'audit trail. Les données ne doivent jamais être perdues silencieusement. Vous pouvez les filtrer manuellement avec `df[df['_flag'].isna()]`.

### Q: Comment appliquer un profil ERP ?
**R:** 
```python
processor = IntelligentRetraitement(config=Config(erp_profile="sap"))
# ou
result = process_file("data.xlsx", erp_profile="sage")
```

---

## Support et questions

Pour toute question : Consultez la docstring de chaque module (processor.py, config.py, etc.).

---

**Dernière mise à jour :** 2024-01-15  
**Version :** 2.0.0  
**Auteur :** AuditPro Engine
