# 🚀 AuditPro - Guide d'Installation et Utilisation

## 📋 Prérequis

- **Python 3.9+** (recommandé: Python 3.10+)
- **Windows 10+**

## 🔧 Installation

### 1. Installer les dépendances

```bash
pip install -r requirements.txt
```

### 2. Lancer l'application

```bash
python main.py
```

## 📁 Structure du Projet

```
AuditPro/
├── core/              # Cœur fonctionnel (config, registry, profils)
├── ui/                # Interface PyQt6 (main_window, workspace, modules)
├── modules/           # Modules d'audit (circularisation, tva, extraction, etc.)
├── resources/         # Images, logos, icônes
├── data/              # Données persistantes (profils, historique)
├── main.py            # Point d'entrée
└── requirements.txt   # Dépendances Python
```

## ⚙️ Configuration des Modules

Chaque module a un fichier `module.py` qui :
- Dialogue avec l'interface via `InputSchema` et `OutputSchema`
- Charge les paramètres depuis les profils sauvegardés
- Accepte n'importe quel fichier Excel/CSV compatible

### Ajouter un Nouveau Module

1. Créer un dossier `modules/mon_module/`
2. Ajouter les fichiers :
   ```
   modules/mon_module/
   ├── __init__.py
   ├── module.py         # Wrapper (InputSchema, execute, OutputSchema)
   └── engine.py         # Logique métier
   ```
3. Le registre détecte automatiquement le nouveau module

## 🎨 Personnalisation

- **Thème**: Éditer `core/config.py` → section `COLORS`
- **Styles**: Éditer `ui/styles.py` → section `STYLESHEET`
- **Couleurs modules**: Modifier `#00ffff` dans `ui/styles.py`

## 🧪 Test Rapide

```bash
# Vérifier que l'app démarre
python -m py_compile main.py ui/main_window.py

# Lancer avec debug
python main.py
```

## 📝 Fichiers Supportés

- `.xlsx`, `.xls`, `.xlsm`
- `.csv` (UTF-8)

## 💡 Conseils d'Utilisation

1. **Profils**: Créer un profil pour chaque type d'audit (colonnes, filtres)
2. **Historique**: L'app garde un historique des 20 derniers traitements
3. **Export**: Les résultats sont générés dans `data/outputs/`

## 🐛 Troubleshooting

**L'app ne démarre pas:**
- Vérifier: `python main.py --version`
- Réinstaller: `pip install -r requirements.txt --force-reinstall`

**Les modules ne chargent pas:**
- Vérifier structure: `modules/nom_module/module.py` doit exister
- Regarder console pour les logs d'erreur

---

**Développé par:** Abderrahmane Chougua  
**Contact:** abderrahmanechougua@gmail.com
