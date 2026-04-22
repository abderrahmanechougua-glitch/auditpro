# AuditPro v1.0.0 - Dossier Partagé

## ✅ Contenu du Dossier

Ce dossier contient une version **propre et optimisée** d'AuditPro prête à être utilisée.

- ✓ **8 modules d'audit** présente et fonctionnels
- ✓ **Interface PyQt6** nettoyée et optimisée
- ✓ **Couleur modules**: Violet clair (#ee82ee)
- ✓ **Tous les fichiers inutiles supprimés** (build, dist, cache)

## 🚀 Démarrage Rapide

### 1. Installer les dépendances
```powershell
pip install -r requirements.txt
```

### 2. Lancer l'application
```powershell
python main.py
```

## 📦 Modules Disponibles

1. **Circularisation des Tiers** - Génération courriers de circularisation
2. **Centralisation CNSS** - Extraction et centralisation données CNSS
3. **Extraction Factures** - Extraction factures depuis fichiers Excel
4. **Extraction IR** - Extraction revenu imposable
5. **Lettrage Grand Livre** - Lettrage comptable automatique
6. **Retraitement Comptable** - Ajustements comptables
7. **SRM Generator** - Génération SRM
8. **Centralisation TVA** - Centralisation données TVA

## 📋 Configuration

- **Modifier les couleurs**: Éditer `core/config.py` → section `COLORS`
- **Personnaliser styles**: Éditer `ui/styles.py`
- **Ajouter modules**: Créer un dossier dans `modules/`

## 🧪 Tests

```powershell
# Vérifier installation
python -c "from core.config import APP_NAME; print(f'{APP_NAME} v1.0.0 ✓')"

# Voir les modules chargés
python -c "from core.module_registry import ModuleRegistry; r = ModuleRegistry(); print(f'{len(r.get_all())} modules trouvés')"
```

## 📖 Documentation

Voir `SETUP.md` pour l'installation complète et le troubleshooting.

## 👤 Auteur

**Abderrahmane Chougua**  
📧 abderrahmanechougua@gmail.com

---

**Date de création**: 14/04/2026  
**Version**: 1.0.0  
**État**: ✅ Production Ready
