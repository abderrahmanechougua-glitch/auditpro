@echo off
REM Script de démarrage AuditPro v1.0.0
REM Cet script installe les dépendances et lance l'application

cls
echo.
echo ╔══════════════════════════════════════════════════════╗
echo ║         AUDITPRO v1.0.0 - Démarrage                  ║
echo ╚══════════════════════════════════════════════════════╝
echo.

REM Vérifier Python
python --version >nul 2>&1
if errorlevel 1 (
    echo ✗ ERREUR: Python n'est pas installé
    echo   Installez Python 3.9+ depuis https://www.python.org
    pause
    exit /b 1
)

echo ✓ Python détecté
echo.

REM Installer les dépendances
echo Vérification des dépendances...
pip install -r requirements.txt --quiet

if errorlevel 1 (
    echo ✗ ERREUR lors de l'installation des dépendances
    pause
    exit /b 1
)

echo ✓ Dépendances OK
echo.

REM Lancer l'app
echo Démarrage d'AuditPro...
echo.

python main.py

if errorlevel 1 (
    echo.
    echo ✗ ERREUR lors du démarrage
    echo   Consultez la sortie d'erreur ci-dessus
    pause
    exit /b 1
)

exit /b 0
