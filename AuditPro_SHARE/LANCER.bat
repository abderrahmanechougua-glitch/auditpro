@echo off
REM AuditPro v1.0.0 - Launcher principal
REM Lance SOIT le mode éditable (run_internal.py) SOIT l'exe packagé, JAMAIS les deux

setlocal enabledelayedexpansion
cd /d "%~dp0"

echo.
echo ╔══════════════════════════════════════════════════════╗
echo ║         AUDITPRO v1.0.0 - Démarrage                  ║
echo ╚══════════════════════════════════════════════════════╝
echo.

REM ── Vérifier Python ───────────────────────────────────────
set "PYTHON_CMD=python"
if exist "..\.venv\Scripts\python.exe" (
    set "PYTHON_CMD=..\.venv\Scripts\python.exe"
)

"%PYTHON_CMD%" --version >nul 2>&1
if errorlevel 1 (
    echo Erreur: Python n'est pas trouvé.
    echo Installez Python 3.9+ depuis https://www.python.org
    pause
    exit /b 1
)

echo Python détecté
echo.

REM ── STRATEGY: Try ONE method, not multiple ───────────────
REM Priority 1: Mode éditable (run_internal.py)
if exist "run_internal.py" (
    echo Lancement en mode éditable...
    echo.
    "%PYTHON_CMD%" "run_internal.py"
    REM On exit du mode éditable, terminer complètement
    exit /b %ERRORLEVEL%
)

REM Priority 2: Version packagée (.exe) - seulement si run_internal.py n'existe pas
if exist "dist\AuditPro\AuditPro.exe" (
    echo Lancement de la version packagée...
    echo.
    start "" "dist\AuditPro\AuditPro.exe"
    exit /b 0
)

REM ── Erreur: Aucune méthode de démarrage disponible ────────
echo.
echo ERREUR: Impossible de démarrer AuditPro
echo  - run_internal.py n'existe pas
echo  - dist\AuditPro\AuditPro.exe n'existe pas
echo.
pause
exit /b 1
