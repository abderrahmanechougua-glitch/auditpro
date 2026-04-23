@echo off
REM Script de démarrage AuditPro v1.0.0
REM Lance en priorité le runtime editable (_internal) pour refléter les correctifs locaux.

cls
echo.
echo ╔══════════════════════════════════════════════════════╗
echo ║         AUDITPRO v1.0.0 - Démarrage                  ║
echo ╚══════════════════════════════════════════════════════╝
echo.

set "PYTHON_CMD=python"
if exist "..\.venv\Scripts\python.exe" (
    set "PYTHON_CMD=..\.venv\Scripts\python.exe"
)

"%PYTHON_CMD%" --version >nul 2>&1
if errorlevel 1 (
    echo ✗ ERREUR: Python n'est pas installé
    echo   Installez Python 3.9+ depuis https://www.python.org
    pause
    exit /b 1
)

echo ✓ Python détecté
echo.

REM Mode runtime editable (recommandé pour voir immédiatement les changements)
if exist "run_internal.py" (
    echo Lancement en mode runtime editable...
    "%PYTHON_CMD%" "run_internal.py"
    set "RET=%ERRORLEVEL%"
    if "%RET%"=="0" exit /b 0
    echo.
    echo [INFO] Echec du mode editable, fallback vers executable packagé...
)

REM Fallback: version packagée
if exist "dist\AuditPro\AuditPro.exe" (
    echo Lancement de la version packagée...
    start "" "dist\AuditPro\AuditPro.exe"
    exit /b 0
)

echo.
echo ✗ ERREUR lors du démarrage
echo   Aucun mode de lancement disponible.
pause
exit /b 1
