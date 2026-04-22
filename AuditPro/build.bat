@echo off
REM ═══════════════════════════════════════════════════════════════
REM  AuditPro — Build Windows
REM  Double-cliquer pour compiler l'application
REM ═══════════════════════════════════════════════════════════════

REM Se placer dans le dossier du script
cd /d "%~dp0"

REM Auto-déblocage (résout "Windows ne peut pas accéder...")
powershell -NoProfile -ExecutionPolicy Bypass -Command ^
  "Get-ChildItem -Path '%~dp0' -Recurse | Unblock-File -ErrorAction SilentlyContinue" >nul 2>&1

echo.
echo  ====================================================
echo   AuditPro — Build Windows
echo  ====================================================
echo.

REM ── Vérifier Python ─────────────────────────────────────────
python --version >nul 2>&1
if errorlevel 1 (
    echo.
    echo [ERREUR] Python non trouve.
    echo.
    echo Solutions :
    echo   1. Installez Python 3.10+ depuis https://python.org
    echo   2. Cochez "Add Python to PATH" lors de l'installation
    echo   3. Relancez ce script
    echo.
    pause & exit /b 1
)
for /f "tokens=*" %%v in ('python --version') do echo  Python detecte : %%v

REM ── Dépendances ─────────────────────────────────────────────
echo [1/4] Installation des dependances...
pip install -r requirements.txt --quiet
if errorlevel 1 (
    echo [WARN] pip install a signale des avertissements. Tentative de continuer...
)

REM ── Icône ───────────────────────────────────────────────────
echo [2/4] Generation de l'icone...
if not exist "resources" mkdir resources
if not exist "resources\AuditPro.ico" (
    python create_icon.py
    if errorlevel 1 echo [WARN] Icone non generee - build sans icone.
)

REM ── Nettoyage ────────────────────────────────────────────────
echo [3/4] Nettoyage + Compilation (2-5 min)...
if exist dist  rmdir /s /q dist  >nul 2>&1
if exist build rmdir /s /q build >nul 2>&1

REM ── PyInstaller ──────────────────────────────────────────────
python -m PyInstaller AuditPro.spec --noconfirm --clean
if errorlevel 1 (
    echo.
    echo [ERREUR] La compilation a echoue.
    echo.
    echo Causes frequentes :
    echo   - Antivirus bloque PyInstaller (ajouter une exception)
    echo   - Manque de RAM (fermer les autres applications)
    echo   - Chemin avec caracteres speciaux
    echo.
    pause & exit /b 1
)

REM ── Données ──────────────────────────────────────────────────
echo [4/4] Copie des donnees...
if exist data xcopy /E /I /Y data dist\AuditPro\data >nul 2>&1

echo.
echo  ====================================================
echo   BUILD TERMINE !
echo  ====================================================
echo   Executable : dist\AuditPro\AuditPro.exe
echo   Double-clic pour lancer.
echo.
echo   Pour un package ZIP distribuable :
echo   → Lancez dist_package.bat
echo  ====================================================
echo.
pause
