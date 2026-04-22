@echo off
REM ═══════════════════════════════════════════════════════════════
REM  dist_package.bat — Build + ZIP prêt à partager
REM  Résultat : AuditPro_v1.0.0_Windows.zip (partageable à n'importe qui)
REM  Usage    : Double-cliquer (ou clic droit → Admin recommandé)
REM ═══════════════════════════════════════════════════════════════

cd /d "%~dp0"

REM Auto-déblocage silencieux
powershell -NoProfile -ExecutionPolicy Bypass -Command ^
  "Get-ChildItem -Path '%~dp0' -Recurse | Unblock-File -ErrorAction SilentlyContinue" >nul 2>&1

echo.
echo  ====================================================
echo   AuditPro — Build + Package de distribution
echo  ====================================================
echo.

REM ── Étape 1 : Vérifier Python ───────────────────────────────
python --version >nul 2>&1
if errorlevel 1 (
    echo [ERREUR] Python non trouve dans le PATH.
    echo Installez Python 3.10+ depuis https://python.org
    pause & exit /b 1
)
for /f "tokens=*" %%v in ('python --version') do echo  Python : %%v

REM ── Étape 2 : Installer les dépendances ─────────────────────
echo.
echo [1/5] Installation des dependances...
pip install -r requirements.txt --quiet
if errorlevel 1 (
    echo [ERREUR] Echec pip install. Verifiez votre connexion.
    pause & exit /b 1
)
echo  OK

REM ── Étape 3 : Générer l'icône ────────────────────────────────
echo [2/5] Generation de l'icone...
if not exist "resources\AuditPro.ico" (
    python create_icon.py
)
echo  OK

REM ── Étape 4 : Nettoyer et compiler ──────────────────────────
echo [3/5] Nettoyage...
if exist dist  rmdir /s /q dist
if exist build rmdir /s /q build
echo  OK

echo [4/5] Compilation PyInstaller (2-5 minutes)...
python -m PyInstaller AuditPro.spec --noconfirm --clean
if errorlevel 1 (
    echo.
    echo [ERREUR] PyInstaller a echoue !
    echo Consultez les messages d'erreur ci-dessus.
    pause & exit /b 1
)

REM ── Copier les données ───────────────────────────────────────
if exist data xcopy /E /I /Y data dist\AuditPro\data >nul 2>&1
echo  OK - Compilation terminee

REM ── Étape 5 : Créer le ZIP de distribution ──────────────────
echo [5/5] Creation du package ZIP...
set ZIPNAME=AuditPro_v1.0.0_Windows
if exist "%ZIPNAME%.zip" del "%ZIPNAME%.zip"

powershell -NoProfile -ExecutionPolicy Bypass -Command ^
  "Compress-Archive -Path 'dist\AuditPro\*' -DestinationPath '%ZIPNAME%.zip' -Force; Write-Host 'ZIP cree.'"

if exist "%ZIPNAME%.zip" (
    for %%f in ("%ZIPNAME%.zip") do set ZIPSIZE=%%~zf
    echo.
    echo  ====================================================
    echo   BUILD ET PACKAGING TERMINES !
    echo  ====================================================
    echo.
    echo   Executable : dist\AuditPro\AuditPro.exe
    echo   Package ZIP : %ZIPNAME%.zip
    echo.
    echo   POUR PARTAGER :
    echo   Envoyez le fichier %ZIPNAME%.zip
    echo   Le destinataire dezippe et double-clique AuditPro.exe
    echo   Aucune installation de Python requise !
    echo.
) else (
    echo [ERREUR] Impossible de creer le ZIP.
    echo Partagez manuellement le dossier dist\AuditPro\
)

REM Optionnel : Compiler l'installateur Inno Setup s'il est installé
where iscc >nul 2>&1
if not errorlevel 1 (
    echo [BONUS] Inno Setup detecte — Creation de l'installateur...
    iscc setup_auditpro.iss /Q
    if exist "Output\AuditPro_Setup_v1.0.0.exe" (
        echo   Installateur : Output\AuditPro_Setup_v1.0.0.exe
    )
)

pause
