@echo off
REM ═══════════════════════════════════════════════════════════════
REM  UNBLOCK — Débloque tous les fichiers AuditPro
REM  (requis quand les fichiers viennent d'Internet / ZIP)
REM  → Clic droit → "Exécuter en tant qu'administrateur"
REM ═══════════════════════════════════════════════════════════════

echo.
echo  ====================================================
echo   AuditPro — Deblocage des fichiers Windows
echo  ====================================================
echo.
echo  Deblocage en cours...

REM Méthode PowerShell — débloquer tous les fichiers récursivement
powershell -NoProfile -ExecutionPolicy Bypass -Command ^
  "Get-ChildItem -Path '%~dp0' -Recurse | Unblock-File -ErrorAction SilentlyContinue; Write-Host '[OK] Tous les fichiers debloque.'"

REM Vérification que ça a fonctionné
if errorlevel 1 (
    echo.
    echo [METHODE 2] Tentative alternative...
    for /r "%~dp0" %%f in (*) do (
        powershell -NoProfile -Command "Unblock-File -Path '%%f' -ErrorAction SilentlyContinue"
    )
)

echo.
echo  ====================================================
echo   Deblocage termine !
echo   Vous pouvez maintenant double-cliquer sur build.bat
echo  ====================================================
echo.
pause
