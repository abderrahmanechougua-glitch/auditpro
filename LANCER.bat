@echo off
REM Canonical launcher for this workspace: starts AuditPro_SHARE
cd /d "%~dp0"

if exist "AuditPro_SHARE\LANCER.bat" (
    pushd "AuditPro_SHARE"
    call "LANCER.bat"
    set "RET=%ERRORLEVEL%"
    popd
    exit /b %RET%
)

echo [ERREUR] AuditPro_SHARE\LANCER.bat introuvable.
echo Verifiez que le dossier AuditPro_SHARE existe ici.
pause
exit /b 1
