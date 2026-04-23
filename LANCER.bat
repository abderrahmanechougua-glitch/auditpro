@echo off
REM DEPRECATED - Use AuditPro_SHARE\LANCER.bat directly
REM This file is kept for backwards compatibility only
echo.
echo [AVERTISSEMENT] Cette version de LANCER.bat est obsolete.
echo Utilisez plutot: AuditPro_SHARE\LANCER.bat
echo.
cd /d "%~dp0"

if exist "AuditPro_SHARE\LANCER.bat" (
    call "AuditPro_SHARE\LANCER.bat" %*
    exit /b %ERRORLEVEL%
)

echo [ERREUR] AuditPro_SHARE\LANCER.bat introuvable.
pause
exit /b 1
