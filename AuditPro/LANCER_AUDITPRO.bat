@echo off
REM ═══════════════════════════════════════════════════════════════
REM  LANCER_AUDITPRO.bat — Lanceur de secours (si raccourci absent)
REM  Ne pas supprimer — utilisé par les raccourcis
REM ═══════════════════════════════════════════════════════════════
cd /d "%~dp0"

REM Déblocage silencieux
powershell -NoProfile -ExecutionPolicy Bypass -Command "Get-ChildItem '%~dp0' -Recurse | Unblock-File -EA SilentlyContinue" >nul 2>&1

REM Chercher pythonw.exe (chemin exact)
set PYTHONW=C:\Users\Abderrahmane.CHOUGUA\AppData\Local\Programs\Python\Python314\pythonw.exe

if not exist "%PYTHONW%" (
    REM Fallback : chercher dans le PATH
    for /f "tokens=*" %%p in ('where pythonw 2^>nul') do set PYTHONW=%%p
)

if not exist "%PYTHONW%" (
    echo Python introuvable. Installez Python depuis python.org
    pause
    exit /b 1
)

start "" "%PYTHONW%" "%~dp0main.py"
exit /b 0
