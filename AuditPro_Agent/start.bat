@echo off
echo ╔═══════════════════════════════════════════════════════════╗
echo ║         AuditPro AI Agent - Startup Script               ║
echo ╚═══════════════════════════════════════════════════════════╗
echo.

REM Check Python
python --version >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Python n'est pas installe ou pas dans le PATH
    pause
    exit /b 1
)
echo [OK] Python installe

REM Check Ollama
curl -s http://localhost:11434/api/version >nul 2>&1
if errorlevel 1 (
    echo [WARNING] Ollama ne repond pas sur localhost:11434
    echo          Lancer: ollama serve
    echo.
) else (
    echo [OK] Ollama est en cours d'execution
)

REM Check model
ollama list | findstr llama3.2 >nul 2>&1
if errorlevel 1 (
    echo [INFO] Telechargement du modele llama3.2...
    ollama pull llama3.2
) else (
    echo [OK] Modele llama3.2 disponible
)

REM Install dependencies
echo.
echo [INFO] Installation des dependances...
pip install -r requirements.txt -q
echo [OK] Dependances installees

REM Start server
echo.
echo [INFO] Demarrage du serveur API...
echo       URL: http://localhost:8000
echo       Docs: http://localhost:8000/docs
echo.
python server.py
