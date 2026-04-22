"""
AuditPro AI Agent - Enhanced Runner
Checks dependencies and starts the server with proper error handling.
"""
import os
import sys
import time
import subprocess
from pathlib import Path

# Colors for output
GREEN = "\033[92m"
RED = "\033[91m"
YELLOW = "\033[93m"
BLUE = "\033[94m"
RESET = "\033[0m"

def check_python():
    """Verify Python is available."""
    print(f"{BLUE}[*] Checking Python...{RESET}")
    print(f"{GREEN}[+] Python {sys.version.split()[0]}{RESET}")
    return True

def check_ollama():
    """Verify Ollama is running."""
    print(f"{BLUE}[*] Checking Ollama...{RESET}")
    try:
        import requests
        resp = requests.get("http://localhost:11434/api/version", timeout=5)
        if resp.status_code == 200:
            version = resp.json().get("version", "unknown")
            print(f"{GREEN}[+] Ollama v{version}{RESET}")
            return True
    except:
        pass
    print(f"{YELLOW}[!] Ollama not responding. Start it with: ollama serve{RESET}")
    return False

def check_model(model_name="llama3.2"):
    """Verify the LLM model is available."""
    print(f"{BLUE}[*] Checking model {model_name}...{RESET}")
    try:
        import requests
        resp = requests.get("http://localhost:11434/api/tags", timeout=5)
        models = [m["name"] for m in resp.json().get("models", [])]
        if any(model_name in m for m in models):
            print(f"{GREEN}[+] Model {model_name} available{RESET}")
            return True
        print(f"{YELLOW}[!] Model {model_name} not found. Downloading...{RESET}")
        subprocess.run(["ollama", "pull", model_name], check=True)
        print(f"{GREEN}[+] Model {model_name} downloaded{RESET}")
        return True
    except Exception as e:
        print(f"{RED}[-] Error checking model: {e}{RESET}")
        return False

def check_auditpro():
    """Verify AuditPro directory exists."""
    print(f"{BLUE}[*] Checking AuditPro...{RESET}")
    auditpro_dir = Path(__file__).parent / "AuditPro"
    if auditpro_dir.exists():
        print(f"{GREEN}[+] AuditPro found at {auditpro_dir}{RESET}")

        # Check core modules
        core_files = ["core/module_registry.py", "modules/base_module.py"]
        for f in core_files:
            if (auditpro_dir / f).exists():
                print(f"    [OK] {f}")
            else:
                print(f"    [!] Missing: {f}")
        return True
    print(f"{YELLOW}[!] AuditPro not found. Create symlink or copy folder.{RESET}")
    return False

def install_dependencies():
    """Install required Python packages."""
    print(f"{BLUE}[*] Installing dependencies...{RESET}")
    try:
        import fastapi
        import uvicorn
        import pandas
        import requests
        print(f"{GREEN}[+] All dependencies installed{RESET}")
        return True
    except ImportError as e:
        print(f"{YELLOW}[!] Installing missing packages...{RESET}")
        subprocess.run([sys.executable, "-m", "pip", "install", "-r", "requirements.txt"], check=True)
        print(f"{GREEN}[+] Dependencies installed{RESET}")
        return True

def start_server(host="0.0.0.0", port=8000):
    """Start the FastAPI server."""
    print()
    print(f"{GREEN}{'='*60}{RESET}")
    print(f"{GREEN}Starting AuditPro AI Agent Server{RESET}")
    print(f"{GREEN}{'='*60}{RESET}")
    print(f"  Host: {BLUE}{host}{RESET}")
    print(f"  Port: {BLUE}{port}{RESET}")
    print(f"  API:  {BLUE}http://localhost:{port}{RESET}")
    print(f"  Docs: {BLUE}http://localhost:{port}/docs{RESET}")
    print(f"{GREEN}{'='*60}{RESET}")
    print()

    import uvicorn
    from server import app
    uvicorn.run(app, host=host, port=port)

def main():
    """Main entry point."""
    print()
    print(f"{GREEN}╔═══════════════════════════════════════════════════════════╗{RESET}")
    print(f"{GREEN}║         AuditPro AI Agent - Startup Check                ║{RESET}")
    print(f"{GREEN}╚═══════════════════════════════════════════════════════════╝{RESET}")
    print()

    checks = [
        ("Python", check_python),
        ("Ollama", check_ollama),
        ("Model", check_model),
        ("AuditPro", check_auditpro),
        ("Dependencies", install_dependencies),
    ]

    results = []
    for name, check_func in checks:
        try:
            result = check_func()
            results.append((name, result))
        except Exception as e:
            print(f"{RED}[-] {name} check failed: {e}{RESET}")
            results.append((name, False))
        print()

    # Summary
    print(f"{BLUE}{'='*60}{RESET}")
    print("Summary:")
    for name, result in results:
        status = f"{GREEN}OK{RESET}" if result else f"{RED}FAIL{RESET}"
        print(f"  {name}: {status}")
    print(f"{BLUE}{'='*60}{RESET}")
    print()

    failed = [n for n, r in results if not r]
    if failed:
        print(f"{YELLOW}[!] Some checks failed: {', '.join(failed)}{RESET}")
        print(f"{YELLOW}    Server may not work correctly.{RESET}")
        response = input("Continue anyway? (y/n): ")
        if response.lower() != 'y':
            sys.exit(1)

    try:
        start_server()
    except KeyboardInterrupt:
        print(f"\n{YELLOW}[!] Server stopped by user{RESET}")
    except Exception as e:
        print(f"{RED}[-] Server error: {e}{RESET}")
        sys.exit(1)

if __name__ == "__main__":
    main()
