"""
Test script for AuditPro AI Agent API.
Run this after starting the server to verify all endpoints work.
"""
import requests
import time
import sys

BASE_URL = "http://localhost:8000"


def test_health():
    """Test health endpoint."""
    print("\n[TEST] Health check...")
    resp = requests.get(f"{BASE_URL}/health", timeout=10)
    print(f"  Status: {resp.status_code}")
    data = resp.json()
    print(f"  Server status: {data.get('status')}")
    print(f"  Ollama available: {data.get('ollama')}")
    print(f"  Modules loaded: {data.get('modules_loaded')}")
    return data.get('status') in ['healthy', 'degraded']


def test_root():
    """Test root endpoint."""
    print("\n[TEST] Root endpoint...")
    resp = requests.get(BASE_URL, timeout=10)
    print(f"  Status: {resp.status_code}")
    data = resp.json()
    print(f"  Service: {data.get('service')}")
    print(f"  Modules: {data.get('modules_disponibles')}")
    return resp.status_code == 200


def test_list_modules():
    """Test modules listing."""
    print("\n[TEST] List modules...")
    resp = requests.get(f"{BASE_URL}/modules", timeout=10)
    print(f"  Status: {resp.status_code}")
    if resp.status_code == 200:
        modules = resp.json()
        print(f"  Found {len(modules)} modules:")
        for m in modules[:5]:
            print(f"    - {m['name']}: {m['description'][:50]}...")
    return resp.status_code == 200


def test_chat_simple():
    """Test chat with a simple question."""
    print("\n[TEST] Chat (simple question)...")
    resp = requests.post(
        f"{BASE_URL}/chat",
        json={"message": "Bonjour, quels modules sont disponibles pour l'audit TVA ?"},
        timeout=60
    )
    print(f"  Status: {resp.status_code}")
    if resp.status_code == 200:
        data = resp.json()
        print(f"  Response: {data.get('response', '')[:100]}...")
    return resp.status_code == 200


def test_chat_with_intent():
    """Test chat with module trigger."""
    print("\n[TEST] Chat (with module intent)...")
    resp = requests.post(
        f"{BASE_URL}/chat",
        json={"message": "Je veux centraliser les déclarations TVA de 2025"},
        timeout=120
    )
    print(f"  Status: {resp.status_code}")
    if resp.status_code == 200:
        data = resp.json()
        print(f"  Response: {data.get('response', '')[:100]}...")
        if data.get('module_used'):
            print(f"  Module used: {data.get('module_used')}")
    return resp.status_code == 200


def test_upload_file(file_path=None):
    """Test file upload."""
    print("\n[TEST] File upload...")

    if not file_path:
        # Create a test file
        import tempfile
        import pandas as pd
        with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as f:
            df = pd.DataFrame({"Compte": ["411001", "401001"], "Libellé": ["Client A", "Fournisseur B"], "Montant": [1000, -500]})
            df.to_excel(f.name, index=False)
            file_path = f.name
            print(f"  Created test file: {file_path}")

    try:
        with open(file_path, 'rb') as f:
            files = {'file': (file_path.split('/')[-1], f, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')}
            resp = requests.post(f"{BASE_URL}/upload", files=files, timeout=30)

        print(f"  Status: {resp.status_code}")
        if resp.status_code == 200:
            data = resp.json()
            print(f"  Detected modules: {data.get('detected_modules')}")
        return resp.status_code == 200
    except FileNotFoundError:
        print(f"  File not found: {file_path}")
        return False


def run_all_tests():
    """Run all API tests."""
    print("=" * 60)
    print("AuditPro AI Agent - API Tests")
    print("=" * 60)

    # Check if server is running
    try:
        requests.get(BASE_URL, timeout=5)
    except:
        print(f"\n[ERROR] Server not running at {BASE_URL}")
        print("Start the server first: python server.py")
        sys.exit(1)

    tests = [
        ("Root", test_root),
        ("Health", test_health),
        ("List modules", test_list_modules),
        ("Chat simple", test_chat_simple),
        ("Chat with intent", test_chat_with_intent),
        ("File upload", test_upload_file),
    ]

    results = []
    for name, test_func in tests:
        try:
            result = test_func()
            results.append((name, result))
        except Exception as e:
            print(f"  ERROR: {e}")
            results.append((name, False))

    # Summary
    print("\n" + "=" * 60)
    print("Test Summary:")
    print("=" * 60)
    passed = sum(1 for _, r in results if r)
    total = len(results)

    for name, result in results:
        status = "PASS" if result else "FAIL"
        print(f"  [{status}] {name}")

    print(f"\nTotal: {passed}/{total} tests passed")

    if passed == total:
        print("\n[SUCCESS] All tests passed!")
    else:
        print(f"\n[WARNING] {total - passed} test(s) failed")

    return passed == total


if __name__ == "__main__":
    success = run_all_tests()
    sys.exit(0 if success else 1)
