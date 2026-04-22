"""
Client Python pour l'API AuditPro AI Agent.
Exemples d'utilisation de l'agent AI avec Ollama.
"""
import requests
from pathlib import Path
from typing import Optional, List, Dict, Any


class AuditProAgent:
    """Client pour interagir avec l'API AuditPro AI Agent."""

    def __init__(self, base_url: str = "http://localhost:8000"):
        self.base_url = base_url
        self.conversation: List[Dict[str, str]] = []

    def health_check(self) -> Dict[str, Any]:
        """Vérifie l'état du service."""
        resp = requests.get(f"{self.base_url}/health", timeout=10)
        resp.raise_for_status()
        return resp.json()

    def list_modules(self) -> List[Dict]:
        """Liste les modules disponibles."""
        resp = requests.get(f"{self.base_url}/modules", timeout=10)
        resp.raise_for_status()
        return resp.json()

    def chat(self, message: str, clear_history: bool = False) -> Dict:
        """
        Envoie un message en langage naturel à l'agent.

        Args:
            message: Votre demande en langage naturel
            clear_history: Si True, efface l'historique de conversation

        Returns:
            Dict avec response, module_used, output_path, data
        """
        if clear_history:
            self.conversation = []

        resp = requests.post(
            f"{self.base_url}/chat",
            json={"message": message, "conversation": self.conversation},
            timeout=120
        )
        resp.raise_for_status()
        result = resp.json()

        # Update conversation history
        self.conversation.append({"role": "user", "content": message})
        if not result.get("error"):
            self.conversation.append({"role": "assistant", "content": result["response"]})

        return result

    def execute_module(self, module_name: str, inputs: Dict, params: Dict = None) -> Dict:
        """
        Exécute un module spécifique.

        Args:
            module_name: Nom du module (ex: "TVA", "CNSS", "Lettrage")
            inputs: Inputs requis par le module
            params: Paramètres optionnels

        Returns:
            Dict avec success, message, output_path, stats
        """
        resp = requests.post(
            f"{self.base_url}/execute/{module_name}",
            json={
                "module_name": module_name,
                "inputs": inputs,
                "params": params or {}
            },
            timeout=300
        )
        resp.raise_for_status()
        return resp.json()

    def upload_file(self, file_path: str) -> Dict:
        """
        Upload un fichier pour analyse et détection automatique.

        Args:
            file_path: Chemin vers le fichier Excel/CSV

        Returns:
            Dict avec detected_modules, columns, recommendation
        """
        path = Path(file_path)
        if not path.exists():
            raise FileNotFoundError(f"Fichier non trouvé: {file_path}")

        with open(path, "rb") as f:
            resp = requests.post(
                f"{self.base_url}/upload",
                files={"file": (path.name, f)},
                timeout=60
            )
        resp.raise_for_status()
        return resp.json()

    def analyze_file(self, file_path: str, question: str = "Que contient ce fichier ?") -> Dict:
        """
        Analyse un fichier avec l'IA et répond à une question.

        Args:
            file_path: Chemin vers le fichier
            question: Question à poser sur les données

        Returns:
            Dict avec ai_analysis, rows, columns
        """
        resp = requests.post(
            f"{self.base_url}/analyze",
            params={"file_path": file_path, "question": question},
            timeout=120
        )
        resp.raise_for_status()
        return resp.json()

    def conversation_summary(self) -> str:
        """Retourne un résumé de la conversation actuelle."""
        if not self.conversation:
            return "Aucune conversation en cours."

        lines = []
        for msg in self.conversation[-5:]:  # Last 5 messages
            role = "Vous" if msg["role"] == "user" else "Agent"
            content = msg["content"][:80] + "..." if len(msg["content"]) > 80 else msg["content"]
            lines.append(f"  {role}: {content}")
        return "\n".join(lines)


# ── Exemples d'utilisation ────────────────────────────────
def demo():
    """Démo complète de l'agent AuditPro."""
    print("=" * 60)
    print("AuditPro AI Agent - Démo")
    print("=" * 60)

    agent = AuditProAgent()

    # 1. Health check
    print("\n[1] Vérification du service...")
    try:
        health = agent.health_check()
        print(f"    Status: {health['status']}")
        print(f"    Ollama: {'En ligne' if health['ollama'] else 'Hors ligne'}")
        print(f"    Modules: {health['modules_loaded']} chargés")
    except Exception as e:
        print(f"    ERREUR: {e}")
        print("    Assurez-vous que le serveur est démarré: python server.py")
        return

    # 2. Lister les modules
    print("\n[2] Modules disponibles:")
    try:
        modules = agent.list_modules()
        for m in modules[:5]:
            print(f"    • {m['name']} — {m['description'][:50]}...")
        if len(modules) > 5:
            print(f"    ... et {len(modules) - 5} autres")
    except Exception as e:
        print(f"    ERREUR: {e}")

    # 3. Chat simple
    print("\n[3] Test du chat (question simple)...")
    try:
        response = agent.chat("Bonjour, quels modules sont disponibles pour l'audit fiscal ?")
        print(f"    Réponse: {response['response'][:100]}...")
        if response.get('module_used'):
            print(f"    Module utilisé: {response['module_used']}")
    except Exception as e:
        print(f"    ERREUR: {e}")

    # 4. Chat avec intention
    print("\n[4] Test du chat (avec intention de module)...")
    try:
        response = agent.chat("Je veux centraliser les déclarations TVA de 2025")
        print(f"    Réponse: {response['response'][:100]}...")
        if response.get('module_used'):
            print(f"    Module utilisé: {response['module_used']}")
        if response.get('output_path'):
            print(f"    Sortie: {response['output_path']}")
    except Exception as e:
        print(f"    ERREUR: {e}")

    # 5. Upload de fichier (si fichier de test disponible)
    print("\n[5] Test d'upload de fichier...")
    test_file = Path(__file__).parent / "test_data.xlsx"
    if test_file.exists():
        try:
            result = agent.upload_file(str(test_file))
            print(f"    Modules détectés: {result.get('detected_modules', [])}")
        except Exception as e:
            print(f"    ERREUR: {e}")
    else:
        print("    (skipped - no test file)")

    print("\n" + "=" * 60)
    print("Démo terminée!")
    print("=" * 60)


if __name__ == "__main__":
    import sys

    if len(sys.argv) > 1 and sys.argv[1] == "--demo":
        demo()
    else:
        # Interactive mode
        agent = AuditProAgent()

        print("AuditPro AI Agent - Mode interactif")
        print("Tapez 'quit' pour quitter, 'reset' pour effacer l'historique")
        print()

        while True:
            try:
                user_input = input("> ").strip()
                if user_input.lower() in ['quit', 'exit', 'q']:
                    print("Au revoir!")
                    break
                if user_input.lower() == 'reset':
                    agent.conversation = []
                    print("Historique effacé.")
                    continue
                if user_input.lower() == 'history':
                    print(agent.conversation_summary())
                    continue

                response = agent.chat(user_input)
                print(f"\nAgent: {response['response']}")
                if response.get('module_used'):
                    print(f"[Module: {response['module_used']}]")
                if response.get('output_path'):
                    print(f"[Sortie: {response['output_path']}]")
                print()

            except KeyboardInterrupt:
                print("\nAu revoir!")
                break
            except Exception as e:
                print(f"Erreur: {e}")
