"""
╔══════════════════════════════════════════════════════════════════╗
║         GUIDE D'INTÉGRATION — Comment brancher vos scripts      ║
╠══════════════════════════════════════════════════════════════════╣
║  Ce fichier montre le pattern EXACT pour chaque module.         ║
║  Copiez le template, adaptez les 3 zones marquées ★             ║
╚══════════════════════════════════════════════════════════════════╝

═══════════════════════════════════════════════════════════════════
 PRINCIPE GÉNÉRAL
═══════════════════════════════════════════════════════════════════

Votre script existant (ex: centralisation_tva.py) reste INCHANGÉ.
Vous créez UNIQUEMENT un fichier module.py (~40 lignes) qui fait le pont.

L'astuce : dans execute(), vous importez et appelez votre script.

═══════════════════════════════════════════════════════════════════
 ÉTAPES POUR CHAQUE MODULE
═══════════════════════════════════════════════════════════════════

1. Copiez votre script.py dans modules/<nom>/
2. Créez module.py dans le même dossier (voir templates ci-dessous)
3. Relancez l'app → le module apparaît automatiquement

═══════════════════════════════════════════════════════════════════
 TEMPLATE 1 : Centralisation TVA
═══════════════════════════════════════════════════════════════════

Fichiers :
  modules/tva/
  ├── module.py              ← le wrapper (ci-dessous)
  └── centralisation_tva.py  ← VOTRE script (inchangé)

Dans module.py, remplacez la méthode execute() par :

    def execute(self, inputs, output_dir, progress_callback=None):
        import sys, os
        from pathlib import Path
        
        # ★ ZONE 1 : Récupérer les inputs
        fichier = inputs["declarations"]
        exercice = inputs.get("exercice", "2025")
        
        if progress_callback:
            progress_callback(10, "Lecture des déclarations...")
        
        # ★ ZONE 2 : Appeler votre script
        # Option A — si votre script a une fonction main/run :
        from modules.tva.centralisation_tva import run_centralisation
        output_path = os.path.join(output_dir, f"Centralisation_TVA_{exercice}.xlsx")
        result_df = run_centralisation(fichier, output_path)
        
        # Option B — si votre script est un script linéaire sans fonction :
        # Ajoutez une fonction wrapper dans votre script, ex:
        #   def run(input_path, output_path): ...
        # Puis appelez-la ici.
        
        if progress_callback:
            progress_callback(100, "Terminé !")
        
        # ★ ZONE 3 : Retourner le résultat
        return ModuleResult(
            success=True,
            output_path=output_path,
            message=f"Centralisation TVA {exercice} générée",
            stats={
                "Exercice": exercice,
                "Lignes": len(result_df) if result_df is not None else 0,
            }
        )


═══════════════════════════════════════════════════════════════════
 TEMPLATE 2 : Centralisation CNSS
═══════════════════════════════════════════════════════════════════

Même pattern que TVA. Dans execute() :

    def execute(self, inputs, output_dir, progress_callback=None):
        from modules.cnss.centralisation_cnss import run_cnss
        
        fichier = inputs["bordereaux"]
        exercice = inputs.get("exercice", "2025")
        output_path = os.path.join(output_dir, f"Centralisation_CNSS_{exercice}.xlsx")
        
        if progress_callback:
            progress_callback(20, "Traitement des bordereaux...")
        
        result = run_cnss(fichier, output_path)
        
        if progress_callback:
            progress_callback(100, "Terminé !")
        
        return ModuleResult(success=True, output_path=output_path, ...)


═══════════════════════════════════════════════════════════════════
 TEMPLATE 3 : Lettrage Grand Livre
═══════════════════════════════════════════════════════════════════

    def execute(self, inputs, output_dir, progress_callback=None):
        from modules.lettrage.lettrage_gl import run_lettrage
        
        gl_path = inputs["grand_livre"]
        tolerance = inputs.get("param_tolerance", 1)
        type_tiers = inputs.get("param_type_tiers", "Tous")
        output_path = os.path.join(output_dir, f"Lettrage_GL.xlsx")
        
        if progress_callback:
            progress_callback(10, "Lecture du grand livre...")
        
        stats = run_lettrage(gl_path, output_path,
                             tolerance=tolerance,
                             type_tiers=type_tiers)
        
        if progress_callback:
            progress_callback(100, "Terminé !")
        
        return ModuleResult(
            success=True,
            output_path=output_path,
            stats=stats  # ex: {"Lettrés": 450, "Non lettrés": 23, "Écart total": 1.50}
        )


═══════════════════════════════════════════════════════════════════
 TEMPLATE 4 : Retraitement Comptable (main.py v7)
═══════════════════════════════════════════════════════════════════

Celui-ci est plus complexe car main.py v7 crée sa propre structure
input/output. L'astuce : passer les chemins via des variables.

    def execute(self, inputs, output_dir, progress_callback=None):
        import importlib, sys
        from pathlib import Path
        
        fichier = inputs["fichier_comptable"]
        
        if progress_callback:
            progress_callback(10, "Détection du format...")
        
        # Votre main.py attend des fichiers dans input/
        # On copie le fichier dans un dossier temporaire structuré
        import shutil, tempfile
        with tempfile.TemporaryDirectory() as tmpdir:
            # Créer la structure attendue par main.py
            input_dir = Path(tmpdir) / "input" / "AUTO"
            input_dir.mkdir(parents=True)
            shutil.copy2(fichier, input_dir)
            
            out_dir = Path(tmpdir) / "output"
            out_dir.mkdir()
            
            if progress_callback:
                progress_callback(30, "Retraitement en cours...")
            
            # Appeler votre script
            # Option : import et appel de la fonction principale
            from modules.retraitement.main import process_files
            process_files(str(input_dir.parent), str(out_dir))
            
            if progress_callback:
                progress_callback(80, "Copie des résultats...")
            
            # Copier les outputs vers le vrai dossier de sortie
            for f in out_dir.glob("*.xlsx"):
                dest = Path(output_dir) / f.name
                shutil.copy2(f, dest)
                final_output = str(dest)
        
        if progress_callback:
            progress_callback(100, "Terminé !")
        
        return ModuleResult(success=True, output_path=final_output)


═══════════════════════════════════════════════════════════════════
 TEMPLATE 5 : SRM Generator
═══════════════════════════════════════════════════════════════════

Output = Word (.docx), pas Excel. Même pattern.

    def execute(self, inputs, output_dir, progress_callback=None):
        from modules.srm_generator.srmgen import generate_srm
        
        excel_path = inputs["tableau_srm"]
        output_path = os.path.join(output_dir, "SRM_generated.docx")
        
        if progress_callback:
            progress_callback(20, "Analyse du tableau SRM...")
        
        generate_srm(excel_path, output_path)
        
        if progress_callback:
            progress_callback(100, "Terminé !")
        
        return ModuleResult(success=True, output_path=output_path)


═══════════════════════════════════════════════════════════════════
 TEMPLATE 6 : Circularisation des Tiers
═══════════════════════════════════════════════════════════════════

Multi-output (Excel sélection + Word lettres). Retourner le dossier.

    def execute(self, inputs, output_dir, progress_callback=None):
        from modules.circularisation.script import run_circularisation
        
        balance = inputs["balance_aux"]
        societe = inputs["societe"]
        exercice = inputs.get("exercice", "2025")
        couverture = inputs.get("param_couverture", 90) / 100
        
        if progress_callback:
            progress_callback(10, "Parsing de la balance...")
        
        result = run_circularisation(
            balance_path=balance,
            societe=societe,
            exercice=exercice,
            couverture_cible=couverture,
            output_dir=output_dir,
            progress_fn=progress_callback
        )
        
        return ModuleResult(
            success=True,
            output_path=output_dir,  # Dossier avec tous les outputs
            message=f"Circularisation générée pour {societe}",
            stats=result.get("stats", {})
        )


═══════════════════════════════════════════════════════════════════
 ADAPTATION DE VOS SCRIPTS EXISTANTS
═══════════════════════════════════════════════════════════════════

Si votre script est un script linéaire (pas de fonction main), 
il faut l'adapter MINIMALEMENT :

AVANT (script linéaire) :
    # centralisation_tva.py
    import pandas as pd
    df = pd.read_excel("input/declarations.xlsx")
    # ... traitement ...
    df.to_excel("output/result.xlsx")

APRÈS (avec une fonction wrapper) :
    # centralisation_tva.py
    import pandas as pd
    
    def run_centralisation(input_path, output_path):
        df = pd.read_excel(input_path)
        # ... même traitement qu'avant ...
        df.to_excel(output_path, index=False)
        return df
    
    # Pour garder la compatibilité standalone :
    if __name__ == "__main__":
        run_centralisation("input/declarations.xlsx", "output/result.xlsx")

Cette modification est MINIMALE : vous ajoutez une fonction et un 
if __name__ == "__main__". Le reste du code ne change PAS.

═══════════════════════════════════════════════════════════════════
 GESTION DES print() → progress_callback()
═══════════════════════════════════════════════════════════════════

Si votre script fait des print() pour montrer la progression,
vous pouvez les remplacer par des appels au callback :

AVANT :
    print("Étape 1 : Lecture du fichier...")
    print("Étape 2 : Calcul des totaux...")

APRÈS :
    if progress_callback:
        progress_callback(25, "Lecture du fichier...")
    # ... traitement ...
    if progress_callback:
        progress_callback(75, "Calcul des totaux...")

Ou, pour éviter de modifier votre script, intercepter les print :

    import io, contextlib
    buffer = io.StringIO()
    with contextlib.redirect_stdout(buffer):
        run_mon_script(...)
    # Les print sont silencieux, l'UI utilise la progress bar

═══════════════════════════════════════════════════════════════════
"""
