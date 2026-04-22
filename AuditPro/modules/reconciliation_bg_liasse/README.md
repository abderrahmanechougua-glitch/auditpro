# Module Réconciliation BG vs Liasse

Ce module rapproche une Balance Générale (Excel) avec une Liasse Fiscale (Excel/PDF).

## Entrées
- `fichier_bg` : `.xlsx` ou `.xls`
- `fichier_liasse` : `.xlsx`, `.xls` ou `.pdf`

## Sorties
- Rapport Excel avec :
  - tableau de réconciliation (montants BG, montants liasse, écarts, statut)
  - mise en forme conditionnelle (vert/jaune/rouge)
  - feuille de résumé (statistiques globales)

## Capacités
- détection automatique de ligne d'en-tête
- extraction Excel et PDF (pdfplumber/tabula-py selon disponibilité)
- mapping compte→rubrique (plan comptable marocain simplifié)
- traitement batch (1↔N, N↔1 ou N↔N)
