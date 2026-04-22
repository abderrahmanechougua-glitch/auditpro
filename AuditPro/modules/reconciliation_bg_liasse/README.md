# Réconciliation BG vs Liasse Fiscale

Ce module rapproche automatiquement la **Balance Générale** et la **Liasse Fiscale** :

- BG : Excel (`.xlsx`, `.xls`, `.xlsm`)
- Liasse : Excel (`.xlsx`, `.xls`) ou PDF (`.pdf`)

## Fonctionnalités

- Détection automatique de l'onglet et de la ligne d'en-tête BG
- Extraction des rubriques liasse (ACTIF, PASSIF, CPC)
- Mapping comptes → rubriques (plan comptable marocain)
- Calcul des écarts absolus et en pourcentage
- Rapport Excel avec code couleur :
  - Vert : OK
  - Jaune : Écart > 100 DH
  - Rouge : Écart > 10 000 DH

## Fichier généré

`Reconciliation_BG_Liasse_YYYYMMDD_HHMMSS.xlsx` avec :

- Feuille `Rapprochement`
- Feuille `Résumé`
