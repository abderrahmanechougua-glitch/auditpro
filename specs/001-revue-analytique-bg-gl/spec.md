# Feature Specification: Automatisation Revue Analytique BG-GL

**Feature Branch**: `001-revue-analytique-bg-gl`  
**Created**: 2026-04-24  
**Status**: Draft  
**Input**: User description: "Automatisation de la revue analytique depuis la Balance Generale (BG) et le Grand Livre (GL)"

## User Scenarios & Testing *(mandatory)*

### User Story 1 - Generer les feuilles analytiques par compte 4 chiffres (Priority: P1)

En tant qu'auditeur, je veux qu'une feuille soit creee automatiquement pour chaque compte a 4 chiffres present dans la BG, afin d'obtenir une revue analytique structuree sans preparation manuelle.

**Why this priority**: C'est la valeur coeur de la fonctionnalite: sans generation automatique des feuilles par compte, la revue analytique n'existe pas.

**Independent Test**: Peut etre teste independamment en important une BG contenant plusieurs comptes 4 chiffres et en verifiant qu'une feuille distincte est creee pour chacun, avec le bon nom de feuille.

**Acceptance Scenarios**:

1. **Given** une feuille BG valide contenant des comptes a 8 chiffres, **When** l'utilisateur lance la generation, **Then** le systeme cree une feuille par compte 4 chiffres detecte.
2. **Given** des sous-comptes commencant par le meme prefixe a 4 chiffres, **When** la feuille du compte parent est generee, **Then** seuls ces sous-comptes sont listes dans cette feuille.

---

### User Story 2 - Produire des tableaux dynamiques bases sur la feuille BG (Priority: P2)

En tant qu'auditeur, je veux que les valeurs affichees dans chaque feuille analytique proviennent de formules referencees sur la feuille BG, afin de garantir une mise a jour dynamique et d'eviter les saisies en dur.

**Why this priority**: La fiabilite de la revue depend de la tracabilite et de la synchronisation avec la source BG.

**Independent Test**: Peut etre teste en modifiant une valeur dans la feuille BG apres generation, puis en verifiant que les cellules dependantes dans les feuilles analytiques se mettent a jour via les formules.

**Acceptance Scenarios**:

1. **Given** une feuille analytique generee, **When** l'utilisateur inspecte les cellules de donnees, **Then** les cellules contiennent des formules pointant vers la feuille BG et non des valeurs statiques.
2. **Given** plusieurs sous-comptes dans une feuille, **When** la ligne de total du compte 4 chiffres est calculee, **Then** elle utilise une formule de somme des sous-comptes de la feuille.

---

### User Story 3 - Ajouter un commentaire analytique standard par feuille (Priority: P3)

En tant qu'auditeur, je veux disposer automatiquement d'un commentaire standard sous chaque tableau pour accelerer la redaction de l'analyse et homogeniser la restitution.

**Why this priority**: Le commentaire ameliore l'exploitation metier, mais la valeur primaire reste la structuration et le calcul des tableaux.

**Independent Test**: Peut etre teste en generant une feuille analytique et en verifiant la presence d'un commentaire standard avec le numero de compte, les soldes N et N-1, et la variation calculee.

**Acceptance Scenarios**:

1. **Given** une feuille de compte 4 chiffres generee, **When** le tableau est complete, **Then** un commentaire standard est ajoute sous le tableau avec les montants N, N-1 et la variation.
2. **Given** une variation positive ou negative, **When** le commentaire est genere, **Then** le signe de variation est explicite et coherent avec les soldes affiches.

---

### Edge Cases

- La feuille BG est absente du classeur: la generation doit s'arreter avec un message explicite et actionnable.
- La feuille BG existe mais une ou plusieurs colonnes obligatoires manquent (Compte, Intitule, Solde N, Solde N-1): aucune feuille analytique ne doit etre produite.
- Un compte est present mais ne suit pas le format attendu (longueur invalide, caracteres non numeriques): la ligne doit etre ignoree et signalee dans un resume d'execution.
- Un compte 4 chiffres n'a aucun sous-compte 8 chiffres associe: la feuille peut etre creee avec une section vide et un total a zero, pour conserver la completude de la revue.
- Le nom de feuille de compte existe deja: la regle de remplacement ou de regeneration doit etre appliquee de maniere coherente sans dupliquer la feuille.

## Requirements *(mandatory)*

### Functional Requirements

- **FR-001**: Le systeme MUST lire la feuille source nommee "BG" et verifier la presence des colonnes obligatoires: Compte, Intitule, Solde N, Solde N-1.
- **FR-002**: Le systeme MUST identifier les comptes agreges a 4 chiffres a partir des comptes presents dans la BG.
- **FR-003**: Le systeme MUST associer a chaque compte 4 chiffres tous les sous-comptes a 8 chiffres dont le prefixe correspond aux 4 premiers chiffres du compte parent.
- **FR-004**: Le systeme MUST creer une feuille distincte pour chaque compte 4 chiffres detecte, et le nom de la feuille MUST etre exactement le numero de ce compte.
- **FR-005**: Chaque feuille de compte MUST contenir un tableau avec les colonnes suivantes, dans cet ordre: Compte, Intitule, Ref, 31/12/N, 31/12/N-1.
- **FR-006**: Les cellules des lignes de sous-comptes MUST etre alimentees par des formules referencees sur la feuille BG; les valeurs en dur sont interdites pour ces donnees.
- **FR-007**: La ligne de total du compte 4 chiffres MUST etre presente en bas du tableau, visuellement differenciee (style gras), et calculee par formule de somme des sous-comptes.
- **FR-008**: Le systeme MUST inserer sous chaque tableau un commentaire analytique standard mentionnant le compte 4 chiffres, le solde N, le solde N-1 et la variation N vs N-1.
- **FR-009**: La variation mentionnee dans le commentaire MUST etre calculee automatiquement a partir des montants du tableau et afficher explicitement son signe.
- **FR-010**: Le systeme MUST traiter l'ensemble des comptes eligibles de la BG en une execution et produire un classeur final contenant toutes les feuilles analytiques attendues.
- **FR-011**: En cas de structure BG invalide (feuille ou colonnes manquantes), le systeme MUST interrompre la generation et fournir un message d'erreur comprehensible.

### Key Entities *(include if feature involves data)*

- **Ligne BG**: Enregistrement source de la balance generale, comprenant au minimum un numero de compte, un intitule et les soldes N et N-1.
- **Compte agrege (4 chiffres)**: Compte de regroupement representant la somme logique des sous-comptes partageant le meme prefixe a 4 chiffres.
- **Sous-compte (8 chiffres)**: Compte detaille rattache a un compte agrege via son prefixe a 4 chiffres.
- **Feuille analytique**: Feuille de sortie dediee a un compte agrege, contenant le tableau detaille, la ligne de total et le commentaire standard.
- **Commentaire analytique**: Phrase standardisee de restitution metier construite a partir des montants agreges N, N-1 et de leur variation.

## Success Criteria *(mandatory)*

### Measurable Outcomes

- **SC-001**: Pour un classeur BG valide, 100% des comptes agreges a 4 chiffres detectes disposent chacun d'une feuille analytique correspondante.
- **SC-002**: 100% des cellules de donnees des sous-comptes dans les feuilles analytiques sont basees sur des references dynamiques a la feuille BG (aucune valeur statique de recopie).
- **SC-003**: 100% des feuilles analytiques generees contiennent une ligne de total calculee automatiquement et un commentaire standard complet.
- **SC-004**: Sur un echantillon de controle metier, au moins 95% des comptes controles presentent une coherence exacte entre total du compte 4 chiffres et somme des sous-comptes affiches.
- **SC-005**: Un utilisateur metier peut passer d'une BG valide au classeur analytique complet en moins de 5 minutes sans intervention manuelle feuille par feuille.

## Assumptions

- La BG d'entree est fournie dans un classeur unique et accessible au moment de l'execution.
- Les numeros de compte exploitables pour cette feature sont representes sous forme numerique ou texte numerique normalisable.
- La colonne Ref de la sortie correspond a une valeur de reference attendue par le metier et reste incluse meme si la source ne precise pas son mode de calcul.
- La generation peut recreer ou mettre a jour des feuilles analytiques existantes du meme nom selon une politique uniforme definie par l'application.
- Le perimetre de cette version couvre la production des feuilles et du commentaire standard, sans interpretation avancee des causes de variation.
