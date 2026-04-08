# Instructions pour Claude

## Approche projet
- TOUJOURS proposer un plan d'architecture AVANT de coder
- Pour tout nouveau plan financier : exiger le cahier des charges complet (voir memory/framework_plan_financier.md)
- Demander les maquettes/wireframes AVANT d'implementer l'UI
- Identifier les contraintes techniques (export, auth, multi-user) EN AMONT
- Proposer une structure modulaire (1 fichier = 1 responsabilite, max 500 lignes)

## Regles metier critiques (Belgique)
- Personnel : systeme belge avec pecule vacances (juillet) et 13e mois (decembre) au prorata
- Cash vs P&L : le P&L est lisse (masse/12), le cash suit le prorata reel
- TVA : paiement le mois suivant la cloture du trimestre (avril, juillet, octobre, janvier)
- Le modele travaille en HT : ajouter TVA nette au cash flow mensuel
- ISOC : charge en decembre, paiement en juin N+1
- Prets : cle par index (pas par nom) pour eviter les doublons
- Delais de paiement : par nature de fourniture (pas par service)
- Termes : "mensualite constante" (pas "annuite constante")

## Developpement
- Creer des fichiers separes par domaine, pas un monolithe
- Pour les graphiques : definir la charte graphique (couleurs, tailles) une seule fois
- Pour l'export PDF : tester le rendu des le premier graphique, pas apres 20
- Ne pas iterer 10 fois sur le CSS : proposer 2-3 options et demander validation
- Auto-save a chaque section critique (pas seulement en fin de page)
- Parametres communs : verrouilles par defaut, bouton Modifier pour editer
- Tableau d'amortissement par pret dans un expander

## Communication
- Quand un probleme technique est identifie (ex: kaleido ne marche pas),
  le signaler immediatement et proposer des alternatives AVANT d'essayer 5 approches
- Regrouper les modifications liees (ex: tous les graphiques d'une section en une fois)
- Demander confirmation sur les choix visuels AVANT d'implementer

## Stack technique
- Framework : Streamlit (Python)
- Calculs : pandas, numpy
- Graphiques : Plotly (texte en gras, taille 14px min, couleurs contrastees)
- Stockage : fichiers JSON dans /plans/
- Export : PDF (fpdf2), PPTX (python-pptx), HTML, CSV
- Auth : codes d'acces en dur (ArgenteauEdit! / ArgenteauVisu!)
- Deploiement : Streamlit Community Cloud (GitHub public, repo Romain-cmyk2/hotel-plan-financier)
