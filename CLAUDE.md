# Instructions pour Claude

## Approche projet
- TOUJOURS proposer un plan d'architecture AVANT de coder
- Demander les maquettes/wireframes AVANT d'implementer l'UI
- Identifier les contraintes techniques (export, auth, multi-user) EN AMONT
- Proposer une structure modulaire (1 fichier = 1 responsabilite, max 500 lignes)

## Developpement
- Creer des fichiers separes par domaine, pas un monolithe
- Pour les graphiques : definir la charte graphique (couleurs, tailles) une seule fois
- Pour l'export PDF : tester le rendu des le premier graphique, pas apres 20
- Ne pas iterer 10 fois sur le CSS : proposer 2-3 options et demander validation

## Communication
- Quand un probleme technique est identifie (ex: kaleido ne marche pas),
  le signaler immediatement et proposer des alternatives AVANT d'essayer 5 approches
- Regrouper les modifications liees (ex: tous les graphiques d'une section en une fois)
- Demander confirmation sur les choix visuels AVANT d'implementer

## Stack technique
- Framework : Streamlit (Python)
- Calculs : pandas, numpy
- Graphiques : Plotly
- Stockage : fichiers JSON dans /plans/
- Export : HTML + impression navigateur (kaleido ne fonctionne pas sur ce systeme)
- Auth : codes d'acces en dur (ArgenteauEdit! / ArgenteauVisu!)
- Deploiement prevu : Streamlit Community Cloud
