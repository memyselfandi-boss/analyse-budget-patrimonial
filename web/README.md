# CL Conseils — Analyse patrimoniale mobile

Version web privée des quatre étapes du formulaire Streamlit. Les champs et six jeux de résultats attendus sont dérivés du module Python `patrimoine.py` du dépôt analyse-budget-patrimonial.

## Fonctions
- Formulaire responsive, indicateurs instantanés, client fictif intégré.
- Sauvegarde automatique après saisie d’un nom/prénom, bouton Enregistrer et liste des 500 derniers dossiers. Aucun dossier client n’est conservé dans localStorage.
- Stockage serveur D1, attaché à l’identité du compte connecté ; contrôle du propriétaire pour chaque lecture et écriture.
- Contrôle de version contre l’écrasement entre deux sessions, possibilité de créer une copie en cas de conflit.
- Téléchargements PDF et Excel, suppression confirmée et contrôle de version à la suppression.
- Accès hébergé réservé au propriétaire du Site. Ne pas désactiver cette protection.

## Développement
Node 22.13+ pour le site, Node 24 pour les tests utilisant node:sqlite.
`npm ci` ; `node --test tests/*.test.mjs` ; `node node_modules/typescript/bin/tsc --noEmit`.

La configuration et les scripts de déploiement sont fournis par l’environnement Sites. Le manifeste `.openai/hosting.json` déclare DB pour D1. Les migrations SQL sont dans drizzle/ et sont appliquées par la publication. Ne pas modifier une migration déjà publiée.

## Limites
Le stockage de cette version web est indépendant des sessions de l’ancienne interface Streamlit. Les exports sont des instantanés, sans formules Excel. L’envoi SMTP de la version Streamlit n’est pas activé dans la version web. Les polices PDF standard couvrent les caractères occidentaux ; un caractère non pris en charge est remplacé par un point d’interrogation.
