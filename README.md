# analyse-budget-patrimonial
Application publique d’analyse patrimoniale : formulaire client en étapes (identité, budget, profil financier), puis export automatique en PDF + Excel. Envoi des documents au cabinet par email (adresse de réception paramétrable).

## Lancement

```sh
python -m pip install -r requirements.txt
streamlit run app.py
```

## Saisie et calculs

- Un DataFrame propre à chaque session conserve identité, budget et profil lors des réexécutions et des retours entre étapes. Les données d'exemple du modèle ne préremplissent plus les champs client.
- Les salaires totaux sont calculés à partir des deux déclarants. Les loyers sont saisis hors charges, avant décote ; le total retenu applique 80 % une seule fois.
- Revenus retenus = salaires des déclarants + 80 % des loyers des deux biens.
- Reste à vivre = revenus retenus - charges fixes - essence - impôts - emprunt RP/loyer - autres crédits.
- Épargne possible = reste à vivre - budget plaisirs/voyages.
- Mensualité restante = max(0, 35 % des revenus retenus - emprunt RP/loyer - autres crédits).
- La réponse « Crédits — souscripteur ? » est descriptive : elle ne masque ni n'annule les mensualités renseignées. L'emprunt RP et le loyer restent regroupés dans le calcul simplifié existant. Il s'agit d'une simulation, pas d'une décision de financement.

La persistance concerne la session Streamlit, sans sauvegarde durable : un rechargement créant une nouvelle session, la fermeture de session ou un redémarrage du serveur peut effacer la saisie. Exporter les documents avant de quitter.

## Documents

Le PDF et l'Excel utilisent la même extraction des données conservées et les mêmes calculs. Ils contiennent tous les champs du formulaire, les 13 réponses du profil et les six indicateurs. Les autres crédits et le budget plaisirs/voyages figurent aussi dans le PDF.

L'Excel est désormais une feuille « Synthèse client », avec libellés explicites et valeurs numériques, remplaçant l'ancien export de la grille du modèle. Il constitue un instantané des résultats : il ne contient pas de formules à recalculer après modification dans Excel. `template.xlsx` reste inchangé et sert uniquement de structure interne au chargement. Sa mise en page et ses rubriques non proposées dans le formulaire ne sont pas exportées.

## Tests

```sh
python -m pip install -r requirements-dev.txt
python -m pytest -q
```

Les versions de test sont fixées dans `requirements-dev.txt` ; les dépendances de production restent celles de `requirements.txt`. Le workflow GitHub Actions exécute la même commande sous Python 3.12.

Les tests Streamlit `AppTest` couvrent le parcours des quatre étapes, les retours, les modifications et effacements, les réexécutions, l'isolation de deux sessions, le chargement unique du modèle, chaque entrée budgétaire, le budget nul, les centimes et le surendettement. Les vrais octets proposés au téléchargement sont rouverts et vérifiés avec `pypdf` et `openpyxl`, champ par champ pour un client fictif. L'envoi SMTP est simulé et ses pièces jointes sont comparées aux téléchargements ; aucun email réel n'est envoyé.

Ces tests n'exercent pas un navigateur réel ni un déploiement distant. Ils ne valident pas la délivrabilité SMTP ou le fonctionnement dans Microsoft Excel.

## Version iPhone avec stockage serveur

Version privée publiée : https://cl-conseils-patrimoine.groupe-fi-6456.chatgpt.site/

Connexion avec le compte ChatGPT propriétaire. Cette version web reprend les quatre étapes, les calculs et les exports, et ajoute l’enregistrement automatique en base D1, la liste et la reprise des dossiers, la suppression et la protection contre l’écrasement entre deux appareils. Les données restent rattachées à l’identité du compte connecté. Les dossiers clients ne sont jamais écrits dans GitHub ni dans localStorage.

Les sources complètes de la version déployée sont dans `web/`. Commandes : `cd web`, `npm ci`, `node --test tests/*.test.mjs`, `node node_modules/typescript/bin/tsc --noEmit`. Node 24 est utilisé pour les tests SQL. La migration de création de la table est dans `web/drizzle/`. Le stockage est provisionné par l’hébergement lors de la publication.

L’interface Streamlit Python corrigée et la version web ont des stockages distincts : la persistance durable est activée sur la version web ci-dessus. L’envoi SMTP demeure dans Streamlit ; la version web propose le téléchargement PDF/Excel. Les tests web comparent six cas de calcul à des résultats produits par Python et contrôlent l’isolation des propriétaires et les conflits de version.
