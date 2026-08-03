# Outil BIM — test local puis déploiement Firebase + Railway + GitHub

## Architecture retenue

- `frontend/` : interface statique HTML/CSS/JavaScript, destinée à Firebase Hosting.
- `backend/` : API Flask et moteur Python, destinée à Railway.
- `backend/outil_bim_v1.py` : moteur existant, adapté pour accepter les réponses du navigateur dans un fichier JSON.
- `backend/dashboard_template.py` : dashboard validé avec la nouvelle structure des exigences.
- `backend/reporting_v2.py` : génération du plan d’action HTML interactif. Le dashboard HTML centralise l’analyse et les preuves.
- `backend/Parametrage_Outil_BIM_V1.xlsx` : paramétrage de référence. Une copie est créée pour chaque analyse afin d'isoler l'historique.

## 1. Test local

### Prérequis

- Python 3.11 ou 3.12
- Une connexion Internet pour l'installation initiale des bibliothèques Python

### Terminal 1 — backend Python

#### Windows PowerShell

```powershell
cd backend
py -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
python app.py
```

#### macOS / Linux

```bash
cd backend
python3 -m venv .venv
source .venv/bin/activate
python -m pip install -r requirements.txt
python app.py
```

Le backend doit répondre sur :

```text
http://127.0.0.1:8000/health
```

Résultat attendu :

```json
{"status":"ok"}
```

### Terminal 2 — frontend local

Depuis le dossier racine `BIM_web_local` :

```bash
python -m http.server 5173 --directory frontend
```

Ouvrir ensuite :

```text
http://127.0.0.1:5173
```

### Ordre de test recommandé

1. Charger une convention BIM et éventuellement un CCTP.
2. Cocher **Test rapide sans auto-évaluation**.
3. Suivre la progression réelle sur la page principale.
4. Vérifier l’ouverture du dashboard et du plan d’action HTML.
5. Refaire le test sans cocher cette option, répondre au questionnaire et contrôler l’historique.

Les résultats temporaires se trouvent dans :

```text
backend/runtime/jobs/<identifiant_du_test>/
```

Chaque dossier contient notamment :

- `Rapport_BIM.html` : dashboard interactif centralisant l’analyse, les preuves et le glossaire ;
- `Plan_Actions_Offre_BIM.html` : plan d’action interactif ;
- `status.json` : progression et état du traitement ;
- `job.json` : informations utilisées par l’historique ;
- `analyse.log` ;
- la copie du fichier Excel utilisée pour le test.

Les PDF de synthèse, d’annexe et de plan d’action ne sont plus générés. Les analyses sont conservées par défaut pendant 30 jours afin d’alimenter l’historique de la page principale.

## 2. Mise sur GitHub

Créer un dépôt contenant l'ensemble du dossier :

```text
BIM_web_local/
├── backend/
├── frontend/
├── firebase.json
├── .firebaserc.example
└── .gitignore
```

Exemple :

```bash
git init
git add .
git commit -m "Version web locale de l'outil BIM"
git branch -M main
git remote add origin <adresse-du-depot-github>
git push -u origin main
```

Ne pas publier de documents contractuels réels, de clés d'API ou de dossiers `runtime/jobs`.

## 3. Déploiement du backend sur Railway

1. Créer un projet Railway depuis le dépôt GitHub.
2. Choisir le dossier racine du service :

```text
backend
```

3. Railway utilisera le `Dockerfile` et le fichier `railway.toml`.
4. Ajouter les variables :

```text
ALLOWED_ORIGINS=https://votre-projet.web.app,https://votre-projet.firebaseapp.com
MAX_UPLOAD_MB=80
ANALYSIS_TIMEOUT_SECONDS=600
JOB_TTL_HOURS=720
```

5. Générer un domaine public dans les paramètres réseau de Railway.
6. Tester :

```text
https://votre-backend.up.railway.app/health
```

### Remarque sur le stockage Railway

La version fournie conserve les résultats sur le disque temporaire du conteneur. Ils peuvent disparaître lors d'un redémarrage ou d'un nouveau déploiement. Ce fonctionnement convient au test et au téléchargement immédiat.

Pour un archivage durable, prévoir ensuite Firebase Storage, Google Cloud Storage ou un stockage compatible S3.

## 4. Déploiement du frontend sur Firebase Hosting

Modifier d'abord :

```text
frontend/config.js
```

Remplacer :

```javascript
window.BIM_API_URL = "http://127.0.0.1:8000";
```

par le domaine Railway :

```javascript
window.BIM_API_URL = "https://votre-backend.up.railway.app";
```

Installer Firebase CLI, puis depuis la racine :

```bash
npm install -g firebase-tools
firebase login
firebase use --add
firebase deploy --only hosting
```

Le fichier `firebase.json` publie uniquement le dossier `frontend/`.

## 5. Points à sécuriser avant ouverture publique

La version actuelle est adaptée aux essais locaux et à une démonstration privée. Avant une ouverture publique :

- ajouter une authentification, par exemple Firebase Authentication ;
- vérifier le jeton Firebase côté backend ;
- ajouter une politique de confidentialité pour les documents déposés ;
- limiter le nombre d'analyses par utilisateur ;
- ajouter un stockage durable seulement si l'archivage est nécessaire ;
- supprimer automatiquement les documents après génération ;
- éviter toute clé secrète dans `frontend/config.js` ou dans GitHub.

## 6. Diagnostic rapide

### Le frontend affiche « Connexion au backend impossible »

- Vérifier que `python app.py` fonctionne.
- Ouvrir `http://127.0.0.1:8000/health`.
- Vérifier l'adresse dans `frontend/config.js`.

### Erreur CORS après déploiement Firebase

Ajouter les domaines Firebase exacts dans la variable Railway `ALLOWED_ORIGINS`, séparés par une virgule, puis redéployer le backend.

### L'analyse échoue

Consulter :

```text
backend/runtime/jobs/<job>/analyse.log
```

### L'analyse est interrompue sur Railway

Augmenter `ANALYSIS_TIMEOUT_SECONDS` et vérifier la mémoire du service. Le conteneur est volontairement configuré avec un seul worker Gunicorn, car l'analyse PDF, Matplotlib et la génération du rapport consomment de la mémoire.
