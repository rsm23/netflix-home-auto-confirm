# confirm-netflix-house

Une petite appli en Python qui surveille un compte Gmail et ouvre le lien de mise à jour du foyer Netflix quand un email correspondant arrive.

## Fonctionnalités
- Se connecte à Gmail via OAuth2 et l'API Gmail.
- Filtre les messages dont l'expéditeur est `info@account.netflix.com` et dont l'objet contient `comment mettre à jour votre foyer Netflix`.
- Parcourt le contenu pour trouver un lien contenant `/update-primary-location` et l'ouvre dans le navigateur par défaut.
- Peut fonctionner en mode "polling" (vérification périodique) simple et fiable.
- Après clic sur le bouton Netflix, extrait le texte du champ "Demande effectuée par" dans l'email et l'enregistre dans un fichier `.txt` dans un dossier de sortie configurable.

## Prérequis
- Python 3.9+ recommandé.
- Un projet Google Cloud avec l'API Gmail activée.
- Un fichier `credentials.json` (Client OAuth 2.0) placé à la racine du projet.

## Installation

```pwsh
# Depuis la racine du projet
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -U pip
pip install -r requirements.txt
# Installer les navigateurs Playwright (Edge/Chrome/Chromium)
python -m playwright install --with-deps chromium
```

> Important: assurez-vous que votre venv est bien activé (le prompt affiche `(venv)` ou similaire) avant d'exécuter les commandes `pip install` et `python -m playwright install`.

## Configuration
1. Activez l'API Gmail dans Google Cloud Console et créez des identifiants OAuth 2.0 (application de bureau).
2. Téléchargez le fichier `credentials.json` et placez-le à la racine du dépôt.
3. Au premier lancement, une fenêtre de consentement s'ouvrira pour autoriser l'application à lire vos emails.
4. Le jeton persistant sera stocké dans `token.json` (crée automatiquement).

## Utilisation

- Exécuter une vérification unique (récents et nouveaux messages) :

```pwsh
python -m src.main once
```

- Lancer une surveillance en boucle (polling toutes les 60s par défaut) :

```pwsh
python -m src.main watch --interval 60
```

- Options utiles :

```pwsh
python -m src.main watch --interval 30 --open-once
```

Par défaut, l'application ignore les emails reçus avant son démarrage. Dès qu'un lien est cliqué, l'horodatage de référence est mis à jour, ce qui évite de retraiter d'anciens emails.

- Pour contrôler manuellement le point de départ en mode `once`, vous pouvez préciser un timestamp (ms epoch) avec `--since-epoch-ms`:

```pwsh
python -m src.main once --since-epoch-ms 1730726400000 --auto-click
```

### Mode auto-click (ouverture + clic + fermeture)

- Une fois Playwright installé (voir Installation), vous pouvez demander au script d’ouvrir la page Netflix, cliquer sur le bouton, puis fermer après 10s (configurable):

```pwsh
python -m src.main once --auto-click --close-delay 10
```

- En mode watch:

```pwsh
python -m src.main watch --auto-click --close-delay 10
```

- Pour réutiliser votre session Netflix existante (recommandé), le script essaie d’ouvrir Edge (ou Chrome) avec le profil utilisateur par défaut (Windows). Vous pouvez préciser:
  - `PLAYWRIGHT_CHANNEL` (ex: `msedge`, `chrome`) 
  - `BROWSER_USER_DATA_DIR` chemin vers le profil à réutiliser.

## Application de barre système (GUI minimal)

Vous pouvez lancer un petit utilitaire en zone de notification qui surveille en arrière-plan:

```pwsh
python -m src.tray_app
```

Un menu s’affiche sur l’icône: Start / Stop / Quit.

Dans le menu "Settings", vous pouvez régler:
- `Interval (s)` : l’intervalle de polling.
- `Close Delay (s)` : le délai de fermeture de l’onglet après le clic Playwright.
- `Output Folder` : le dossier où seront enregistrés les fichiers `.txt` contenant le texte "Demande effectuée par". Après changement, arrêtez puis redémarrez le watcher (Stop puis Start) pour appliquer.

## Construire un .exe (Windows)

Nous recommandons PyInstaller pour packager un .exe autonome:

```pwsh
.\.venv\Scripts\Activate.ps1
pip install pyinstaller
pyinstaller --noconsole --name confirm-netflix-house --add-data "credentials.json;." --hidden-import playwright --hidden-import bs4 --hidden-import googleapiclient --hidden-import google.oauth2 --hidden-import google_auth_oauthlib --hidden-import html5lib --hidden-import pystray --hidden-import PIL --collect-all playwright --collect-all bs4 --collect-all googleapiclient --collect-all google --collect-all google_auth_oauthlib --collect-all html5lib --collect-all pystray --collect-all PIL -m src.tray_app
```

Notes:
- `--noconsole` masque la console (GUI tray uniquement).
- `--add-data "credentials.json;."` embarque vos credentials si vous le souhaitez; sinon, placez `credentials.json` à côté du .exe.
- `--collect-all` rassemble les ressources/données nécessaires (Playwright, bs4, Google libs, etc.).
- Playwright: vous pouvez aussi installer les navigateurs sur la machine cible (au premier lancement si nécessaire).

Le binaire sera dans `dist/confirm-netflix-house/confirm-netflix-house.exe`.

## Exécuter en arrière-plan au démarrage de Windows

Deux approches pratiques:

1) Planificateur de tâches (Task Scheduler)
  - Créez une tâche basique qui lance le .exe “À l’ouverture de session”.
  - Configurez “Exécuter si l’utilisateur est connecté” et “Exécuter avec les autorisations maximales” si besoin.

2) Dossier Démarrage
  - Ouvrez `shell:startup` dans l’explorateur.
  - Collez un raccourci vers `confirm-netflix-house.exe`.

Variables utiles pour le mode service/tray:
- `POLL_INTERVAL` (ex: `60`) pour l’intervalle de scan.
- `GMAIL_QUERY` si vous souhaitez surcharger le filtre Gmail.
- `PLAYWRIGHT_CHANNEL` (ex: `msedge`, `chrome`).
- `BROWSER_USER_DATA_DIR` chemin du profil navigateur à réutiliser.

- Variables d'environnement (optionnelles) :
  - `GMAIL_QUERY` : surcharge la requête Gmail (par défaut : `from:info@account.netflix.com subject:"comment mettre à jour votre foyer Netflix"`).
  - `POLL_INTERVAL` : intervalle de polling par défaut en secondes.
  - `OUTPUT_DIR` : dossier de sortie par défaut pour enregistrer les fichiers `.txt` (peut aussi être défini via `--output-dir` en CLI ou dans la GUI Settings).
  - Note: par défaut, la requête inclut `is:unread` et les messages cliqués sont marqués comme lus.

### Export du texte "Demande effectuée par"
Lorsqu’un clic réussi est effectué sur le bouton de confirmation Netflix via Playwright, l’application:
1) Analyse l’email pour trouver la cellule `<td>` contenant "Demande effectuée par".
2) Écrit un fichier texte nommé `requester_<timestamp>.txt` dans le dossier de sortie, contenant l’objet, l’ID du message et le texte extrait.

Vous pouvez configurer le dossier de sortie via:
- la GUI (`Settings` → `Output Folder`),
- l’argument CLI `--output-dir`,
- la variable d’environnement `OUTPUT_DIR`.

## Sécurité et respect
- Le script ouvre automatiquement un lien reçu par email; utilisez-le uniquement pour votre propre compte et en comprenant l'impact.
- Ne partagez jamais `credentials.json`/`token.json`.

## Dépannage
- Si l'authentification échoue, supprimez `token.json` et relancez pour reconsentir.
- Vérifiez que l'API Gmail est activée et que le type d'application est "Desktop".
- Les emails HTML peuvent être volumineux; si aucun lien n'est détecté, sauvegardez le corps pour inspection (voir `--debug`).

### Erreur: `redirect_uri_mismatch`
Cette erreur survient lorsque les identifiants OAuth ne correspondent pas au flux de redirection utilisé.

Solutions:
- Recommandé: créez un client OAuth de type "Application de bureau" (Desktop) et remplacez `credentials.json`.
- Alternative (client OAuth de type Web):
  1. Ajoutez une URI de redirection autorisée exacte, par exemple `http://localhost:8080/` dans la console GCP.
  2. Fixez le port local côté appli avant de lancer:
     ```pwsh
     $env:OAUTH_LOCAL_SERVER_PORT = "8080"
     python -m src.main once
     ```
  3. En cas de cache invalide, supprimez `token.json` puis relancez.

### Portée OAuth requise
L'application utilise désormais la portée `https://www.googleapis.com/auth/gmail.modify` pour pouvoir marquer les messages comme lus (`UNREAD` -> retiré). Au premier lancement après cette modification, supprimez `token.json` pour reconsentir avec la nouvelle portée:

```pwsh
Remove-Item -ErrorAction SilentlyContinue .\token.json
python -m src.main once
```

### Expiration et rafraîchissement du token
L'application gère automatiquement le rafraîchissement du token. Si le rafraîchissement échoue (token révoqué/changé), elle relance le flux OAuth pour reconsentir. En cas d'erreur liée à l'URI de redirection, suivez la section `redirect_uri_mismatch`.
