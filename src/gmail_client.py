import base64
import os
import sys
import logging
from typing import List, Optional, Tuple

from bs4 import BeautifulSoup
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# Scopes pour lire et modifier (marquer comme lu)
SCOPES = ["https://www.googleapis.com/auth/gmail.modify"]

# Évite une erreur quand Google renvoie un scope élargi (ex: readonly + modify)
# par rapport à celui demandé. Cela ne réduit pas la sécurité car l'ensemble
# retourné est un superset et la lib google gère les permissions effectives.
os.environ.setdefault("OAUTHLIB_RELAX_TOKEN_SCOPE", "1")

DEFAULT_QUERY = 'from:info@account.netflix.com subject:"comment mettre à jour votre foyer Netflix" is:unread'


def _resolve_credentials_path(preferred_path: Optional[str]) -> str:
    """Trouve un chemin valide pour credentials.json en contexte normal ou PyInstaller.

    Ordre de recherche:
    - variable d'environnement CREDENTIALS_PATH (si définie et existante)
    - chemin fourni (absolu ou relatif) si existant
    - répertoire courant (credentials.json)
    - répertoire de l'exécutable (si frozen)
    - dossier temporaire de PyInstaller (sys._MEIPASS) si présent
    - répertoire du module (src) et son parent
    """
    env_path = os.getenv("CREDENTIALS_PATH")
    candidates: List[str] = []
    if env_path:
        candidates.append(env_path)
    if preferred_path:
        candidates.append(os.path.abspath(preferred_path))
    # cwd
    candidates.append(os.path.join(os.getcwd(), "credentials.json"))
    # executable dir (frozen)
    if getattr(sys, "frozen", False):
        exe_dir = os.path.dirname(sys.executable)
        candidates.append(os.path.join(exe_dir, "credentials.json"))
        if hasattr(sys, "_MEIPASS"):
            candidates.append(os.path.join(getattr(sys, "_MEIPASS"), "credentials.json"))
    # module dir
    mod_dir = os.path.dirname(os.path.abspath(__file__))
    candidates.append(os.path.join(mod_dir, "credentials.json"))
    candidates.append(os.path.join(os.path.dirname(mod_dir), "credentials.json"))

    for p in candidates:
        try:
            if p and os.path.exists(p):
                return p
        except Exception:
            continue
    # Retourne le préféré ou le nom par défaut (pour l'erreur explicite plus loin)
    return preferred_path or "credentials.json"


class GmailWatcher:
    def __init__(self, credentials_path: str = "credentials.json", token_path: str = "token.json") -> None:
        self.credentials_path = _resolve_credentials_path(credentials_path)
        self.token_path = token_path
        self.creds: Optional[Credentials] = None
        self.service = None

    def _load_credentials(self) -> Credentials:
        logging.info("Chargement des identifiants OAuth (token: %s, creds: %s)", self.token_path, self.credentials_path)
        creds = None
        if os.path.exists(self.token_path):
            logging.info("token.json trouvé, tentative de chargement…")
            try:
                creds = Credentials.from_authorized_user_file(self.token_path, SCOPES)
            except ValueError as e:
                # token.json mal formé ou sans refresh_token -> forcer une nouvelle autorisation
                logging.warning(
                    "token.json invalide (%s). Suppression et relance du flux OAuth avec consent forcé…",
                    e,
                )
                try:
                    os.remove(self.token_path)
                except Exception:
                    pass
                creds = None
        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                logging.info("Token expiré, tentative de rafraîchissement…")
                try:
                    creds.refresh(Request())
                    logging.info("Rafraîchissement du token réussi.")
                except Exception:
                    # Rafraîchissement impossible (token révoqué/expiré sans refresh valide) -> reconsentir
                    logging.warning("Rafraîchissement impossible, lancement du flux OAuth local pour reconsentir…")
                    flow = InstalledAppFlow.from_client_secrets_file(self.credentials_path, SCOPES)
                    try:
                        port = int(os.getenv("OAUTH_LOCAL_SERVER_PORT", "6969"))
                    except ValueError:
                        port = 6969
                    try:
                        logging.info("Ouverture du serveur local OAuth sur le port %s", port)
                        # Forcer un refresh_token en mode offline + consent
                        creds = flow.run_local_server(
                            port=port,
                            access_type="offline",
                            prompt="consent",
                        )
                    except Exception as e:
                        msg = str(e)
                        if "redirect_uri_mismatch" in msg or "MismatchingRedirectURIError" in msg:
                            raise RuntimeError(
                                "redirect_uri_mismatch: Vos identifiants OAuth ne correspondent pas au flux utilisé. "
                                "Créez un client OAuth de type 'Application de bureau' (recommandé) et remplacez credentials.json, "
                                "ou utilisez un client 'Web' avec une redirection autorisée exacte (ex: http://localhost:8080/) "
                                "et définissez OAUTH_LOCAL_SERVER_PORT=8080."
                            ) from e
                        raise
            else:
                logging.info("Aucun token valide trouvé, lancement du flux OAuth local…")
                # Recalcul défensif au cas où l'environnement change
                self.credentials_path = _resolve_credentials_path(self.credentials_path)
                logging.info("Utilisation du credentials.json: %s", self.credentials_path)
                flow = InstalledAppFlow.from_client_secrets_file(self.credentials_path, SCOPES)
                # Permet d'imposer un port fixe si vous utilisez un client OAuth de type "Web"
                # avec une redirection autorisée spécifique (ex: http://localhost:8080/)
                try:
                    port = int(os.getenv("OAUTH_LOCAL_SERVER_PORT", "6969"))
                except ValueError:
                    port = 6969
                try:
                    logging.info("Ouverture du serveur local OAuth sur le port %s", port)
                    # Forcer un refresh_token en mode offline + consent
                    creds = flow.run_local_server(
                        port=port,
                        access_type="offline",
                        prompt="consent",
                    )
                except Exception as e:
                    msg = str(e)
                    if "redirect_uri_mismatch" in msg or "MismatchingRedirectURIError" in msg:
                        raise RuntimeError(
                            "redirect_uri_mismatch: Vos identifiants OAuth ne correspondent pas au flux utilisé. "
                            "Créez un client OAuth de type 'Application de bureau' (recommandé) et remplacez credentials.json, "
                            "ou utilisez un client 'Web' avec une redirection autorisée exacte (ex: http://localhost:8080/) "
                            "et définissez OAUTH_LOCAL_SERVER_PORT=8080."
                        ) from e
                    raise
            # A ce stade on peut encore se retrouver sans refresh_token (comptes/clients particuliers)
            if creds and not getattr(creds, "refresh_token", None):
                logging.warning(
                    "Les informations d'authentification ne contiennent pas de refresh_token. "
                    "Relance du flux OAuth avec 'prompt=consent' pour l'obtenir…"
                )
                flow = InstalledAppFlow.from_client_secrets_file(self.credentials_path, SCOPES)
                try:
                    port = int(os.getenv("OAUTH_LOCAL_SERVER_PORT", "6969"))
                except ValueError:
                    port = 6969
                creds = flow.run_local_server(
                    port=port,
                    access_type="offline",
                    prompt="consent",
                )
            with open(self.token_path, "w") as token:
                token.write(creds.to_json())
            logging.info("token.json enregistré.")
        self.creds = creds
        return creds

    def _ensure_service(self):
        if self.service is None:
            creds = self._load_credentials()
            self.service = build("gmail", "v1", credentials=creds, cache_discovery=False)
            logging.info("Client Gmail initialisé.")
        return self.service

    def search_messages(self, query: Optional[str] = None, max_results: int = 10) -> List[str]:
        service = self._ensure_service()
        q = query or os.getenv("GMAIL_QUERY") or DEFAULT_QUERY
        try:
            logging.info("Recherche des messages avec la requête: %s (max_results=%s)", q, max_results)
            results = service.users().messages().list(userId="me", q=q, maxResults=max_results).execute()
            messages = results.get("messages", [])
            ids = [m["id"] for m in messages]
            logging.info("%s message(s) retourné(s) par la recherche.", len(ids))
            return ids
        except HttpError as e:
            raise RuntimeError(f"Erreur lors de la recherche Gmail: {e}")

    def get_message_raw(self, message_id: str) -> dict:
        service = self._ensure_service()
        try:
            logging.info("Récupération du message id=%s", message_id)
            msg = service.users().messages().get(userId="me", id=message_id, format="full").execute()
            try:
                headers = {h['name'].lower(): h['value'] for h in msg.get('payload', {}).get('headers', [])}
                subject = headers.get('subject', '(sans sujet)')
                sender = headers.get('from', '(inconnu)')
                logging.info("Message récupéré: subject=%s | from=%s | id=%s", subject, sender, message_id)
            except Exception:
                pass
            return msg
        except HttpError as e:
            raise RuntimeError(f"Erreur lors de la récupération du message {message_id}: {e}")

    def mark_as_read(self, message_id: str) -> None:
        service = self._ensure_service()
        try:
            service.users().messages().modify(
                userId="me",
                id=message_id,
                body={"removeLabelIds": ["UNREAD"]},
            ).execute()
            logging.info("Message marqué comme lu: id=%s", message_id)
        except HttpError as e:
            raise RuntimeError(f"Erreur lors du marquage en lu du message {message_id}: {e}")

    def _gather_parts(self, payload: dict, message_id: Optional[str]) -> List[Tuple[str, bytes]]:
        parts: List[Tuple[str, bytes]] = []
        if not payload:
            return parts
        # Helper pour ajouter un part
        def add_part(mime: str, body: dict):
            data = body.get('data')
            attachment_id = body.get('attachmentId')
            if data:
                parts.append((mime, base64.urlsafe_b64decode(data)))
            elif attachment_id and message_id:
                # Récupérer la pièce jointe (utile si HTML/texte est externalisé)
                try:
                    svc = self._ensure_service()
                    att = svc.users().messages().attachments().get(userId="me", messageId=message_id, id=attachment_id).execute()
                    a_data = att.get('data')
                    if a_data:
                        parts.append((mime, base64.urlsafe_b64decode(a_data)))
                except HttpError:
                    pass

        if 'parts' in payload and payload['parts']:
            for p in payload['parts']:
                mime = p.get('mimeType', '')
                body = p.get('body', {})
                if body:
                    add_part(mime, body)
                # Multipart nested
                parts.extend(self._gather_parts(p, message_id))
        else:
            mime = payload.get('mimeType', '')
            body = payload.get('body', {})
            if body:
                add_part(mime, body)
        try:
            logging.info("Collecte des parties du message: %s partie(s) trouvée(s).", len(parts))
        except Exception:
            pass
        return parts

    def extract_update_link_from_message(self, msg: dict) -> Optional[str]:
        payload = msg.get('payload', {})
        message_id = msg.get('id')
        parts = self._gather_parts(payload, message_id)
        logging.info("Extraction du lien d'update: %s partie(s) collectée(s).", len(parts))

        # Parcourir HTML d'abord, puis texte
        html_candidates: List[str] = []
        text_candidates: List[str] = []
        for mime, content in parts:
            try:
                text = content.decode('utf-8', errors='ignore')
            except Exception:
                continue
            if mime.startswith('text/html'):
                html_candidates.append(text)
            elif mime.startswith('text/plain'):
                text_candidates.append(text)

        # Chercher dans HTML
        for html in html_candidates:
            soup = BeautifulSoup(html, 'html5lib')
            for a in soup.find_all('a', href=True):
                href = a['href']
                if '/update-primary-location' in href:
                    logging.info("Lien '/update-primary-location' trouvé dans HTML: %s", href)
                    return href
        # Fallback texte brut: capturer URLs
        import re
        url_re = re.compile(r"https?://\S+")
        for txt in text_candidates:
            for url in url_re.findall(txt):
                if '/update-primary-location' in url:
                    # Nettoyage basique si traînent des ponctuations
                    url = url.rstrip(").,>]')\"")
                    logging.info("Lien '/update-primary-location' trouvé dans texte: %s", url)
                    return url
        logging.info("Aucun lien '/update-primary-location' trouvé dans le message id=%s", message_id)
        return None

    @staticmethod
    def extract_requester_text_from_message(msg: dict) -> Optional[str]:
        """Extrait le contenu du <td> qui contient le libellé 'Demande effectuée par'.

        On recherche dans les parties HTML en priorité.
        """
        payload = msg.get('payload', {})
        # Réutiliser le collecteur sans attachements (texte suffira la plupart du temps)
        # mais on peut appeler _gather_parts si besoin du message id
        message_id = msg.get('id')
        # Note: utiliser la méthode d'instance pour récupérer les parties si nécessaire
        parts: List[Tuple[str, bytes]] = []
        if hasattr(GmailWatcher, "_gather_parts"):
            # type: ignore[attr-defined]
            gw = GmailWatcher()
            parts = gw._gather_parts(payload, message_id)  # type: ignore
        else:
            parts = []

        html_candidates: List[str] = []
        for mime, content in parts:
            try:
                text = content.decode('utf-8', errors='ignore')
            except Exception:
                continue
            if mime.startswith('text/html'):
                html_candidates.append(text)

        from bs4 import BeautifulSoup
        for html in html_candidates:
            soup = BeautifulSoup(html, 'html5lib')
            # Rechercher toutes les cellules <td> contenant la phrase
            for td in soup.find_all('td'):
                txt = td.get_text(separator=' ', strip=True)
                if not txt:
                    continue
                if 'demande effectuée par' in txt.lower():
                    logging.info("Texte 'Demande effectuée par' trouvé: %s", txt)
                    return txt
        logging.info("Texte 'Demande effectuée par' non trouvé dans le message id=%s", msg.get('id'))
        return None
