import base64
import os
from typing import List, Optional, Tuple

from bs4 import BeautifulSoup
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# Scopes pour lire et modifier (marquer comme lu)
SCOPES = ["https://www.googleapis.com/auth/gmail.modify"]

DEFAULT_QUERY = 'from:info@account.netflix.com subject:"comment mettre à jour votre foyer Netflix" is:unread'


class GmailWatcher:
    def __init__(self, credentials_path: str = "credentials.json", token_path: str = "token.json") -> None:
        self.credentials_path = credentials_path
        self.token_path = token_path
        self.creds: Optional[Credentials] = None
        self.service = None

    def _load_credentials(self) -> Credentials:
        creds = None
        if os.path.exists(self.token_path):
            creds = Credentials.from_authorized_user_file(self.token_path, SCOPES)
        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                try:
                    creds.refresh(Request())
                except Exception:
                    # Rafraîchissement impossible (token révoqué/expiré sans refresh valide) -> reconsentir
                    flow = InstalledAppFlow.from_client_secrets_file(self.credentials_path, SCOPES)
                    try:
                        port = int(os.getenv("OAUTH_LOCAL_SERVER_PORT", "0"))
                    except ValueError:
                        port = 0
                    try:
                        creds = flow.run_local_server(port=port)
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
                flow = InstalledAppFlow.from_client_secrets_file(self.credentials_path, SCOPES)
                # Permet d'imposer un port fixe si vous utilisez un client OAuth de type "Web"
                # avec une redirection autorisée spécifique (ex: http://localhost:8080/)
                try:
                    port = int(os.getenv("OAUTH_LOCAL_SERVER_PORT", "0"))
                except ValueError:
                    port = 0
                try:
                    creds = flow.run_local_server(port=port)
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
            with open(self.token_path, "w") as token:
                token.write(creds.to_json())
        self.creds = creds
        return creds

    def _ensure_service(self):
        if self.service is None:
            creds = self._load_credentials()
            self.service = build("gmail", "v1", credentials=creds, cache_discovery=False)
        return self.service

    def search_messages(self, query: Optional[str] = None, max_results: int = 10) -> List[str]:
        service = self._ensure_service()
        q = query or os.getenv("GMAIL_QUERY") or DEFAULT_QUERY
        try:
            results = service.users().messages().list(userId="me", q=q, maxResults=max_results).execute()
            messages = results.get("messages", [])
            return [m["id"] for m in messages]
        except HttpError as e:
            raise RuntimeError(f"Erreur lors de la recherche Gmail: {e}")

    def get_message_raw(self, message_id: str) -> dict:
        service = self._ensure_service()
        try:
            msg = service.users().messages().get(userId="me", id=message_id, format="full").execute()
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
        return parts

    def extract_update_link_from_message(self, msg: dict) -> Optional[str]:
        payload = msg.get('payload', {})
        message_id = msg.get('id')
        parts = self._gather_parts(payload, message_id)

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
                    return href
        # Fallback texte brut: capturer URLs
        import re
        url_re = re.compile(r"https?://\S+")
        for txt in text_candidates:
            for url in url_re.findall(txt):
                if '/update-primary-location' in url:
                    # Nettoyage basique si traînent des ponctuations
                    url = url.rstrip(").,>]')\"")
                    return url
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
                    return txt
        return None
