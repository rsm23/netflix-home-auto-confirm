import os
import re
from typing import List

# Adresse e-mail de l’expéditeur Netflix
SENDER_EMAIL: str = os.getenv("SENDER_EMAIL", "info@account.netflix.com")

# Motifs à rechercher dans les URLs à l’intérieur des emails
# Variable d’environnement LINK_SUBSTRINGS peut surcharger, ex: "pattern1,pattern2"
_default_link_substrings = "update-primary-location,update-primary,set-primary"
LINK_SUBSTRINGS: List[str] = [
    s.strip().lower()
    for s in os.getenv("LINK_SUBSTRINGS", _default_link_substrings).split(",")
    if s.strip()
]

# Sélecteur du bouton de confirmation sur Netflix
CONFIRM_BUTTON_SELECTOR: str = os.getenv(
    "CONFIRM_BUTTON_SELECTOR",
    '[data-uia="set-primary-location-action"]',
)

# Regex du texte alternatif du bouton (insensible à la casse)
CONFIRM_TEXT_REGEX: str = os.getenv(
    "CONFIRM_TEXT_REGEX",
    r"Confirmer\s+la\s+mise\s+à\s+jour",
)
CONFIRM_TEXT_RE = re.compile(CONFIRM_TEXT_REGEX, re.I)

# Requête Gmail par défaut (modulable via GMAIL_QUERY_DEFAULT)
DEFAULT_GMAIL_QUERY: str = os.getenv(
    "GMAIL_QUERY_DEFAULT",
    f"from:{SENDER_EMAIL} is:unread in:inbox",
)

# Libellé ("dossier") Gmail vers lequel déplacer les emails traités
TARGET_LABEL: str = os.getenv("TARGET_LABEL", "Netflix Location Update")
