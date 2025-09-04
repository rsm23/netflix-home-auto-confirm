import os
import re
import time
import logging
from typing import Optional

from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError


def _default_browser_channel() -> Optional[str]:
    # Par défaut sur Windows, tenter Edge puis Chrome
    ch = os.getenv("PLAYWRIGHT_CHANNEL") or os.getenv("BROWSER_CHANNEL")
    if ch:
        return ch
    if os.name == "nt":
        return "msedge"
    return "chrome"


def _default_user_data_dir(channel: Optional[str]) -> Optional[str]:
    # Permet de réutiliser la session existante (cookies Netflix)
    udd = os.getenv("BROWSER_USER_DATA_DIR")
    if udd:
        return udd
    if os.name != "nt":
        return None
    local = os.getenv("LOCALAPPDATA")
    if not local:
        return None
    if channel == "msedge":
        return os.path.join(local, "Microsoft", "Edge", "User Data")
    # fallback Chrome
    return os.path.join(local, "Google", "Chrome", "User Data")


def confirm_netflix_primary_location(
    url: str,
    close_delay_seconds: int = 10,
    channel: Optional[str] = None,
    user_data_dir: Optional[str] = None,
    nav_timeout_ms: int = 30000,
    click_timeout_ms: int = 30000,
) -> bool:
    """Ouvre l'URL Netflix, clique sur le bouton de confirmation et ferme après close_delay_seconds.

    - Tente d'utiliser un profil persistant (Edge/Chrome) pour réutiliser la session Netflix.
    - Boutons ciblés:
      * [data-uia="set-primary-location-action"]
      * Un bouton/lien contenant le texte "Confirmer la mise à jour" (insensible à la casse)
    """
    channel = channel or _default_browser_channel()
    user_data_dir = user_data_dir or _default_user_data_dir(channel)
    logging.info("Lancement Playwright: channel=%s | user_data_dir=%s | url=%s", channel, user_data_dir, url)

    with sync_playwright() as p:
        # Sélectionner le moteur chromium
        browser_type = p.chromium
        launch_kwargs = {}

        # Utiliser un channel (msedge/chrome) si disponible
        if channel:
            launch_kwargs["channel"] = channel

        context = None
        try:
            if user_data_dir:
                # Contexte persistant -> réutilise le profil existant (cookies)
                logging.info("Ouverture d'un contexte persistant avec user_data_dir=%s", user_data_dir)
                context = browser_type.launch_persistent_context(
                    user_data_dir=user_data_dir,
                    **launch_kwargs,
                )
            else:
                # Contexte non persistant (peut nécessiter une reconnexion Netflix)
                logging.info("Ouverture d'un navigateur non persistant")
                browser = browser_type.launch(**launch_kwargs)
                context = browser.new_context()

            page = context.new_page()
            page.set_default_timeout(nav_timeout_ms)
            logging.info("Navigation vers l'URL Netflix…")
            page.goto(url, wait_until="load")

            # Chercher le bouton par data attribute
            btn = page.locator('[data-uia="set-primary-location-action"]')
            found = False
            try:
                btn.wait_for(state="visible", timeout=click_timeout_ms)
                btn.click()
                found = True
                logging.info("Bouton [data-uia='set-primary-location-action'] cliqué")
            except PlaywrightTimeoutError:
                pass

            if not found:
                # Chercher par texte
                # Essayer rôle bouton puis lien
                re_txt = re.compile(r"Confirmer\s+la\s+mise\s+à\s+jour", re.I)
                try:
                    page.get_by_role("button", name=re_txt).first.click(timeout=click_timeout_ms)
                    found = True
                    logging.info("Bouton par texte cliqué (role=button)")
                except PlaywrightTimeoutError:
                    try:
                        page.get_by_role("link", name=re_txt).first.click(timeout=click_timeout_ms)
                        found = True
                        logging.info("Lien par texte cliqué (role=link)")
                    except PlaywrightTimeoutError:
                        pass

            if not found:
                logging.info("Bouton de confirmation introuvable dans le délai imparti; la page reste ouverte.")
                return False

            # Laisser l'action se compléter puis fermer
            logging.info("Attente de %ss avant fermeture…", max(0, close_delay_seconds))
            time.sleep(max(0, close_delay_seconds))
            return True
        finally:
            try:
                if context:
                    logging.info("Fermeture du contexte navigateur")
                    context.close()
            except Exception:
                pass
