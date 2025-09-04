import argparse
import os
import sys
import time
import logging
import webbrowser
from typing import Optional, Tuple

try:
    # Contexte package
    from .gmail_client import GmailWatcher  # type: ignore
except Exception:
    # Contexte script/pyinstaller
    from gmail_client import GmailWatcher  # type: ignore

# Charger un éventuel .env
try:
    from dotenv import load_dotenv  # type: ignore
    load_dotenv()
except Exception:
    pass


def open_update_link(link: str, open_once: bool = False, auto_click: bool = False, close_delay: int = 10) -> None:
    logging.info("Ouverture du lien: %s", link)
    if auto_click:
        # Utilise Playwright pour ouvrir, cliquer, puis fermer après délai
        try:
            from .automation import confirm_netflix_primary_location  # lazy import
        except ModuleNotFoundError as e:
            if "playwright" in str(e):
                raise RuntimeError(
                    "Le module 'playwright' est introuvable. Activez votre venv puis installez-le:\n"
                    "  .\\.venv\\Scripts\\Activate.ps1\n"
                    "  pip install -r requirements.txt\n"
                    "  python -m playwright install --with-deps chromium\n"
                ) from e
            raise
        ok = confirm_netflix_primary_location(link, close_delay_seconds=close_delay)
        logging.info("Automatisation Playwright terminée avec succès=%s", ok)
    else:
        webbrowser.open(link, new=2)
    if open_once:
        os.environ["LAST_OPENED_LINK"] = link


def _now_ms() -> int:
    return int(time.time() * 1000)


def process_once(
    query: Optional[str] = None,
    open_once: bool = False,
    debug: bool = False,
    auto_click: bool = False,
    close_delay: int = 10,
    anchor_ts_ms: Optional[int] = None,
    output_dir: Optional[str] = None,
) -> Tuple[int, bool, Optional[int]]:
    watcher = GmailWatcher()
    ids = watcher.search_messages(query=query, max_results=10)
    if not ids:
        logging.info("Aucun message correspondant trouvé.")
        return 1, False, None
    logging.info("%s message(s) candidat(s) à analyser.", len(ids))
    for mid in ids:
        msg = watcher.get_message_raw(mid)
        # Filtrer par date de réception (internalDate en ms depuis epoch)
        try:
            internal_ms = int(msg.get('internalDate'))
        except (TypeError, ValueError):
            internal_ms = 0
        logging.info("Analyse message id=%s | internalDate=%s | ancre=%s", mid, internal_ms, anchor_ts_ms)
        if anchor_ts_ms is not None and internal_ms <= anchor_ts_ms:
            # Ignorer les messages reçus avant/ancré
            logging.info("Ignoré (avant l'ancre): id=%s", mid)
            continue
        link = watcher.extract_update_link_from_message(msg)
        if debug and not link:
            # Dump minimal sujet/expéditeur pour debug
            headers = {h['name'].lower(): h['value'] for h in msg.get('payload', {}).get('headers', [])}
            subject = headers.get('subject', '(sans sujet)')
            sender = headers.get('from', '(inconnu)')
            logging.info("Pas de lien trouvé dans: subject=%s | from=%s | id=%s", subject, sender, mid)
        if link:
            logging.info("Lien Netflix détecté pour id=%s: %s", mid, link)
            last = os.getenv("LAST_OPENED_LINK")
            if open_once and last == link:
                logging.info("Lien déjà ouvert précédemment et open_once actif, on saute: %s", link)
                continue
            # Ouvre et tente le clic via Playwright (si activé)
            success = False
            if auto_click:
                try:
                    from .automation import confirm_netflix_primary_location  # lazy import
                    success = confirm_netflix_primary_location(link, close_delay_seconds=close_delay)
                    logging.info("Résultat du clic Playwright: %s", success)
                except Exception as e:
                    logging.warning("Automatisation Playwright échouée: %s", e)
            else:
                open_update_link(link, open_once=open_once, auto_click=False, close_delay=close_delay)
                success = True  # Considérer le clic externe comme succès logique d'ouverture

            if success:
                # Extraire texte 'Demande effectuée par' et écrire dans un .txt
                try:
                    requester = watcher.extract_requester_text_from_message(msg)
                    headers = {h['name'].lower(): h['value'] for h in msg.get('payload', {}).get('headers', [])}
                    subject = headers.get('subject', '(sans sujet)')
                    ts = _now_ms()
                    out_dir = output_dir or os.getenv('OUTPUT_DIR') or os.path.join(os.getcwd(), 'out')
                    os.makedirs(out_dir, exist_ok=True)
                    out_path = os.path.join(out_dir, f"requester_{ts}.txt")
                    with open(out_path, 'w', encoding='utf-8') as f:
                        f.write(f"Subject: {subject}\n")
                        f.write(f"Message-ID: {mid}\n")
                        f.write("--- Demande effectuée par ---\n")
                        f.write((requester or "(non trouvé)") + "\n")
                    logging.info("Détails sauvegardés dans %s", out_path)
                except Exception as e:
                    logging.warning("Impossible d'extraire/écrire le détails demandeur: %s", e)

            # Marquer le message comme lu après un clic
            try:
                watcher.mark_as_read(mid)
            except Exception as e:
                logging.warning("Impossible de marquer le message comme lu: %s", e)
            # Mettre à jour l'ancre au moment courant pour éviter les anciens emails
            new_anchor = _now_ms()
            logging.info("Traitement terminé pour id=%s, nouvelle ancre=%s", mid, new_anchor)
            return 0, True, new_anchor
    logging.info("Aucun lien '/update-primary-location' trouvé dans les messages récents.")
    return 2, False, None


def watch_loop(interval: int = 60, query: Optional[str] = None, open_once: bool = False, debug: bool = False, auto_click: bool = False, close_delay: int = 10, output_dir: Optional[str] = None) -> int:
    logging.info("Surveillance démarrée. Intervalle: %ss | auto_click=%s | close_delay=%ss | output_dir=%s", interval, auto_click, close_delay, output_dir)
    # Ancrage à l'ouverture de l'application: ignorer les anciens emails
    anchor_ts_ms = _now_ms()
    while True:
        try:
            code, clicked, new_anchor = process_once(
                query=query,
                open_once=open_once,
                debug=debug,
                auto_click=auto_click,
                close_delay=close_delay,
                anchor_ts_ms=anchor_ts_ms,
                output_dir=output_dir,
            )
            if clicked and new_anchor is not None:
                anchor_ts_ms = new_anchor
            logging.info("Cycle terminé avec code=%s | clicked=%s | new_anchor=%s", code, clicked, new_anchor)
        except KeyboardInterrupt:
            logging.info("Arrêt par l'utilisateur.")
            return 0
        except Exception as e:
            logging.exception("Erreur dans la boucle: %s", e)
        time.sleep(interval)


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(description="Ouvre automatiquement le lien Netflix d'update du foyer depuis Gmail.")
    sub = parser.add_subparsers(dest="cmd", required=False)

    p_once = sub.add_parser("once", help="Vérifie une fois")
    p_once.add_argument("--query", type=str, default=None, help="Requête Gmail personnalisée")
    p_once.add_argument("--open-once", action="store_true", help="N'ouvre pas le même lien deux fois")
    p_once.add_argument("--debug", action="store_true", help="Affiche des infos de debug si aucun lien trouvé")
    p_once.add_argument("--auto-click", action="store_true", help="Ouvre la page avec Playwright, clique sur le bouton et ferme")
    p_once.add_argument("--close-delay", type=int, default=10, help="Délai avant fermeture après clic (secondes)")
    p_once.add_argument("--since-epoch-ms", type=int, default=None, help="Filtrer uniquement les emails reçus après ce timestamp (ms epoch). Par défaut: maintenant.")
    p_once.add_argument("--output-dir", type=str, default=None, help="Dossier où enregistrer les fichiers .txt")

    p_watch = sub.add_parser("watch", help="Surveille en boucle")
    p_watch.add_argument("--interval", type=int, default=int(os.getenv("POLL_INTERVAL", "60")))
    p_watch.add_argument("--query", type=str, default=None, help="Requête Gmail personnalisée")
    p_watch.add_argument("--open-once", action="store_true", help="N'ouvre pas le même lien deux fois")
    p_watch.add_argument("--debug", action="store_true", help="Affiche des infos de debug si aucun lien trouvé")
    p_watch.add_argument("--auto-click", action="store_true", help="Ouvre la page avec Playwright, clique sur le bouton et ferme")
    p_watch.add_argument("--close-delay", type=int, default=10, help="Délai avant fermeture après clic (secondes)")
    p_watch.add_argument("--output-dir", type=str, default=None, help="Dossier où enregistrer les fichiers .txt")

    args = parser.parse_args(argv)

    if args.cmd == "watch":
        return watch_loop(interval=args.interval, query=args.query, open_once=args.open_once, debug=args.debug, auto_click=args.auto_click, close_delay=args.close_delay, output_dir=args.output_dir)
    # default: once
    anchor = args.since_epoch_ms if args.since_epoch_ms is not None else _now_ms()
    code, _clicked, _new_anchor = process_once(query=args.query, open_once=args.open_once, debug=args.debug, auto_click=args.auto_click, close_delay=args.close_delay, anchor_ts_ms=anchor, output_dir=args.output_dir)
    return code


if __name__ == "__main__":
    raise SystemExit(main())
