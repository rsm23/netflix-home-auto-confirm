import os
import sys
import threading
import time
import logging
from typing import Optional

import pystray
from pystray import MenuItem as item
from PIL import Image, ImageDraw
import tkinter as tk
from tkinter import ttk, messagebox

# Charger .env si présent
try:
    from dotenv import load_dotenv  # type: ignore
    load_dotenv()
except Exception:
    pass

# Importer les fonctions de l'app
from .main import process_once, _now_ms


def _make_image(color_bg=(30, 144, 255), color_fg=(255, 255, 255)) -> Image.Image:
    # Génère une icône simple (64x64)
    img = Image.new('RGB', (64, 64), color_bg)
    d = ImageDraw.Draw(img)
    d.ellipse((12, 12, 52, 52), outline=color_fg, width=4)
    d.rectangle((30, 20, 34, 44), fill=color_fg)
    return img


class WatcherThread(threading.Thread):
    def __init__(
        self,
        stop_event: threading.Event,
        interval: int = 60,
        query: Optional[str] = None,
        open_once: bool = True,
        debug: bool = False,
        auto_click: bool = True,
        close_delay: int = 10,
    ) -> None:
        super().__init__(daemon=True)
        self.stop_event = stop_event
        self.interval = interval
        self.query = query
        self.open_once = open_once
        self.debug = debug
        self.auto_click = auto_click
        self.close_delay = close_delay

    def run(self) -> None:
        logging.info("Watcher démarré: interval=%ss auto_click=%s", self.interval, self.auto_click)
        anchor = _now_ms()
        while not self.stop_event.is_set():
            try:
                code, clicked, new_anchor = process_once(
                    query=self.query,
                    open_once=self.open_once,
                    debug=self.debug,
                    auto_click=self.auto_click,
                    close_delay=self.close_delay,
                    anchor_ts_ms=anchor,
                )
                if clicked and new_anchor is not None:
                    anchor = new_anchor
            except Exception as e:
                logging.exception("Erreur watcher: %s", e)
            # Attente avec sortie anticipée si stop_event
            if self.stop_event.wait(self.interval):
                break
        logging.info("Watcher arrêté")


class TrayApp:
    def __init__(self) -> None:
        self.icon = pystray.Icon("confirm_netflix_house", _make_image(), "Netflix House Watcher")
        self.stop_event = threading.Event()
        self.worker: Optional[WatcherThread] = None
        self.interval = int(os.getenv("POLL_INTERVAL", "60"))
        self.query = os.getenv("GMAIL_QUERY")
        self.auto_click = True
        self.close_delay = int(os.getenv("AUTO_CLOSE_DELAY", "10"))

        self.icon.menu = pystray.Menu(
            item(lambda: f"Status: {'RUNNING' if self.worker and self.worker.is_alive() else 'STOPPED'}", lambda: None, enabled=False),
            item("Connect", self.connect),
            item("Disconnect", self.disconnect),
            item("Settings", self.open_settings),
            item("Start", self.start),
            item("Stop", self.stop),
            item("Quit", self.quit),
        )

        # Setup logging
        logs_dir = os.path.join(os.getcwd(), "logs")
        os.makedirs(logs_dir, exist_ok=True)
        logging.basicConfig(
            level=logging.INFO,
            format="%(asctime)s [%(levelname)s] %(message)s",
            handlers=[
                logging.FileHandler(os.path.join(logs_dir, "app.log"), encoding="utf-8"),
                logging.StreamHandler(sys.stdout),
            ],
        )

    def start(self, icon: Optional[pystray.Icon] = None, item_clicked: Optional[item] = None):
        if self.worker and self.worker.is_alive():
            return
        self.stop_event.clear()
        self.worker = WatcherThread(
            stop_event=self.stop_event,
            interval=self.interval,
            query=self.query,
            open_once=True,
            debug=False,
            auto_click=self.auto_click,
            close_delay=self.close_delay,
        )
        self.worker.start()
        logging.info("Start demandé")
        self.icon.title = "Netflix House Watcher (RUNNING)"

    def stop(self, icon: Optional[pystray.Icon] = None, item_clicked: Optional[item] = None):
        if not (self.worker and self.worker.is_alive()):
            return
        self.stop_event.set()
        self.worker.join(timeout=5)
        logging.info("Stop demandé")
        self.icon.title = "Netflix House Watcher (STOPPED)"

    def connect(self, icon: Optional[pystray.Icon] = None, item_clicked: Optional[item] = None):
        """Force un appel aux APIs qui déclenchera le flux OAuth si nécessaire."""
        try:
            from .gmail_client import GmailWatcher
            gw = GmailWatcher()
            # Appel léger: récupérer 0 mail déclenche juste l'auth si besoin
            gw.search_messages(max_results=1)
            messagebox.showinfo("Connecté", "Connexion/consentement effectué avec succès.")
        except Exception as e:
            logging.exception("Erreur connect: %s", e)
            messagebox.showerror("Erreur connexion", str(e))

    def disconnect(self, icon: Optional[pystray.Icon] = None, item_clicked: Optional[item] = None):
        """Supprime le token local pour forcer un reconsentement au prochain appel."""
        try:
            token_path = os.path.join(os.getcwd(), 'token.json')
            if os.path.exists(token_path):
                os.remove(token_path)
                messagebox.showinfo("Déconnecté", "Token supprimé. Le prochain appel redemandera l'autorisation.")
            else:
                messagebox.showinfo("Info", "Aucun token à supprimer.")
        except Exception as e:
            logging.exception("Erreur disconnect: %s", e)
            messagebox.showerror("Erreur déconnexion", str(e))

    def open_settings(self, icon: Optional[pystray.Icon] = None, item_clicked: Optional[item] = None):
        """Ouvre une petite fenêtre pour régler intervalle et délai de fermeture."""
        if getattr(self, "_settings_open", False):
            return
        self._settings_open = True

        def on_close():
            self._settings_open = False
            win.destroy()

        win = tk.Tk()
        win.title("Settings - Netflix House Watcher")
        win.protocol("WM_DELETE_WINDOW", on_close)

        frm = ttk.Frame(win, padding=10)
        frm.grid(row=0, column=0, sticky="nsew")

        ttk.Label(frm, text="Interval (s)").grid(row=0, column=0, sticky="w", padx=5, pady=5)
        interval_var = tk.StringVar(value=str(self.interval))
        interval_entry = ttk.Entry(frm, textvariable=interval_var, width=10)
        interval_entry.grid(row=0, column=1, sticky="w", padx=5, pady=5)

        ttk.Label(frm, text="Close Delay (s)").grid(row=1, column=0, sticky="w", padx=5, pady=5)
        delay_var = tk.StringVar(value=str(self.close_delay))
        delay_entry = ttk.Entry(frm, textvariable=delay_var, width=10)
        delay_entry.grid(row=1, column=1, sticky="w", padx=5, pady=5)

        def save_and_close():
            try:
                new_interval = int(interval_var.get())
                new_delay = int(delay_var.get())
                if new_interval <= 0 or new_delay < 0:
                    raise ValueError("Valeurs invalides")
                self.interval = new_interval
                self.close_delay = new_delay
                messagebox.showinfo("OK", "Paramètres sauvegardés. Redémarrez le watcher pour appliquer.")
                on_close()
            except Exception:
                messagebox.showerror("Erreur", "Veuillez entrer des nombres valides.")

        buttons = ttk.Frame(frm)
        buttons.grid(row=2, column=0, columnspan=2, pady=10)
        ttk.Button(buttons, text="Save", command=save_and_close).grid(row=0, column=0, padx=5)
        ttk.Button(buttons, text="Cancel", command=on_close).grid(row=0, column=1, padx=5)

        win.mainloop()

    def quit(self, icon: Optional[pystray.Icon] = None, item_clicked: Optional[item] = None):
        try:
            self.stop()
        finally:
            self.icon.stop()

    def run(self):
        # Démarrer automatiquement
        self.start()
        self.icon.run()


def main():
    app = TrayApp()
    app.run()


if __name__ == "__main__":
    main()
