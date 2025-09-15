import os
import json
import sys
import threading
import time
import logging
import atexit
from typing import Optional

import pystray
from pystray import MenuItem as item
from PIL import Image, ImageDraw
import tkinter as tk
from tkinter import ttk, messagebox, filedialog

# Charger .env si présent
try:
    from dotenv import load_dotenv  # type: ignore
    load_dotenv()
except Exception:
    pass

# --- Single-instance (Windows .exe) -----------------------------------------------------------
_single_instance_handle = None

def _enforce_single_instance_if_frozen() -> None:
    """Empêche plusieurs instances lorsque packagé en .exe (Windows).

    Crée un mutex nommé Global\\confirm-netflix-house-singleton. Si déjà présent, affiche
    un message et termine le processus immédiatement.
    """
    if not getattr(sys, "frozen", False):
        return
    if os.name != "nt":
        return
    try:
        import ctypes
        from ctypes import wintypes

        kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
        CreateMutexW = kernel32.CreateMutexW
        CreateMutexW.argtypes = [wintypes.LPVOID, wintypes.BOOL, wintypes.LPCWSTR]
        CreateMutexW.restype = wintypes.HANDLE

        mutex_name = "Global\\confirm-netflix-house-singleton"
        h_mutex = CreateMutexW(None, False, mutex_name)
        if not h_mutex:
            return
        err = ctypes.get_last_error()
        global _single_instance_handle
        _single_instance_handle = h_mutex
        if err == 183:  # ERROR_ALREADY_EXISTS
            # Optionnel: informer l'utilisateur
            try:
                user32 = ctypes.WinDLL("user32", use_last_error=True)
                MB_ICONEXCLAMATION = 0x30
                user32.MessageBoxW(None, "L'application est déjà en cours d'exécution.", "confirm-netflix-house", MB_ICONEXCLAMATION)
            except Exception:
                pass
            # Sortie immédiate
            sys.exit(0)
    except Exception:
        # En cas d'échec du mécanisme, ne pas bloquer l'app.
        pass

def _release_single_instance() -> None:
    if os.name != "nt":
        return
    try:
        import ctypes
        if _single_instance_handle:
            kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
            kernel32.CloseHandle(_single_instance_handle)
    except Exception:
        pass

# Importer les fonctions de l'app (compatible package et exécutable PyInstaller)
try:
    # Contexte package (python -m src.tray_app)
    from .main import process_once, _now_ms  # type: ignore
except Exception:
    try:
        # PyInstaller avec package 'src' conservé
        from src.main import process_once, _now_ms  # type: ignore
    except Exception:
        # Contexte script/pyinstaller (imports absolus)
        from main import process_once, _now_ms  # type: ignore


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
        output_dir: Optional[str] = None,
    ) -> None:
        super().__init__(daemon=True)
        self.stop_event = stop_event
        self.interval = interval
        self.query = query
        self.open_once = open_once
        self.debug = debug
        self.auto_click = auto_click
        self.close_delay = close_delay
        self.output_dir = output_dir

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
                    output_dir=self.output_dir,
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
        self.output_dir: Optional[str] = os.getenv("OUTPUT_DIR")
        # Logs: activés par défaut, dossier par défaut ./logs
        self.logging_enabled: bool = True
        self.log_dir: str = os.path.join(self._app_base_dir(), "logs")
        # Run at startup (Windows only)
        self.run_at_startup: bool = False
        # Port OAuth local (par défaut 6969)
        try:
            self.oauth_port: int = int(os.getenv("OAUTH_LOCAL_SERVER_PORT", "6969"))
        except ValueError:
            self.oauth_port = 6969

        # Charger une configuration persistée si disponible
        self._load_config()

        self.icon.menu = pystray.Menu(
            item(lambda _item: f"Status: {'RUNNING' if self.worker and self.worker.is_alive() else 'STOPPED'} | OAuth Port: {self.oauth_port}", None, enabled=False),
            item(lambda _item: f"Config: {self._config_path()}", None, enabled=False),
            item("Connect", self.connect),
            item("Disconnect", self.disconnect),
            item("Settings", self.open_settings, default=True),
            item("Start", self.start),
            item("Stop", self.stop),
            item("Quit", self.quit),
        )

        # Setup logging selon configuration
        self._setup_logging()
        # Appliquer le démarrage auto si configuré
        self._apply_run_at_startup(self.run_at_startup)

    def start(self, icon: Optional[pystray.Icon] = None, item_clicked: Optional[item] = None):
        if self.worker and self.worker.is_alive():
            return
        self.stop_event.clear()
        # Appliquer le port OAuth choisi dans l'environnement
        os.environ["OAUTH_LOCAL_SERVER_PORT"] = str(self.oauth_port)
        self.worker = WatcherThread(
            stop_event=self.stop_event,
            interval=self.interval,
            query=self.query,
            open_once=True,
            debug=False,
            auto_click=self.auto_click,
            close_delay=self.close_delay,
            output_dir=self.output_dir,
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
            # Appliquer le port OAuth choisi dans l'environnement
            os.environ["OAUTH_LOCAL_SERVER_PORT"] = str(self.oauth_port)
            try:
                from .gmail_client import GmailWatcher  # type: ignore
            except Exception:
                try:
                    from src.gmail_client import GmailWatcher  # type: ignore
                except Exception:
                    from gmail_client import GmailWatcher  # type: ignore
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
        """Ouvre une petite fenêtre pour régler intervalle, délai de fermeture et dossier de sortie."""
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
        # Close delay
        ttk.Label(frm, text="Close Delay (s)").grid(row=1, column=0, sticky="w", padx=5, pady=5)
        delay_var = tk.StringVar(value=str(self.close_delay))
        delay_entry = ttk.Entry(frm, textvariable=delay_var, width=10)
        delay_entry.grid(row=1, column=1, sticky="w", padx=5, pady=5)
        # Output folder
        ttk.Label(frm, text="Output Folder").grid(row=2, column=0, sticky="w", padx=5, pady=5)
        out_var = tk.StringVar(value=str(self.output_dir or ""))
        out_entry = ttk.Entry(frm, textvariable=out_var, width=40)
        out_entry.grid(row=2, column=1, sticky="w", padx=5, pady=5)

        def browse_folder():
            folder = filedialog.askdirectory()
            if folder:
                out_var.set(folder)

        ttk.Button(frm, text="Browse...", command=browse_folder).grid(row=2, column=2, sticky="w", padx=5, pady=5)

        # OAuth port
        ttk.Label(frm, text="OAuth Port").grid(row=3, column=0, sticky="w", padx=5, pady=5)
        oauth_var = tk.StringVar(value=str(self.oauth_port))
        oauth_entry = ttk.Entry(frm, textvariable=oauth_var, width=10)
        oauth_entry.grid(row=3, column=1, sticky="w", padx=5, pady=5)

        # Logging enable/disable
        log_enable_var = tk.BooleanVar(value=bool(self.logging_enabled))
        log_enable_chk = ttk.Checkbutton(frm, text="Enable Logging", variable=log_enable_var)
        log_enable_chk.grid(row=4, column=0, sticky="w", padx=5, pady=5)

        # Logs folder
        ttk.Label(frm, text="Logs Folder").grid(row=5, column=0, sticky="w", padx=5, pady=5)
        log_dir_var = tk.StringVar(value=str(self.log_dir or ""))
        log_dir_entry = ttk.Entry(frm, textvariable=log_dir_var, width=40)
        log_dir_entry.grid(row=5, column=1, sticky="w", padx=5, pady=5)

        def browse_log_folder():
            folder = filedialog.askdirectory()
            if folder:
                log_dir_var.set(folder)

        ttk.Button(frm, text="Browse...", command=browse_log_folder).grid(row=5, column=2, sticky="w", padx=5, pady=5)

        # Run at startup
        run_startup_var = tk.BooleanVar(value=bool(self.run_at_startup))
        run_startup_chk = ttk.Checkbutton(frm, text="Run at Windows startup", variable=run_startup_var)
        run_startup_chk.grid(row=6, column=0, sticky="w", padx=5, pady=5)

        # Config path display
        ttk.Label(frm, text="Config file").grid(row=7, column=0, sticky="w", padx=5, pady=5)
        cfg_path = self._config_path()
        cfg_label = ttk.Label(frm, text=cfg_path, wraplength=420)
        cfg_label.grid(row=7, column=1, sticky="w", padx=5, pady=5)

        def open_config_folder():
            try:
                folder = os.path.dirname(cfg_path)
                if os.path.isdir(folder):
                    os.startfile(folder)  # Windows
            except Exception as e:
                logging.warning("Impossible d'ouvrir le dossier de config: %s", e)

        ttk.Button(frm, text="Open folder", command=open_config_folder).grid(row=7, column=2, sticky="w", padx=5, pady=5)

        def save_and_close():
            try:
                was_running = bool(self.worker and self.worker.is_alive())
                new_interval = int(interval_var.get())
                new_delay = int(delay_var.get())
                if new_interval <= 0 or new_delay < 0:
                    raise ValueError("Valeurs invalides")
                self.interval = new_interval
                self.close_delay = new_delay
                new_out = out_var.get().strip()
                self.output_dir = new_out or None
                # Port OAuth
                new_port = int(oauth_var.get())
                if new_port <= 0 or new_port > 65535:
                    raise ValueError("Port invalide")
                self.oauth_port = new_port
                # Logging settings
                self.logging_enabled = bool(log_enable_var.get())
                new_log_dir = log_dir_var.get().strip()
                if not new_log_dir:
                    # si vide, remettre dossier par défaut
                    new_log_dir = os.path.join(self._app_base_dir(), "logs")
                self.log_dir = new_log_dir
                # Run at startup
                new_run_startup = bool(run_startup_var.get())
                if new_run_startup != self.run_at_startup:
                    self.run_at_startup = new_run_startup
                    self._apply_run_at_startup(self.run_at_startup)
                # Répercuter immédiatement dans l'environnement
                os.environ["OAUTH_LOCAL_SERVER_PORT"] = str(self.oauth_port)
                if self.output_dir:
                    os.environ["OUTPUT_DIR"] = self.output_dir
                else:
                    os.environ.pop("OUTPUT_DIR", None)
                # Sauvegarder la configuration persistée
                self._save_config()
                # Reconfigurer le logging maintenant
                self._setup_logging()
                # Redémarrer automatiquement le watcher si nécessaire
                if was_running:
                    logging.info("Redémarrage du watcher suite au changement de configuration…")
                    try:
                        self.stop()
                        # Petite latence pour laisser le thread se terminer proprement
                        time.sleep(0.2)
                        self.start()
                        message = "Paramètres sauvegardés. Le watcher a été redémarré."
                    except Exception as e:
                        logging.exception("Echec du redémarrage du watcher: %s", e)
                        message = "Paramètres sauvegardés, mais le redémarrage automatique a échoué. Redémarrez manuellement."
                else:
                    message = "Paramètres sauvegardés."
                messagebox.showinfo("OK", message)
                on_close()
            except Exception:
                messagebox.showerror("Erreur", "Veuillez entrer des nombres valides.")

        buttons = ttk.Frame(frm)
        buttons.grid(row=8, column=0, columnspan=3, pady=10)
        ttk.Button(buttons, text="Save", command=save_and_close).grid(row=0, column=0, padx=5)
        ttk.Button(buttons, text="Cancel", command=on_close).grid(row=0, column=1, padx=5)
        win.mainloop()

    # --- Persistence helpers ---
    def _config_path(self) -> str:
        return os.path.join(self._app_base_dir(), "settings.json")

    def _app_base_dir(self) -> str:
        if getattr(sys, "frozen", False):
            return os.path.dirname(sys.executable)
        # fallback dev
        return os.getcwd()

    def _load_config(self) -> None:
        try:
            cfg_path = self._config_path()
            if not os.path.exists(cfg_path):
                return
            with open(cfg_path, "r", encoding="utf-8") as f:
                cfg = json.load(f)
            # Appliquer valeurs si présentes et valides
            interval = int(cfg.get("interval", self.interval))
            if interval > 0:
                self.interval = interval
            close_delay = int(cfg.get("close_delay", self.close_delay))
            if close_delay >= 0:
                self.close_delay = close_delay
            output_dir = cfg.get("output_dir")
            if isinstance(output_dir, str) and output_dir.strip():
                self.output_dir = output_dir.strip()
                os.environ["OUTPUT_DIR"] = self.output_dir
            oauth_port = int(cfg.get("oauth_port", self.oauth_port))
            if 0 < oauth_port <= 65535:
                self.oauth_port = oauth_port
                os.environ["OAUTH_LOCAL_SERVER_PORT"] = str(self.oauth_port)
            # Logging
            logging_enabled = cfg.get("logging_enabled", self.logging_enabled)
            self.logging_enabled = bool(logging_enabled)
            log_dir = cfg.get("log_dir", self.log_dir)
            if isinstance(log_dir, str) and log_dir.strip():
                self.log_dir = log_dir.strip()
            # Run at startup
            ras = cfg.get("run_at_startup", self.run_at_startup)
            self.run_at_startup = bool(ras)
        except Exception:
            # Ignorer les erreurs de lecture/parse et garder les valeurs actuelles
            logging.warning("Impossible de charger settings.json, valeurs par défaut conservées.")

    def _save_config(self) -> None:
        try:
            cfg = {
                "interval": self.interval,
                "close_delay": self.close_delay,
                "output_dir": self.output_dir,
                "oauth_port": self.oauth_port,
                "logging_enabled": self.logging_enabled,
                "log_dir": self.log_dir,
                "run_at_startup": self.run_at_startup,
            }
            with open(self._config_path(), "w", encoding="utf-8") as f:
                json.dump(cfg, f, ensure_ascii=False, indent=2)
        except Exception as e:
            logging.warning("Impossible d'enregistrer settings.json: %s", e)

    def _setup_logging(self) -> None:
        """Configure les handlers de logging selon les paramètres utilisateur.
        Format requis: "[full date time] : message" -> on utilise [YYYY-MM-DD HH:MM:SS] : message
        """
        root = logging.getLogger()
        root.setLevel(logging.INFO)
        # Nettoyer les handlers existants
        for h in list(root.handlers):
            try:
                h.flush()
                h.close()
            except Exception:
                pass
            root.removeHandler(h)

        # Toujours avoir un flux console pour debug local
        fmt = logging.Formatter(fmt="[%(asctime)s] : %(message)s", datefmt="%Y-%m-%d %H:%M:%S")
        sh = logging.StreamHandler(sys.stdout)
        sh.setLevel(logging.INFO)
        sh.setFormatter(fmt)
        root.addHandler(sh)

        # Ajouter FileHandler si activé
        if self.logging_enabled:
            try:
                os.makedirs(self.log_dir, exist_ok=True)
                log_file = os.path.join(self.log_dir, "app.log.txt")
                fh = logging.FileHandler(log_file, encoding="utf-8")
                fh.setLevel(logging.INFO)
                fh.setFormatter(fmt)
                root.addHandler(fh)
                logging.info("Fichier de log: %s", log_file)
            except Exception as e:
                logging.warning("Impossible d'initialiser le fichier de log: %s", e)

    # --- Windows startup (HKCU Run) -----------------------------------------------------------
    def _apply_run_at_startup(self, enabled: bool) -> None:
        if os.name != "nt":
            return
        try:
            import winreg  # type: ignore
        except Exception:
            logging.warning("winreg non disponible: impossible de configurer le démarrage automatique.")
            return
        try:
            key_path = r"Software\Microsoft\Windows\CurrentVersion\Run"
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, key_path, 0, winreg.KEY_SET_VALUE) as key:
                name = "confirm-netflix-house"
                if enabled:
                    if getattr(sys, "frozen", False):
                        cmd = f'"{sys.executable}"'
                    else:
                        py = sys.executable.replace("/", "\\")
                        # Démarrer le module explicitement en dev
                        cmd = f'"{py}" -m src.tray_app'
                    winreg.SetValueEx(key, name, 0, winreg.REG_SZ, cmd)
                    logging.info("Démarrage automatique activé (%s)", cmd)
                else:
                    try:
                        winreg.DeleteValue(key, name)
                        logging.info("Démarrage automatique désactivé")
                    except FileNotFoundError:
                        pass
        except Exception as e:
            logging.warning("Configuration du démarrage automatique échouée: %s", e)

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
    _enforce_single_instance_if_frozen()
    atexit.register(_release_single_instance)
    app = TrayApp()
    app.run()


if __name__ == "__main__":
    main()
