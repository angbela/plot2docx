"""app.py — Local Web App launcher & native Windows file dialog host.

Starts a local HTTP server bound to 127.0.0.1, opens the modern offline web UI
in an app-mode Edge/Chrome window (falling back to the default browser), and
services native Windows folder/file picker dialogs from the main thread.
"""

import argparse
import datetime
import os
from pathlib import Path
import queue
import shutil
import subprocess
import threading
import webbrowser

try:
    import tkinter as tk
    from tkinter import filedialog
    HAS_TK = True
except ImportError:
    HAS_TK = False

from server import create_server


class TkDialogHost:
    """Thread-safe bridge that executes tkinter filedialog calls on the main thread."""

    def __init__(self):
        self._queue: queue.Queue = queue.Queue()

    def _call(self, fn):
        if not HAS_TK:
            raise RuntimeError("Tkinter file dialogs are unavailable in this environment.")
        done = threading.Event()
        box = {"result": "", "error": None}
        self._queue.put((fn, box, done))
        done.wait()
        if box["error"] is not None:
            raise box["error"]
        return box["result"]

    def process_pending(self, timeout: float = 0.15) -> None:
        try:
            fn, box, done = self._queue.get(timeout=timeout)
        except queue.Empty:
            return
        root = None
        try:
            root = tk.Tk()
            root.withdraw()
            root.attributes("-topmost", True)
            root.lift()
            root.focus_force()
            box["result"] = fn(root) or ""
        except Exception as exc:
            box["error"] = exc
        finally:
            if root is not None:
                try:
                    root.destroy()
                except Exception:
                    pass
            done.set()

    def browse_folder(self, initial_dir: str = "") -> str:
        def _run(root):
            kw = {"parent": root, "title": "Select Plots Folder"}
            if initial_dir and os.path.isdir(initial_dir):
                kw["initialdir"] = initial_dir
            selected = filedialog.askdirectory(**kw)
            return os.path.normpath(selected) if selected else ""

        return self._call(_run)

    def save_docx(self, initial_dir: str = "", initial_file: str = "", default_ext: str = ".docx") -> str:
        ext = ".html" if str(default_ext).lower() == ".html" else ".docx"

        def _run(root):
            if ext == ".html":
                default_name = initial_file or f"Plot_Report_{datetime.datetime.now().strftime('%Y%m%d-%H%M')}.html"
                kw = {
                    "parent": root,
                    "title": "Save Generated HTML Report",
                    "defaultextension": ".html",
                    "filetypes": [("HTML Report", "*.html")],
                    "initialfile": default_name,
                }
            else:
                default_name = initial_file or f"Plot_Report_{datetime.datetime.now().strftime('%Y%m%d-%H%M')}.docx"
                kw = {
                    "parent": root,
                    "title": "Save Generated Report",
                    "defaultextension": ".docx",
                    "filetypes": [("Word Document", "*.docx")],
                    "initialfile": default_name,
                }
            if initial_dir and os.path.isdir(initial_dir):
                kw["initialdir"] = initial_dir
            selected = filedialog.asksaveasfilename(**kw)
            if selected and not selected.lower().endswith(ext):
                selected += ext
            return os.path.normpath(selected) if selected else ""

        return self._call(_run)

    def save_config(self, initial_dir: str = "", initial_file: str = "") -> str:
        def _run(root):
            default_name = initial_file or f"Report_Config_{datetime.datetime.now().strftime('%Y%m%d')}.plot2docx.json"
            kw = {
                "parent": root,
                "title": "Save Plot2Docx Configuration",
                "defaultextension": ".plot2docx.json",
                "filetypes": [("Plot2Docx Config", "*.plot2docx.json *.json"), ("All Files", "*.*")],
                "initialfile": default_name,
            }
            if initial_dir and os.path.isdir(initial_dir):
                kw["initialdir"] = initial_dir
            selected = filedialog.asksaveasfilename(**kw)
            return os.path.normpath(selected) if selected else ""

        return self._call(_run)

    def open_config(self, initial_dir: str = "") -> str:
        def _run(root):
            kw = {
                "parent": root,
                "title": "Load Plot2Docx Configuration",
                "filetypes": [("Plot2Docx Config", "*.plot2docx.json *.json"), ("All Files", "*.*")],
            }
            if initial_dir and os.path.isdir(initial_dir):
                kw["initialdir"] = initial_dir
            selected = filedialog.askopenfilename(**kw)
            return os.path.normpath(selected) if selected else ""

        return self._call(_run)


def _find_chromium_app_browser() -> str | None:
    candidates = [
        shutil.which("msedge"),
        shutil.which("chrome"),
        r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
        r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
        r"C:\Program Files\Google\Chrome\Application\chrome.exe",
        r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
        str(Path(os.environ.get("LOCALAPPDATA", "")) / r"Google\Chrome\Application\chrome.exe"),
    ]
    for path in candidates:
        if path and os.path.isfile(path):
            return path
    return None


def launch_ui(url: str, force_browser: bool = False) -> None:
    if not force_browser:
        browser_exe = _find_chromium_app_browser()
        if browser_exe:
            try:
                subprocess.Popen(
                    [browser_exe, f"--app={url}", "--window-size=1280,880", "--new-window"],
                    stdout=subprocess.DEVNULL,
                    stderr=subprocess.DEVNULL,
                )
                return
            except Exception:
                pass
    webbrowser.open(url)


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Plot to Docx local web application")
    parser.add_argument("--port", type=int, default=0, help="Port to bind on 127.0.0.1 (0 = ephemeral)")
    parser.add_argument("--no-open", action="store_true", help="Start server without opening a window")
    parser.add_argument("--no-auto-exit", action="store_true", help="Keep server running after UI closes")
    parser.add_argument("--browser", action="store_true", help="Open in default browser tab instead of app window")
    args = parser.parse_args(argv)

    shutdown_event = threading.Event()
    dialog_host = TkDialogHost() if HAS_TK else None

    server = create_server(
        host="127.0.0.1",
        port=args.port,
        dialog_bridge=dialog_host,
        auto_exit=not args.no_auto_exit,
        on_shutdown=shutdown_event.set,
    )
    actual_port = server.server_address[1]
    url = f"http://127.0.0.1:{actual_port}/"

    server_thread = threading.Thread(target=server.serve_forever, daemon=True)
    server_thread.start()

    if not args.no_open:
        launch_ui(url, force_browser=args.browser)
    else:
        print(f"Plot to Docx running at {url}", flush=True)

    try:
        while not shutdown_event.is_set():
            if dialog_host is not None:
                dialog_host.process_pending(timeout=0.15)
            else:
                shutdown_event.wait(0.15)
    except KeyboardInterrupt:
        pass
    finally:
        server.shutdown()
        server.server_close()

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
