"""server.py — Local HTTP API server, thumbnail cache, and generation job runner.

Binds strictly to 127.0.0.1 and serves the self-contained frontend from web/
alongside JSON endpoints for directory browsing, guarded scanning, config
persistence/remapping, native Windows dialogs, and Word/PDF generation.
"""

from collections import OrderedDict
from http import HTTPStatus
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
import io
import json
import math
import mimetypes
import os
from pathlib import Path
import sys
import threading
import uuid
from urllib.parse import parse_qs, urlparse

from PIL import Image, ImageOps

import config_store
import doc_generator
from scanner import ScanManager, is_image_filename, list_directory


MAX_BODY_BYTES = 20 * 1024 * 1024
ALLOWED_OPEN_EXTENSIONS = {".docx", ".pdf", ".html"}


def _parse_int(value, field_name: str) -> int:
    if value is None or isinstance(value, bool):
        raise ValueError(f"'{field_name}' must be an integer.")
    try:
        return int(value)
    except (TypeError, ValueError) as exc:
        raise ValueError(f"'{field_name}' must be an integer.") from exc


def _parse_optional_pos_float(value, field_name: str) -> float | None:
    if value in (None, ""):
        return None
    if isinstance(value, bool):
        raise ValueError(f"'{field_name}' must be a positive number.")
    try:
        parsed = float(value)
    except (TypeError, ValueError) as exc:
        raise ValueError(f"'{field_name}' must be a positive number.") from exc
    if not math.isfinite(parsed) or parsed <= 0:
        raise ValueError(f"'{field_name}' must be a positive number.")
    return parsed


def resource_path(relative_path: str) -> Path:
    base = getattr(sys, "_MEIPASS", None)
    if base:
        return Path(base) / relative_path
    return Path(__file__).resolve().parent / relative_path


class ThumbnailCache:
    """Thread-safe bounded in-memory thumbnail cache keyed by file state."""

    def __init__(self, max_entries: int = 300):
        self.max_entries = max_entries
        self._lock = threading.Lock()
        self._cache: OrderedDict[tuple, tuple[bytes, str]] = OrderedDict()

    def get_thumbnail(self, file_path: str, size: int = 120) -> tuple[bytes, str]:
        norm_path = os.path.normpath(file_path)
        if not is_image_filename(norm_path):
            raise ValueError("Unsupported image extension.")
        if not os.path.isfile(norm_path):
            raise FileNotFoundError(f"Image not found: {norm_path}")

        clamped_size = max(48, min(512, int(size)))
        st = os.stat(norm_path)
        key = (norm_path, st.st_mtime_ns, st.st_size, clamped_size)

        with self._lock:
            cached = self._cache.get(key)
            if cached is not None:
                self._cache.move_to_end(key)
                return cached

        with Image.open(norm_path) as img:
            img = ImageOps.exif_transpose(img)
            has_alpha = "A" in img.getbands() or "transparency" in img.info
            img = img.convert("RGBA" if has_alpha else "RGB")
            img.thumbnail((clamped_size, clamped_size), Image.Resampling.LANCZOS)

            buf = io.BytesIO()
            if has_alpha:
                img.save(buf, format="PNG", optimize=True)
                mime = "image/png"
            else:
                img.save(buf, format="JPEG", quality=82)
                mime = "image/jpeg"
            payload = (buf.getvalue(), mime)

        with self._lock:
            self._cache[key] = payload
            self._cache.move_to_end(key)
            while len(self._cache) > self.max_entries:
                self._cache.popitem(last=False)

        return payload


class GenerationManager:
    """Runs build_report_documents() in background threads and tracks progress."""

    MAX_JOBS = 20

    def __init__(self):
        self._lock = threading.Lock()
        self._jobs: dict[str, dict] = {}

    def start(
        self,
        sections: list,
        save_file: str,
        orientation: str = "landscape",
        num_cols: int = 2,
        max_width_cm: float | None = None,
        max_height_cm: float | None = None,
        export_pdf: bool = False,
        export_docx: bool = True,
        export_html: bool = False,
        html_embed_images: bool = False,
    ) -> str:
        if not save_file or not str(save_file).strip():
            raise ValueError("Output file path cannot be empty.")

        job_id = uuid.uuid4().hex[:12]
        job = {
            "job_id": job_id,
            "status": "running",
            "progress": 0.0,
            "message": "Starting document generation...",
            "result": None,
            "error": None,
        }
        with self._lock:
            self._jobs[job_id] = job
            if len(self._jobs) > self.MAX_JOBS:
                for old_id, old_job in list(self._jobs.items()):
                    if len(self._jobs) <= self.MAX_JOBS:
                        break
                    if old_id != job_id and old_job.get("status") in {"done", "error"}:
                        del self._jobs[old_id]

        def progress_cb(msg: str, pct: float) -> None:
            with self._lock:
                job["message"] = msg
                job["progress"] = max(0.0, min(1.0, float(pct)))

        def _normalize_gen_sections(raw_secs: list, depth: int = 0) -> list[dict]:
            default_pb = depth == 0
            out = []
            for sec in raw_secs or []:
                plots = []
                for item in sec.get("plots") or []:
                    p = item.get("path") if isinstance(item, dict) else str(item)
                    if p:
                        plots.append(p)
                subs = (
                    _normalize_gen_sections(sec.get("subsections") or [], depth + 1)
                    if depth < doc_generator.MAX_SUBSECTION_DEPTH
                    else []
                )
                out.append({
                    "title": sec.get("title", sec.get("name", "")),
                    "page_break": bool(sec["page_break"]) if "page_break" in sec else default_pb,
                    "plots": plots,
                    "subsections": subs,
                })
            return out

        def worker() -> None:
            try:
                normalized_sections = _normalize_gen_sections(sections or [], depth=0)

                result = doc_generator.build_report_documents(
                    normalized_sections,
                    save_file,
                    orientation=orientation,
                    num_cols=int(num_cols),
                    user_max_w_cm=max_width_cm,
                    user_max_h_cm=max_height_cm,
                    export_pdf=bool(export_pdf),
                    progress_callback=progress_cb,
                    export_docx=bool(export_docx),
                    export_html=bool(export_html),
                    html_embed_images=bool(html_embed_images),
                )
                with self._lock:
                    job["result"] = result
                    if result.get("success"):
                        job["status"] = "done"
                        job["progress"] = 1.0
                        job["message"] = f"Generated report with {result.get('count', 0)} plot(s)."
                    else:
                        job["status"] = "error"
                        skipped = result.get("skipped") or []
                        job["error"] = "\n".join(
                            msg if path == "<none>" else f"{path}: {msg}"
                            for path, msg in skipped
                        ) or "Document generation failed."
            except Exception as exc:
                with self._lock:
                    job["status"] = "error"
                    job["error"] = str(exc)

        threading.Thread(target=worker, daemon=True).start()
        return job_id

    def get_status(self, job_id: str) -> dict:
        with self._lock:
            job = self._jobs.get(job_id)
            if not job:
                raise KeyError(f"Unknown generation job: {job_id}")
            return dict(job)


class Plot2DocxServer(ThreadingHTTPServer):
    allow_reuse_address = True
    daemon_threads = True

    def __init__(
        self,
        server_address,
        RequestHandlerClass,
        *,
        web_dir: Path,
        dialog_bridge=None,
        auto_exit: bool = True,
        on_shutdown=None,
    ):
        self.web_dir = web_dir
        self.dialog_bridge = dialog_bridge
        self.auto_exit = auto_exit
        self.on_shutdown = on_shutdown
        self.scan_manager = ScanManager()
        self.gen_manager = GenerationManager()
        self.thumb_cache = ThumbnailCache(max_entries=300)
        self._lifecycle_lock = threading.Lock()
        self._shutdown_timer: threading.Timer | None = None
        super().__init__(server_address, RequestHandlerClass)

    def cancel_scheduled_shutdown(self) -> None:
        with self._lifecycle_lock:
            if self._shutdown_timer is not None:
                self._shutdown_timer.cancel()
                self._shutdown_timer = None

    def schedule_shutdown(self, delay_seconds: float = 5.0) -> None:
        if not self.auto_exit or self.on_shutdown is None:
            return
        with self._lifecycle_lock:
            if self._shutdown_timer is not None:
                self._shutdown_timer.cancel()
            timer = threading.Timer(delay_seconds, self.on_shutdown)
            timer.daemon = True
            timer.start()
            self._shutdown_timer = timer


class RequestHandler(BaseHTTPRequestHandler):
    protocol_version = "HTTP/1.1"

    def log_message(self, format, *args):
        # Keep console output clean unless debugging
        pass

    def _send_bytes(self, status: int, data: bytes, content_type: str, cache_control: str = "no-store") -> None:
        try:
            self.send_response(status)
            self.send_header("Content-Type", content_type)
            self.send_header("Content-Length", str(len(data)))
            self.send_header("Cache-Control", cache_control)
            self.end_headers()
            self.wfile.write(data)
        except (BrokenPipeError, ConnectionAbortedError, ConnectionResetError):
            pass

    def _send_json(self, status: int, payload: dict) -> None:
        raw = json.dumps(payload).encode("utf-8")
        self._send_bytes(status, raw, "application/json; charset=utf-8")

    def _check_origin(self) -> bool:
        origin = self.headers.get("Origin")
        if origin is None:
            return True
        port = self.server.server_address[1]
        allowed = {f"http://127.0.0.1:{port}", f"http://localhost:{port}"}
        if origin not in allowed:
            self._send_json(HTTPStatus.FORBIDDEN, {"error": "Cross-origin requests are not allowed."})
            return False
        return True

    def _read_json_body(self) -> dict:
        raw_len = self.headers.get("Content-Length")
        if raw_len is None:
            return {}
        try:
            length = int(raw_len)
        except ValueError as exc:
            raise ValueError("Invalid Content-Length header.") from exc
        if length < 0 or length > MAX_BODY_BYTES:
            raise ValueError("Invalid or oversized request body.")
        if length == 0:
            return {}
        raw = self.rfile.read(length)
        if not raw:
            return {}
        data = json.loads(raw.decode("utf-8"))
        if not isinstance(data, dict):
            raise ValueError("JSON body must be an object.")
        return data

    def _serve_static(self, rel_name: str) -> None:
        web_dir: Path = self.server.web_dir
        web_root = web_dir.resolve()
        target = (web_dir / rel_name).resolve()
        if not target.is_relative_to(web_root) or not target.is_file():
            self._send_json(HTTPStatus.NOT_FOUND, {"error": "Static file not found."})
            return
        mime, _ = mimetypes.guess_type(str(target))
        if target.suffix == ".js":
            mime = "application/javascript; charset=utf-8"
        elif target.suffix == ".css":
            mime = "text/css; charset=utf-8"
        elif target.suffix == ".html":
            mime = "text/html; charset=utf-8"
        elif not mime:
            mime = "application/octet-stream"
        self._send_bytes(HTTPStatus.OK, target.read_bytes(), mime)

    def do_GET(self):
        parsed = urlparse(self.path)
        path = parsed.path
        qs = parse_qs(parsed.query)

        try:
            if path in ("/", "/index.html"):
                self._serve_static("index.html")
                return
            if path in ("/styles.css", "/app.js"):
                self._serve_static(path.lstrip("/"))
                return

            if path.startswith("/api/") and not self._check_origin():
                return

            if path == "/api/health":
                self._send_json(HTTPStatus.OK, {
                    "ok": True,
                    "has_win32com": doc_generator.HAS_WIN32COM,
                    "has_dialogs": self.server.dialog_bridge is not None,
                })
                return

            if path == "/api/fs/list":
                target_path = (qs.get("path") or [""])[0]
                listing = list_directory(target_path)
                self._send_json(HTTPStatus.OK, listing)
                return

            if path == "/api/scan/status":
                scan_id = (qs.get("id") or [""])[0]
                since = _parse_int((qs.get("since") or ["0"])[0], "since")
                status = self.server.scan_manager.get_status(scan_id, since_index=since)
                self._send_json(HTTPStatus.OK, status)
                return

            if path == "/api/thumb":
                img_path = (qs.get("path") or [""])[0]
                size = _parse_int((qs.get("size") or ["120"])[0], "size")
                data, mime = self.server.thumb_cache.get_thumbnail(img_path, size=size)
                self._send_bytes(HTTPStatus.OK, data, mime, cache_control="private, max-age=300")
                return

            if path == "/api/config/recent":
                self._send_json(HTTPStatus.OK, {"recent": config_store.get_recent_configs()})
                return

            if path == "/api/generate/status":
                job_id = (qs.get("id") or [""])[0]
                status = self.server.gen_manager.get_status(job_id)
                self._send_json(HTTPStatus.OK, status)
                return

            self._send_json(HTTPStatus.NOT_FOUND, {"error": f"Unknown endpoint: {path}"})

        except FileNotFoundError as exc:
            self._send_json(HTTPStatus.NOT_FOUND, {"error": str(exc)})
        except NotADirectoryError as exc:
            self._send_json(HTTPStatus.BAD_REQUEST, {"error": str(exc)})
        except KeyError as exc:
            self._send_json(HTTPStatus.NOT_FOUND, {"error": str(exc)})
        except (ValueError, TypeError) as exc:
            self._send_json(HTTPStatus.BAD_REQUEST, {"error": str(exc)})
        except Exception as exc:
            self._send_json(HTTPStatus.INTERNAL_SERVER_ERROR, {"error": str(exc)})

    def do_POST(self):
        parsed = urlparse(self.path)
        path = parsed.path

        try:
            if not self._check_origin():
                return
            body = self._read_json_body()

            if path == "/api/ping":
                self.server.cancel_scheduled_shutdown()
                self._send_json(HTTPStatus.OK, {"ok": True})
                return

            if path == "/api/bye":
                self.server.schedule_shutdown(delay_seconds=5.0)
                self._send_json(HTTPStatus.OK, {"ok": True})
                return

            if path == "/api/dialog/browse-folder":
                bridge = self.server.dialog_bridge
                if bridge is None:
                    self._send_json(HTTPStatus.NOT_IMPLEMENTED, {"error": "Desktop dialog host is not available."})
                    return
                selected = bridge.browse_folder(initial_dir=body.get("initial_dir", ""))
                self._send_json(HTTPStatus.OK, {"path": selected, "cancelled": not selected})
                return

            if path == "/api/dialog/save-docx":
                bridge = self.server.dialog_bridge
                if bridge is None:
                    self._send_json(HTTPStatus.NOT_IMPLEMENTED, {"error": "Desktop dialog host is not available."})
                    return
                default_ext = str(body.get("default_ext", ".docx")).lower()
                if default_ext not in {".docx", ".html"}:
                    raise ValueError("'default_ext' must be '.docx' or '.html'.")
                selected = bridge.save_docx(
                    initial_dir=body.get("initial_dir", ""),
                    initial_file=body.get("initial_file", ""),
                    default_ext=default_ext,
                )
                self._send_json(HTTPStatus.OK, {"path": selected, "cancelled": not selected})
                return

            if path == "/api/dialog/save-config":
                bridge = self.server.dialog_bridge
                if bridge is None:
                    self._send_json(HTTPStatus.NOT_IMPLEMENTED, {"error": "Desktop dialog host is not available."})
                    return
                selected = bridge.save_config(
                    initial_dir=body.get("initial_dir", ""),
                    initial_file=body.get("initial_file", ""),
                )
                self._send_json(HTTPStatus.OK, {"path": selected, "cancelled": not selected})
                return

            if path == "/api/dialog/open-config":
                bridge = self.server.dialog_bridge
                if bridge is None:
                    self._send_json(HTTPStatus.NOT_IMPLEMENTED, {"error": "Desktop dialog host is not available."})
                    return
                selected = bridge.open_config(initial_dir=body.get("initial_dir", ""))
                self._send_json(HTTPStatus.OK, {"path": selected, "cancelled": not selected})
                return

            if path == "/api/scan/start":
                root_path = body.get("path", "")
                soft_limit = _parse_int(body.get("soft_limit", 500), "soft_limit")
                scan_id = self.server.scan_manager.start_scan(root_path, soft_limit=soft_limit)
                self._send_json(HTTPStatus.OK, {"scan_id": scan_id})
                return

            if path == "/api/scan/continue":
                scan_id = body.get("scan_id", "")
                new_limit = _parse_int(body.get("new_limit", 0), "new_limit")
                keep_loaded = bool(body.get("keep_loaded", False))
                status = self.server.scan_manager.continue_scan(scan_id, new_limit=new_limit, keep_loaded=keep_loaded)
                self._send_json(HTTPStatus.OK, status)
                return

            if path == "/api/scan/cancel":
                scan_id = body.get("scan_id", "")
                status = self.server.scan_manager.cancel_scan(scan_id)
                self._send_json(HTTPStatus.OK, status)
                return

            if path == "/api/files/validate":
                paths = body.get("paths") or []
                if not isinstance(paths, list):
                    raise ValueError("'paths' must be a list.")
                self._send_json(HTTPStatus.OK, {"file_status": config_store.validate_files(paths)})
                return

            if path == "/api/config/save":
                file_path = body.get("path", "")
                state = body.get("state") or {}
                saved = config_store.save_config(file_path, state)
                self._send_json(HTTPStatus.OK, saved)
                return

            if path == "/api/config/load":
                file_path = body.get("path", "")
                loaded = config_store.load_config(file_path)
                self._send_json(HTTPStatus.OK, loaded)
                return

            if path == "/api/config/remap":
                sections = body.get("sections") or []
                new_source = body.get("new_source_folder", "")
                old_source = body.get("old_source_folder", "")
                remapped = config_store.remap_config_source(
                    sections, new_source, old_source_folder=old_source
                )
                self._send_json(HTTPStatus.OK, remapped)
                return

            if path == "/api/generate/start":
                sections = body.get("sections") or []
                save_file = body.get("save_file", "")
                orientation = str(body.get("orientation", "landscape")).lower()
                if orientation not in {"landscape", "portrait"}:
                    raise ValueError("'orientation' must be 'landscape' or 'portrait'.")
                num_cols = _parse_int(body.get("num_cols", 2), "num_cols")
                if num_cols not in (1, 2, 3):
                    raise ValueError("'num_cols' must be 1, 2, or 3.")
                max_w_val = _parse_optional_pos_float(body.get("max_width_cm"), "max_width_cm")
                max_h_val = _parse_optional_pos_float(body.get("max_height_cm"), "max_height_cm")
                export_docx = bool(body.get("export_docx", True))
                export_html = bool(body.get("export_html", False))
                html_embed_images = bool(body.get("html_embed_images", False))
                if not export_docx and not export_html:
                    raise ValueError("Select at least one output format (.docx or .html).")
                export_pdf = bool(body.get("export_pdf", False)) and export_docx

                job_id = self.server.gen_manager.start(
                    sections=sections,
                    save_file=save_file,
                    orientation=orientation,
                    num_cols=num_cols,
                    max_width_cm=max_w_val,
                    max_height_cm=max_h_val,
                    export_pdf=export_pdf,
                    export_docx=export_docx,
                    export_html=export_html,
                    html_embed_images=html_embed_images,
                )
                self._send_json(HTTPStatus.OK, {"job_id": job_id})
                return

            if path == "/api/open-file":
                target = body.get("path", "")
                if not target or not isinstance(target, str) or not os.path.exists(target):
                    raise FileNotFoundError(f"Target does not exist: {target}")
                abs_target = os.path.abspath(target)
                if (
                    not os.path.isdir(abs_target)
                    and Path(abs_target).suffix.lower() not in ALLOWED_OPEN_EXTENSIONS
                ):
                    raise ValueError("Only generated .docx/.html/.pdf files and folders can be opened.")
                os.startfile(abs_target)
                self._send_json(HTTPStatus.OK, {"ok": True})
                return

            self._send_json(HTTPStatus.NOT_FOUND, {"error": f"Unknown endpoint: {path}"})

        except FileNotFoundError as exc:
            self._send_json(HTTPStatus.NOT_FOUND, {"error": str(exc)})
        except NotADirectoryError as exc:
            self._send_json(HTTPStatus.BAD_REQUEST, {"error": str(exc)})
        except KeyError as exc:
            self._send_json(HTTPStatus.NOT_FOUND, {"error": str(exc)})
        except (ValueError, TypeError) as exc:
            self._send_json(HTTPStatus.BAD_REQUEST, {"error": str(exc)})
        except Exception as exc:
            self._send_json(HTTPStatus.INTERNAL_SERVER_ERROR, {"error": str(exc)})


def create_server(
    host: str = "127.0.0.1",
    port: int = 0,
    dialog_bridge=None,
    auto_exit: bool = True,
    on_shutdown=None,
) -> Plot2DocxServer:
    web_dir = resource_path("web")
    return Plot2DocxServer(
        (host, port),
        RequestHandler,
        web_dir=web_dir,
        dialog_bridge=dialog_bridge,
        auto_exit=auto_exit,
        on_shutdown=on_shutdown,
    )
