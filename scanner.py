"""scanner.py — Fast lazy folder browser & cancellable guarded recursive scanner.

Uses os.scandir() so local drives and Windows UNC network shares (\\\\Server\\...)
enumerate cleanly without redundant stat round-trips, supports instant scan
cancellation, and pauses at a configurable soft limit to prevent accidentally
loading huge directories.
"""

from collections import deque
from dataclasses import dataclass, field
import os
from pathlib import Path
import threading
import uuid


IMAGE_EXTENSIONS = {".png", ".jpg", ".jpeg", ".bmp", ".tiff", ".tif", ".gif", ".webp"}


def format_size(size_in_bytes: int) -> str:
    size = float(size_in_bytes)
    for unit in ["B", "KB", "MB", "GB"]:
        if size < 1024.0:
            return f"{size:.1f} {unit}"
        size /= 1024.0
    return f"{size:.1f} TB"


def is_image_filename(name: str) -> bool:
    return Path(name).suffix.lower() in IMAGE_EXTENSIONS


def list_directory(path: str) -> dict:
    """List immediate subfolders and image files in a single directory."""
    if not path or not path.strip():
        raise ValueError("Folder path cannot be empty.")

    norm_path = os.path.normpath(path.strip())
    if not os.path.exists(norm_path):
        raise FileNotFoundError(f"Path does not exist or is not accessible: {norm_path}")
    if not os.path.isdir(norm_path):
        raise NotADirectoryError(f"Path is not a folder: {norm_path}")

    folders = []
    files = []

    with os.scandir(norm_path) as it:
        for entry in it:
            try:
                if entry.is_dir(follow_symlinks=False):
                    folders.append({
                        "name": entry.name,
                        "path": entry.path,
                    })
                elif entry.is_file(follow_symlinks=False) and is_image_filename(entry.name):
                    st = entry.stat(follow_symlinks=False)
                    ext = Path(entry.name).suffix.upper().lstrip(".")
                    files.append({
                        "name": entry.name,
                        "path": entry.path,
                        "rel_path": entry.name,
                        "folder": norm_path,
                        "ext": ext,
                        "size": st.st_size,
                        "size_str": format_size(st.st_size),
                        "mtime": st.st_mtime,
                    })
            except OSError:
                continue

    folders.sort(key=lambda item: item["name"].lower())
    files.sort(key=lambda item: item["name"].lower())

    parent_dir = os.path.dirname(norm_path)
    parent = parent_dir if parent_dir and parent_dir != norm_path else None

    return {
        "path": norm_path,
        "name": os.path.basename(norm_path) or norm_path,
        "parent": parent,
        "folders": folders,
        "files": files,
    }


@dataclass
class _ScanJob:
    scan_id: str
    root: str
    soft_limit: int
    status: str = "running"  # running | paused_limit | done | cancelled | error
    files: list = field(default_factory=list)
    total_dirs: int = 1
    scanned_dirs: int = 0
    error: str | None = None
    stop_at_limit: bool = False
    cancel_event: threading.Event = field(default_factory=threading.Event)
    continue_event: threading.Event = field(default_factory=threading.Event)


class ScanManager:
    """Manages cancellable recursive image scans with soft-limit guardrails."""

    MAX_JOBS = 20

    def __init__(self):
        self._lock = threading.Lock()
        self._jobs: dict[str, _ScanJob] = {}

    def start_scan(self, root_dir: str, soft_limit: int = 500) -> str:
        if not root_dir or not root_dir.strip():
            raise ValueError("Folder path cannot be empty.")

        norm_root = os.path.normpath(root_dir.strip())
        if not os.path.exists(norm_root):
            raise FileNotFoundError(f"Path does not exist or is not accessible: {norm_root}")
        if not os.path.isdir(norm_root):
            raise NotADirectoryError(f"Path is not a folder: {norm_root}")

        scan_id = uuid.uuid4().hex[:12]
        job = _ScanJob(scan_id=scan_id, root=norm_root, soft_limit=max(0, int(soft_limit)))
        with self._lock:
            self._jobs[scan_id] = job
            if len(self._jobs) > self.MAX_JOBS:
                for old_id, old_job in list(self._jobs.items()):
                    if len(self._jobs) <= self.MAX_JOBS:
                        break
                    if old_id != scan_id and old_job.status in {"done", "cancelled", "error"}:
                        del self._jobs[old_id]

        thread = threading.Thread(target=self._run_scan, args=(job,), daemon=True)
        thread.start()
        return scan_id

    def _run_scan(self, job: _ScanJob) -> None:
        queue = deque([job.root])
        root_len = len(job.root.rstrip("\\/"))

        try:
            while queue:
                if job.cancel_event.is_set():
                    with self._lock:
                        job.status = "cancelled"
                    return

                current_dir = queue.popleft()
                subdirs = []
                dir_files = []

                try:
                    with os.scandir(current_dir) as it:
                        for entry in it:
                            if job.cancel_event.is_set():
                                with self._lock:
                                    job.status = "cancelled"
                                return
                            try:
                                if entry.is_dir(follow_symlinks=False):
                                    subdirs.append(entry.path)
                                elif entry.is_file(follow_symlinks=False) and is_image_filename(entry.name):
                                    st = entry.stat(follow_symlinks=False)
                                    ext = Path(entry.name).suffix.upper().lstrip(".")
                                    rel = entry.path[root_len:].lstrip("\\/") if entry.path.startswith(job.root) else os.path.relpath(entry.path, job.root)
                                    dir_files.append({
                                        "name": entry.name,
                                        "path": entry.path,
                                        "rel_path": rel or entry.name,
                                        "folder": current_dir,
                                        "ext": ext,
                                        "size": st.st_size,
                                        "size_str": format_size(st.st_size),
                                        "mtime": st.st_mtime,
                                    })
                            except OSError:
                                continue
                except OSError:
                    with self._lock:
                        job.scanned_dirs += 1
                    continue

                subdirs.sort(key=lambda p: os.path.basename(p).lower())
                dir_files.sort(key=lambda item: item["name"].lower())

                with self._lock:
                    job.scanned_dirs += 1
                    job.total_dirs += len(subdirs)
                queue.extend(subdirs)

                for file_item in dir_files:
                    if job.cancel_event.is_set():
                        with self._lock:
                            job.status = "cancelled"
                        return

                    should_pause = False
                    with self._lock:
                        if job.soft_limit > 0 and len(job.files) >= job.soft_limit:
                            job.status = "paused_limit"
                            job.continue_event.clear()
                            should_pause = True
                        else:
                            job.files.append(file_item)

                    if should_pause:
                        job.continue_event.wait()
                        if job.cancel_event.is_set():
                            with self._lock:
                                job.status = "cancelled"
                            return
                        with self._lock:
                            if job.stop_at_limit:
                                job.status = "done"
                                return
                            job.files.append(file_item)

            with self._lock:
                if job.cancel_event.is_set():
                    job.status = "cancelled"
                else:
                    job.status = "done"

        except Exception as exc:
            with self._lock:
                job.status = "error"
                job.error = str(exc)

    def continue_scan(self, scan_id: str, new_limit: int = 0, keep_loaded: bool = False) -> dict:
        with self._lock:
            job = self._jobs.get(scan_id)
            if not job:
                raise KeyError(f"Unknown scan_id: {scan_id}")
            if keep_loaded:
                job.stop_at_limit = True
                job.status = "done"
            else:
                job.soft_limit = max(0, int(new_limit))
                if job.status == "paused_limit":
                    job.status = "running"
            job.continue_event.set()
        return self.get_status(scan_id)

    def cancel_scan(self, scan_id: str) -> dict:
        with self._lock:
            job = self._jobs.get(scan_id)
            if not job:
                raise KeyError(f"Unknown scan_id: {scan_id}")
            job.cancel_event.set()
            job.status = "cancelled"
            job.continue_event.set()
        return self.get_status(scan_id)

    def get_status(self, scan_id: str, since_index: int = 0) -> dict:
        with self._lock:
            job = self._jobs.get(scan_id)
            if not job:
                raise KeyError(f"Unknown scan_id: {scan_id}")
            start = max(0, int(since_index))
            new_files = list(job.files[start:])
            total_files = len(job.files)
            return {
                "scan_id": job.scan_id,
                "status": job.status,
                "root": job.root,
                "total_files": total_files,
                "total_dirs": job.total_dirs,
                "scanned_dirs": job.scanned_dirs,
                "soft_limit": job.soft_limit,
                "next_index": total_files,
                "new_files": new_files,
                "error": job.error,
            }
