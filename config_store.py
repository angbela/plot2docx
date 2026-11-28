"""config_store.py — Save, load, validate, and remap Plot to Docx project configs.

Allows saving section structures, ordering, page breaks, and document layout
options into portable JSON files so reports can be regenerated whenever plots
are replotted under the same filenames or a new base folder.
"""

import datetime
import json
import math
import os
from pathlib import Path

from scanner import format_size, is_image_filename


SCHEMA_VERSION = 1
RECENT_MAX = 10
MAX_SUBSECTION_DEPTH = 2


def _recent_file_path() -> Path:
    local_app = os.environ.get("LOCALAPPDATA")
    if local_app:
        return Path(local_app) / "PlotToDocx" / "recent.json"
    return Path.home() / ".plot2docx" / "recent.json"


def _sanitize_rel_path(rel_path: str | None) -> str | None:
    if rel_path is None:
        return None
    raw = str(rel_path).strip()
    if not raw or os.path.isabs(raw) or os.path.splitdrive(raw)[0] or raw.startswith(("\\", "/")):
        return None
    parts = [p for p in raw.replace("\\", "/").split("/") if p and p != "."]
    if not parts or any(p == ".." for p in parts):
        return None
    norm = os.path.normpath(raw)
    if norm in (".", "") or norm.startswith("..") or os.path.isabs(norm) or os.path.splitdrive(norm)[0]:
        return None
    return norm


def _safe_rel_path(plot_path: str, source_folder: str) -> str | None:
    if not source_folder:
        return None
    try:
        rel = os.path.relpath(os.path.normpath(plot_path), os.path.normpath(source_folder))
        if rel == "." or rel.startswith("..") or os.path.isabs(rel):
            return None
        return rel
    except ValueError:
        return None


def _parse_optional_pos_float(value) -> float | None:
    if value in (None, "") or isinstance(value, bool):
        return None
    try:
        parsed = float(value)
    except (TypeError, ValueError):
        return None
    if not math.isfinite(parsed) or parsed <= 0:
        return None
    return parsed


def _parse_num_cols(value) -> int:
    if value in (None, "") or isinstance(value, bool):
        return 2
    try:
        cols = int(value)
    except (TypeError, ValueError):
        return 2
    return cols if cols in (1, 2, 3) else 2


def validate_files(paths: list[str]) -> dict:
    """Return existence, size, and modification timestamps for a list of paths."""
    status = {}
    for raw_path in paths:
        p = str(raw_path)
        if p in status:
            continue
        try:
            st = os.stat(p)
            if os.path.isfile(p):
                mtime_str = datetime.datetime.fromtimestamp(st.st_mtime).strftime("%Y-%m-%d %H:%M:%S")
                status[p] = {
                    "exists": True,
                    "size": st.st_size,
                    "size_str": format_size(st.st_size),
                    "mtime": st.st_mtime,
                    "mtime_str": mtime_str,
                }
            else:
                status[p] = {
                    "exists": False,
                    "size": None,
                    "size_str": "Missing",
                    "mtime": None,
                    "mtime_str": "",
                }
        except OSError:
            status[p] = {
                "exists": False,
                "size": None,
                "size_str": "Missing",
                "mtime": None,
                "mtime_str": "",
            }
    return status


def iter_section_plot_paths(sections: list, depth: int = 0) -> list[str]:
    """Collect all plot paths across sections and nested subsections."""
    paths: list[str] = []
    for sec in sections or []:
        for item in sec.get("plots") or []:
            p = item.get("path") if isinstance(item, dict) else str(item)
            if p:
                paths.append(str(p))
        if depth < MAX_SUBSECTION_DEPTH:
            paths.extend(iter_section_plot_paths(sec.get("subsections") or [], depth + 1))
    return paths


def _normalize_sections(raw_sections: list, source_folder: str, depth: int = 0) -> list[dict]:
    sections_out: list[dict] = []
    default_pb = depth == 0
    for idx, sec in enumerate(raw_sections or [], start=1):
        plots_out = []
        for item in sec.get("plots") or []:
            if isinstance(item, dict):
                raw_plot = str(item.get("path", "")).strip()
                raw_rel = item.get("rel_path")
            else:
                raw_plot = str(item).strip()
                raw_rel = None
            if not raw_plot:
                continue
            plot_path = os.path.normpath(raw_plot)
            rel_path = _sanitize_rel_path(raw_rel) or _safe_rel_path(plot_path, source_folder)
            plots_out.append({"path": plot_path, "rel_path": rel_path})

        subsections_out = (
            _normalize_sections(sec.get("subsections") or [], source_folder, depth + 1)
            if depth < MAX_SUBSECTION_DEPTH
            else []
        )
        sections_out.append({
            "id": str(sec.get("id") or f"sec_{idx}"),
            "name": str(sec.get("name", f"Section {idx}")),
            "page_break": bool(sec["page_break"]) if "page_break" in sec else default_pb,
            "plots": plots_out,
            "subsections": subsections_out,
        })
    return sections_out


def build_config_payload(state: dict) -> dict:
    """Normalize UI state into the persisted schema_version=1 config structure."""
    source_folder = (state.get("source_folder") or "").strip()
    if source_folder:
        source_folder = os.path.normpath(source_folder)

    sections_out = _normalize_sections(state.get("sections") or [], source_folder, depth=0)

    doc_opts_in = state.get("doc_options") or {}
    doc_options = {
        "orientation": "portrait" if str(doc_opts_in.get("orientation", "landscape")).lower() == "portrait" else "landscape",
        "num_cols": _parse_num_cols(doc_opts_in.get("num_cols")),
        "max_width_cm": _parse_optional_pos_float(doc_opts_in.get("max_width_cm")),
        "max_height_cm": _parse_optional_pos_float(doc_opts_in.get("max_height_cm")),
        "export_docx": bool(doc_opts_in.get("export_docx", True)),
        "export_html": bool(doc_opts_in.get("export_html", False)),
        "html_embed_images": bool(doc_opts_in.get("html_embed_images", False)),
        "export_pdf": bool(doc_opts_in.get("export_pdf", False)),
        "last_output_path": doc_opts_in.get("last_output_path") or None,
    }

    grouping_mode = str(state.get("grouping_mode") or "manual").lower()
    if grouping_mode not in {"manual", "auto", "flat"}:
        grouping_mode = "manual"

    return {
        "schema_version": SCHEMA_VERSION,
        "saved_at": datetime.datetime.now().isoformat(timespec="seconds"),
        "name": str(state.get("name") or "Plot2Docx Report"),
        "source_folder": source_folder,
        "grouping_mode": grouping_mode,
        "sections": sections_out,
        "doc_options": doc_options,
    }


def _atomic_write_json(target: Path, payload: dict | list) -> None:
    target.parent.mkdir(parents=True, exist_ok=True)
    tmp = target.with_name(f".{target.name}.{os.getpid()}.tmp")
    tmp.write_text(json.dumps(payload, indent=2), encoding="utf-8")
    os.replace(tmp, target)


def get_recent_configs() -> list[dict]:
    recent_path = _recent_file_path()
    try:
        data = json.loads(recent_path.read_text(encoding="utf-8"))
        if not isinstance(data, list):
            return []
    except Exception:
        return []

    filtered = []
    for entry in data:
        if not isinstance(entry, dict):
            continue
        p = entry.get("path")
        if p and os.path.isfile(p):
            filtered.append(entry)
    return filtered[:RECENT_MAX]


def record_recent_config(file_path: str, name: str = "", plot_count: int = 0) -> None:
    abs_path = os.path.abspath(file_path)
    entries = [e for e in get_recent_configs() if os.path.abspath(e.get("path", "")) != abs_path]
    entries.insert(0, {
        "path": abs_path,
        "name": name or Path(abs_path).stem,
        "plot_count": int(plot_count),
        "opened_at": datetime.datetime.now().isoformat(timespec="seconds"),
    })
    try:
        _atomic_write_json(_recent_file_path(), entries[:RECENT_MAX])
    except OSError:
        pass


def save_config(file_path: str, state: dict) -> dict:
    if not file_path or not str(file_path).strip():
        raise ValueError("Config file path cannot be empty.")

    target = Path(file_path).expanduser()
    if not target.suffix:
        target = target.with_suffix(".plot2docx.json")

    if not state.get("name"):
        state = {**state, "name": target.stem.replace(".plot2docx", "")}

    payload = build_config_payload(state)
    _atomic_write_json(target, payload)

    all_paths = iter_section_plot_paths(payload["sections"])
    record_recent_config(str(target), payload["name"], len(all_paths))

    return {
        "path": str(target.resolve()),
        "config": payload,
        "file_status": validate_files(all_paths),
    }


def load_config(file_path: str) -> dict:
    if not file_path or not str(file_path).strip():
        raise ValueError("Config file path cannot be empty.")

    target = Path(file_path).expanduser()
    if not target.is_file():
        raise FileNotFoundError(f"Config file not found: {target}")

    raw = json.loads(target.read_text(encoding="utf-8"))
    if not isinstance(raw, dict):
        raise ValueError("Invalid config file format.")
    version = raw.get("schema_version")
    if version != SCHEMA_VERSION:
        raise ValueError(f"Unsupported config schema_version: {version}")

    payload = build_config_payload(raw)
    all_paths = iter_section_plot_paths(payload["sections"])
    file_status = validate_files(all_paths)

    plot_count = len(all_paths)
    record_recent_config(str(target), payload["name"], plot_count)

    return {
        "path": str(target.resolve()),
        "config": payload,
        "file_status": file_status,
    }


def _build_basename_index(root_dir: str) -> dict[str, list[str]]:
    """Index image files under root_dir by lowercase basename for fallback remapping."""
    index: dict[str, list[str]] = {}
    stack = [root_dir]
    while stack:
        curr = stack.pop()
        try:
            with os.scandir(curr) as it:
                for entry in it:
                    try:
                        if entry.is_dir(follow_symlinks=False):
                            stack.append(entry.path)
                        elif entry.is_file(follow_symlinks=False) and is_image_filename(entry.name):
                            index.setdefault(entry.name.lower(), []).append(entry.path)
                    except OSError:
                        continue
        except OSError:
            continue
    return index


def remap_config_source(sections: list, new_source_folder: str, old_source_folder: str = "") -> dict:
    """Resolve section plots against a new base source folder using rel_path or filename."""
    if not new_source_folder or not str(new_source_folder).strip():
        raise ValueError("New source folder cannot be empty.")

    norm_source = os.path.normpath(new_source_folder.strip())
    if not os.path.isdir(norm_source):
        raise NotADirectoryError(f"Source folder does not exist: {norm_source}")

    norm_old = (
        os.path.normpath(str(old_source_folder).strip())
        if old_source_folder and str(old_source_folder).strip()
        else ""
    )

    basename_index: dict[str, list[str]] | None = None
    all_paths = []
    matched_count = 0

    def _remap_list(raw_secs: list, depth: int = 0) -> list[dict]:
        nonlocal matched_count
        default_pb = depth == 0
        remapped_secs: list[dict] = []
        for idx, sec in enumerate(raw_secs or [], start=1):
            updated_plots = []
            for item in sec.get("plots") or []:
                if isinstance(item, dict):
                    old_path = str(item.get("path", ""))
                    raw_rel = item.get("rel_path")
                else:
                    old_path = str(item)
                    raw_rel = None

                rel_path = _sanitize_rel_path(raw_rel)
                if rel_path is None and norm_old and old_path:
                    rel_path = _safe_rel_path(old_path, norm_old)

                resolved = None
                new_rel = rel_path

                # 1. Try relative path directly under new_source_folder
                if rel_path:
                    candidate = os.path.normpath(os.path.join(norm_source, rel_path))
                    if _safe_rel_path(candidate, norm_source) is not None and os.path.isfile(candidate):
                        resolved = candidate

                # 2. Try direct basename in new_source_folder
                base_name = os.path.basename(old_path)
                if resolved is None and base_name:
                    direct_candidate = os.path.normpath(os.path.join(norm_source, base_name))
                    if os.path.isfile(direct_candidate):
                        resolved = direct_candidate
                        new_rel = base_name

                # 3. Fallback: walk new_source_folder once for unique basename (or unique parent+basename) match
                if resolved is None and base_name:
                    if basename_index is None:
                        basename_index = _build_basename_index(norm_source)
                    matches = sorted(basename_index.get(base_name.lower(), []))
                    if len(matches) == 1:
                        resolved = matches[0]
                        new_rel = _safe_rel_path(resolved, norm_source)
                    elif len(matches) > 1:
                        old_parent = os.path.basename(os.path.dirname(os.path.normpath(old_path))).lower()
                        if old_parent:
                            parent_matches = [
                                m for m in matches if os.path.basename(os.path.dirname(m)).lower() == old_parent
                            ]
                            if len(parent_matches) == 1:
                                resolved = parent_matches[0]
                                new_rel = _safe_rel_path(resolved, norm_source)

                if resolved is not None:
                    matched_count += 1
                    final_path = resolved
                else:
                    # Keep candidate path if sanitized rel_path is known, else old_path
                    final_path = os.path.normpath(os.path.join(norm_source, rel_path)) if rel_path else old_path

                updated_plots.append({"path": final_path, "rel_path": new_rel})
                all_paths.append(final_path)

            remapped_subs = (
                _remap_list(sec.get("subsections") or [], depth + 1)
                if depth < MAX_SUBSECTION_DEPTH
                else []
            )
            remapped_secs.append({
                "id": str(sec.get("id") or f"sec_{idx}"),
                "name": str(sec.get("name", f"Section {idx}")),
                "page_break": bool(sec["page_break"]) if "page_break" in sec else default_pb,
                "plots": updated_plots,
                "subsections": remapped_subs,
            })
        return remapped_secs

    updated_sections = _remap_list(sections or [], depth=0)

    file_status = validate_files(all_paths)
    missing_count = sum(1 for p in all_paths if not file_status.get(p, {}).get("exists", False))

    return {
        "source_folder": norm_source,
        "sections": updated_sections,
        "file_status": file_status,
        "matched_count": matched_count,
        "missing_count": missing_count,
    }
