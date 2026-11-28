"""html_generator.py: Standalone scrollable HTML report generator.

Generates a single offline HTML5 report with collapsible Table of Contents,
Light/Dark theme toggle, click-to-zoom lightbox, and lazy-loaded network or
local image links (with an optional base64 embed mode).
"""

import base64
import html
import io
import os
from pathlib import Path
from urllib.parse import quote

from PIL import Image, ImageOps

from doc_generator import MAX_SUBSECTION_DEPTH, count_plots_recursive, has_any_plots


def image_src_for_html(img_path: str, html_dir: str, embed_images: bool = False) -> str:
    """Return an HTML img src attribute value (relative path, file:// URI, or data URI)."""
    if embed_images:
        with Image.open(img_path) as image:
            image = ImageOps.exif_transpose(image)
            has_alpha = "A" in image.getbands() or "transparency" in image.info
            fmt = "PNG" if (has_alpha or Path(img_path).suffix.lower() == ".png") else "JPEG"
            mime = "image/png" if fmt == "PNG" else "image/jpeg"
            converted = image.convert("RGBA" if fmt == "PNG" and has_alpha else "RGB")
            buf = io.BytesIO()
            if fmt == "PNG":
                converted.save(buf, format="PNG", optimize=True)
            else:
                converted.save(buf, format="JPEG", quality=92)
            encoded = base64.b64encode(buf.getvalue()).decode("ascii")
            return f"data:{mime};base64,{encoded}"

    abs_img = os.path.abspath(img_path)
    abs_dir = os.path.abspath(html_dir)
    try:
        rel = os.path.relpath(abs_img, abs_dir)
        if rel != "." and not rel.startswith("..") and not os.path.isabs(rel):
            return quote(rel.replace("\\", "/"), safe="/")
    except ValueError:
        pass
    return Path(abs_img).as_uri()


def _render_html_report_shell(
    title: str,
    num_cols: int,
    user_max_w_cm: float | None,
    user_max_h_cm: float | None,
    toc_items: list[tuple[int, str, str]],
    body_html: str,
    total_count: int,
) -> str:
    max_w_css = f"{float(user_max_w_cm):.2f}cm" if user_max_w_cm else "100%"
    max_h_css = f"{float(user_max_h_cm):.2f}cm" if user_max_h_cm else "75vh"
    toc_links = "\n".join(
        f'<a class="toc-link toc-d{depth}" href="#{html.escape(anchor)}">{html.escape(label)}</a>'
        for depth, anchor, label in toc_items
    )

    return f"""<!DOCTYPE html>
<html lang="en" data-theme="light">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>{html.escape(title)}</title>
<style>
  :root, html[data-theme="light"] {{
    --bg: #f5f4f0;
    --surface: #ffffff;
    --surface-2: #eeece6;
    --border: #d0ccc2;
    --text: #161a20;
    --muted: #4d5563;
    --accent: #1b5e91;
    --accent-soft: rgba(27, 94, 145, 0.11);
    --warn-bg: #fef3c7;
    --warn-text: #8c3b06;
  }}
  html[data-theme="dark"] {{
    --bg: #13161b;
    --surface: #1a1e25;
    --surface-2: #222730;
    --border: #373f4d;
    --text: #eef1f6;
    --muted: #9da7b8;
    --accent: #7cc4fa;
    --accent-soft: rgba(96, 173, 232, 0.14);
    --warn-bg: rgba(245, 158, 11, 0.14);
    --warn-text: #fbbf24;
  }}
  * {{ box-sizing: border-box; }}
  body {{
    margin: 0;
    font-family: "Segoe UI Variable Text", "Segoe UI", system-ui, sans-serif;
    background: var(--bg);
    color: var(--text);
    font-size: 14px;
    line-height: 1.45;
  }}
  .layout {{ display: flex; min-height: 100vh; }}
  .toc {{
    width: 270px;
    flex-shrink: 0;
    position: sticky;
    top: 0;
    height: 100vh;
    overflow-y: auto;
    background: var(--surface);
    border-right: 1px solid var(--border);
    padding: 16px 12px;
  }}
  .toc.collapsed {{ display: none; }}
  .toc-head {{
    font-size: 12px;
    font-weight: 700;
    text-transform: uppercase;
    letter-spacing: 0.04em;
    color: var(--muted);
    margin: 0 0 10px 6px;
  }}
  .toc-link {{
    display: block;
    padding: 5px 8px;
    border-radius: 5px;
    color: var(--text);
    text-decoration: none;
    font-size: 13px;
    overflow: hidden;
    text-overflow: ellipsis;
    white-space: nowrap;
  }}
  .toc-link:hover {{ background: var(--surface-2); color: var(--accent); }}
  .toc-d1 {{ padding-left: 20px; font-size: 12.5px; color: var(--muted); }}
  .toc-d2 {{ padding-left: 34px; font-size: 12px; color: var(--muted); }}
  .main {{ flex: 1; min-width: 0; padding: 18px 24px 48px 24px; }}
  .topbar {{
    display: flex;
    align-items: center;
    justify-content: space-between;
    gap: 12px;
    flex-wrap: wrap;
    padding-bottom: 14px;
    margin-bottom: 18px;
    border-bottom: 1px solid var(--border);
  }}
  .topbar h1 {{ margin: 0; font-size: 20px; font-weight: 700; }}
  .topbar-meta {{ color: var(--muted); font-size: 12.5px; font-family: "Cascadia Code", Consolas, monospace; }}
  .topbar-actions {{ display: flex; gap: 8px; flex-wrap: wrap; }}
  .btn {{
    padding: 6px 12px;
    border-radius: 6px;
    border: 1px solid var(--border);
    background: var(--surface);
    color: var(--text);
    font-family: inherit;
    font-size: 12.5px;
    font-weight: 600;
    cursor: pointer;
  }}
  .btn:hover {{ background: var(--surface-2); }}
  .sec-heading {{ margin: 26px 0 12px 0; font-weight: 700; scroll-margin-top: 16px; }}
  .sec-heading.d0 {{ font-size: 18px; border-bottom: 2px solid var(--border); padding-bottom: 6px; }}
  .sec-heading.d1 {{ font-size: 15.5px; margin-top: 20px; }}
  .sec-heading.d2 {{ font-size: 14px; margin-top: 16px; }}
  .sec-num {{ color: var(--accent); }}
  .fig-grid {{
    display: grid;
    grid-template-columns: repeat(var(--cols, {num_cols}), minmax(0, 1fr));
    gap: 16px;
    margin-bottom: 16px;
  }}
  .fig-card {{
    background: var(--surface);
    border: 1px solid var(--border);
    border-radius: 8px;
    padding: 10px;
    display: flex;
    flex-direction: column;
    gap: 8px;
  }}
  .fig-zoom {{
    border: none;
    background: var(--surface-2);
    border-radius: 6px;
    padding: 6px;
    cursor: zoom-in;
    display: flex;
    align-items: center;
    justify-content: center;
    min-height: 80px;
  }}
  .fig-img {{
    display: block;
    max-width: {max_w_css};
    max-height: {max_h_css};
    width: auto;
    height: auto;
    object-fit: contain;
    image-orientation: from-image;
  }}
  .fig-caption {{ font-size: 12.5px; display: flex; flex-direction: column; gap: 2px; }}
  .fig-label {{ font-weight: 700; font-family: "Cascadia Code", Consolas, monospace; color: var(--accent); }}
  .fig-path {{
    font-size: 11px;
    color: var(--muted);
    font-family: "Cascadia Code", Consolas, monospace;
    word-break: break-all;
  }}
  .fig-fallback {{
    padding: 10px;
    border-radius: 6px;
    background: var(--warn-bg);
    color: var(--warn-text);
    font-size: 12px;
    display: none;
    flex-direction: column;
    gap: 6px;
  }}
  .fig-card.is-broken .fig-zoom {{ display: none; }}
  .fig-card.is-broken .fig-fallback {{ display: flex; }}
  .lightbox {{
    position: fixed;
    inset: 0;
    background: rgba(0, 0, 0, 0.78);
    display: none;
    align-items: center;
    justify-content: center;
    padding: 20px;
    z-index: 100;
  }}
  .lightbox.open {{ display: flex; }}
  .lightbox-inner {{
    background: var(--surface);
    border: 1px solid var(--border);
    border-radius: 8px;
    max-width: 94vw;
    max-height: 92vh;
    padding: 14px;
    display: flex;
    flex-direction: column;
    gap: 10px;
  }}
  .lightbox-inner img {{ max-width: 90vw; max-height: 76vh; object-fit: contain; image-orientation: from-image; }}
  .lightbox-head {{ display: flex; justify-content: space-between; align-items: center; gap: 12px; }}
  @media (max-width: 820px) {{
    .toc {{ display: none; }}
    .fig-grid {{ grid-template-columns: 1fr; }}
    .main {{ padding: 14px; }}
  }}
  @media print {{
    .toc, .topbar-actions, .lightbox {{ display: none !important; }}
    .main {{ padding: 0; }}
    .fig-card {{ break-inside: avoid; }}
  }}
</style>
</head>
<body>
<div class="layout">
  <aside class="toc" id="toc" aria-label="Table of contents">
    <div class="toc-head">Contents</div>
    {toc_links or '<div class="topbar-meta">All Plots</div>'}
  </aside>
  <main class="main">
    <header class="topbar">
      <div>
        <h1>{html.escape(title)}</h1>
        <div class="topbar-meta">{total_count} figure(s)</div>
      </div>
      <div class="topbar-actions">
        <button type="button" class="btn" id="btn-toc">Toggle Contents</button>
        <button type="button" class="btn" id="btn-theme">Toggle Theme</button>
      </div>
    </header>
    {body_html}
  </main>
</div>
<div class="lightbox" id="lightbox" role="dialog" aria-modal="true" aria-label="Image preview">
  <div class="lightbox-inner">
    <div class="lightbox-head">
      <strong id="lb-title">Preview</strong>
      <button type="button" class="btn" id="lb-close">Close</button>
    </div>
    <img id="lb-img" alt="">
    <div class="fig-path" id="lb-path"></div>
  </div>
</div>
<script>
(() => {{
  const lb = document.getElementById("lightbox");
  const lbImg = document.getElementById("lb-img");
  const lbTitle = document.getElementById("lb-title");
  const lbPath = document.getElementById("lb-path");
  let lastFocus = null;

  window.plotBroken = (img) => {{
    const card = img.closest(".fig-card");
    if (card) card.classList.add("is-broken");
  }};
  document.addEventListener("click", (e) => {{
    const copyBtn = e.target.closest("[data-copy]");
    if (copyBtn && navigator.clipboard) {{
      navigator.clipboard.writeText(copyBtn.dataset.copy).catch(() => {{}});
      copyBtn.textContent = "Copied";
      return;
    }}
    const zoomBtn = e.target.closest(".fig-zoom");
    if (zoomBtn) {{
      lastFocus = zoomBtn;
      lbImg.src = zoomBtn.dataset.src;
      lbImg.alt = zoomBtn.dataset.label || "Plot preview";
      lbTitle.textContent = zoomBtn.dataset.label || "Preview";
      lbPath.textContent = zoomBtn.dataset.raw || "";
      lb.classList.add("open");
      document.getElementById("lb-close").focus();
    }}
  }});
  const closeLb = () => {{
    lb.classList.remove("open");
    if (lastFocus) lastFocus.focus();
  }};
  document.getElementById("lb-close").addEventListener("click", closeLb);
  lb.addEventListener("click", (e) => {{ if (e.target === lb) closeLb(); }});
  document.addEventListener("keydown", (e) => {{ if (e.key === "Escape" && lb.classList.contains("open")) closeLb(); }});
  document.getElementById("btn-toc").addEventListener("click", () => {{
    document.getElementById("toc").classList.toggle("collapsed");
  }});
  document.getElementById("btn-theme").addEventListener("click", () => {{
    const root = document.documentElement;
    root.setAttribute("data-theme", root.getAttribute("data-theme") === "dark" ? "light" : "dark");
  }});
}})();
</script>
</body>
</html>
"""


def build_html_report(
    sections: list[dict],
    html_path: str,
    num_cols: int = 2,
    user_max_w_cm: float | None = None,
    user_max_h_cm: float | None = None,
    embed_images: bool = False,
    progress_callback=None,
) -> dict:
    """Build a standalone HTML report following Level-1 figure numbering."""
    num_cols = max(1, int(num_cols))
    html_abs = os.path.abspath(html_path)
    html_dir = os.path.dirname(html_abs)
    total_plots = count_plots_recursive(sections)
    processed = 0
    inserted = 0
    skipped: list[tuple[str, str]] = []
    toc_items: list[tuple[int, str, str]] = []
    body_parts: list[str] = []
    rendered_sec_idx = 0
    anchor_counter = 0

    def render_plots(plot_paths: list[str], sec_num_for_figs: int | None, sec_fig_counter: list[int]) -> None:
        nonlocal processed, inserted
        if not plot_paths:
            return
        cards: list[str] = []
        for img_p in plot_paths:
            processed += 1
            if progress_callback:
                progress_callback(f"Preparing HTML plot {processed} of {max(1, total_plots)}...", processed / max(1, total_plots))

            if not os.path.isfile(img_p):
                skipped.append((img_p, "File not found on disk."))
                continue
            try:
                src = image_src_for_html(img_p, html_dir, embed_images=embed_images)
            except Exception as err:
                skipped.append((img_p, f"Cannot read image: {err}"))
                continue

            sec_fig_counter[0] += 1
            inserted += 1
            fig_label = f"Figure {sec_num_for_figs}.{sec_fig_counter[0]}" if sec_num_for_figs is not None else f"Figure {sec_fig_counter[0]}"
            esc_src = html.escape(src, quote=True)
            esc_label = html.escape(fig_label, quote=True)
            esc_raw = html.escape(img_p, quote=True)
            esc_name = html.escape(os.path.basename(img_p), quote=True)

            cards.append(
                f'<article class="fig-card">'
                f'<button type="button" class="fig-zoom" data-src="{esc_src}" data-label="{esc_label}" data-raw="{esc_raw}">'
                f'<img class="fig-img" src="{esc_src}" alt="{esc_name}" loading="lazy" decoding="async" onerror="window.plotBroken(this)">'
                f'</button>'
                f'<div class="fig-fallback">'
                f'<strong>Image unavailable from browser</strong>'
                f'<span class="fig-path">{esc_raw}</span>'
                f'<span>Check network access or open this file in Microsoft Edge or Chrome if using Firefox.</span>'
                f'<div><button type="button" class="btn" data-copy="{esc_raw}">Copy Path</button></div>'
                f'</div>'
                f'<div class="fig-caption">'
                f'<span class="fig-label">{esc_label}</span>'
                f'<span class="fig-path">{esc_raw}</span>'
                f'</div>'
                f'</article>'
            )
        if cards:
            body_parts.append(f'<div class="fig-grid">{"".join(cards)}</div>')

    def render_node(sec: dict, depth: int, number_label: str | None, sec_num_for_figs: int | None, sec_fig_counter: list[int]) -> None:
        nonlocal anchor_counter
        if not has_any_plots(sec, depth):
            return
        sec_title = (sec.get("title") or sec.get("name") or "").strip()
        if sec_title and number_label:
            anchor_counter += 1
            anchor = f"sec-anchor-{anchor_counter}"
            prefix = f"Section {number_label}: " if depth == 0 else f"{number_label} "
            toc_items.append((depth, anchor, f"{prefix}{sec_title}"))
            body_parts.append(
                f'<h{min(depth + 2, 4)} id="{anchor}" class="sec-heading d{depth}">'
                f'<span class="sec-num">{html.escape(prefix)}</span>{html.escape(sec_title)}'
                f'</h{min(depth + 2, 4)}>'
            )

        render_plots(sec.get("plots") or [], sec_num_for_figs, sec_fig_counter)

        if depth < MAX_SUBSECTION_DEPTH:
            sub_idx = 0
            for sub in sec.get("subsections") or []:
                if not has_any_plots(sub, depth + 1):
                    continue
                sub_idx += 1
                sub_label = f"{number_label}.{sub_idx}" if number_label else str(sub_idx)
                render_node(sub, depth + 1, sub_label, sec_num_for_figs, sec_fig_counter)

    for sec in sections:
        if not has_any_plots(sec, 0):
            continue
        sec_title = (sec.get("title") or sec.get("name") or "").strip()
        if sec_title:
            rendered_sec_idx += 1
            sec_num = rendered_sec_idx
            label = str(rendered_sec_idx)
        else:
            sec_num = None
            label = None
        render_node(sec, 0, label, sec_num, [0])

    if inserted == 0:
        return {"count": 0, "skipped": skipped or [("<none>", "No valid plot images could be inserted.")]}

    os.makedirs(html_dir, exist_ok=True)
    report_title = Path(html_abs).stem.replace("_", " ") or "Plot to Docx Report"
    doc_html = _render_html_report_shell(
        title=report_title,
        num_cols=num_cols,
        user_max_w_cm=user_max_w_cm,
        user_max_h_cm=user_max_h_cm,
        toc_items=toc_items,
        body_html="\n".join(body_parts),
        total_count=inserted,
    )
    Path(html_abs).write_text(doc_html, encoding="utf-8")
    return {"count": inserted, "skipped": skipped}
