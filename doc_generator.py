"""doc_generator.py — Core Word (.docx) & high-resolution PDF generation engine.

Builds Word reports from structured sections (with custom titles, page breaks,
column grids, and section-based figure numbering like Figure 1.1, Figure 1.2)
and optionally exports lossless high-resolution PDFs via Microsoft Word +
pdf_images.restore_pdf_images.
"""

import io
import os
import tempfile
import traceback
from pathlib import Path

try:
    import pythoncom
    import win32com.client
    HAS_WIN32COM = True
except ImportError:
    HAS_WIN32COM = False

from docx import Document
from docx.shared import Inches, Cm, Pt, RGBColor
from docx.enum.section import WD_ORIENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

from PIL import Image, ImageOps

from pdf_images import PLOT_TAG_PREFIX, restore_pdf_images


MAX_SUBSECTION_DEPTH = 2


def has_any_plots(sec: dict, depth: int = 0) -> bool:
    if sec.get("plots"):
        return True
    if depth >= MAX_SUBSECTION_DEPTH:
        return False
    return any(has_any_plots(sub, depth + 1) for sub in (sec.get("subsections") or []))


def count_plots_recursive(sections: list[dict], depth: int = 0) -> int:
    total = 0
    for sec in sections or []:
        total += len(sec.get("plots") or [])
        if depth < MAX_SUBSECTION_DEPTH:
            total += count_plots_recursive(sec.get("subsections") or [], depth + 1)
    return total


# --------------------------------------------------------------------------- #
# Page-layout helpers                                                         #
# --------------------------------------------------------------------------- #

def set_table_borderless(table):
    tblPr = table._tbl.tblPr
    tblBorders = OxmlElement('w:tblBorders')
    for border_name in ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']:
        border = OxmlElement(f'w:{border_name}')
        border.set(qn('w:val'), 'none')
        tblBorders.append(border)
    tblPr.append(tblBorders)


def lock_row_cant_split(row):
    trPr = row._tr.get_or_add_trPr()
    trPr.append(OxmlElement('w:cantSplit'))


def prepare_image(img_path):
    with Image.open(img_path) as image:
        image = ImageOps.exif_transpose(image)
        width, height = image.size
        has_alpha = "A" in image.getbands() or "transparency" in image.info
        stream = io.BytesIO()
        image.convert("RGBA" if has_alpha else "RGB").save(stream, format="PNG")
        stream.seek(0)
    return width, height, stream


def compute_display_size(pil_w, pil_h, max_col_w_cm, user_max_w_cm, user_max_h_cm, usable_h_cm):
    if pil_w <= 0 or pil_h <= 0:
        raise ValueError("Image has invalid dimensions.")
    if user_max_w_cm is not None and user_max_w_cm <= 0:
        raise ValueError("Max width must be positive.")
    if user_max_h_cm is not None and user_max_h_cm <= 0:
        raise ValueError("Max height must be positive.")
    aspect_ratio = pil_w / float(pil_h)
    effective_max_w = min(max_col_w_cm, user_max_w_cm) if user_max_w_cm else max_col_w_cm
    target_w_cm = effective_max_w
    target_h_cm = target_w_cm / aspect_ratio
    allowed_max_h = max(2.5, usable_h_cm - 3.5)
    if user_max_h_cm:
        allowed_max_h = min(allowed_max_h, user_max_h_cm)
    if target_h_cm > allowed_max_h:
        target_h_cm = max(1.0, allowed_max_h)
        target_w_cm = target_h_cm * aspect_ratio
    return round(target_w_cm, 2), round(target_h_cm, 2)


# --------------------------------------------------------------------------- #
# Word -> PDF conversion (Windows COM)                                        #
# --------------------------------------------------------------------------- #

def convert_docx_to_pdf_win32(docx_path, pdf_path):
    if not HAS_WIN32COM:
        raise RuntimeError("PDF export requires pywin32 and Microsoft Word.")
    pythoncom.CoInitialize()
    word = None
    try:
        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False
        with tempfile.TemporaryDirectory(prefix=".plot2docx-", dir=os.path.dirname(os.path.abspath(pdf_path))) as temp_dir:
            word_pdf = os.path.join(temp_dir, "word.pdf")
            restored_pdf = os.path.join(temp_dir, "restored.pdf")
            doc = word.Documents.Open(os.path.abspath(docx_path), ReadOnly=True, AddToRecentFiles=False)
            try:
                doc.ExportAsFixedFormat(OutputFileName=word_pdf, ExportFormat=17, OptimizeFor=0, DocStructureTags=True)
            finally:
                doc.Close(False)
            restore_pdf_images(docx_path, word_pdf, restored_pdf)
            os.replace(restored_pdf, os.path.abspath(pdf_path))
    finally:
        try:
            if word:
                word.Quit()
        finally:
            pythoncom.CoUninitialize()


# --------------------------------------------------------------------------- #
# Core document builder                                                       #
# --------------------------------------------------------------------------- #

def _create_document(orientation, num_cols):
    num_cols = int(num_cols)
    if num_cols < 1:
        raise ValueError("num_cols must be at least 1.")
    doc = Document()
    for sec in doc.sections:
        sec.top_margin = Inches(0.8)
        sec.bottom_margin = Inches(0.8)
        sec.left_margin = Inches(0.8)
        sec.right_margin = Inches(0.8)

    last_sec = doc.sections[-1]
    curr_w = last_sec.page_width
    curr_h = last_sec.page_height
    if orientation == "landscape":
        last_sec.orientation = WD_ORIENT.LANDSCAPE
        last_sec.page_width = max(curr_w, curr_h)
        last_sec.page_height = min(curr_w, curr_h)
    else:
        last_sec.orientation = WD_ORIENT.PORTRAIT
        last_sec.page_width = min(curr_w, curr_h)
        last_sec.page_height = max(curr_w, curr_h)

    page_w_emus = int(last_sec.page_width)
    page_h_emus = int(last_sec.page_height)
    left_emus = int(last_sec.left_margin)
    right_emus = int(last_sec.right_margin)
    top_emus = int(last_sec.top_margin)
    bottom_emus = int(last_sec.bottom_margin)

    usable_w_cm = ((page_w_emus - left_emus - right_emus) / 914400.0) * 2.54
    usable_h_cm = ((page_h_emus - top_emus - bottom_emus) / 914400.0) * 2.54

    col_gap_cm = 0.6 if num_cols > 1 else 0.0
    max_col_w_cm = (usable_w_cm - (col_gap_cm * (num_cols - 1))) / num_cols
    max_col_w_cm = max(3.0, max_col_w_cm)

    return doc, usable_w_cm, usable_h_cm, max_col_w_cm


def _insert_section_heading(doc, title, number_label, level=0):
    """Insert formatted section/subsection heading with Word outline level."""
    h = doc.add_paragraph()
    sizes = {0: 14, 1: 12, 2: 10.5}
    space_before = {0: 14, 1: 10, 2: 8}
    space_after = {0: 8, 1: 6, 2: 4}
    pt_size = sizes.get(level, 10.5)

    h.paragraph_format.space_before = Pt(space_before.get(level, 8))
    h.paragraph_format.space_after = Pt(space_after.get(level, 4))
    h.paragraph_format.keep_with_next = True

    pPr = h._p.get_or_add_pPr()
    outline = OxmlElement("w:outlineLvl")
    outline.set(qn("w:val"), str(min(2, max(0, int(level)))))
    pPr.append(outline)

    prefix = f"Section {number_label}: " if level == 0 else f"{number_label} "
    run_num = h.add_run(prefix)
    run_num.bold = True
    run_num.font.size = Pt(pt_size)
    run_num.font.name = "Segoe UI"
    run_num.font.color.rgb = RGBColor(31, 106, 165)

    run_title = h.add_run(title)
    run_title.bold = True
    run_title.font.size = Pt(pt_size)
    run_title.font.name = "Segoe UI"
    run_title.font.color.rgb = RGBColor(35, 35, 35)


def _insert_single_column_images(
    doc,
    plot_paths,
    sec_idx,
    usable_h_cm,
    max_col_w_cm,
    user_max_w_cm,
    user_max_h_cm,
    total_overall,
    global_counter,
    export_pdf,
    progress_callback,
    sec_fig_counter=None,
):
    if sec_fig_counter is None:
        sec_fig_counter = [0]
    inserted = 0
    skipped = []
    has_section = sec_idx is not None

    for img_p in plot_paths:
        global_counter[0] += 1
        pct = (global_counter[0] / max(1, total_overall)) * (0.8 if export_pdf else 1.0)
        if progress_callback:
            progress_callback(f"Processing plot {global_counter[0]} of {total_overall}...", pct)

        try:
            orig_w, orig_h, img_stream = prepare_image(img_p)
        except Exception as img_err:
            skipped.append((img_p, f"Cannot read image: {img_err}"))
            continue

        disp_w, disp_h = compute_display_size(
            orig_w, orig_h, max_col_w_cm, user_max_w_cm, user_max_h_cm, usable_h_cm
        )
        try:
            p_img = doc.add_paragraph()
            p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p_img.paragraph_format.space_before = Pt(8)
            p_img.paragraph_format.space_after = Pt(2)
            p_img.paragraph_format.keep_with_next = True
            run_img = p_img.add_run()
            shape = run_img.add_picture(img_stream, width=Cm(disp_w), height=Cm(disp_h))

            # Use global figure tag for win32 PDF high-res restoration
            shape._inline.docPr.set("descr", f"{PLOT_TAG_PREFIX}{global_counter[0]}")

            sec_fig_counter[0] += 1
            inserted += 1
            fig_label = f"Figure {sec_idx}.{sec_fig_counter[0]}" if has_section else f"Figure {global_counter[0]}"
            p_cap = doc.add_paragraph()
            p_cap.alignment = WD_ALIGN_PARAGRAPH.LEFT
            p_cap.paragraph_format.space_before = Pt(2)
            p_cap.paragraph_format.space_after = Pt(14)
            r_bold = p_cap.add_run(f"{fig_label}\t")
            r_bold.bold = True
            r_bold.font.size = Pt(9.5)
            r_bold.font.name = "Segoe UI"
            r_bold.font.color.rgb = RGBColor(40, 40, 40)

            r_path = p_cap.add_run(img_p)
            r_path.font.size = Pt(8.5)
            r_path.font.name = "Segoe UI"
            r_path.font.color.rgb = RGBColor(90, 90, 90)
        except Exception as add_err:
            skipped.append((img_p, f"Word insertion error: {add_err}"))
            continue

    return inserted, skipped


def _insert_multi_column_images(
    doc,
    plot_paths,
    sec_idx,
    max_col_w_cm,
    num_cols,
    user_max_w_cm,
    user_max_h_cm,
    usable_h_cm,
    total_overall,
    global_counter,
    export_pdf,
    progress_callback,
    sec_fig_counter=None,
):
    if sec_fig_counter is None:
        sec_fig_counter = [0]
    table = None
    inserted = 0
    skipped = []
    has_section = sec_idx is not None
    row_buf = []

    def _flush_row(items):
        nonlocal table, inserted
        if not items:
            return
        if table is None:
            table = doc.add_table(rows=0, cols=num_cols)
            table.alignment = WD_TABLE_ALIGNMENT.CENTER
            set_table_borderless(table)
            for col in table.columns:
                col.width = Cm(max_col_w_cm)

        img_row = table.add_row()
        cap_row = table.add_row()
        lock_row_cant_split(img_row)
        lock_row_cant_split(cap_row)

        for cell in img_row.cells:
            cell.width = Cm(max_col_w_cm)
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        for cell in cap_row.cells:
            cell.width = Cm(max_col_w_cm)
            cell.vertical_alignment = WD_ALIGN_VERTICAL.TOP

        for col_idx, (img_p, img_stream, disp_w, disp_h, tag_num) in enumerate(items):
            try:
                cell_img = img_row.cells[col_idx]
                p_img = cell_img.paragraphs[0]
                p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p_img.paragraph_format.space_before = Pt(4)
                p_img.paragraph_format.space_after = Pt(2)
                p_img.paragraph_format.keep_with_next = True
                run_img = p_img.add_run()
                shape = run_img.add_picture(img_stream, width=Cm(disp_w), height=Cm(disp_h))
                shape._inline.docPr.set("descr", f"{PLOT_TAG_PREFIX}{tag_num}")

                sec_fig_counter[0] += 1
                inserted += 1
                fig_label = f"Figure {sec_idx}.{sec_fig_counter[0]}" if has_section else f"Figure {tag_num}"
                cell_cap = cap_row.cells[col_idx]
                p_cap = cell_cap.paragraphs[0]
                p_cap.alignment = WD_ALIGN_PARAGRAPH.LEFT
                p_cap.paragraph_format.space_before = Pt(2)
                p_cap.paragraph_format.space_after = Pt(12)
                r_bold = p_cap.add_run(f"{fig_label}\t")
                r_bold.bold = True
                r_bold.font.size = Pt(9.0)
                r_bold.font.name = "Segoe UI"
                r_bold.font.color.rgb = RGBColor(40, 40, 40)

                r_path = p_cap.add_run(img_p)
                r_path.font.size = Pt(8.0)
                r_path.font.name = "Segoe UI"
                r_path.font.color.rgb = RGBColor(90, 90, 90)
            except Exception as add_err:
                skipped.append((img_p, f"Word insertion error: {add_err}"))

    for img_p in plot_paths:
        global_counter[0] += 1
        pct = (global_counter[0] / max(1, total_overall)) * (0.8 if export_pdf else 1.0)
        if progress_callback:
            progress_callback(f"Processing plot {global_counter[0]} of {total_overall}...", pct)

        try:
            orig_w, orig_h, img_stream = prepare_image(img_p)
            disp_w, disp_h = compute_display_size(
                orig_w, orig_h, max_col_w_cm, user_max_w_cm, user_max_h_cm, usable_h_cm
            )
        except Exception as img_err:
            skipped.append((img_p, f"Cannot read image: {img_err}"))
            continue

        row_buf.append((img_p, img_stream, disp_w, disp_h, global_counter[0]))
        if len(row_buf) == num_cols:
            _flush_row(row_buf)
            row_buf = []

    if row_buf:
        _flush_row(row_buf)

    return inserted, skipped


def build_report_documents(
    sections_or_paths,
    save_file,
    orientation="landscape",
    num_cols=2,
    user_max_w_cm=None,
    user_max_h_cm=None,
    export_pdf=False,
    progress_callback=None,
    export_docx=True,
    export_html=False,
    html_embed_images=False,
):
    """Build Word (.docx), standalone HTML (.html), and optional high-res PDF."""
    result = {
        "success": False,
        "docx_path": None,
        "html_path": None,
        "pdf_path": None,
        "count": 0,
        "skipped": [],
    }

    if not sections_or_paths:
        result["skipped"].append(("<none>", "No plots provided to generate."))
        return result

    if isinstance(sections_or_paths[0], dict):
        sections = sections_or_paths
    else:
        sections = [{"title": "", "page_break": False, "plots": list(sections_or_paths), "subsections": []}]

    total_plots = count_plots_recursive(sections)
    if total_plots == 0:
        result["skipped"].append(("<none>", "No plots found in sections."))
        return result

    try:
        if not export_docx and not export_html:
            raise ValueError("Select at least one output format (.docx or .html).")
        if user_max_w_cm is not None and float(user_max_w_cm) <= 0:
            raise ValueError("Max width must be positive.")
        if user_max_h_cm is not None and float(user_max_h_cm) <= 0:
            raise ValueError("Max height must be positive.")

        num_cols = int(num_cols)
        if num_cols < 1:
            raise ValueError("num_cols must be at least 1.")

        effective_pdf = bool(export_pdf) and bool(export_docx)
        base_path = Path(save_file)
        docx_out = str(base_path.with_suffix(".docx")) if export_docx else None
        html_out = str(base_path.with_suffix(".html")) if export_html else None

        inserted_total = 0
        all_skipped = []

        if export_docx and docx_out is not None:
            doc, _, usable_h_cm, max_col_w_cm = _create_document(orientation, num_cols)
            global_counter = [0]
            is_first_node = [True]
            rendered_sec_idx = 0

            def _render_word_node(sec, depth, number_label, sec_num_for_figs, sec_fig_counter):
                nonlocal inserted_total
                if not has_any_plots(sec, depth):
                    return
                sec_title = (sec.get("title") or sec.get("name") or "").strip()
                page_break = bool(sec["page_break"]) if "page_break" in sec else (depth == 0)

                if not is_first_node[0] and page_break:
                    doc.add_page_break()
                is_first_node[0] = False

                if sec_title and number_label:
                    _insert_section_heading(doc, sec_title, number_label, level=depth)

                sec_plots = sec.get("plots") or []
                if sec_plots:
                    if num_cols == 1:
                        inserted, skipped = _insert_single_column_images(
                            doc, sec_plots, sec_num_for_figs, usable_h_cm, max_col_w_cm,
                            user_max_w_cm, user_max_h_cm, total_plots, global_counter,
                            effective_pdf, progress_callback, sec_fig_counter=sec_fig_counter,
                        )
                    else:
                        inserted, skipped = _insert_multi_column_images(
                            doc, sec_plots, sec_num_for_figs, max_col_w_cm, num_cols,
                            user_max_w_cm, user_max_h_cm, usable_h_cm, total_plots,
                            global_counter, effective_pdf, progress_callback, sec_fig_counter=sec_fig_counter,
                        )
                    inserted_total += inserted
                    all_skipped.extend(skipped)

                if depth < MAX_SUBSECTION_DEPTH:
                    sub_idx = 0
                    for sub in sec.get("subsections") or []:
                        if not has_any_plots(sub, depth + 1):
                            continue
                        sub_idx += 1
                        sub_label = f"{number_label}.{sub_idx}" if number_label else str(sub_idx)
                        _render_word_node(sub, depth + 1, sub_label, sec_num_for_figs, sec_fig_counter)

            for sec in sections:
                if not has_any_plots(sec, 0):
                    continue
                sec_title = (sec.get("title") or sec.get("name") or "").strip()
                if sec_title:
                    rendered_sec_idx += 1
                    sec_num = rendered_sec_idx
                    num_label = str(rendered_sec_idx)
                else:
                    sec_num = None
                    num_label = None
                _render_word_node(sec, 0, num_label, sec_num, [0])

            if inserted_total == 0:
                result["skipped"] = all_skipped or [("<none>", "No valid plot images could be inserted.")]
                return result

            if progress_callback:
                progress_callback("Saving Word document...", 0.80 if effective_pdf else 0.90)
            os.makedirs(os.path.dirname(os.path.abspath(docx_out)), exist_ok=True)
            doc.save(docx_out)
            result["docx_path"] = docx_out

        if export_html and html_out is not None:
            import html_generator

            if progress_callback:
                progress_callback("Generating HTML report...", 0.84 if effective_pdf else 0.95)
            html_res = html_generator.build_html_report(
                sections,
                html_out,
                num_cols=num_cols,
                user_max_w_cm=user_max_w_cm,
                user_max_h_cm=user_max_h_cm,
                embed_images=bool(html_embed_images),
                progress_callback=None if export_docx else progress_callback,
            )
            if html_res["count"] == 0 and not export_docx:
                result["skipped"] = html_res["skipped"]
                return result
            result["html_path"] = html_out
            if not export_docx:
                inserted_total = html_res["count"]
                all_skipped.extend(html_res["skipped"])

        pdf_path = None
        if effective_pdf and docx_out is not None:
            pdf_path = str(base_path.with_suffix(".pdf"))
            if progress_callback:
                progress_callback("Exporting high-resolution PDF via Microsoft Word...", 0.88)
            try:
                convert_docx_to_pdf_win32(docx_out, pdf_path)
            except Exception as pdf_err:
                raise RuntimeError(
                    f"The Word document was saved to:\n{docx_out}\n\nHigh-resolution PDF export failed: {pdf_err}"
                ) from pdf_err

        if progress_callback:
            progress_callback("Done.", 1.0)

        result["success"] = True
        result["pdf_path"] = pdf_path
        result["count"] = inserted_total
        result["skipped"] = all_skipped

    except PermissionError:
        result["docx_path"] = save_file
        result["skipped"].append((
            "<none>",
            f"Permission Denied when saving to:\n{save_file}\n\n"
            f"Please make sure the destination document is NOT open in Microsoft Word, or choose another filename.",
        ))
    except (ValueError, RuntimeError) as err:
        result["skipped"].append(("<none>", str(err)))
    except Exception:
        result["skipped"].append(("<none>", f"Error details:\n{traceback.format_exc()}"))

    return result
