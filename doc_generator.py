"""doc_generator.py — Module 3: document & PDF generation.

Receives page-layout options and structured sections (with custom titles and
page breaks) from plot_organizer (or raw plots), builds the Word (.docx) report
with section-based figure numbering (Figure 1.1, Figure 1.2, etc.), and
optionally converts to high-fidelity .pdf via Microsoft Word.
"""

import datetime
import os
import tempfile
import traceback
import threading
import io
from pathlib import Path

import customtkinter as ctk
import tkinter as tk
from tkinter import filedialog, messagebox

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

from PIL import Image

from pdf_images import PLOT_TAG_PREFIX, restore_pdf_images
from plot_selector import get_selected_paths


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
        width, height = image.size
        has_alpha = "A" in image.getbands() or "transparency" in image.info
        stream = io.BytesIO()
        image.convert("RGBA" if has_alpha else "RGB").save(stream, format="PNG")
        stream.seek(0)
    return width, height, stream


def compute_display_size(pil_w, pil_h, max_col_w_cm, user_max_w_cm, user_max_h_cm, usable_h_cm):
    aspect_ratio = pil_w / float(pil_h)
    effective_max_w = min(max_col_w_cm, user_max_w_cm) if user_max_w_cm else max_col_w_cm
    target_w_cm = effective_max_w
    target_h_cm = target_w_cm / aspect_ratio
    allowed_max_h = usable_h_cm - 3.5
    if user_max_h_cm:
        allowed_max_h = min(allowed_max_h, user_max_h_cm)
    if target_h_cm > allowed_max_h:
        target_h_cm = max(2.5, allowed_max_h)
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


def _insert_section_heading(doc, title, sec_idx):
    """Insert formatted section heading styled like an engineering report."""
    h = doc.add_paragraph()
    h.paragraph_format.space_before = Pt(14)
    h.paragraph_format.space_after = Pt(8)
    h.paragraph_format.keep_with_next = True
    
    run_num = h.add_run(f"Section {sec_idx}: ")
    run_num.bold = True
    run_num.font.size = Pt(14)
    run_num.font.name = "Segoe UI"
    run_num.font.color.rgb = RGBColor(31, 106, 165)

    run_title = h.add_run(title)
    run_title.bold = True
    run_title.font.size = Pt(14)
    run_title.font.name = "Segoe UI"
    run_title.font.color.rgb = RGBColor(35, 35, 35)


def _insert_single_column_images(doc, plot_paths, sec_idx, usable_h_cm, max_col_w_cm, user_max_w_cm, user_max_h_cm, total_overall, global_counter, export_pdf, progress_callback):
    fig_idx = 1
    skipped = []
    has_section = (sec_idx is not None)

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

        disp_w, disp_h = compute_display_size(orig_w, orig_h, max_col_w_cm, user_max_w_cm, user_max_h_cm, usable_h_cm)
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

            fig_label = f"Figure {sec_idx}.{fig_idx}" if has_section else f"Figure {global_counter[0]}"
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

            fig_idx += 1
        except Exception as add_err:
            skipped.append((img_p, f"Word insertion error: {add_err}"))
            continue

    return fig_idx - 1, skipped


def _insert_multi_column_images(doc, plot_paths, sec_idx, max_col_w_cm, num_cols, usable_h_cm, total_overall, global_counter, export_pdf, progress_callback):
    table = doc.add_table(rows=0, cols=num_cols)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borderless(table)
    for col in table.columns:
        col.width = Cm(max_col_w_cm)

    fig_idx = 1
    skipped = []
    has_section = (sec_idx is not None)

    for i in range(0, len(plot_paths), num_cols):
        chunk = plot_paths[i:i + num_cols]
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

        for col_idx, img_p in enumerate(chunk):
            global_counter[0] += 1
            pct = (global_counter[0] / max(1, total_overall)) * (0.8 if export_pdf else 1.0)
            if progress_callback:
                progress_callback(f"Processing plot {global_counter[0]} of {total_overall}...", pct)

            try:
                orig_w, orig_h, img_stream = prepare_image(img_p)
            except Exception as img_err:
                skipped.append((img_p, f"Cannot read image: {img_err}"))
                continue

            disp_w, disp_h = compute_display_size(orig_w, orig_h, max_col_w_cm, None, None, usable_h_cm)
            try:
                cell_img = img_row.cells[col_idx]
                p_img = cell_img.paragraphs[0]
                p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p_img.paragraph_format.space_before = Pt(4)
                p_img.paragraph_format.space_after = Pt(2)
                p_img.paragraph_format.keep_with_next = True
                run_img = p_img.add_run()
                shape = run_img.add_picture(img_stream, width=Cm(disp_w), height=Cm(disp_h))
                shape._inline.docPr.set("descr", f"{PLOT_TAG_PREFIX}{global_counter[0]}")

                fig_label = f"Figure {sec_idx}.{fig_idx}" if has_section else f"Figure {global_counter[0]}"
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

                fig_idx += 1
            except Exception as add_err:
                skipped.append((img_p, f"Word insertion error: {add_err}"))
                continue

    return fig_idx - 1, skipped


def build_report_documents(sections_or_paths, save_file, orientation, num_cols, user_max_w_cm, user_max_h_cm, export_pdf, progress_callback=None):
    """
    Accepts either:
      - List of section dicts: [{'title': str, 'page_break': bool, 'plots': [...]}, ...]
      - Or simple list of image paths: [str, ...]
    """
    result = {"success": False, "docx_path": None, "pdf_path": None, "count": 0, "skipped": []}

    # Normalize inputs to structured sections
    if not sections_or_paths:
        result["skipped"].append(("<none>", "No plots provided to generate."))
        return result

    if isinstance(sections_or_paths[0], dict):
        sections = sections_or_paths
    else:
        sections = [{"title": "", "page_break": False, "plots": list(sections_or_paths)}]

    total_plots = sum(len(s.get("plots", [])) for s in sections)
    if total_plots == 0:
        result["skipped"].append(("<none>", "No plots found in sections."))
        return result

    try:
        doc, usable_w_cm, usable_h_cm, max_col_w_cm = _create_document(orientation, num_cols)
        global_counter = [0]
        all_skipped = []
        is_first_section = True

        for s_idx, sec in enumerate(sections, start=1):
            sec_title = sec.get("title", "").strip()
            sec_plots = sec.get("plots", [])
            page_break = sec.get("page_break", True)

            if not sec_plots:
                continue

            # Page break before new section (except the very first page)
            if not is_first_section and page_break:
                doc.add_page_break()

            # Insert section heading if title exists
            sec_number = s_idx if sec_title else None
            if sec_title:
                _insert_section_heading(doc, sec_title, s_idx)

            # Insert section plots
            if num_cols == 1:
                _, skipped = _insert_single_column_images(
                    doc, sec_plots, sec_number, usable_h_cm, max_col_w_cm,
                    user_max_w_cm, user_max_h_cm, total_plots, global_counter,
                    export_pdf, progress_callback
                )
            else:
                _, skipped = _insert_multi_column_images(
                    doc, sec_plots, sec_number, max_col_w_cm, num_cols,
                    usable_h_cm, total_plots, global_counter,
                    export_pdf, progress_callback
                )

            all_skipped.extend(skipped)
            is_first_section = False

        doc.save(save_file)
        pdf_path = None
        if export_pdf:
            pdf_path = str(Path(save_file).with_suffix(".pdf"))
            try:
                convert_docx_to_pdf_win32(save_file, pdf_path)
            except Exception as pdf_err:
                raise RuntimeError(
                    f"The Word document was saved to:\n{save_file}\n\nHigh-resolution PDF export failed: {pdf_err}"
                ) from pdf_err

        result["success"] = True
        result["docx_path"] = str(save_file)
        result["pdf_path"] = pdf_path
        result["count"] = global_counter[0]
        result["skipped"] = all_skipped

    except PermissionError:
        result["docx_path"] = save_file
        result["skipped"].append((
            "<none>",
            f"Permission Denied when saving to:\n{save_file}\n\n"
            f"Please make sure the destination document is NOT open in Microsoft Word, or choose another filename."
        ))
    except Exception:
        result["skipped"].append(("<none>", f"Error details:\n{traceback.format_exc()}"))

    return result


# --------------------------------------------------------------------------- #
# Module 3 — Document & PDF generation (frame)                                #
# --------------------------------------------------------------------------- #

class DocGeneratorFrame(ctk.CTkFrame):
    def __init__(self, parent, organizer_source=None, **kw):
        super().__init__(parent, **kw)
        self.organizer_source = organizer_source
        self.sections_data = []
        self.generating = False
        self._build_ui()
        self.refresh_summary()

    def refresh_summary(self):
        if self.organizer_source and hasattr(self.organizer_source, "get_structured_sections"):
            self.sections_data = self.organizer_source.get_structured_sections()
        else:
            raw = get_selected_paths()
            self.sections_data = [{"title": "", "page_break": False, "plots": raw}] if raw else []

        total_plots = sum(len(s.get("plots", [])) for s in self.sections_data)
        num_secs = len(self.sections_data)

        if total_plots:
            sec_desc = f"{num_secs} section(s)" if any(s.get("title") for s in self.sections_data) else "Unsectioned"
            preview_plots = []
            for s in self.sections_data:
                preview_plots.extend(s.get("plots", []))
            preview = ", ".join(preview_plots[:2])
            extra = f"\n... and {total_plots - 2} more" if total_plots > 2 else ""
            self.summary_lbl.configure(
                text=f"Ready to export: {total_plots} plots across {sec_desc}.\nPreview: {preview}{extra}"
            )
        else:
            self.summary_lbl.configure(
                text="No plots selected yet.\nSwitch to 'Plot Selection', scan a folder, organize sections, then return here."
            )

    def _build_ui(self):
        layout = ctk.CTkFrame(self, corner_radius=10)
        layout.pack(fill="x", padx=15, pady=(8, 4))
        ctk.CTkLabel(layout, text="Page Layout, Sizing & Output Options:", font=ctk.CTkFont(size=13, weight="bold")).pack(anchor="w", padx=15, pady=(6, 4))

        row = ctk.CTkFrame(layout, fg_color="transparent")
        row.pack(fill="x", padx=15, pady=(0, 8))

        obox = ctk.CTkFrame(row, fg_color="transparent")
        obox.pack(side="left", padx=(0, 18))
        ctk.CTkLabel(obox, text="Page Orientation:", font=ctk.CTkFont(size=11, weight="bold")).pack(anchor="w")
        self.orient_var = ctk.StringVar(value="Landscape")
        ctk.CTkSegmentedButton(obox, values=["Landscape", "Portrait"], variable=self.orient_var).pack(anchor="w", pady=(2, 0))

        cbox = ctk.CTkFrame(row, fg_color="transparent")
        cbox.pack(side="left", padx=(0, 18))
        ctk.CTkLabel(cbox, text="Columns:", font=ctk.CTkFont(size=11, weight="bold")).pack(anchor="w")
        self.cols_var = ctk.StringVar(value="2 Columns")
        ctk.CTkSegmentedButton(cbox, values=["1 Column", "2 Columns", "3 Columns"], variable=self.cols_var).pack(anchor="w", pady=(2, 0))

        wbox = ctk.CTkFrame(row, fg_color="transparent")
        wbox.pack(side="left", padx=(0, 12))
        ctk.CTkLabel(wbox, text="Max Width (cm):", font=ctk.CTkFont(size=11, weight="bold")).pack(anchor="w")
        self.max_w = ctk.CTkEntry(wbox, width=90, height=28, placeholder_text="Auto")
        self.max_w.pack(anchor="w", pady=(2, 0))

        hbox = ctk.CTkFrame(row, fg_color="transparent")
        hbox.pack(side="left", padx=(0, 15))
        ctk.CTkLabel(hbox, text="Max Height (cm):", font=ctk.CTkFont(size=11, weight="bold")).pack(anchor="w")
        self.max_h = ctk.CTkEntry(hbox, width=90, height=28, placeholder_text="Auto")
        self.max_h.pack(anchor="w", pady=(2, 0))

        pbox = ctk.CTkFrame(row, fg_color="transparent")
        pbox.pack(side="left", padx=(5, 0), pady=(14, 0))
        self.export_pdf_var = ctk.BooleanVar(value=False)
        ctk.CTkCheckBox(pbox, text="Also export PDF", variable=self.export_pdf_var, font=ctk.CTkFont(size=12, weight="bold")).pack(anchor="w")

        summary = ctk.CTkFrame(self, corner_radius=10)
        summary.pack(fill="x", padx=15, pady=4)
        ctk.CTkLabel(summary, text="Report Structure Summary", font=ctk.CTkFont(size=13, weight="bold")).pack(anchor="w", padx=15, pady=(6, 4))
        self.summary_lbl = ctk.CTkLabel(summary, text="", justify="left", font=ctk.CTkFont(size=11))
        self.summary_lbl.pack(anchor="w", padx=15, pady=(0, 8))

        action = ctk.CTkFrame(self, corner_radius=10)
        action.pack(fill="x", padx=15, pady=(2, 6))
        self.export_btn = ctk.CTkButton(
            action,
            text="⚡ Generate Word Document with Organized Sections",
            height=38,
            font=ctk.CTkFont(size=14, weight="bold"),
            fg_color="#28a745",
            hover_color="#1e7e34",
            command=self._on_generate,
        )
        self.export_btn.pack(fill="x", padx=15, pady=6)

        status = ctk.CTkFrame(self, height=30, corner_radius=0)
        status.pack(fill="x", side="bottom")
        self.status_lbl = ctk.CTkLabel(status, text="Ready. Configure options and generate.", font=ctk.CTkFont(size=12))
        self.status_lbl.pack(side="left", padx=15, pady=2)
        self.progress_bar = ctk.CTkProgressBar(status, width=180, height=10)
        self.progress_bar.set(0)
        self.progress_bar.pack(side="right", padx=15, pady=2)

    def _on_generate(self):
        total_plots = sum(len(s.get("plots", [])) for s in self.sections_data)
        if total_plots == 0:
            messagebox.showwarning("No Plots", "No plots to generate. Please select and organize plots first.")
            return

        max_w = None
        w_val = self.max_w.get().strip()
        if w_val:
            try:
                max_w = float(w_val)
                if max_w <= 0:
                    raise ValueError
            except ValueError:
                messagebox.showerror("Invalid Width", "Max Width must be a positive number.")
                return

        max_h = None
        h_val = self.max_h.get().strip()
        if h_val:
            try:
                max_h = float(h_val)
                if max_h <= 0:
                    raise ValueError
            except ValueError:
                messagebox.showerror("Invalid Height", "Max Height must be a positive number.")
                return

        num_cols = 1
        if "2" in self.cols_var.get():
            num_cols = 2
        elif "3" in self.cols_var.get():
            num_cols = 3

        export_pdf = self.export_pdf_var.get()

        save_file = filedialog.asksaveasfilename(
            title="Save Generated Report",
            defaultextension=".docx",
            filetypes=[("Word Document", "*.docx")],
            initialfile=f"Plot_Report_{datetime.datetime.now().strftime('%Y%m%d-%H%M')}.docx",
        )
        if not save_file:
            return
        if not save_file.lower().endswith(".docx"):
            save_file += ".docx"

        self.generating = True
        self.export_btn.configure(state="disabled", text="Generating Document(s)... Please wait")
        self.progress_bar.configure(mode="determinate")
        self.progress_bar.set(0)

        threading.Thread(
            target=self._gen_worker,
            args=(self.sections_data, save_file, self.orient_var.get().lower(), num_cols, max_w, max_h, export_pdf),
            daemon=True,
        ).start()

    def _gen_worker(self, sections, save_file, orientation, num_cols, max_w, max_h, export_pdf):
        def progress(msg, pct):
            self.after(0, lambda: self._update_gen_progress(msg, pct))

        try:
            result = build_report_documents(
                sections, save_file, orientation, num_cols, max_w, max_h, export_pdf, progress_callback=progress
            )
            self.after(0, lambda: self._finish_generation(result))
        except Exception as e:
            self.after(0, lambda: self._finish_generation({
                "success": False, "docx_path": None, "pdf_path": None, "count": 0,
                "skipped": [("<none>", f"Unexpected error: {e}")],
            }))

    def _update_gen_progress(self, msg, pct):
        self.progress_bar.set(pct)
        self.status_lbl.configure(text=msg)

    def _finish_generation(self, result):
        self.generating = False
        self.export_btn.configure(state="normal", text="⚡ Generate Word Document with Organized Sections")
        self.progress_bar.set(1.0)
        if result["success"]:
            pdf_note = f"\nPDF: {result['pdf_path']}" if result["pdf_path"] else ""
            self.status_lbl.configure(text=f"Done! Generated report with {result['count']} plots.")
            extra = f"\n\nNote: {len(result['skipped'])} images were skipped." if result["skipped"] else ""
            open_it = messagebox.askyesno(
                "Document Generated",
                f"Successfully exported {result['count']} plots across sections!{extra}{pdf_note}\n\nWord Doc:\n{result['docx_path']}{pdf_note}\n\nOpen document now?",
            )
            if open_it:
                try:
                    target = result["pdf_path"] if (result["pdf_path"] and os.path.exists(result["pdf_path"])) else result["docx_path"]
                    os.startfile(target)
                except Exception as e:
                    messagebox.showerror("Error Opening", f"Could not launch file:\n{str(e)}")
        else:
            self.status_lbl.configure(text="Error generating document.")
            _show_error_dialog(self, "Generation Error", "\n".join(f"• {s}" for s in result["skipped"]) or result["skipped"][0][1])


def _show_error_dialog(parent, title, message):
    win = ctk.CTkToplevel(parent)
    win.title(title)
    win.geometry("600x400")
    win.grab_set()
    ctk.CTkLabel(win, text="An error occurred while generating the report:", font=ctk.CTkFont(weight="bold")).pack(anchor="w", padx=15, pady=(15, 5))
    txt = ctk.CTkTextbox(win, width=560, height=280)
    txt.pack(fill="both", expand=True, padx=15, pady=5)
    txt.insert("0.0", message)
    txt.configure(state="disabled")
    ctk.CTkButton(win, text="Close", width=100, command=win.destroy).pack(pady=10)
