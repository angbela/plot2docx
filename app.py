import os
import sys
import io
import datetime
import threading
import traceback
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from pathlib import Path
import customtkinter as ctk
from PIL import Image
import docx
from docx.shared import Inches, Cm, Pt, RGBColor
from docx.enum.section import WD_ORIENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

# Windows COM for Word to PDF conversion
try:
    import pythoncom
    import win32com.client
    HAS_WIN32COM = True
except ImportError:
    HAS_WIN32COM = False

IMAGE_EXTENSIONS = {".png", ".jpg", ".jpeg", ".bmp", ".tiff", ".tif", ".gif", ".webp"}

def set_table_borderless(table):
    """Remove all borders from a Word table so images float cleanly side-by-side."""
    tblPr = table._tbl.tblPr
    tblBorders = OxmlElement('w:tblBorders')
    for border_name in ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']:
        border = OxmlElement(f'w:{border_name}')
        border.set(qn('w:val'), 'none')
        tblBorders.append(border)
    tblPr.append(tblBorders)

def lock_row_cant_split(row):
    """Prevents Word from breaking this table row across two pages."""
    trPr = row._tr.get_or_add_trPr()
    trPr.append(OxmlElement('w:cantSplit'))

def convert_docx_to_pdf_win32(docx_path, pdf_path):
    """Use native installed Microsoft Word via COM to convert DOCX to 100% fidelity PDF."""
    pythoncom.CoInitialize()
    word = None
    try:
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False
        doc = word.Documents.Open(os.path.abspath(docx_path))
        # 17 = wdFormatPDF
        doc.SaveAs(os.path.abspath(pdf_path), FileFormat=17)
        doc.Close()
    finally:
        if word:
            word.Quit()
        pythoncom.CoUninitialize()

class PlotToDocxApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("Plot to Docx")
        self.geometry("1020x860")
        self.minsize(860, 680)
        
        # Set modern theme
        ctk.set_appearance_mode("System")
        ctk.set_default_color_theme("blue")

        self.scanning = False
        self.generating = False
        self.node_info = {}   # item_id -> {"type": "folder"/"file", "name": str, "path": path, "checked": bool, "children": []}
        self.parent_map = {}  # item_id -> parent_id

        self._build_ui()

    def _build_ui(self):
        # 1. Header Frame
        header_frame = ctk.CTkFrame(self, corner_radius=10)
        header_frame.pack(fill="x", padx=15, pady=(12, 6))

        title_lbl = ctk.CTkLabel(
            header_frame, 
            text="Plot to Docx", 
            font=ctk.CTkFont(size=22, weight="bold")
        )
        title_lbl.pack(anchor="w", padx=15, pady=(8, 2))

        subtitle_lbl = ctk.CTkLabel(
            header_frame, 
            text="Batch index plots, customize layout, columns, and dimensions, and export to Word & PDF.",
            font=ctk.CTkFont(size=12),
            text_color="gray"
        )
        subtitle_lbl.pack(anchor="w", padx=15, pady=(0, 8))

        # 2. File / Directory Inputs Frame
        inputs_frame = ctk.CTkFrame(self, corner_radius=10)
        inputs_frame.pack(fill="x", padx=15, pady=4)

        lbl_img_dir = ctk.CTkLabel(inputs_frame, text="Plots Folder / UNC Path:", font=ctk.CTkFont(weight="bold"))
        lbl_img_dir.grid(row=0, column=0, padx=15, pady=(8, 2), sticky="w")

        self.path_entry = ctk.CTkEntry(
            inputs_frame, 
            placeholder_text=r"e.g. C:\Pictures or \\Pentagon\Projects",
            height=32
        )
        self.path_entry.grid(row=1, column=0, padx=(15, 8), pady=(0, 10), sticky="ew")

        browse_dir_btn = ctk.CTkButton(inputs_frame, text="Browse...", width=80, height=32, command=self._on_browse_dir)
        browse_dir_btn.grid(row=1, column=1, padx=(0, 8), pady=(0, 10))

        self.scan_btn = ctk.CTkButton(
            inputs_frame, 
            text="Scan Plots", 
            width=95, 
            height=32,
            fg_color="#1f6aa5",
            hover_color="#144870",
            command=self._start_scan
        )
        self.scan_btn.grid(row=1, column=2, padx=(0, 15), pady=(0, 10))

        inputs_frame.grid_columnconfigure(0, weight=1)

        # 3. Layout & Dimension Controls Frame
        layout_frame = ctk.CTkFrame(self, corner_radius=10)
        layout_frame.pack(fill="x", padx=15, pady=4)

        lbl_layout = ctk.CTkLabel(layout_frame, text="Page Layout, Sizing & Output Options:", font=ctk.CTkFont(size=13, weight="bold"))
        lbl_layout.pack(anchor="w", padx=15, pady=(6, 4))

        controls_row = ctk.CTkFrame(layout_frame, fg_color="transparent")
        controls_row.pack(fill="x", padx=15, pady=(0, 8))

        # Page Orientation
        orient_box = ctk.CTkFrame(controls_row, fg_color="transparent")
        orient_box.pack(side="left", padx=(0, 18))
        ctk.CTkLabel(orient_box, text="Page Orientation:", font=ctk.CTkFont(size=11, weight="bold")).pack(anchor="w")
        self.orient_var = ctk.StringVar(value="Landscape")
        self.orient_combo = ctk.CTkSegmentedButton(
            orient_box, 
            values=["Landscape", "Portrait"], 
            variable=self.orient_var
        )
        self.orient_combo.pack(anchor="w", pady=(2, 0))

        # Columns per Page
        cols_box = ctk.CTkFrame(controls_row, fg_color="transparent")
        cols_box.pack(side="left", padx=(0, 18))
        ctk.CTkLabel(cols_box, text="Columns per Page:", font=ctk.CTkFont(size=11, weight="bold")).pack(anchor="w")
        self.cols_var = ctk.StringVar(value="2 Columns")
        self.cols_combo = ctk.CTkSegmentedButton(
            cols_box, 
            values=["1 Column", "2 Columns", "3 Columns"], 
            variable=self.cols_var
        )
        self.cols_combo.pack(anchor="w", pady=(2, 0))

        # Max Width (cm)
        w_box = ctk.CTkFrame(controls_row, fg_color="transparent")
        w_box.pack(side="left", padx=(0, 12))
        ctk.CTkLabel(w_box, text="Max Width (cm):", font=ctk.CTkFont(size=11, weight="bold")).pack(anchor="w")
        self.max_w_entry = ctk.CTkEntry(w_box, width=75, height=28, placeholder_text="Auto")
        self.max_w_entry.pack(anchor="w", pady=(2, 0))

        # Max Height (cm)
        h_box = ctk.CTkFrame(controls_row, fg_color="transparent")
        h_box.pack(side="left", padx=(0, 15))
        ctk.CTkLabel(h_box, text="Max Height (cm):", font=ctk.CTkFont(size=11, weight="bold")).pack(anchor="w")
        self.max_h_entry = ctk.CTkEntry(h_box, width=75, height=28, placeholder_text="Auto")
        self.max_h_entry.pack(anchor="w", pady=(2, 0))

        # PDF Export Checkbox
        pdf_box = ctk.CTkFrame(controls_row, fg_color="transparent")
        pdf_box.pack(side="left", padx=(5, 0), pady=(14, 0))
        self.export_pdf_var = ctk.BooleanVar(value=False)
        self.pdf_chk = ctk.CTkCheckBox(
            pdf_box, 
            text="Also export PDF", 
            variable=self.export_pdf_var,
            font=ctk.CTkFont(size=12, weight="bold")
        )
        self.pdf_chk.pack(anchor="w")

        # 4. Middle Content Frame: Tree View
        content_frame = ctk.CTkFrame(self, corner_radius=10)
        content_frame.pack(fill="both", expand=True, padx=15, pady=4)

        tree_header_frame = ctk.CTkFrame(content_frame, fg_color="transparent")
        tree_header_frame.pack(fill="x", padx=15, pady=(6, 4))

        self.tree_lbl = ctk.CTkLabel(
            tree_header_frame, 
            text="Plot Hierarchy (0 plots found, 0 selected)", 
            font=ctk.CTkFont(size=13, weight="bold")
        )
        self.tree_lbl.pack(side="left")

        unselect_all_btn = ctk.CTkButton(
            tree_header_frame, 
            text="Deselect All", 
            width=85, 
            height=24,
            fg_color="gray40",
            command=lambda: self._set_all_selection(False)
        )
        unselect_all_btn.pack(side="right", padx=(5, 0))

        select_all_btn = ctk.CTkButton(
            tree_header_frame, 
            text="Select All", 
            width=85, 
            height=24,
            fg_color="#1f6aa5",
            command=lambda: self._set_all_selection(True)
        )
        select_all_btn.pack(side="right")

        # Treeview styling via ttk
        style = ttk.Style()
        style.theme_use("clam")
        is_dark = ctk.get_appearance_mode() == "Dark"
        bg_col = "#2b2b2b" if is_dark else "#f9f9f9"
        fg_col = "white" if is_dark else "black"

        style.configure(
            "Treeview",
            background=bg_col,
            foreground=fg_col,
            fieldbackground=bg_col,
            rowheight=24,
            font=("Segoe UI", 10)
        )
        style.configure(
            "Treeview.Heading",
            background="#1f6aa5",
            foreground="white",
            relief="flat",
            font=("Segoe UI", 10, "bold")
        )
        style.map("Treeview.Heading", background=[('active', '#144870')])
        style.map("Treeview", background=[('selected', '#1f6aa5')], foreground=[('selected', 'white')])

        tree_container = tk.Frame(content_frame, bg=bg_col)
        tree_container.pack(fill="both", expand=True, padx=15, pady=(0, 8))

        self.tree = ttk.Treeview(
            tree_container,
            columns=("type", "size", "full_path"),
            show="tree headings",
            selectmode="browse"
        )
        self.tree.heading("#0", text="  Plot / Folder (Click checkbox to select)", anchor="w")
        self.tree.heading("type", text="Type", anchor="center")
        self.tree.heading("size", text="File Size", anchor="e")
        self.tree.heading("full_path", text="Full Path", anchor="w")

        self.tree.column("#0", width=380, minwidth=240)
        self.tree.column("type", width=80, minwidth=60, anchor="center")
        self.tree.column("size", width=90, minwidth=70, anchor="e")
        self.tree.column("full_path", width=380, minwidth=200)

        self.tree.bind("<ButtonRelease-1>", self._on_tree_click)

        vsb = ttk.Scrollbar(tree_container, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(tree_container, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")

        tree_container.grid_rowconfigure(0, weight=1)
        tree_container.grid_columnconfigure(0, weight=1)

        # 5. Action Frame
        action_frame = ctk.CTkFrame(self, corner_radius=10)
        action_frame.pack(fill="x", padx=15, pady=(2, 6))

        self.export_btn = ctk.CTkButton(
            action_frame,
            text="⚡ Generate Word Document with Selected Plots",
            height=38,
            font=ctk.CTkFont(size=14, weight="bold"),
            fg_color="#28a745",
            hover_color="#1e7e34",
            command=self._start_generation
        )
        self.export_btn.pack(fill="x", padx=15, pady=6)

        # 6. Status Bar
        status_frame = ctk.CTkFrame(self, height=30, corner_radius=0)
        status_frame.pack(fill="x", side="bottom")

        self.status_lbl = ctk.CTkLabel(
            status_frame, 
            text="Ready. Select plot folder and scan.",
            font=ctk.CTkFont(size=12)
        )
        self.status_lbl.pack(side="left", padx=15, pady=2)

        self.progress_bar = ctk.CTkProgressBar(status_frame, width=180, height=10)
        self.progress_bar.set(0)
        self.progress_bar.pack(side="right", padx=15, pady=2)

    # --- File/Folder Pickers ---
    def _on_browse_dir(self):
        selected_dir = filedialog.askdirectory(title="Select Plots Folder")
        if selected_dir:
            self.path_entry.delete(0, tk.END)
            self.path_entry.insert(0, os.path.normpath(selected_dir))

    def _format_size(self, size_in_bytes):
        for unit in ['B', 'KB', 'MB', 'GB']:
            if size_in_bytes < 1024.0:
                return f"{size_in_bytes:.1f} {unit}"
            size_in_bytes /= 1024.0
        return f"{size_in_bytes:.1f} TB"

    # --- Tree & Checkbox Logic ---
    def _get_item_label(self, name, item_type, is_checked):
        box = "☑" if is_checked else "☐"
        icon = "📁" if item_type == "folder" else "🖼"
        return f"{box} {icon} {name}"

    def _on_tree_click(self, event):
        item_id = self.tree.identify_row(event.y)
        if not item_id or item_id not in self.node_info:
            return

        current_checked = self.node_info[item_id]["checked"]
        new_state = not current_checked
        
        self._set_node_and_descendants_checked(item_id, new_state)
        self._update_ancestors(item_id)
        self._update_counts_display()

    def _set_node_and_descendants_checked(self, item_id, state):
        self.node_info[item_id]["checked"] = state
        node = self.node_info[item_id]
        new_text = self._get_item_label(node["name"], node["type"], state)
        self.tree.item(item_id, text=new_text)

        for child_id in node["children"]:
            self._set_node_and_descendants_checked(child_id, state)

    def _update_ancestors(self, item_id):
        parent_id = self.parent_map.get(item_id)
        if not parent_id or parent_id not in self.node_info:
            return

        p_node = self.node_info[parent_id]
        all_children_checked = all(self.node_info[cid]["checked"] for cid in p_node["children"])
        if p_node["checked"] != all_children_checked:
            p_node["checked"] = all_children_checked
            new_text = self._get_item_label(p_node["name"], p_node["type"], all_children_checked)
            self.tree.item(parent_id, text=new_text)

        self._update_ancestors(parent_id)

    def _set_all_selection(self, state):
        for item_id in self.tree.get_children():
            self._set_node_and_descendants_checked(item_id, state)
        self._update_counts_display()

    def _update_counts_display(self):
        total_plots = sum(1 for n in self.node_info.values() if n["type"] == "file")
        selected_plots = sum(1 for n in self.node_info.values() if n["type"] == "file" and n["checked"])
        self.tree_lbl.configure(text=f"Plot Hierarchy ({total_plots} plots found, {selected_plots} selected)")

    # --- Scanning ---
    def _start_scan(self):
        target_path = self.path_entry.get().strip()
        if not target_path:
            messagebox.showwarning("Empty Path", "Please enter or browse to a valid folder path.")
            return

        if not os.path.exists(target_path):
            messagebox.showerror("Path Not Found", f"The path does not exist or is not accessible:\n{target_path}")
            return

        if self.scanning or self.generating:
            return

        self.scanning = True
        self.scan_btn.configure(state="disabled", text="Scanning...")
        self.status_lbl.configure(text=f"Scanning plots in: {target_path} ...")
        self.progress_bar.configure(mode="indeterminate")
        self.progress_bar.start()

        for item in self.tree.get_children():
            self.tree.delete(item)
        self.node_info.clear()
        self.parent_map.clear()

        threading.Thread(target=self._scan_worker, args=(target_path,), daemon=True).start()

    def _scan_worker(self, root_dir):
        root_dir = os.path.normpath(root_dir)
        folder_id_map = {}

        root_name = os.path.basename(root_dir) or root_dir
        root_node_id = self.tree.insert(
            "", 
            "end", 
            text=self._get_item_label(root_name, "folder", True),
            values=("Folder", "", root_dir),
            open=True
        )
        self.node_info[root_node_id] = {
            "type": "folder", 
            "name": root_name, 
            "path": root_dir, 
            "checked": True, 
            "children": []
        }
        folder_id_map[root_dir] = root_node_id

        total_images = 0

        try:
            for current_root, dirs, files in os.walk(root_dir):
                img_files = [f for f in sorted(files) if Path(f).suffix.lower() in IMAGE_EXTENSIONS]
                
                if img_files or dirs:
                    rel = os.path.relpath(current_root, root_dir)
                    if rel != ".":
                        parts = rel.split(os.sep)
                        accum_path = root_dir
                        for part in parts:
                            prev_path = accum_path
                            accum_path = os.path.join(accum_path, part)
                            if accum_path not in folder_id_map:
                                parent_id = folder_id_map.get(prev_path, root_node_id)
                                fid = self.tree.insert(
                                    parent_id,
                                    "end",
                                    text=self._get_item_label(part, "folder", True),
                                    values=("Folder", "", accum_path),
                                    open=True
                                )
                                self.node_info[fid] = {
                                    "type": "folder",
                                    "name": part,
                                    "path": accum_path,
                                    "checked": True,
                                    "children": []
                                }
                                self.node_info[parent_id]["children"].append(fid)
                                self.parent_map[fid] = parent_id
                                folder_id_map[accum_path] = fid

                    curr_folder_id = folder_id_map.get(current_root, root_node_id)

                    for img_name in img_files:
                        img_path = os.path.join(current_root, img_name)
                        try:
                            fsize = os.path.getsize(img_path)
                            size_str = self._format_size(fsize)
                        except Exception:
                            size_str = "Unknown"

                        ext = Path(img_name).suffix.upper().replace(".", "")
                        img_item_id = self.tree.insert(
                            curr_folder_id,
                            "end",
                            text=self._get_item_label(img_name, "file", True),
                            values=(ext, size_str, img_path)
                        )
                        self.node_info[img_item_id] = {
                            "type": "file",
                            "name": img_name,
                            "path": img_path,
                            "checked": True,
                            "children": []
                        }
                        self.node_info[curr_folder_id]["children"].append(img_item_id)
                        self.parent_map[img_item_id] = curr_folder_id
                        total_images += 1

        except Exception as e:
            self.after(0, lambda: messagebox.showerror("Scan Error", f"Scan error:\n{str(e)}"))

        self.after(0, lambda: self._finish_scan(total_images))

    def _finish_scan(self, total_images):
        self.scanning = False
        self.scan_btn.configure(state="normal", text="Scan Plots")
        self.progress_bar.stop()
        self.progress_bar.configure(mode="determinate")
        self.progress_bar.set(1.0)
        self._update_counts_display()
        self.status_lbl.configure(text=f"Indexed {total_images} plots. All selected by default.")

    # --- Docx Generation ---
    def _start_generation(self):
        selected_plot_paths = [
            n["path"] for n in self.node_info.values() 
            if n["type"] == "file" and n["checked"]
        ]

        if not selected_plot_paths:
            messagebox.showwarning("No Plots Selected", "Please select at least one plot in the tree before generating.")
            return

        # Parse user dimensions
        max_w_user = None
        w_val = self.max_w_entry.get().strip()
        if w_val:
            try:
                max_w_user = float(w_val)
                if max_w_user <= 0:
                    raise ValueError
            except ValueError:
                messagebox.showerror("Invalid Width", "Please enter a positive numeric value for Max Width in cm (or leave empty).")
                return

        max_h_user = None
        h_val = self.max_h_entry.get().strip()
        if h_val:
            try:
                max_h_user = float(h_val)
                if max_h_user <= 0:
                    raise ValueError
            except ValueError:
                messagebox.showerror("Invalid Height", "Please enter a positive numeric value for Max Height in cm (or leave empty).")
                return

        # Parse columns
        cols_text = self.cols_var.get()
        num_cols = 1
        if "2" in cols_text:
            num_cols = 2
        elif "3" in cols_text:
            num_cols = 3

        orientation = self.orient_var.get().lower()
        export_pdf = self.export_pdf_var.get()

        # Generate default timestamped filename: Plot_Report_YYYYMMDD-HHMM.docx
        now_str = datetime.datetime.now().strftime("%Y%m%d-%H%M")
        default_filename = f"Plot_Report_{now_str}.docx"

        save_file = filedialog.asksaveasfilename(
            title="Save Generated Report",
            defaultextension=".docx",
            filetypes=[("Word Document", "*.docx")],
            initialfile=default_filename
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
            target=self._docx_worker, 
            args=(selected_plot_paths, save_file, orientation, num_cols, max_w_user, max_h_user, export_pdf), 
            daemon=True
        ).start()

    def _compute_image_display_size(self, pil_w, pil_h, max_col_w_cm, user_max_w_cm, user_max_h_cm, max_page_h_cm):
        """
        Calculate scaled width and height in centimeters while strictly locking aspect ratio.
        Guarantees that image height leaves ample room for its caption on the same page.
        """
        aspect_ratio = pil_w / float(pil_h)

        effective_max_w = min(max_col_w_cm, user_max_w_cm) if user_max_w_cm else max_col_w_cm
        target_w_cm = effective_max_w
        target_h_cm = target_w_cm / aspect_ratio

        # Determine effective maximum allowed height:
        # Either the user-defined max height, or page height minus caption headroom (~3.5cm)
        allowed_max_h = max_page_h_cm - 3.5
        if user_max_h_cm:
            allowed_max_h = min(allowed_max_h, user_max_h_cm)

        if target_h_cm > allowed_max_h:
            target_h_cm = max(2.5, allowed_max_h)
            target_w_cm = target_h_cm * aspect_ratio

        return target_w_cm, target_h_cm

    def _docx_worker(self, plot_paths, save_file, orientation, num_cols, user_max_w_cm, user_max_h_cm, export_pdf):
        total = len(plot_paths)
        try:
            # 1. Create Document with standard margins
            doc = docx.Document()
            for sec in doc.sections:
                sec.top_margin = Inches(0.8)
                sec.bottom_margin = Inches(0.8)
                sec.left_margin = Inches(0.8)
                sec.right_margin = Inches(0.8)

            # 2. Apply Page Orientation across ALL sections
            for sec in doc.sections:
                curr_w = sec.page_width
                curr_h = sec.page_height
                if orientation == "landscape":
                    sec.orientation = WD_ORIENT.LANDSCAPE
                    sec.page_width = max(curr_w, curr_h)
                    sec.page_height = min(curr_w, curr_h)
                else:
                    sec.orientation = WD_ORIENT.PORTRAIT
                    sec.page_width = min(curr_w, curr_h)
                    sec.page_height = max(curr_w, curr_h)

            # 3. Compute usable printable dimensions in cm
            last_sec = doc.sections[-1]
            try:
                page_w_emus = int(last_sec.page_width)
                page_h_emus = int(last_sec.page_height)
                left_emus = int(last_sec.left_margin)
                right_emus = int(last_sec.right_margin)
                top_emus = int(last_sec.top_margin)
                bottom_emus = int(last_sec.bottom_margin)

                usable_w_cm = ((page_w_emus - left_emus - right_emus) / 914400.0) * 2.54
                usable_h_cm = ((page_h_emus - top_emus - bottom_emus) / 914400.0) * 2.54
            except Exception:
                usable_w_cm = 24.0 if orientation == "landscape" else 16.5
                usable_h_cm = 16.0 if orientation == "landscape" else 24.0

            col_gap_cm = 0.6 if num_cols > 1 else 0.0
            max_col_w_cm = (usable_w_cm - (col_gap_cm * (num_cols - 1))) / num_cols
            max_col_w_cm = max(3.0, max_col_w_cm)

            fig_num = 1
            skipped_images = []

            # 4. Plot Images
            if num_cols == 1:
                for idx, img_p in enumerate(plot_paths):
                    pct = (idx + 1) / total * (0.8 if export_pdf else 1.0)
                    self.after(0, lambda i=idx+1, p=pct: self._update_gen_progress(f"Processing plot {i} of {total}...", p))

                    try:
                        with Image.open(img_p) as pil_img:
                            orig_w, orig_h = pil_img.size
                            img_stream = io.BytesIO()
                            if pil_img.mode in ("RGBA", "LA") or (pil_img.mode == "P" and "transparency" in pil_img.info):
                                pil_img.save(img_stream, format="PNG")
                            else:
                                pil_img.convert("RGB").save(img_stream, format="JPEG", quality=95)
                            img_stream.seek(0)
                    except Exception as img_err:
                        skipped_images.append((img_p, f"Cannot read image: {img_err}"))
                        continue

                    disp_w_cm, disp_h_cm = self._compute_image_display_size(
                        orig_w, orig_h, max_col_w_cm, user_max_w_cm, user_max_h_cm, usable_h_cm
                    )

                    try:
                        # Image paragraph: keep_with_next = True ensures it stays with caption
                        p_img = doc.add_paragraph()
                        p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        p_img.paragraph_format.space_before = Pt(8)
                        p_img.paragraph_format.space_after = Pt(2)
                        p_img.paragraph_format.keep_with_next = True

                        run_img = p_img.add_run()
                        run_img.add_picture(img_stream, width=Cm(disp_w_cm), height=Cm(disp_h_cm))

                        # Caption paragraph
                        p_cap = doc.add_paragraph()
                        p_cap.alignment = WD_ALIGN_PARAGRAPH.LEFT
                        p_cap.paragraph_format.space_before = Pt(2)
                        p_cap.paragraph_format.space_after = Pt(14)

                        r_bold = p_cap.add_run(f"Figure {fig_num}\t")
                        r_bold.bold = True
                        r_bold.font.size = Pt(9.5)
                        r_bold.font.name = "Segoe UI"
                        r_bold.font.color.rgb = RGBColor(40, 40, 40)

                        r_path = p_cap.add_run(img_p)
                        r_path.font.size = Pt(8.5)
                        r_path.font.name = "Segoe UI"
                        r_path.font.color.rgb = RGBColor(90, 90, 90)

                        fig_num += 1
                    except Exception as add_err:
                        skipped_images.append((img_p, f"Word insertion error: {add_err}"))
                        continue

            else:
                table = doc.add_table(rows=0, cols=num_cols)
                table.alignment = WD_TABLE_ALIGNMENT.CENTER
                set_table_borderless(table)

                for col in table.columns:
                    col.width = Cm(max_col_w_cm)

                for i in range(0, total, num_cols):
                    chunk = plot_paths[i:i + num_cols]
                    pct = min(1.0, (i + len(chunk)) / total) * (0.8 if export_pdf else 1.0)
                    self.after(0, lambda c=i+len(chunk), p=pct: self._update_gen_progress(f"Processing plot {c} of {total}...", p))

                    img_row = table.add_row()
                    cap_row = table.add_row()

                    # Prevent Word from splitting individual rows across pages
                    lock_row_cant_split(img_row)
                    lock_row_cant_split(cap_row)

                    for cell in img_row.cells:
                        cell.width = Cm(max_col_w_cm)
                        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
                    for cell in cap_row.cells:
                        cell.width = Cm(max_col_w_cm)
                        cell.vertical_alignment = WD_ALIGN_VERTICAL.TOP

                    for col_idx, img_p in enumerate(chunk):
                        try:
                            with Image.open(img_p) as pil_img:
                                orig_w, orig_h = pil_img.size
                                img_stream = io.BytesIO()
                                if pil_img.mode in ("RGBA", "LA") or (pil_img.mode == "P" and "transparency" in pil_img.info):
                                    pil_img.save(img_stream, format="PNG")
                                else:
                                    pil_img.convert("RGB").save(img_stream, format="JPEG", quality=95)
                                img_stream.seek(0)
                        except Exception as img_err:
                            skipped_images.append((img_p, f"Cannot read image: {img_err}"))
                            continue

                        disp_w_cm, disp_h_cm = self._compute_image_display_size(
                            orig_w, orig_h, max_col_w_cm, user_max_w_cm, user_max_h_cm, usable_h_cm
                        )

                        try:
                            cell_img = img_row.cells[col_idx]
                            p_img = cell_img.paragraphs[0]
                            p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
                            p_img.paragraph_format.space_before = Pt(4)
                            p_img.paragraph_format.space_after = Pt(2)
                            p_img.paragraph_format.keep_with_next = True

                            run_img = p_img.add_run()
                            run_img.add_picture(img_stream, width=Cm(disp_w_cm), height=Cm(disp_h_cm))

                            cell_cap = cap_row.cells[col_idx]
                            p_cap = cell_cap.paragraphs[0]
                            p_cap.alignment = WD_ALIGN_PARAGRAPH.LEFT
                            p_cap.paragraph_format.space_before = Pt(2)
                            p_cap.paragraph_format.space_after = Pt(12)

                            r_bold = p_cap.add_run(f"Figure {fig_num}\t")
                            r_bold.bold = True
                            r_bold.font.size = Pt(9.0)
                            r_bold.font.name = "Segoe UI"
                            r_bold.font.color.rgb = RGBColor(40, 40, 40)

                            r_path = p_cap.add_run(img_p)
                            r_path.font.size = Pt(8.0)
                            r_path.font.name = "Segoe UI"
                            r_path.font.color.rgb = RGBColor(90, 90, 90)

                            fig_num += 1
                        except Exception as add_err:
                            skipped_images.append((img_p, f"Word insertion error: {add_err}"))
                            continue

            # Save DOCX
            doc.save(save_file)

            # 5. Export to PDF if requested
            pdf_path = None
            if export_pdf:
                self.after(0, lambda: self._update_gen_progress("Converting to PDF via Microsoft Word...", 0.9))
                pdf_path = str(Path(save_file).with_suffix(".pdf"))
                convert_docx_to_pdf_win32(save_file, pdf_path)

            successful_count = fig_num - 1
            self.after(0, lambda: self._finish_generation(True, save_file, pdf_path, successful_count, skipped_images))

        except PermissionError:
            err_msg = (
                f"Permission Denied when saving to:\n{save_file}\n\n"
                "Please make sure the destination document is NOT currently open in Microsoft Word, "
                "or choose a different file name."
            )
            self.after(0, lambda: self._finish_generation(False, err_msg, None, 0, []))
        except Exception as e:
            err_details = traceback.format_exc()
            err_msg = f"Error details:\n{str(e)}\n\nTraceback:\n{err_details}"
            self.after(0, lambda: self._finish_generation(False, err_msg, None, 0, []))

    def _update_gen_progress(self, msg, pct):
        self.progress_bar.set(pct)
        self.status_lbl.configure(text=msg)

    def _finish_generation(self, success, docx_info, pdf_info, count, skipped):
        self.generating = False
        self.export_btn.configure(state="normal", text="⚡ Generate Word Document with Selected Plots")
        self.progress_bar.set(1.0)

        if success:
            extra_msg = ""
            if skipped:
                extra_msg = f"\n\nNote: {len(skipped)} images were skipped due to format issues."

            pdf_note = f"\nPDF: {pdf_info}" if pdf_info else ""
            self.status_lbl.configure(text=f"Done! Successfully generated document with {count} plots.")

            ask_open = messagebox.askyesno(
                "Documents Generated",
                f"Successfully plotted {count} images!{extra_msg}\n\nWord Doc:\n{docx_info}{pdf_note}\n\nDo you want to open the document now?"
            )
            if ask_open:
                try:
                    target_to_open = pdf_info if pdf_info and os.path.exists(pdf_info) else docx_info
                    os.startfile(target_to_open)
                except Exception as e:
                    messagebox.showerror("Error Opening File", f"Could not launch file:\n{str(e)}")
        else:
            self.status_lbl.configure(text="Error generating document.")
            self._show_error_dialog("Generation Error", docx_info)

    def _show_error_dialog(self, title, message):
        err_window = ctk.CTkToplevel(self)
        err_window.title(title)
        err_window.geometry("600x400")
        err_window.grab_set()

        lbl = ctk.CTkLabel(err_window, text="An error occurred while generating the report:", font=ctk.CTkFont(weight="bold"))
        lbl.pack(anchor="w", padx=15, pady=(15, 5))

        txt = ctk.CTkTextbox(err_window, width=560, height=280)
        txt.pack(fill="both", expand=True, padx=15, pady=5)
        txt.insert("0.0", message)
        txt.configure(state="disabled")

        btn = ctk.CTkButton(err_window, text="Close", width=100, command=err_window.destroy)
        btn.pack(pady=10)

if __name__ == "__main__":
    app = PlotToDocxApp()
    app.mainloop()
