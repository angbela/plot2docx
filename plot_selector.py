"""plot_selector.py — Module 1: plots selection.

Embeddable frame that:
  - accepts a folder / UNC path,
  - scans and builds the interactive plot hierarchy tree,
  - lets the user select which plots to include,
  - keeps an in-memory snapshot and persists it to selection.json so Module 2
    (doc_generator) can read it.

The document itself (page layout, columns, dimensions, Generate button) is
handled by doc_generator.py.

Also exposes a standalone `main()` for running this module alone.
"""

import datetime
import json
import os
import threading
from pathlib import Path

import customtkinter as ctk
import tkinter as tk
import tkinter.font as tkfont
from tkinter import filedialog, messagebox
from tkinter import ttk

# --------------------------------------------------------------------------- #
# Shared state file (read by Module 2 for external/non-tab usage)            #
# --------------------------------------------------------------------------- #

SELECTION_FILE = Path(__file__).resolve().parent / "selection.json"


def load_selected_paths() -> list:
    """Backup: read the selection snapshot written by PlotSelectionFrame."""
    try:
        data = json.loads(SELECTION_FILE.read_text(encoding="utf-8"))
        return [str(p) for p in data.get("selected_plots", [])]
    except Exception:
        return []


def get_selected_paths() -> list:
    """Return the current selection snapshot (plot_selector or disk fallback)."""
    return load_selected_paths()


def save_selected_paths(paths: list) -> None:
    """Atomically write the selection snapshot."""
    payload = {
        "saved_at": datetime.datetime.now().isoformat(timespec="seconds"),
        "selected_plots": [str(p) for p in paths],
    }
    tmp = SELECTION_FILE.with_suffix(".json.tmp")
    tmp.write_text(json.dumps(payload, indent=2), encoding="utf-8")
    tmp.replace(SELECTION_FILE)


# --------------------------------------------------------------------------- #
# Helpers                                                                   #
# --------------------------------------------------------------------------- #

IMAGE_EXTENSIONS = {".png", ".jpg", ".jpeg", ".bmp", ".tiff", ".tif", ".gif", ".webp"}


def format_size(size_in_bytes: int) -> str:
    for unit in ["B", "KB", "MB", "GB"]:
        if size_in_bytes < 1024.0:
            return f"{size_in_bytes:.1f} {unit}"
        size_in_bytes /= 1024.0
    return f"{size_in_bytes:.1f} TB"


# --------------------------------------------------------------------------- #
# Module 1 — Plots selection (frame)                                        #
# --------------------------------------------------------------------------- #

class PlotSelectionFrame(ctk.CTkFrame):
    """The plots-selection UI, embeddable in a notebook/tab."""

    def __init__(self, parent, **kw):
        super().__init__(parent, **kw)

        self.selected_paths = []   # in-memory snapshot of checked file paths
        self.scanning = False
        self.generating = False
        self.node_info = {}        # item_id -> {...}
        self.parent_map = {}       # item_id -> parent_id
        self.selection_changed = None  # callback when selection changes

        self._build_ui()

    # -- UI ---------------------------------------------------------------- #

    def _build_ui(self):
        # 2. Inputs ---------------------------------------------------------- #
        inputs = ctk.CTkFrame(self, corner_radius=10)
        inputs.pack(fill="x", padx=15, pady=(8, 4))

        ctk.CTkLabel(inputs, text="Plots Folder / UNC Path:", font=ctk.CTkFont(weight="bold")).grid(row=0, column=0, padx=15, pady=(8, 2), sticky="w")

        self.path_entry = ctk.CTkEntry(inputs, placeholder_text=r"e.g. C:\Pictures or \\Server\Projects", height=32)
        self.path_entry.grid(row=1, column=0, padx=(15, 8), pady=(0, 10), sticky="ew")

        ctk.CTkButton(inputs, text="Browse...", width=80, height=32, command=self._on_browse_dir).grid(row=1, column=1, padx=(0, 8), pady=(0, 10))

        self.scan_btn = ctk.CTkButton(
            inputs, text="Scan Plots", width=95, height=32, fg_color="#1f6aa5", hover_color="#144870", command=self._start_scan
        )
        self.scan_btn.grid(row=1, column=2, padx=(0, 15), pady=(0, 10))

        inputs.grid_columnconfigure(0, weight=1)

        # 3. Tree view -------------------------------------------------------- #
        content = ctk.CTkFrame(self, corner_radius=10)
        content.pack(fill="both", expand=True, padx=15, pady=4)

        header_frame = ctk.CTkFrame(content, fg_color="transparent")
        header_frame.pack(fill="x", padx=15, pady=(6, 4))

        self.tree_lbl = ctk.CTkLabel(header_frame, text="Plot Hierarchy (0 plots found, 0 selected)", font=ctk.CTkFont(size=13, weight="bold"))
        self.tree_lbl.pack(side="left")

        ctk.CTkButton(header_frame, text="Select All", width=85, height=24, fg_color="#1f6aa5", command=lambda: self._set_all_selection(True)).pack(side="right")
        ctk.CTkButton(header_frame, text="Deselect All", width=85, height=24, fg_color="gray40", command=lambda: self._set_all_selection(False)).pack(side="right", padx=(5, 0))

        style = ttk.Style()
        style.theme_use("clam")
        is_dark = ctk.get_appearance_mode() == "Dark"
        bg = "#2b2b2b" if is_dark else "#f9f9f9"
        style.configure("Treeview", background=bg, foreground="white" if is_dark else "black", fieldbackground=bg, rowheight=28, font=("Segoe UI", 14))
        style.configure("Treeview.Heading", background="#1f6aa5", foreground="white", relief="flat", font=("Segoe UI", 14, "bold"))
        style.map("Treeview.Heading", background=[("active", "#144870")])
        style.map("Treeview", background=[("selected", "#1f6aa5")], foreground=[("selected", "white")])

        self._tree_font = tkfont.Font(family="Segoe UI", size=14)
        self._tree_font_bold = tkfont.Font(family="Segoe UI", size=14, weight="bold")

        tree_container = tk.Frame(content, bg=bg)
        tree_container.pack(fill="both", expand=True, padx=15, pady=(0, 8))

        self.tree = ttk.Treeview(tree_container, columns=("type", "size", "full_path"), show="tree headings", selectmode="browse")
        self.tree.heading("#0", text="  Plot / Folder (Click checkbox to select)", anchor="w")
        self.tree.heading("type", text="Type", anchor="center")
        self.tree.heading("size", text="File Size", anchor="e")
        self.tree.heading("full_path", text="Full Path", anchor="w")
        self.tree.column("#0", width=420, minwidth=240, stretch=False)
        self.tree.column("type", width=55, minwidth=50, anchor="center", stretch=False)
        self.tree.column("size", width=80, minwidth=70, anchor="e", stretch=False)
        self.tree.column("full_path", width=900, minwidth=200, stretch=True)
        self.tree.bind("<ButtonRelease-1>", self._on_tree_click)

        vsb = ttk.Scrollbar(tree_container, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(tree_container, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        tree_container.grid_rowconfigure(0, weight=1)
        tree_container.grid_columnconfigure(0, weight=1)

        # 4. Status bar ------------------------------------------------------ #
        status = ctk.CTkFrame(self, height=30, corner_radius=0)
        status.pack(fill="x", side="bottom")
        self.status_lbl = ctk.CTkLabel(status, text="Ready. Enter a path and scan.", font=ctk.CTkFont(size=12))
        self.status_lbl.pack(side="left", padx=15, pady=2)
        self.progress_bar = ctk.CTkProgressBar(status, width=180, height=10)
        self.progress_bar.set(0)
        self.progress_bar.pack(side="right", padx=15, pady=2)

        # Selection is persisted on every change (scan / checkbox), so the
        # only remaining close path is the default window close.

    # -- File / Folder pickers --------------------------------------------- #

    def _on_browse_dir(self):
        selected = filedialog.askdirectory(title="Select Plots Folder")
        if selected:
            self.path_entry.delete(0, tk.END)
            self.path_entry.insert(0, os.path.normpath(selected))

    # -- Tree & checkbox logic -------------------------------------------- #

    def _get_item_label(self, name, item_type, is_checked):
        box = "☑" if is_checked else "☐"
        icon = "📁" if item_type == "folder" else "🖼"
        return f"{box} {icon} {name}"

    def _on_tree_click(self, event):
        item_id = self.tree.identify_row(event.y)
        if not item_id or item_id not in self.node_info:
            return
        new_state = not self.node_info[item_id]["checked"]
        self._set_node_and_descendants_checked(item_id, new_state)
        self._update_ancestors(item_id)
        self.refresh_selection()

    def _set_node_and_descendants_checked(self, item_id, state):
        self.node_info[item_id]["checked"] = state
        self.tree.item(item_id, text=self._get_item_label(self.node_info[item_id]["name"], self.node_info[item_id]["type"], state))
        for child_id in self.node_info[item_id]["children"]:
            self._set_node_and_descendants_checked(child_id, state)

    def _update_ancestors(self, item_id):
        parent_id = self.parent_map.get(item_id)
        if not parent_id or parent_id not in self.node_info:
            return
        p = self.node_info[parent_id]
        all_checked = all(self.node_info[cid]["checked"] for cid in p["children"])
        if p["checked"] != all_checked:
            p["checked"] = all_checked
            self.tree.item(parent_id, text=self._get_item_label(p["name"], p["type"], all_checked))
        self._update_ancestors(parent_id)

    def _set_all_selection(self, state):
        for item_id in self.tree.get_children():
            self._set_node_and_descendants_checked(item_id, state)
        self.refresh_selection()

    def _update_counts_display(self):
        total = sum(1 for n in self.node_info.values() if n["type"] == "file")
        selected = sum(1 for n in self.node_info.values() if n["type"] == "file" and n["checked"])
        self.tree_lbl.configure(text=f"Plot Hierarchy ({total} plots found, {selected} selected)")

    # -- Scanning ---------------------------------------------------------- #

    def _start_scan(self):
        target = self.path_entry.get().strip()
        if not target:
            messagebox.showwarning("Empty Path", "Please enter or browse to a valid folder path.")
            return
        if not os.path.exists(target):
            messagebox.showerror("Path Not Found", f"The path does not exist or is not accessible:\n{target}")
            return
        if self.scanning or self.generating:
            return

        self.scanning = True
        self.scan_btn.configure(state="disabled", text="Scanning...")
        self.status_lbl.configure(text=f"Scanning plots in: {target} ...")
        self.progress_bar.configure(mode="indeterminate")
        self.progress_bar.start()

        for item in self.tree.get_children():
            self.tree.delete(item)
        self.node_info.clear()
        self.parent_map.clear()

        threading.Thread(target=self._scan_worker, args=(target,), daemon=True).start()

    def _scan_worker(self, root_dir):
        root_dir = os.path.normpath(root_dir)
        folder_id_map = {}
        root_name = os.path.basename(root_dir) or root_dir
        root_id = self.tree.insert("", "end", text=self._get_item_label(root_name, "folder", True), values=("Folder", "", root_dir), open=True)
        self.node_info[root_id] = {"type": "folder", "name": root_name, "path": root_dir, "checked": True, "children": []}
        folder_id_map[root_dir] = root_id
        total = 0

        try:
            for current_root, dirs, files in os.walk(root_dir):
                img_files = [f for f in sorted(files) if Path(f).suffix.lower() in IMAGE_EXTENSIONS]
                if img_files or dirs:
                    rel = os.path.relpath(current_root, root_dir)
                    if rel != ".":
                        parts = rel.split(os.sep)
                        accum = root_dir
                        for part in parts:
                            prev = accum
                            accum = os.path.join(accum, part)
                            if accum not in folder_id_map:
                                parent = folder_id_map.get(prev, root_id)
                                fid = self.tree.insert(parent, "end", text=self._get_item_label(part, "folder", True), values=("Folder", "", accum), open=True)
                                self.node_info[fid] = {"type": "folder", "name": part, "path": accum, "checked": True, "children": []}
                                self.node_info[parent]["children"].append(fid)
                                self.parent_map[fid] = parent
                                folder_id_map[accum] = fid
                curr = folder_id_map.get(current_root, root_id)
                for img_name in img_files:
                    img_path = os.path.join(current_root, img_name)
                    try:
                        size_str = format_size(os.path.getsize(img_path))
                    except Exception:
                        size_str = "Unknown"
                    ext = Path(img_name).suffix.upper().replace(".", "")
                    img_id = self.tree.insert(curr, "end", text=self._get_item_label(img_name, "file", True), values=(ext, size_str, img_path))
                    self.node_info[img_id] = {"type": "file", "name": img_name, "path": img_path, "checked": True, "children": []}
                    self.node_info[curr]["children"].append(img_id)
                    self.parent_map[img_id] = curr
                    total += 1
        except Exception as e:
            self.after(0, lambda: messagebox.showerror("Scan Error", f"Scan error:\n{str(e)}"))

        self.after(0, lambda: self._finish_scan(total))

    def _finish_scan(self, total_images):
        self.scanning = False
        self.scan_btn.configure(state="normal", text="Scan Plots")
        self.progress_bar.stop()
        self.progress_bar.configure(mode="determinate")
        self.progress_bar.set(1.0)
        self.refresh_selection()
        self.status_lbl.configure(text=f"Indexed {total_images} plots. All selected by default.")


    def refresh_selection(self):
        """Rebuild the snapshot, update counters, persist to disk, notify Module 2."""
        self.selected_paths = [n["path"] for n in self.node_info.values() if n["type"] == "file" and n["checked"]]
        self._update_counts_display()
        save_selected_paths(self.selected_paths)
        if self.selection_changed:
            self.selection_changed()

    def get_selected_plots(self):
        """Return the current in-memory selected file paths."""
        return list(self.selected_paths)

# --------------------------------------------------------------------------- #
# Standalone entrypoint (optional — the tab app uses the frame below)        #
# --------------------------------------------------------------------------- #

def main():
    root = ctk.CTk()
    root.title("Plot to Docx — Plot Selection")
    root.geometry("1020x860")
    root.minsize(860, 680)
    ctk.set_appearance_mode("System")
    ctk.set_default_color_theme("blue")
    PlotSelectionFrame(root).pack(fill="both", expand=True)
    root.mainloop()


if __name__ == "__main__":
    main()
