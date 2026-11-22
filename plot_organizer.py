"""plot_organizer.py — Module 2: Organize & Group plots into sections.

Lets user:
  - Toggle between "Manual Grouping", "Auto-Group by Folder", and "No Grouping (Flat)".
  - Add, rename, delete, and reorder sections (Move Up / Down).
  - Bulk actions using checkboxes:
      - Select All / Deselect All across or within sections.
      - Bulk ▲ Move Up and ▼ Move Down.
      - Bulk "Move to Section...".
  - Configure whether each section starts on a page break.
  - Preview the final figure numbering (Section-based, e.g. Figure 1.1, Figure 1.2, Figure 2.1...).
"""

import os
from pathlib import Path
import customtkinter as ctk
import tkinter as tk
from tkinter import messagebox, simpledialog

from plot_selector import get_selected_paths


class PlotOrganizerFrame(ctk.CTkFrame):
    def __init__(self, parent, plot_source=None, **kw):
        super().__init__(parent, **kw)
        self.plot_source = plot_source
        self.raw_plots = []

        # Data structure for sections:
        # list of dicts: [
        #   {"id": str, "name": str, "page_break": bool, "plots": [path1, path2, ...]}
        # ]
        self.sections = []
        self._section_counter = 1

        # Checked plots set for bulk actions: set of plot paths (str)
        self.checked_plots = set()

        self.on_structure_changed = None

        self._build_ui()
        self.sync_from_source()

    def sync_from_source(self):
        """Called when user selects plots or switches to this tab."""
        if self.plot_source and hasattr(self.plot_source, "get_selected_plots"):
            selected = self.plot_source.get_selected_plots()
        else:
            selected = get_selected_paths()

        if selected == self.raw_plots and self.sections:
            self._render_sections_view()
            return

        self.raw_plots = list(selected)
        # Retain checked plots that are still in raw_plots
        self.checked_plots = self.checked_plots.intersection(set(self.raw_plots))
        self._rebuild_sections_from_raw()
        self._render_sections_view()

    def _rebuild_sections_from_raw(self):
        if not self.raw_plots:
            self.sections = []
            return

        mode = self.grouping_mode_var.get()
        if mode == "auto":
            groups = {}
            for p in self.raw_plots:
                folder_name = os.path.basename(os.path.dirname(p)) or "General"
                groups.setdefault(folder_name, []).append(p)

            self.sections = []
            for idx, (folder_name, plots) in enumerate(groups.items(), start=1):
                self.sections.append({
                    "id": f"sec_{idx}",
                    "name": folder_name,
                    "page_break": True,
                    "plots": plots,
                })
            self._section_counter = len(self.sections) + 1
        else:
            self.sections = [
                {
                    "id": "sec_1",
                    "name": "General Plots",
                    "page_break": True,
                    "plots": list(self.raw_plots),
                }
            ]
            self._section_counter = 2

    def _build_ui(self):
        # 1. Top Controls Bar
        top_bar = ctk.CTkFrame(self, corner_radius=10)
        top_bar.pack(fill="x", padx=15, pady=(8, 6))

        mode_box = ctk.CTkFrame(top_bar, fg_color="transparent")
        mode_box.pack(side="left", padx=15, pady=8)

        ctk.CTkLabel(mode_box, text="Grouping Strategy:", font=ctk.CTkFont(size=12, weight="bold")).pack(side="left", padx=(0, 10))

        self.grouping_mode_var = ctk.StringVar(value="manual")
        self.mode_segmented = ctk.CTkSegmentedButton(
            mode_box,
            values=["Manual Grouping", "Auto-Group by Folder", "No Grouping (Flat)"],
            command=self._on_mode_change,
            height=28,
        )
        self.mode_segmented.set("Manual Grouping")
        self.mode_segmented.pack(side="left")

        btn_box = ctk.CTkFrame(top_bar, fg_color="transparent")
        btn_box.pack(side="right", padx=15, pady=8)

        self.add_sec_btn = ctk.CTkButton(
            btn_box,
            text="+ Add New Section",
            width=140,
            height=28,
            fg_color="#1f6aa5",
            hover_color="#144870",
            command=self._add_section_dialog,
        )
        self.add_sec_btn.pack(side="left", padx=4)

        # 2. Scrollable Container for Section Cards
        self.scroll_canvas = ctk.CTkScrollableFrame(self, corner_radius=10)
        self.scroll_canvas.pack(fill="both", expand=True, padx=15, pady=(0, 6))

        # 3. Bottom Bulk Action Toolbar (Right before Total summary)
        self.toolbar_frame = ctk.CTkFrame(self, corner_radius=8, fg_color=("#e0e0e0", "#252525"))
        self.toolbar_frame.pack(fill="x", padx=15, pady=(0, 4))

        # Checkbox management tools on left
        tb_left = ctk.CTkFrame(self.toolbar_frame, fg_color="transparent")
        tb_left.pack(side="left", padx=10, pady=6)

        ctk.CTkButton(
            tb_left,
            text="Select All",
            width=80,
            height=26,
            fg_color="#3a3a3a",
            command=lambda: self._set_all_checks(True),
        ).pack(side="left", padx=(0, 4))

        ctk.CTkButton(
            tb_left,
            text="Deselect All",
            width=85,
            height=26,
            fg_color="#3a3a3a",
            command=lambda: self._set_all_checks(False),
        ).pack(side="left", padx=(0, 10))

        self.selected_badge = ctk.CTkLabel(
            tb_left,
            text="0 plots selected",
            font=ctk.CTkFont(size=12, weight="bold"),
            text_color="gray",
        )
        self.selected_badge.pack(side="left")

        # Bulk operation action buttons on right
        tb_right = ctk.CTkFrame(self.toolbar_frame, fg_color="transparent")
        tb_right.pack(side="right", padx=10, pady=6)

        self.bulk_up_btn = ctk.CTkButton(
            tb_right,
            text="▲ Move Up",
            width=90,
            height=26,
            fg_color="#3a3a3a",
            command=lambda: self._bulk_reorder(-1),
        )
        self.bulk_up_btn.pack(side="left", padx=4)

        self.bulk_down_btn = ctk.CTkButton(
            tb_right,
            text="▼ Move Down",
            width=100,
            height=26,
            fg_color="#3a3a3a",
            command=lambda: self._bulk_reorder(1),
        )
        self.bulk_down_btn.pack(side="left", padx=4)

        self.bulk_move_btn = ctk.CTkButton(
            tb_right,
            text="Move to Section...",
            width=140,
            height=26,
            fg_color="#1f6aa5",
            hover_color="#144870",
            command=self._bulk_move_dialog,
        )
        self.bulk_move_btn.pack(side="left", padx=(4, 0))

        # 4. Bottom Summary Info Bar
        bottom_bar = ctk.CTkFrame(self, height=30, corner_radius=0)
        bottom_bar.pack(fill="x", side="bottom")

        self.summary_lbl = ctk.CTkLabel(
            bottom_bar,
            text="0 sections, 0 plots.",
            font=ctk.CTkFont(size=12),
        )
        self.summary_lbl.pack(side="left", padx=15, pady=2)

    def _on_mode_change(self, val):
        if val == "Auto-Group by Folder":
            self.grouping_mode_var.set("auto")
            self.add_sec_btn.configure(state="disabled")
            self._rebuild_sections_from_raw()
        elif val == "No Grouping (Flat)":
            self.grouping_mode_var.set("flat")
            self.add_sec_btn.configure(state="disabled")
            self.sections = [{
                "id": "sec_flat",
                "name": "",
                "page_break": False,
                "plots": list(self.raw_plots),
            }]
        else:
            self.grouping_mode_var.set("manual")
            self.add_sec_btn.configure(state="normal")
            if not self.sections or self.sections[0].get("id") == "sec_flat":
                self._rebuild_sections_from_raw()

        self._render_sections_view()

    def _set_all_checks(self, state):
        if state:
            all_p = []
            for s in self.sections:
                all_p.extend(s["plots"])
            self.checked_plots = set(all_p)
        else:
            self.checked_plots.clear()
        self._render_sections_view()

    def _update_bulk_toolbar_state(self):
        count = len(self.checked_plots)
        self.selected_badge.configure(text=f"{count} plot(s) selected for bulk action")
        is_manual = (self.grouping_mode_var.get() == "manual")
        has_multiple_secs = (len(self.sections) > 1)

        btn_state = "normal" if count > 0 else "disabled"
        self.bulk_up_btn.configure(state=btn_state)
        self.bulk_down_btn.configure(state=btn_state)
        self.bulk_move_btn.configure(state="normal" if (count > 0 and is_manual and has_multiple_secs) else "disabled")

    def _render_sections_view(self):
        for widget in self.scroll_canvas.winfo_children():
            widget.destroy()

        total_plots = sum(len(s["plots"]) for s in self.sections)
        mode = self.grouping_mode_var.get()

        if not self.raw_plots:
            empty_lbl = ctk.CTkLabel(
                self.scroll_canvas,
                text="No plots selected yet.\nPlease go to 'Plot Selection', scan a directory and select plots first.",
                font=ctk.CTkFont(size=13),
                text_color="gray",
            )
            empty_lbl.pack(pady=40)
            self.summary_lbl.configure(text="0 sections, 0 plots selected.")
            self._update_bulk_toolbar_state()
            return

        self.summary_lbl.configure(
            text=f"Total: {len(self.sections)} section(s) | {total_plots} plot(s) ready for generation."
        )

        for s_idx, sec in enumerate(self.sections, start=1):
            self._create_section_card(sec, s_idx, len(self.sections))

        self._update_bulk_toolbar_state()

        if self.on_structure_changed:
            self.on_structure_changed()

    def _create_section_card(self, sec, sec_num, total_secs):
        mode = self.grouping_mode_var.get()
        is_flat = (mode == "flat")

        card = ctk.CTkFrame(self.scroll_canvas, corner_radius=8, fg_color=("#e6e6e6", "#232323"))
        card.pack(fill="x", pady=6, padx=4)

        header = ctk.CTkFrame(card, fg_color=("#d0d0d0", "#2c2c2c"), corner_radius=6)
        header.pack(fill="x", padx=6, pady=6)

        title_text = f"Section {sec_num}: {sec['name']}" if not is_flat else "All Plots (No Sections)"
        lbl_sec = ctk.CTkLabel(
            header,
            text=title_text,
            font=ctk.CTkFont(size=14, weight="bold"),
        )
        lbl_sec.pack(side="left", padx=10, pady=6)

        badge = ctk.CTkLabel(
            header,
            text=f"({len(sec['plots'])} plots)",
            font=ctk.CTkFont(size=12),
            text_color="gray",
        )
        badge.pack(side="left", padx=4)

        # Select all in section button
        def toggle_sec_plots(s=sec):
            sec_set = set(s["plots"])
            if sec_set.issubset(self.checked_plots):
                self.checked_plots.difference_update(sec_set)
            else:
                self.checked_plots.update(sec_set)
            self._render_sections_view()

        sec_all_checked = bool(sec["plots"]) and set(sec["plots"]).issubset(self.checked_plots)
        ctk.CTkButton(
            header,
            text="Select Section" if not sec_all_checked else "Deselect Section",
            width=110,
            height=22,
            font=ctk.CTkFont(size=11),
            fg_color="#3a3a3a",
            command=toggle_sec_plots,
        ).pack(side="left", padx=10)

        # Section Controls
        if not is_flat:
            pb_var = ctk.BooleanVar(value=sec.get("page_break", True))

            def on_toggle_pb(s=sec, v=pb_var):
                s["page_break"] = v.get()

            pb_chk = ctk.CTkCheckBox(
                header,
                text="Page Break Before",
                variable=pb_var,
                command=on_toggle_pb,
                font=ctk.CTkFont(size=11),
                width=14,
            )
            pb_chk.pack(side="right", padx=10)

            if mode == "manual" and total_secs > 1:
                del_btn = ctk.CTkButton(
                    header,
                    text="✕",
                    width=28,
                    height=24,
                    fg_color="#a83232",
                    hover_color="#7a2222",
                    command=lambda idx=sec_num-1: self._delete_section(idx),
                )
                del_btn.pack(side="right", padx=2)

            if sec_num < total_secs:
                down_sec_btn = ctk.CTkButton(
                    header,
                    text="▼",
                    width=28,
                    height=24,
                    fg_color="#3a3a3a",
                    command=lambda idx=sec_num-1: self._move_section(idx, 1),
                )
                down_sec_btn.pack(side="right", padx=2)

            if sec_num > 1:
                up_sec_btn = ctk.CTkButton(
                    header,
                    text="▲",
                    width=28,
                    height=24,
                    fg_color="#3a3a3a",
                    command=lambda idx=sec_num-1: self._move_section(idx, -1),
                )
                up_sec_btn.pack(side="right", padx=2)

            if mode == "manual":
                ren_btn = ctk.CTkButton(
                    header,
                    text="Rename",
                    width=65,
                    height=24,
                    fg_color="#3a3a3a",
                    command=lambda s=sec: self._rename_section(s),
                )
                ren_btn.pack(side="right", padx=4)

        plots_container = ctk.CTkFrame(card, fg_color="transparent")
        plots_container.pack(fill="x", padx=10, pady=(0, 6))

        if not sec["plots"]:
            empty_item = ctk.CTkLabel(
                plots_container,
                text="(No plots in this section)",
                font=ctk.CTkFont(size=11, slant="italic"),
                text_color="gray",
            )
            empty_item.pack(anchor="w", padx=10, pady=4)

        for p_idx, plot_path in enumerate(sec["plots"], start=1):
            fig_tag = f"Figure {sec_num}.{p_idx}" if not is_flat else f"Figure {p_idx}"
            self._create_plot_row(
                plots_container,
                sec,
                p_idx - 1,
                len(sec["plots"]),
                plot_path,
                fig_tag,
            )

    def _create_plot_row(self, container, sec, p_idx, total_plots, plot_path, fig_tag):
        row = ctk.CTkFrame(container, height=28, corner_radius=4, fg_color=("#f2f2f2", "#1e1e1e"))
        row.pack(fill="x", pady=2)

        # Plot selection checkbox for bulk action
        chk_var = ctk.BooleanVar(value=(plot_path in self.checked_plots))

        def on_chk_toggle(p=plot_path, v=chk_var):
            if v.get():
                self.checked_plots.add(p)
            else:
                self.checked_plots.discard(p)
            self._update_bulk_toolbar_state()

        chk = ctk.CTkCheckBox(
            row,
            text="",
            variable=chk_var,
            width=20,
            command=on_chk_toggle,
        )
        chk.pack(side="left", padx=(8, 4))

        lbl_fig = ctk.CTkLabel(
            row,
            text=fig_tag,
            width=85,
            font=ctk.CTkFont(size=11, weight="bold"),
            anchor="w",
        )
        lbl_fig.pack(side="left", padx=4)

        filename = os.path.basename(plot_path)
        lbl_file = ctk.CTkLabel(
            row,
            text=filename,
            font=ctk.CTkFont(size=12),
            anchor="w",
        )
        lbl_file.pack(side="left", padx=4)

        lbl_full = ctk.CTkLabel(
            row,
            text=f"({plot_path})",
            font=ctk.CTkFont(size=10),
            text_color="gray",
            anchor="w",
        )
        lbl_full.pack(side="left", padx=4)

        # Quick individual actions on right
        if len(self.sections) > 1 and self.grouping_mode_var.get() == "manual":
            move_sec_btn = ctk.CTkButton(
                row,
                text="Move to...",
                width=65,
                height=20,
                font=ctk.CTkFont(size=10),
                fg_color="#3a3a3a",
                command=lambda s=sec, idx=p_idx: self._move_plot_to_section_dialog(s, idx),
            )
            move_sec_btn.pack(side="right", padx=4)

        if p_idx < total_plots - 1:
            btn_down = ctk.CTkButton(
                row,
                text="▼",
                width=24,
                height=20,
                font=ctk.CTkFont(size=10),
                fg_color="#3a3a3a",
                command=lambda s=sec, idx=p_idx: self._reorder_plot(s, idx, 1),
            )
            btn_down.pack(side="right", padx=2)

        if p_idx > 0:
            btn_up = ctk.CTkButton(
                row,
                text="▲",
                width=24,
                height=20,
                font=ctk.CTkFont(size=10),
                fg_color="#3a3a3a",
                command=lambda s=sec, idx=p_idx: self._reorder_plot(s, idx, -1),
            )
            btn_up.pack(side="right", padx=2)

    # ----------------------------------------------------------------------- #
    # Bulk & Individual Reordering / Moving Logic                             #
    # ----------------------------------------------------------------------- #

    def _bulk_reorder(self, direction):
        """
        Move all checked plots up (direction = -1) or down (direction = 1)
        within their respective sections simultaneously.
        """
        if not self.checked_plots:
            return

        for sec in self.sections:
            plots = sec["plots"]
            n = len(plots)
            if direction == -1:
                # Move up: iterate from index 1 to n-1
                for i in range(1, n):
                    if plots[i] in self.checked_plots and plots[i - 1] not in self.checked_plots:
                        plots[i], plots[i - 1] = plots[i - 1], plots[i]
            else:
                # Move down: iterate from index n-2 down to 0
                for i in range(n - 2, -1, -1):
                    if plots[i] in self.checked_plots and plots[i + 1] not in self.checked_plots:
                        plots[i], plots[i + 1] = plots[i + 1], plots[i]

        self._render_sections_view()

    def _bulk_move_dialog(self):
        """Transfer all checked plots to a chosen target section."""
        if not self.checked_plots or len(self.sections) <= 1:
            return

        dialog = ctk.CTkToplevel(self)
        dialog.title("Bulk Move Plots")
        dialog.geometry("380x220")
        dialog.grab_set()

        ctk.CTkLabel(
            dialog,
            text=f"Move {len(self.checked_plots)} checked plot(s) to section:",
            font=ctk.CTkFont(size=12, weight="bold"),
        ).pack(padx=15, pady=(15, 8))

        var_choice = ctk.StringVar(value=self.sections[0]["name"])
        combo = ctk.CTkComboBox(
            dialog,
            values=[s["name"] for s in self.sections],
            variable=var_choice,
            width=260,
        )
        combo.pack(pady=10)

        def do_bulk_move():
            chosen_name = var_choice.get()
            target_sec = None
            for s in self.sections:
                if s["name"] == chosen_name:
                    target_sec = s
                    break

            if target_sec:
                plots_to_move = list(self.checked_plots)
                for s in self.sections:
                    if s["id"] != target_sec["id"]:
                        s["plots"] = [p for p in s["plots"] if p not in self.checked_plots]
                # Append to target section while avoiding duplicate items
                existing_set = set(target_sec["plots"])
                for p in plots_to_move:
                    if p not in existing_set:
                        target_sec["plots"].append(p)
                        existing_set.add(p)

            dialog.destroy()
            self._render_sections_view()

        ctk.CTkButton(dialog, text="Confirm Move", command=do_bulk_move, width=120).pack(pady=(15, 5))
        ctk.CTkButton(dialog, text="Cancel", command=dialog.destroy, width=120, fg_color="gray40").pack(pady=4)

    def _reorder_plot(self, sec, p_idx, direction):
        target = p_idx + direction
        if 0 <= target < len(sec["plots"]):
            sec["plots"][p_idx], sec["plots"][target] = sec["plots"][target], sec["plots"][p_idx]
            self._render_sections_view()

    def _move_section(self, s_idx, direction):
        target = s_idx + direction
        if 0 <= target < len(self.sections):
            self.sections[s_idx], self.sections[target] = self.sections[target], self.sections[s_idx]
            self._render_sections_view()

    def _rename_section(self, sec):
        new_name = simpledialog.askstring("Rename Section", "Enter new section name:", initialvalue=sec["name"])
        if new_name and new_name.strip():
            sec["name"] = new_name.strip()
            self._render_sections_view()

    def _delete_section(self, s_idx):
        if len(self.sections) <= 1:
            return
        sec = self.sections[s_idx]
        confirm = messagebox.askyesno(
            "Delete Section",
            f"Are you sure you want to delete section '{sec['name']}'?\n"
            f"Its {len(sec['plots'])} plot(s) will be moved to the first section."
        )
        if confirm:
            plots_to_salvage = sec["plots"]
            del self.sections[s_idx]
            self.sections[0]["plots"].extend(plots_to_salvage)
            self._render_sections_view()

    def _add_section_dialog(self):
        name = simpledialog.askstring("New Section", "Enter section name:", initialvalue=f"Section {self._section_counter}")
        if name and name.strip():
            new_id = f"sec_{self._section_counter}"
            self._section_counter += 1
            self.sections.append({
                "id": new_id,
                "name": name.strip(),
                "page_break": True,
                "plots": [],
            })
            self._render_sections_view()

    def _move_plot_to_section_dialog(self, source_sec, p_idx):
        plot_p = source_sec["plots"][p_idx]
        target_options = [s for s in self.sections if s["id"] != source_sec["id"]]
        if not target_options:
            return

        dialog = ctk.CTkToplevel(self)
        dialog.title("Move Plot to Section")
        dialog.geometry("380x220")
        dialog.grab_set()

        ctk.CTkLabel(
            dialog,
            text=f"Move '{os.path.basename(plot_p)}' to:",
            font=ctk.CTkFont(size=12, weight="bold"),
        ).pack(padx=15, pady=(15, 8))

        var_choice = ctk.StringVar(value=target_options[0]["name"])
        combo = ctk.CTkComboBox(
            dialog,
            values=[s["name"] for s in target_options],
            variable=var_choice,
            width=260,
        )
        combo.pack(pady=10)

        def do_move():
            chosen_name = var_choice.get()
            for ts in target_options:
                if ts["name"] == chosen_name:
                    source_sec["plots"].remove(plot_p)
                    ts["plots"].append(plot_p)
                    break
            dialog.destroy()
            self._render_sections_view()

        ctk.CTkButton(dialog, text="Confirm Move", command=do_move, width=120).pack(pady=(15, 5))
        ctk.CTkButton(dialog, text="Cancel", command=dialog.destroy, width=120, fg_color="gray40").pack(pady=4)

    def get_structured_sections(self):
        result = []
        is_flat = (self.grouping_mode_var.get() == "flat")
        for s in self.sections:
            if not s["plots"]:
                continue
            result.append({
                "title": "" if is_flat else s["name"],
                "page_break": False if is_flat else s.get("page_break", True),
                "plots": list(s["plots"]),
            })
        return result
