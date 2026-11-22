"""app.py — main application entry point.

Unified modern window with 3-step navigation:
  1. Select Plots       (plot_selector.PlotSelectionFrame)
  2. Organize & Group   (plot_organizer.PlotOrganizerFrame)
  3. Generate Document  (doc_generator.DocGeneratorFrame)
"""

import customtkinter as ctk

from plot_selector import PlotSelectionFrame
from plot_organizer import PlotOrganizerFrame
from doc_generator import DocGeneratorFrame


class MainWindow(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Plot to Docx")
        self.geometry("1180x900")
        self.minsize(920, 740)

        ctk.set_appearance_mode("System")
        ctk.set_default_color_theme("blue")

        self.current_tab = "Select Plots"
        self._build_header_and_nav()

        # Content container where active module frames live
        self.content_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.content_frame.pack(fill="both", expand=True)

        # 3 modular tabs
        self.plot_tab = PlotSelectionFrame(self.content_frame)
        self.organizer_tab = PlotOrganizerFrame(self.content_frame, plot_source=self.plot_tab)
        self.doc_tab = DocGeneratorFrame(self.content_frame, organizer_source=self.organizer_tab)

        # Start on tab 1
        self.plot_tab.pack(fill="both", expand=True)

        self._link_events()

    def _build_header_and_nav(self):
        # 1. Main Application Header (persistent across modules)
        header_frame = ctk.CTkFrame(self, corner_radius=10)
        header_frame.pack(fill="x", padx=15, pady=(12, 6))

        title_lbl = ctk.CTkLabel(
            header_frame,
            text="Plot to Docx",
            font=ctk.CTkFont(size=22, weight="bold"),
        )
        title_lbl.pack(anchor="w", padx=15, pady=(8, 2))

        self.subtitle_lbl = ctk.CTkLabel(
            header_frame,
            text="Enter a folder or UNC path to index plots, click Scan, then select which to include.",
            font=ctk.CTkFont(size=12),
            text_color="gray",
        )
        self.subtitle_lbl.pack(anchor="w", padx=15, pady=(0, 10))

        # 2. Navigation bar with modern pill buttons
        nav_frame = ctk.CTkFrame(self, fg_color="transparent")
        nav_frame.pack(fill="x", padx=15, pady=(0, 4))

        self.nav_plot_btn = ctk.CTkButton(
            nav_frame,
            text="1. Select Plots",
            width=160,
            height=34,
            font=ctk.CTkFont(size=13, weight="bold"),
            command=lambda: self._select_tab("Select Plots"),
        )
        self.nav_plot_btn.pack(side="left", padx=(0, 8))

        self.nav_org_btn = ctk.CTkButton(
            nav_frame,
            text="2. Organize & Group",
            width=175,
            height=34,
            font=ctk.CTkFont(size=13, weight="bold"),
            command=lambda: self._select_tab("Organize & Group"),
        )
        self.nav_org_btn.pack(side="left", padx=(0, 8))

        self.nav_doc_btn = ctk.CTkButton(
            nav_frame,
            text="3. Generate Document",
            width=185,
            height=34,
            font=ctk.CTkFont(size=13, weight="bold"),
            command=lambda: self._select_tab("Generate Document"),
        )
        self.nav_doc_btn.pack(side="left")

        self._update_nav_styles()

    def _update_nav_styles(self):
        buttons = [
            ("Select Plots", self.nav_plot_btn),
            ("Organize & Group", self.nav_org_btn),
            ("Generate Document", self.nav_doc_btn),
        ]

        for tab_name, btn in buttons:
            if self.current_tab == tab_name:
                btn.configure(
                    fg_color=("#1f6aa5", "#1f6aa5"),
                    text_color="white",
                    border_width=1,
                    border_color=("#3b8ed0", "#3b8ed0"),
                )
            else:
                btn.configure(
                    fg_color=("#3a3a3a", "#2b2b2b"),
                    text_color=("gray85", "gray70"),
                    border_width=0,
                )

    def _select_tab(self, tab_name):
        if self.current_tab == tab_name:
            return

        self.current_tab = tab_name
        self._update_nav_styles()

        # Hide all tabs
        self.plot_tab.pack_forget()
        self.organizer_tab.pack_forget()
        self.doc_tab.pack_forget()

        if tab_name == "Select Plots":
            self.plot_tab.pack(fill="both", expand=True)
            self.subtitle_lbl.configure(
                text="Enter a folder or UNC path to index plots, click Scan, then select which to include."
            )
        elif tab_name == "Organize & Group":
            self.organizer_tab.sync_from_source()
            self.organizer_tab.pack(fill="both", expand=True)
            self.subtitle_lbl.configure(
                text="Reorder plots (▲/▼), add or rename sections, move plots between sections, and configure page breaks."
            )
        else:
            self.doc_tab.refresh_summary()
            self.doc_tab.pack(fill="both", expand=True)
            self.subtitle_lbl.configure(
                text="Customize page layout, columns, dimensions, and generate Word (.docx) & high-res PDF."
            )

    def _link_events(self):
        # When user checks/unchecks plots in Tab 1, sync to Tab 2 and Tab 3
        def on_selection_updated():
            self.organizer_tab.sync_from_source()
            self.doc_tab.refresh_summary()

        self.plot_tab.selection_changed = on_selection_updated
        self.organizer_tab.on_structure_changed = self.doc_tab.refresh_summary


if __name__ == "__main__":
    app = MainWindow()
    app.mainloop()
