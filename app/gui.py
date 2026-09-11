from __future__ import annotations

import os
import json
import queue
import re
import subprocess
import sys
import threading
import time
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Optional

import tkinter as tk
from tkinter import filedialog, font as tkfont, messagebox, simpledialog, ttk

from PIL import Image, ImageTk

try:
    from ttkbootstrap import Window
except Exception:
    Window = None

from .exporter import export
from .mapping import (
    DepartmentPlantListError,
    load_department_plant_file,
    load_mapping,
    merge_mappings,
    parse_department_plant_text,
)
from .models import ChangeRuleDraftRecord, FileRecord, MappingIndex, PageRecord, norm
from .preview import VisioRenderer, render
from .package_io import (
    AnalysisPackageError,
    import_analysis_package,
    save_analysis_package,
)
from .vsdx import analyze, apply_mapping, load_page, quick_scan
from .routing import (
    ChangeRuleRow, RoutingReference, compare_change_rule_rows, department_readiness,
    export_change_rule_subset, export_into_template, export_routing_generation_package,
    generate_change_rule_draft, load_routing_reference, merge_preserve_manual_edits,
    readiness_summary, validate_change_rule_rows,
)
from .material_selection import (
    MaterialSelectionError, MaterialSelectionSources, export_material_selection_template,
    generate_material_selection, load_code_reference, load_p41_routing, workbook_sheets,
)


TITLE = "VSM to ROUTING Preparation Tool v.1.3.8"


class ToolTip:
    """Small, dependency-free tooltip for controls whose labels use routing jargon."""
    def __init__(self, widget, text: str):
        self.widget, self.text, self.window = widget, text, None
        widget.bind("<Enter>", self.show, add="+")
        widget.bind("<Leave>", self.hide, add="+")

    def show(self, _event=None):
        if self.window:
            return
        x = self.widget.winfo_rootx() + 18
        y = self.widget.winfo_rooty() + self.widget.winfo_height() + 4
        self.window = tk.Toplevel(self.widget)
        self.window.wm_overrideredirect(True)
        self.window.wm_geometry(f"+{x}+{y}")
        ttk.Label(self.window, text=self.text, style="ToolTip.TLabel", padding=(7, 4), wraplength=300).pack()

    def hide(self, _event=None):
        if self.window:
            self.window.destroy()
            self.window = None


class RoundedButton(tk.Canvas):
    """A compact canvas button with dependable rounded corners on Windows."""

    def __init__(self, master, text="", command=None, style="TButton", state=tk.NORMAL, **kwargs):
        self._text = text
        self._command = command
        self._style_name = style or "TButton"
        self._button_state = state
        self._hovered = False
        self._font = ("Aptos", 9, "bold")
        parent_style = ttk.Style(master)
        parent_background = parent_style.lookup(master.winfo_class(), "background") or "#EEF5FF"
        super().__init__(master, background=parent_background, highlightthickness=0, borderwidth=0, takefocus=1, cursor="hand2", **kwargs)
        self.bind("<Configure>", self._redraw, add="+")
        self.bind("<Enter>", self._on_enter, add="+")
        self.bind("<Leave>", self._on_leave, add="+")
        self.bind("<ButtonRelease-1>", self._on_click, add="+")
        self._set_requested_size()

    def _set_requested_size(self):
        font = tkfont.Font(font=self._font)
        super().configure(width=max(94, font.measure(self._text) + 30), height=36)

    def _palette(self):
        disabled = self._button_state == tk.DISABLED
        primary = self._style_name == "Primary.TButton"
        secondary = self._style_name == "Secondary.TButton"
        if disabled:
            return "#DCE7F8" if primary else "#EAF0F9", "#60759C", "#C9D8EF"
        if primary:
            return ("#0D4BD7" if self._hovered else "#1463F3"), "#FFFFFF", "#0A3FBC"
        if secondary:
            return ("#E9F1FF" if self._hovered else "#FFFFFF"), "#0B1F4D", "#9DB8E8"
        return ("#F2F6FD" if self._hovered else "#FFFFFF"), "#0B1F4D", "#B9C9E4"

    def _rounded_rectangle(self, x1, y1, x2, y2, radius, fill, outline):
        # Fill pieces have no outline, preventing internal construction lines.
        self.create_rectangle(x1 + radius, y1, x2 - radius, y2, fill=fill, outline="")
        self.create_rectangle(x1, y1 + radius, x2, y2 - radius, fill=fill, outline="")
        corners = (
            (x1, y1, x1 + radius * 2, y1 + radius * 2),
            (x1, y2 - radius * 2, x1 + radius * 2, y2),
            (x2 - radius * 2, y2 - radius * 2, x2, y2),
            (x2 - radius * 2, y1, x2, y1 + radius * 2),
        )
        for box in corners:
            self.create_oval(*box, fill=fill, outline="")
        # Draw just one outside border around the finished surface.
        self.create_line(x1 + radius, y1, x2 - radius, y1, fill=outline)
        self.create_line(x2, y1 + radius, x2, y2 - radius, fill=outline)
        self.create_line(x2 - radius, y2, x1 + radius, y2, fill=outline)
        self.create_line(x1, y2 - radius, x1, y1 + radius, fill=outline)
        for box, start in zip(corners, (90, 180, 270, 0)):
            self.create_arc(*box, start=start, extent=90, style=tk.ARC, outline=outline)

    def _redraw(self, _event=None):
        width, height = max(self.winfo_width(), 2), max(self.winfo_height(), 2)
        fill, foreground, outline = self._palette()
        self.delete("all")
        self._rounded_rectangle(1, 1, width - 1, height - 1, min(13, height // 2 - 1), fill, outline)
        self.create_text(width / 2, height / 2, text=self._text, fill=foreground, font=self._font)

    def _on_enter(self, _event=None):
        if self._button_state != tk.DISABLED:
            self._hovered = True
            self._redraw()

    def _on_leave(self, _event=None):
        self._hovered = False
        self._redraw()

    def _on_click(self, _event=None):
        if self._button_state != tk.DISABLED and self._command:
            self._command()

    def configure(self, cnf=None, **kwargs):
        if isinstance(cnf, dict):
            kwargs = {**cnf, **kwargs}
        for key, attribute in (("text", "_text"), ("command", "_command"), ("style", "_style_name"), ("state", "_button_state")):
            if key in kwargs:
                setattr(self, attribute, kwargs.pop(key))
        result = super().configure(**kwargs) if kwargs else None
        if hasattr(self, "_text"):
            self._set_requested_size()
            self._redraw()
        return result

    config = configure

    def cget(self, key):
        if key == "text":
            return self._text
        if key == "command":
            return self._command
        if key == "style":
            return self._style_name
        if key == "state":
            return self._button_state
        return super().cget(key)


# Route this desktop application's controls through the rounded renderer.
ttk.Button = RoundedButton


class App:
    def __init__(self):
        # Windows' native ttk theme ignores several foreground/background maps
        # for disabled buttons.  Use clam so the operational button states stay
        # visibly distinct on every supported Windows desktop.
        self.root = tk.Tk()
        self.root.title(TITLE)
        self.root.geometry("1660x960")
        self.root.minsize(1250, 760)

        self.workbook_mapping: Optional[MappingIndex] = None
        self.manual_mapping: Optional[MappingIndex] = None
        self.mapping: Optional[MappingIndex] = None
        self.routing_reference: Optional[RoutingReference] = None
        self.routing_template_path: str = ""
        self.material_sources = MaterialSelectionSources()
        self.material_rows = []
        self.material_unmatched = []
        self.step0_loading = False
        self.routing_source_ready = False

        self.files: List[FileRecord] = []
        self.file_map: Dict[str, FileRecord] = {}
        self.page_map: Dict[str, PageRecord] = {}
        self.selected_file: Optional[FileRecord] = None
        self.selected_page: Optional[PageRecord] = None

        self.events: queue.Queue = queue.Queue()
        self.logs: List[str] = []
        self.running = False
        self.started = 0.0
        self.preview_photo = None
        self.zoom = 1.0
        self.exact_thread = None
        self.preview_request_token = 0
        self.auto_preview = tk.BooleanVar(value=True)
        self.highlight_workcenters = tk.BooleanVar(value=False)
        self.show_process_path = tk.BooleanVar(value=False)
        self.show_shape_ids = tk.BooleanVar(value=False)
        self.plant_saved_value = ""
        self.dirty_reasons = set()
        self.ui_settings_path = Path.home() / ".vsdx_routing_generation_ui.json"

        self.style()
        self.build()
        self.load_ui_settings()
        self.bind_shortcuts()
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)
        self.poll()
        self.elapsed()
        self.auto_mapping()

    def style(self):
        style = ttk.Style(self.root)
        style.theme_use("clam")
        families = set(tkfont.families())
        display_font = "Aptos Display" if "Aptos Display" in families else "Segoe UI Variable Display" if "Segoe UI Variable Display" in families else "Segoe UI"
        body_font = "Aptos" if "Aptos" in families else "Segoe UI Variable" if "Segoe UI Variable" in families else "Segoe UI"
        bg, surface, ink = "#EEF5FF", "#FFFFFF", "#0B1F4D"
        muted, blue, blue_active, border = "#5F739C", "#1463F3", "#0D4BD7", "#D8E4F7"
        self.root.configure(background=bg)
        style.configure("TFrame", background=bg)
        style.configure("TLabelframe", background=surface, bordercolor=border)
        style.configure("TLabelframe.Label", background=surface, foreground=ink)
        style.configure("Title.TLabel", background=bg, foreground=ink, font=(display_font, 21, "bold"))
        style.configure(
            "Sub.TLabel",
            background=bg,
            font=(body_font, 9),
            foreground=muted,
        )
        style.configure("Card.TLabelframe", background=surface, padding=14, bordercolor=border, relief="solid")
        style.configure(
            "Card.TLabelframe.Label",
            background=surface,
            foreground=ink,
            font=(display_font, 10, "bold"),
        )
        style.configure("Treeview", background=surface, fieldbackground=surface, foreground=ink, bordercolor=border, rowheight=31, font=(body_font, 9))
        style.configure(
            "Treeview.Heading",
            background="#F7FAFF",
            foreground=ink,
            bordercolor=border,
            font=(body_font, 9, "bold"),
            padding=(8, 8),
        )
        style.configure(
            "Kpi.TLabel",
            background=bg,
            font=(body_font, 10, "bold"),
            foreground=blue,
        )
        style.configure("TButton", background=surface, foreground=ink, bordercolor="#B9C9E4", font=(body_font, 9), padding=(10, 6), relief="raised", borderwidth=1)
        style.map("TButton", background=[("disabled", "#E6EDF8"), ("active", "#F2F6FD")], foreground=[("disabled", "#6B7D9F")], bordercolor=[("disabled", "#D4DFEF")])
        style.configure("Primary.TButton", background=blue, foreground="#FFFFFF", bordercolor=blue_active, lightcolor=blue, darkcolor="#0A3FBC", focuscolor=blue, font=(body_font, 10, "bold"), padding=(16, 9), relief="raised", borderwidth=1)
        style.map("Primary.TButton", background=[("disabled", "#DCE7F8"), ("pressed", "#0A3FBC"), ("active", blue_active), ("!disabled", blue)], foreground=[("disabled", "#60759C"), ("!disabled", "#FFFFFF")], bordercolor=[("disabled", "#C9D8EF"), ("pressed", "#082F8F"), ("!disabled", blue_active)])
        style.configure("Secondary.TButton", background=surface, foreground=ink, bordercolor="#9DB8E8", font=(body_font, 10, "bold"), padding=(14, 8), relief="raised", borderwidth=1)
        style.map("Secondary.TButton", background=[("disabled", "#EAF0F9"), ("pressed", "#DCEBFF"), ("active", "#E9F1FF"), ("!disabled", surface)], foreground=[("disabled", "#64748B"), ("!disabled", ink)], bordercolor=[("disabled", "#D4DFEF"), ("!disabled", "#9DB8E8")])
        style.configure("TEntry", fieldbackground="#FBFDFF", foreground=ink, bordercolor=border, padding=5)
        style.configure("TCombobox", fieldbackground="#FBFDFF", foreground=ink, bordercolor=border, padding=4)
        style.configure("ToolTip.TLabel", background="#0B1F4D", foreground="#FFFFFF", font=(body_font, 9))
        style.map("Treeview", background=[("selected", blue)], foreground=[("selected", "#FFFFFF")])
        style.configure("Draft.Treeview", rowheight=31)
        style.map("Draft.Treeview", background=[("selected", blue)], foreground=[("selected", "#FFFFFF")])

    def build(self):
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(2, weight=1)

        header = ttk.Frame(self.root, padding=(16, 12))
        header.grid(row=0, column=0, sticky="ew")
        header.columnconfigure(0, weight=1)

        ttk.Label(header, text=TITLE, style="Title.TLabel").grid(
            row=0, column=0, sticky="w"
        )
        ttk.Label(
            header,
            text=(
                "Confirm raw business data, read VSM (VSDX) process flows, then prepare reviewable Change Rule and Material Selection drafts."
            ),
            style="Sub.TLabel",
        ).grid(row=1, column=0, sticky="w")

        self.context_var = tk.StringVar(value="No VSDX selected — choose a file to establish context.")
        self.context_bar = ttk.Label(header, textvariable=self.context_var, style="Sub.TLabel", anchor="w")
        self.context_bar.grid(row=2, column=0, columnspan=2, sticky="ew", pady=(7, 0))

        self.status = ttk.Label(
            header, text="Ready", style="Kpi.TLabel"
        )
        self.status.grid(row=0, column=1, rowspan=2, sticky="e")

        self.build_wizard_rail()

        self.body = ttk.Panedwindow(self.root, orient=tk.HORIZONTAL)
        self.body.grid(
            row=2,
            column=0,
            sticky="nsew",
            padx=12,
            pady=(0, 8),
        )

        self.left = ttk.Frame(self.body, padding=4)
        self.center = ttk.Frame(self.body, padding=4)
        self.right = ttk.Frame(self.body, padding=4)

        self.body.add(self.left, weight=3)
        self.body.add(self.center, weight=6)
        self.body.add(self.right, weight=3)

        self.build_left()
        self.build_center()
        self.build_right()
        self.build_footer()
        self.show_stage(0)

    def build_wizard_rail(self):
        """A manufacturing-style process rail; only its current work area is shown."""
        rail = ttk.Frame(self.root, padding=(16, 0, 16, 8))
        rail.grid(row=1, column=0, sticky="ew")
        rail.columnconfigure(9, weight=1)
        self.stage_var = tk.StringVar(value="Step 0 of 7 — Data sources")
        self.stage_help_var = tk.StringVar(value="Select and confirm the workbook tabs used by the routing job.")
        self.stage_names = [
            "0  Data sources", "1  Load VSM", "2  Assign context", "3  Inspect process",
            "4  Generate rule", "5  Review rule", "6  Material selection", "7  Export",
        ]
        self.stage_buttons = []
        for index, label in enumerate(self.stage_names):
            button = ttk.Button(rail, text=label, command=lambda i=index: self.show_stage(i))
            button.grid(row=0, column=index, padx=(0, 5), sticky="w")
            self.stage_buttons.append(button)
        self.advanced_btn = ttk.Button(rail, text="Advanced workspace", command=self.show_advanced)
        self.advanced_btn.grid(row=0, column=8, padx=(12, 0), sticky="e")
        ttk.Label(rail, textvariable=self.stage_var, style="Kpi.TLabel").grid(row=1, column=0, columnspan=3, sticky="w", pady=(5, 0))
        ttk.Label(rail, textvariable=self.stage_help_var, style="Sub.TLabel").grid(row=1, column=3, columnspan=6, sticky="e", pady=(5, 0))

    def build_left(self):
        self.left.columnconfigure(0, weight=1)
        self.left.rowconfigure(2, weight=1)

        intake_card = ttk.Labelframe(
            self.left, text="Step 0. Confirm business data sources", style="Card.TLabelframe",
        )
        intake_card.grid(row=0, column=0, sticky="nsew", pady=(0, 8))
        intake_card.columnconfigure(1, weight=1)
        self.source_intake_card = intake_card
        ttk.Label(
            intake_card,
            text="Browse each workbook, verify the proposed tab names, then click Confirm & Load. VSM upload unlocks only when all three sources pass.",
            style="Sub.TLabel", wraplength=1120,
        ).grid(row=0, column=0, columnspan=4, sticky="w", pady=(0, 12))

        self.p41_path_var = tk.StringVar()
        self.p41_material_sheet = tk.StringVar(value="MATERIAL_IN_SCOPE")
        self.p41_mapl_sheet = tk.StringVar(value="MAPL")
        self.p41_plpo_sheet = tk.StringVar(value="PLPO")
        self.p41_source_info = tk.StringVar(value="Not confirmed")
        ttk.Label(intake_card, text="1. P41 Routing extraction", font=("Segoe UI", 10, "bold")).grid(row=1, column=0, sticky="w")
        ttk.Entry(intake_card, textvariable=self.p41_path_var).grid(row=1, column=1, sticky="ew", padx=8)
        ttk.Button(intake_card, text="Browse…", command=lambda: self.browse_step0_source("p41")).grid(row=1, column=2)
        ttk.Button(intake_card, text="Confirm & Load", command=self.confirm_p41_source, style="Primary.TButton").grid(row=1, column=3, padx=(8, 0))
        p41_sheets = ttk.Frame(intake_card)
        p41_sheets.grid(row=2, column=1, columnspan=3, sticky="ew", padx=8, pady=(6, 3))
        for index, (label, variable) in enumerate((("Material", self.p41_material_sheet), ("Routing assignment", self.p41_mapl_sheet), ("Operations", self.p41_plpo_sheet))):
            ttk.Label(p41_sheets, text=f"{label} tab:").grid(row=0, column=index * 2, sticky="w", padx=(0 if index == 0 else 12, 4))
            combo = ttk.Combobox(p41_sheets, textvariable=variable, state="readonly", width=22)
            combo.grid(row=0, column=index * 2 + 1, sticky="w")
            setattr(self, f"p41_sheet_combo_{index}", combo)
        ttk.Label(intake_card, textvariable=self.p41_source_info, style="Sub.TLabel").grid(row=3, column=1, columnspan=3, sticky="w", padx=8)
        ttk.Separator(intake_card).grid(row=4, column=0, columnspan=4, sticky="ew", pady=12)

        self.workcenter_path_var = tk.StringVar()
        self.workcenter_sheet_var = tk.StringVar(value="PTS03_Work_Center")
        self.workcenter_source_info = tk.StringVar(value="Not confirmed")
        ttk.Label(intake_card, text="2. PTS03 WorkCenter mapping", font=("Segoe UI", 10, "bold")).grid(row=5, column=0, sticky="w")
        ttk.Entry(intake_card, textvariable=self.workcenter_path_var).grid(row=5, column=1, sticky="ew", padx=8)
        ttk.Button(intake_card, text="Browse…", command=lambda: self.browse_step0_source("workcenter")).grid(row=5, column=2)
        ttk.Button(intake_card, text="Confirm & Load", command=self.confirm_workcenter_source, style="Primary.TButton").grid(row=5, column=3, padx=(8, 0))
        wc_line = ttk.Frame(intake_card)
        wc_line.grid(row=6, column=1, columnspan=3, sticky="w", padx=8, pady=(6, 3))
        ttk.Label(wc_line, text="WorkCenter tab:").pack(side=tk.LEFT, padx=(0, 4))
        self.workcenter_sheet_combo = ttk.Combobox(wc_line, textvariable=self.workcenter_sheet_var, state="readonly", width=34)
        self.workcenter_sheet_combo.pack(side=tk.LEFT)
        ttk.Label(intake_card, textvariable=self.workcenter_source_info, style="Sub.TLabel").grid(row=7, column=1, columnspan=3, sticky="w", padx=8)
        ttk.Separator(intake_card).grid(row=8, column=0, columnspan=4, sticky="ew", pady=12)

        self.code_path_var = tk.StringVar()
        self.mrp_sheet_var = tk.StringVar(value="Final list of MRP Controllers")
        self.supervisor_sheet_var = tk.StringVar(value="Production Supervisor")
        self.code_source_info = tk.StringVar(value="Not confirmed")
        ttk.Label(intake_card, text="3. MRP / Production Supervisor reference", font=("Segoe UI", 10, "bold")).grid(row=9, column=0, sticky="w")
        ttk.Entry(intake_card, textvariable=self.code_path_var).grid(row=9, column=1, sticky="ew", padx=8)
        ttk.Button(intake_card, text="Browse…", command=lambda: self.browse_step0_source("codes")).grid(row=9, column=2)
        ttk.Button(intake_card, text="Confirm & Load", command=self.confirm_code_source, style="Primary.TButton").grid(row=9, column=3, padx=(8, 0))
        code_sheets = ttk.Frame(intake_card)
        code_sheets.grid(row=10, column=1, columnspan=3, sticky="w", padx=8, pady=(6, 3))
        ttk.Label(code_sheets, text="MRP tab:").pack(side=tk.LEFT, padx=(0, 4))
        self.mrp_sheet_combo = ttk.Combobox(code_sheets, textvariable=self.mrp_sheet_var, state="readonly", width=32)
        self.mrp_sheet_combo.pack(side=tk.LEFT)
        ttk.Label(code_sheets, text="Supervisor tab:").pack(side=tk.LEFT, padx=(16, 4))
        self.supervisor_sheet_combo = ttk.Combobox(code_sheets, textvariable=self.supervisor_sheet_var, state="readonly", width=30)
        self.supervisor_sheet_combo.pack(side=tk.LEFT)
        ttk.Label(intake_card, textvariable=self.code_source_info, style="Sub.TLabel").grid(row=11, column=1, columnspan=3, sticky="w", padx=8)
        self.step0_summary = tk.StringVar(value="0 of 3 sources confirmed")
        ttk.Label(intake_card, textvariable=self.step0_summary, style="Kpi.TLabel").grid(row=12, column=0, columnspan=4, sticky="w", pady=(14, 0))

        source_card = ttk.Labelframe(
            self.left,
            text="Step 2 support. Department / CRID mapping",
            style="Card.TLabelframe",
        )
        source_card.grid(row=1, column=0, sticky="ew", pady=(0, 8))
        source_card.columnconfigure(0, weight=1)
        # Reserve vertical working room for the VSDX file table below.
        source_card.configure(height=180)
        source_card.grid_propagate(False)
        self.source_card = source_card

        source_tabs = ttk.Notebook(source_card)
        source_tabs.grid(row=0, column=0, sticky="ew")
        self.source_tabs = source_tabs

        excel_tab = ttk.Frame(source_tabs, padding=8)
        manual_tab = ttk.Frame(source_tabs, padding=8)
        source_tabs.add(excel_tab, text="Excel workcentre mapping")
        source_tabs.add(manual_tab, text="Department | Plant list")

        excel_tab.columnconfigure(0, weight=1)
        self.map_path = tk.StringVar()
        ttk.Entry(excel_tab, textvariable=self.map_path).grid(
            row=0, column=0, sticky="ew", padx=(0, 6)
        )
        ttk.Button(
            excel_tab,
            text="Browse…",
            command=self.browse_mapping,
        ).grid(row=0, column=1)

        self.map_info = tk.StringVar(
            value="Loading bundled mapping workbook…"
        )
        ttk.Label(
            excel_tab,
            textvariable=self.map_info,
            style="Sub.TLabel",
            wraplength=410,
        ).grid(
            row=1,
            column=0,
            columnspan=2,
            sticky="w",
            pady=(7, 0),
        )

        manual_tab.columnconfigure(0, weight=1)
        manual_tab.rowconfigure(1, weight=1)

        ttk.Label(
            manual_tab,
            text=(
                "Paste two columns. Accepted separators: |, tab, comma, "
                "semicolon, or multiple spaces. Header is optional."
            ),
            style="Sub.TLabel",
            wraplength=410,
        ).grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 5))

        list_frame = ttk.Frame(manual_tab)
        list_frame.grid(row=1, column=0, columnspan=2, sticky="nsew")
        list_frame.columnconfigure(0, weight=1)
        list_frame.rowconfigure(0, weight=1)

        self.department_list_text = tk.Text(
            list_frame,
            # Keep the source editor compact in the guided workspace; the
            # scrollbars preserve full editing while leaving working height
            # for the VSDX / Plant assignment list below.
            height=2,
            wrap="none",
            font=("Consolas", 9),
            relief=tk.FLAT,
            background="#F8FAFC",
            undo=True,
        )
        list_y = ttk.Scrollbar(
            list_frame,
            orient=tk.VERTICAL,
            command=self.department_list_text.yview,
        )
        list_x = ttk.Scrollbar(
            list_frame,
            orient=tk.HORIZONTAL,
            command=self.department_list_text.xview,
        )
        self.department_list_text.configure(
            yscrollcommand=list_y.set,
            xscrollcommand=list_x.set,
        )
        self.department_list_text.grid(
            row=0, column=0, sticky="nsew"
        )
        list_y.grid(row=0, column=1, sticky="ns")
        list_x.grid(row=1, column=0, sticky="ew")
        self.department_list_text.insert(
            "1.0", "DEPARTMENT | PLANT\n"
        )

        manual_buttons = ttk.Frame(manual_tab)
        manual_buttons.grid(
            row=2,
            column=0,
            columnspan=2,
            sticky="ew",
            pady=(7, 0),
        )
        ttk.Button(
            manual_buttons,
            text="Load / Apply list",
            command=self.apply_department_plant_list,
        ).pack(side=tk.LEFT, padx=(0, 4))
        ttk.Button(
            manual_buttons,
            text="Paste & load",
            command=self.paste_department_plant_list,
        ).pack(side=tk.LEFT, padx=(0, 4))
        ttk.Button(
            manual_buttons,
            text="Import TXT/CSV",
            command=self.import_department_plant_list,
        ).pack(side=tk.LEFT, padx=(0, 4))
        ttk.Button(
            manual_buttons,
            text="Clear",
            command=self.clear_department_plant_list,
        ).pack(side=tk.LEFT)

        self.manual_info = tk.StringVar(
            value="No Department-Plant list loaded."
        )
        ttk.Label(
            manual_tab,
            textvariable=self.manual_info,
            style="Sub.TLabel",
            wraplength=410,
        ).grid(
            row=3,
            column=0,
            columnspan=2,
            sticky="w",
            pady=(6, 0),
        )

        # Step 2 replaces source-file editing with the mapping records that
        # actually guide the Plant/CRID assignment decision.
        self.context_mapping_frame = ttk.Frame(source_card, padding=8)
        self.context_mapping_frame.columnconfigure(0, weight=1)
        self.context_mapping_frame.rowconfigure(2, weight=1)
        self.context_mapping_title = tk.StringVar(value="Available mapping")
        ttk.Label(self.context_mapping_frame, textvariable=self.context_mapping_title, font=("Segoe UI", 9, "bold")).grid(row=0, column=0, sticky="w")
        self.context_mapping_filter = tk.StringVar()
        mapping_filter = ttk.Entry(self.context_mapping_frame, textvariable=self.context_mapping_filter)
        mapping_filter.grid(row=1, column=0, sticky="ew", pady=(5, 5))
        mapping_filter.insert(0, "Filter CRID or workcenter")
        mapping_filter.bind("<FocusIn>", self.clear_mapping_filter_hint)
        mapping_filter.bind("<KeyRelease>", lambda _event: self.refresh_context_mapping())
        map_table = ttk.Frame(self.context_mapping_frame)
        map_table.grid(row=2, column=0, sticky="nsew")
        map_table.columnconfigure(0, weight=1)
        map_table.rowconfigure(0, weight=1)
        self.context_mapping_tree = ttk.Treeview(map_table, columns=("crid", "wc", "source"), show="headings", height=10)
        for column, title, width in [("crid", "CRID", 105), ("wc", "WORKCENTER", 105), ("source", "SOURCE", 120)]:
            self.context_mapping_tree.heading(column, text=title)
            self.context_mapping_tree.column(column, width=width, anchor="w")
        mapping_y = ttk.Scrollbar(map_table, orient=tk.VERTICAL, command=self.context_mapping_tree.yview)
        self.context_mapping_tree.configure(yscrollcommand=mapping_y.set)
        self.context_mapping_tree.grid(row=0, column=0, sticky="nsew")
        mapping_y.grid(row=0, column=1, sticky="ns")
        self.context_mapping_frame.grid(row=0, column=0, sticky="nsew")
        self.context_mapping_frame.grid_remove()

        files_card = ttk.Labelframe(
            self.left,
            text="VSM files and plant assignment",
            style="Card.TLabelframe",
        )
        files_card.grid(row=2, column=0, sticky="nsew")
        files_card.columnconfigure(0, weight=1)
        files_card.rowconfigure(2, weight=1)

        file_toolbar = ttk.Frame(files_card)
        file_toolbar.grid(row=0, column=0, sticky="ew", pady=(0, 6))
        ttk.Button(
            file_toolbar,
            text="Add VSDX",
            command=self.add_files,
        ).pack(side=tk.LEFT, padx=(0, 4))
        ttk.Button(
            file_toolbar,
            text="Add folder",
            command=self.add_folder,
        ).pack(side=tk.LEFT, padx=(0, 4))
        ttk.Button(
            file_toolbar,
            text="Import package",
            command=self.import_package,
        ).pack(side=tk.LEFT, padx=(0, 4))
        ttk.Button(
            file_toolbar,
            text="Remove",
            command=self.remove_file,
        ).pack(side=tk.LEFT, padx=(0, 4))
        ttk.Button(
            file_toolbar,
            text="Clear",
            command=self.clear_files,
        ).pack(side=tk.LEFT)

        # Plant selection is deliberately placed above the file list so it
        # remains visible even when the lower panel is short.
        plant_frame = ttk.Frame(files_card, padding=(0, 2, 0, 6))
        plant_frame.grid(row=1, column=0, sticky="ew")
        plant_frame.columnconfigure(1, weight=1)
        ttk.Label(
            plant_frame,
            text="Selected file Plant",
            font=("Segoe UI", 9, "bold"),
        ).grid(row=0, column=0, sticky="w", padx=(0, 8))

        self.plant = tk.StringVar()
        self.plant_choice_frame = ttk.Frame(plant_frame)
        self.plant_choice_frame.grid(row=0, column=1, sticky="ew", padx=(0, 6))
        self.plant_buttons = {}
        self.plant_apply_btn = ttk.Button(
            plant_frame,
            text="Apply",
            command=self.apply_plant,
            state=tk.DISABLED,
        )
        self.plant_apply_btn.grid(row=0, column=2)
        self.plant_state = tk.StringVar(value="Select a VSDX file")
        ttk.Label(plant_frame, textvariable=self.plant_state, style="Sub.TLabel").grid(row=1, column=1, columnspan=2, sticky="w", pady=(3, 0))
        self.refresh_plant_buttons(["CA02", "CA06", "CA15"])

        self.file_tree = ttk.Treeview(
            files_card,
            columns=("file", "plant", "pages", "status"),
            show="headings",
            selectmode="browse",
            height=7,
        )
        file_specs = [
            ("file", "VSDX file", 220, "w"),
            ("plant", "Plant", 70, "center"),
            ("pages", "Tabs", 45, "center"),
            ("status", "Status", 80, "center"),
        ]
        for column, title, width, anchor in file_specs:
            self.file_tree.heading(column, text=title)
            self.file_tree.column(column, width=width, anchor=anchor)

        file_y = ttk.Scrollbar(
            files_card,
            orient=tk.VERTICAL,
            command=self.file_tree.yview,
        )
        self.file_tree.configure(yscrollcommand=file_y.set)
        self.file_tree.grid(row=2, column=0, sticky="nsew")
        file_y.grid(row=2, column=1, sticky="ns")
        self.file_tree.bind("<<TreeviewSelect>>", self.file_selected)
        self.file_tree.bind("<Double-1>", self.file_double)

        action_card = ttk.Labelframe(
            self.left,
            text="Step 7. Analysis package and export",
            style="Card.TLabelframe",
        )
        action_card.grid(row=3, column=0, sticky="ew", pady=(8, 0))
        action_card.columnconfigure(0, weight=1)
        action_card.columnconfigure(1, weight=1)
        self.files_card = files_card
        self.action_card = action_card

        self.save_package_btn = ttk.Button(
            action_card,
            text="Save analysis package",
            command=self.save_package,
        )
        self.save_package_btn.grid(
            row=0, column=0, sticky="ew", padx=(0, 3), pady=(0, 5)
        )
        ttk.Button(
            action_card,
            text="Import analysis package",
            command=self.import_package,
        ).grid(row=0, column=1, sticky="ew", padx=(3, 0), pady=(0, 5))

        self.run_btn = ttk.Button(
            action_card,
            text="Run metadata analysis",
            command=self.run_analysis,
            style="Primary.TButton",
            state=tk.DISABLED,
        )
        self.run_btn.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(0, 5))

        self.export_btn = ttk.Button(
            action_card,
            text="Export structured Excel",
            command=self.export_results,
            state=tk.DISABLED,
        )
        self.export_btn.grid(row=2, column=0, columnspan=2, sticky="ew")

    def build_center(self):
        self.center.columnconfigure(0, weight=1)
        self.center.rowconfigure(1, weight=1)

        page_card = ttk.Labelframe(
            self.center, text="VSM tabs / pages", style="Card.TLabelframe"
        )
        self.page_card = page_card
        page_card.grid(row=0, column=0, sticky="ew", pady=(0, 8))
        page_card.columnconfigure(0, weight=1)
        page_card.rowconfigure(1, weight=1)

        columns = ("tab", "shapes", "wcs", "suggested", "confidence", "assigned", "status", "clear_assignment")
        self.page_tree = ttk.Treeview(page_card, columns=columns, show="headings", height=8, selectmode="browse")
        page_specs = [
            ("tab", "Tab", 135, "w"), ("shapes", "Shp", 42, "center"),
            ("wcs", "WC", 32, "center"), ("suggested", "Suggested", 75, "w"),
            ("confidence", "Confidence", 70, "center"), ("assigned", "Assigned", 95, "w"),
            ("status", "Status", 105, "center"),
            ("clear_assignment", "✕", 28, "center"),
        ]
        for column, title, width, anchor in page_specs:
            self.page_tree.heading(column, text=title)
            self.page_tree.column(column, width=width, anchor=anchor)
        page_toolbar = ttk.Frame(page_card)
        page_toolbar.grid(row=0, column=0, sticky="ew", pady=(0, 5))
        page_toolbar.columnconfigure(0, weight=1)
        ttk.Label(
            page_toolbar,
            text="Choose a VSM tab, then select an action.",
            style="Sub.TLabel",
        ).grid(row=0, column=0, sticky="w")
        self.regenerate_all_btn = ttk.Button(page_toolbar, text="Regenerate All", command=self.regenerate_all_vsms, style="Primary.TButton", state=tk.DISABLED)
        self.regenerate_all_btn.grid(row=0, column=2, sticky="e")
        self.selected_tab_btn = ttk.Button(page_toolbar, text="Selected Tab", command=self.analyze_selected_vsm, style="Secondary.TButton", state=tk.DISABLED)
        self.selected_tab_btn.grid(row=0, column=1, sticky="e", padx=(0, 8))
        ToolTip(self.selected_tab_btn, "Refreshes the selected VSM tab only: its shape metadata, detected workcentres and CRID mapping evidence are reloaded without re-running every tab.")
        ToolTip(self.regenerate_all_btn, "Runs the established full metadata analysis for every tab in every loaded VSDX file.")
        self.page_tree.grid(row=1, column=0, sticky="ew")
        self.page_tree.bind("<<TreeviewSelect>>", self.page_selected)
        self.page_tree.bind("<ButtonRelease-1>", self.page_click)
        self.page_tree.bind("<Double-1>", self.page_double)
        ToolTip(self.page_tree, "Status shows whether a tab is ready to generate, needs review, or is still loading.")

        self.center_notebook = ttk.Notebook(self.center)
        self.center_notebook.grid(row=1, column=0, sticky="nsew")

        preview_card = ttk.Frame(self.center_notebook, padding=6)
        routing_card = ttk.Frame(self.center_notebook, padding=6)
        material_card = ttk.Frame(self.center_notebook, padding=6)
        self.center_notebook.add(preview_card, text="Visual Preview")
        self.center_notebook.add(routing_card, text="Change Rule Draft")
        self.center_notebook.add(material_card, text="Material Selection")

        preview_card.columnconfigure(0, weight=1)
        preview_card.rowconfigure(1, weight=1)
        preview_toolbar = ttk.Frame(preview_card)
        preview_toolbar.grid(row=0, column=0, sticky="ew", pady=(0, 6))
        ttk.Button(preview_toolbar, text="Fit", command=lambda: self.set_zoom(1)).pack(side=tk.LEFT, padx=(0, 4))
        self.zoom_var = tk.StringVar(value="100%")
        self.zoom_combo = ttk.Combobox(preview_toolbar, textvariable=self.zoom_var, values=["50%", "75%", "100%", "125%", "150%", "200%"], width=7, state="readonly")
        self.zoom_combo.pack(side=tk.LEFT, padx=(0, 4))
        self.zoom_combo.bind("<<ComboboxSelected>>", self.zoom_selected)
        ttk.Button(preview_toolbar, text="Zoom +", command=lambda: self.set_zoom(min(3, self.zoom * 1.2))).pack(side=tk.LEFT, padx=(0, 4))
        ttk.Button(preview_toolbar, text="Zoom −", command=lambda: self.set_zoom(max(0.45, self.zoom / 1.2))).pack(side=tk.LEFT, padx=(0, 4))
        ttk.Button(preview_toolbar, text="Reload visual", command=self.reload_preview).pack(side=tk.LEFT, padx=(8, 4))
        ttk.Checkbutton(preview_toolbar, text="Auto preview", variable=self.auto_preview, command=self.auto_preview_changed).pack(side=tk.LEFT, padx=(4, 4))
        self.exact_btn = ttk.Button(preview_toolbar, text="Exact Visio", command=self.exact_preview)
        self.exact_btn.pack(side=tk.LEFT, padx=(8, 4))
        ToolTip(self.exact_btn, "Exact Visio Preview uses an installed Visio renderer when available and never changes routing data.")
        if not VisioRenderer.supported():
            self.exact_btn.configure(state=tk.DISABLED)
        ttk.Checkbutton(preview_toolbar, text="Highlight detected Workcenters", variable=self.highlight_workcenters, command=self.render_preview).pack(side=tk.LEFT, padx=(8, 2))
        ttk.Checkbutton(preview_toolbar, text="Show process path", variable=self.show_process_path, command=self.render_preview).pack(side=tk.LEFT, padx=(2, 2))
        ttk.Checkbutton(preview_toolbar, text="Shape IDs", variable=self.show_shape_ids, command=self.render_preview).pack(side=tk.LEFT, padx=(2, 4))
        self.preview_mode = tk.StringVar(value="Select a tab to preview.")
        ttk.Label(preview_toolbar, textvariable=self.preview_mode, style="Sub.TLabel").pack(side=tk.RIGHT)

        canvas_frame = ttk.Frame(preview_card)
        canvas_frame.grid(row=1, column=0, sticky="nsew")
        canvas_frame.columnconfigure(0, weight=1)
        canvas_frame.rowconfigure(0, weight=1)
        self.canvas = tk.Canvas(canvas_frame, background="#E2E8F0", highlightthickness=0)
        canvas_x = ttk.Scrollbar(canvas_frame, orient=tk.HORIZONTAL, command=self.canvas.xview)
        canvas_y = ttk.Scrollbar(canvas_frame, orient=tk.VERTICAL, command=self.canvas.yview)
        self.canvas.configure(xscrollcommand=canvas_x.set, yscrollcommand=canvas_y.set)
        self.canvas.grid(row=0, column=0, sticky="nsew")
        canvas_x.grid(row=1, column=0, sticky="ew")
        canvas_y.grid(row=0, column=1, sticky="ns")
        self.canvas.bind("<Configure>", lambda _event: self.render_preview())

        # Routing-generation preparation panel.
        routing_card.columnconfigure(0, weight=1)
        routing_card.rowconfigure(4, weight=1)
        source = ttk.Labelframe(routing_card, text="Routing template / OLD WC reference", style="Card.TLabelframe")
        source.grid(row=0, column=0, sticky="ew", pady=(0, 6))
        source.columnconfigure(1, weight=1)
        ttk.Label(source, text="Template").grid(row=0, column=0, sticky="w", padx=(0, 6))
        self.routing_template_var = tk.StringVar()
        ttk.Entry(source, textvariable=self.routing_template_var).grid(row=0, column=1, sticky="ew", padx=(0, 6))
        ttk.Button(source, text="Browse…", command=self.browse_routing_template).grid(row=0, column=2)
        self.routing_reference_info = tk.StringVar(value="No routing reference loaded.")
        ttk.Label(source, textvariable=self.routing_reference_info, style="Sub.TLabel").grid(row=1, column=0, columnspan=3, sticky="w", pady=(4, 0))

        toolbar = ttk.Frame(routing_card)
        toolbar.grid(row=1, column=0, sticky="ew", pady=(0, 4))
        self.generate_selected_btn = ttk.Button(toolbar, text="Generate Selected CRID", command=self.generate_change_rule, style="Primary.TButton")
        self.generate_selected_btn.pack(side=tk.LEFT, padx=(0, 4))
        self.generate_plant_btn = ttk.Button(toolbar, text="Generate All CRIDs for Plant", command=self.generate_all_crids_for_plant, style="Primary.TButton")
        self.generate_plant_btn.pack(side=tk.LEFT, padx=(0, 4))
        ttk.Button(toolbar, text="Generate All Loaded Plants", command=self.generate_all_loaded_plants).pack(side=tk.LEFT, padx=(0, 4))
        self.validate_btn = ttk.Button(toolbar, text="Validate", command=self.validate_change_rule)
        self.validate_btn.pack(side=tk.LEFT, padx=(8, 4))
        self.clear_draft_btn = ttk.Button(toolbar, text="Clear Draft", command=self.clear_change_rule)
        self.clear_draft_btn.pack(side=tk.LEFT, padx=(0, 4))
        self.routing_summary = tk.StringVar(value="Select a tab and generate a routing draft.")
        ttk.Label(toolbar, textvariable=self.routing_summary, style="Kpi.TLabel").pack(side=tk.RIGHT)

        batchbar = ttk.Frame(routing_card)
        batchbar.grid(row=2, column=0, sticky="ew", pady=(0, 6))
        ttk.Label(batchbar, text="View:").pack(side=tk.LEFT)
        self.routing_filter_var = tk.StringVar(value="ALL DEPARTMENTS")
        self.routing_filter_combo = ttk.Combobox(batchbar, textvariable=self.routing_filter_var, state="readonly", width=22)
        self.routing_filter_combo.pack(side=tk.LEFT, padx=(4, 10))
        self.routing_filter_combo.bind("<<ComboboxSelected>>", lambda _e: self.refresh_routing_tree())
        ttk.Label(batchbar, text="Status:").pack(side=tk.LEFT, padx=(6, 0))
        self.routing_status_filter = tk.StringVar(value="All")
        self.routing_status_combo = ttk.Combobox(batchbar, textvariable=self.routing_status_filter, values=["All", "Ready", "Review Required"], state="readonly", width=16)
        self.routing_status_combo.pack(side=tk.LEFT, padx=(4, 8))
        self.routing_status_combo.bind("<<ComboboxSelected>>", lambda _e: self.refresh_routing_tree())
        self.routing_search = tk.StringVar()
        search_entry = ttk.Entry(batchbar, textvariable=self.routing_search, width=22)
        search_entry.pack(side=tk.LEFT, padx=(0, 8))
        search_entry.bind("<KeyRelease>", lambda _e: self.refresh_routing_tree())
        ttk.Label(batchbar, text="Rerun policy:").pack(side=tk.LEFT)
        self.rerun_policy_var = tk.StringVar(value="Preview changes first")
        ttk.Combobox(
            batchbar, textvariable=self.rerun_policy_var, state="readonly", width=24,
            values=["Preview changes first", "Replace existing draft", "Merge and preserve manual edits"],
        ).pack(side=tk.LEFT, padx=(4, 10))
        self.rerun_btn = ttk.Menubutton(batchbar, text="Rerun ▼")
        rerun_menu = tk.Menu(self.rerun_btn, tearoff=False)
        rerun_menu.add_command(label="Rerun Selected CRID", command=self.rerun_selected_crid)
        rerun_menu.add_command(label="Rerun Selected Plant", command=self.rerun_selected_plant)
        rerun_menu.add_command(label="Rerun Everything", command=self.rerun_everything)
        self.rerun_btn.configure(menu=rerun_menu)
        self.rerun_btn.pack(side=tk.LEFT, padx=(0, 8))
        ToolTip(self.rerun_btn, "Preview is safest; Merge keeps manual edits; Replace rebuilds the selected scope.")

        # Keep export actions on their own row. They must remain reachable in
        # the narrower Assign Context and Advanced workspace layouts.
        exportbar = ttk.Frame(routing_card)
        exportbar.grid(row=3, column=0, sticky="ew", pady=(0, 6))
        ttk.Label(exportbar, text="Export routing:").pack(side=tk.LEFT, padx=(0, 8))
        self.export_selected_btn = ttk.Button(exportbar, text="Export Change Rule", command=self.export_change_rule, style="Primary.TButton")
        self.export_selected_btn.pack(side=tk.LEFT, padx=(0, 4))
        ttk.Button(exportbar, text="Export Plant Package", command=self.export_plant_routing_package).pack(side=tk.LEFT, padx=(0, 4))
        ttk.Button(exportbar, text="Export All Plants", command=self.export_all_routing_package).pack(side=tk.LEFT, padx=(0, 4))
        ttk.Button(exportbar, text="Export into Template", command=self.export_change_rule_template).pack(side=tk.LEFT, padx=(0, 4))

        table_frame = ttk.Frame(routing_card)
        table_frame.grid(row=4, column=0, sticky="nsew")
        table_frame.columnconfigure(0, weight=1)
        table_frame.rowconfigure(0, weight=1)
        cols = ("crid","plant","alternate","sequence","subseq","old_wc","new_wc","description","rule_type","column","status","notes")
        self.routing_tree = ttk.Treeview(table_frame, columns=cols, show="headings", selectmode="browse", style="Draft.Treeview")
        specs = [
            ("crid","CRID",90),("plant","PLANT",55),("alternate","ALT",45),("sequence","SEQ",45),("subseq","SUB",45),
            ("old_wc","OLD WC",90),("new_wc","NEW WC",90),("description","OP_DESCRIPTIONS",180),("rule_type","RULE TYPE",70),
            ("column","COLUMN",65),("status","STATUS",70),("notes","REVIEW NOTES",200),
        ]
        for col, title, width in specs:
            self.routing_tree.heading(col, text=title)
            self.routing_tree.column(col, width=width, anchor="w" if col in {"description","notes","crid","old_wc","new_wc"} else "center")
        rx = ttk.Scrollbar(table_frame, orient=tk.HORIZONTAL, command=self.routing_tree.xview)
        ry = ttk.Scrollbar(table_frame, orient=tk.VERTICAL, command=self.routing_tree.yview)
        self.routing_tree.configure(xscrollcommand=rx.set, yscrollcommand=ry.set)
        self.routing_tree.grid(row=0, column=0, sticky="nsew")
        rx.grid(row=1, column=0, sticky="ew")
        ry.grid(row=0, column=1, sticky="ns")
        self.routing_tree.bind("<Double-1>", self.routing_double_click)
        for col in cols:
            self.routing_tree.heading(col, command=lambda c=col: self.sort_routing_tree(c, False))
        self.routing_row_map = {}

        readiness = ttk.Labelframe(routing_card, text="Department / CRID readiness", style="Card.TLabelframe")
        readiness.grid(row=5, column=0, sticky="ew", pady=(6, 0))
        readiness.columnconfigure(0, weight=1)
        self.readiness_tree = ttk.Treeview(
            readiness, columns=("plant","crid","rows","ready","review","missing","readiness","status","generated"),
            show="headings", height=2,
        )
        for col, title, width in [
            ("plant","PLANT",60),("crid","CRID",130),("rows","ROWS",50),("ready","READY",55),
            ("review","REVIEW",55),("missing","MISSING OLD WC",100),("readiness","READINESS",75),("status","STATUS",105),("generated","LAST GENERATED",135),
        ]:
            self.readiness_tree.heading(col, text=title)
            self.readiness_tree.column(col, width=width, anchor="center" if col != "crid" else "w")
        self.readiness_tree.grid(row=0, column=0, sticky="ew")
        self.readiness_tree.bind("<<TreeviewSelect>>", self.readiness_selected)

        # Step 6 consumes the raw P41 and code reference already confirmed in Step 0.
        material_card.columnconfigure(0, weight=1)
        material_card.rowconfigure(3, weight=1)
        material_source = ttk.Labelframe(material_card, text="Confirmed Step 0 sources", style="Card.TLabelframe")
        material_source.grid(row=0, column=0, sticky="ew", pady=(0, 6))
        material_source.columnconfigure(1, weight=1)
        self.material_source_info = tk.StringVar(value="Complete Step 0 before generating Material Selection.")
        ttk.Label(material_source, textvariable=self.material_source_info, style="Sub.TLabel", wraplength=1100).grid(row=0, column=0, sticky="w")

        material_actions = ttk.Frame(material_card)
        material_actions.grid(row=1, column=0, sticky="ew", pady=(0, 6))
        self.generate_material_btn = ttk.Button(material_actions, text="Generate Material Selection", command=self.generate_material_selection, style="Primary.TButton")
        self.generate_material_btn.pack(side=tk.LEFT, padx=(0, 5))
        self.export_material_btn = ttk.Button(material_actions, text="Export into Material Template", command=self.export_material_selection, state=tk.DISABLED)
        self.export_material_btn.pack(side=tk.LEFT)
        self.material_summary = tk.StringVar(value="No material-selection rows generated.")
        ttk.Label(material_actions, textvariable=self.material_summary, style="Kpi.TLabel").pack(side=tk.RIGHT)

        material_notice = ttk.Label(material_card, text="First round: P41 variants are matched to OLD WC evidence. CA02 Calendar uses the agreed majority-path assumption and remains Review Required.", style="Sub.TLabel")
        material_notice.grid(row=2, column=0, sticky="w", pady=(0, 4))
        columns = ("material", "plant", "group", "counter", "variant", "mrp", "supervisor", "crid", "alternate", "seq1", "seq2", "seq3", "seq4", "status", "notes")
        self.material_tree = ttk.Treeview(material_card, columns=columns, show="headings", style="Draft.Treeview")
        for col, label, width in [
            ("material", "MATERIAL", 120), ("plant", "PLANT", 60), ("group", "ROUTING GROUP", 110), ("counter", "COUNTER", 70),
            ("variant", "VARIANT AS IS", 190), ("mrp", "MRP", 65), ("supervisor", "PROD SUPERVISOR", 105), ("crid", "CRID", 105), ("alternate", "ALT", 55),
            ("seq1", "SEQ1", 85), ("seq2", "SEQ2", 85), ("seq3", "SEQ3", 85), ("seq4", "SEQ4", 85), ("status", "STATUS", 105), ("notes", "REVIEW NOTES", 330),
        ]:
            self.material_tree.heading(col, text=label)
            self.material_tree.column(col, width=width, anchor="w" if col in {"material", "variant", "crid", "notes"} else "center")
        my = ttk.Scrollbar(material_card, orient=tk.VERTICAL, command=self.material_tree.yview)
        mx = ttk.Scrollbar(material_card, orient=tk.HORIZONTAL, command=self.material_tree.xview)
        self.material_tree.configure(yscrollcommand=my.set, xscrollcommand=mx.set)
        self.material_tree.grid(row=3, column=0, sticky="nsew")
        my.grid(row=3, column=1, sticky="ns")
        mx.grid(row=4, column=0, sticky="ew")

    def build_right(self):
        self.right.columnconfigure(0, weight=1)
        self.right.rowconfigure(2, weight=1)
        self.right.rowconfigure(3, weight=1)

        assignment_card = ttk.Labelframe(
            self.right,
            text="Selected tab assignment",
            style="Card.TLabelframe",
        )
        assignment_card.grid(
            row=0, column=0, sticky="ew", pady=(0, 8)
        )
        assignment_card.columnconfigure(0, weight=1)

        self.department = tk.StringVar()
        self.department_combo = ttk.Combobox(
            assignment_card,
            textvariable=self.department,
            state="normal",
        )
        self.department_combo.grid(
            row=0, column=0, sticky="ew", padx=(0, 6)
        )
        ttk.Button(
            assignment_card,
            text="Apply",
            command=self.apply_department,
        ).grid(row=0, column=1, padx=(0, 4))
        ttk.Button(
            assignment_card,
            text="✕ Clear",
            command=self.clear_selected_department,
        ).grid(row=0, column=2)

        metadata_card = ttk.Labelframe(
            self.right,
            text="Metadata quick glance",
            style="Card.TLabelframe",
        )
        metadata_card.grid(
            row=1, column=0, sticky="ew", pady=(0, 8)
        )
        metadata_card.columnconfigure(0, weight=1)

        self.meta = tk.Text(
            metadata_card,
            height=14,
            wrap="word",
            font=("Consolas", 9),
            relief=tk.FLAT,
            background="#F8FAFC",
        )
        self.meta.grid(row=0, column=0, sticky="ew")
        self.meta.configure(state=tk.DISABLED)

        candidate_card = ttk.Labelframe(
            self.right,
            text="Possible Department / CRID",
            style="Card.TLabelframe",
        )
        candidate_card.grid(
            row=2, column=0, sticky="nsew", pady=(0, 8)
        )
        candidate_card.columnconfigure(0, weight=1)
        candidate_card.rowconfigure(0, weight=1)

        self.candidates = ttk.Treeview(
            candidate_card,
            columns=("dept", "match", "evidence"),
            show="headings",
            height=8,
        )
        self.candidates.heading(
            "dept", text="Department / CRID"
        )
        self.candidates.heading("match", text="Match")
        self.candidates.heading(
            "evidence", text="Workcenters / source"
        )
        self.candidates.column("dept", width=130, anchor="w")
        self.candidates.column(
            "match", width=52, anchor="center"
        )
        self.candidates.column(
            "evidence", width=190, anchor="w"
        )
        self.candidates.grid(row=0, column=0, sticky="nsew")
        self.candidates.bind(
            "<Double-1>", self.use_candidate
        )

        log_card = ttk.Labelframe(
            self.right,
            text="Real-time processing log",
            style="Card.TLabelframe",
        )
        log_card.grid(row=3, column=0, sticky="nsew")
        log_card.columnconfigure(0, weight=1)
        log_card.rowconfigure(0, weight=1)

        self.logbox = tk.Text(
            log_card,
            height=13,
            wrap="word",
            font=("Consolas", 8),
            relief=tk.FLAT,
            background="#0F172A",
            foreground="#E2E8F0",
            insertbackground="white",
        )
        log_y = ttk.Scrollbar(
            log_card, command=self.logbox.yview
        )
        self.logbox.configure(yscrollcommand=log_y.set)
        self.logbox.grid(row=0, column=0, sticky="nsew")
        log_y.grid(row=0, column=1, sticky="ns")

    def build_footer(self):
        footer = ttk.Frame(
            self.root, padding=(16, 5, 16, 10)
        )
        footer.grid(row=3, column=0, sticky="ew")
        footer.columnconfigure(0, weight=1)

        self.progress = ttk.Progressbar(
            footer, mode="determinate", maximum=100
        )
        self.progress.grid(
            row=0, column=0, sticky="ew", padx=(0, 10)
        )

        self.progress_text = tk.StringVar(value="0%")
        ttk.Label(
            footer,
            textvariable=self.progress_text,
            width=7,
        ).grid(row=0, column=1)

        self.elapsed_text = tk.StringVar(value="Elapsed 00:00")
        ttk.Label(
            footer,
            textvariable=self.elapsed_text,
            width=16,
        ).grid(row=0, column=2)

        ttk.Button(footer, text="‹ Previous stage", command=self.previous_stage).grid(row=0, column=3, padx=(18, 4))
        self.next_stage_btn = ttk.Button(footer, text="Next stage ›", command=self.next_stage, style="Primary.TButton")
        self.next_stage_btn.grid(row=0, column=4)
        ttk.Label(
            footer,
            text="Software by COA Data Team Copyright 2026",
            style="Sub.TLabel",
        ).grid(row=1, column=0, columnspan=5, sticky="w", pady=(6, 0))

    def browse_step0_source(self, kind: str):
        labels = {
            "p41": "Select P41 Routing extraction",
            "workcenter": "Select PTS03 WorkCenter mapping",
            "codes": "Select MRP / Production Supervisor reference",
        }
        path = filedialog.askopenfilename(title=labels[kind], filetypes=[("Excel workbook", "*.xlsx *.xlsm"), ("All files", "*.*")])
        if not path:
            return
        try:
            sheets = workbook_sheets(path)
        except Exception as exc:
            messagebox.showerror("Data source", str(exc))
            return
        if kind == "p41":
            self.p41_path_var.set(path)
            for combo in (self.p41_sheet_combo_0, self.p41_sheet_combo_1, self.p41_sheet_combo_2):
                combo.configure(values=sheets)
            self.p41_source_info.set("Workbook opened — confirm the three P41 tabs before data is read.")
            self.material_sources.p41_path = ""
            self.material_sources.variants.clear()
            self.material_sources.workcenters.clear()
        elif kind == "workcenter":
            self.workcenter_path_var.set(path)
            self.workcenter_sheet_combo.configure(values=sheets)
            preferred = next((name for name in sheets if "WORK" in name.upper() and "CENTER" in name.upper()), sheets[0] if sheets else "")
            self.workcenter_sheet_var.set(preferred)
            self.workcenter_source_info.set("Workbook opened — select and confirm the WorkCenter tab.")
            self.routing_source_ready = False
        else:
            self.code_path_var.set(path)
            self.mrp_sheet_combo.configure(values=sheets)
            self.supervisor_sheet_combo.configure(values=sheets)
            if "Final list of MRP Controllers" in sheets:
                self.mrp_sheet_var.set("Final list of MRP Controllers")
            if "Production Supervisor" in sheets:
                self.supervisor_sheet_var.set("Production Supervisor")
            self.code_source_info.set("Workbook opened — select and confirm both code tabs.")
            self.material_sources.code_reference_path = ""
            self.material_sources.valid_mrp_codes.clear()
            self.material_sources.valid_supervisor_codes.clear()
        self.refresh_step0_state()

    def confirm_p41_source(self):
        path = self.p41_path_var.get().strip()
        if not path:
            messagebox.showwarning("P41 source", "Browse to the P41 Routing extraction first.")
            return
        chosen = (self.p41_material_sheet.get(), self.p41_mapl_sheet.get(), self.p41_plpo_sheet.get())
        if not all(chosen):
            messagebox.showwarning("P41 source", "Select the Material, routing-assignment and operations tabs.")
            return
        if self.step0_loading:
            return
        self.step0_loading = True
        self.status.configure(text="Loading confirmed P41 tabs in background…")
        self.p41_source_info.set("Reading the confirmed P41 tabs — the app remains available.")

        def worker():
            try:
                loaded = load_p41_routing(path, *chosen)
                self.events.put(("p41_loaded", loaded, path, chosen))
            except Exception as exc:
                self.events.put(("p41_error", str(exc)))

        threading.Thread(target=worker, daemon=True).start()

    def confirm_workcenter_source(self):
        path, sheet = self.workcenter_path_var.get().strip(), self.workcenter_sheet_var.get().strip()
        if not path or not sheet:
            messagebox.showwarning("WorkCenter source", "Browse the workbook and select its WorkCenter tab first.")
            return
        try:
            reference = load_routing_reference(path, sheet)
        except Exception as exc:
            self.workcenter_source_info.set("Validation failed — correct the file or tab selection.")
            messagebox.showerror("WorkCenter source", str(exc))
            return
        self.routing_reference = reference
        self.routing_source_ready = True
        self.workcenter_source_info.set(f"Ready — {len(reference.records):,} old/new WorkCenter mapping records loaded from {sheet}.")
        self.routing_reference_info.set(f"Step 0 WorkCenter reference: {Path(path).name} / {sheet} ({len(reference.records):,} rows)")
        self.status.configure(text="PTS03 WorkCenter source confirmed")
        self.log(f"Step 0 WorkCenter confirmed: {Path(path).name} | tab: {sheet}")
        self.refresh_step0_state()

    def confirm_code_source(self):
        path = self.code_path_var.get().strip()
        mrp_sheet, supervisor_sheet = self.mrp_sheet_var.get().strip(), self.supervisor_sheet_var.get().strip()
        if not path or not mrp_sheet or not supervisor_sheet:
            messagebox.showwarning("Code reference", "Browse the workbook and select both code tabs first.")
            return
        try:
            self.material_sources = load_code_reference(path, self.material_sources, mrp_sheet, supervisor_sheet)
        except Exception as exc:
            self.code_source_info.set("Validation failed — correct the file or tab selection.")
            messagebox.showerror("Code reference", str(exc))
            return
        self.code_source_info.set(f"Ready — {len(self.material_sources.valid_mrp_codes):,} MRP and {len(self.material_sources.valid_supervisor_codes):,} supervisor Plant/code pairs loaded.")
        self.status.configure(text="MRP and Production Supervisor source confirmed")
        self.log(f"Step 0 code reference confirmed: {Path(path).name} | tabs: {mrp_sheet}; {supervisor_sheet}")
        self.refresh_step0_state()

    def data_sources_ready(self) -> bool:
        return bool(self.routing_source_ready and self.material_sources.p41_ready and self.material_sources.code_reference_ready)

    def refresh_step0_state(self):
        ready_count = sum((self.material_sources.p41_ready, self.routing_source_ready, self.material_sources.code_reference_ready))
        self.step0_summary.set("All 3 sources confirmed — Step 1 VSM upload unlocked" if ready_count == 3 else f"{ready_count} of 3 sources confirmed")
        if self.material_sources.p41_ready:
            p41_name = Path(self.material_sources.p41_path).name
            code_name = Path(self.material_sources.code_reference_path).name if self.material_sources.code_reference_ready else "code reference pending"
            self.material_source_info.set(f"P41: {p41_name} | Codes: {code_name} | {len(self.material_sources.variants):,} routing variants available")
        self.update_action_states()

    def stage_available(self, stage: int) -> bool:
        file, page = self.selected_file, self.selected_page
        if stage == 0:
            return True
        if stage == 1:
            return self.data_sources_ready()
        if stage == 2:
            return bool(self.files)
        if stage == 3:
            return bool(file and page)
        if stage == 4:
            return bool(file and page and file.plant and self._page_crid(page))
        if stage == 5:
            return bool(file and any(page.change_rule_draft for page in file.pages))
        if stage == 6:
            return bool(self.material_sources.ready and any(page.change_rule_draft for file in self.files for page in file.pages))
        return bool(any(page.change_rule_draft for file in self.files for page in file.pages))

    def previous_stage(self):
        self.show_stage(max(0, getattr(self, "current_stage", 0) - 1))

    def next_stage(self):
        current = getattr(self, "current_stage", 0)
        next_stage = min(7, current + 1)
        if next_stage == current:
            return
        if not self.stage_available(next_stage):
            requirements = [
                "Confirm all three business data sources first.",
                "Add a VSDX file first.",
                "Select a file and tab first.",
                "Assign Plant and CRID first.",
                "Generate at least one Change Rule draft first.",
                "Generate a Change Rule draft before Material Selection.",
                "Generate Material Selection or continue to export routing outputs.",
            ]
            self.status.configure(text=requirements[max(0, current)])
            return
        self.show_stage(next_stage)

    def update_stage_buttons(self):
        if not hasattr(self, "stage_buttons"):
            return
        for index, button in enumerate(self.stage_buttons):
            button.configure(state=tk.NORMAL if self.stage_available(index) else tk.DISABLED)
        if hasattr(self, "next_stage_btn"):
            next_stage = min(7, getattr(self, "current_stage", 0) + 1)
            self.next_stage_btn.configure(
                state=tk.NORMAL if next_stage != getattr(self, "current_stage", 0) and self.stage_available(next_stage) else tk.DISABLED
            )

    def show_stage(self, stage: int):
        if not self.stage_available(stage):
            self.status.configure(text="Complete the previous manufacturing stage first")
            return
        panes = {
            0: [(self.left, 12)],
            1: [(self.left, 12)],
            2: [(self.left, 3), (self.center, 6), (self.right, 3)],
            3: [(self.center, 8), (self.right, 4)],
            4: [(self.center, 12)],
            5: [(self.center, 12)],
            6: [(self.center, 12)],
            7: [(self.left, 4), (self.center, 8)],
        }[stage]
        for pane in (self.left, self.center, self.right):
            try:
                self.body.forget(pane)
            except tk.TclError:
                pass
        for pane, weight in panes:
            self.body.add(pane, weight=weight)
        if stage in (4, 5, 6, 7):
            self.page_card.grid_remove()
        else:
            self.page_card.grid()
        for card in (self.source_intake_card, self.source_card, self.files_card, self.action_card):
            card.grid_remove()
        if stage == 0:
            self.source_intake_card.grid()
        elif stage == 1:
            self.files_card.grid()
        elif stage == 2:
            self.source_card.grid()
            self.files_card.grid()
        elif stage == 7:
            self.files_card.grid()
            self.action_card.grid()
        self.show_context_mapping(stage == 2)
        if stage == 2:
            # Context assignment deliberately excludes generation/export
            # controls; the user only sees the file, Plant, tab and CRID job.
            self.center_notebook.grid_remove()
        else:
            self.center_notebook.grid()
        if stage == 3:
            self.center_notebook.select(0)
        elif stage in (4, 5, 7):
            self.center_notebook.select(1)
        elif stage == 6:
            self.center_notebook.select(2)
        titles = [
            ("Step 0 of 7 — Data sources", "Confirm the exact workbook tabs used for P41, WorkCenter and MRP/Supervisor data."),
            ("Step 1 of 7 — Load VSM files", "Upload Visio files only after all three raw business sources are ready."),
            ("Step 2 of 7 — Assign context", "Assign Plant and CRID against the selected VSM tab."),
            ("Step 3 of 7 — Inspect process", "Review the Visio process and detected workcentres before creating rules."),
            ("Step 4 of 7 — Generate Change Rule", "Generate rules for the selected CRID, Plant, or all loaded Plants."),
            ("Step 5 of 7 — Review Change Rule", "Resolve Review Required records, then validate readiness."),
            ("Step 6 of 7 — Material Selection", "Use the confirmed P41 and code tabs to create the first-draft material scope."),
            ("Step 7 of 7 — Export package", "Export routing and Material Selection templates."),
        ]
        self.current_stage = stage
        self.stage_var.set(titles[stage][0])
        self.stage_help_var.set(titles[stage][1])
        for index, button in enumerate(self.stage_buttons):
            button.configure(style="Primary.TButton" if index == stage else "TButton")
        self.update_stage_buttons()

    def show_advanced(self):
        """Power-user escape hatch: restores the original all-pane workspace."""
        for pane in (self.left, self.center, self.right):
            try:
                self.body.forget(pane)
            except tk.TclError:
                pass
        self.body.add(self.left, weight=3)
        self.body.add(self.center, weight=6)
        self.body.add(self.right, weight=3)
        self.center_notebook.grid()
        self.page_card.grid()
        self.source_intake_card.grid_remove()
        self.source_card.grid()
        self.files_card.grid()
        self.action_card.grid()
        self.show_context_mapping(False)
        self.stage_var.set("Advanced workspace")
        self.stage_help_var.set("All controls are visible for experienced reviewers; routing logic is unchanged.")
        for button in self.stage_buttons:
            button.configure(style="TButton")

    def clear_mapping_filter_hint(self, _event=None):
        if self.context_mapping_filter.get() == "Filter CRID or workcenter":
            self.context_mapping_filter.set("")

    def show_context_mapping(self, show: bool):
        if not hasattr(self, "context_mapping_frame"):
            return
        if show:
            self.source_card.configure(text="Active Plant mapping", height=300)
            self.source_tabs.grid_remove()
            self.context_mapping_frame.grid()
            self.refresh_context_mapping()
        else:
            self.source_card.configure(text="Department / CRID mapping", height=180)
            self.context_mapping_frame.grid_remove()
            self.source_tabs.grid()

    def refresh_context_mapping(self):
        if not hasattr(self, "context_mapping_tree"):
            return
        self.context_mapping_tree.delete(*self.context_mapping_tree.get_children())
        plant = norm(self.selected_file.plant) if self.selected_file else ""
        query = self.context_mapping_filter.get().strip().upper()
        if query == "FILTER CRID OR WORKCENTER":
            query = ""
        title = f"Mapping records for Plant {plant}" if plant else "Select a VSDX file to view its Plant mapping"
        self.context_mapping_title.set(title)
        if not self.mapping or not plant:
            return
        records = [record for record in self.mapping.records if norm(record.plant) == plant]
        for record in sorted(records, key=lambda item: (norm(item.department), norm(item.workcenter), item.source_sheet)):
            searchable = " ".join([norm(record.department), norm(record.workcenter), record.source_sheet, record.source_file]).upper()
            if query and query not in searchable:
                continue
            source = record.source_sheet or Path(record.source_file).name or "Manual list"
            self.context_mapping_tree.insert("", tk.END, values=(norm(record.department), norm(record.workcenter) or "—", source))

    def mark_dirty(self, reason: str):
        self.dirty_reasons.add(reason)
        self.status.configure(text="● Unsaved changes")

    def mark_saved(self):
        self.dirty_reasons.clear()
        self.status.configure(text="✓ Saved")

    def mark_plant_dirty(self, _event=None):
        if not self.selected_file:
            return
        changed = norm(self.plant.get()) != self.plant_saved_value
        self.plant_state.set("● Unsaved Plant change — Apply to save" if changed else "✓ Saved")
        self.plant_apply_btn.configure(state=tk.NORMAL if changed else tk.DISABLED)
        if changed:
            self.mark_dirty("plant")

    def refresh_plant_buttons(self, values=None):
        """Render Plant values as direct choices instead of a long combobox."""
        if not hasattr(self, "plant_choice_frame"):
            return
        if values is None:
            values = self.mapping.plants if self.mapping else []
            values = list(values) + ([self.selected_file.plant] if self.selected_file and self.selected_file.plant else [])
        choices = []
        for value in values:
            plant = norm(value)
            if plant and plant not in choices:
                choices.append(plant)
        for button in self.plant_choice_frame.winfo_children():
            button.destroy()
        self.plant_buttons = {}
        for plant in choices:
            button = ttk.Button(
                self.plant_choice_frame,
                text=plant,
                command=lambda selected=plant: self.choose_plant(selected),
                style="Primary.TButton" if norm(self.plant.get()) == plant else "TButton",
            )
            button.pack(side=tk.LEFT, padx=(0, 4))
            self.plant_buttons[plant] = button
        ttk.Button(self.plant_choice_frame, text="Other…", command=self.choose_other_plant).pack(side=tk.LEFT)

    def choose_plant(self, plant: str):
        if not self.selected_file:
            return
        self.plant.set(norm(plant))
        self.refresh_plant_buttons()
        self.mark_plant_dirty()

    def choose_other_plant(self):
        if not self.selected_file:
            return
        value = simpledialog.askstring("Other Plant", "Enter the Plant code:", initialvalue=self.plant.get(), parent=self.root)
        if value is None:
            return
        self.plant.set(norm(value))
        self.refresh_plant_buttons()
        self.mark_plant_dirty()

    def update_context(self):
        file = self.selected_file
        page = self.selected_page
        if not file:
            self.context_var.set("No VSDX selected — choose a file to establish context.")
            return
        rows = len(page.change_rule_draft) if page else 0
        summary = readiness_summary(self._page_to_routing_rows(page)) if page and rows else {"readiness": 0}
        self.context_var.set(
            f"Plant: {file.plant or 'Unassigned'}  |  VSDX: {file.file_name}  |  "
            f"Tab: {page.name if page else 'No tab selected'}  |  CRID: {self._page_crid(page) if page else 'Unassigned'}  |  "
            f"Draft rows: {rows}  |  Readiness: {summary['readiness']:.0f}%"
        )

    def update_action_states(self):
        file, page = self.selected_file, self.selected_page
        generated = bool(page and page.change_rule_draft)
        if hasattr(self, "plant_apply_btn"):
            self.plant_apply_btn.configure(state=tk.NORMAL if file and norm(self.plant.get()) != self.plant_saved_value else tk.DISABLED)
        for name, enabled in {
            "run_btn": bool(self.files and all(item.plant for item in self.files) and not self.running),
            "selected_tab_btn": bool(file and page and file.plant and not self.running),
            "regenerate_all_btn": bool(self.files and all(item.plant for item in self.files) and not self.running),
            "generate_selected_btn": bool(file and page and file.plant and self._page_crid(page)),
            "generate_plant_btn": bool(file and file.plant),
            "validate_btn": generated,
            "clear_draft_btn": generated,
            "rerun_btn": generated,
            "export_selected_btn": generated,
            "generate_material_btn": bool(self.material_sources.ready and any(page.change_rule_draft for item in self.files for page in item.pages)),
            "export_material_btn": bool(self.material_rows),
        }.items():
            if hasattr(self, name):
                getattr(self, name).configure(state=tk.NORMAL if enabled else tk.DISABLED)
        self.update_stage_buttons()

    def analyze_selected_vsm(self):
        """Run the established page-load analysis for the highlighted VSM only."""
        if not self.selected_file or not self.selected_page:
            messagebox.showwarning("Selected Tab", "Select a VSM tab first.")
            return
        if not self.selected_file.plant:
            messagebox.showwarning("Selected Tab", "Assign the Plant for this VSDX file first.")
            return
        self.log(f"VSM analysis requested: {self.selected_file.file_name} → {self.selected_page.name}")
        self.status.configure(text=f"Analyzing VSM: {self.selected_page.name}")
        self.request_selected_preview(force=True)

    def regenerate_all_vsms(self):
        """Re-run the established full analysis for every loaded VSDX tab."""
        if self.running:
            return
        if not self.files:
            messagebox.showwarning("Regenerate all", "Add at least one VSDX file.")
            return
        self.log("VSM regeneration requested for all loaded VSDX files.")
        self.run_analysis()

    def zoom_selected(self, _event=None):
        try:
            self.set_zoom(float(self.zoom_var.get().rstrip("%")) / 100)
        except ValueError:
            pass

    def bind_shortcuts(self):
        self.root.bind_all("<Control-o>", lambda _e: self.add_files())
        self.root.bind_all("<Control-s>", lambda _e: self.save_package())
        self.root.bind_all("<Control-e>", lambda _e: self.export_change_rule())
        self.root.bind_all("<Control-f>", self.focus_search)
        self.root.bind_all("<F5>", lambda _e: self.rerun_selected_crid())
        self.root.bind_all("<Delete>", self.shortcut_clear_assignment)

    def focus_search(self, _event=None):
        if hasattr(self, "routing_search"):
            self.center_notebook.select(1)
            # The search entry is deliberately the first Entry in the rerun toolbar.
            for child in self.center_notebook.nametowidget(self.center_notebook.tabs()[1]).winfo_children():
                for item in child.winfo_children():
                    if isinstance(item, ttk.Entry) and str(item.cget("textvariable")) == str(self.routing_search):
                        item.focus_set()
                        return "break"

    def shortcut_clear_assignment(self, _event=None):
        if self.selected_page and self.center_notebook.index("current") == 0:
            self.clear_selected_department()
            return "break"

    def load_ui_settings(self):
        try:
            saved = json.loads(self.ui_settings_path.read_text(encoding="utf-8"))
            if saved.get("geometry"):
                self.root.geometry(saved["geometry"])
            if saved.get("workspace_tab") in (0, 1):
                self.center_notebook.select(saved["workspace_tab"])
            if saved.get("zoom"):
                self.set_zoom(float(saved["zoom"]))
        except (OSError, ValueError, json.JSONDecodeError):
            pass

    def save_ui_settings(self):
        try:
            self.ui_settings_path.write_text(json.dumps({
                "geometry": self.root.geometry(), "workspace_tab": self.center_notebook.index("current"),
                "zoom": self.zoom, "last_directory": getattr(self, "last_directory", ""),
            }, indent=2), encoding="utf-8")
        except OSError:
            pass

    def on_close(self):
        self.save_ui_settings()
        self.root.destroy()

    def bundle_roots(self) -> List[Path]:
        bundle_root = Path(
            getattr(
                sys,
                "_MEIPASS",
                Path(__file__).resolve().parent.parent,
            )
        )
        return [
            Path.cwd(),
            Path(sys.executable).resolve().parent,
            bundle_root,
            bundle_root / "data",
            Path(__file__).resolve().parent.parent / "data",
        ]

    def auto_mapping(self):
        candidates: List[Path] = []
        seen = set()

        for folder in self.bundle_roots():
            for pattern in (
                "*PTS03*CRMS*.xlsx",
                "*Workcentre*CRMS*.xlsx",
            ):
                try:
                    for candidate in folder.glob(pattern):
                        key = str(candidate.resolve()).lower()
                        if key not in seen:
                            seen.add(key)
                            candidates.append(candidate)
                except Exception:
                    pass

        if candidates:
            self.load_workbook_mapping(str(candidates[0]))
        else:
            self.map_info.set(
                "No Excel mapping loaded. Browse to select one, "
                "or use the Department | Plant list tab."
            )
            self.rebuild_mapping()


    def auto_routing_reference(self):
        candidates = []
        for folder in self.bundle_roots():
            for pattern in ("Routing_Generation_Template.xlsx", "*ROUTING*GENERATION*.xlsx", "*NEW TEMPLATE*.xlsx"):
                try:
                    candidates.extend(folder.glob(pattern))
                except Exception:
                    pass
        if candidates:
            self.load_routing_template(str(candidates[0]))

    def browse_routing_template(self):
        path = filedialog.askopenfilename(
            title="Select routing-generation template / PTS03 reference",
            filetypes=[("Excel workbook", "*.xlsx"), ("All files", "*.*")],
        )
        if path:
            self.load_routing_template(path)

    def load_routing_template(self, path: str):
        try:
            self.routing_reference = load_routing_reference(path)
            self.routing_template_path = path
            self.routing_template_var.set(path)
            self.routing_reference_info.set(
                f"Loaded {len(self.routing_reference.records):,} PTS03/workcenter rows for OLD WC lookup."
            )
            self.log(f"Routing reference loaded: {Path(path).name} ({len(self.routing_reference.records):,} rows).")
        except Exception as exc:
            self.routing_reference = None
            self.routing_reference_info.set(f"Routing reference error: {exc}")
            messagebox.showerror("Routing reference", str(exc))

    def browse_mapping(self):
        path = filedialog.askopenfilename(
            title="Select Workcentre vs CRMS analysis workbook",
            filetypes=[
                ("Excel workbook", "*.xlsx"),
                ("All files", "*.*"),
            ],
        )
        if path:
            self.load_workbook_mapping(path)

    def load_workbook_mapping(self, path: str):
        try:
            self.status.configure(text="Loading mapping…")
            self.root.update_idletasks()

            self.workbook_mapping = load_mapping(path, self.log)
            self.map_path.set(path)
            mapping = self.workbook_mapping
            self.map_info.set(
                f"Sheet: {mapping.sheet_name} | "
                f"Header row: {mapping.header_row} | "
                f"{len(mapping.records):,} rows | "
                f"Plants: {', '.join(mapping.plants)}"
            )
            self.rebuild_mapping()
            self.invalidate_analyzed_files()
            self.status.configure(text="Mapping ready")
        except Exception as exc:
            self.status.configure(text="Mapping error")
            messagebox.showerror("Mapping workbook", str(exc))

    def apply_department_plant_list(self):
        text = self.department_list_text.get("1.0", tk.END)
        try:
            mapping, warnings = parse_department_plant_text(text)
            self.manual_mapping = mapping
            self.manual_info.set(
                f"Loaded {mapping.department_plant_pair_count:,} "
                f"Department-Plant pairs across "
                f"{len(mapping.plants)} plants."
                + (
                    f" Skipped {len(warnings)} invalid line(s)."
                    if warnings
                    else ""
                )
            )
            self.rebuild_mapping()
            self.log(
                f"Department-Plant list loaded: "
                f"{mapping.department_plant_pair_count:,} pairs, "
                f"{len(mapping.plants)} plants."
            )
            for warning in warnings[:10]:
                self.log("Department-Plant warning: " + warning)
            if len(warnings) > 10:
                self.log(
                    f"Department-Plant warnings: "
                    f"{len(warnings) - 10} additional line(s) omitted."
                )
            self.status.configure(text="Department list ready")
        except DepartmentPlantListError as exc:
            messagebox.showerror(
                "Department | Plant list", str(exc)
            )

    def paste_department_plant_list(self):
        try:
            clipboard = self.root.clipboard_get()
        except tk.TclError:
            messagebox.showwarning(
                "Clipboard",
                "The clipboard does not contain text.",
            )
            return

        self.department_list_text.delete("1.0", tk.END)
        self.department_list_text.insert("1.0", clipboard)
        self.apply_department_plant_list()

    def import_department_plant_list(self):
        path = filedialog.askopenfilename(
            title="Import Department | Plant list",
            filetypes=[
                ("Text and CSV", "*.txt *.csv *.tsv"),
                ("All files", "*.*"),
            ],
        )
        if not path:
            return

        try:
            mapping, warnings = load_department_plant_file(path)
            text = Path(path).read_text(
                encoding="utf-8-sig", errors="replace"
            )
            self.department_list_text.delete("1.0", tk.END)
            self.department_list_text.insert("1.0", text)
            self.manual_mapping = mapping
            self.manual_info.set(
                f"Imported {mapping.department_plant_pair_count:,} "
                f"Department-Plant pairs from {Path(path).name}."
                + (
                    f" Skipped {len(warnings)} invalid line(s)."
                    if warnings
                    else ""
                )
            )
            self.rebuild_mapping()
            self.log(
                f"Department-Plant file imported: {Path(path).name} "
                f"({mapping.department_plant_pair_count:,} pairs)."
            )
        except Exception as exc:
            messagebox.showerror(
                "Department | Plant import", str(exc)
            )

    def clear_department_plant_list(self):
        self.department_list_text.delete("1.0", tk.END)
        self.department_list_text.insert(
            "1.0", "DEPARTMENT | PLANT\n"
        )
        self.manual_mapping = None
        self.manual_info.set(
            "No Department-Plant list loaded."
        )
        self.rebuild_mapping()
        self.log("Department-Plant list cleared.")

    def rebuild_mapping(self):
        self.mapping = merge_mappings(
            self.workbook_mapping,
            self.manual_mapping,
        )

        plants = set(
            self.mapping.plants if self.mapping else []
        )
        plants.update(file.plant for file in self.files if file.plant)
        self.refresh_plant_buttons(sorted(plants))
        self.refresh_context_mapping()

        # Recalculate suggestions immediately for tabs already loaded for preview.
        for file in self.files:
            for page in file.pages:
                if page.loaded:
                    apply_mapping(page, file.plant, self.mapping)

        self.refresh_dept_values()
        if self.selected_page:
            self.show_candidates()
            self.show_page_meta()
            self.render_preview()

    def invalidate_analyzed_files(self):
        loaded_pages = sum(
            page.loaded
            for file in self.files
            for page in file.pages
        )
        if loaded_pages:
            self.refresh_pages()
            self.log(
                f"Mapping changed. Refreshed Department/CRID suggestions "
                f"for {loaded_pages:,} loaded tab(s)."
            )

    def add_files(self):
        if not self.data_sources_ready():
            messagebox.showwarning("Step 0 required", "Confirm all three business data sources and their tab selections before uploading VSM files.")
            self.show_stage(0)
            return
        paths = filedialog.askopenfilenames(
            title="Select VSDX files",
            filetypes=[
                ("Visio VSDX", "*.vsdx"),
                ("All files", "*.*"),
            ],
        )
        self.add_paths(paths)

    def add_folder(self):
        if not self.data_sources_ready():
            messagebox.showwarning("Step 0 required", "Confirm all three business data sources and their tab selections before uploading a VSM folder.")
            self.show_stage(0)
            return
        path = filedialog.askdirectory(
            title="Select folder containing VSDX files"
        )
        if path:
            self.add_paths(
                sorted(str(item) for item in Path(path).glob("*.vsdx"))
            )

    def add_paths(self, paths):
        existing = {
            str(Path(file.path).resolve()).lower()
            for file in self.files
        }

        for path in paths:
            resolved = str(Path(path).resolve())
            if resolved.lower() in existing:
                continue

            try:
                pages = quick_scan(resolved)
                plant_match = re.search(
                    r"(?<![A-Z0-9])([A-Z]{2}\d{2})(?![A-Z0-9])",
                    Path(resolved).name.upper(),
                )
                plant = (
                    plant_match.group(1) if plant_match else ""
                )
                record = FileRecord(
                    resolved,
                    plant,
                    pages,
                    "Scanned",
                    "",
                )
                self.files.append(record)
                existing.add(resolved.lower())
                self.log(
                    f"Quick scan: {record.file_name} — "
                    f"{len(pages)} tabs: "
                    + ", ".join(page.name for page in pages)
                )
            except Exception as exc:
                messagebox.showerror(
                    "VSDX quick scan",
                    f"{Path(path).name}\n\n{exc}",
                )

        self.rebuild_mapping()
        self.refresh_files()
        self.export_btn.configure(state=tk.DISABLED)
        self.status.configure(
            text=f"{len(self.files)} file(s) ready"
        )

    def refresh_files(self):
        selected_path = (
            self.selected_file.path if self.selected_file else ""
        )
        self.file_tree.delete(*self.file_tree.get_children())
        self.file_map = {}

        for index, file in enumerate(self.files):
            item_id = f"f{index}"
            self.file_map[item_id] = file
            self.file_tree.insert(
                "",
                tk.END,
                iid=item_id,
                values=(
                    file.file_name,
                    file.plant,
                    len(file.pages),
                    file.status,
                ),
            )
            if file.path == selected_path:
                self.file_tree.selection_set(item_id)

        if self.files and not self.file_tree.selection():
            item_id = self.file_tree.get_children()[0]
            self.file_tree.selection_set(item_id)
            self.file_tree.focus(item_id)
            self.file_selected()
        self.update_stage_buttons()

    def remove_file(self):
        selection = self.file_tree.selection()
        file = (
            self.file_map.get(selection[0])
            if selection
            else None
        )
        if file in self.files:
            self.files.remove(file)

        self.selected_file = None
        self.selected_page = None
        self.rebuild_mapping()
        self.refresh_files()
        self.refresh_pages()

    def clear_files(self):
        if self.files and not messagebox.askyesno(
            "Clear files", "Remove all loaded VSDX files?"
        ):
            return

        self.files = []
        self.selected_file = None
        self.selected_page = None
        self.rebuild_mapping()
        self.refresh_files()
        self.refresh_pages()
        self.canvas.delete("all")
        self.export_btn.configure(state=tk.DISABLED)

    def file_selected(self, _event=None):
        selection = self.file_tree.selection()
        self.selected_file = (
            self.file_map.get(selection[0])
            if selection
            else None
        )
        self.plant.set(
            self.selected_file.plant
            if self.selected_file
            else ""
        )
        self.plant_saved_value = self.selected_file.plant if self.selected_file else ""
        self.plant_state.set("✓ Saved" if self.selected_file else "Select a VSDX file")
        self.refresh_plant_buttons()
        self.selected_page = None
        self.refresh_pages()
        self.refresh_context_mapping()
        self.update_context()
        self.update_action_states()

    def file_double(self, event):
        row = self.file_tree.identify_row(event.y)
        column = self.file_tree.identify_column(event.x)
        if row and column == "#2":
            values = (
                self.mapping.plants if self.mapping else []
            )
            self.edit_cell(
                self.file_tree,
                row,
                column,
                values,
                lambda value: self.set_file_plant(row, value),
            )

    def set_file_plant(self, item_id: str, value: str):
        file = self.file_map.get(item_id)
        if file:
            file.plant = norm(value)
            file.status = "Scanned"
            self.plant.set(file.plant)
            self.refresh_plant_buttons()
            self.rebuild_mapping()
            self.refresh_files()
            self.refresh_pages()
            self.export_btn.configure(state=tk.DISABLED)

    def apply_plant(self):
        if not self.selected_file:
            return

        self.selected_file.plant = norm(self.plant.get())
        self.selected_file.status = "Scanned"
        self.plant_saved_value = self.selected_file.plant
        self.plant_state.set("✓ Saved")
        self.refresh_plant_buttons()
        self.log(
            f"Plant assigned: {self.selected_file.file_name} → "
            f"{self.selected_file.plant or '(blank)'}"
        )
        self.rebuild_mapping()
        self.refresh_files()
        self.refresh_pages()
        self.refresh_context_mapping()
        self.export_btn.configure(state=tk.DISABLED)
        self.mark_saved()
        self.update_context()
        self.update_action_states()

    def refresh_pages(self):
        selected_page_id = (
            self.selected_page.page_id
            if self.selected_page
            else ""
        )
        self.page_tree.delete(*self.page_tree.get_children())
        self.page_map = {}

        if not self.selected_file:
            self.show_meta("")
            self.candidates.delete(
                *self.candidates.get_children()
            )
            self.update_context()
            self.update_action_states()
            return

        selected_item = None
        for index, page in enumerate(self.selected_file.pages):
            item_id = f"p{index}"
            self.page_map[item_id] = page
            self.page_tree.insert(
                "",
                tk.END,
                iid=item_id,
                values=(
                    page.name,
                    (
                        len(page.shapes)
                        if page.loaded
                        else ("…" if page.loading else "")
                    ),
                    (
                        len(page.workcenter_counts)
                        if page.loaded
                        else ("…" if page.loading else "")
                    ),
                    page.suggested_department,
                    (
                        f"{page.confidence:.0%}"
                        if page.candidate_counts
                        else ""
                    ),
                    page.assigned_department,
                    self.page_status(page),
                    "✕" if page.assigned_department else "",
                ),
            )
            if page.page_id == selected_page_id:
                selected_item = item_id

        if self.selected_file.pages:
            item_id = (
                selected_item
                or self.page_tree.get_children()[0]
            )
            self.page_tree.selection_set(item_id)
            self.page_tree.focus(item_id)
            self.page_selected()

    def page_selected(self, _event=None):
        selection = self.page_tree.selection()
        self.selected_page = (
            self.page_map.get(selection[0])
            if selection
            else None
        )
        if not self.selected_page:
            return

        self.department.set(
            self.selected_page.assigned_department
        )
        self.refresh_dept_values()
        self.show_page_meta()
        self.show_candidates()
        self.zoom = 1
        self.render_preview()
        self.refresh_routing_tree()
        self.update_context()
        self.update_action_states()
        if self.auto_preview.get() and not self.selected_page.loaded:
            self.request_selected_preview()

    def auto_preview_changed(self):
        if self.auto_preview.get():
            self.request_selected_preview()

    def reload_preview(self):
        self.request_selected_preview(force=True)

    def request_selected_preview(self, force: bool = False):
        file = self.selected_file
        page = self.selected_page
        if not file or not page:
            return
        if page.loading:
            return
        if page.loaded and not force:
            self.render_preview()
            return

        self.preview_request_token += 1
        token = self.preview_request_token
        page.loading = True
        page.load_error = ""
        self.preview_mode.set(f"Loading visual preview: {page.name}")
        self.status.configure(text="Loading selected tab…")
        self.render_preview()
        self.refresh_pages()
        self.log(
            f"Visual preview loading: {file.file_name} → {page.name}"
        )

        def worker():
            started = time.perf_counter()
            try:
                load_page(
                    file.path,
                    page,
                    file.plant,
                    self.mapping,
                    force=force,
                )
                self.events.put(
                    (
                        "preview_loaded",
                        file.path,
                        page.page_id,
                        token,
                        time.perf_counter() - started,
                    )
                )
            except Exception as exc:
                self.events.put(
                    (
                        "preview_error",
                        file.path,
                        page.page_id,
                        token,
                        str(exc),
                    )
                )

        threading.Thread(target=worker, daemon=True).start()

    def page_click(self, event):
        row = self.page_tree.identify_row(event.y)
        column = self.page_tree.identify_column(event.x)
        if row and column == "#7":
            self.clear_page_department(row)

    def page_double(self, event):
        row = self.page_tree.identify_row(event.y)
        column = self.page_tree.identify_column(event.x)
        if row and column == "#6":
            self.edit_cell(
                self.page_tree,
                row,
                column,
                self.dept_values(self.page_map.get(row)),
                lambda value: self.set_page_dept(row, value),
            )

    def dept_values(
        self, page: Optional[PageRecord] = None
    ) -> List[str]:
        values = set(
            page.candidate_counts if page else []
        )

        if self.mapping and self.selected_file:
            values.update(
                self.mapping.departments_for_plant(
                    self.selected_file.plant
                )
            )

        if page and page.assigned_department:
            values.add(page.assigned_department)

        return sorted(values)

    def refresh_dept_values(self):
        self.department_combo["values"] = self.dept_values(
            self.selected_page
        )

    def set_page_dept(self, item_id: str, value: str):
        page = self.page_map.get(item_id)
        if not page:
            return
        department = norm(value)
        page.assigned_department = department
        page.assignment_mode = "manual" if department else "cleared"
        self.department.set(page.assigned_department)
        self.refresh_pages()

    def clear_page_department(self, item_id: str):
        page = self.page_map.get(item_id)
        if not page:
            return
        if page.assignment_mode == "manual" and not messagebox.askyesno(
            "Clear manual CRID assignment", f"Clear the manual CRID assignment for '{page.name}'?"
        ):
            return
        page.assigned_department = ""
        page.assignment_mode = "cleared"
        if self.selected_page is page:
            self.department.set("")
        self.log(f"Department/CRID cleared: {page.name}")
        self.mark_dirty("CRID assignment")
        self.refresh_pages()

    def clear_selected_department(self):
        selection = self.page_tree.selection()
        if selection:
            self.clear_page_department(selection[0])
        elif self.selected_page:
            self.selected_page.assigned_department = ""
            self.selected_page.assignment_mode = "cleared"
            self.department.set("")
            self.refresh_pages()

    def apply_department(self):
        if not self.selected_page:
            return

        department = norm(self.department.get())
        self.selected_page.assigned_department = department
        self.selected_page.assignment_mode = "manual" if department else "cleared"
        self.log(
            f"Department/CRID assigned: "
            f"{self.selected_page.name} → "
            f"{self.selected_page.assigned_department or '(cleared)'}"
        )
        self.mark_dirty("CRID assignment")
        self.refresh_pages()

    def page_status(self, page: PageRecord) -> str:
        if page.change_rule_draft:
            rows = self._page_to_routing_rows(page)
            return "Ready" if readiness_summary(rows)["readiness"] >= 100 else "Review Required"
        if page.loading:
            return "Loading"
        if page.loaded:
            return "Analyzed" if page.assigned_department else "Review Required"
        return "Scanned" if page.assigned_department else "Loaded"

    def use_candidate(self, _event=None):
        selection = self.candidates.selection()
        if not selection:
            return
        self.department.set(
            self.candidates.item(
                selection[0], "values"
            )[0]
        )
        self.apply_department()

    def edit_cell(
        self,
        tree,
        row,
        column,
        values,
        callback,
    ):
        box = tree.bbox(row, column)
        if not box:
            return

        x, y, width, height = box
        editor = ttk.Combobox(
            tree, values=values, state="normal"
        )
        editor.set(tree.set(row, column))
        editor.place(
            x=x, y=y, width=width, height=height
        )
        editor.focus_set()
        editor.selection_range(0, tk.END)

        committed = {"value": False}

        def commit(_event=None):
            if committed["value"]:
                return
            committed["value"] = True
            callback(editor.get())
            editor.destroy()

        editor.bind("<Return>", commit)
        editor.bind("<<ComboboxSelected>>", commit)
        editor.bind("<FocusOut>", commit)
        editor.bind(
            "<Escape>", lambda _event: editor.destroy()
        )


    def _page_to_routing_rows(self, page: Optional[PageRecord] = None):
        page = page or self.selected_page
        if not page:
            return []
        return [
            ChangeRuleRow(
                crid=row.crid,
                plant=row.plant,
                alternate=row.alternate,
                sequence=int(row.sequence or 0),
                subseq=int(row.subseq or 1),
                old_wc=row.old_wc,
                new_wc=row.new_wc,
                op_descriptions=row.op_descriptions,
                rule_type=row.rule_type,
                column=row.column,
                status=row.status,
                notes=row.notes,
                source_shape_id=row.source_shape_id,
            )
            for row in page.change_rule_draft
        ]

    def _draft_to_routing_rows(self):
        return self._page_to_routing_rows(self.selected_page)

    def _store_rows_on_page(self, page: PageRecord, rows):
        page.change_rule_draft = [
            ChangeRuleDraftRecord(
                crid=row.crid,
                plant=row.plant,
                alternate=row.alternate,
                sequence=int(row.sequence or 0),
                subseq=int(row.subseq or 1),
                old_wc=row.old_wc,
                new_wc=row.new_wc,
                op_descriptions=row.op_descriptions,
                rule_type=row.rule_type,
                column=row.column,
                status=row.status,
                notes=row.notes,
                source_shape_id=row.source_shape_id,
            )
            for row in rows
        ]

    def _store_routing_rows(self, rows):
        if self.selected_page:
            self._store_rows_on_page(self.selected_page, rows)

    def _valid_workcenters(self, file: FileRecord):
        return self.mapping.workcenters_for_plant(file.plant) if self.mapping else []

    def _page_crid(self, page: PageRecord) -> str:
        if page.assignment_mode == "cleared":
            return ""
        return page.assigned_department or page.suggested_department

    def _ensure_page_for_routing(self, file: FileRecord, page: PageRecord, force: bool = False):
        load_page(file.path, page, file.plant, self.mapping, force=force)
        return self._page_crid(page)

    def _history_item(self, file: FileRecord, page: PageRecord):
        return {
            "plant": file.plant,
            "crid": self._page_crid(page),
            "generated_at": page.change_rule_generated_at,
            "rerun_at": page.change_rule_rerun_at,
            "source": page.change_rule_source or file.file_name,
            "status": "Updated" if page.change_rule_rerun_at else ("Generated" if page.change_rule_generated_at else ""),
            "count": page.change_rule_generation_count,
        }

    def _all_history(self, files=None):
        output = []
        for file in files or self.files:
            for page in file.pages:
                if page.change_rule_generated_at or page.change_rule_rerun_at:
                    output.append(self._history_item(file, page))
        return output

    def _preview_changes_dialog(self, page: PageRecord, old_rows, new_rows):
        changes = compare_change_rule_rows(old_rows, new_rows)
        if not changes:
            return "replace"

        choice = {"value": None}
        dialog = tk.Toplevel(self.root)
        dialog.title(f"Rerun preview — {page.name}")
        dialog.geometry("980x520")
        dialog.transient(self.root)
        dialog.grab_set()
        dialog.columnconfigure(0, weight=1)
        dialog.rowconfigure(1, weight=1)

        ttk.Label(
            dialog,
            text=(
                f"{len(changes)} detected change(s). Review before replacing the existing draft. "
                "Ambiguous routing logic remains flagged for analyst review."
            ),
            style="Sub.TLabel",
            wraplength=920,
        ).grid(row=0, column=0, sticky="ew", padx=12, pady=10)

        frame = ttk.Frame(dialog)
        frame.grid(row=1, column=0, sticky="nsew", padx=12)
        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(0, weight=1)
        tree = ttk.Treeview(frame, columns=("change","field","before","after"), show="headings")
        for col, title, width in [
            ("change","CHANGE",80),("field","FIELD",140),("before","BEFORE",320),("after","AFTER",320)
        ]:
            tree.heading(col, text=title)
            tree.column(col, width=width, anchor="w")
        for idx, change in enumerate(changes[:500]):
            tree.insert("", tk.END, iid=f"d{idx}", values=(change["change"], change["field"], change["before"], change["after"]))
        y = ttk.Scrollbar(frame, orient=tk.VERTICAL, command=tree.yview)
        tree.configure(yscrollcommand=y.set)
        tree.grid(row=0, column=0, sticky="nsew")
        y.grid(row=0, column=1, sticky="ns")

        buttons = ttk.Frame(dialog)
        buttons.grid(row=2, column=0, sticky="e", padx=12, pady=12)
        def finish(value):
            choice["value"] = value
            dialog.destroy()
        ttk.Button(buttons, text="Cancel", command=lambda: finish(None)).pack(side=tk.RIGHT, padx=(6,0))
        ttk.Button(buttons, text="Merge & preserve edits", command=lambda: finish("merge")).pack(side=tk.RIGHT, padx=(6,0))
        ttk.Button(buttons, text="Replace draft", command=lambda: finish("replace"), style="Primary.TButton").pack(side=tk.RIGHT)
        dialog.protocol("WM_DELETE_WINDOW", lambda: finish(None))
        self.root.wait_window(dialog)
        return choice["value"]

    def _rerun_policy(self):
        value = self.rerun_policy_var.get() if hasattr(self, "rerun_policy_var") else "Preview changes first"
        if value.startswith("Replace"):
            return "replace"
        if value.startswith("Merge"):
            return "merge"
        return "preview"

    def _generate_page_change_rule(self, file: FileRecord, page: PageRecord, rerun: bool = False, policy: str = "replace"):
        if not file.plant:
            return False, "Missing Plant"
        try:
            crid = self._ensure_page_for_routing(file, page, force=rerun)
        except Exception as exc:
            self.log(f"Routing generation failed for {file.file_name} / {page.name}: {exc}")
            return False, str(exc)
        if not crid:
            self.log(f"Skipped routing draft: {file.file_name} / {page.name} — no assigned Department/CRID.")
            return False, "No CRID"

        valid = self._valid_workcenters(file)
        fresh = generate_change_rule_draft(
            page=page,
            plant=file.plant,
            crid=crid,
            reference=self.routing_reference,
            valid_workcenters=valid,
        )
        if not fresh:
            self.log(f"Skipped routing draft: {file.file_name} / {page.name} — no workcenter operations detected.")
            return False, "No routing operations"

        old = self._page_to_routing_rows(page)
        applied = fresh
        if rerun and old:
            selected_policy = policy
            if selected_policy == "preview":
                selected_policy = self._preview_changes_dialog(page, old, fresh)
                if selected_policy is None:
                    return False, "Cancelled"
            if selected_policy == "merge":
                applied = merge_preserve_manual_edits(old, fresh)
                applied = validate_change_rule_rows(applied, set(valid))

        self._store_rows_on_page(page, applied)
        now = datetime.now().isoformat(timespec="seconds")
        if not page.change_rule_generated_at:
            page.change_rule_generated_at = now
        if rerun:
            page.change_rule_rerun_at = now
        page.change_rule_source = file.file_name
        page.change_rule_generation_count = int(page.change_rule_generation_count or 0) + 1
        summary = readiness_summary(applied)
        file.status = "Ready" if summary["readiness"] >= 100 else "Review Required"
        self.log(
            f"{'Rerun' if rerun else 'Generated'}: {file.plant} / {crid} / {page.name} — "
            f"{summary['total']} rows; {summary['ready']} ready; {summary['review']} review."
        )
        return True, "OK"

    def generate_change_rule(self):
        file = self.selected_file
        page = self.selected_page
        if not file or not page:
            messagebox.showwarning("Change Rule Draft", "Select a VSDX tab first.")
            return
        if not file.plant:
            messagebox.showwarning("Change Rule Draft", "Assign the Plant for the selected VSDX file first.")
            return
        ok, message = self._generate_page_change_rule(file, page, rerun=False, policy="replace")
        if not ok and message == "No CRID":
            messagebox.showwarning("Change Rule Draft", "Assign a Department/CRID to the selected tab before generating a routing draft.")
            return
        self.routing_filter_var.set(self._page_crid(page) or "ALL DEPARTMENTS")
        self.refresh_pages()
        self.refresh_routing_tree()
        self.center_notebook.select(1)
        self.status.configure(text="Change Rule draft generated" if ok else "No draft generated")
        if ok:
            self.mark_dirty("draft generation")

    def _generate_scope(self, files, rerun=False, policy="replace"):
        generated = 0
        skipped = 0
        for file in files:
            if not file.plant:
                skipped += len(file.pages)
                self.log(f"Skipped {file.file_name}: Plant is blank.")
                continue
            for page in file.pages:
                # Explicitly cleared pages are intentionally excluded.
                if page.assignment_mode == "cleared":
                    skipped += 1
                    continue
                ok, _ = self._generate_page_change_rule(file, page, rerun=rerun, policy=policy)
                generated += int(ok)
                skipped += int(not ok)
        self.refresh_pages()
        self.refresh_routing_tree()
        self.center_notebook.select(1)
        self.status.configure(text=f"Routing generation complete — {generated} CRID tab(s)")
        self.log(f"Batch routing generation complete: {generated} generated/updated; {skipped} skipped.")
        return generated

    def generate_all_crids_for_plant(self):
        if not self.selected_file:
            messagebox.showwarning("Generate All CRIDs", "Select a VSDX file / Plant first.")
            return
        self.routing_filter_var.set("ALL DEPARTMENTS")
        self._generate_scope([self.selected_file], rerun=False, policy="replace")

    def generate_all_loaded_plants(self):
        if not self.files:
            messagebox.showwarning("Generate All Plants", "Load at least one VSDX file first.")
            return
        if not messagebox.askyesno("Generate All Loaded Plants", f"Generate Change Rule drafts for all assigned CRIDs across {len(self.files)} loaded VSDX file(s)?"):
            return
        self.routing_filter_var.set("ALL DEPARTMENTS")
        self._generate_scope(self.files, rerun=False, policy="replace")

    def rerun_selected_crid(self):
        if not self.selected_file or not self.selected_page:
            messagebox.showwarning("Rerun Selected CRID", "Select a Plant and Department/CRID tab first.")
            return
        self._generate_page_change_rule(self.selected_file, self.selected_page, rerun=True, policy=self._rerun_policy())
        self.refresh_pages()
        self.refresh_routing_tree()

    def rerun_selected_plant(self):
        if not self.selected_file:
            messagebox.showwarning("Rerun Selected Plant", "Select a VSDX file / Plant first.")
            return
        self.routing_filter_var.set("ALL DEPARTMENTS")
        self._generate_scope([self.selected_file], rerun=True, policy=self._rerun_policy())

    def rerun_everything(self):
        if not self.files:
            return
        if not messagebox.askyesno("Rerun Everything", "Rerun routing preparation for every assigned Department/CRID in all loaded Plants?"):
            return
        self.routing_filter_var.set("ALL DEPARTMENTS")
        self._generate_scope(self.files, rerun=True, policy=self._rerun_policy())

    def _routing_pages_for_current_view(self):
        if not self.selected_file:
            return []
        selected = self.routing_filter_var.get() if hasattr(self, "routing_filter_var") else "ALL DEPARTMENTS"
        pages = [page for page in self.selected_file.pages if page.change_rule_draft]
        if selected and selected != "ALL DEPARTMENTS":
            pages = [page for page in pages if norm(self._page_crid(page)) == norm(selected)]
        return pages

    def validate_change_rule(self):
        pages = self._routing_pages_for_current_view()
        if not pages and self.selected_page and self.selected_page.change_rule_draft:
            pages = [self.selected_page]
        if not pages:
            messagebox.showwarning("Validate Draft", "Generate a Change Rule draft first.")
            return
        for page in pages:
            file = self.selected_file
            valid = set(self._valid_workcenters(file)) if file else set()
            rows = validate_change_rule_rows(self._page_to_routing_rows(page), valid)
            self._store_rows_on_page(page, rows)
        self.refresh_routing_tree()
        all_rows = [row for page in pages for row in self._page_to_routing_rows(page)]
        summary = readiness_summary(all_rows)
        self.log(
            f"Change Rule validation: {summary['ready']}/{summary['total']} ready; "
            f"{summary['review']} need review; {summary['missing_old']} missing OLD WC."
        )
        self.status.configure(text=f"Routing draft readiness {summary['readiness']:.0f}%")

    def clear_change_rule(self):
        if not self.selected_page:
            return
        if self.selected_page.change_rule_draft and not messagebox.askyesno(
            "Clear Change Rule Draft", "Clear all draft Change Rule rows for this selected tab?"
        ):
            return
        self.selected_page.change_rule_draft = []
        self.selected_page.change_rule_generated_at = ""
        self.selected_page.change_rule_rerun_at = ""
        self.selected_page.change_rule_generation_count = 0
        self.refresh_routing_tree()
        self.log(f"Change Rule draft cleared: {self.selected_page.name}")

    def refresh_routing_tree(self):
        if not hasattr(self, "routing_tree"):
            return
        self.routing_tree.delete(*self.routing_tree.get_children())
        self.routing_row_map = {}
        if hasattr(self, "readiness_tree"):
            self.readiness_tree.delete(*self.readiness_tree.get_children())
        if not self.selected_file:
            self.routing_summary.set("Select a Plant/VSDX file to prepare routing generation.")
            return

        crids = sorted({self._page_crid(page) for page in self.selected_file.pages if self._page_crid(page)})
        values = ["ALL DEPARTMENTS"] + crids
        current = self.routing_filter_var.get() if hasattr(self, "routing_filter_var") else "ALL DEPARTMENTS"
        if current not in values:
            current = "ALL DEPARTMENTS"
            self.routing_filter_var.set(current)
        self.routing_filter_combo["values"] = values

        pages = self._routing_pages_for_current_view()
        shown_rows = []
        row_no = 0
        for page in pages:
            for index, row in enumerate(page.change_rule_draft):
                status_filter = self.routing_status_filter.get() if hasattr(self, "routing_status_filter") else "All"
                searchable = " ".join(map(str, [row.crid, row.plant, row.old_wc, row.new_wc, row.op_descriptions, row.status, row.notes])).lower()
                query = self.routing_search.get().strip().lower() if hasattr(self, "routing_search") else ""
                is_ready = row.status.strip().lower() == "ready"
                if (status_filter == "Ready" and not is_ready) or (status_filter == "Review Required" and is_ready) or (query and query not in searchable):
                    continue
                iid = f"cr-{row_no}"
                self.routing_row_map[iid] = (page, index)
                self.routing_tree.insert(
                    "", tk.END, iid=iid, values=(
                        row.crid, row.plant, row.alternate, row.sequence, row.subseq,
                        row.old_wc, row.new_wc, row.op_descriptions, row.rule_type,
                        row.column, row.status, row.notes,
                    ), tags=("even" if row_no % 2 == 0 else "odd", "ready" if is_ready else "review")
                )
                shown_rows.append(self._page_to_routing_rows(page)[index])
                row_no += 1

        all_file_rows = [row for page in self.selected_file.pages for row in self._page_to_routing_rows(page)]
        if hasattr(self, "readiness_tree"):
            for idx, item in enumerate(department_readiness(all_file_rows)):
                self.readiness_tree.insert(
                    "", tk.END, iid=f"ready-{idx}", values=(
                        item["plant"], item["crid"], item["rows"], item["ready"], item["review"],
                        item["missing_old"], f"{item['readiness']:.0f}%", item["status"],
                        self.last_generated_for(item["plant"], item["crid"]),
                    )
                )

        summary = readiness_summary(shown_rows)
        if shown_rows:
            self.routing_summary.set(
                f"{summary['ready']}/{summary['total']} Ready | Review {summary['review']} | "
                f"Missing OLD WC {summary['missing_old']} | Readiness {summary['readiness']:.0f}%"
            )
        else:
            self.routing_summary.set(
                f"{self.selected_file.plant or 'No Plant'} — no Change Rule drafts in current view"
            )
        self.routing_tree.tag_configure("even", background="#F8FAFC")
        self.routing_tree.tag_configure("odd", background="#EEF6F5")
        self.routing_tree.tag_configure("review", foreground="#9A3412")
        self.update_context()
        self.update_action_states()
        self.update_stage_buttons()

    def last_generated_for(self, plant: str, crid: str) -> str:
        if not self.selected_file:
            return ""
        matches = [page for page in self.selected_file.pages if norm(page.plant if hasattr(page, "plant") else self.selected_file.plant) == norm(plant) and norm(self._page_crid(page)) == norm(crid)]
        return max((page.change_rule_rerun_at or page.change_rule_generated_at for page in matches), default="")

    def readiness_selected(self, _event=None):
        selection = self.readiness_tree.selection()
        if not selection:
            return
        values = self.readiness_tree.item(selection[0], "values")
        if len(values) >= 2:
            self.routing_filter_var.set(values[1])
            self.refresh_routing_tree()

    def sort_routing_tree(self, column: str, reverse: bool):
        rows = [(self.routing_tree.set(item, column), item) for item in self.routing_tree.get_children("")]
        try:
            rows.sort(key=lambda pair: float(str(pair[0]).rstrip("%")) if str(pair[0]).replace(".", "", 1).rstrip("%").isdigit() else str(pair[0]).lower(), reverse=reverse)
        except ValueError:
            rows.sort(key=lambda pair: str(pair[0]).lower(), reverse=reverse)
        for index, (_value, item) in enumerate(rows):
            self.routing_tree.move(item, "", index)
        self.routing_tree.heading(column, command=lambda: self.sort_routing_tree(column, not reverse))

    def routing_double_click(self, event):
        row_id = self.routing_tree.identify_row(event.y)
        column = self.routing_tree.identify_column(event.x)
        if not row_id or not column or row_id not in self.routing_row_map:
            return
        page, index = self.routing_row_map[row_id]
        if index >= len(page.change_rule_draft):
            return
        field_order = [
            "crid", "plant", "alternate", "sequence", "subseq", "old_wc", "new_wc",
            "op_descriptions", "rule_type", "column", "status", "notes",
        ]
        col_index = int(column.lstrip("#")) - 1
        if col_index < 0 or col_index >= len(field_order):
            return
        field = field_order[col_index]
        if field in {"status", "notes"}:
            return
        box = self.routing_tree.bbox(row_id, column)
        if not box:
            return
        x, y, width, height = box
        current = getattr(page.change_rule_draft[index], field)
        values = ["ALL", "WC", "LIST", "SELECT"] if field == "rule_type" else []
        editor = ttk.Combobox(self.routing_tree, values=values, state="normal") if values else ttk.Entry(self.routing_tree)
        if isinstance(editor, ttk.Combobox):
            editor.set(str(current))
        else:
            editor.insert(0, str(current))
        editor.place(x=x, y=y, width=max(width, 80), height=height)
        editor.focus_set()

        def commit(_event=None):
            value = editor.get().strip()
            record = page.change_rule_draft[index]
            try:
                if field in {"sequence", "subseq"}:
                    setattr(record, field, int(value or 0))
                else:
                    setattr(record, field, value.upper() if field in {"crid","plant","old_wc","new_wc","rule_type","column"} else value)
                if field == "sequence" and record.sequence > 0:
                    record.column = f"SEQ{record.sequence}"
            except ValueError:
                messagebox.showwarning("Invalid value", f"{field} must be numeric.")
            editor.destroy()
            file = self.selected_file
            valid = set(self._valid_workcenters(file)) if file else set()
            self._store_rows_on_page(page, validate_change_rule_rows(self._page_to_routing_rows(page), valid))
            self.mark_dirty("Change Rule edits")
            self.refresh_routing_tree()

        editor.bind("<Return>", commit)
        editor.bind("<FocusOut>", commit)
        editor.bind("<Escape>", lambda _event: editor.destroy())

    def generate_material_selection(self):
        rows = [row for item in self.files for page in item.pages for row in self._page_to_routing_rows(page)]
        try:
            self.material_rows, self.material_unmatched = generate_material_selection(rows, self.material_sources)
        except MaterialSelectionError as exc:
            messagebox.showwarning("Material Selection", str(exc))
            return
        for item in self.material_tree.get_children():
            self.material_tree.delete(item)
        for item in self.material_rows:
            self.material_tree.insert("", "end", values=(
                item.material, item.plant, item.routing_group, item.counter, item.variant_as_is,
                item.mrp_controller, item.production_supervisor, item.crid, item.alternate, *item.sequences, item.status, item.notes,
            ))
        for missing in self.material_unmatched:
            plant, crid, old_wc = (missing.split(" | ", 2) + ["", "", ""])[:3]
            self.material_tree.insert("", "end", values=("", plant, "", "", "", "", "", crid, "", "", "", "", "", "No P41 match", old_wc))
        self.material_summary.set(
            f"{len(self.material_rows):,} matched rows | {len(self.material_unmatched):,} source gaps for later rules"
        )
        self.status.configure(text="Material Selection first-round generation complete")
        self.log(f"Material Selection generated: {len(self.material_rows)} matched row(s), {len(self.material_unmatched)} unmatched OLD WC key(s).")
        self.update_action_states()

    def export_material_selection(self):
        if not self.material_rows:
            messagebox.showwarning("Export Material Selection", "Generate Material Selection first.")
            return
        template = filedialog.askopenfilename(
            title="Select Material Selection template",
            filetypes=[("Excel workbook", "*.xlsx *.xlsm")],
        )
        if not template:
            return
        output = filedialog.asksaveasfilename(
            title="Export Material Selection into Template", defaultextension=".xlsx",
            initialfile="CA02_CALENDAR_Material_Selection.xlsx",
            filetypes=[("Excel workbook", "*.xlsx")],
        )
        if not output:
            return
        try:
            path = export_material_selection_template(template, output, self.material_rows)
        except Exception as exc:
            messagebox.showerror("Export Material Selection", str(exc))
            return
        self.status.configure(text="Material Selection template export complete")
        self.log(f"Material Selection template populated: {path}")

    def export_change_rule(self):
        rows = self._draft_to_routing_rows()
        if not rows:
            messagebox.showwarning("Export Change Rule", "Generate a Change Rule draft for the selected tab first.")
            return
        output = filedialog.asksaveasfilename(
            title="Export Selected Change Rule Draft",
            defaultextension=".xlsx",
            initialfile=f"{(self.selected_file.plant if self.selected_file else 'PLANT')}_{(self._page_crid(self.selected_page) if self.selected_page else 'CRID')}_Change_Rule_Draft.xlsx",
            filetypes=[("Excel workbook", "*.xlsx")],
        )
        if not output:
            return
        path = export_change_rule_subset(output, rows)
        self.log(f"Selected Change Rule draft exported: {path}")
        self.status.configure(text="Change Rule export complete")

    def export_plant_routing_package(self):
        if not self.selected_file:
            return
        rows = [row for page in self.selected_file.pages for row in self._page_to_routing_rows(page)]
        if not rows:
            messagebox.showwarning("Export Plant Package", "Generate Change Rule drafts for this Plant first.")
            return
        output = filedialog.asksaveasfilename(
            title="Export Plant Routing Generation Package",
            defaultextension=".xlsx",
            initialfile=f"{self.selected_file.plant}_Routing_Generation_Package.xlsx",
            filetypes=[("Excel workbook", "*.xlsx")],
        )
        if not output:
            return
        path = export_routing_generation_package(output, rows, self._all_history([self.selected_file]))
        self.log(f"Plant routing-generation package exported: {path}")
        self.status.configure(text="Plant routing package export complete")

    def export_all_routing_package(self):
        rows = [row for file in self.files for page in file.pages for row in self._page_to_routing_rows(page)]
        if not rows:
            messagebox.showwarning("Export All Plants", "Generate Change Rule drafts first.")
            return
        output = filedialog.asksaveasfilename(
            title="Export All Plants Routing Generation Package",
            defaultextension=".xlsx",
            initialfile="All_Plants_Routing_Generation_Package.xlsx",
            filetypes=[("Excel workbook", "*.xlsx")],
        )
        if not output:
            return
        path = export_routing_generation_package(output, rows, self._all_history())
        self.log(f"All-plants routing-generation package exported: {path}")
        self.status.configure(text="All-plants routing package export complete")

    def export_change_rule_template(self):
        rows = self._draft_to_routing_rows()
        if not rows:
            messagebox.showwarning("Export into Template", "Generate a Change Rule draft for the selected tab first.")
            return
        if not self.routing_template_path or not Path(self.routing_template_path).exists():
            messagebox.showwarning("Export into Template", "Load the routing-generation template first.")
            return
        output = filedialog.asksaveasfilename(
            title="Export into Routing Generation Template",
            defaultextension=".xlsx",
            initialfile=f"{(self.selected_file.plant if self.selected_file else 'PLANT')}_{(self._page_crid(self.selected_page) if self.selected_page else 'CRID')}_Routing_Generation_Draft.xlsx",
            filetypes=[("Excel workbook", "*.xlsx")],
        )
        if not output:
            return
        sheet = f"{self.selected_file.plant}_{self._page_crid(self.selected_page)}"
        try:
            path = export_into_template(self.routing_template_path, output, rows, target_sheet=sheet)
            self.log(f"Routing template populated: {path}")
            self.status.configure(text="Routing template export complete")
        except Exception as exc:
            messagebox.showerror("Export into Template", str(exc))

    def save_package(self):
        if self.running:
            return
        if not self.files:
            messagebox.showwarning(
                "No files",
                "Add at least one VSDX file before saving an analysis package.",
            )
            return

        output = filedialog.asksaveasfilename(
            title="Save analysis package",
            defaultextension=".vpa",
            initialfile=(
                "VSDX_Analysis_Package_"
                f"{time.strftime('%Y%m%d_%H%M%S')}.vpa"
            ),
            filetypes=[
                ("VSDX Analysis Package", "*.vpa"),
                ("All files", "*.*"),
            ],
        )
        if not output:
            return

        workbook_path = ""
        if self.workbook_mapping:
            candidate = Path(self.workbook_mapping.source_path)
            if candidate.exists() and candidate.is_file():
                workbook_path = str(candidate)

        manual_text = self.department_list_text.get("1.0", tk.END).strip()
        if manual_text == "DEPARTMENT | PLANT":
            manual_text = ""

        try:
            self.status.configure(text="Saving analysis package…")
            self.root.update_idletasks()
            package_path = save_analysis_package(
                output_path=output,
                files=self.files,
                workbook_mapping_path=workbook_path,
                manual_department_plant_text=manual_text,
                routing_template_path=self.routing_template_path,
                logs=self.logs,
                app_version="1.3.8",
            )
            self.status.configure(text="Analysis package saved")
            self.log(f"Analysis package saved: {package_path}")
            messagebox.showinfo(
                "Analysis package saved",
                "The package contains the selected VSDX files, Plant values, "
                "detected page metadata, visual metadata, Department/CRID assignments, "
                "Change Rule drafts/rerun history, mapping inputs and the routing template reference. "
                f"\n\n{package_path}",
            )
        except Exception as exc:
            self.status.configure(text="Package save error")
            messagebox.showerror("Save analysis package", str(exc))

    def import_package(self):
        if self.running:
            return
        package_path = filedialog.askopenfilename(
            title="Import analysis package",
            filetypes=[
                ("VSDX Analysis Package", "*.vpa"),
                ("All files", "*.*"),
            ],
        )
        if not package_path:
            return

        if self.files and not messagebox.askyesno(
            "Replace current workspace",
            "Importing an analysis package replaces the currently selected "
            "VSDX files and assignments. Continue?",
        ):
            return

        try:
            self.status.configure(text="Importing analysis package…")
            self.root.update_idletasks()
            imported = import_analysis_package(package_path)

            workbook_mapping = None
            if imported.workbook_mapping_path:
                workbook_mapping = load_mapping(
                    imported.workbook_mapping_path,
                    self.log,
                )

            manual_mapping = None
            warnings = []
            if imported.manual_department_plant_text.strip():
                manual_mapping, warnings = parse_department_plant_text(
                    imported.manual_department_plant_text,
                    source_name="Imported analysis package",
                )

            self.files = imported.files
            self.selected_file = None
            self.selected_page = None
            self.file_map = {}
            self.page_map = {}

            self.workbook_mapping = workbook_mapping
            self.manual_mapping = manual_mapping
            self.mapping = merge_mappings(
                self.workbook_mapping,
                self.manual_mapping,
            )

            if workbook_mapping:
                self.map_path.set(imported.workbook_mapping_path)
                self.map_info.set(
                    f"Imported package mapping | Sheet: {workbook_mapping.sheet_name} | "
                    f"{len(workbook_mapping.records):,} rows | "
                    f"Plants: {', '.join(workbook_mapping.plants)}"
                )
            else:
                self.map_path.set("")
                self.map_info.set("No Excel mapping was stored in this package.")

            self.department_list_text.delete("1.0", tk.END)
            self.department_list_text.insert(
                "1.0",
                imported.manual_department_plant_text
                or "DEPARTMENT | PLANT\n",
            )
            if manual_mapping:
                self.manual_info.set(
                    f"Imported {manual_mapping.department_plant_pair_count:,} "
                    "Department-Plant pairs from the analysis package."
                    + (
                        f" Skipped {len(warnings)} invalid line(s)."
                        if warnings
                        else ""
                    )
                )
            else:
                self.manual_info.set(
                    "No Department-Plant text list was stored in this package."
                )

            self.rebuild_mapping()
            if imported.routing_template_path:
                self.load_routing_template(imported.routing_template_path)
            self.refresh_files()
            has_metadata = any(
                page.loaded or page.shapes
                for file in self.files
                for page in file.pages
            )
            self.export_btn.configure(
                state=tk.NORMAL if has_metadata else tk.DISABLED
            )
            self.progress["value"] = 100 if has_metadata else 0
            self.progress_text.set("100%" if has_metadata else "0%")
            self.status.configure(
                text=f"Package imported: {len(self.files)} file(s)"
            )
            self.log(
                f"Analysis package imported: {Path(package_path).name} — "
                f"{len(self.files)} VSDX file(s), "
                f"{sum(len(file.pages) for file in self.files)} tabs/pages."
            )
        except Exception as exc:
            self.status.configure(text="Package import error")
            messagebox.showerror("Import analysis package", str(exc))

    def run_analysis(self):
        if self.running:
            return
        if not self.files:
            messagebox.showwarning(
                "No files", "Add at least one VSDX file."
            )
            return

        missing = [
            file.file_name
            for file in self.files
            if not file.plant
        ]
        if missing:
            messagebox.showwarning(
                "Plant assignment required",
                "Assign a plant code to every file:\n\n"
                + "\n".join(missing),
            )
            return

        self.running = True
        self.started = time.time()
        self.progress["value"] = 0
        self.progress_text.set("0%")
        self.run_btn.configure(state=tk.DISABLED)
        self.export_btn.configure(state=tk.DISABLED)
        self.update_action_states()
        self.status.configure(text="Analyzing…")

        total = sum(len(file.pages) for file in self.files)
        self.log(
            f"Analysis started: {len(self.files)} file(s), "
            f"{total} tabs/pages."
        )

        def worker():
            completed = 0
            try:
                for file in self.files:
                    self.events.put(
                        ("file", file.path, "Processing")
                    )
                    old_assignments = {
                        page.page_id: (
                            page.assigned_department,
                            page.assignment_mode,
                        )
                        for page in file.pages
                    }

                    def page_callback(_index, _total, _page):
                        nonlocal completed
                        completed += 1
                        percent = (
                            completed / total * 100
                            if total
                            else 100
                        )
                        self.events.put(
                            (
                                "progress",
                                percent,
                                f"{completed}/{total} tabs",
                            )
                        )

                    analyzed = analyze(
                        file.path,
                        file.plant,
                        self.mapping,
                        lambda line: self.events.put(
                            ("log", line)
                        ),
                        page_callback,
                        pages=file.pages,
                    )

                    for page in analyzed.pages:
                        saved = old_assignments.get(page.page_id)
                        if saved and saved[1] in ("manual", "cleared"):
                            page.assigned_department = saved[0]
                            page.assignment_mode = saved[1]

                    file.pages = analyzed.pages
                    file.status = analyzed.status
                    file.error = analyzed.error
                    self.events.put(
                        ("file", file.path, analyzed.status)
                    )

                self.events.put(("done",))
            except Exception as exc:
                self.events.put(("error", str(exc)))

        threading.Thread(
            target=worker, daemon=True
        ).start()

    def export_results(self):
        if self.running:
            return
        if not any(
            page.shapes
            for file in self.files
            for page in file.pages
        ):
            messagebox.showwarning(
                "No analysis",
                "Run metadata analysis first.",
            )
            return

        output = filedialog.asksaveasfilename(
            title="Export structured Excel",
            defaultextension=".xlsx",
            initialfile=(
                "VSDX_Process_Metadata_"
                f"{time.strftime('%Y%m%d_%H%M%S')}.xlsx"
            ),
            filetypes=[("Excel workbook", "*.xlsx")],
        )
        if not output:
            return

        try:
            self.status.configure(text="Exporting…")
            self.root.update_idletasks()
            path = export(
                output,
                self.files,
                self.mapping,
                self.logs,
            )
            self.log(f"Excel exported: {path}")
            self.status.configure(text="Export complete")
            if messagebox.askyesno(
                "Export complete",
                f"Saved:\n{path}\n\nOpen containing folder?",
            ):
                self.open_folder(Path(path).parent)
        except Exception as exc:
            self.status.configure(text="Export error")
            messagebox.showerror(
                "Export failed", str(exc)
            )

    def open_folder(self, path: Path):
        try:
            if sys.platform.startswith("win"):
                os.startfile(str(path))
            elif sys.platform == "darwin":
                subprocess.Popen(["open", str(path)])
            else:
                subprocess.Popen(["xdg-open", str(path)])
        except Exception:
            pass

    def set_zoom(self, zoom: float):
        self.zoom = zoom
        if hasattr(self, "zoom_var"):
            self.zoom_var.set(f"{zoom * 100:.0f}%")
        self.render_preview()

    def render_preview(self):
        if not self.selected_page:
            return

        try:
            image = render(
                self.selected_page,
                max(700, self.canvas.winfo_width() - 8),
                max(420, self.canvas.winfo_height() - 8),
                self.zoom,
            )
            self.show_image(image)
            if self.selected_page.loading:
                self.preview_mode.set("Loading selected tab from VSDX XML…")
            elif self.selected_page.loaded:
                self.preview_mode.set("Visual preview ready")
            else:
                self.preview_mode.set("Waiting to load selected tab")
        except Exception as exc:
            self.preview_mode.set(
                f"Preview error: {exc}"
            )

    def show_image(self, image):
        self.preview_photo = ImageTk.PhotoImage(image)
        self.canvas.delete("all")
        self.canvas.create_image(
            0,
            0,
            image=self.preview_photo,
            anchor="nw",
        )
        self.canvas.configure(
            scrollregion=(0, 0, image.width, image.height)
        )

    def exact_preview(self):
        if not self.selected_file or not self.selected_page:
            return
        if (
            self.exact_thread
            and self.exact_thread.is_alive()
        ):
            return

        file = self.selected_file
        page = self.selected_page
        self.preview_mode.set(
            "Rendering through Microsoft Visio…"
        )
        self.exact_btn.configure(state=tk.DISABLED)

        def worker():
            try:
                self.events.put(
                    (
                        "exact",
                        str(VisioRenderer.export(file.path, page)),
                        page.page_id,
                    )
                )
            except Exception as exc:
                self.events.put(
                    ("exact_error", str(exc))
                )

        self.exact_thread = threading.Thread(
            target=worker, daemon=True
        )
        self.exact_thread.start()

    def show_page_meta(self):
        page = self.selected_page
        if not page:
            self.show_meta("")
            return

        available = (
            self.mapping.departments_for_plant(
                self.selected_file.plant
            )
            if self.mapping and self.selected_file
            else []
        )
        lines = [
            f"Tab name       : {page.name}",
            f"Internal name  : {page.name_u}",
            f"Page ID        : {page.page_id}",
            (
                "Preview status : Loading"
                if page.loading
                else (
                    "Preview status : Ready"
                    if page.loaded
                    else "Preview status : Not loaded"
                )
            ),
            (
                f"Page size      : "
                f"{page.width:.2f} × {page.height:.2f}"
            ),
            f"Shapes         : {len(page.shapes):,}",
            f"Text shapes    : {page.text_shape_count:,}",
            f"Connections    : {len(page.connections):,}",
            (
                f"Workcenters    : "
                f"{len(page.workcenter_counts):,}"
            ),
            (
                f"Suggested      : "
                f"{page.suggested_department or '-'}"
            ),
            (
                f"Suggestion via : "
                f"{page.suggestion_source or '-'}"
            ),
            (
                f"Confidence     : {page.confidence:.1%}"
                if page.candidate_counts
                else "Confidence     : -"
            ),
            (
                f"Assigned       : "
                f"{page.assigned_department or '-'}"
            ),
            f"Assignment mode: {page.assignment_mode}",
            (
                f"Plant options  : "
                f"{len(available):,} departments/CRIDs"
            ),
            "",
            "Detected workcenters:",
            (
                ", ".join(page.workcenter_counts)
                if page.workcenter_counts
                else (
                    "(Loading selected tab…)"
                    if page.loading
                    else "(Load the visual preview or run full analysis.)"
                )
            ),
            "",
            "Unmatched workcenters:",
            (
                ", ".join(page.unmatched_wcs)
                if page.unmatched_wcs
                else "-"
            ),
        ]
        self.show_meta("\n".join(lines))

    def show_meta(self, text: str):
        self.meta.configure(state=tk.NORMAL)
        self.meta.delete("1.0", tk.END)
        self.meta.insert("1.0", text)
        self.meta.configure(state=tk.DISABLED)

    def show_candidates(self):
        self.candidates.delete(
            *self.candidates.get_children()
        )
        page = self.selected_page
        if not page:
            return

        inserted = set()
        for index, (department, count) in enumerate(
            page.candidate_counts.items()
        ):
            workcenters = ", ".join(
                sorted(
                    page.candidate_wcs.get(
                        department, set()
                    )
                )
            )
            self.candidates.insert(
                "",
                tk.END,
                iid=f"matched-{index}",
                values=(
                    department,
                    count,
                    workcenters or "Workcenter mapping",
                ),
            )
            inserted.add(department)

        if self.mapping and self.selected_file:
            available = self.mapping.departments_for_plant(
                self.selected_file.plant
            )
            next_index = len(inserted)
            for department in available:
                if department in inserted:
                    continue
                sources = self.mapping.sources_for_department(
                    self.selected_file.plant,
                    department,
                )
                self.candidates.insert(
                    "",
                    tk.END,
                    iid=f"available-{next_index}",
                    values=(
                        department,
                        "",
                        ", ".join(sources)
                        or "Available for plant",
                    ),
                )
                next_index += 1

    def poll(self):
        try:
            while True:
                event = self.events.get_nowait()
                kind = event[0]

                if kind == "log":
                    self.log(event[1])

                elif kind == "progress":
                    self.progress["value"] = event[1]
                    self.progress_text.set(
                        f"{event[1]:.0f}%"
                    )
                    self.status.configure(
                        text="Analyzing " + event[2]
                    )

                elif kind == "file":
                    for file in self.files:
                        if file.path == event[1]:
                            file.status = event[2]
                            break
                    self.refresh_files()

                elif kind == "p41_loaded":
                    loaded, path, chosen = event[1], event[2], event[3]
                    previous = self.material_sources
                    loaded.code_reference_path = previous.code_reference_path
                    loaded.controller_sheets = previous.controller_sheets
                    loaded.valid_mrp_codes = previous.valid_mrp_codes
                    loaded.valid_supervisor_codes = previous.valid_supervisor_codes
                    self.material_sources = loaded
                    self.step0_loading = False
                    self.p41_source_info.set(f"Ready — {len(loaded.variants):,} routing variants and {len(loaded.workcenters):,} Plant/WorkCenter links loaded.")
                    self.status.configure(text="P41 source confirmed")
                    self.log(f"Step 0 P41 confirmed: {Path(path).name} | tabs: {', '.join(chosen)}")
                    self.refresh_step0_state()

                elif kind == "p41_error":
                    self.step0_loading = False
                    self.p41_source_info.set("Validation failed — correct the file or tab selection.")
                    self.status.configure(text="P41 validation failed")
                    messagebox.showerror("P41 source", event[1])
                    self.refresh_step0_state()

                elif kind == "done":
                    self.running = False
                    self.progress["value"] = 100
                    self.progress_text.set("100%")
                    self.run_btn.configure(state=tk.NORMAL)
                    self.export_btn.configure(state=tk.NORMAL)
                    self.status.configure(
                        text="Analysis complete"
                    )
                    self.log(
                        "Analysis completed successfully."
                    )
                    self.refresh_files()
                    self.refresh_pages()
                    self.update_action_states()

                elif kind == "error":
                    self.running = False
                    self.run_btn.configure(state=tk.NORMAL)
                    self.status.configure(
                        text="Analysis error"
                    )
                    self.update_action_states()
                    self.log("ERROR: " + event[1])
                    messagebox.showerror(
                        "Analysis failed", event[1]
                    )

                elif kind == "preview_loaded":
                    file_path, page_id, token, seconds = (
                        event[1], event[2], event[3], event[4]
                    )
                    page = next(
                        (
                            candidate
                            for file in self.files
                            if file.path == file_path
                            for candidate in file.pages
                            if candidate.page_id == page_id
                        ),
                        None,
                    )
                    if page:
                        page.loading = False
                    self.log(
                        f"Visual preview ready: {page.name if page else page_id} "
                        f"in {seconds:.2f}s."
                    )
                    if (
                        token == self.preview_request_token
                        and self.selected_file
                        and self.selected_file.path == file_path
                        and self.selected_page
                        and self.selected_page.page_id == page_id
                    ):
                        self.status.configure(text="Preview ready")
                        self.refresh_pages()
                        self.show_page_meta()
                        self.show_candidates()
                        self.render_preview()

                elif kind == "preview_error":
                    file_path, page_id, token, error = (
                        event[1], event[2], event[3], event[4]
                    )
                    page = next(
                        (
                            candidate
                            for file in self.files
                            if file.path == file_path
                            for candidate in file.pages
                            if candidate.page_id == page_id
                        ),
                        None,
                    )
                    if page:
                        page.loading = False
                        page.load_error = error
                    self.log(f"Visual preview error: {error}")
                    if token == self.preview_request_token:
                        self.status.configure(text="Preview error")
                        self.preview_mode.set("Visual preview could not be loaded")
                        self.refresh_pages()
                        self.show_page_meta()
                        self.render_preview()

                elif kind == "exact":
                    if (
                        self.selected_page
                        and self.selected_page.page_id
                        == event[2]
                    ):
                        self.show_image(
                            Image.open(event[1]).convert("RGB")
                        )
                        self.preview_mode.set(
                            "Exact preview exported by "
                            "Microsoft Visio"
                        )
                    self.exact_btn.configure(state=tk.NORMAL)

                elif kind == "exact_error":
                    self.exact_btn.configure(
                        state=(
                            tk.NORMAL
                            if VisioRenderer.supported()
                            else tk.DISABLED
                        )
                    )
                    self.preview_mode.set(
                        "XML vector preview"
                    )
                    self.log(
                        "Exact preview unavailable: "
                        + event[1]
                    )

        except queue.Empty:
            pass

        self.root.after(100, self.poll)

    def elapsed(self):
        if self.running:
            seconds = int(time.time() - self.started)
            self.elapsed_text.set(
                f"Elapsed {seconds // 60:02d}:"
                f"{seconds % 60:02d}"
            )
        self.root.after(500, self.elapsed)

    def log(self, message: str):
        line = f"[{time.strftime('%H:%M:%S')}] {message}"
        self.logs.append(line)
        self.logbox.configure(state=tk.NORMAL)
        self.logbox.insert(tk.END, line + "\n")
        self.logbox.see(tk.END)
        self.logbox.configure(state=tk.DISABLED)

    def run(self):
        self.root.mainloop()


def launch():
    App().run()
