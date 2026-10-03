import os
import subprocess
import re
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from datetime import datetime

import pandas as pd
import pymupdf
from PIL import Image, ImageTk


# ============================================================
# APPLICATION CONFIGURATION
# ============================================================

APP_NAME = "KAMA JOBBAG PROCESSOR"
APP_SUBTITLE = "Document Processing & SO / Production Order Management"

BG = "#F5F7FB"
CARD = "#FFFFFF"

NAVY = "#172B4D"
NAVY_DARK = "#0F1F35"

BLUE = "#2563EB"
BLUE_DARK = "#1D4ED8"

GREEN = "#16A34A"
GREEN_DARK = "#15803D"

ORANGE = "#F59E0B"
ORANGE_DARK = "#D97706"

RED = "#DC2626"
RED_LIGHT = "#FEE2E2"

PURPLE = "#7C3AED"
PURPLE_DARK = "#6D28D9"

TEXT = "#1E293B"
MUTED = "#64748B"
BORDER = "#E2E8F0"

LIGHT_BLUE = "#EFF6FF"
LIGHT_GREEN = "#F0FDF4"
LIGHT_PURPLE = "#F5F3FF"
LIGHT_ORANGE = "#FFFBEB"

FONT = "Segoe UI"


# ============================================================
# MAIN APPLICATION
# ============================================================

class JobbagApplication:

    def __init__(self, root):

        self.root = root

        self.root.title(APP_NAME)

        self.root.geometry("1400x850")
        self.root.minsize(1150, 700)

        self.root.configure(
            bg=BG
        )

        # ----------------------------------------------------
        # DATA
        # ----------------------------------------------------

        self.pdf_files = []
        self.excel_file = ""
        self.results = []
        self.output_folder = ""
        self.a4_2up_path = ""

        # ----------------------------------------------------
        # PREVIEW
        # ----------------------------------------------------

        self.preview_window = None
        self.preview_doc = None
        self.preview_page = None
        self.preview_image = None
        self.preview_zoom = 1.0

        # ----------------------------------------------------
        # VARIABLES
        # ----------------------------------------------------

        self.total_pdf_var = tk.StringVar(
            value="0"
        )

        self.total_jobbag_var = tk.StringVar(
            value="0"
        )

        self.ready_var = tk.StringVar(
            value="0"
        )

        self.missing_var = tk.StringVar(
            value="0"
        )

        self.page_title_var = tk.StringVar(
            value="Dashboard"
        )

        # ----------------------------------------------------
        # STYLE
        # ----------------------------------------------------

        self.setup_styles()

        # ----------------------------------------------------
        # BUILD UI
        # ----------------------------------------------------

        self.build_layout()

        self.show_dashboard()

    # ========================================================
    # STYLES
    # ========================================================

    def setup_styles(self):

        style = ttk.Style()

        try:
            style.theme_use(
                "clam"
            )
        except Exception:
            pass

        style.configure(
            "Treeview",
            background=CARD,
            foreground=TEXT,
            fieldbackground=CARD,
            rowheight=38,
            font=(FONT, 10)
        )

        style.configure(
            "Treeview.Heading",
            background=NAVY,
            foreground="white",
            font=(FONT, 10, "bold"),
            padding=10
        )

        style.map(
            "Treeview",
            background=[
                ("selected", "#DBEAFE")
            ],
            foreground=[
                ("selected", NAVY)
            ]
        )

        style.configure(
            "Horizontal.TProgressbar",
            troughcolor=BORDER,
            background=BLUE,
            bordercolor=BORDER,
            lightcolor=BLUE,
            darkcolor=BLUE
        )

    # ========================================================
    # MAIN LAYOUT
    # ========================================================

    def build_layout(self):

        # ====================================================
        # LEFT SIDEBAR
        # ====================================================

        self.sidebar = tk.Frame(
            self.root,
            bg=NAVY_DARK,
            width=245
        )

        self.sidebar.pack(
            side="left",
            fill="y"
        )

        self.sidebar.pack_propagate(
            False
        )

        # ----------------------------------------------------
        # Logo
        # ----------------------------------------------------

        logo_frame = tk.Frame(
            self.sidebar,
            bg=NAVY_DARK,
            height=100
        )

        logo_frame.pack(
            fill="x"
        )

        logo_frame.pack_propagate(
            False
        )

        tk.Label(
            logo_frame,
            text="KAMA",
            font=(FONT, 24, "bold"),
            fg="white",
            bg=NAVY_DARK
        ).pack(
            anchor="w",
            padx=24,
            pady=(20, 0)
        )

        tk.Label(
            logo_frame,
            text="JOBBAG PROCESSOR",
            font=(FONT, 9, "bold"),
            fg="#93C5FD",
            bg=NAVY_DARK
        ).pack(
            anchor="w",
            padx=25,
            pady=(2, 0)
        )

        # ----------------------------------------------------
        # Navigation label
        # ----------------------------------------------------

        tk.Label(
            self.sidebar,
            text="WORKSPACE",
            font=(FONT, 8, "bold"),
            fg="#64748B",
            bg=NAVY_DARK
        ).pack(
            anchor="w",
            padx=25,
            pady=(18, 8)
        )

        # ----------------------------------------------------
        # Navigation buttons
        # ----------------------------------------------------

        self.nav_buttons = {}

        self.nav_buttons["dashboard"] = self.create_nav_button(
            "▦",
            "Dashboard",
            self.show_dashboard
        )

        self.nav_buttons["process"] = self.create_nav_button(
            "▣",
            "Process Documents",
            self.show_process
        )

        self.nav_buttons["results"] = self.create_nav_button(
            "☷",
            "Results",
            self.show_results
        )

        # ----------------------------------------------------
        # Divider
        # ----------------------------------------------------

        tk.Frame(
            self.sidebar,
            bg="#243B5A",
            height=1
        ).pack(
            fill="x",
            padx=22,
            pady=20
        )

        tk.Label(
            self.sidebar,
            text="SYSTEM",
            font=(FONT, 8, "bold"),
            fg="#64748B",
            bg=NAVY_DARK
        ).pack(
            anchor="w",
            padx=25,
            pady=(0, 8)
        )

        self.nav_buttons["settings"] = self.create_nav_button(
            "⚙",
            "Settings",
            self.show_settings
        )

        self.nav_buttons["about"] = self.create_nav_button(
            "ⓘ",
            "About",
            self.show_about
        )

        # ----------------------------------------------------
        # Sidebar bottom
        # ----------------------------------------------------

        sidebar_bottom = tk.Frame(
            self.sidebar,
            bg=NAVY_DARK
        )

        sidebar_bottom.pack(
            side="bottom",
            fill="x",
            padx=20,
            pady=20
        )

        tk.Label(
            sidebar_bottom,
            text="SYSTEM STATUS",
            font=(FONT, 8, "bold"),
            fg="#64748B",
            bg=NAVY_DARK
        ).pack(
            anchor="w"
        )

        status_row = tk.Frame(
            sidebar_bottom,
            bg=NAVY_DARK
        )

        status_row.pack(
            anchor="w",
            pady=(7, 0)
        )

        tk.Label(
            status_row,
            text="●",
            font=(FONT, 12),
            fg="#22C55E",
            bg=NAVY_DARK
        ).pack(
            side="left"
        )

        tk.Label(
            status_row,
            text="Ready",
            font=(FONT, 9),
            fg="#CBD5E1",
            bg=NAVY_DARK
        ).pack(
            side="left",
            padx=5
        )

        # ====================================================
        # RIGHT AREA
        # ====================================================

        self.main_area = tk.Frame(
            self.root,
            bg=BG
        )

        self.main_area.pack(
            side="right",
            fill="both",
            expand=True
        )

        # ----------------------------------------------------
        # TOP HEADER
        # ----------------------------------------------------

        self.topbar = tk.Frame(
            self.main_area,
            bg=CARD,
            height=78,
            highlightbackground=BORDER,
            highlightthickness=1
        )

        self.topbar.pack(
            fill="x"
        )

        self.topbar.pack_propagate(
            False
        )

        tk.Label(
            self.topbar,
            textvariable=self.page_title_var,
            font=(FONT, 20, "bold"),
            fg=NAVY,
            bg=CARD
        ).pack(
            side="left",
            padx=28
        )

        self.header_status = tk.Label(
            self.topbar,
            text="● READY",
            font=(FONT, 9, "bold"),
            fg=GREEN,
            bg=LIGHT_GREEN,
            padx=14,
            pady=7
        )

        self.header_status.pack(
            side="right",
            padx=28
        )

        # ----------------------------------------------------
        # CONTENT
        # ----------------------------------------------------

        self.content = tk.Frame(
            self.main_area,
            bg=BG
        )

        self.content.pack(
            fill="both",
            expand=True,
            padx=28,
            pady=25
        )

    # ========================================================
    # NAV BUTTON
    # ========================================================

    def create_nav_button(
        self,
        icon,
        text,
        command
    ):

        frame = tk.Frame(
            self.sidebar,
            bg=NAVY_DARK,
            height=48
        )

        frame.pack(
            fill="x",
            padx=12,
            pady=3
        )

        frame.pack_propagate(
            False
        )

        icon_label = tk.Label(
            frame,
            text=icon,
            font=(FONT, 16),
            fg="#94A3B8",
            bg=NAVY_DARK,
            width=3
        )

        icon_label.pack(
            side="left"
        )

        text_label = tk.Label(
            frame,
            text=text,
            font=(FONT, 10),
            fg="#CBD5E1",
            bg=NAVY_DARK,
            anchor="w"
        )

        text_label.pack(
            side="left",
            fill="x",
            expand=True
        )

        for widget in (
            frame,
            icon_label,
            text_label
        ):

            widget.bind(
                "<Button-1>",
                lambda e: command()
            )

        return frame

    # ========================================================
    # CLEAR CONTENT
    # ========================================================

    def clear_content(self):

        for widget in self.content.winfo_children():
            widget.destroy()

    # ========================================================
    # SET ACTIVE NAV
    # ========================================================

    def set_active_nav(
        self,
        name
    ):

        for key, frame in self.nav_buttons.items():

            if key == name:

                frame.configure(
                    bg="#1E3A5F"
                )

                for child in frame.winfo_children():

                    child.configure(
                        bg="#1E3A5F",
                        fg="white"
                    )

            else:

                frame.configure(
                    bg=NAVY_DARK
                )

                for child in frame.winfo_children():

                    child.configure(
                        bg=NAVY_DARK,
                        fg="#CBD5E1"
                    )

    # ========================================================
    # DASHBOARD
    # ========================================================

    def show_dashboard(self):

        self.page_title_var.set(
            "SALES ORDER PROCESSING"
        )

        self.set_active_nav(
            "dashboard"
        )

        self.clear_content()

        # ----------------------------------------------------
        # Welcome
        # ----------------------------------------------------

        welcome = tk.Frame(
            self.content,
            bg=BG
        )

        welcome.pack(
            fill="x"
        )

        tk.Label(
            welcome,
            text="Jobbag Processing Overview",
            font=(FONT, 16, "bold"),
            fg=TEXT,
            bg=BG
        ).pack(
            anchor="w"
        )

        tk.Label(
            welcome,
            text="Process PDF jobbags, map Sales Orders and generate ready-to-print documents.",
            font=(FONT, 10),
            fg=MUTED,
            bg=BG
        ).pack(
            anchor="w",
            pady=(4, 20)
        )

        # ----------------------------------------------------
        # KPI CARDS
        # ----------------------------------------------------

        cards = tk.Frame(
            self.content,
            bg=BG
        )

        cards.pack(
            fill="x"
        )

        self.create_kpi_card(
            cards,
            "PDF FILES",
            self.total_pdf_var,
            "Documents selected",
            BLUE,
            LIGHT_BLUE
        ).pack(
            side="left",
            fill="both",
            expand=True,
            padx=(0, 10)
        )

        self.create_kpi_card(
            cards,
            "JOBBAGS",
            self.total_jobbag_var,
            "Jobbags processed",
            PURPLE,
            LIGHT_PURPLE
        ).pack(
            side="left",
            fill="both",
            expand=True,
            padx=10
        )

        self.create_kpi_card(
            cards,
            "READY",
            self.ready_var,
            "SO & PO successfully mapped",
            GREEN,
            LIGHT_GREEN
        ).pack(
            side="left",
            fill="both",
            expand=True,
            padx=10
        )

        self.create_kpi_card(
            cards,
            "SO / PO MISSING",
            self.missing_var,
            "Requires attention",
            RED,
            RED_LIGHT
        ).pack(
            side="left",
            fill="both",
            expand=True,
            padx=(10, 0)
        )

        # ----------------------------------------------------
        # MAIN ACTION CARD
        # ----------------------------------------------------

        action_card = tk.Frame(
            self.content,
            bg=CARD,
            highlightbackground=BORDER,
            highlightthickness=1
        )

        action_card.pack(
            fill="x",
            pady=25
        )

        tk.Label(
            action_card,
            text="START A NEW PROCESS",
            font=(FONT, 11, "bold"),
            fg=NAVY,
            bg=CARD
        ).pack(
            anchor="w",
            padx=25,
            pady=(22, 4)
        )

        tk.Label(
            action_card,
            text="Select your PDF documents and Excel Jobbag/SO mapping file.",
            font=(FONT, 10),
            fg=MUTED,
            bg=CARD
        ).pack(
            anchor="w",
            padx=25
        )

        buttons = tk.Frame(
            action_card,
            bg=CARD
        )

        buttons.pack(
            anchor="w",
            padx=25,
            pady=22
        )

        self.create_colored_button(
            buttons,
            "▣  PROCESS DOCUMENTS",
            GREEN,
            GREEN_DARK,
            self.show_process,
            25
        ).pack(
            side="left"
        )

        self.create_colored_button(
            buttons,
            "☷  VIEW RESULTS",
            BLUE,
            BLUE_DARK,
            self.show_results,
            18
        ).pack(
            side="left",
            padx=10
        )

        # ----------------------------------------------------
        # INFORMATION CARD
        # ----------------------------------------------------

        info_card = tk.Frame(
            self.content,
            bg=NAVY
        )

        info_card.pack(
            fill="x"
        )

        tk.Label(
            info_card,
            text="PROCESS FLOW",
            font=(FONT, 10, "bold"),
            fg="#93C5FD",
            bg=NAVY
        ).pack(
            anchor="w",
            padx=25,
            pady=(18, 4)
        )

        tk.Label(
            info_card,
            text="PDF  →  Jobbag Detection  →  Excel SO Mapping  →  Individual PDF  →  Preview / Print",
            font=(FONT, 11, "bold"),
            fg="white",
            bg=NAVY
        ).pack(
            anchor="w",
            padx=25,
            pady=(0, 18)
        )

    # ========================================================
    # KPI CARD
    # ========================================================

    def create_kpi_card(
        self,
        parent,
        title,
        variable,
        subtitle,
        color,
        light_color
    ):

        card = tk.Frame(
            parent,
            bg=CARD,
            height=135,
            highlightbackground=BORDER,
            highlightthickness=1
        )

        card.pack_propagate(
            False
        )

        top = tk.Frame(
            card,
            bg=CARD
        )

        top.pack(
            fill="x",
            padx=18,
            pady=(17, 0)
        )

        tk.Label(
            top,
            text=title,
            font=(FONT, 9, "bold"),
            fg=MUTED,
            bg=CARD
        ).pack(
            side="left"
        )

        tk.Label(
            card,
            textvariable=variable,
            font=(FONT, 27, "bold"),
            fg=color,
            bg=CARD
        ).pack(
            anchor="w",
            padx=18,
            pady=(7, 0)
        )

        tk.Label(
            card,
            text=subtitle,
            font=(FONT, 9),
            fg=MUTED,
            bg=CARD
        ).pack(
            anchor="w",
            padx=18
        )

        return card

    # ========================================================
    # PROCESS PAGE
    # ========================================================

    def show_process(self):

        self.page_title_var.set(
            "Process Documents"
        )

        self.set_active_nav(
            "process"
        )

        self.clear_content()

        # ----------------------------------------------------
        # Description
        # ----------------------------------------------------

        tk.Label(
            self.content,
            text="Process PDF Documents",
            font=(FONT, 17, "bold"),
            fg=TEXT,
            bg=BG
        ).pack(
            anchor="w"
        )

        tk.Label(
            self.content,
            text="Select the source PDFs and Excel mapping file to generate individual Jobbag documents.",
            font=(FONT, 10),
            fg=MUTED,
            bg=BG
        ).pack(
            anchor="w",
            pady=(4, 18)
        )

        # ----------------------------------------------------
        # FILE CARDS
        # ----------------------------------------------------

        files_container = tk.Frame(
            self.content,
            bg=BG
        )

        files_container.pack(
            fill="x"
        )

        # PDF card
        pdf_card = tk.Frame(
            files_container,
            bg=CARD,
            highlightbackground=BORDER,
            highlightthickness=1
        )

        pdf_card.pack(
            side="left",
            fill="both",
            expand=True,
            padx=(0, 10)
        )

        tk.Label(
            pdf_card,
            text="PDF DOCUMENTS",
            font=(FONT, 10, "bold"),
            fg=NAVY,
            bg=CARD
        ).pack(
            anchor="w",
            padx=22,
            pady=(20, 5)
        )

        tk.Label(
            pdf_card,
            text="Select one or multiple PDF files.",
            font=(FONT, 9),
            fg=MUTED,
            bg=CARD
        ).pack(
            anchor="w",
            padx=22
        )

        self.pdf_selection_label = tk.Label(
            pdf_card,
            text="No PDF files selected",
            font=(FONT, 10),
            fg=TEXT,
            bg="#F8FAFC",
            anchor="w",
            padx=12
        )

        self.pdf_selection_label.pack(
            fill="x",
            padx=22,
            pady=15,
            ipady=8
        )

        self.create_colored_button(
            pdf_card,
            "＋  SELECT PDF FILES",
            BLUE,
            BLUE_DARK,
            self.select_pdfs,
            22
        ).pack(
            anchor="w",
            padx=22,
            pady=(0, 22)
        )

        # Excel card
        excel_card = tk.Frame(
            files_container,
            bg=CARD,
            highlightbackground=BORDER,
            highlightthickness=1
        )

        excel_card.pack(
            side="left",
            fill="both",
            expand=True,
            padx=(10, 0)
        )

        tk.Label(
            excel_card,
            text="EXCEL MAPPING",
            font=(FONT, 10, "bold"),
            fg=NAVY,
            bg=CARD
        ).pack(
            anchor="w",
            padx=22,
            pady=(20, 5)
        )

        tk.Label(
            excel_card,
            text="Required columns: JOBBAG_NUMBER, SO_NUMBER and PRODUCTION_ORDER.",
            font=(FONT, 9),
            fg=MUTED,
            bg=CARD
        ).pack(
            anchor="w",
            padx=22
        )

        self.excel_selection_label = tk.Label(
            excel_card,
            text="No Excel file selected",
            font=(FONT, 10),
            fg=TEXT,
            bg="#F8FAFC",
            anchor="w",
            padx=12
        )

        self.excel_selection_label.pack(
            fill="x",
            padx=22,
            pady=15,
            ipady=8
        )

        self.create_colored_button(
            excel_card,
            "＋  SELECT EXCEL FILE",
            PURPLE,
            PURPLE_DARK,
            self.select_excel,
            22
        ).pack(
            anchor="w",
            padx=22,
            pady=(0, 22)
        )

        # ----------------------------------------------------
        # PROCESS CARD
        # ----------------------------------------------------

        process_card = tk.Frame(
            self.content,
            bg=NAVY
        )

        process_card.pack(
            fill="x",
            pady=22
        )

        tk.Label(
            process_card,
            text="READY TO PROCESS",
            font=(FONT, 10, "bold"),
            fg="#93C5FD",
            bg=NAVY
        ).pack(
            anchor="w",
            padx=25,
            pady=(18, 3)
        )

        tk.Label(
            process_card,
            text="Generate individual Jobbag PDFs with their mapped Sales Order numbers.",
            font=(FONT, 11),
            fg="white",
            bg=NAVY
        ).pack(
            anchor="w",
            padx=25
        )

        self.process_main_button = self.create_colored_button(
            process_card,
            "▶  PROCESS DOCUMENTS",
            GREEN,
            GREEN_DARK,
            self.process_files,
            26
        )

        self.process_main_button.pack(
            anchor="w",
            padx=25,
            pady=18
        )

        # ----------------------------------------------------
        # PROGRESS
        # ----------------------------------------------------

        progress_card = tk.Frame(
            self.content,
            bg=CARD,
            highlightbackground=BORDER,
            highlightthickness=1
        )

        progress_card.pack(
            fill="x"
        )

        progress_top = tk.Frame(
            progress_card,
            bg=CARD
        )

        progress_top.pack(
            fill="x",
            padx=20,
            pady=(15, 7)
        )

        tk.Label(
            progress_top,
            text="PROCESSING STATUS",
            font=(FONT, 9, "bold"),
            fg=MUTED,
            bg=CARD
        ).pack(
            side="left"
        )

        self.progress_label = tk.Label(
            progress_top,
            text="Waiting for files...",
            font=(FONT, 9),
            fg=TEXT,
            bg=CARD
        )

        self.progress_label.pack(
            side="right"
        )

        self.progress = ttk.Progressbar(
            progress_card,
            orient="horizontal",
            mode="determinate",
            style="Horizontal.TProgressbar"
        )

        self.progress.pack(
            fill="x",
            padx=20,
            pady=(0, 16)
        )

    # ========================================================
    # RESULTS PAGE
    # ========================================================

    def show_results(self):

        self.page_title_var.set(
            "Results"
        )

        self.set_active_nav(
            "results"
        )

        self.clear_content()

        # ----------------------------------------------------
        # Header
        # ----------------------------------------------------

        heading = tk.Frame(
            self.content,
            bg=BG
        )

        heading.pack(
            fill="x"
        )

        tk.Label(
            heading,
            text="Processing Results",
            font=(FONT, 17, "bold"),
            fg=TEXT,
            bg=BG
        ).pack(
            side="left"
        )

        self.result_summary_label = tk.Label(
            heading,
            text=f"{len(self.results)} records",
            font=(FONT, 9, "bold"),
            fg=MUTED,
            bg=BG
        )

        self.result_summary_label.pack(
            side="right"
        )

        # ----------------------------------------------------
        # Summary Cards
        # ----------------------------------------------------

        summary = tk.Frame(
            self.content,
            bg=BG
        )

        summary.pack(
            fill="x",
            pady=(15, 18)
        )

        self.create_small_stat(
            summary,
            "Total Jobbags",
            self.total_jobbag_var,
            BLUE
        ).pack(
            side="left",
            fill="both",
            expand=True,
            padx=(0, 8)
        )

        self.create_small_stat(
            summary,
            "Ready",
            self.ready_var,
            GREEN
        ).pack(
            side="left",
            fill="both",
            expand=True,
            padx=8
        )

        self.create_small_stat(
            summary,
            "SO Missing",
            self.missing_var,
            RED
        ).pack(
            side="left",
            fill="both",
            expand=True,
            padx=8
        )

        # ----------------------------------------------------
        # Table card
        # ----------------------------------------------------

        table_card = tk.Frame(
            self.content,
            bg=CARD,
            highlightbackground=BORDER,
            highlightthickness=1
        )

        table_card.pack(
            fill="both",
            expand=True
        )

        table_top = tk.Frame(
            table_card,
            bg=CARD
        )

        table_top.pack(
            fill="x",
            padx=20,
            pady=(18, 10)
        )

        tk.Label(
            table_top,
            text="JOBBAG DETAILS",
            font=(FONT, 10, "bold"),
            fg=NAVY,
            bg=CARD
        ).pack(
            side="left"
        )

        # ----------------------------------------------------
        # Table
        # ----------------------------------------------------

        table_frame = tk.Frame(
            table_card,
            bg=CARD
        )

        table_frame.pack(
            fill="both",
            expand=True,
            padx=20,
            pady=(0, 20)
        )

        columns = (
            "jobbag",
            "so",
            "po",
            "source",
            "status",
            "output"
        )

        self.results_tree = ttk.Treeview(
            table_frame,
            columns=columns,
            show="headings",
            selectmode="browse"
        )

        self.results_tree.heading(
            "jobbag",
            text="JOBBAG"
        )

        self.results_tree.heading(
            "so",
            text="SO NUMBER"
        )

        self.results_tree.heading(
            "po",
            text="PRODUCTION ORDER"
        )

        self.results_tree.heading(
            "source",
            text="SOURCE PDF"
        )

        self.results_tree.heading(
            "status",
            text="STATUS"
        )

        self.results_tree.heading(
            "output",
            text="OUTPUT"
        )

        self.results_tree.column(
            "jobbag",
            width=140,
            anchor="center"
        )

        self.results_tree.column(
            "so",
            width=140,
            anchor="center"
        )

        self.results_tree.column(
            "po",
            width=170,
            anchor="center"
        )

        self.results_tree.column(
            "source",
            width=220
        )

        self.results_tree.column(
            "status",
            width=120,
            anchor="center"
        )

        self.results_tree.column(
            "output",
            width=330
        )

        scroll = ttk.Scrollbar(
            table_frame,
            orient="vertical",
            command=self.results_tree.yview
        )

        self.results_tree.configure(
            yscrollcommand=scroll.set
        )

        self.results_tree.pack(
            side="left",
            fill="both",
            expand=True
        )

        scroll.pack(
            side="right",
            fill="y"
        )

        self.results_tree.bind(
            "<<TreeviewSelect>>",
            self.result_selected
        )

        self.results_tree.tag_configure(
            "ready",
            foreground=GREEN
        )

        self.results_tree.tag_configure(
            "missing",
            foreground=RED
        )

        self.populate_results_table()

        # ----------------------------------------------------
        # Bottom actions
        # ----------------------------------------------------

        actions = tk.Frame(
            table_card,
            bg=CARD
        )

        actions.pack(
            fill="x",
            padx=20,
            pady=(0, 18)
        )

        self.create_colored_button(
            actions,
            "👁  PREVIEW",
            BLUE,
            BLUE_DARK,
            self.preview_selected,
            15
        ).pack(
            side="left"
        )

        self.create_colored_button(
            actions,
            "🖨  PRINT",
            PURPLE,
            PURPLE_DARK,
            self.print_selected,
            13
        ).pack(
            side="left",
            padx=8
        )

        self.create_colored_button(
            actions,
            "▣  A4 2-UP PRINT",
            GREEN,
            GREEN_DARK,
            self.print_a4_2up,
            17
        ).pack(
            side="left",
            padx=8
        )

        self.create_colored_button(
            actions,
            "📂  OPEN OUTPUT FOLDER",
            ORANGE,
            ORANGE_DARK,
            self.open_output_folder,
            21
        ).pack(
            side="left"
        )

    # ========================================================
    # SMALL STAT
    # ========================================================

    def create_small_stat(
        self,
        parent,
        title,
        variable,
        color
    ):

        frame = tk.Frame(
            parent,
            bg=CARD,
            height=80,
            highlightbackground=BORDER,
            highlightthickness=1
        )

        frame.pack_propagate(
            False
        )

        tk.Label(
            frame,
            text=title,
            font=(FONT, 9),
            fg=MUTED,
            bg=CARD
        ).pack(
            anchor="w",
            padx=15,
            pady=(10, 0)
        )

        tk.Label(
            frame,
            textvariable=variable,
            font=(FONT, 18, "bold"),
            fg=color,
            bg=CARD
        ).pack(
            anchor="w",
            padx=15
        )

        return frame

    # ========================================================
    # SETTINGS
    # ========================================================

    def show_settings(self):

        self.page_title_var.set(
            "Settings"
        )

        self.set_active_nav(
            "settings"
        )

        self.clear_content()

        tk.Label(
            self.content,
            text="Settings",
            font=(FONT, 18, "bold"),
            fg=TEXT,
            bg=BG
        ).pack(
            anchor="w"
        )

        tk.Label(
            self.content,
            text="Application configuration",
            font=(FONT, 10),
            fg=MUTED,
            bg=BG
        ).pack(
            anchor="w",
            pady=(4, 20)
        )

        card = tk.Frame(
            self.content,
            bg=CARD,
            highlightbackground=BORDER,
            highlightthickness=1
        )

        card.pack(
            fill="x"
        )

        tk.Label(
            card,
            text="PDF OUTPUT",
            font=(FONT, 10, "bold"),
            fg=NAVY,
            bg=CARD
        ).pack(
            anchor="w",
            padx=25,
            pady=(22, 5)
        )

        tk.Label(
            card,
            text="Processed PDFs are stored inside an Output\\Processed_TIMESTAMP folder next to the selected PDF.",
            font=(FONT, 10),
            fg=MUTED,
            bg=CARD,
            wraplength=900,
            justify="left"
        ).pack(
            anchor="w",
            padx=25,
            pady=(0, 22)
        )

        card2 = tk.Frame(
            self.content,
            bg=CARD,
            highlightbackground=BORDER,
            highlightthickness=1
        )

        card2.pack(
            fill="x",
            pady=15
        )

        tk.Label(
            card2,
            text="SO / PRODUCTION ORDER PLACEMENT",
            font=(FONT, 10, "bold"),
            fg=NAVY,
            bg=CARD
        ).pack(
            anchor="w",
            padx=25,
            pady=(22, 5)
        )

        tk.Label(
            card2,
            text="SO keeps the approved X position (+108). PO is printed as a separate labeled line with spacing to avoid overlap.",
            font=(FONT, 10),
            fg=MUTED,
            bg=CARD
        ).pack(
            anchor="w",
            padx=25,
            pady=(0, 22)
        )

    # ========================================================
    # ABOUT
    # ========================================================

    def show_about(self):

        self.page_title_var.set(
            "About"
        )

        self.set_active_nav(
            "about"
        )

        self.clear_content()

        card = tk.Frame(
            self.content,
            bg=CARD,
            highlightbackground=BORDER,
            highlightthickness=1
        )

        card.pack(
            fill="both",
            expand=True
        )

        tk.Label(
            card,
            text="KAMA JOBBAG PROCESSOR ",
            font=(FONT, 24, "bold"),
            fg=NAVY,
            bg=CARD
        ).pack(
            pady=(70, 8)
        )

        tk.Label(
            card,
            text="PDF Jobbag Processing & Sales Order Management",
            font=(FONT, 11),
            fg=MUTED,
            bg=CARD
        ).pack()

        tk.Label(
            card,
            text="",
            bg=CARD
        ).pack(
            pady=5
        )

        tk.Label(
            card,
            text="Workflow",
            font=(FONT, 11, "bold"),
            fg=TEXT,
            bg=CARD
        ).pack(
            pady=(20, 5)
        )

        tk.Label(
            card,
            text=(
                "PDF → Jobbag Detection → Excel SO Mapping\n"
                "→ Individual Jobbag PDF → Preview → Print"
            ),
            font=(FONT, 10),
            fg=MUTED,
            bg=CARD,
            justify="center"
        ).pack()

        tk.Label(
            card,
            text="Built with Python • Tkinter • PyMuPDF • Pandas",
            font=(FONT, 9),
            fg="#94A3B8",
            bg=CARD
        ).pack(
            side="bottom",
            pady=25
        )

    # ========================================================
    # BUTTON
    # ========================================================

    def create_colored_button(
        self,
        parent,
        text,
        color,
        hover_color,
        command,
        width=18
    ):

        button = tk.Button(
            parent,
            text=text,
            command=command,
            bg=color,
            fg="white",
            activebackground=hover_color,
            activeforeground="white",
            font=(FONT, 10, "bold"),
            relief="flat",
            bd=0,
            cursor="hand2",
            width=width,
            padx=10,
            pady=10
        )

        def enter(event):

            button.configure(
                bg=hover_color
            )

        def leave(event):

            button.configure(
                bg=color
            )

        button.bind(
            "<Enter>",
            enter
        )

        button.bind(
            "<Leave>",
            leave
        )

        return button

    # ========================================================
    # SELECT PDF
    # ========================================================

    def select_pdfs(self):

        files = filedialog.askopenfilenames(
            title="Select PDF Documents",
            filetypes=[
                ("PDF Files", "*.pdf")
            ]
        )

        if not files:
            return

        self.pdf_files = list(
            files
        )

        count = len(
            self.pdf_files
        )

        if count == 1:

            text = os.path.basename(
                self.pdf_files[0]
            )

        else:

            text = (
                f"{count} PDF files selected"
            )

        if hasattr(
            self,
            "pdf_selection_label"
        ):

            self.pdf_selection_label.config(
                text=text
            )

        self.total_pdf_var.set(
            str(count)
        )

        self.header_status.config(
            text="● READY",
            fg=GREEN,
            bg=LIGHT_GREEN
        )

    # ========================================================
    # SELECT EXCEL
    # ========================================================

    def select_excel(self):

        file = filedialog.askopenfilename(
            title="Select Excel Mapping File",
            filetypes=[
                ("Excel Files", "*.xlsx *.xls")
            ]
        )

        if not file:
            return

        self.excel_file = file

        if hasattr(
            self,
            "excel_selection_label"
        ):

            self.excel_selection_label.config(
                text=os.path.basename(file)
            )

    # ========================================================
    # PROCESS
    # ========================================================

    def process_files(self):

        if not self.pdf_files:

            messagebox.showwarning(
                "PDF Required",
                "Please select at least one PDF file."
            )

            return

        if not self.excel_file:

            messagebox.showwarning(
                "Excel Required",
                "Please select the Excel mapping file."
            )

            return

        try:

            # ------------------------------------------------
            # Excel
            # ------------------------------------------------

            df = pd.read_excel(
                self.excel_file
            )

            required_columns = [
                "JOBBAG_NUMBER",
                "SO_NUMBER",
                "PRODUCTION_ORDER"
            ]

            missing_columns = [
                col
                for col in required_columns
                if col not in df.columns
            ]

            if missing_columns:

                messagebox.showerror(
                    "Excel Error",
                    "Required columns are missing:\n\n"
                    + "\n".join(
                        missing_columns
                    )
                )

                return

            # Normalize Excel values so numeric cells such as 10437212.0
            # still match the 8-digit Jobbag extracted from the PDF.
            def normalize_value(value):
                if pd.isna(value):
                    return ""
                text = str(value).strip()
                if text.endswith(".0") and text[:-2].isdigit():
                    text = text[:-2]
                return text

            df["JOBBAG_NUMBER"] = df["JOBBAG_NUMBER"].apply(normalize_value)
            df["SO_NUMBER"] = df["SO_NUMBER"].apply(normalize_value)
            df["PRODUCTION_ORDER"] = df["PRODUCTION_ORDER"].apply(normalize_value)

            mapping = {}
            for _, row in df.iterrows():
                jobbag_key = row["JOBBAG_NUMBER"]
                if jobbag_key:
                    mapping[jobbag_key] = {
                        "so": row["SO_NUMBER"],
                        "po": row["PRODUCTION_ORDER"]
                    }

            # ------------------------------------------------
            # Output
            # ------------------------------------------------

            timestamp = datetime.now().strftime(
                "%Y%m%d_%H%M%S"
            )

            source_dir = os.path.normpath(
                os.path.dirname(
                    os.path.abspath(self.pdf_files[0])
                )
            )

            self.output_folder = os.path.normpath(
                os.path.join(
                    source_dir,
                    "Output",
                    f"Processed_{timestamp}"
                )
            )

            os.makedirs(
                self.output_folder,
                exist_ok=True
            )

            self.results = []

            # ------------------------------------------------
            # Progress
            # ------------------------------------------------

            self.progress["maximum"] = len(
                self.pdf_files
            )

            self.progress["value"] = 0

            self.header_status.config(
                text="● PROCESSING",
                fg=ORANGE_DARK,
                bg=LIGHT_ORANGE
            )

            # ------------------------------------------------
            # Process PDFs
            # ------------------------------------------------

            for index, pdf_path in enumerate(
                self.pdf_files,
                start=1
            ):

                filename = os.path.basename(
                    pdf_path
                )

                self.progress_label.config(
                    text=(
                        f"Processing {filename} "
                        f"({index}/{len(self.pdf_files)})"
                    )
                )

                self.root.update_idletasks()

                self.process_single_pdf(
                    pdf_path,
                    mapping
                )

                self.progress["value"] = index

                self.root.update_idletasks()

            # ------------------------------------------------
            # CREATE A4 2-UP PRINT FILE
            # ------------------------------------------------
            self.create_a4_2up_pdf()

            # ------------------------------------------------
            # Counts
            # ------------------------------------------------

            total = len(
                self.results
            )

            ready = sum(
                1
                for r in self.results
                if r["status"] == "Ready"
            )

            missing = sum(
                1
                for r in self.results
                if r["status"] == "SO / PO Missing"
            )

            self.total_jobbag_var.set(
                str(total)
            )

            self.ready_var.set(
                str(ready)
            )

            self.missing_var.set(
                str(missing)
            )

            self.header_status.config(
                text="● COMPLETED",
                fg=GREEN,
                bg=LIGHT_GREEN
            )

            self.progress_label.config(
                text=(
                    f"Completed • "
                    f"{total} Jobbags"
                )
            )

            # ------------------------------------------------
            # Results
            # ------------------------------------------------

            self.show_results()

            messagebox.showinfo(
                "Processing Complete",
                f"Processing completed successfully.\n\n"
                f"PDFs processed: {len(self.pdf_files)}\n"
                f"Jobbags found: {total}\n"
                f"Ready: {ready}\n"
                f"SO / PO Missing: {missing}\n\n"
                f"Output folder:\n{self.output_folder}\n\n"
                + (
                    "A4 2-UP file created: KAMA_Jobbags_A4_2UP_Print.pdf"
                    if self.a4_2up_path
                    else "No A4 2-UP file created because no Jobbags were Ready."
                )
            )

        except Exception as e:

            self.header_status.config(
                text="● ERROR",
                fg=RED,
                bg=RED_LIGHT
            )

            messagebox.showerror(
                "Processing Error",
                str(e)
            )

    # ========================================================
    # CREATE A4 2-UP PRINT PDF
    # ========================================================

    def create_a4_2up_pdf(self):

        ready_results = [
            result
            for result in self.results
            if result.get("status") == "Ready"
            and os.path.isfile(self._windows_path(result.get("output", "")))
        ]

        self.a4_2up_path = ""

        if not ready_results:
            return

        # IMPORTANT:
        # Each Jobbag output is A5 LANDSCAPE (210 x 148 mm).
        # Two A5 LANDSCAPE pages fit perfectly on one A4 PORTRAIT page
        # when placed one above the other. This preserves the complete
        # Jobbag image without rotating, cropping, or squeezing it.
        # Side-by-side A5 LANDSCAPE on A4 would require rotating the
        # Jobbag, so top/bottom is the correct print-ready arrangement.
        A4_WIDTH = 595.276
        A4_HEIGHT = 841.89
        HALF_HEIGHT = A4_HEIGHT / 2

        output_path = os.path.join(
            self.output_folder,
            "KAMA_Jobbags_A4_2UP_Print.pdf"
        )

        combined = pymupdf.open()

        for start in range(0, len(ready_results), 2):

            a4_page = combined.new_page(
                width=A4_WIDTH,
                height=A4_HEIGHT
            )

            pair = ready_results[start:start + 2]

            for position, result in enumerate(pair):

                pdf_path = self._windows_path(result["output"])

                try:
                    source_doc = pymupdf.open(pdf_path)

                    if len(source_doc) > 0:
                        # Full-width A5 landscape area.
                        # First Jobbag = top half; second Jobbag = bottom half.
                        target_rect = pymupdf.Rect(
                            0,
                            position * HALF_HEIGHT,
                            A4_WIDTH,
                            (position + 1) * HALF_HEIGHT
                        )

                        # Fill the exact A5-sized area. No centering blank space,
                        # no cropping and no rotation.
                        a4_page.show_pdf_page(
                            target_rect,
                            source_doc,
                            0,
                            keep_proportion=False
                        )

                    source_doc.close()

                except Exception:
                    continue

        if len(combined) > 0:
            combined.save(output_path, garbage=4, deflate=True)
            self.a4_2up_path = output_path

        combined.close()

    # ========================================================
    # PROCESS SINGLE PDF
    # ========================================================

    def process_single_pdf(
        self,
        pdf_path,
        mapping
    ):

        doc = pymupdf.open(
            pdf_path
        )

        source_name = os.path.basename(
            pdf_path
        )

        for page_number in range(
            len(doc)
        ):

            page = doc[
                page_number
            ]

            page_width = page.rect.width
            page_height = page.rect.height

            words = page.get_text(
                "words"
            )

            jobbags = []

            for word in words:

                text = word[4].strip()

                if (
                    re.fullmatch(
                        r"\d{8}",
                        text
                    )
                    and word[1] < 150
                ):

                    x = word[0]
                    y = word[1]

                    if not any(
                        j["jobbag"] == text
                        for j in jobbags
                    ):

                        jobbags.append(
                            {
                                "jobbag": text,
                                "x": x,
                                "y": y
                            }
                        )

            if not jobbags:
                continue

            jobbags.sort(
                key=lambda j: (
                    round(j["y"], 1),
                    j["x"]
                )
            )

            count = len(
                jobbags
            )

            section_width = (
                page_width / count
            )

            for i, job in enumerate(
                jobbags
            ):

                jobbag = job[
                    "jobbag"
                ]

                mapped = mapping.get(
                    jobbag,
                    {}
                )

                so_number = mapped.get("so", "")
                production_order = mapped.get("po", "")

                status = (
                    "Ready"
                    if so_number and production_order
                    else "SO / PO Missing"
                )

                left = (
                    i * section_width
                )

                right = (
                    (i + 1)
                    * section_width
                )

                clip = pymupdf.Rect(
                    left,
                    0,
                    right,
                    page_height
                )

                # -------------------------------------------------
                # OUTPUT PAGE: A5 LANDSCAPE
                # -------------------------------------------------
                # The original PDF section is scaled to a real A5
                # landscape page (210 x 148 mm). Overlay coordinates
                # are scaled by the same factor so SO/PO stay in the
                # approved blank area.
                A5_WIDTH = 595.276
                A5_HEIGHT = 419.528

                new_doc = pymupdf.open()

                new_page = new_doc.new_page(
                    width=A5_WIDTH,
                    height=A5_HEIGHT
                )

                new_page.show_pdf_page(
                    new_page.rect,
                    doc,
                    page_number,
                    clip=clip
                )

                # Scale source coordinates into the A5 output page.
                scale_x = A5_WIDTH / clip.width
                scale_y = A5_HEIGHT / clip.height

                # Keep the approved v4 placement relative to the
                # original Jobbag/blank header area.
                source_box_x = (
                    job["x"]
                    - clip.x0
                    + 148
                )
                source_box_y = job["y"] + 3

                box_x = source_box_x * scale_x
                box_y = source_box_y * scale_y

                # A little wider than the previous box so the full
                # labels remain visible after A5 scaling.
                box_width = 120
                line_height = 15

                orange = (1.0, 0.35, 0.0)

                # IMPORTANT:
                # Do not paint a white rectangle over the source PDF.
                # The original blank area is preserved.

                if so_number:
                    new_page.insert_text(
                        (box_x, box_y + 10),
                        f"SO: {so_number}",
                        fontsize=8.0,
                        fontname="hebo",
                        color=orange
                    )

                if production_order:
                    new_page.insert_text(
                        (box_x, box_y + 10 + line_height),
                        f"PO: {production_order}",
                        fontsize=8.0,
                        fontname="hebo",
                        color=orange
                    )

                safe_so = (
                    so_number
                    if so_number
                    else "NO_SO"
                )

                safe_po = (
                    production_order
                    if production_order
                    else "NO_PO"
                )

                output_name = (
                    f"{jobbag}_SO_{safe_so}_PO_{safe_po}.pdf"
                )

                output_path = os.path.join(
                    self.output_folder,
                    output_name
                )

                new_doc.save(
                    output_path
                )

                new_doc.close()

                self.results.append(
                    {
                        "jobbag": jobbag,
                        "so": so_number,
                        "po": production_order,
                        "source": source_name,
                        "status": status,
                        "output": output_path
                    }
                )

        doc.close()

    # ========================================================
    # POPULATE RESULTS
    # ========================================================

    def populate_results_table(self):

        if not hasattr(
            self,
            "results_tree"
        ):
            return

        for item in self.results_tree.get_children():
            self.results_tree.delete(
                item
            )

        for index, result in enumerate(
            self.results
        ):

            iid = f"result_{index}"

            self.results_tree.insert(
                "",
                "end",
                iid=iid,
                values=(
                    result["jobbag"],
                    result["so"]
                    if result["so"]
                    else "-",
                    result["po"]
                    if result["po"]
                    else "-",
                    result["source"],
                    result["status"],
                    os.path.basename(
                        result["output"]
                    )
                ),
                tags=(
                    "ready"
                    if result["status"] == "Ready"
                    else "missing"
                )
            )

        self.result_summary_label.config(
            text=f"{len(self.results)} records"
        )

        # Select first Ready item
        for index, result in enumerate(
            self.results
        ):

            if result["status"] == "Ready":

                iid = f"result_{index}"

                self.results_tree.selection_set(
                    iid
                )

                self.results_tree.focus(
                    iid
                )

                self.results_tree.see(
                    iid
                )

                break

    # ========================================================
    # SELECT RESULT
    # ========================================================

    def result_selected(
        self,
        event=None
    ):

        result = self.get_selected_result()

        if not result:
            return

    # ========================================================
    # GET SELECTED
    # ========================================================

    def get_selected_result(self):

        if not hasattr(
            self,
            "results_tree"
        ):
            return None

        selection = (
            self.results_tree.selection()
        )

        if not selection:
            return None

        iid = selection[0]

        try:

            index = int(
                iid.split("_")[1]
            )

        except Exception:

            return None

        if (
            index < 0
            or index >= len(
                self.results
            )
        ):

            return None

        return self.results[
            index
        ]

    # ========================================================
    # PREVIEW
    # ========================================================

    def preview_selected(self):

        result = (
            self.get_selected_result()
        )

        if not result:

            messagebox.showwarning(
                "No Selection",
                "Please select a Jobbag."
            )

            return

        if not os.path.exists(
            result["output"]
        ):

            messagebox.showerror(
                "File Missing",
                "Output PDF was not found."
            )

            return

        self.open_preview(
            result["output"],
            result["jobbag"]
        )

    # ========================================================
    # PREVIEW WINDOW
    # ========================================================

    def open_preview(
        self,
        pdf_path,
        jobbag
    ):

        if (
            self.preview_window
            and self.preview_window.winfo_exists()
        ):

            self.preview_window.destroy()

        self.preview_window = tk.Toplevel(
            self.root
        )

        self.preview_window.title(
            f"Preview - {jobbag}"
        )

        self.preview_window.geometry(
            "1100x800"
        )

        self.preview_window.configure(
            bg=NAVY_DARK
        )

        # ----------------------------------------------------
        # Header
        # ----------------------------------------------------

        header = tk.Frame(
            self.preview_window,
            bg=NAVY,
            height=65
        )

        header.pack(
            fill="x"
        )

        header.pack_propagate(
            False
        )

        tk.Label(
            header,
            text=f"JOBBAG PREVIEW  •  {jobbag}",
            font=(FONT, 14, "bold"),
            fg="white",
            bg=NAVY
        ).pack(
            side="left",
            padx=25
        )

        # ----------------------------------------------------
        # Toolbar
        # ----------------------------------------------------

        toolbar = tk.Frame(
            self.preview_window,
            bg="#E2E8F0",
            height=55
        )

        toolbar.pack(
            fill="x"
        )

        toolbar.pack_propagate(
            False
        )

        self.create_colored_button(
            toolbar,
            "−  Zoom Out",
            NAVY,
            NAVY_DARK,
            self.zoom_out,
            12
        ).pack(
            side="left",
            padx=(15, 5),
            pady=8
        )

        self.create_colored_button(
            toolbar,
            "＋  Zoom In",
            BLUE,
            BLUE_DARK,
            self.zoom_in,
            12
        ).pack(
            side="left",
            padx=5,
            pady=8
        )

        self.create_colored_button(
            toolbar,
            "Fit Page",
            GREEN,
            GREEN_DARK,
            self.fit_page,
            11
        ).pack(
            side="left",
            padx=5,
            pady=8
        )

        self.create_colored_button(
            toolbar,
            "Open PDF",
            ORANGE,
            ORANGE_DARK,
            lambda: self.open_pdf_file(
                pdf_path
            ),
            12
        ).pack(
            side="right",
            padx=15,
            pady=8
        )

        # ----------------------------------------------------
        # Canvas
        # ----------------------------------------------------

        canvas_frame = tk.Frame(
            self.preview_window,
            bg="#94A3B8"
        )

        canvas_frame.pack(
            fill="both",
            expand=True
        )

        self.preview_canvas = tk.Canvas(
            canvas_frame,
            bg="#94A3B8",
            highlightthickness=0
        )

        self.preview_canvas.pack(
            fill="both",
            expand=True
        )

        try:

            self.preview_doc = pymupdf.open(
                pdf_path
            )

            self.preview_page = (
                self.preview_doc[0]
            )

            self.preview_zoom = 1.0

            self.preview_window.after(
                150,
                self.fit_page
            )

        except Exception as e:

            messagebox.showerror(
                "Preview Error",
                str(e)
            )

    # ========================================================
    # RENDER PREVIEW
    # ========================================================

    def render_preview(self):

        if (
            not self.preview_page
            or not self.preview_window
            or not self.preview_window.winfo_exists()
        ):

            return

        matrix = pymupdf.Matrix(
            self.preview_zoom,
            self.preview_zoom
        )

        pix = self.preview_page.get_pixmap(
            matrix=matrix,
            alpha=False
        )

        image = Image.frombytes(
            "RGB",
            [pix.width, pix.height],
            pix.samples
        )

        self.preview_image = ImageTk.PhotoImage(
            image
        )

        self.preview_canvas.delete(
            "all"
        )

        canvas_width = (
            self.preview_canvas.winfo_width()
        )

        canvas_height = (
            self.preview_canvas.winfo_height()
        )

        x = max(
            (canvas_width - image.width) // 2,
            10
        )

        y = max(
            (canvas_height - image.height) // 2,
            10
        )

        self.preview_canvas.create_image(
            x,
            y,
            anchor="nw",
            image=self.preview_image
        )

    # ========================================================
    # ZOOM
    # ========================================================

    def zoom_in(self):

        self.preview_zoom *= 1.2

        if self.preview_zoom > 4:
            self.preview_zoom = 4

        self.render_preview()

    def zoom_out(self):

        self.preview_zoom /= 1.2

        if self.preview_zoom < 0.3:
            self.preview_zoom = 0.3

        self.render_preview()

    # ========================================================
    # FIT
    # ========================================================

    def fit_page(self):

        if not self.preview_page:
            return

        width = (
            self.preview_canvas.winfo_width()
        )

        height = (
            self.preview_canvas.winfo_height()
        )

        if width <= 50 or height <= 50:
            return

        rect = self.preview_page.rect

        width_ratio = (
            width - 50
        ) / rect.width

        height_ratio = (
            height - 50
        ) / rect.height

        self.preview_zoom = min(
            width_ratio,
            height_ratio
        )

        if self.preview_zoom <= 0:
            self.preview_zoom = 1

        self.render_preview()

    # ========================================================
    # WINDOWS FILE / FOLDER HELPERS
    # ========================================================

    def _windows_path(self, path):
        """Return a clean absolute Windows path."""
        if not path:
            return ""
        return os.path.normpath(os.path.abspath(str(path)))

    def _open_with_windows(self, path):
        """Open a file/folder through Windows Explorer/default association."""
        path = self._windows_path(path)

        if not path or not os.path.exists(path):
            return False

        try:
            # Explorer is more reliable from a PyInstaller EXE than relying
            # only on os.startfile().
            subprocess.Popen(["explorer.exe", path], close_fds=True)
            return True
        except Exception:
            pass

        try:
            os.startfile(path)
            return True
        except Exception:
            return False

    # ========================================================
    # OPEN PDF
    # ========================================================

    def open_pdf_file(self, path):

        path = self._windows_path(path)

        if not os.path.exists(path):
            messagebox.showerror(
                "Open PDF",
                f"PDF file was not found:\n\n{path}"
            )
            return

        if not self._open_with_windows(path):
            messagebox.showerror(
                "Open PDF",
                "Windows could not open this PDF.\n\n"
                "Please make sure a PDF reader such as Microsoft Edge or Adobe Acrobat is installed."
            )

    # ========================================================
    # OPEN OUTPUT
    # ========================================================

    def open_output_folder(self):

        folder = self._windows_path(self.output_folder)

        if not folder or not os.path.isdir(folder):
            messagebox.showwarning(
                "Output Folder",
                "No processed output folder is available yet."
            )
            return

        if not self._open_with_windows(folder):
            messagebox.showerror(
                "Output Folder",
                f"Windows could not open the output folder:\n\n{folder}"
            )

    # ========================================================
    # PRINT A4 2-UP
    # ========================================================

    def print_a4_2up(self):

        if not self.a4_2up_path:
            self.create_a4_2up_pdf()

        pdf_path = self._windows_path(self.a4_2up_path)

        if not pdf_path or not os.path.isfile(pdf_path):
            messagebox.showwarning(
                "A4 2-UP",
                "No ready Jobbags are available for A4 2-up printing."
            )
            return

        # Try the Windows Print verb first.
        try:
            os.startfile(pdf_path, "print")
            return
        except Exception:
            pass

        # PowerShell fallback.
        try:
            ps_script = (
                "$p = Get-Item -LiteralPath "
                + repr(pdf_path)
                + "; Start-Process -FilePath $p.FullName -Verb Print"
            )

            completed = subprocess.run(
                [
                    "powershell.exe",
                    "-NoProfile",
                    "-ExecutionPolicy",
                    "Bypass",
                    "-Command",
                    ps_script,
                ],
                capture_output=True,
                text=True,
                timeout=15,
                creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
            )

            if completed.returncode == 0:
                return

        except Exception:
            pass

        if self._open_with_windows(pdf_path):
            messagebox.showinfo(
                "A4 2-UP Print",
                "The A4 2-UP PDF has been opened.\n\n"
                "It contains 2 Jobbags side-by-side on each A4 landscape page.\n\n"
                "Press Ctrl+P and select A4 paper for printing."
            )
            return

        messagebox.showerror(
            "A4 2-UP Print",
            "The A4 2-UP PDF could not be opened for printing."
        )

    # ========================================================
    # PRINT
    # ========================================================

    def print_selected(self):

        result = self.get_selected_result()

        if not result:
            messagebox.showwarning(
                "No Selection",
                "Please select a Jobbag."
            )
            return

        if result["status"] != "Ready":
            messagebox.showwarning(
                "SO / PO Missing",
                "This Jobbag does not have both SO Number and Production Order."
            )
            return

        pdf_path = self._windows_path(result["output"])

        if not os.path.isfile(pdf_path):
            messagebox.showerror(
                "File Missing",
                f"Output PDF was not found:\n\n{pdf_path}"
            )
            return

        # First try the normal Windows PDF print verb.
        try:
            os.startfile(pdf_path, "print")
            return
        except OSError:
            pass
        except Exception:
            pass

        # Some Windows installations expose the Print verb through PowerShell
        # even when os.startfile(..., "print") is unavailable to Python.
        try:
            ps_script = (
                "$p = Get-Item -LiteralPath "
                + repr(pdf_path)
                + "; Start-Process -FilePath $p.FullName -Verb Print"
            )
            completed = subprocess.run(
                [
                    "powershell.exe",
                    "-NoProfile",
                    "-ExecutionPolicy",
                    "Bypass",
                    "-Command",
                    ps_script,
                ],
                capture_output=True,
                text=True,
                timeout=15,
                creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
            )

            if completed.returncode == 0:
                return
        except Exception:
            pass

        # Final safe fallback: open the PDF so the user can press Ctrl+P.
        if self._open_with_windows(pdf_path):
            messagebox.showinfo(
                "Print",
                "Your PDF has been opened.\n\n"
                "Windows does not have a PDF Print action registered for this computer, "
                "so please press Ctrl+P in the PDF viewer and select your printer."
            )
            return

        messagebox.showerror(
            "Print Error",
            "The PDF could not be opened for printing.\n\n"
            "Please install or repair your PDF reader (Microsoft Edge or Adobe Acrobat)."
        )


# ============================================================
# START APPLICATION
# ============================================================

if __name__ == "__main__":

    root = tk.Tk()

    app = JobbagApplication(
        root
    )

    root.mainloop()