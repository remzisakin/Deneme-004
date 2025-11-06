"""Modern sales data entry application with Excel persistence and reporting hooks.

This module implements the user-facing GUI described in the specification.  It
provides a Tkinter based user interface with modern styling, input
validation, Excel synchronisation and integration with the ``sales_reporting``
module.  The goal of the implementation is to provide a pleasant and reliable
experience for day-to-day sales data management.
"""
from __future__ import annotations

import json
import os
import threading
import tkinter as tk
from dataclasses import dataclass
from datetime import date, datetime, timedelta
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from functools import partial
from pathlib import Path
from tkinter import filedialog, messagebox, simpledialog, ttk
from typing import Dict, Iterable, List, Optional, Tuple, TYPE_CHECKING

import pandas as pd
from tkcalendar import DateEntry

import sales_reporting

try:  # Optional dependency for visual reporting
    from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
    from matplotlib.figure import Figure
except ImportError:  # pragma: no cover - runtime guard
    FigureCanvasTkAgg = None  # type: ignore[assignment]
    Figure = None  # type: ignore[assignment]

if TYPE_CHECKING:
    from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg as FigureCanvasTkAggType
else:  # pragma: no cover - runtime fallback
    FigureCanvasTkAggType = object


APP_TITLE = "Satış Veri Giriş Sistemi"
APP_GEOMETRY = "1200x820"
DATA_FILE = "sales_data_master.xlsx"
CONFIG_FILE = "config.json"
BACKUP_DIR = Path("backups")
REPORT_DIR = Path("reports")
AUTO_SAVE_INTERVAL = 5 * 60 * 1000  # 5 minutes in milliseconds
PAGE_SIZE = 15
FORM_BUTTON_WIDTH = 12

ASSETS_DIR = Path("assets")
DESOUTTER_THEME_KEY = "desoutter"

COLUMNS = [
    "Date of Request",
    "Date of Issue",
    "Date of Delivery",
    "Sales Man",
    "Customer Name",
    "Customer PO No",
    "Definition",
    "Sales Ticket Reference",
    "SO No",
    "PTD PO No",
    "NON-EDI PO No",
    "Amount",
    "Total Discount",
    "CPI",
    "CPS",
    "QI Forecast",
    "Delivery Note",
    "Invoiced",
    "Invoiced Amount",
]

DEFAULT_SALES_REPS = ["Fatih Aykut", "Ridvan Yasar", "Rami Sakin"]
OTHER_SALES_REP_OPTION = "Diğer..."
DEFAULT_SALES_REP_PASSWORD = "Remzi123"
DEFAULT_THEME = "light"
THEMES = {"light": "Açık", "dark": "Koyu", DESOUTTER_THEME_KEY: "Desoutter Tema"}

COLORS = {
    "primary": "#667eea",
    "secondary": "#764ba2",
    "success": "#10b981",
    "danger": "#ef4444",
    "warning": "#f59e0b",
    "info": "#3b82f6",
    "bg_light": "#f9fafb",
    "bg_dark": "#1f2937",
    "text_dark": "#111827",
    "text_light": "#6b7280",
    "border": "#e5e7eb",
}

REQUIRED_FIELDS = {
    "Date of Request",
    "Date of Issue",
    "Sales Man",
    "Customer Name",
    "Definition",
    "Amount",
    "QI Forecast",
    "Invoiced",
}

FLOAT_FIELDS = {"Amount", "Total Discount", "CPI", "CPS", "Invoiced Amount"}

CURRENCY_FIELDS = {"Amount", "CPS", "CPI", "Invoiced Amount"}


@dataclass
class FilterOptions:
    search_text: str = ""
    salesman: str = ""
    invoiced: str = ""
    so_no: str = ""
    start_date: Optional[datetime] = None
    end_date: Optional[datetime] = None


class SalesEntryApp:
    """Main application class wrapping the Tkinter GUI."""

    def __init__(self) -> None:
        self.root = tk.Tk()
        self.root.title(APP_TITLE)
        self.root.geometry(APP_GEOMETRY)
        self.root.minsize(1100, 760)
        self.df = pd.DataFrame(columns=COLUMNS)
        self.current_page = 1
        self.total_pages = 1
        self.selected_index: Optional[int] = None
        self._new_entry_mode = False
        self.history: List[pd.DataFrame] = []
        self.redo_stack: List[pd.DataFrame] = []
        self.filter_options = FilterOptions()
        self._updating_cpi_field = False
        self._suspend_delivery_autofill = False
        self._theme_settings: Dict[str, str] = {}
        self._desoutter_logo: Optional[tk.PhotoImage] = None
        self._report_canvases: List[FigureCanvasTkAggType] = []
        
        self._config = self._load_config()
        self.sales_reps = self._load_sales_reps()
        self.sales_rep_password = self._config.get(
            "sales_rep_password", DEFAULT_SALES_REP_PASSWORD
        )
        self._load_theme_assets()
        self._apply_theme(self._config.get("theme", DEFAULT_THEME))

        self._ensure_directories()
        self._ensure_excel_file()

        self._create_styles()
        self._create_status_bar()
        self._create_menu()
        self._create_main_frames()
        self._create_form()
        self._create_table()
        self._create_action_buttons()
        self._create_bottom_buttons()

        self.load_data()
        self.schedule_auto_backup()

    # ------------------------------------------------------------------ setup
    def _load_config(self) -> Dict[str, object]:
        if not os.path.exists(CONFIG_FILE):
            config = {
                "theme": DEFAULT_THEME,
                "sales_reps": DEFAULT_SALES_REPS,
                "sales_rep_password": DEFAULT_SALES_REP_PASSWORD,
            }
            self._save_config(config)
            return config
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as fh:
                config = json.load(fh)
        except (json.JSONDecodeError, OSError):
            config = {
                "theme": DEFAULT_THEME,
                "sales_reps": DEFAULT_SALES_REPS,
                "sales_rep_password": DEFAULT_SALES_REP_PASSWORD,
            }
        if "theme" not in config:
            config["theme"] = DEFAULT_THEME
        sales_reps = config.get("sales_reps")
        if not isinstance(sales_reps, list) or not all(isinstance(item, str) for item in sales_reps):
            config["sales_reps"] = DEFAULT_SALES_REPS.copy()
        password = config.get("sales_rep_password")
        if not isinstance(password, str) or not password.strip():
            config["sales_rep_password"] = DEFAULT_SALES_REP_PASSWORD
        return config

    def _save_config(self, config: Dict[str, object]) -> None:
        with open(CONFIG_FILE, "w", encoding="utf-8") as fh:
            json.dump(config, fh, indent=2, ensure_ascii=False)

    def _update_sales_rep_password(self, new_password: str) -> None:
        self.sales_rep_password = new_password
        self._config["sales_rep_password"] = new_password
        self._save_config(self._config)

    def _load_sales_reps(self) -> List[str]:
        reps = self._config.get("sales_reps", DEFAULT_SALES_REPS)
        cleaned = [rep.strip() for rep in reps if isinstance(rep, str) and rep.strip()]
        if not cleaned:
            cleaned = DEFAULT_SALES_REPS.copy()
        return cleaned

    def _save_sales_reps(self) -> None:
        self._config["sales_reps"] = self.sales_reps
        self._save_config(self._config)

    def _get_sales_rep_options(self) -> List[str]:
        return [*self.sales_reps, OTHER_SALES_REP_OPTION]

    def _load_theme_assets(self) -> None:
        logo_path = ASSETS_DIR / "desoutter_logo_dark.png"
        if logo_path.exists():
            try:
                self._desoutter_logo = tk.PhotoImage(file=str(logo_path))
            except tk.TclError:
                self._desoutter_logo = None

    def _apply_theme(self, theme: str) -> None:
        if theme == "dark":
                settings = {
                "bg": COLORS["bg_dark"],
                "fg": "white",
                "accent": COLORS["primary"],
                "secondary": COLORS["secondary"],
                "action_bg": "#374151",
                "action_fg": "white",
                "disabled_bg": "#4b5563",
                "disabled_fg": "#9ca3af",
                "card_bg": "#2a2f3a",
                "table_bg": "#1f2530",
                "table_fg": "#f4f4f5",
            }
        elif theme == DESOUTTER_THEME_KEY:
            settings = {
                "bg": "#101015",
                "fg": "#f3f4f6",
                "accent": "#E4002B",
                "secondary": "#b30f27",
                "action_bg": "#1e1e26",
                "action_fg": "#f3f4f6",
                "disabled_bg": "#2b2b35",
                "disabled_fg": "#9ca3af",
                "card_bg": "#16161f",
                "table_bg": "#1d1d27",
                "table_fg": "#f3f4f6",
            }        
        else:
            settings = {
                "bg": COLORS["bg_light"],
                "fg": COLORS["text_dark"],
                "accent": COLORS["primary"],
                "secondary": COLORS["secondary"],
                "action_bg": "#e5e7eb",
                "action_fg": COLORS["text_dark"],
                "disabled_bg": "#d1d5db",
                "disabled_fg": COLORS["text_light"],
                "card_bg": "#ffffff",
                "table_bg": "#ffffff",
                "table_fg": COLORS["text_dark"],
            }

        self._theme_settings = settings
        self.root.configure(bg=settings["bg"])
        
        style = ttk.Style(self.root)
        style.theme_use("clam")
        style.configure("TLabel", background=settings["bg"], foreground=settings["fg"], font=("Segoe UI", 10))
        style.configure(
            "Header.TLabel",
            font=("Segoe UI", 14, "bold"),
            foreground=settings.get("accent", COLORS["primary"]),
            background=settings["bg"],
        )
        style.configure("TFrame", background=settings["bg"])
        style.configure("TButton", font=("Segoe UI", 10), padding=6)
        style.map(
            "TButton",
            background=[("active", settings.get("accent", COLORS["primary"]))],
            foreground=[("active", "white")],
        )
        style.configure(
            "TLabelframe",
            background=settings["bg"],
            borderwidth=0,
            padding=(10, 8),
        )
        style.configure(
            "TLabelframe.Label",
            background=settings["bg"],
            foreground=settings["fg"],
            font=("Segoe UI", 11, "bold"),
        )
        style.configure(
            "Card.TLabelframe",
            background=settings["card_bg"],
            borderwidth=0,
            padding=(14, 10),
        )
        style.configure(
            "Card.TLabelframe.Label",
            background=settings["card_bg"],
            foreground=settings["fg"],
            font=("Segoe UI", 12, "bold"),
        )
        style.configure("Card.TFrame", background=settings["card_bg"])
        style.configure("Card.TLabel", background=settings["card_bg"], foreground=settings["fg"], font=("Segoe UI", 10))
        style.configure(
            "Action.TButton",
            font=("Segoe UI", 11, "bold"),
            padding=(10, 6),
            background=settings["action_bg"],
            foreground=settings["action_fg"],
        )
        style.map(
            "Action.TButton",
            background=[("disabled", settings["disabled_bg"]), ("active", settings.get("accent", COLORS["primary"]))],
            foreground=[("disabled", settings["disabled_fg"]), ("active", "white")],
        )
        style.configure(
            "ActionPrimary.TButton",
            font=("Segoe UI", 11, "bold"),
            padding=(10, 6),
            background=settings.get("accent", COLORS["primary"]),
            foreground="white",
            borderwidth=0,
        )
        style.map(
            "ActionPrimary.TButton",
            background=[("disabled", settings["disabled_bg"]), ("active", settings.get("secondary", COLORS["secondary"]))],
            foreground=[("disabled", settings["disabled_fg"]), ("active", "white")],
        )
        style.configure(
            "Accent.TButton",
            background=settings.get("accent", COLORS["primary"]),
            foreground="white",
            borderwidth=0,
        )
        style.map(
            "Accent.TButton",
            background=[("active", settings.get("secondary", COLORS["secondary"]))],
        )        
        style.configure(
            "ActionAccent.TButton",
            font=("Segoe UI", 11, "bold"),
            padding=(10, 6),
            background=COLORS["success"],
            foreground="white",
            borderwidth=0,
        )
        style.map(
            "ActionAccent.TButton",
            background=[("disabled", settings["disabled_bg"]), ("active", "#059669")],
            foreground=[("disabled", settings["disabled_fg"]), ("active", "white")],
        )
        style.configure(
            "ActionDanger.TButton",
            font=("Segoe UI", 11, "bold"),
            padding=(10, 6),
            background=COLORS["danger"],
            foreground="white",
            borderwidth=0,
        )
        style.map(
            "ActionDanger.TButton",
            background=[("disabled", settings["disabled_bg"]), ("active", "#b91c1c")],
            foreground=[("disabled", settings["disabled_fg"]), ("active", "white")],
        )
        style.configure(
            "Custom.Treeview",
            background=settings["table_bg"],
            foreground=settings["table_fg"],
            fieldbackground=settings["table_bg"],
            rowheight=26,
            font=("Segoe UI", 10),
            borderwidth=0,
        )
        style.map(
            "Custom.Treeview",
            background=[("selected", settings.get("accent", COLORS["secondary"]))],
            foreground=[("selected", "white")],
        )
        style.configure(
            "Custom.Treeview.Heading",
            font=("Segoe UI", 10, "bold"),
            background=settings["card_bg"],
            foreground=settings["fg"],
        )
        style.configure(
            "Report.Treeview",
            background=settings["card_bg"],
            foreground=settings["fg"],
            fieldbackground=settings["card_bg"],
            rowheight=28,
            font=("Segoe UI", 10),
            borderwidth=0,
        )
        style.map(
            "Report.Treeview",
            background=[("selected", settings.get("accent", COLORS["primary"]))],
            foreground=[("selected", "white")],
        )
        style.configure(
            "Report.Treeview.Heading",
            font=("Segoe UI", 10, "bold"),
            background=settings["card_bg"],
            foreground=settings["fg"],
        )
        
        self._apply_branding(theme)

    def _apply_branding(self, theme: str) -> None:
        if not hasattr(self, "logo_label"):
            return
        if theme == DESOUTTER_THEME_KEY and self._desoutter_logo is not None:
            bg = self._theme_settings.get("bg", COLORS["bg_dark"])
            self.logo_label.configure(image=self._desoutter_logo, background=bg)
            if not self.logo_label.winfo_manager():
                self.logo_label.pack(side="right", padx=(12, 0))
        else:
            bg = self._theme_settings.get("bg", COLORS["bg_light"])
            self.logo_label.configure(image="", background=bg)
            if self.logo_label.winfo_manager():
                self.logo_label.pack_forget()
        
    def _ensure_directories(self) -> None:
        BACKUP_DIR.mkdir(exist_ok=True)
        REPORT_DIR.mkdir(exist_ok=True)

    def _ensure_excel_file(self) -> None:
        if os.path.exists(DATA_FILE):
            return
        df = pd.DataFrame(columns=COLUMNS)
        df.to_excel(DATA_FILE, index=False)

    def _create_styles(self) -> None:
        # Additional style customisation for treeview
        style = ttk.Style(self.root)
        style.configure(
            "Custom.Treeview",
            rowheight=26,
            font=("Segoe UI", 10),
        )
        style.configure("Custom.Treeview.Heading", font=("Segoe UI", 10, "bold"))
        style.configure("Report.Treeview", rowheight=28, font=("Segoe UI", 10))        

    def _create_status_bar(self) -> None:
        self.status_var = tk.StringVar()
        self.status_label = ttk.Label(self.root, textvariable=self.status_var, anchor="w")
        self.status_label.pack(side="bottom", fill="x", padx=8, pady=(0, 6))
        self._update_status("Hazır")

    def _create_menu(self) -> None:
        menubar = tk.Menu(self.root)

        file_menu = tk.Menu(menubar, tearoff=0)
        file_menu.add_command(label="Yeni Dosya", command=self.create_new_file)
        file_menu.add_command(label="Dosya Aç", command=self.open_existing_file)
        file_menu.add_separator()
        file_menu.add_command(label="Kaydet", command=self.save_current_dataframe)
        file_menu.add_command(label="Farklı Kaydet", command=self.save_as)
        file_menu.add_separator()
        file_menu.add_command(label="Çıkış", accelerator="Ctrl+Q", command=self.root.quit)
        menubar.add_cascade(label="Dosya", menu=file_menu)

        edit_menu = tk.Menu(menubar, tearoff=0)
        edit_menu.add_command(label="Yeni Kayıt", accelerator="Ctrl+N", command=self.start_new_entry)
        edit_menu.add_command(label="Kaydı Düzenle", command=self.populate_form_from_selection)
        edit_menu.add_command(label="Kaydı Sil", command=self.delete_data)
        edit_menu.add_separator()
        edit_menu.add_command(label="Yenile", accelerator="F5", command=self.load_data)
        menubar.add_cascade(label="Düzenle", menu=edit_menu)

        report_menu = tk.Menu(menubar, tearoff=0)
        report_menu.add_command(label="Satış Raporu", command=self.generate_report)
        report_menu.add_command(label="Aylık Özet", command=self.generate_report)
        report_menu.add_command(label="Satış Elemanı Raporu", command=self.generate_report)
        menubar.add_cascade(label="Raporlar", menu=report_menu)

        help_menu = tk.Menu(menubar, tearoff=0)
        help_menu.add_command(label="Kullanım Kılavuzu", command=self.show_help)
        help_menu.add_command(label="Hakkında", command=self.show_about)
        menubar.add_cascade(label="Yardım", menu=help_menu)

        theme_menu = tk.Menu(menubar, tearoff=0)
        self._theme_var = tk.StringVar(value=self._config.get("theme", DEFAULT_THEME))
        for key, label in THEMES.items():
            theme_menu.add_radiobutton(
                label=label,
                variable=self._theme_var,
                value=key,
                command=partial(self.set_theme, key),
            )
        menubar.add_cascade(label="Tema", menu=theme_menu)

        self.root.config(menu=menubar)

        self.root.bind("<Control-s>", lambda event: self.save_data())
        self.root.bind("<Control-S>", lambda event: self.save_data())
        self.root.bind("<Control-n>", self.start_new_entry)
        self.root.bind("<Control-N>", self.start_new_entry)
        self.root.bind("<Control-z>", lambda event: self.undo_last_change())
        self.root.bind("<Control-Z>", lambda event: self.undo_last_change())
        self.root.bind("<Control-y>", lambda event: self.redo_last_change())
        self.root.bind("<Control-Y>", lambda event: self.redo_last_change())
        self.root.bind("<Control-f>", lambda event: self.open_filter_window())
        self.root.bind("<Control-F>", lambda event: self.open_filter_window())
        self.root.bind("<Control-q>", lambda event: self.root.quit())
        self.root.bind("<Control-Q>", lambda event: self.root.quit())
        self.root.bind("<F5>", lambda event: self.load_data())

    def _create_main_frames(self) -> None:
        self.header_frame = ttk.Frame(self.root)
        self.header_frame.pack(fill="x", padx=16, pady=(16, 8))

        ttk.Label(self.header_frame, text=APP_TITLE, style="Header.TLabel").pack(side="left")
        self.file_info_var = tk.StringVar(value=f"Dosya: {DATA_FILE} | Kayıt: 0")
        self.file_info_label = ttk.Label(self.header_frame, textvariable=self.file_info_var)
        self.file_info_label.pack(side="right")
        self.logo_label = tk.Label(self.header_frame, bd=0, highlightthickness=0, borderwidth=0)
        self.logo_label.pack_forget()
        self._apply_branding(self._config.get("theme", DEFAULT_THEME))
        
        self.content_frame = ttk.Frame(self.root)
        self.content_frame.pack(fill="both", expand=True, padx=16, pady=8)

        self.form_frame = ttk.LabelFrame(self.content_frame, text="Veri Giriş Formu")
        self.form_frame.pack(side="left", fill="y", padx=(0, 8), pady=4)
        self.form_frame.configure(width=420)
        self.form_frame.pack_propagate(False)

        self.table_frame = ttk.LabelFrame(self.content_frame, text="Kayıtlı Veriler")
        self.table_frame.pack(side="right", fill="both", expand=True, pady=4)

    def _create_form(self) -> None:
        self.form_vars: Dict[str, tk.Variable] = {}
        self.date_entries: Dict[str, DateEntry] = {}
        self._form_widgets: List[Tuple[tk.Widget, str]] = []

        # Üst kısımda yer alan aksiyon butonları
        top_action_frame = ttk.Frame(self.form_frame)
        top_action_frame.pack(fill="x", pady=(0, 12))
        for column in range(4):
            top_action_frame.columnconfigure(column, weight=1)

        self.new_entry_button = ttk.Button(
            top_action_frame,
            text="＋ Yeni Kayıt",
            style="ActionPrimary.TButton",
            command=self.start_new_entry,
        )
        self.new_entry_button.grid(row=0, column=0, sticky="ew", padx=4)
        self.new_entry_button.configure(width=FORM_BUTTON_WIDTH)

        self.update_button = ttk.Button(
            top_action_frame,
            text="✏️ Güncelle",
            style="ActionPrimary.TButton",
            command=self.update_data,
        )
        self.update_button.grid(row=0, column=1, sticky="ew", padx=4)
        self.update_button.configure(width=FORM_BUTTON_WIDTH)

        self.undo_button = ttk.Button(
            top_action_frame,
            text="↩",
            style="Action.TButton",
            command=self.undo_last_change,
        )
        self.undo_button.grid(row=0, column=2, sticky="ew", padx=4)
        self.undo_button.configure(width=6)

        self.redo_button = ttk.Button(
            top_action_frame,
            text="↪",
            style="Action.TButton",
            command=self.redo_last_change,
        )
        self.redo_button.grid(row=0, column=3, sticky="ew", padx=4)
        self.redo_button.configure(width=6)

        def create_labeled_row(parent, label_text, widget_factory):
            """Create a labelled row and ensure the widget is packed correctly."""

            row = ttk.Frame(parent)
            row.pack(fill="x", pady=3)
            ttk.Label(row, text=label_text, width=22).pack(side="left")

            widget = widget_factory(row)
            if isinstance(widget, tk.Widget):
                widget.pack(side="left", fill="x", expand=True)
                self._register_form_widget(widget)
            return widget

        today = datetime.today()

        # Tarih alanları
        self.form_vars["Date of Request"] = tk.StringVar()
        self.form_vars["Date of Issue"] = tk.StringVar()
        self.form_vars["Date of Delivery"] = tk.StringVar()

        date_labels = {
            "Date of Request": "Talep Tarihi",
            "Date of Issue": "Düzenleme Tarihi",
            "Date of Delivery": "Teslimat Tarihi",
        }

        for field in ("Date of Request", "Date of Issue", "Date of Delivery"):
            var = self.form_vars[field]

            def widget_factory(parent, v=var):
                entry = DateEntry(
                    parent,
                    textvariable=v,
                    date_pattern="dd.mm.yyyy",
                    locale="tr_TR",
                    font=("Segoe UI", 10),
                    width=18,
                )
                entry.set_date(today)
                return entry

            widget = create_labeled_row(self.form_frame, date_labels[field], widget_factory)
            if isinstance(widget, DateEntry):
                self.date_entries[field] = widget

        self.form_vars["Date of Request"].trace_add("write", self._handle_request_date_change)
        self._handle_request_date_change()

        # Satış bilgileri
        salesman_options = self._get_sales_rep_options()
        default_salesman = salesman_options[0] if salesman_options else ""
        self.form_vars["Sales Man"] = tk.StringVar(value=default_salesman)

        def salesman_factory(parent: tk.Widget) -> ttk.Combobox:
            combo = ttk.Combobox(
                parent,
                textvariable=self.form_vars["Sales Man"],
                values=self._get_sales_rep_options(),
                state="readonly",
            )
            return combo

        self.salesman_combo = create_labeled_row(self.form_frame, "Satış Elemanı", salesman_factory)
        self.salesman_combo.bind("<<ComboboxSelected>>", self._handle_salesman_selection)

        text_fields = [
            ("Customer Name", "Müşteri Adı"),
            ("Customer PO No", "Müşteri Sipariş No"),
            ("Definition", "Ürün Tanımı"),
            ("Sales Ticket Reference", "SalesForce Ref"),
            ("SO No", "SO No"),
            ("PTD PO No", "PTD PO No"),
            ("NON-EDI PO No", "NON-EDI PO No"),
        ]
        for field, label in text_fields:
            self.form_vars[field] = tk.StringVar()
            create_labeled_row(
                self.form_frame,
                label,
                lambda parent, v=self.form_vars[field]: ttk.Entry(parent, textvariable=v),
            )

        # Finansal alanlar
        self.form_vars["Amount"] = tk.StringVar()
        self.form_vars["DiscountPercent"] = tk.StringVar()
        self.form_vars["CPS"] = tk.StringVar()
        self.form_vars["Invoiced Amount"] = tk.StringVar()

        self.amount_entry = create_labeled_row(
            self.form_frame,
            "Tutar (€)",
            lambda parent: ttk.Entry(parent, textvariable=self.form_vars["Amount"]),
        )
        self.amount_entry.bind("<FocusIn>", lambda _event: self._on_currency_focus_in("Amount"))
        self.amount_entry.bind("<FocusOut>", lambda _event: self._format_currency_entry("Amount"))
        self.discount_entry = create_labeled_row(
            self.form_frame,
            "İndirim (%)",
            lambda parent: ttk.Entry(parent, textvariable=self.form_vars["DiscountPercent"]),
        )
        self.cps_entry = create_labeled_row(
            self.form_frame,
            "CPS (€)",
            lambda parent: ttk.Entry(parent, textvariable=self.form_vars["CPS"]),
        )
        self.cps_entry.bind("<FocusIn>", lambda _event: self._on_currency_focus_in("CPS"))
        self.cps_entry.bind("<FocusOut>", lambda _event: self._format_currency_entry("CPS"))
        self.cpi_total_entry = create_labeled_row(
            self.form_frame,
            "CPI Tutarı (€)",
            lambda parent: ttk.Entry(
                parent,
                textvariable=self.form_vars["Invoiced Amount"],
                state="readonly",
            ),
        )

        self.form_vars["Delivery Note"] = tk.StringVar()
        notes_row = ttk.Frame(self.form_frame)
        notes_row.pack(fill="both", pady=3)
        ttk.Label(notes_row, text="Notlar", width=22).pack(side="left", anchor="n")
        self.notes_text = tk.Text(notes_row, height=2, wrap="word", font=("Segoe UI", 10))
        self.notes_text.pack(side="left", fill="both", expand=True)
        self._register_form_widget(self.notes_text)

        self.discount_entry.bind("<FocusIn>", self._on_discount_focus_in)
        self.discount_entry.bind("<FocusOut>", self._format_discount_entry)
        self.form_vars["Amount"].trace_add("write", lambda *args: self._update_cpi_field())
        self.form_vars["CPS"].trace_add("write", lambda *args: self._update_cpi_field())
        self._update_cpi_field()

        # Boolean seçenekler
        self.form_vars["QI Forecast"] = tk.StringVar(value="NO")
        self.form_vars["Invoiced"] = tk.StringVar(value="NO")

        def create_radio(label_text: str, var_name: str) -> None:
            container = ttk.Frame(self.form_frame)
            container.pack(fill="x", pady=3)
            ttk.Label(container, text=label_text, width=22).pack(side="left")
            for choice in ("YES", "NO"):
                radio = ttk.Radiobutton(
                    container,
                    text=choice,
                    value=choice,
                    variable=self.form_vars[var_name],
                )
                radio.pack(side="left", padx=2)
                self._register_form_widget(radio)

        create_radio("QI Forecast", "QI Forecast")
        create_radio("Faturalandi", "Invoiced")

        # Aksiyon butonları
        action_frame = ttk.LabelFrame(self.form_frame, text="Kayıt İşlemleri")
        action_frame.pack(fill="x", pady=(18, 0))
        for column in range(3):
            action_frame.columnconfigure(column, weight=1)

        self.save_button = ttk.Button(
            action_frame,
            text="💾 Kaydet",
            style="ActionAccent.TButton",
            command=self.save_data,
        )
        self.save_button.grid(row=0, column=0, sticky="ew", padx=4, pady=4)

        ttk.Button(
            action_frame,
            text="🗑 Sil",
            style="ActionDanger.TButton",
            command=self.delete_data,
        ).grid(row=0, column=1, sticky="ew", padx=4, pady=4)

        ttk.Button(
            action_frame,
            text="🧹 Temizle",
            style="Action.TButton",
            command=self.reset_form,
        ).grid(row=0, column=2, sticky="ew", padx=4, pady=4)

        self._apply_form_state()
        self._update_button_states()

    def _register_form_widget(self, widget: tk.Widget) -> None:
        if widget is None:
            return
        try:
            state = widget.cget("state")
        except tk.TclError:
            state = "normal"
        if not state:
            state = "normal"
        if not hasattr(self, "_form_widgets"):
            self._form_widgets = []
        self._form_widgets.append((widget, state))

    def _set_form_state(self, enabled: bool) -> None:
        self._form_enabled = enabled
        for widget, default_state in getattr(self, "_form_widgets", []):
            state_to_use = default_state if enabled else "disabled"
            if isinstance(widget, tk.Text):
                widget.configure(state="normal")
                widget.configure(state=state_to_use if enabled else "disabled")
            else:
                widget.configure(state=state_to_use)

    def _apply_form_state(self) -> None:
        should_enable = self._new_entry_mode or self.selected_index is not None
        self._set_form_state(should_enable)

    def _create_table(self) -> None:
        columns = ["#", *COLUMNS]
        tree_container = ttk.Frame(self.table_frame)
        tree_container.pack(fill="both", expand=True)

        self.tree = ttk.Treeview(
            tree_container,
            columns=columns,
            show="headings",
            style="Custom.Treeview",
        )
        self.tree.pack(side="left", fill="both", expand=True)

        self.tree.tag_configure("invoiced", background="#d1fae5")

        display_names = {
            "Invoiced Amount": "CPI Tutarı",
            "Delivery Note": "Notlar",
            "Sales Ticket Reference": "SalesForce Ref",
        }
        for col in columns:
            heading_text = display_names.get(col, col)
            self.tree.heading(col, text=heading_text, command=partial(self.sort_by_column, col))
            self.tree.column(col, width=120, anchor="center")
        self.tree.column("#", width=60, anchor="center")
        self.tree.column("Definition", width=200, anchor="w")
        self.tree.column("Delivery Note", width=220, anchor="w")

        scrollbar_y = ttk.Scrollbar(tree_container, orient="vertical", command=self.tree.yview)
        scrollbar_y.pack(side="right", fill="y")
        self.tree.configure(yscrollcommand=scrollbar_y.set)

        scrollbar_x = ttk.Scrollbar(self.table_frame, orient="horizontal", command=self.tree.xview)
        scrollbar_x.pack(fill="x")
        self.tree.configure(xscrollcommand=scrollbar_x.set)

        self.tree.bind("<<TreeviewSelect>>", lambda event: self.populate_form_from_selection())
        self.tree.bind("<Button-3>", self._show_context_menu)

        pagination_frame = ttk.Frame(self.table_frame)
        pagination_frame.pack(fill="x", pady=6)
        ttk.Button(pagination_frame, text="◀ Önceki", command=self.prev_page).pack(side="left")
        self.page_var = tk.StringVar(value="Sayfa 1/1")
        ttk.Label(pagination_frame, textvariable=self.page_var).pack(side="left", padx=8)
        ttk.Button(pagination_frame, text="Sonraki ▶", command=self.next_page).pack(side="left")

    def _create_action_buttons(self) -> None:
        top_buttons = ttk.Frame(self.header_frame)
        top_buttons.pack(side="right", padx=(0, 12))
        ttk.Button(
            top_buttons,
            text="📊 Rapor Oluştur",
            style="Accent.TButton",
            command=self.generate_report,
        ).pack(side="left", padx=(0, 4))

        spacer = ttk.Frame(top_buttons, width=16)
        spacer.pack(side="left")
        spacer.pack_propagate(False)

        ttk.Button(top_buttons, text="Ara / Filtrele", command=self.open_filter_window).pack(
            side="left", padx=(12, 4)
        )
        ttk.Button(top_buttons, text="Satış Elemanları", command=self.open_sales_rep_manager).pack(side="left", padx=4)
        ttk.Button(top_buttons, text="Dışa Aktar", command=self.export_filtered_data).pack(side="left", padx=4)

    def _create_bottom_buttons(self) -> None:
        bottom_frame = ttk.Frame(self.root)
        bottom_frame.pack(fill="x", padx=16, pady=8)
        quick_actions = ttk.Frame(bottom_frame)
        quick_actions.pack(side="left")

        self.quick_save_button = ttk.Button(
            quick_actions,
            text="💾 Kaydet",
            style="ActionAccent.TButton",
            command=self.save_data,
            width=FORM_BUTTON_WIDTH,
        )
        self.quick_save_button.pack(side="left", padx=4)

        self.quick_delete_button = ttk.Button(
            quick_actions,
            text="🗑 Sil",
            style="ActionDanger.TButton",
            command=self.delete_data,
            width=FORM_BUTTON_WIDTH,
        )
        self.quick_delete_button.pack(side="left", padx=4)

        self.quick_reset_button = ttk.Button(
            quick_actions,
            text="🧹 Temizle",
            style="Action.TButton",
            command=self.reset_form,
            width=FORM_BUTTON_WIDTH,
        )
        self.quick_reset_button.pack(side="left", padx=4)

        ttk.Button(bottom_frame, text="Yedeklemeyi Aç", command=self.open_backup_directory).pack(side="right")

        self._update_button_states()

    # ----------------------------------------------------------------- helpers
    def _set_date_field(self, field: str, date_value: datetime) -> None:
        entry = self.date_entries.get(field)
        if entry:
            entry.set_date(date_value)
        else:
            self.form_vars[field].set(date_value.strftime("%d.%m.%Y"))

    def _handle_request_date_change(self, *_args) -> None:
        if self._suspend_delivery_autofill:
            return
        request_value = self.form_vars["Date of Request"].get().strip()
        if not request_value:
            return
        request_date = self._parse_date_str(request_value)
        if request_date is None:
            return
        if isinstance(request_date, pd.Timestamp):
            request_date = request_date.to_pydatetime()
        if not isinstance(request_date, datetime):
            return
        if request_date.tzinfo is not None:
            request_date = request_date.replace(tzinfo=None)
        delivery_date = request_date + timedelta(weeks=8)
        self._set_date_field("Date of Delivery", delivery_date)

    def _update_status(self, text: str) -> None:
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.status_var.set(f"[{timestamp}] {text}")

    def _push_history(self) -> None:
        self.history.append(self.df.copy(deep=True))
        if len(self.history) > 20:
            self.history = self.history[-20:]
        self.redo_stack.clear()
        self._update_button_states()

    def _clear_history(self) -> None:
        self.history.clear()
        self.redo_stack.clear()
        self._update_button_states()

    def _update_button_states(self) -> None:
        if not hasattr(self, "save_button"):
            return

        if self._new_entry_mode and self.selected_index is None:
            self.save_button.state(["!disabled"])
        else:
            self.save_button.state(["disabled"])

        if hasattr(self, "quick_save_button"):
            if self._new_entry_mode and self.selected_index is None:
                self.quick_save_button.state(["!disabled"])
            else:
                self.quick_save_button.state(["disabled"])

        if self.selected_index is not None:
            self.update_button.state(["!disabled"])
        else:
            self.update_button.state(["disabled"])

        if hasattr(self, "quick_delete_button"):
            if self.selected_index is not None:
                self.quick_delete_button.state(["!disabled"])
            else:
                self.quick_delete_button.state(["disabled"])

        if self.history:
            self.undo_button.state(["!disabled"])
        else:
            self.undo_button.state(["disabled"])

        if self.redo_stack:
            self.redo_button.state(["!disabled"])
        else:
            self.redo_button.state(["disabled"])

    def start_new_entry(self, _event=None) -> None:
        self._new_entry_mode = True
        self.reset_form(preserve_new_mode=True)
        self._update_status("Yeni kayıt için form hazırlandı")
        self._update_button_states()

    def undo_last_change(self) -> None:
        if not self.history:
            messagebox.showinfo("Bilgi", "Geri alınacak bir değişiklik yok")
            return
        self.redo_stack.append(self.df.copy(deep=True))
        self.df = self.history.pop()
        self.save_current_dataframe()
        self.apply_filters()
        self.reset_form()
        self._update_status("Son değişiklik geri alındı")
        self._update_button_states()

    def redo_last_change(self) -> None:
        if not self.redo_stack:
            messagebox.showinfo("Bilgi", "İleri alınacak bir değişiklik yok")
            return
        self.history.append(self.df.copy(deep=True))
        if len(self.history) > 20:
            self.history = self.history[-20:]
        self.df = self.redo_stack.pop()
        self.save_current_dataframe()
        self.apply_filters()
        self.reset_form()
        self._update_status("İleri alma işlemi uygulandı")
        self._update_button_states()

    def set_theme(self, theme: str) -> None:
        self._config["theme"] = theme
        self._save_config(self._config)
        self._apply_theme(theme)

    def _handle_salesman_selection(self, _event) -> None:
        if self.form_vars["Sales Man"].get() == OTHER_SALES_REP_OPTION:
            new_name = simpledialog.askstring("Yeni Satış Elemanı", "İsim giriniz", parent=self.root)
            if new_name:
                new_name = new_name.strip()
                if new_name and new_name not in self.sales_reps:
                    self.sales_reps.append(new_name)
                    self._save_sales_reps()
                if new_name:
                    self.salesman_combo.config(values=self._get_sales_rep_options())
                    self.form_vars["Sales Man"].set(new_name)
                    return
            options = self._get_sales_rep_options()
            if options:
                self.form_vars["Sales Man"].set(options[0])

    def open_sales_rep_manager(self) -> None:
        password = simpledialog.askstring(
            "Satış Elemanları", "Şifreyi giriniz", show="*", parent=self.root
        )
        if password is None:
            return
        if password != self.sales_rep_password:
            messagebox.showerror("Yetkisiz İşlem", "Şifre hatalı")
            return

        if hasattr(self, "_sales_rep_window") and self._sales_rep_window.winfo_exists():
            self._sales_rep_window.lift()
            return

        window = tk.Toplevel(self.root)
        window.title("Satış Elemanları")
        window.geometry("360x400")
        window.transient(self.root)
        window.grab_set()
        self._sales_rep_window = window

        ttk.Label(window, text="Satış Mühendisi Listesi", style="Header.TLabel").pack(pady=(12, 8))

        listbox = tk.Listbox(window, height=12, selectmode="extended")
        listbox.pack(fill="both", expand=True, padx=16)
        for rep in self.sales_reps:
            listbox.insert("end", rep)

        entry_frame = ttk.Frame(window)
        entry_frame.pack(fill="x", padx=16, pady=8)
        ttk.Label(entry_frame, text="Yeni İsim").pack(anchor="w")
        new_rep_var = tk.StringVar()
        ttk.Entry(entry_frame, textvariable=new_rep_var).pack(fill="x", pady=(2, 0))

        def populate_entry_from_selection(event: tk.Event) -> None:
            selected = listbox.curselection()
            if len(selected) == 1:
                new_rep_var.set(listbox.get(selected[0]))

        listbox.bind("<<ListboxSelect>>", populate_entry_from_selection)

        def add_rep() -> None:
            name = new_rep_var.get().strip()
            if not name:
                return
            if name in listbox.get(0, "end"):
                messagebox.showinfo("Bilgi", "Bu isim zaten listede", parent=window)
                return
            listbox.insert("end", name)
            new_rep_var.set("")

        def remove_selected() -> None:
            selected = listbox.curselection()
            if not selected:
                return
            for index in reversed(selected):
                listbox.delete(index)

        def edit_selected() -> None:
            selected = listbox.curselection()
            if len(selected) != 1:
                messagebox.showinfo(
                    "Bilgi", "Lütfen düzenlemek için tek bir isim seçiniz", parent=window
                )
                return
            new_name = new_rep_var.get().strip()
            if not new_name:
                messagebox.showinfo(
                    "Bilgi", "Yeni isim alanı boş olamaz", parent=window
                )
                return
            current_index = selected[0]
            current_name = listbox.get(current_index)
            existing_names = listbox.get(0, "end")
            if new_name in existing_names and new_name != current_name:
                messagebox.showinfo(
                    "Bilgi", "Bu isim zaten listede", parent=window
                )
                return
            listbox.delete(current_index)
            listbox.insert(current_index, new_name)
            listbox.selection_set(current_index)

        button_frame = ttk.Frame(window)
        button_frame.pack(fill="x", padx=16, pady=4)
        ttk.Button(button_frame, text="Ekle", command=add_rep).pack(
            side="left", expand=True, padx=4
        )
        ttk.Button(button_frame, text="Düzenle", command=edit_selected).pack(
            side="left", expand=True, padx=4
        )
        ttk.Button(button_frame, text="Sil", command=remove_selected).pack(
            side="left", expand=True, padx=4
        )

        def change_password() -> None:
            current = simpledialog.askstring(
                "Şifreyi Değiştir",
                "Mevcut şifreyi giriniz",
                show="*",
                parent=window,
            )
            if current is None:
                return
            if current != self.sales_rep_password:
                messagebox.showerror("Hata", "Mevcut şifre yanlış", parent=window)
                return
            new_password = simpledialog.askstring(
                "Şifreyi Değiştir",
                "Yeni şifreyi giriniz",
                show="*",
                parent=window,
            )
            if new_password is None:
                return
            new_password = new_password.strip()
            if not new_password:
                messagebox.showinfo("Bilgi", "Şifre değişmedi", parent=window)
                return
            confirm_password = simpledialog.askstring(
                "Şifreyi Değiştir",
                "Yeni şifreyi tekrar giriniz",
                show="*",
                parent=window,
            )
            if confirm_password is None:
                return
            confirm_password = confirm_password.strip()
            if new_password != confirm_password:
                messagebox.showerror(
                    "Hata", "Yeni şifreler eşleşmiyor", parent=window
                )
                return
            self._update_sales_rep_password(new_password)
            messagebox.showinfo(
                "Başarılı", "Şifre güncellendi", parent=window
            )

        def save_and_close() -> None:
            raw_reps = [listbox.get(i) for i in range(listbox.size())]
            unique_reps: List[str] = []
            for rep in raw_reps:
                cleaned = rep.strip()
                if cleaned and cleaned not in unique_reps:
                    unique_reps.append(cleaned)
            self.sales_reps = unique_reps
            self._save_sales_reps()
            options = self._get_sales_rep_options()
            self.salesman_combo.config(values=options)
            current = self.form_vars["Sales Man"].get()
            if current not in options:
                self.form_vars["Sales Man"].set(options[0] if options else "")
            if self.filter_options.salesman and self.filter_options.salesman not in self.sales_reps:
                self.filter_options.salesman = ""
            window.destroy()

        def cancel() -> None:
            window.destroy()

        action_frame = ttk.Frame(window)
        action_frame.pack(fill="x", padx=16, pady=(8, 12))
        ttk.Button(
            action_frame,
            text="Şifreyi Değiştir",
            command=change_password,
        ).pack(side="left", expand=True, fill="x", padx=4)
        ttk.Button(action_frame, text="Kaydet", style="Accent.TButton", command=save_and_close).pack(
            side="left", expand=True, fill="x", padx=4
        )
        ttk.Button(action_frame, text="İptal", command=cancel).pack(
            side="left", expand=True, fill="x", padx=4
        )

    def _show_context_menu(self, event) -> None:
        menu = tk.Menu(self.root, tearoff=0)
        menu.add_command(label="Düzenle", command=self.populate_form_from_selection)
        menu.add_command(label="Sil", command=self.delete_data)
        menu.add_command(label="Kopyala", command=self.copy_selected_row)
        menu.add_command(label="Detay", command=self.show_detail_popup)
        menu.tk_popup(event.x_root, event.y_root)

    # ----------------------------------------------------------------- data i/o
    def load_data(self) -> None:
        try:
            self.df = pd.read_excel(DATA_FILE)
        except FileNotFoundError:
            self._ensure_excel_file()
            self.df = pd.read_excel(DATA_FILE)
        except Exception as exc:
            messagebox.showerror("Hata", f"Veri yüklenemedi: {exc}")
            return
        if "PO No" in self.df.columns and "PTD PO No" not in self.df.columns:
            self.df = self.df.rename(columns={"PO No": "PTD PO No"})
        for column in COLUMNS:
            if column not in self.df.columns:
                self.df[column] = ""
        extra_columns = [col for col in self.df.columns if col not in COLUMNS]
        if extra_columns:
            self.df = self.df.drop(columns=extra_columns)
        self.df = self.df[COLUMNS]
        self._normalise_discount_values()
        self._clear_history()
        self.apply_filters()
        self._update_status("Veri yüklendi")

    def save_current_dataframe(self) -> None:
        try:
            self.df.to_excel(DATA_FILE, index=False)
            self._update_status("Dosya kaydedildi")
        except Exception as exc:
            messagebox.showerror("Kaydetme Hatası", str(exc))

    def _normalise_discount_values(self) -> None:
        if "Total Discount" not in self.df.columns:
            return
        updated = False
        for idx, row in self.df.iterrows():
            discount_raw = self._to_float(row.get("Total Discount"))

            if discount_raw is not None:
                if discount_raw > 1:
                    amount = self._to_float(row.get("Amount"))
                    if amount:
                        fraction = discount_raw / amount
                    else:
                        fraction = discount_raw / 100
                else:
                    fraction = discount_raw
                fraction = max(0.0, min(fraction, 1.0))
                if discount_raw != fraction:
                    self.df.at[idx, "Total Discount"] = fraction
                    updated = True
        if updated:
            self._update_status("İndirim verileri güncellendi")

    def save_as(self) -> None:
        filename = filedialog.asksaveasfilename(
            title="Farklı Kaydet",
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx")],
        )
        if not filename:
            return
        try:
            self.df.to_excel(filename, index=False)
            self._update_status(f"Dosya kaydedildi: {filename}")
        except Exception as exc:
            messagebox.showerror("Hata", str(exc))

    def create_new_file(self) -> None:
        if messagebox.askyesno("Onay", "Yeni bir dosya oluşturmak istediğinize emin misiniz?"):
            self.df = pd.DataFrame(columns=COLUMNS)
            self.save_current_dataframe()
            self.load_data()

    def open_existing_file(self) -> None:
        filename = filedialog.askopenfilename(title="Excel Dosyası", filetypes=[("Excel", "*.xlsx")])
        if not filename:
            return
        global DATA_FILE
        DATA_FILE = filename
        self.file_info_var.set(f"Dosya: {DATA_FILE}")
        self.load_data()

    def open_backup_directory(self) -> None:
        os.startfile(BACKUP_DIR) if os.name == "nt" else os.system(f"xdg-open '{BACKUP_DIR}'")

    def export_filtered_data(self) -> None:
        filename = filedialog.asksaveasfilename(
            title="Dışa Aktar",
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx")],
        )
        if not filename:
            return
        try:
            filtered_df = self.get_filtered_dataframe()
            filtered_df.to_excel(filename, index=False)
            self._update_status(f"Dışa aktarıldı: {filename}")
        except Exception as exc:
            messagebox.showerror("Dışa Aktarım Hatası", str(exc))

    # ----------------------------------------------------------------- backups
    def schedule_auto_backup(self) -> None:
        self.root.after(AUTO_SAVE_INTERVAL, self.perform_backup)

    def perform_backup(self) -> None:
        if self.df.empty:
            self.schedule_auto_backup()
            return
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        backup_path = BACKUP_DIR / f"backup_{timestamp}.xlsx"
        try:
            self.df.to_excel(backup_path, index=False)
            self._update_status(f"Yedek oluşturuldu: {backup_path.name}")
        except Exception as exc:
            messagebox.showwarning("Yedekleme Hatası", str(exc))
        finally:
            self.schedule_auto_backup()

    # --------------------------------------------------------------- table ops
    def update_table(self, dataframe: Optional[pd.DataFrame] = None) -> None:
        df = dataframe if dataframe is not None else self.df
        for item in self.tree.get_children():
            self.tree.delete(item)

        if df.empty:
            self.page_var.set("Sayfa 1/1")
            self.file_info_var.set(f"Dosya: {DATA_FILE} | Kayıt: 0")
            return

        self.total_pages = max(1, (len(df) + PAGE_SIZE - 1) // PAGE_SIZE)
        self.current_page = min(self.current_page, self.total_pages)
        start = (self.current_page - 1) * PAGE_SIZE
        end = start + PAGE_SIZE
        page_df = df.iloc[start:end]

        for idx, (_, row) in enumerate(page_df.iterrows(), start=start + 1):
            formatted_values: List[str] = []
            for col in COLUMNS:
                if col == "Total Discount":
                    formatted_values.append(self._format_discount_fraction(row[col]))
                elif col in CURRENCY_FIELDS:
                    formatted_values.append(self._format_currency(row[col]))
                else:
                    formatted_values.append(self._format_value(row[col]))
            values = [idx] + formatted_values
            tags: Tuple[str, ...] = ()
            if str(row.get("Invoiced", "")).upper() == "YES":
                tags = ("invoiced",)
            self.tree.insert("", "end", values=values, tags=tags)

        self.page_var.set(f"Sayfa {self.current_page}/{self.total_pages}")
        self.file_info_var.set(f"Dosya: {DATA_FILE} | Kayıt: {len(self.df)}")

    def next_page(self) -> None:
        if self.current_page < self.total_pages:
            self.current_page += 1
            self.update_table(self.get_filtered_dataframe())

    def prev_page(self) -> None:
        if self.current_page > 1:
            self.current_page -= 1
            self.update_table(self.get_filtered_dataframe())

    def sort_by_column(self, column: str) -> None:
        if column == "#":
            return
        df = self.get_filtered_dataframe().sort_values(by=column, na_position="last")
        self.update_table(df)

    def get_filtered_dataframe(self) -> pd.DataFrame:
        df = self.df.copy()
        opts = self.filter_options
        if opts.search_text:
            df = df[df["Customer Name"].str.contains(opts.search_text, case=False, na=False)]
        if opts.so_no:
            df = df[df["SO No"].astype(str).str.contains(opts.so_no, case=False, na=False)]
        if opts.salesman:
            df = df[df["Sales Man"] == opts.salesman]
        if opts.invoiced:
            df = df[df["Invoiced"].str.upper() == opts.invoiced.upper()]
        if opts.start_date:
            df = df[pd.to_datetime(df["Date of Issue"], errors="coerce") >= opts.start_date]
        if opts.end_date:
            df = df[pd.to_datetime(df["Date of Issue"], errors="coerce") <= opts.end_date]
        return df

    def apply_filters(self) -> None:
        self.current_page = 1
        self.update_table(self.get_filtered_dataframe())

    # --------------------------------------------------------------- form logic
    def reset_form(self, _event=None, *, preserve_new_mode: bool = False) -> None:
        for key, var in self.form_vars.items():
            if isinstance(var, tk.StringVar):
                if key in ("QI Forecast", "Invoiced"):
                    var.set("NO")
                elif key == "Sales Man":
                    options = self._get_sales_rep_options()
                    var.set(options[0] if options else "")
                elif key in ("Date of Request", "Date of Issue", "Date of Delivery"):
                    continue
                else:
                    var.set("")
        if hasattr(self, "notes_text"):
            self.notes_text.configure(state="normal")
            self.notes_text.delete("1.0", "end")
            if "Delivery Note" in self.form_vars:
                self.form_vars["Delivery Note"].set("")
        self.selected_index = None
        self.tree.selection_remove(self.tree.selection())
        today = datetime.today()
        delivery_date = today + timedelta(weeks=8)
        self._suspend_delivery_autofill = True
        self._set_date_field("Date of Request", today)
        self._set_date_field("Date of Issue", today)
        self._set_date_field("Date of Delivery", delivery_date)
        self._suspend_delivery_autofill = False
        self._update_cpi_field()
        if not preserve_new_mode:
            self._new_entry_mode = False
        self._update_status("Form temizlendi")
        self._apply_form_state()
        self._update_button_states()

    def populate_form_from_selection(self) -> None:
        selected = self.tree.selection()
        if not selected:
            return
        item = self.tree.item(selected[0])
        values = item["values"][1:]
        if len(values) != len(COLUMNS):
            return
        row_data = dict(zip(COLUMNS, values))
        if hasattr(self, "notes_text"):
            self.notes_text.configure(state="normal")
        self._suspend_delivery_autofill = True
        try:
            for col, value in row_data.items():
                if col in ("QI Forecast", "Invoiced"):
                    self.form_vars[col].set(str(value).upper())
                elif col in ("Date of Request", "Date of Issue", "Date of Delivery"):
                    try:
                        dt = pd.to_datetime(value)
                        if pd.isna(dt):
                            raise ValueError
                        if isinstance(dt, pd.Timestamp):
                            dt = dt.to_pydatetime()
                        if isinstance(dt, datetime) and dt.tzinfo is not None:
                            dt = dt.replace(tzinfo=None)
                        self._set_date_field(col, dt)
                    except Exception:
                        self.form_vars[col].set(value)
                elif col == "Delivery Note":
                    if value is None or pd.isna(value):
                        text_value = ""
                    else:
                        text_value = str(value)
                    self.form_vars[col].set(text_value)
                    if hasattr(self, "notes_text"):
                        self.notes_text.delete("1.0", "end")
                        self.notes_text.insert("1.0", text_value)
                elif col in CURRENCY_FIELDS and col in self.form_vars:
                    self.form_vars[col].set(self._format_currency(value))
                elif col in self.form_vars:
                    self.form_vars[col].set(str(value) if value is not None else "")
        finally:
            self._suspend_delivery_autofill = False

        amount = self._to_float(row_data.get("Amount")) or 0
        discount_value = self._to_float(row_data.get("Total Discount"))
        if discount_value is None:
            self.form_vars["DiscountPercent"].set("")
        else:
            if discount_value > 1 and amount:
                percent = (discount_value / amount) * 100
            elif discount_value > 1:
                percent = discount_value
            else:
                percent = discount_value * 100
            self.form_vars["DiscountPercent"].set(self._format_percent(percent))
        self._update_cpi_field()
        try:
            row_index = int(item["values"][0]) - 1
        except Exception:
            row_index = None
        self.selected_index = row_index
        self._update_status("Kayıt düzenleme için yüklendi")
        self._new_entry_mode = False
        self._apply_form_state()
        self._update_button_states()

    def _normalise_currency_value(self, value) -> Optional[Decimal]:
        if isinstance(value, Decimal):
            target = value
        else:
            numeric = self._to_float(value)
            if numeric is None:
                return None
            target = Decimal(str(numeric))
        try:
            return target.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
        except (InvalidOperation, ValueError):
            return None

    def _format_currency(self, value) -> str:
        normalised = self._normalise_currency_value(value)
        if normalised is None:
            return ""
        formatted = f"{normalised:.2f}".replace(".", ",")
        return f"{formatted} €"

    def _on_currency_focus_in(self, field: str) -> None:
        value = self.form_vars[field].get().strip()
        if value.endswith("€"):
            self.form_vars[field].set(value[:-1].strip())

    def _format_currency_entry(self, field: str) -> None:
        raw_value = self.form_vars[field].get()
        amount = self._parse_float(raw_value)
        if amount is None:
            cleaned = raw_value.strip()
            if cleaned.endswith("€"):
                cleaned = cleaned[:-1].strip()
            if not cleaned:
                self.form_vars[field].set("")
            return
        formatted = self._format_currency(amount)
        self.form_vars[field].set(formatted)

    def _format_value(self, value) -> str:
        if pd.isna(value):
            return ""
        if isinstance(value, pd.Timestamp):
            if pd.isna(value):
                return ""
            return value.strftime("%d.%m.%Y")
        if isinstance(value, datetime):
            return value.strftime("%d.%m.%Y")
        if isinstance(value, date):
            return value.strftime("%d.%m.%Y")            
        if isinstance(value, (float, int)):
            return f"{value:,.2f}" if value > 999 else f"{value:.2f}"
        if isinstance(value, str):
            stripped = value.strip()
            if not stripped:
                return ""
            if any(sep in stripped for sep in ("-", ".", "/")):
                try:
                    parsed = pd.to_datetime(stripped, dayfirst=True, errors="coerce")
                except Exception:
                    parsed = None
                if parsed is not None and not pd.isna(parsed):
                    return parsed.strftime("%d.%m.%Y")
            return stripped            
        return str(value)

    def _format_discount_fraction(self, value) -> str:
        fraction = self._to_float(value)
        if fraction is None:
            return ""
        return f"{fraction:.2f}".replace(".", ",")

    def _parse_float(self, value: str) -> Optional[float]:
        if not value:
            return None
        cleaned = value.replace("€", "").replace(" ", "").replace("%", "")
        cleaned = cleaned.replace(".", "").replace(",", ".")
        try:
            return float(cleaned)
        except ValueError:
            return None

    def _format_percent(self, value: float) -> str:
        return f"%{value:.2f}".replace(".", ",")

    def _on_discount_focus_in(self, _event) -> None:
        value = self.form_vars["DiscountPercent"].get().strip()
        if value.startswith("%"):
            self.form_vars["DiscountPercent"].set(value[1:])

    def _format_discount_entry(self, _event=None) -> None:
        value = self._parse_float(self.form_vars["DiscountPercent"].get())
        if value is None:
            if not self.form_vars["DiscountPercent"].get().strip():
                self.form_vars["DiscountPercent"].set("")
            return
        self.form_vars["DiscountPercent"].set(self._format_percent(value))

    def _update_cpi_field(self) -> None:
        if self._updating_cpi_field:
            return
        self._updating_cpi_field = True
        try:
            amount_value = self._parse_float(self.form_vars["Amount"].get())
            cps_value = self._parse_float(self.form_vars["CPS"].get())
            if amount_value is None:
                self.form_vars["Invoiced Amount"].set("")
                return
            amount_decimal = self._normalise_currency_value(amount_value)
            cps_decimal = (
                self._normalise_currency_value(cps_value)
                if cps_value is not None
                else Decimal("0")
            )
            if amount_decimal is None:
                self.form_vars["Invoiced Amount"].set("")
                return
            if cps_decimal is None:
                cps_decimal = Decimal("0")
            cpi_total = (amount_decimal - cps_decimal).quantize(
                Decimal("0.01"), rounding=ROUND_HALF_UP
            )
            self.form_vars["Invoiced Amount"].set(self._format_currency(cpi_total))
        finally:
            self._updating_cpi_field = False

    def _to_float(self, value) -> Optional[float]:
        if pd.isna(value) or value == "":
            return None
        if isinstance(value, (int, float)):
            return float(value)
        return self._parse_float(str(value))

    def _collect_form_data(self) -> Tuple[Optional[pd.Series], Optional[str]]:
        data: Dict[str, object] = {}
        errors: List[str] = []

        if hasattr(self, "notes_text"):
            self.form_vars["Delivery Note"].set(
                self.notes_text.get("1.0", "end").strip()
            )

        for column in COLUMNS:
            if column in ("Total Discount", "CPI"):
                # computed later
                continue
            if column not in self.form_vars:
                data[column] = ""
                continue
            value = self.form_vars[column].get().strip()
            if column in REQUIRED_FIELDS and not value:
                errors.append(f"{column} boş bırakılamaz")
            data[column] = value

        amount_raw = self._parse_float(self.form_vars["Amount"].get())
        amount_decimal = self._normalise_currency_value(amount_raw)
        if amount_decimal is None:
            errors.append("Geçerli bir tutar girin")
        else:
            self.form_vars["Amount"].set(self._format_currency(amount_decimal))
        discount_percent_value = self._parse_float(self.form_vars["DiscountPercent"].get()) or 0.0
        self._format_discount_entry()
        if discount_percent_value < 0:
            errors.append("İndirim yüzdesi negatif olamaz")
        discount_fraction = discount_percent_value / 100
        if discount_fraction > 1:
            errors.append("İndirim 100%'ü aşamaz")
        cps_raw = self._parse_float(self.form_vars["CPS"].get())
        cps_decimal = self._normalise_currency_value(cps_raw)
        if cps_decimal is None:
            cps_decimal = Decimal("0")
        else:
            if cps_raw is not None:
                self.form_vars["CPS"].set(self._format_currency(cps_decimal))

        if errors:
            return None, "\n".join(errors)

        discount_fraction = max(0.0, min(discount_fraction, 1.0))
        amount_decimal = amount_decimal or Decimal("0")
        cpi_total = (amount_decimal - cps_decimal).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
        self.form_vars["Invoiced Amount"].set(self._format_currency(cpi_total))

        data["Amount"] = float(amount_decimal)
        data["Total Discount"] = discount_fraction
        data["CPI"] = float(cpi_total)
        data["CPS"] = float(cps_decimal)
        data["Invoiced Amount"] = float(cpi_total)
        data["QI Forecast"] = data["QI Forecast"].upper() if data["QI Forecast"] else "NO"
        data["Invoiced"] = data["Invoiced"].upper() if data["Invoiced"] else "NO"

        for date_field in ("Date of Request", "Date of Issue", "Date of Delivery"):
            raw = data.get(date_field)
            if raw:
                try:
                    data[date_field] = datetime.strptime(raw, "%d.%m.%Y")
                except ValueError:
                    try:
                        data[date_field] = pd.to_datetime(raw, dayfirst=True)
                    except Exception:
                        errors.append(f"Tarih formatı hatalı: {date_field}")

        if errors:
            return None, "\n".join(errors)

        return pd.Series(data), None

    def save_data(self) -> None:
        if not self._new_entry_mode:
            messagebox.showwarning(
                "Yeni Kayıt Gerekli",
                "Yeni veri eklemek için önce 'Yeni Sipariş Verisi' butonuna tıklayınız.",
            )
            return
        if self.selected_index is not None:
            messagebox.showwarning(
                "Yeni Kayıt", "Kayıtlı bir veri seçiliyken yeni kayıt ekleyemezsiniz."
            )
            return
        record, error = self._collect_form_data()
        if error:
            messagebox.showerror("Geçersiz Veri", error)
            return
        assert record is not None

        confirm = messagebox.askyesno(
            "Onay",
            "Yeni veri girişini onaylıyor musunuz?",
            parent=self.root,
        )
        if not confirm:
            return

        self._push_history()

        self.df = pd.concat([self.df, record.to_frame().T], ignore_index=True)
        self.save_current_dataframe()
        self.apply_filters()
        self.reset_form()
        self._update_button_states()
        messagebox.showinfo("Başarılı", "Kayıt eklendi")

    def update_data(self) -> None:
        if self.selected_index is None:
            messagebox.showwarning("Seçim Yok", "Güncellenecek kayıt seçiniz")
            return
        record, error = self._collect_form_data()
        if error:
            messagebox.showerror("Geçersiz Veri", error)
            return
        assert record is not None
        confirm = messagebox.askyesno(
            "Onay",
            "Seçili kaydı güncellemek istediğinize emin misiniz?",
            parent=self.root,
        )
        if not confirm:
            return
        self._push_history()
        self.df.loc[self.selected_index, record.index] = record.values
        self.save_current_dataframe()
        self.apply_filters()
        self.reset_form()
        self._update_button_states()
        messagebox.showinfo("Güncellendi", "Kayıt başarıyla güncellendi")

    def delete_data(self) -> None:
        selected = self.tree.selection()
        if not selected:
            messagebox.showwarning("Seçim Yok", "Silmek için kayıt seçiniz")
            return
        confirm = messagebox.askyesno(
            "Emin misiniz?",
            "Seçili kaydı silmek istediğinize emin misiniz? Bu işlem geri alınamaz.",
            parent=self.root,
            icon="warning",
        )
        if not confirm:
            return
        item = self.tree.item(selected[0])
        index = int(item["values"][0]) - 1
        self._push_history()
        self.df = self.df.drop(self.df.index[index]).reset_index(drop=True)
        self.save_current_dataframe()
        self.apply_filters()
        self.reset_form()
        self._update_status("Kayıt silindi")
        self._update_button_states()

    def copy_selected_row(self) -> None:
        selected = self.tree.selection()
        if not selected:
            return
        values = self.tree.item(selected[0])["values"]
        clipboard_text = "\t".join(str(v) for v in values)
        self.root.clipboard_clear()
        self.root.clipboard_append(clipboard_text)
        self._update_status("Satır panoya kopyalandı")

    def show_detail_popup(self) -> None:
        selected = self.tree.selection()
        if not selected:
            return
        values = self.tree.item(selected[0])["values"][1:]
        detail = tk.Toplevel(self.root)
        detail.title("Kayıt Detayı")
        detail.geometry("420x520")
        text = tk.Text(detail, wrap="word")
        text.pack(expand=True, fill="both")
        for col, value in zip(COLUMNS, values):
            text.insert("end", f"{col}: {value}\n")
        text.configure(state="disabled")

    # -------------------------------------------------------------- filter UI
    def open_filter_window(self) -> None:
        popup = tk.Toplevel(self.root)
        popup.title("Ara / Filtrele")
        popup.geometry("320x360")

        opts = self.filter_options

        ttk.Label(popup, text="Müşteri Adı").pack(pady=4)
        search_var = tk.StringVar(value=opts.search_text)
        ttk.Entry(popup, textvariable=search_var).pack(fill="x", padx=16)

        ttk.Label(popup, text="SO No").pack(pady=4)
        so_no_var = tk.StringVar(value=opts.so_no)
        ttk.Entry(popup, textvariable=so_no_var).pack(fill="x", padx=16)

        ttk.Label(popup, text="Satış Elemanı").pack(pady=4)
        salesman_var = tk.StringVar(value=opts.salesman)
        ttk.Combobox(
            popup,
            textvariable=salesman_var,
            values=["", *self.sales_reps],
            state="readonly",
        ).pack(fill="x", padx=16)

        ttk.Label(popup, text="Fatura Durumu").pack(pady=4)
        invoiced_var = tk.StringVar(value=opts.invoiced)
        ttk.Combobox(popup, textvariable=invoiced_var, values=["", "YES", "NO"], state="readonly").pack(fill="x", padx=16)

        ttk.Label(popup, text="Tarih Aralığı").pack(pady=4)
        start_var = tk.StringVar(value=opts.start_date.strftime("%d.%m.%Y") if opts.start_date else "")
        end_var = tk.StringVar(value=opts.end_date.strftime("%d.%m.%Y") if opts.end_date else "")
        ttk.Entry(popup, textvariable=start_var).pack(fill="x", padx=16)
        ttk.Entry(popup, textvariable=end_var).pack(fill="x", padx=16, pady=(0, 8))

        def apply():
            self.filter_options.search_text = search_var.get().strip()
            self.filter_options.so_no = so_no_var.get().strip()
            self.filter_options.salesman = salesman_var.get().strip()
            self.filter_options.invoiced = invoiced_var.get().strip()
            self.filter_options.start_date = self._parse_date_str(start_var.get())
            self.filter_options.end_date = self._parse_date_str(end_var.get())
            self.apply_filters()
            popup.destroy()

        def reset_filters() -> None:
            self.filter_options = FilterOptions()
            search_var.set("")
            so_no_var.set("")
            salesman_var.set("")
            invoiced_var.set("")
            start_var.set("")
            end_var.set("")
            self.apply_filters()

        button_frame = ttk.Frame(popup)
        button_frame.pack(pady=6)
        ttk.Button(button_frame, text="Uygula", command=apply).pack(side="left", padx=4)
        ttk.Button(button_frame, text="Filtreleri Temizle", command=reset_filters).pack(side="left", padx=4)
        
    def _parse_date_str(self, value: str) -> Optional[datetime]:
        if not value:
            return None
        try:
            return datetime.strptime(value, "%d.%m.%Y")
        except ValueError:
            try:
                return pd.to_datetime(value, dayfirst=True)
            except Exception:
                return None

    # ------------------------------------------------------------- reporting
    def _show_reporting_dashboard(self, df: pd.DataFrame) -> None:
        if Figure is None or FigureCanvasTkAgg is None:
            messagebox.showerror(
                "Eksik Bileşen",
                "Grafikleri görüntülemek için 'matplotlib' kütüphanesinin yüklü olması gerekir.\n"
                "Lütfen 'pip install matplotlib' komutu ile kurulumu tamamlayın.",
            )
            return
        dashboard = tk.Toplevel(self.root)
        dashboard.title("Satış Raporu Paneli")
        dashboard.geometry("1240x900")
        dashboard.minsize(1080, 720)
        settings = self._theme_settings or {}
        bg_color = settings.get("bg", COLORS["bg_light"])
        fg_color = settings.get("fg", COLORS["text_dark"])
        accent = settings.get("accent", COLORS["primary"])
        secondary = settings.get("secondary", COLORS["secondary"])
        card_bg = settings.get("card_bg", "#ffffff")
        dashboard.configure(bg=bg_color)
        dashboard.transient(self.root)

        container = ttk.Frame(dashboard)
        container.pack(fill="both", expand=True, padx=16, pady=12)
        container.grid_columnconfigure(0, weight=1)

        data = df.copy()
        for col in ("Amount", "CPI", "CPS", "Invoiced Amount"):
            if col in data:
                data[col] = pd.to_numeric(data[col], errors="coerce")
        data["Sales Man"] = data.get("Sales Man", "").fillna("Bilinmiyor").astype(str)
        data["Invoiced"] = data.get("Invoiced", "").astype(str).str.upper()
        data["Date of Delivery"] = pd.to_datetime(data.get("Date of Delivery"), errors="coerce")

        self._report_canvases.clear()

        def normalise_numeric(values: Iterable) -> List[float]:
            result: List[float] = []
            for item in values:
                if item is None or (isinstance(item, float) and pd.isna(item)):
                    result.append(0.0)
                else:
                    try:
                        result.append(float(item))
                    except (TypeError, ValueError):
                        result.append(0.0)
            return result

        def format_currency(value: float) -> str:
            if value is None or pd.isna(value):
                value = 0.0
            return self._format_currency(value)

        def build_table(parent: ttk.Frame, columns: List[str], rows: List[Tuple[str, ...]], *, height: int = 6) -> ttk.Treeview:
            tree = ttk.Treeview(parent, columns=columns, show="headings", style="Report.Treeview", height=min(max(len(rows), 1), height))
            for col in columns:
                tree.heading(col, text=col)
                anchor = "center"
                if col.lower().startswith("satış") or col.lower() in {"kategori", "ay"}:
                    anchor = "w"
                tree.column(col, anchor=anchor, width=160, stretch=True)
            for row in rows:
                tree.insert("", "end", values=row)
            tree.pack(fill="both", expand=True)
            return tree

        def render_bar_chart(parent: ttk.Frame, labels: List[str], values: List[float], *, palette: Optional[List[str]] = None) -> None:
            if not palette:
                palette = [accent, secondary, COLORS.get("info", "#3b82f6"), COLORS.get("warning", "#f59e0b")]
            fig = Figure(figsize=(4.6, 2.6), dpi=100)
            ax = fig.add_subplot(111)
            fig.patch.set_facecolor(card_bg)
            ax.set_facecolor(card_bg)
            norm_values = normalise_numeric(values)
            bars = ax.bar(labels, norm_values, color=palette[: len(labels)])
            ax.tick_params(colors=fg_color, labelrotation=0)
            for spine in ax.spines.values():
                spine.set_color(fg_color)
            ax.set_ylabel("Tutar (€)", color=fg_color)
            ax.set_title("Özet", color=fg_color, pad=8)
            max_value = max(norm_values + [0]) if norm_values else 0
            ax.set_ylim(0, max_value * 1.15 if max_value else 1)
            for bar, value in zip(bars, norm_values):
                ax.bar_label([bar], labels=[format_currency(value)], padding=4, color=fg_color, fontsize=9, rotation=90 if len(labels) > 6 else 0)
            fig.tight_layout()
            canvas = FigureCanvasTkAgg(fig, master=parent)
            canvas.draw()
            widget = canvas.get_tk_widget()
            widget.pack(fill="both", expand=True)
            self._report_canvases.append(canvas)

        def render_grouped_chart(parent: ttk.Frame, grouped_df: pd.DataFrame) -> None:
            if grouped_df.empty:
                ttk.Label(parent, text="Veri bulunamadı", style="Card.TLabel").pack(fill="both", expand=True, pady=8)
                return
            labels = grouped_df["Sales Man"].tolist()
            totals = normalise_numeric(grouped_df["Amount"].tolist())
            cpi_vals = normalise_numeric(grouped_df["CPI"].tolist())
            cps_vals = normalise_numeric(grouped_df["CPS"].tolist())
            fig = Figure(figsize=(5.2, 3.0), dpi=100)
            ax = fig.add_subplot(111)
            fig.patch.set_facecolor(card_bg)
            ax.set_facecolor(card_bg)
            x = range(len(labels))
            width = 0.25
            bars1 = ax.bar([pos - width for pos in x], totals, width=width, color=accent, label="Toplam")
            bars2 = ax.bar(x, cpi_vals, width=width, color=secondary, label="CPI")
            bars3 = ax.bar([pos + width for pos in x], cps_vals, width=width, color=COLORS.get("info", "#3b82f6"), label="CPS")
            ax.set_xticks(list(x))
            ax.set_xticklabels(labels, rotation=20, ha="right", color=fg_color)
            ax.tick_params(axis="y", colors=fg_color)
            for spine in ax.spines.values():
                spine.set_color(fg_color)
            ax.set_ylabel("Tutar (€)", color=fg_color)
            ax.legend(loc="upper right", frameon=False, fontsize=9)
            for bar_group in (bars1, bars2, bars3):
                for bar in bar_group:
                    value = bar.get_height()
                    ax.text(bar.get_x() + bar.get_width() / 2, value, format_currency(value), ha="center", va="bottom", fontsize=8, color=fg_color, rotation=90 if len(labels) > 6 else 0)
            fig.tight_layout()
            canvas = FigureCanvasTkAgg(fig, master=parent)
            canvas.draw()
            canvas.get_tk_widget().pack(fill="both", expand=True)
            self._report_canvases.append(canvas)

        def render_monthly_chart(parent: ttk.Frame, monthly_df: pd.DataFrame) -> None:
            if monthly_df.empty:
                ttk.Label(parent, text="Gelecek teslimatlar bulunamadı", style="Card.TLabel").pack(fill="both", expand=True, pady=8)
                return
            labels = monthly_df["Ay"].tolist()
            totals = normalise_numeric(monthly_df["Amount"].tolist())
            cpi_vals = normalise_numeric(monthly_df["CPI"].tolist())
            cps_vals = normalise_numeric(monthly_df["CPS"].tolist())
            fig = Figure(figsize=(5.4, 2.8), dpi=100)
            ax = fig.add_subplot(111)
            fig.patch.set_facecolor(card_bg)
            ax.set_facecolor(card_bg)
            ax.plot(labels, totals, marker="o", color=accent, label="Toplam")
            ax.plot(labels, cpi_vals, marker="o", color=secondary, label="CPI")
            ax.plot(labels, cps_vals, marker="o", color=COLORS.get("info", "#3b82f6"), label="CPS")
            ax.set_xticks(range(len(labels)))
            ax.set_xticklabels(labels, rotation=25, ha="right", color=fg_color)
            ax.tick_params(axis="y", colors=fg_color)
            for spine in ax.spines.values():
                spine.set_color(fg_color)
            ax.set_ylabel("Tutar (€)", color=fg_color)
            ax.legend(loc="upper left", frameon=False, fontsize=9)
            fig.tight_layout()
            canvas = FigureCanvasTkAgg(fig, master=parent)
            canvas.draw()
            canvas.get_tk_widget().pack(fill="both", expand=True)
            self._report_canvases.append(canvas)

        # Genel satış özeti
        overall_frame = ttk.LabelFrame(container, text="Genel Satış Özeti", style="Card.TLabelframe")
        overall_frame.grid(row=0, column=0, sticky="nsew", pady=6)
        overall_frame.grid_columnconfigure(0, weight=1)
        overall_frame.grid_columnconfigure(1, weight=1)

        summary_values = [
            ("Toplam Satış Tutarı", data["Amount"].sum()),
            ("CPI Satış Tutarı", data["CPI"].sum()),
            ("CPS Satış Tutarı", data["CPS"].sum()),
        ]
        summary_rows = [(label, format_currency(value)) for label, value in summary_values]
        summary_table = ttk.Frame(overall_frame, style="Card.TFrame")
        summary_table.grid(row=0, column=0, sticky="nsew", padx=(0, 12))
        build_table(summary_table, ["Kategori", "Tutar"], summary_rows, height=4)
        summary_chart = ttk.Frame(overall_frame, style="Card.TFrame")
        summary_chart.grid(row=0, column=1, sticky="nsew")
        render_bar_chart(summary_chart, [label for label, _ in summary_values], [value for _, value in summary_values])

        # Satış mühendisleri özeti
        sales_frame = ttk.LabelFrame(container, text="Satış Mühendisleri Dağılımı", style="Card.TLabelframe")
        sales_frame.grid(row=1, column=0, sticky="nsew", pady=6)
        sales_frame.grid_columnconfigure(0, weight=1)
        sales_frame.grid_columnconfigure(1, weight=1)

        grouped_sales = (
            data.groupby("Sales Man")[["Amount", "CPI", "CPS"]]
            .sum()
            .reset_index()
            .sort_values("Amount", ascending=False)
        )
        sales_rows = [
            (
                row["Sales Man"],
                format_currency(row["Amount"]),
                format_currency(row["CPI"]),
                format_currency(row["CPS"]),
            )
            for _, row in grouped_sales.iterrows()
        ]
        sales_table = ttk.Frame(sales_frame, style="Card.TFrame")
        sales_table.grid(row=0, column=0, sticky="nsew", padx=(0, 12))
        build_table(sales_table, ["Satış Mühendisi", "Toplam", "CPI", "CPS"], sales_rows, height=8)
        sales_chart = ttk.Frame(sales_frame, style="Card.TFrame")
        sales_chart.grid(row=0, column=1, sticky="nsew")
        render_grouped_chart(sales_chart, grouped_sales)

        # Faturalandırılan tutarlar
        invoiced_frame = ttk.LabelFrame(container, text="Faturalandırılan Tutarlar", style="Card.TLabelframe")
        invoiced_frame.grid(row=2, column=0, sticky="nsew", pady=6)
        invoiced_frame.grid_columnconfigure(0, weight=1)
        invoiced_frame.grid_columnconfigure(1, weight=1)

        invoiced_df = data[data["Invoiced"] == "YES"]
        invoiced_values = [
            ("Faturalandırılan Toplam", invoiced_df["Invoiced Amount"].sum()),
            ("CPI Tutarı", invoiced_df["CPI"].sum()),
            ("CPS Tutarı", invoiced_df["CPS"].sum()),
        ]
        invoiced_rows = [(label, format_currency(value)) for label, value in invoiced_values]
        invoiced_table = ttk.Frame(invoiced_frame, style="Card.TFrame")
        invoiced_table.grid(row=0, column=0, sticky="nsew", padx=(0, 12))
        build_table(invoiced_table, ["Kategori", "Tutar"], invoiced_rows, height=4)
        invoiced_chart = ttk.Frame(invoiced_frame, style="Card.TFrame")
        invoiced_chart.grid(row=0, column=1, sticky="nsew")
        render_bar_chart(invoiced_chart, [label for label, _ in invoiced_values], [value for _, value in invoiced_values], palette=[accent, secondary, COLORS.get("info", "#3b82f6")])

        invoiced_sales_frame = ttk.LabelFrame(container, text="Satış Mühendisi Bazında Faturalama", style="Card.TLabelframe")
        invoiced_sales_frame.grid(row=3, column=0, sticky="nsew", pady=6)
        invoiced_sales_frame.grid_columnconfigure(0, weight=1)
        invoiced_sales_frame.grid_columnconfigure(1, weight=1)

        invoiced_grouped = (
            invoiced_df.groupby("Sales Man")[["Invoiced Amount", "CPI", "CPS"]]
            .sum()
            .reset_index()
            .sort_values("Invoiced Amount", ascending=False)
        )
        invoiced_sales_rows = [
            (
                row["Sales Man"],
                format_currency(row["Invoiced Amount"]),
                format_currency(row["CPI"]),
                format_currency(row["CPS"]),
            )
            for _, row in invoiced_grouped.iterrows()
        ]
        invoiced_sales_table = ttk.Frame(invoiced_sales_frame, style="Card.TFrame")
        invoiced_sales_table.grid(row=0, column=0, sticky="nsew", padx=(0, 12))
        build_table(invoiced_sales_table, ["Satış Mühendisi", "Faturalı", "CPI", "CPS"], invoiced_sales_rows, height=8)
        invoiced_sales_chart = ttk.Frame(invoiced_sales_frame, style="Card.TFrame")
        invoiced_sales_chart.grid(row=0, column=1, sticky="nsew")
        if not invoiced_grouped.empty:
            chart_df = invoiced_grouped.rename(columns={"Invoiced Amount": "Amount"})[["Sales Man", "Amount", "CPI", "CPS"]]
            render_grouped_chart(invoiced_sales_chart, chart_df)
        else:
            ttk.Label(invoiced_sales_chart, text="Veri bulunamadı", style="Card.TLabel").pack(fill="both", expand=True, pady=8)

        # Teslimat bazlı gelecek faturalama
        forecast_frame = ttk.LabelFrame(container, text="Teslimat Tarihine Göre Faturalama Planı", style="Card.TLabelframe")
        forecast_frame.grid(row=4, column=0, sticky="nsew", pady=6)
        forecast_frame.grid_columnconfigure(0, weight=1)
        forecast_frame.grid_columnconfigure(1, weight=1)

        today = pd.Timestamp.today().normalize()
        future_df = data[data["Date of Delivery"] >= today]
        monthly = (
            future_df.groupby(future_df["Date of Delivery"].dt.to_period("M"))[["Amount", "CPI", "CPS"]]
            .sum()
            .reset_index()
        )
        if not monthly.empty:
            monthly["Ay"] = monthly["Date of Delivery"].dt.to_timestamp().dt.strftime("%Y %B")
        else:
            monthly = monthly.assign(Ay=pd.Series(dtype=str))
        monthly_rows = [
            (
                row.get("Ay", "-"),
                format_currency(row.get("Amount", 0)),
                format_currency(row.get("CPI", 0)),
                format_currency(row.get("CPS", 0)),
            )
            for _, row in monthly.iterrows()
        ]
        forecast_table = ttk.Frame(forecast_frame, style="Card.TFrame")
        forecast_table.grid(row=0, column=0, sticky="nsew", padx=(0, 12))
        build_table(forecast_table, ["Ay", "Toplam", "CPI", "CPS"], monthly_rows, height=6)
        forecast_chart = ttk.Frame(forecast_frame, style="Card.TFrame")
        forecast_chart.grid(row=0, column=1, sticky="nsew")
        render_monthly_chart(forecast_chart, monthly)

        for idx in range(5):
            container.grid_rowconfigure(idx, weight=1)
    
    def generate_report(self) -> None:
        input_file = DATA_FILE
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_file = REPORT_DIR / f"sales_report_{timestamp}.xlsx"

        filtered_df = self.get_filtered_dataframe()
        if filtered_df.empty:
            messagebox.showinfo(
                "Bilgi",
                "Raporlanacak veri bulunamadı. Excel çıktısı yine de oluşturulacaktır.",
            )
        else:
            self._show_reporting_dashboard(filtered_df)

        progress_window = tk.Toplevel(self.root)
        progress_window.title("Rapor Oluşturuluyor")
        ttk.Label(progress_window, text="Rapor hazırlanıyor, lütfen bekleyin...").pack(padx=16, pady=12)
        progress = ttk.Progressbar(progress_window, orient="horizontal", length=280, mode="determinate")
        progress.pack(padx=16, pady=12)

        def update_progress(value: float, message: str) -> None:
            progress["value"] = value * 100
            progress_window.update_idletasks()
            self._update_status(message)

        def worker() -> None:
            try:
                sales_reporting.generate_sales_report(
                    input_file,
                    str(output_file),
                    progress_callback=update_progress,
                )
                messagebox.showinfo("Başarılı", f"Rapor oluşturuldu: {output_file}")
            except Exception as exc:
                messagebox.showerror("Rapor Hatası", str(exc))
            finally:
                progress_window.destroy()

        threading.Thread(target=worker, daemon=True).start()

    # -------------------------------------------------------------- utilities
    def schedule_refresh(self, delay_ms: int = 200) -> None:
        self.root.after(delay_ms, self.apply_filters)

    def show_help(self) -> None:
        messagebox.showinfo(
            "Kullanım",
            "Formu doldurun, Kaydet ile ekleyin. Kayıt seçerek Güncelle veya Sil yapın."
            " Ara / Filtrele menüsü ile müşteri, tarih ve durum filtreleri uygulayın.",
        )

    def show_about(self) -> None:
        messagebox.showinfo("Hakkında", "Satış Veri Giriş Sistemi\nSürüm 1.0")

    def run(self) -> None:
        self.root.mainloop()


if __name__ == "__main__":
    app = SalesEntryApp()
    app.run()
