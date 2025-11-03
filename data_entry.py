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
from datetime import datetime
from functools import partial
from pathlib import Path
from tkinter import filedialog, messagebox, simpledialog, ttk
from typing import Dict, List, Optional, Tuple

import pandas as pd
from tkcalendar import DateEntry

import sales_reporting


APP_TITLE = "Satış Veri Giriş Sistemi"
APP_GEOMETRY = "1200x820"
DATA_FILE = "sales_data_master.xlsx"
CONFIG_FILE = "config.json"
BACKUP_DIR = Path("backups")
REPORT_DIR = Path("reports")
AUTO_SAVE_INTERVAL = 5 * 60 * 1000  # 5 minutes in milliseconds
PAGE_SIZE = 15

COLUMNS = [
    "Date of Request",
    "Date of Issue",
    "Date of Delivery",
    "Sales Man",
    "Customer Name",
    "Customer DO No",
    "Definition",
    "C4C Code",
    "Sales Ticket Reference",
    "SO No",
    "PO No",
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
SALES_REP_PASSWORD = "Remzi123"
DEFAULT_THEME = "light"
THEMES = {"light": "Açık", "dark": "Koyu"}

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


@dataclass
class FilterOptions:
    search_text: str = ""
    salesman: str = ""
    invoiced: str = ""
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
        self.filter_options = FilterOptions()
        self._updating_cpi_field = False

        self._config = self._load_config()
        self.sales_reps = self._load_sales_reps()
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
            config = {"theme": DEFAULT_THEME, "sales_reps": DEFAULT_SALES_REPS}
            self._save_config(config)
            return config
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as fh:
                config = json.load(fh)
        except (json.JSONDecodeError, OSError):
            config = {"theme": DEFAULT_THEME, "sales_reps": DEFAULT_SALES_REPS}
        if "theme" not in config:
            config["theme"] = DEFAULT_THEME
        sales_reps = config.get("sales_reps")
        if not isinstance(sales_reps, list) or not all(isinstance(item, str) for item in sales_reps):
            config["sales_reps"] = DEFAULT_SALES_REPS.copy()
        return config

    def _save_config(self, config: Dict[str, object]) -> None:
        with open(CONFIG_FILE, "w", encoding="utf-8") as fh:
            json.dump(config, fh, indent=2, ensure_ascii=False)

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

    def _apply_theme(self, theme: str) -> None:
        if theme == "dark":
            bg = COLORS["bg_dark"]
            fg = "white"
        else:
            bg = COLORS["bg_light"]
            fg = COLORS["text_dark"]
        self.root.configure(bg=bg)
        style = ttk.Style(self.root)
        style.theme_use("clam")
        style.configure("TLabel", background=bg, foreground=fg, font=("Segoe UI", 10))
        style.configure("Header.TLabel", font=("Segoe UI", 14, "bold"), foreground=COLORS["primary"])
        style.configure("TFrame", background=bg)
        style.configure("TButton", font=("Segoe UI", 10), padding=6)
        style.map(
            "TButton",
            background=[("active", COLORS["primary"])],
            foreground=[("active", "white")],
        )
        style.configure(
            "Accent.TButton",
            background=COLORS["primary"],
            foreground="white",
            borderwidth=0,
        )
        style.map("Accent.TButton", background=[("active", COLORS["secondary"])])
        style.configure("Danger.TButton", background=COLORS["danger"], foreground="white")
        style.map("Danger.TButton", background=[("active", "#b91c1c")])

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
            background="white",
            foreground=COLORS["text_dark"],
            fieldbackground="white",
            rowheight=26,
            font=("Segoe UI", 10),
        )
        style.map("Custom.Treeview", background=[("selected", COLORS["secondary"])], foreground=[("selected", "white")])
        style.configure("Custom.Treeview.Heading", font=("Segoe UI", 10, "bold"))

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
        edit_menu.add_command(label="Yeni Kayıt", accelerator="Ctrl+N", command=self.reset_form)
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
        self.root.bind("<Control-n>", lambda event: self.reset_form())
        self.root.bind("<Control-N>", lambda event: self.reset_form())
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
        ttk.Label(self.header_frame, textvariable=self.file_info_var).pack(side="right")

        self.content_frame = ttk.Frame(self.root)
        self.content_frame.pack(fill="both", expand=True, padx=16, pady=8)

        self.form_frame = ttk.LabelFrame(self.content_frame, text="Veri Giriş Formu")
        self.form_frame.pack(side="left", fill="y", padx=(0, 8), pady=4)

        self.table_frame = ttk.LabelFrame(self.content_frame, text="Kayıtlı Veriler")
        self.table_frame.pack(side="right", fill="both", expand=True, pady=4)

    def _create_form(self) -> None:
        self.form_vars: Dict[str, tk.Variable] = {}

        def create_labeled_row(parent, label_text, widget_factory):
            """Create a labelled row and ensure the widget is packed correctly."""

            row = ttk.Frame(parent)
            row.pack(fill="x", pady=3)
            ttk.Label(row, text=label_text, width=22).pack(side="left")

            widget = widget_factory(row)
            if isinstance(widget, tk.Widget):
                widget.pack(side="left", fill="x", expand=True)
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

            create_labeled_row(self.form_frame, date_labels[field], widget_factory)

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
            ("Customer DO No", "Müşteri Sipariş No"),
            ("Definition", "Ürün Tanımı"),
            ("C4C Code", "C4C Kodu"),
            ("Sales Ticket Reference", "Satış Ref"),
            ("SO No", "SO No"),
            ("PO No", "PO No"),
            ("Delivery Note", "Teslimat Notu"),
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
        self.cpi_total_entry = create_labeled_row(
            self.form_frame,
            "CPI Tutarı (€)",
            lambda parent: ttk.Entry(
                parent,
                textvariable=self.form_vars["Invoiced Amount"],
                state="readonly",
            ),
        )

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
                ttk.Radiobutton(
                    container,
                    text=choice,
                    value=choice,
                    variable=self.form_vars[var_name],
                ).pack(side="left", padx=2)

        create_radio("QI Forecast", "QI Forecast")
        create_radio("Faturalandi", "Invoiced")

        # Aksiyon butonları
        btn_frame = ttk.Frame(self.form_frame)
        btn_frame.pack(fill="x", pady=(12, 0))
        ttk.Button(btn_frame, text="Temizle", command=self.reset_form).pack(side="left", expand=True, fill="x", padx=2)
        ttk.Button(btn_frame, text="Kaydet", style="Accent.TButton", command=self.save_data).pack(
            side="left", expand=True, fill="x", padx=2
        )
        ttk.Button(btn_frame, text="Güncelle", command=self.update_data).pack(side="left", expand=True, fill="x", padx=2)

    def _create_table(self) -> None:
        columns = ["#", *COLUMNS]
        self.tree = ttk.Treeview(
            self.table_frame,
            columns=columns,
            show="headings",
            style="Custom.Treeview",
        )
        self.tree.pack(side="left", fill="both", expand=True)

        self.tree.tag_configure("invoiced", background="#d1fae5")

        display_names = {"Invoiced Amount": "CPI Tutarı"}
        for col in columns:
            heading_text = display_names.get(col, col)
            self.tree.heading(col, text=heading_text, command=partial(self.sort_by_column, col))
            self.tree.column(col, width=120, anchor="center")
        self.tree.column("#", width=60, anchor="center")
        self.tree.column("Definition", width=200, anchor="w")
        self.tree.column("Delivery Note", width=220, anchor="w")

        scrollbar_y = ttk.Scrollbar(self.table_frame, orient="vertical", command=self.tree.yview)
        scrollbar_y.pack(side="right", fill="y")
        self.tree.configure(yscrollcommand=scrollbar_y.set)

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
        ttk.Button(top_buttons, text="Ara / Filtrele", command=self.open_filter_window).pack(side="left", padx=4)
        ttk.Button(top_buttons, text="Satış Elemanları", command=self.open_sales_rep_manager).pack(side="left", padx=4)
        ttk.Button(top_buttons, text="Dışa Aktar", command=self.export_filtered_data).pack(side="left", padx=4)

    def _create_bottom_buttons(self) -> None:
        bottom_frame = ttk.Frame(self.root)
        bottom_frame.pack(fill="x", padx=16, pady=8)
        ttk.Button(bottom_frame, text="📊 Rapor Oluştur", style="Accent.TButton", command=self.generate_report).pack(
            side="left"
        )
        ttk.Button(bottom_frame, text="Yedeklemeyi Aç", command=self.open_backup_directory).pack(side="right")

    # ----------------------------------------------------------------- helpers
    def _update_status(self, text: str) -> None:
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.status_var.set(f"[{timestamp}] {text}")

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
        if password != SALES_REP_PASSWORD:
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
        self._normalise_discount_values()
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
            if discount_raw is None:
                continue
            amount = self._to_float(row.get("Amount"))
            if discount_raw > 1:
                if amount:
                    fraction = discount_raw / amount
                else:
                    fraction = discount_raw / 100
            else:
                fraction = discount_raw
            fraction = max(0.0, min(fraction, 1.0))
            if discount_raw != fraction:
                self.df.at[idx, "Total Discount"] = fraction
                if amount is not None:
                    self.df.at[idx, "CPI"] = amount * (1 - fraction)
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
    def reset_form(self) -> None:
        for key, var in self.form_vars.items():
            if isinstance(var, tk.StringVar):
                if key in ("QI Forecast", "Invoiced"):
                    var.set("NO")
                elif key == "Sales Man":
                    options = self._get_sales_rep_options()
                    var.set(options[0] if options else "")
                elif "Date" in key:
                    var.set(datetime.today().strftime("%d.%m.%Y"))
                else:
                    var.set("")
        self.selected_index = None
        self.tree.selection_remove(self.tree.selection())
        self._update_cpi_field()
        self._update_status("Form temizlendi")

    def populate_form_from_selection(self) -> None:
        selected = self.tree.selection()
        if not selected:
            return
        item = self.tree.item(selected[0])
        values = item["values"][1:]
        if len(values) != len(COLUMNS):
            return
        row_data = dict(zip(COLUMNS, values))
        for col, value in row_data.items():
            if col in ("QI Forecast", "Invoiced"):
                self.form_vars[col].set(str(value).upper())
            elif col in ("Date of Request", "Date of Issue", "Date of Delivery"):
                try:
                    dt = pd.to_datetime(value)
                    self.form_vars[col].set(dt.strftime("%d.%m.%Y"))
                except Exception:
                    self.form_vars[col].set(value)
            elif col in self.form_vars:
                self.form_vars[col].set(str(value) if value is not None else "")

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

    def _format_value(self, value) -> str:
        if pd.isna(value):
            return ""
        if isinstance(value, (float, int)):
            return f"{value:,.2f}" if value > 999 else f"{value:.2f}"
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
            amount = self._parse_float(self.form_vars["Amount"].get())
            cps_value = self._parse_float(self.form_vars["CPS"].get())
            if amount is None:
                self.form_vars["Invoiced Amount"].set("")
                return
            cps = cps_value or 0.0
            cpi_total = amount - cps
            formatted = f"{cpi_total:.2f}".replace(".", ",")
            self.form_vars["Invoiced Amount"].set(formatted)
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

        amount = self._parse_float(self.form_vars["Amount"].get())
        if amount is None:
            errors.append("Geçerli bir tutar girin")
        discount_percent_value = self._parse_float(self.form_vars["DiscountPercent"].get()) or 0.0
        self._format_discount_entry()
        if discount_percent_value < 0:
            errors.append("İndirim yüzdesi negatif olamaz")
        discount_fraction = discount_percent_value / 100
        if discount_fraction > 1:
            errors.append("İndirim 100%'ü aşamaz")
        cps_value = self._parse_float(self.form_vars["CPS"].get()) or 0.0

        if errors:
            return None, "\n".join(errors)

        discount_fraction = max(0.0, min(discount_fraction, 1.0))
        discount_amount = amount * discount_fraction
        cpi_value = amount - discount_amount
        cpi_total = amount - cps_value

        data["Amount"] = amount
        data["Total Discount"] = discount_fraction
        data["CPI"] = cpi_value
        data["CPS"] = cps_value
        data["Invoiced Amount"] = cpi_total
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
        record, error = self._collect_form_data()
        if error:
            messagebox.showerror("Geçersiz Veri", error)
            return
        assert record is not None

        self.df = pd.concat([self.df, record.to_frame().T], ignore_index=True)
        self.save_current_dataframe()
        self.apply_filters()
        self.reset_form()
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
        self.df.loc[self.selected_index, record.index] = record.values
        self.save_current_dataframe()
        self.apply_filters()
        self.reset_form()
        messagebox.showinfo("Güncellendi", "Kayıt başarıyla güncellendi")

    def delete_data(self) -> None:
        selected = self.tree.selection()
        if not selected:
            messagebox.showwarning("Seçim Yok", "Silmek için kayıt seçiniz")
            return
        if not messagebox.askyesno("Onay", "Seçili kayıt silinsin mi?"):
            return
        item = self.tree.item(selected[0])
        index = int(item["values"][0]) - 1
        self.df = self.df.drop(self.df.index[index]).reset_index(drop=True)
        self.save_current_dataframe()
        self.apply_filters()
        self.reset_form()
        self._update_status("Kayıt silindi")

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
        popup.geometry("320x320")

        opts = self.filter_options

        ttk.Label(popup, text="Müşteri Adı").pack(pady=4)
        search_var = tk.StringVar(value=opts.search_text)
        ttk.Entry(popup, textvariable=search_var).pack(fill="x", padx=16)

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
            self.filter_options.salesman = salesman_var.get().strip()
            self.filter_options.invoiced = invoiced_var.get().strip()
            self.filter_options.start_date = self._parse_date_str(start_var.get())
            self.filter_options.end_date = self._parse_date_str(end_var.get())
            self.apply_filters()
            popup.destroy()

        ttk.Button(popup, text="Uygula", command=apply).pack(pady=6)

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
    def generate_report(self) -> None:
        input_file = DATA_FILE
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_file = REPORT_DIR / f"sales_report_{timestamp}.xlsx"

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
