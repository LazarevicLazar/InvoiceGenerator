from __future__ import annotations

import os
import tkinter as tk
import subprocess
import sys
from datetime import date, datetime, timedelta
from pathlib import Path
from tkinter import filedialog, messagebox, simpledialog, ttk

from app.config import APP_NAME, DATE_FORMAT, DEFAULT_DUE_DAYS
from app.database import DatabaseManager
from app.services.calendar_service import CalendarService
from app.services.customer_service import CustomerService
from app.services.email_service import EmailService
from app.services.google_calendar_service import GoogleCalendarService
from app.services.invoice_service import InvoiceService
from app.services.pdf_service import PDFService, format_usd
from app.services.sms_service import SMSService
from app.services.settings_service import SettingsService

CALENDAR_DATETIME_FORMAT = "%Y-%m-%d %H:%M"


class InvoiceGeneratorApp(tk.Tk):
    def __init__(self, db: DatabaseManager) -> None:
        super().__init__()
        self.db = db

        self.customer_service = CustomerService(db)
        self.calendar_service = CalendarService(db)
        self.settings_service = SettingsService(db)
        self.invoice_service = InvoiceService(db)
        self.pdf_service = PDFService()
        self.email_service = EmailService()
        self.sms_service = SMSService()
        self.google_calendar_service: GoogleCalendarService | None = None
        self.google_connected = False
        self.google_calendar_lookup: dict[str, str] = {}

        self.customer_lookup: dict[str, int] = {}
        self.calendar_customer_lookup: dict[str, int] = {}
        self.cleaner_lookup: dict[str, int] = {}
        self.current_items: list[dict] = []
        self.calendar_jobs_cache: list[dict] = []
        self.calendar_anchor_date = date.today()
        self.customer_frequency_var = tk.StringVar(value="Single")

        self.cleaners_popup: tk.Toplevel | None = None
        self.google_popup: tk.Toplevel | None = None
        self.schedule_popup: tk.Toplevel | None = None

        self.title(APP_NAME)
        self.geometry("1280x780")
        self.minsize(1100, 680)

        self._build_ui()
        self._load_settings_into_form()
        self._refresh_google_status()
        self._load_customers()
        self._load_cleaners()
        self._load_jobs()
        self._load_invoices()
        self._reset_invoice_form()
        self._reset_job_form()

    def _build_ui(self) -> None:
        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill="both", expand=True, padx=10, pady=10)

        self.customers_tab = ttk.Frame(self.notebook)
        self.invoice_tab = ttk.Frame(self.notebook)
        self.calendar_tab = ttk.Frame(self.notebook)
        self.invoices_tab = ttk.Frame(self.notebook)
        self.settings_tab = ttk.Frame(self.notebook)

        self.notebook.add(self.customers_tab, text="Customers")
        self.notebook.add(self.invoice_tab, text="Create Invoice")
        self.notebook.add(self.calendar_tab, text="Calendar")
        self.notebook.add(self.invoices_tab, text="Invoice History")
        self.notebook.add(self.settings_tab, text="Settings")

        self._build_customers_tab()
        self._build_invoice_tab()
        self._build_calendar_tab()
        self._build_invoices_tab()
        self._build_settings_tab()

    def _build_customers_tab(self) -> None:
        main = ttk.Frame(self.customers_tab)
        main.pack(fill="both", expand=True, padx=8, pady=8)
        main.columnconfigure(0, weight=2)
        main.columnconfigure(1, weight=3)
        main.rowconfigure(0, weight=1)

        form = ttk.LabelFrame(main, text="Customer Details")
        form.grid(row=0, column=0, sticky="nsew", padx=(0, 8), pady=0)
        form.columnconfigure(1, weight=1)

        self.customer_id_var = tk.StringVar(value="")
        self.customer_name_var = tk.StringVar()
        self.customer_email_var = tk.StringVar()
        self.customer_phone_var = tk.StringVar()
        self.customer_address_var = tk.StringVar()
        self.customer_bedrooms_var = tk.StringVar()
        self.customer_bathrooms_var = tk.StringVar()
        self.customer_square_feet_var = tk.StringVar()
        self.customer_cleaning_duration_var = tk.StringVar()

        ttk.Label(form, text="Name").grid(row=0, column=0, sticky="w", padx=6, pady=(8, 4))
        ttk.Entry(form, textvariable=self.customer_name_var).grid(row=0, column=1, sticky="ew", padx=6, pady=(8, 4))

        ttk.Label(form, text="Email").grid(row=1, column=0, sticky="w", padx=6, pady=4)
        ttk.Entry(form, textvariable=self.customer_email_var).grid(row=1, column=1, sticky="ew", padx=6, pady=4)

        ttk.Label(form, text="Phone").grid(row=2, column=0, sticky="w", padx=6, pady=4)
        ttk.Entry(form, textvariable=self.customer_phone_var).grid(row=2, column=1, sticky="ew", padx=6, pady=4)

        job_frame = ttk.LabelFrame(form, text="Job Details")
        job_frame.grid(row=4, column=0, columnspan=2, sticky="ew", padx=6, pady=6)
        job_frame.columnconfigure((1, 3), weight=1)

        ttk.Label(job_frame, text="Beds").grid(row=0, column=0, sticky="w", padx=6, pady=4)
        ttk.Entry(job_frame, textvariable=self.customer_bedrooms_var).grid(row=0, column=1, sticky="ew", padx=6, pady=4)

        ttk.Label(job_frame, text="Baths").grid(row=0, column=2, sticky="w", padx=6, pady=4)
        ttk.Entry(job_frame, textvariable=self.customer_bathrooms_var).grid(row=0, column=3, sticky="ew", padx=6, pady=4)

        ttk.Label(job_frame, text="SqFt").grid(row=1, column=0, sticky="w", padx=6, pady=4)
        ttk.Entry(job_frame, textvariable=self.customer_square_feet_var).grid(row=1, column=1, sticky="ew", padx=6, pady=4)

        ttk.Label(job_frame, text="Cleaning Duration").grid(row=1, column=2, sticky="w", padx=6, pady=4)
        ttk.Entry(job_frame, textvariable=self.customer_cleaning_duration_var).grid(row=1, column=3, sticky="ew", padx=6, pady=4)

        ttk.Label(job_frame, text="Frequency").grid(row=2, column=2, sticky="w", padx=6, pady=4)
        frequency_combo = ttk.Combobox(
            job_frame,
            textvariable=self.customer_frequency_var,
            values=("Single", "Weekly", "Bi-Weekly", "Monthly", "Bi-Monthly"),
            state="readonly",
            width=14,
        )
        frequency_combo.grid(row=2, column=3, sticky="ew", padx=6, pady=4)

        ttk.Label(form, text="Address").grid(row=3, column=0, sticky="w", padx=6, pady=4)
        ttk.Entry(form, textvariable=self.customer_address_var).grid(row=3, column=1, sticky="ew", padx=6, pady=4)

        ttk.Label(form, text="Notes").grid(row=5, column=0, sticky="nw", padx=6, pady=4)
        self.customer_notes_text = tk.Text(form, height=5, width=40)
        self.customer_notes_text.grid(row=5, column=1, sticky="ew", padx=6, pady=4)

        button_row = ttk.Frame(form)
        button_row.grid(row=6, column=0, columnspan=2, sticky="ew", padx=6, pady=(10, 8))
        button_row.columnconfigure((0, 1, 2, 3, 4), weight=1)

        ttk.Button(button_row, text="Add", command=self._add_customer).grid(row=0, column=0, sticky="ew", padx=3)
        ttk.Button(button_row, text="Update", command=self._update_customer).grid(row=0, column=1, sticky="ew", padx=3)
        ttk.Button(button_row, text="Delete", command=self._delete_customer).grid(row=0, column=2, sticky="ew", padx=3)
        ttk.Button(button_row, text="Delete + Invoices", command=self._delete_customer_with_invoices).grid(row=0, column=3, sticky="ew", padx=3)
        ttk.Button(button_row, text="Clear", command=self._clear_customer_form).grid(row=0, column=4, sticky="ew", padx=3)

        table = ttk.LabelFrame(main, text="Customers")
        table.grid(row=0, column=1, sticky="nsew")
        table.columnconfigure(0, weight=1)
        table.rowconfigure(0, weight=1)

        columns = ("id", "name", "email", "phone", "address", "frequency", "last_clean")
        self.customers_tree = ttk.Treeview(table, columns=columns, show="headings", height=18)
        self.customers_tree.heading("id", text="ID")
        self.customers_tree.heading("name", text="Name")
        self.customers_tree.heading("email", text="Email")
        self.customers_tree.heading("phone", text="Phone")
        self.customers_tree.heading("address", text="Address")
        self.customers_tree.heading("frequency", text="Frequency")
        self.customers_tree.heading("last_clean", text="Last Clean")

        self.customers_tree.column("id", width=45, anchor="center")
        self.customers_tree.column("name", width=160)
        self.customers_tree.column("email", width=200)
        self.customers_tree.column("phone", width=110)
        self.customers_tree.column("address", width=220)
        self.customers_tree.column("frequency", width=90, anchor="center")
        self.customers_tree.column("last_clean", width=100, anchor="center")

        self.customers_tree.tag_configure("Single", background="#ffffff")
        self.customers_tree.tag_configure("Weekly", background="#c8e6c9")
        self.customers_tree.tag_configure("Bi-Weekly", background="#bbdefb")
        self.customers_tree.tag_configure("Monthly", background="#fff9c4")
        self.customers_tree.tag_configure("Bi-Monthly", background="#e1bee7")

        self.customers_tree.grid(row=0, column=0, sticky="nsew")
        self.customers_tree.bind("<<TreeviewSelect>>", self._on_customer_selected)

        scrollbar = ttk.Scrollbar(table, orient="vertical", command=self.customers_tree.yview)
        self.customers_tree.configure(yscrollcommand=scrollbar.set)
        scrollbar.grid(row=0, column=1, sticky="ns")

    def _build_invoice_tab(self) -> None:
        container = ttk.Frame(self.invoice_tab)
        container.pack(fill="both", expand=True, padx=8, pady=8)
        container.columnconfigure(0, weight=1)
        container.rowconfigure(1, weight=1)

        top = ttk.LabelFrame(container, text="Invoice Header")
        top.grid(row=0, column=0, sticky="ew")
        for i in range(6):
            top.columnconfigure(i, weight=1)

        self.invoice_customer_var = tk.StringVar()
        self.invoice_issue_date_var = tk.StringVar()
        self.invoice_due_date_var = tk.StringVar()
        self.invoice_tax_rate_var = tk.StringVar(value="0")

        ttk.Label(top, text="Customer").grid(row=0, column=0, sticky="w", padx=6, pady=(8, 4))
        self.invoice_customer_combo = ttk.Combobox(
            top,
            textvariable=self.invoice_customer_var,
            state="readonly",
            width=42,
        )
        self.invoice_customer_combo.grid(row=0, column=1, columnspan=2, sticky="ew", padx=6, pady=(8, 4))

        ttk.Label(top, text="Issue Date").grid(row=0, column=3, sticky="w", padx=6, pady=(8, 4))
        ttk.Entry(top, textvariable=self.invoice_issue_date_var).grid(row=0, column=4, sticky="ew", padx=6, pady=(8, 4))

        ttk.Label(top, text="Due Date").grid(row=0, column=5, sticky="w", padx=6, pady=(8, 4))
        ttk.Entry(top, textvariable=self.invoice_due_date_var).grid(row=0, column=5, sticky="e", padx=(80, 6), pady=(8, 4))

        ttk.Label(top, text="Tax %").grid(row=1, column=0, sticky="w", padx=6, pady=(4, 8))
        tax_entry = ttk.Entry(top, textvariable=self.invoice_tax_rate_var, width=10)
        tax_entry.grid(row=1, column=1, sticky="w", padx=6, pady=(4, 8))
        tax_entry.bind("<FocusOut>", lambda _event: self._refresh_invoice_item_table())

        ttk.Label(top, text="Notes").grid(row=1, column=2, sticky="w", padx=6, pady=(4, 8))
        self.invoice_notes_text = tk.Text(top, height=3, width=60)
        self.invoice_notes_text.grid(row=1, column=3, columnspan=3, sticky="ew", padx=6, pady=(4, 8))

        middle = ttk.Frame(container)
        middle.grid(row=1, column=0, sticky="nsew", pady=(10, 0))
        middle.columnconfigure(0, weight=1)
        middle.rowconfigure(1, weight=1)

        item_form = ttk.LabelFrame(middle, text="Add Service Item")
        item_form.grid(row=0, column=0, sticky="ew")
        for i in range(5):
            item_form.columnconfigure(i, weight=1)

        self.item_description_var = tk.StringVar()
        self.item_rate_var = tk.StringVar(value="0")

        ttk.Label(item_form, text="Description").grid(row=0, column=0, sticky="w", padx=6, pady=6)
        ttk.Entry(item_form, textvariable=self.item_description_var).grid(row=0, column=1, sticky="ew", padx=6, pady=6)

        ttk.Label(item_form, text="Rate (USD)").grid(row=0, column=2, sticky="w", padx=6, pady=6)
        ttk.Entry(item_form, textvariable=self.item_rate_var, width=12).grid(row=0, column=3, sticky="w", padx=6, pady=6)

        ttk.Button(item_form, text="Add Item", command=self._add_invoice_item).grid(row=0, column=4, sticky="ew", padx=6, pady=6)

        items_frame = ttk.LabelFrame(middle, text="Invoice Items")
        items_frame.grid(row=1, column=0, sticky="nsew", pady=(10, 0))
        items_frame.columnconfigure(0, weight=1)
        items_frame.rowconfigure(0, weight=1)

        self.items_tree = ttk.Treeview(
            items_frame,
            columns=("description", "unit_price", "line_total"),
            show="headings",
            height=12,
        )
        self.items_tree.heading("description", text="Description")
        self.items_tree.heading("unit_price", text="Unit Price")
        self.items_tree.heading("line_total", text="Line Total")

        self.items_tree.column("description", width=620)
        self.items_tree.column("unit_price", width=160, anchor="e")
        self.items_tree.column("line_total", width=160, anchor="e")
        self.items_tree.grid(row=0, column=0, sticky="nsew")

        scroll = ttk.Scrollbar(items_frame, orient="vertical", command=self.items_tree.yview)
        self.items_tree.configure(yscrollcommand=scroll.set)
        scroll.grid(row=0, column=1, sticky="ns")

        action_row = ttk.Frame(items_frame)
        action_row.grid(row=1, column=0, columnspan=2, sticky="ew", padx=4, pady=8)
        action_row.columnconfigure((0, 1, 2, 3, 4, 5), weight=1)

        ttk.Button(action_row, text="Remove Selected Item", command=self._remove_selected_item).grid(row=0, column=0, sticky="ew", padx=4)
        ttk.Button(action_row, text="Clear Items", command=self._clear_items).grid(row=0, column=1, sticky="ew", padx=4)

        self.subtotal_label = ttk.Label(action_row, text="Subtotal: $0.00")
        self.subtotal_label.grid(row=0, column=2, sticky="e", padx=4)
        self.tax_label = ttk.Label(action_row, text="Tax: $0.00")
        self.tax_label.grid(row=0, column=3, sticky="e", padx=4)
        self.total_label = ttk.Label(action_row, text="Total: $0.00")
        self.total_label.grid(row=0, column=4, sticky="e", padx=4)

        ttk.Button(action_row, text="Reset Form", command=self._reset_invoice_form).grid(row=0, column=5, sticky="ew", padx=4)

        bottom = ttk.Frame(container)
        bottom.grid(row=2, column=0, sticky="ew", pady=(12, 0))
        bottom.columnconfigure((0, 1, 2, 3), weight=1)

        ttk.Button(bottom, text="Save Draft", command=lambda: self._save_invoice("draft")).grid(row=0, column=0, sticky="ew", padx=4)
        ttk.Button(bottom, text="Save + Generate PDF", command=lambda: self._save_invoice("pdf")).grid(row=0, column=1, sticky="ew", padx=4)
        ttk.Button(bottom, text="Save + PDF + Email", command=lambda: self._save_invoice("email")).grid(row=0, column=2, sticky="ew", padx=4)
        ttk.Button(bottom, text="Save + Text", command=lambda: self._save_invoice("sms")).grid(row=0, column=3, sticky="ew", padx=4)

    def _build_calendar_tab(self) -> None:
        self.cleaner_id_var = tk.StringVar(value="")
        self.cleaner_name_var = tk.StringVar()
        self.cleaner_phone_var = tk.StringVar()
        self.cleaner_email_var = tk.StringVar()
        self.cleaner_google_calendar_id_var = tk.StringVar()
        self.auto_create_cleaner_calendar_var = tk.IntVar(value=1)
        self.google_status_var = tk.StringVar(value="Status: Not connected")
        self.google_calendars_combo_var = tk.StringVar()
        self.google_target_cleaner_var = tk.StringVar()

        self.job_customer_var = tk.StringVar()
        self.job_cleaner_var = tk.StringVar()
        self.job_title_var = tk.StringVar()
        self.job_start_var = tk.StringVar()
        self.job_end_var = tk.StringVar()
        self.job_location_var = tk.StringVar()

        self.calendar_view_var = tk.StringVar(value="Week")
        self.calendar_range_var = tk.StringVar(value="")

        container = ttk.Frame(self.calendar_tab)
        container.pack(fill="both", expand=True, padx=8, pady=8)
        container.columnconfigure(0, weight=1)
        container.rowconfigure(2, weight=1)

        top_actions = ttk.Frame(container)
        top_actions.grid(row=0, column=0, sticky="ew", pady=(0, 8))
        top_actions.columnconfigure((0, 1, 2, 3), weight=1)
        ttk.Button(top_actions, text="Manage Cleaners", command=self._open_cleaners_popup).grid(row=0, column=0, sticky="ew", padx=4)
        ttk.Button(top_actions, text="Google Calendar", command=self._open_google_popup).grid(row=0, column=1, sticky="ew", padx=4)
        ttk.Button(top_actions, text="Schedule Job", command=self._open_schedule_popup).grid(row=0, column=2, sticky="ew", padx=4)
        ttk.Button(top_actions, text="Refresh", command=self._load_jobs).grid(row=0, column=3, sticky="ew", padx=4)

        nav = ttk.Frame(container)
        nav.grid(row=1, column=0, sticky="ew", pady=(0, 8))
        nav.columnconfigure(4, weight=1)
        ttk.Button(nav, text="<", width=4, command=lambda: self._move_calendar_range(-1)).grid(row=0, column=0, padx=(0, 4))
        ttk.Button(nav, text="Today", command=self._set_calendar_today).grid(row=0, column=1, padx=4)
        ttk.Button(nav, text=">", width=4, command=lambda: self._move_calendar_range(1)).grid(row=0, column=2, padx=4)

        view_combo = ttk.Combobox(
            nav,
            textvariable=self.calendar_view_var,
            values=("Day", "Week", "Month", "Year"),
            state="readonly",
            width=10,
        )
        view_combo.grid(row=0, column=3, padx=(10, 8))
        view_combo.bind("<<ComboboxSelected>>", self._on_calendar_view_changed)
        ttk.Label(nav, textvariable=self.calendar_range_var).grid(row=0, column=4, sticky="w")

        jobs_box = ttk.LabelFrame(container, text="Scheduled Jobs")
        jobs_box.grid(row=2, column=0, sticky="nsew")
        jobs_box.columnconfigure(0, weight=1)
        jobs_box.rowconfigure(0, weight=1)

        self.jobs_tree = ttk.Treeview(jobs_box, show="headings", height=18)
        self.jobs_tree.grid(row=0, column=0, sticky="nsew")

        jobs_scroll = ttk.Scrollbar(jobs_box, orient="vertical", command=self.jobs_tree.yview)
        self.jobs_tree.configure(yscrollcommand=jobs_scroll.set)
        jobs_scroll.grid(row=0, column=1, sticky="ns")

        jobs_actions = ttk.Frame(jobs_box)
        jobs_actions.grid(row=1, column=0, columnspan=2, sticky="ew", padx=6, pady=(8, 8))
        jobs_actions.columnconfigure((0, 1, 2, 3), weight=1)
        ttk.Button(jobs_actions, text="Mark In Progress", command=lambda: self._update_selected_job_status("in-progress")).grid(row=0, column=0, sticky="ew", padx=3)
        ttk.Button(jobs_actions, text="Mark Done", command=lambda: self._update_selected_job_status("completed")).grid(row=0, column=1, sticky="ew", padx=3)
        ttk.Button(jobs_actions, text="Cancel Job", command=lambda: self._update_selected_job_status("cancelled")).grid(row=0, column=2, sticky="ew", padx=3)
        ttk.Button(jobs_actions, text="Delete Job", command=self._delete_selected_job).grid(row=0, column=3, sticky="ew", padx=3)

        self._configure_calendar_columns()
        self._update_calendar_range_label()

    def _build_invoices_tab(self) -> None:
        container = ttk.Frame(self.invoices_tab)
        container.pack(fill="both", expand=True, padx=8, pady=8)
        container.columnconfigure(0, weight=1)
        container.rowconfigure(0, weight=1)

        self.invoices_tree = ttk.Treeview(
            container,
            columns=("id", "invoice_number", "customer", "issue", "due", "total", "status", "sent_at"),
            show="headings",
            height=22,
        )
        self.invoices_tree.heading("id", text="ID")
        self.invoices_tree.heading("invoice_number", text="Invoice #")
        self.invoices_tree.heading("customer", text="Customer")
        self.invoices_tree.heading("issue", text="Issue Date")
        self.invoices_tree.heading("due", text="Due Date")
        self.invoices_tree.heading("total", text="Total")
        self.invoices_tree.heading("status", text="Status")
        self.invoices_tree.heading("sent_at", text="Sent At")

        self.invoices_tree.column("id", width=60, anchor="center")
        self.invoices_tree.column("invoice_number", width=140)
        self.invoices_tree.column("customer", width=220)
        self.invoices_tree.column("issue", width=110)
        self.invoices_tree.column("due", width=110)
        self.invoices_tree.column("total", width=120, anchor="e")
        self.invoices_tree.column("status", width=100, anchor="center")
        self.invoices_tree.column("sent_at", width=170)

        self.invoices_tree.grid(row=0, column=0, sticky="nsew")
        scroll = ttk.Scrollbar(container, orient="vertical", command=self.invoices_tree.yview)
        self.invoices_tree.configure(yscrollcommand=scroll.set)
        scroll.grid(row=0, column=1, sticky="ns")

        actions = ttk.Frame(container)
        actions.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(10, 0))
        actions.columnconfigure((0, 1, 2, 3, 4, 5, 6), weight=1)

        ttk.Button(actions, text="Refresh", command=self._load_invoices).grid(row=0, column=0, sticky="ew", padx=4)
        ttk.Button(actions, text="Generate PDF", command=self._generate_pdf_for_selected).grid(row=0, column=1, sticky="ew", padx=4)
        ttk.Button(actions, text="Send Email", command=self._send_selected_invoice).grid(row=0, column=2, sticky="ew", padx=4)
        ttk.Button(actions, text="Mark Paid", command=self._mark_selected_paid).grid(row=0, column=3, sticky="ew", padx=4)
        ttk.Button(actions, text="Open PDF", command=self._open_selected_pdf).grid(row=0, column=4, sticky="ew", padx=4)
        ttk.Button(actions, text="Delete Invoice", command=self._delete_selected_invoice).grid(row=0, column=5, sticky="ew", padx=4)
        ttk.Button(actions, text="Send Text", command=self._send_selected_sms).grid(row=0, column=6, sticky="ew", padx=4)

    def _build_settings_tab(self) -> None:
        container = ttk.Frame(self.settings_tab)
        container.pack(fill="both", expand=True, padx=8, pady=8)
        container.columnconfigure(1, weight=1)

        self.settings_vars: dict[str, tk.StringVar] = {
            "business_name": tk.StringVar(),
            "business_email": tk.StringVar(),
            "business_phone": tk.StringVar(),
            "business_logo_path": tk.StringVar(),
            "default_tax_rate": tk.StringVar(value="0"),
            "smtp_server": tk.StringVar(value="smtp.gmail.com"),
            "smtp_port": tk.StringVar(value="587"),
            "smtp_username": tk.StringVar(),
            "smtp_password": tk.StringVar(),
            "smtp_from_email": tk.StringVar(),
            "google_credentials_file": tk.StringVar(),
            "google_token_file": tk.StringVar(),
        }
        self.smtp_use_tls_var = tk.IntVar(value=1)

        row = 0
        for label, key in [
            ("Business Name", "business_name"),
            ("Business Email", "business_email"),
            ("Business Phone", "business_phone"),
            ("Default Tax %", "default_tax_rate"),
            ("SMTP Server", "smtp_server"),
            ("SMTP Port", "smtp_port"),
            ("SMTP Username", "smtp_username"),
            ("SMTP Password", "smtp_password"),
            ("SMTP From Email", "smtp_from_email"),
            ("Google Credentials File", "google_credentials_file"),
            ("Google Token File", "google_token_file"),
        ]:
            ttk.Label(container, text=label).grid(row=row, column=0, sticky="w", padx=6, pady=4)
            show = "*" if key == "smtp_password" else ""
            ttk.Entry(container, textvariable=self.settings_vars[key], show=show).grid(
                row=row,
                column=1,
                sticky="ew",
                padx=6,
                pady=4,
            )
            row += 1

        ttk.Label(container, text="Business Logo Path").grid(row=row, column=0, sticky="w", padx=6, pady=4)
        logo_wrap = ttk.Frame(container)
        logo_wrap.grid(row=row, column=1, sticky="ew", padx=6, pady=4)
        logo_wrap.columnconfigure(0, weight=1)
        ttk.Entry(logo_wrap, textvariable=self.settings_vars["business_logo_path"]).grid(row=0, column=0, sticky="ew")
        ttk.Button(logo_wrap, text="Browse", command=self._browse_logo_path).grid(row=0, column=1, padx=(6, 0))
        row += 1

        ttk.Label(container, text="Business Address").grid(row=row, column=0, sticky="nw", padx=6, pady=4)
        self.business_address_text = tk.Text(container, height=4, width=50)
        self.business_address_text.grid(row=row, column=1, sticky="ew", padx=6, pady=4)
        row += 1

        ttk.Label(container, text="Payment Instructions").grid(row=row, column=0, sticky="nw", padx=6, pady=4)
        self.payment_instructions_text = tk.Text(container, height=4, width=50)
        self.payment_instructions_text.grid(row=row, column=1, sticky="ew", padx=6, pady=4)
        row += 1

        ttk.Checkbutton(
            container,
            text="Use TLS for SMTP",
            variable=self.smtp_use_tls_var,
            onvalue=1,
            offvalue=0,
        ).grid(row=row, column=1, sticky="w", padx=6, pady=4)
        row += 1

        actions = ttk.Frame(container)
        actions.grid(row=row, column=0, columnspan=2, sticky="ew", padx=6, pady=(10, 0))
        actions.columnconfigure((0, 1), weight=1)

        ttk.Button(actions, text="Save Settings", command=self._save_settings).grid(row=0, column=0, sticky="ew", padx=4)
        ttk.Button(actions, text="Reload", command=self._load_settings_into_form).grid(row=0, column=1, sticky="ew", padx=4)

    def _open_cleaners_popup(self) -> None:
        if self.cleaners_popup and self.cleaners_popup.winfo_exists():
            self.cleaners_popup.lift()
            self.cleaners_popup.focus_force()
            return

        popup = tk.Toplevel(self)
        popup.title("Manage Cleaners")
        popup.geometry("900x620")
        popup.minsize(760, 520)
        popup.transient(self)
        popup.protocol("WM_DELETE_WINDOW", self._close_cleaners_popup)
        self.cleaners_popup = popup

        container = ttk.Frame(popup)
        container.pack(fill="both", expand=True, padx=10, pady=10)
        container.columnconfigure(0, weight=1)
        container.rowconfigure(1, weight=1)

        cleaner_box = ttk.LabelFrame(container, text="Cleaner")
        cleaner_box.grid(row=0, column=0, sticky="ew", pady=(0, 8))
        cleaner_box.columnconfigure(1, weight=1)

        ttk.Label(cleaner_box, text="Name").grid(row=0, column=0, sticky="w", padx=6, pady=(8, 4))
        ttk.Entry(cleaner_box, textvariable=self.cleaner_name_var).grid(row=0, column=1, sticky="ew", padx=6, pady=(8, 4))

        ttk.Label(cleaner_box, text="Phone").grid(row=1, column=0, sticky="w", padx=6, pady=4)
        ttk.Entry(cleaner_box, textvariable=self.cleaner_phone_var).grid(row=1, column=1, sticky="ew", padx=6, pady=4)

        ttk.Label(cleaner_box, text="Email").grid(row=2, column=0, sticky="w", padx=6, pady=4)
        ttk.Entry(cleaner_box, textvariable=self.cleaner_email_var).grid(row=2, column=1, sticky="ew", padx=6, pady=4)

        ttk.Label(cleaner_box, text="Google Calendar ID").grid(row=3, column=0, sticky="w", padx=6, pady=4)
        ttk.Entry(cleaner_box, textvariable=self.cleaner_google_calendar_id_var).grid(row=3, column=1, sticky="ew", padx=6, pady=4)

        ttk.Checkbutton(
            cleaner_box,
            text="Auto-create Google calendar on add when empty",
            variable=self.auto_create_cleaner_calendar_var,
            onvalue=1,
            offvalue=0,
        ).grid(row=4, column=0, columnspan=2, sticky="w", padx=6, pady=(2, 4))

        ttk.Label(cleaner_box, text="Notes").grid(row=5, column=0, sticky="nw", padx=6, pady=4)
        self.cleaner_notes_text = tk.Text(cleaner_box, height=4, width=34)
        self.cleaner_notes_text.grid(row=5, column=1, sticky="ew", padx=6, pady=4)

        cleaner_btns = ttk.Frame(cleaner_box)
        cleaner_btns.grid(row=6, column=0, columnspan=2, sticky="ew", padx=6, pady=(8, 8))
        cleaner_btns.columnconfigure((0, 1, 2, 3, 4), weight=1)
        ttk.Button(cleaner_btns, text="Add Cleaner", command=self._add_cleaner).grid(row=0, column=0, sticky="ew", padx=3)
        ttk.Button(cleaner_btns, text="Update Cleaner", command=self._update_cleaner).grid(row=0, column=1, sticky="ew", padx=3)
        ttk.Button(cleaner_btns, text="Delete Cleaner", command=self._delete_cleaner).grid(row=0, column=2, sticky="ew", padx=3)
        ttk.Button(cleaner_btns, text="Clear", command=self._clear_cleaner_form).grid(row=0, column=3, sticky="ew", padx=3)
        ttk.Button(cleaner_btns, text="Close", command=self._close_cleaners_popup).grid(row=0, column=4, sticky="ew", padx=3)

        cleaner_table = ttk.LabelFrame(container, text="Cleaner List")
        cleaner_table.grid(row=1, column=0, sticky="nsew")
        cleaner_table.columnconfigure(0, weight=1)
        cleaner_table.rowconfigure(0, weight=1)

        self.cleaners_tree = ttk.Treeview(
            cleaner_table,
            columns=("id", "name", "phone", "email", "calendar"),
            show="headings",
            height=14,
        )
        self.cleaners_tree.heading("id", text="ID")
        self.cleaners_tree.heading("name", text="Name")
        self.cleaners_tree.heading("phone", text="Phone")
        self.cleaners_tree.heading("email", text="Email")
        self.cleaners_tree.heading("calendar", text="Calendar ID")
        self.cleaners_tree.column("id", width=55, anchor="center")
        self.cleaners_tree.column("name", width=170)
        self.cleaners_tree.column("phone", width=120)
        self.cleaners_tree.column("email", width=220)
        self.cleaners_tree.column("calendar", width=290)
        self.cleaners_tree.grid(row=0, column=0, sticky="nsew")
        self.cleaners_tree.bind("<<TreeviewSelect>>", self._on_cleaner_selected)

        cleaner_scroll = ttk.Scrollbar(cleaner_table, orient="vertical", command=self.cleaners_tree.yview)
        self.cleaners_tree.configure(yscrollcommand=cleaner_scroll.set)
        cleaner_scroll.grid(row=0, column=1, sticky="ns")

        self._load_cleaners()
        self._clear_cleaner_form()

    def _close_cleaners_popup(self) -> None:
        if self.cleaners_popup and self.cleaners_popup.winfo_exists():
            self.cleaners_popup.destroy()
        self.cleaners_popup = None

    def _open_google_popup(self) -> None:
        if self.google_popup and self.google_popup.winfo_exists():
            self.google_popup.lift()
            self.google_popup.focus_force()
            return

        popup = tk.Toplevel(self)
        popup.title("Google Calendar")
        popup.geometry("760x280")
        popup.minsize(680, 240)
        popup.transient(self)
        popup.protocol("WM_DELETE_WINDOW", self._close_google_popup)
        self.google_popup = popup

        container = ttk.Frame(popup)
        container.pack(fill="both", expand=True, padx=10, pady=10)
        container.columnconfigure(1, weight=1)

        ttk.Label(container, textvariable=self.google_status_var).grid(
            row=0,
            column=0,
            columnspan=2,
            sticky="w",
            pady=(0, 8),
        )

        ttk.Label(container, text="Google Calendar").grid(row=1, column=0, sticky="w", padx=(0, 6), pady=4)
        self.google_calendars_combo = ttk.Combobox(
            container,
            textvariable=self.google_calendars_combo_var,
            state="readonly",
        )
        self.google_calendars_combo.grid(row=1, column=1, sticky="ew", pady=4)

        ttk.Label(container, text="Assign To Cleaner").grid(row=2, column=0, sticky="w", padx=(0, 6), pady=4)
        self.google_target_cleaner_combo = ttk.Combobox(
            container,
            textvariable=self.google_target_cleaner_var,
            state="readonly",
        )
        self.google_target_cleaner_combo.grid(row=2, column=1, sticky="ew", pady=4)

        btns = ttk.Frame(container)
        btns.grid(row=3, column=0, columnspan=2, sticky="ew", pady=(10, 0))
        btns.columnconfigure((0, 1, 2, 3), weight=1)
        ttk.Button(btns, text="Connect Google", command=self._connect_google_calendar).grid(row=0, column=0, sticky="ew", padx=3)
        ttk.Button(btns, text="Load Calendars", command=self._load_google_calendars).grid(row=0, column=1, sticky="ew", padx=3)
        ttk.Button(btns, text="Assign To Cleaner", command=self._apply_google_calendar_to_cleaner).grid(row=0, column=2, sticky="ew", padx=3)
        ttk.Button(btns, text="Close", command=self._close_google_popup).grid(row=0, column=3, sticky="ew", padx=3)

        self._load_cleaners()
        self._refresh_google_status()

    def _close_google_popup(self) -> None:
        if self.google_popup and self.google_popup.winfo_exists():
            self.google_popup.destroy()
        self.google_popup = None

    def _open_schedule_popup(self) -> None:
        if self.schedule_popup and self.schedule_popup.winfo_exists():
            self.schedule_popup.lift()
            self.schedule_popup.focus_force()
            return

        popup = tk.Toplevel(self)
        popup.title("Schedule Job")
        popup.geometry("900x430")
        popup.minsize(820, 380)
        popup.transient(self)
        popup.protocol("WM_DELETE_WINDOW", self._close_schedule_popup)
        self.schedule_popup = popup

        container = ttk.Frame(popup)
        container.pack(fill="both", expand=True, padx=10, pady=10)
        for idx in range(4):
            container.columnconfigure(idx, weight=1)

        ttk.Label(container, text="Customer").grid(row=0, column=0, sticky="w", padx=6, pady=(4, 4))
        self.job_customer_combo = ttk.Combobox(container, textvariable=self.job_customer_var, state="readonly")
        self.job_customer_combo.grid(row=0, column=1, sticky="ew", padx=6, pady=(4, 4))

        ttk.Label(container, text="Cleaner").grid(row=0, column=2, sticky="w", padx=6, pady=(4, 4))
        self.job_cleaner_combo = ttk.Combobox(container, textvariable=self.job_cleaner_var, state="readonly")
        self.job_cleaner_combo.grid(row=0, column=3, sticky="ew", padx=6, pady=(4, 4))

        ttk.Label(container, text="Job Title").grid(row=1, column=0, sticky="w", padx=6, pady=4)
        ttk.Entry(container, textvariable=self.job_title_var).grid(row=1, column=1, sticky="ew", padx=6, pady=4)

        ttk.Label(container, text="Location").grid(row=1, column=2, sticky="w", padx=6, pady=4)
        ttk.Entry(container, textvariable=self.job_location_var).grid(row=1, column=3, sticky="ew", padx=6, pady=4)

        ttk.Label(container, text="Start (YYYY-MM-DD HH:MM)").grid(row=2, column=0, sticky="w", padx=6, pady=4)
        ttk.Entry(container, textvariable=self.job_start_var).grid(row=2, column=1, sticky="ew", padx=6, pady=4)

        ttk.Label(container, text="End (YYYY-MM-DD HH:MM)").grid(row=2, column=2, sticky="w", padx=6, pady=4)
        ttk.Entry(container, textvariable=self.job_end_var).grid(row=2, column=3, sticky="ew", padx=6, pady=4)

        ttk.Label(container, text="Notes").grid(row=3, column=0, sticky="nw", padx=6, pady=4)
        self.job_notes_text = tk.Text(container, height=6, width=50)
        self.job_notes_text.grid(row=3, column=1, columnspan=3, sticky="ew", padx=6, pady=4)

        btns = ttk.Frame(container)
        btns.grid(row=4, column=0, columnspan=4, sticky="ew", padx=6, pady=(10, 0))
        btns.columnconfigure((0, 1, 2, 3), weight=1)
        ttk.Button(btns, text="Check Availability", command=self._check_job_availability).grid(row=0, column=0, sticky="ew", padx=3)
        ttk.Button(btns, text="Schedule Job", command=self._schedule_job).grid(row=0, column=1, sticky="ew", padx=3)
        ttk.Button(btns, text="Reset", command=self._reset_job_form).grid(row=0, column=2, sticky="ew", padx=3)
        ttk.Button(btns, text="Close", command=self._close_schedule_popup).grid(row=0, column=3, sticky="ew", padx=3)

        self._load_customers()
        self._load_cleaners()
        self._reset_job_form()

    def _close_schedule_popup(self) -> None:
        if self.schedule_popup and self.schedule_popup.winfo_exists():
            self.schedule_popup.destroy()
        self.schedule_popup = None

    def _on_calendar_view_changed(self, _event: tk.Event) -> None:
        self._update_calendar_range_label()
        self._refresh_calendar_view()

    def _set_calendar_today(self) -> None:
        self.calendar_anchor_date = date.today()
        self._update_calendar_range_label()
        self._refresh_calendar_view()

    def _move_calendar_range(self, step: int) -> None:
        mode = self.calendar_view_var.get()
        if mode == "Day":
            self.calendar_anchor_date = self.calendar_anchor_date + timedelta(days=step)
        elif mode == "Week":
            self.calendar_anchor_date = self.calendar_anchor_date + timedelta(days=7 * step)
        elif mode == "Month":
            self.calendar_anchor_date = self._shift_months(self.calendar_anchor_date, step)
        else:
            try:
                self.calendar_anchor_date = self.calendar_anchor_date.replace(
                    year=self.calendar_anchor_date.year + step
                )
            except ValueError:
                # Handle Feb 29 in non-leap years.
                self.calendar_anchor_date = self.calendar_anchor_date.replace(
                    month=2,
                    day=28,
                    year=self.calendar_anchor_date.year + step,
                )

        self._update_calendar_range_label()
        self._refresh_calendar_view()

    @staticmethod
    def _shift_months(base_date: date, month_delta: int) -> date:
        total_months = (base_date.year * 12 + (base_date.month - 1)) + month_delta
        target_year = total_months // 12
        target_month = (total_months % 12) + 1
        target_day = min(base_date.day, 28)
        return date(target_year, target_month, target_day)

    def _get_calendar_window(self) -> tuple[date, date]:
        mode = self.calendar_view_var.get()

        if mode == "Day":
            return self.calendar_anchor_date, self.calendar_anchor_date

        if mode == "Week":
            start = self.calendar_anchor_date - timedelta(days=self.calendar_anchor_date.weekday())
            end = start + timedelta(days=6)
            return start, end

        if mode == "Month":
            start = self.calendar_anchor_date.replace(day=1)
            next_month = (start.replace(day=28) + timedelta(days=4)).replace(day=1)
            end = next_month - timedelta(days=1)
            return start, end

        start = date(self.calendar_anchor_date.year, 1, 1)
        end = date(self.calendar_anchor_date.year, 12, 31)
        return start, end

    def _update_calendar_range_label(self) -> None:
        start, end = self._get_calendar_window()
        if start == end:
            label = start.strftime("%A, %B %d, %Y")
        elif start.year == end.year:
            label = f"{start.strftime('%b %d')} - {end.strftime('%b %d, %Y')}"
        else:
            label = f"{start.strftime('%b %d, %Y')} - {end.strftime('%b %d, %Y')}"
        self.calendar_range_var.set(label)

    def _configure_calendar_columns(self) -> None:
        mode = self.calendar_view_var.get()
        if mode == "Day":
            columns = ("id", "time", "customer", "cleaner", "title", "status", "location")
            headings = {
                "id": "ID",
                "time": "Time",
                "customer": "Customer",
                "cleaner": "Cleaner",
                "title": "Job",
                "status": "Status",
                "location": "Location",
            }
            widths = {
                "id": 0,
                "time": 130,
                "customer": 210,
                "cleaner": 170,
                "title": 240,
                "status": 110,
                "location": 230,
            }
        elif mode == "Year":
            columns = ("id", "month", "date", "time", "cleaner", "title", "status")
            headings = {
                "id": "ID",
                "month": "Month",
                "date": "Date",
                "time": "Time",
                "cleaner": "Cleaner",
                "title": "Job",
                "status": "Status",
            }
            widths = {
                "id": 0,
                "month": 110,
                "date": 130,
                "time": 130,
                "cleaner": 190,
                "title": 320,
                "status": 110,
            }
        else:
            columns = ("id", "date", "time", "customer", "cleaner", "title", "status")
            headings = {
                "id": "ID",
                "date": "Date",
                "time": "Time",
                "customer": "Customer",
                "cleaner": "Cleaner",
                "title": "Job",
                "status": "Status",
            }
            widths = {
                "id": 0,
                "date": 130,
                "time": 130,
                "customer": 220,
                "cleaner": 180,
                "title": 320,
                "status": 110,
            }

        self.jobs_tree["columns"] = columns
        for key in columns:
            self.jobs_tree.heading(key, text=headings[key])
            anchor = "center" if key in {"status", "date", "time", "month"} else "w"
            self.jobs_tree.column(key, width=widths[key], anchor=anchor)

        # Keep internal ID column hidden while preserving row operations.
        self.jobs_tree.column("id", width=0, stretch=False)

    def _refresh_calendar_view(self) -> None:
        if not hasattr(self, "jobs_tree"):
            return

        self._configure_calendar_columns()

        for item in self.jobs_tree.get_children():
            self.jobs_tree.delete(item)

        start_date, end_date = self._get_calendar_window()
        mode = self.calendar_view_var.get()

        for row in self.calendar_jobs_cache:
            try:
                start_dt = datetime.strptime(row["start_at"], CALENDAR_DATETIME_FORMAT)
                end_dt = datetime.strptime(row["end_at"], CALENDAR_DATETIME_FORMAT)
            except ValueError:
                continue

            if not (start_date <= start_dt.date() <= end_date):
                continue

            if mode == "Day":
                values = (
                    row["id"],
                    f"{start_dt.strftime('%I:%M %p')} - {end_dt.strftime('%I:%M %p')}",
                    row["customer_name"],
                    row["cleaner_name"],
                    row["title"],
                    row["status"],
                    row["location"],
                )
            elif mode == "Year":
                values = (
                    row["id"],
                    start_dt.strftime("%B"),
                    start_dt.strftime("%Y-%m-%d"),
                    f"{start_dt.strftime('%I:%M %p')} - {end_dt.strftime('%I:%M %p')}",
                    row["cleaner_name"],
                    row["title"],
                    row["status"],
                )
            else:
                values = (
                    row["id"],
                    start_dt.strftime("%Y-%m-%d"),
                    f"{start_dt.strftime('%I:%M %p')} - {end_dt.strftime('%I:%M %p')}",
                    row["customer_name"],
                    row["cleaner_name"],
                    row["title"],
                    row["status"],
                )

            self.jobs_tree.insert("", tk.END, values=values)

    def _load_settings_into_form(self) -> None:
        settings = self.settings_service.get_settings()
        for key, var in self.settings_vars.items():
            var.set(settings.get(key, ""))

        self.smtp_use_tls_var.set(1 if settings.get("smtp_use_tls", "1") == "1" else 0)

        self.business_address_text.delete("1.0", tk.END)
        self.business_address_text.insert("1.0", settings.get("business_address", ""))

        self.payment_instructions_text.delete("1.0", tk.END)
        self.payment_instructions_text.insert("1.0", settings.get("payment_instructions", ""))

        if not self.invoice_tax_rate_var.get().strip():
            self.invoice_tax_rate_var.set(settings.get("default_tax_rate", "0"))

    def _save_settings(self) -> None:
        updates = {key: var.get().strip() for key, var in self.settings_vars.items()}
        updates["business_address"] = self.business_address_text.get("1.0", tk.END).strip()
        updates["payment_instructions"] = self.payment_instructions_text.get("1.0", tk.END).strip()
        updates["smtp_use_tls"] = "1" if self.smtp_use_tls_var.get() else "0"

        self.settings_service.save_settings(updates)
        self.google_calendar_service = None
        self.google_connected = False
        self._refresh_google_status()
        messagebox.showinfo("Saved", "Settings saved.")

    def _browse_logo_path(self) -> None:
        filename = filedialog.askopenfilename(
            title="Select logo image",
            filetypes=[("Image files", "*.png *.jpg *.jpeg *.gif")],
        )
        if filename:
            self.settings_vars["business_logo_path"].set(filename)

    def _add_customer(self) -> None:
        payload = self._get_customer_form_payload()
        if not payload:
            return

        self.customer_service.create_customer(**payload)
        self._load_customers()
        self._clear_customer_form()
        messagebox.showinfo("Customer", "Customer added.")

    def _update_customer(self) -> None:
        if not self.customer_id_var.get().strip():
            messagebox.showwarning("Select", "Select a customer first.")
            return

        payload = self._get_customer_form_payload()
        if not payload:
            return

        self.customer_service.update_customer(int(self.customer_id_var.get()), **payload)
        self._load_customers()
        messagebox.showinfo("Customer", "Customer updated.")

    def _delete_customer(self) -> None:
        if not self.customer_id_var.get().strip():
            messagebox.showwarning("Select", "Select a customer first.")
            return

        if not messagebox.askyesno("Confirm", "Delete selected customer?"):
            return

        try:
            self.customer_service.delete_customer(int(self.customer_id_var.get()))
        except ValueError as exc:
            messagebox.showwarning("Delete Blocked", str(exc))
            return
        except Exception as exc:
            messagebox.showerror("Delete Failed", f"Could not delete customer: {exc}")
            return

        self._load_customers()
        self._clear_customer_form()

    def _delete_customer_with_invoices(self) -> None:
        if not self.customer_id_var.get().strip():
            messagebox.showwarning("Select", "Select a customer first.")
            return

        customer_id = int(self.customer_id_var.get())
        customer_name = self.customer_name_var.get().strip() or f"Customer {customer_id}"
        invoice_count = self.customer_service.count_customer_invoices(customer_id)

        if invoice_count == 0:
            messagebox.showinfo("Nothing To Remove", "This customer has no linked invoices. Use Delete.")
            return

        if not messagebox.askyesno(
            "Dangerous Action",
            (
                f"This will permanently delete customer '{customer_name}' and "
                f"{invoice_count} linked invoice(s).\n\nThis cannot be undone. Continue?"
            ),
        ):
            return

        confirmation = simpledialog.askstring(
            "Final Confirmation",
            "Type DELETE to confirm permanent deletion:",
            parent=self,
        )
        if confirmation != "DELETE":
            messagebox.showinfo("Cancelled", "Deletion cancelled.")
            return

        try:
            removed_count = self.customer_service.delete_customer_with_invoices(customer_id)
        except Exception as exc:
            messagebox.showerror("Delete Failed", f"Could not delete customer and invoices: {exc}")
            return

        self._load_customers()
        self._load_invoices()
        self._clear_customer_form()
        messagebox.showinfo(
            "Deleted",
            f"Deleted customer '{customer_name}' and {removed_count} linked invoice(s).",
        )

    def _get_customer_form_payload(self) -> dict | None:
        name = self.customer_name_var.get().strip()
        email = self.customer_email_var.get().strip()
        phone = self.customer_phone_var.get().strip()
        bedrooms = self.customer_bedrooms_var.get().strip()
        bathrooms = self.customer_bathrooms_var.get().strip()
        square_feet = self.customer_square_feet_var.get().strip()
        cleaning_duration = self.customer_cleaning_duration_var.get().strip()
        frequency = self.customer_frequency_var.get().strip()
        address = self.customer_address_var.get().strip()
        notes = self.customer_notes_text.get("1.0", tk.END).strip()

        if not name:
            messagebox.showwarning("Validation", "Customer name is required.")
            return None
        if not email:
            messagebox.showwarning("Validation", "Customer email is required.")
            return None
        if not address:
            messagebox.showwarning("Validation", "Customer address is required.")
            return None

        return {
            "name": name,
            "email": email,
            "phone": phone,
            "bedrooms": bedrooms,
            "bathrooms": bathrooms,
            "square_feet": square_feet,
            "cleaning_duration": cleaning_duration,
            "frequency": frequency,
            "address": address,
            "notes": notes,
        }

    def _clear_customer_form(self) -> None:
        self.customer_id_var.set("")
        self.customer_name_var.set("")
        self.customer_email_var.set("")
        self.customer_phone_var.set("")
        self.customer_bedrooms_var.set("")
        self.customer_bathrooms_var.set("")
        self.customer_square_feet_var.set("")
        self.customer_cleaning_duration_var.set("")
        self.customer_address_var.set("")
        self.customer_frequency_var.set("Single")
        self.customer_notes_text.delete("1.0", tk.END)

    def _on_customer_selected(self, _event: tk.Event) -> None:
        selection = self.customers_tree.selection()
        if not selection:
            return

        row_values = self.customers_tree.item(selection[0], "values")
        customer_id = int(row_values[0])
        customer = self.customer_service.get_customer(customer_id)
        if not customer:
            return

        self.customer_id_var.set(str(customer["id"]))
        self.customer_name_var.set(customer["name"])
        self.customer_email_var.set(customer["email"])
        self.customer_phone_var.set(customer["phone"])
        self.customer_bedrooms_var.set(customer["bedrooms"])
        self.customer_bathrooms_var.set(customer["bathrooms"])
        self.customer_square_feet_var.set(customer["square_feet"])
        self.customer_cleaning_duration_var.set(customer["cleaning_duration"])

        self.customer_address_var.set(customer["address"])
        self.customer_frequency_var.set(customer.get("frequency") or "Single")
        self.customer_notes_text.delete("1.0", tk.END)
        self.customer_notes_text.insert("1.0", customer["notes"])

    def _load_customers(self) -> None:
        from datetime import date as date_type

        rows = self.customer_service.list_customers()

        for item in self.customers_tree.get_children():
            self.customers_tree.delete(item)

        customer_labels: list[str] = []
        self.customer_lookup.clear()
        self.calendar_customer_lookup.clear()

        today = date_type.today()

        for row in rows:
            one_line_address = row["address"].replace("\n", " ")
            frequency = row.get("frequency") or "Single"

            last_invoice_date = row.get("last_invoice_date")
            if last_invoice_date:
                try:
                    last_date = date_type.fromisoformat(last_invoice_date)
                    days_ago = (today - last_date).days
                    last_clean = f"{days_ago}d ago"
                except ValueError:
                    last_clean = "—"
            else:
                last_clean = "Never"

            self.customers_tree.insert(
                "",
                tk.END,
                values=(row["id"], row["name"], row["email"], row["phone"], one_line_address, frequency, last_clean),
                tags=(frequency,),
            )
            label = f"{row['id']} - {row['name']} <{row['email']}>"
            customer_labels.append(label)
            self.customer_lookup[label] = row["id"]
            self.calendar_customer_lookup[label] = row["id"]

        self.invoice_customer_combo["values"] = customer_labels
        if hasattr(self, "job_customer_combo") and self.job_customer_combo.winfo_exists():
            self.job_customer_combo["values"] = customer_labels

    def _clear_cleaner_form(self) -> None:
        self.cleaner_id_var.set("")
        self.cleaner_name_var.set("")
        self.cleaner_phone_var.set("")
        self.cleaner_email_var.set("")
        self.cleaner_google_calendar_id_var.set("")
        if hasattr(self, "cleaner_notes_text") and self.cleaner_notes_text.winfo_exists():
            self.cleaner_notes_text.delete("1.0", tk.END)

    def _load_cleaners(self) -> None:
        rows = self.calendar_service.list_cleaners()

        if hasattr(self, "cleaners_tree") and self.cleaners_tree.winfo_exists():
            for item in self.cleaners_tree.get_children():
                self.cleaners_tree.delete(item)

        cleaner_labels: list[str] = []
        self.cleaner_lookup.clear()

        for row in rows:
            if hasattr(self, "cleaners_tree") and self.cleaners_tree.winfo_exists():
                self.cleaners_tree.insert(
                    "",
                    tk.END,
                    values=(
                        row["id"],
                        row["name"],
                        row["phone"],
                        row["email"],
                        row["google_calendar_id"],
                    ),
                )
            label = f"{row['id']} - {row['name']}"
            cleaner_labels.append(label)
            self.cleaner_lookup[label] = row["id"]

        if hasattr(self, "job_cleaner_combo") and self.job_cleaner_combo.winfo_exists():
            self.job_cleaner_combo["values"] = cleaner_labels

        if hasattr(self, "google_target_cleaner_combo") and self.google_target_cleaner_combo.winfo_exists():
            self.google_target_cleaner_combo["values"] = cleaner_labels

    def _get_google_calendar_service(self) -> GoogleCalendarService:
        settings = self.settings_service.get_settings()
        credentials_file = settings.get("google_credentials_file", "").strip()
        token_file = settings.get("google_token_file", "").strip()

        if not credentials_file or not token_file:
            raise ValueError("Set Google credentials and token file paths in Settings first.")

        if (
            self.google_calendar_service is None
            or self.google_calendar_service.credentials_file != credentials_file
            or self.google_calendar_service.token_file != token_file
        ):
            self.google_calendar_service = GoogleCalendarService(credentials_file, token_file)
            self.google_connected = False

        return self.google_calendar_service

    def _refresh_google_status(self) -> None:
        try:
            service = self._get_google_calendar_service()
            if self.google_connected:
                status = "Status: Connected"
            elif service.is_configured():
                status = "Status: Configured (connect required)"
            else:
                status = "Status: Credentials file missing"
        except Exception as exc:
            status = f"Status: Not configured ({exc})"

        if hasattr(self, "google_status_var"):
            self.google_status_var.set(status)

    def _connect_google_calendar(self) -> None:
        try:
            service = self._get_google_calendar_service()
            service.connect()
            self.google_connected = True
            self._refresh_google_status()
        except Exception as exc:
            messagebox.showerror("Google Calendar", str(exc))
            return

        self._load_google_calendars()

    def _load_google_calendars(self) -> None:
        try:
            service = self._get_google_calendar_service()
            calendars = service.list_calendars()
            self.google_connected = True
            self._refresh_google_status()
        except Exception as exc:
            messagebox.showerror("Google Calendar", str(exc))
            return

        labels: list[str] = []
        self.google_calendar_lookup.clear()

        for cal in calendars:
            cal_id = cal.get("id", "").strip()
            summary = cal.get("summary", "Untitled")
            if not cal_id:
                continue
            label = f"{summary} ({cal_id})"
            labels.append(label)
            self.google_calendar_lookup[label] = cal_id

        if hasattr(self, "google_calendars_combo") and self.google_calendars_combo.winfo_exists():
            self.google_calendars_combo["values"] = labels
        if labels:
            self.google_calendars_combo_var.set(labels[0])
            messagebox.showinfo("Google Calendar", f"Loaded {len(labels)} calendar(s).")
        else:
            messagebox.showwarning("Google Calendar", "No calendars were returned by Google.")

    def _apply_google_calendar_to_cleaner(self) -> None:
        selected = self.google_calendars_combo_var.get().strip()
        if selected not in self.google_calendar_lookup:
            messagebox.showwarning("Google Calendar", "Select a calendar first.")
            return

        cleaner_label = self.google_target_cleaner_var.get().strip()
        if cleaner_label not in self.cleaner_lookup:
            messagebox.showwarning("Google Calendar", "Select a cleaner to assign this calendar to.")
            return

        cleaner_id = self.cleaner_lookup[cleaner_label]
        cleaner = self.calendar_service.get_cleaner(cleaner_id)
        if not cleaner:
            messagebox.showwarning("Google Calendar", "Selected cleaner was not found.")
            return

        calendar_id = self.google_calendar_lookup[selected]
        self.calendar_service.update_cleaner(
            cleaner_id=cleaner_id,
            name=cleaner["name"],
            phone=cleaner["phone"],
            email=cleaner["email"],
            notes=cleaner["notes"],
            google_calendar_id=calendar_id,
        )

        self.cleaner_google_calendar_id_var.set(calendar_id)
        self._load_cleaners()
        messagebox.showinfo("Google Calendar", "Calendar assigned to cleaner.")

    def _on_cleaner_selected(self, _event: tk.Event) -> None:
        selection = self.cleaners_tree.selection()
        if not selection:
            return

        values = self.cleaners_tree.item(selection[0], "values")
        cleaner_id = int(values[0])
        cleaner = self.calendar_service.get_cleaner(cleaner_id)
        if not cleaner:
            return

        self.cleaner_id_var.set(str(cleaner["id"]))
        self.cleaner_name_var.set(cleaner["name"])
        self.cleaner_phone_var.set(cleaner["phone"])
        self.cleaner_email_var.set(cleaner["email"])
        self.cleaner_google_calendar_id_var.set(cleaner["google_calendar_id"])
        if hasattr(self, "cleaner_notes_text") and self.cleaner_notes_text.winfo_exists():
            self.cleaner_notes_text.delete("1.0", tk.END)
            self.cleaner_notes_text.insert("1.0", cleaner["notes"])

    def _add_cleaner(self) -> None:
        name = self.cleaner_name_var.get().strip()
        if not name:
            messagebox.showwarning("Validation", "Cleaner name is required.")
            return

        google_calendar_id = self.cleaner_google_calendar_id_var.get().strip()
        if not google_calendar_id and self.auto_create_cleaner_calendar_var.get() == 1:
            try:
                service = self._get_google_calendar_service()
                created = service.create_calendar(
                    summary=f"Cleaner - {name}",
                    description=(
                        "Auto-created by Cleaning Invoice Generator "
                        f"for cleaner '{name}'."
                    ),
                )
                google_calendar_id = created.get("id", "").strip()
                self.cleaner_google_calendar_id_var.set(google_calendar_id)
                self.google_connected = True
                self._refresh_google_status()
            except Exception as exc:
                proceed = messagebox.askyesno(
                    "Google Calendar",
                    (
                        "Could not auto-create a Google calendar for this cleaner:\n"
                        f"{exc}\n\n"
                        "Add cleaner without Google calendar?"
                    ),
                )
                if not proceed:
                    return

        self.calendar_service.create_cleaner(
            name=name,
            phone=self.cleaner_phone_var.get().strip(),
            email=self.cleaner_email_var.get().strip(),
            notes=self.cleaner_notes_text.get("1.0", tk.END).strip(),
            google_calendar_id=google_calendar_id,
        )
        self._load_cleaners()
        self._clear_cleaner_form()
        if google_calendar_id:
            messagebox.showinfo("Cleaner", "Cleaner added and linked to Google Calendar.")
        else:
            messagebox.showinfo("Cleaner", "Cleaner added.")

    def _update_cleaner(self) -> None:
        if not self.cleaner_id_var.get().strip():
            messagebox.showwarning("Select", "Select a cleaner first.")
            return

        name = self.cleaner_name_var.get().strip()
        if not name:
            messagebox.showwarning("Validation", "Cleaner name is required.")
            return

        self.calendar_service.update_cleaner(
            cleaner_id=int(self.cleaner_id_var.get()),
            name=name,
            phone=self.cleaner_phone_var.get().strip(),
            email=self.cleaner_email_var.get().strip(),
            notes=self.cleaner_notes_text.get("1.0", tk.END).strip(),
            google_calendar_id=self.cleaner_google_calendar_id_var.get().strip(),
        )
        self._load_cleaners()
        messagebox.showinfo("Cleaner", "Cleaner updated.")

    def _delete_cleaner(self) -> None:
        if not self.cleaner_id_var.get().strip():
            messagebox.showwarning("Select", "Select a cleaner first.")
            return

        if not messagebox.askyesno("Confirm", "Delete selected cleaner?"):
            return

        cleaner_id = int(self.cleaner_id_var.get())
        cleaner = self.calendar_service.get_cleaner(cleaner_id)

        try:
            google_calendar_id = (cleaner or {}).get("google_calendar_id", "").strip()
            if google_calendar_id:
                active_with_same_calendar = [
                    row
                    for row in self.calendar_service.list_cleaners()
                    if int(row["id"]) != cleaner_id and row.get("google_calendar_id", "").strip() == google_calendar_id
                ]

                if active_with_same_calendar:
                    messagebox.showwarning(
                        "Google Calendar",
                        "This calendar is assigned to another cleaner, so it was not deleted.",
                    )
                else:
                    try:
                        service = self._get_google_calendar_service()
                        service.delete_calendar(google_calendar_id)
                        self.google_connected = True
                        self._refresh_google_status()
                    except Exception as exc:
                        proceed = messagebox.askyesno(
                            "Google Calendar",
                            (
                                "Could not delete the linked Google calendar:\n"
                                f"{exc}\n\n"
                                "Delete cleaner locally anyway?"
                            ),
                        )
                        if not proceed:
                            return

            self.calendar_service.delete_cleaner(cleaner_id)
        except ValueError as exc:
            messagebox.showwarning("Delete Blocked", str(exc))
            return

        self._load_cleaners()
        self._clear_cleaner_form()
        messagebox.showinfo("Cleaner", "Cleaner deleted.")

    def _reset_job_form(self) -> None:
        now = datetime.now().replace(minute=0, second=0, microsecond=0)
        self.job_customer_var.set("")
        self.job_cleaner_var.set("")
        self.job_title_var.set("")
        self.job_location_var.set("")
        self.job_start_var.set(now.strftime(CALENDAR_DATETIME_FORMAT))
        self.job_end_var.set((now + timedelta(hours=2)).strftime(CALENDAR_DATETIME_FORMAT))
        if hasattr(self, "job_notes_text") and self.job_notes_text.winfo_exists():
            self.job_notes_text.delete("1.0", tk.END)

    def _selected_job_id(self) -> int | None:
        selection = self.jobs_tree.selection()
        if not selection:
            messagebox.showwarning("Select", "Select a job first.")
            return None

        values = self.jobs_tree.item(selection[0], "values")
        return int(values[0])

    def _load_jobs(self) -> None:
        self.calendar_jobs_cache = self.calendar_service.list_jobs()
        self._refresh_calendar_view()

    def _validate_job_form(self) -> tuple[int, int, str, str, str, str, str]:
        if not hasattr(self, "job_notes_text") or not self.job_notes_text.winfo_exists():
            raise ValueError("Open the Schedule Job popup first.")

        customer_label = self.job_customer_var.get().strip()
        cleaner_label = self.job_cleaner_var.get().strip()

        if customer_label not in self.calendar_customer_lookup:
            raise ValueError("Select a customer.")
        if cleaner_label not in self.cleaner_lookup:
            raise ValueError("Select a cleaner.")

        title = self.job_title_var.get().strip()
        if not title:
            raise ValueError("Job title is required.")

        start_at = self.job_start_var.get().strip()
        end_at = self.job_end_var.get().strip()

        try:
            start_dt = datetime.strptime(start_at, CALENDAR_DATETIME_FORMAT)
            end_dt = datetime.strptime(end_at, CALENDAR_DATETIME_FORMAT)
        except ValueError as exc:
            raise ValueError("Start/End must use format YYYY-MM-DD HH:MM.") from exc

        if end_dt <= start_dt:
            raise ValueError("End time must be after start time.")

        location = self.job_location_var.get().strip()
        notes = self.job_notes_text.get("1.0", tk.END).strip()

        return (
            self.calendar_customer_lookup[customer_label],
            self.cleaner_lookup[cleaner_label],
            title,
            start_at,
            end_at,
            location,
            notes,
        )

    def _check_job_availability(self) -> None:
        try:
            _, cleaner_id, _, start_at, end_at, _, _ = self._validate_job_form()
        except ValueError as exc:
            messagebox.showwarning("Validation", str(exc))
            return

        conflicts = self.calendar_service.check_cleaner_availability(
            cleaner_id=cleaner_id,
            start_at=start_at,
            end_at=end_at,
        )

        if not conflicts:
            cleaner = self.calendar_service.get_cleaner(cleaner_id)
            google_calendar_id = (cleaner or {}).get("google_calendar_id", "").strip()
            if google_calendar_id:
                try:
                    service = self._get_google_calendar_service()
                    is_busy = service.check_time_busy(
                        google_calendar_id,
                        start_at,
                        end_at,
                        cleaner_id=cleaner_id,
                    )
                    self.google_connected = True
                    self._refresh_google_status()
                except Exception as exc:
                    messagebox.showwarning(
                        "Availability",
                        (
                            "No local conflicts found, but Google Calendar check failed: "
                            f"{exc}"
                        ),
                    )
                    return

                if is_busy:
                    messagebox.showwarning(
                        "Not Available",
                        "Cleaner has a conflict in Google Calendar for the selected time range.",
                    )
                    return

                messagebox.showinfo(
                    "Availability",
                    "Cleaner is available (local schedule + Google Calendar check).",
                )
                return

            messagebox.showinfo("Availability", "Cleaner is available for the selected time range.")
            return

        conflict_text = "\n".join(
            f"- {job['start_at']} to {job['end_at']}: {job['title']}"
            for job in conflicts[:5]
        )
        messagebox.showwarning(
            "Not Available",
            "Cleaner has conflicting jobs:\n" + conflict_text,
        )

    def _schedule_job(self) -> None:
        try:
            customer_id, cleaner_id, title, start_at, end_at, location, notes = self._validate_job_form()
            cleaner = self.calendar_service.get_cleaner(cleaner_id)
            google_calendar_id = (cleaner or {}).get("google_calendar_id", "").strip()
            if google_calendar_id:
                service = self._get_google_calendar_service()
                is_busy = service.check_time_busy(
                    google_calendar_id,
                    start_at,
                    end_at,
                    cleaner_id=cleaner_id,
                )
                self.google_connected = True
                self._refresh_google_status()
                if is_busy:
                    raise ValueError(
                        "Schedule blocked: cleaner has a conflict in Google Calendar for that time range."
                    )

            job_id = self.calendar_service.create_job(
                customer_id=customer_id,
                cleaner_id=cleaner_id,
                title=title,
                start_at=start_at,
                end_at=end_at,
                location=location,
                notes=notes,
            )
        except ValueError as exc:
            messagebox.showwarning("Schedule Blocked", str(exc))
            return
        except Exception as exc:
            messagebox.showwarning("Google Calendar", str(exc))
            return

        if google_calendar_id:
            try:
                service = self._get_google_calendar_service()
                description = notes or "Scheduled from Cleaning Invoice Generator"
                google_event_id = service.create_event(
                    calendar_id=google_calendar_id,
                    title=title,
                    start_at=start_at,
                    end_at=end_at,
                    location=location,
                    description=description,
                    cleaner_id=cleaner_id,
                )
                if google_event_id:
                    self.calendar_service.set_job_google_event_id(job_id, google_event_id)
                self.google_connected = True
                self._refresh_google_status()
            except Exception as exc:
                messagebox.showwarning(
                    "Google Sync",
                    f"Job was scheduled locally, but Google sync failed: {exc}",
                )

        self._load_jobs()
        messagebox.showinfo("Job", "Job scheduled successfully.")

    def _update_selected_job_status(self, status: str) -> None:
        job_id = self._selected_job_id()
        if job_id is None:
            return

        if status == "cancelled":
            job = self.calendar_service.get_job(job_id)
            if job:
                google_event_id = (job.get("google_event_id") or "").strip()
                google_calendar_id = (job.get("google_calendar_id") or "").strip()
                if google_event_id and google_calendar_id:
                    try:
                        service = self._get_google_calendar_service()
                        service.delete_event(google_calendar_id, google_event_id)
                        self.calendar_service.set_job_google_event_id(job_id, "")
                        self.google_connected = True
                        self._refresh_google_status()
                    except Exception:
                        pass

        self.calendar_service.update_job_status(job_id, status)
        self._load_jobs()
        messagebox.showinfo("Job", f"Job updated to '{status}'.")

    def _delete_selected_job(self) -> None:
        job_id = self._selected_job_id()
        if job_id is None:
            return

        if not messagebox.askyesno("Confirm", "Delete selected job?"):
            return

        job = self.calendar_service.get_job(job_id)
        if job:
            google_event_id = (job.get("google_event_id") or "").strip()
            google_calendar_id = (job.get("google_calendar_id") or "").strip()
            if google_event_id and google_calendar_id:
                try:
                    service = self._get_google_calendar_service()
                    service.delete_event(google_calendar_id, google_event_id)
                    self.google_connected = True
                    self._refresh_google_status()
                except Exception:
                    # Local deletion should still proceed even if Google event removal fails.
                    pass

        self.calendar_service.delete_job(job_id)
        self._load_jobs()
        messagebox.showinfo("Job", "Job deleted.")

    def _reset_invoice_form(self) -> None:
        today = date.today()
        self.invoice_issue_date_var.set(today.strftime(DATE_FORMAT))
        self.invoice_due_date_var.set((today + timedelta(days=DEFAULT_DUE_DAYS)).strftime(DATE_FORMAT))

        settings = self.settings_service.get_settings()
        self.invoice_tax_rate_var.set(settings.get("default_tax_rate", "0"))

        self.invoice_customer_var.set("")
        self.invoice_notes_text.delete("1.0", tk.END)

        self.item_description_var.set("")
        self.item_rate_var.set("0")

        self.current_items = []
        self._refresh_invoice_item_table()

    def _add_invoice_item(self) -> None:
        description = self.item_description_var.get().strip()
        if not description:
            messagebox.showwarning("Validation", "Service description is required.")
            return

        rate = self._parse_non_negative_number(self.item_rate_var.get(), "Rate")
        if rate is None:
            return

        self.current_items.append(
            {
                "description": description,
                "quantity": 1.0,
                "unit_price": rate,
            }
        )

        self.item_description_var.set("")
        self.item_rate_var.set("0")
        self._refresh_invoice_item_table()

    def _remove_selected_item(self) -> None:
        selection = self.items_tree.selection()
        if not selection:
            messagebox.showwarning("Select", "Select an item row first.")
            return

        index = self.items_tree.index(selection[0])
        if 0 <= index < len(self.current_items):
            del self.current_items[index]
        self._refresh_invoice_item_table()

    def _clear_items(self) -> None:
        self.current_items = []
        self._refresh_invoice_item_table()

    def _refresh_invoice_item_table(self) -> None:
        for item in self.items_tree.get_children():
            self.items_tree.delete(item)

        for item in self.current_items:
            line_total = float(item["quantity"]) * float(item["unit_price"])
            self.items_tree.insert(
                "",
                tk.END,
                values=(
                    item["description"],
                    format_usd(float(item["unit_price"])),
                    format_usd(line_total),
                ),
            )

        tax_rate = self._parse_non_negative_number(self.invoice_tax_rate_var.get(), "Tax")
        tax_rate = tax_rate if tax_rate is not None else 0

        totals = self.invoice_service.calculate_totals(self.current_items, tax_rate)
        self.subtotal_label.configure(text=f"Subtotal: {format_usd(float(totals.subtotal))}")
        self.tax_label.configure(text=f"Tax: {format_usd(float(totals.tax_amount))}")
        self.total_label.configure(text=f"Total: {format_usd(float(totals.total_amount))}")

    def _save_invoice(self, mode: str) -> None:
        customer_label = self.invoice_customer_var.get().strip()
        if customer_label not in self.customer_lookup:
            messagebox.showwarning("Validation", "Select a customer.")
            return

        if not self.current_items:
            messagebox.showwarning("Validation", "Add at least one service item.")
            return

        issue_date = self.invoice_issue_date_var.get().strip()
        due_date = self.invoice_due_date_var.get().strip()

        if not self._valid_date(issue_date):
            messagebox.showwarning("Validation", f"Issue date must match {DATE_FORMAT}.")
            return

        if not self._valid_date(due_date):
            messagebox.showwarning("Validation", f"Due date must match {DATE_FORMAT}.")
            return

        tax_rate = self._parse_non_negative_number(self.invoice_tax_rate_var.get(), "Tax")
        if tax_rate is None:
            return

        notes = self.invoice_notes_text.get("1.0", tk.END).strip()

        saved = self.invoice_service.create_invoice(
            customer_id=self.customer_lookup[customer_label],
            issue_date=issue_date,
            due_date=due_date,
            tax_rate=tax_rate,
            notes=notes,
            items=self.current_items,
            status="draft",
        )

        invoice_id = int(saved["invoice_id"])
        invoice_number = saved["invoice_number"]
        sms_delivery_gateways: list[str] = []

        if mode in {"pdf", "email"}:
            self._generate_pdf(invoice_id)

        if mode == "email":
            self._send_invoice(invoice_id)
        elif mode == "sms":
            sms_delivery_gateways = self._send_invoice_sms(invoice_id)

        self._load_invoices()
        self._load_customers()
        self._reset_invoice_form()

        if mode == "draft":
            messagebox.showinfo("Saved", f"Invoice {invoice_number} saved as draft.")
        elif mode == "pdf":
            messagebox.showinfo("Saved", f"Invoice {invoice_number} saved and PDF generated.")
        elif mode == "sms":
            messagebox.showinfo(
                "Saved",
                (
                    f"Invoice {invoice_number} saved and text sent "
                    f"(gateway attempts accepted: {len(sms_delivery_gateways)})."
                ),
            )
        else:
            messagebox.showinfo("Saved", f"Invoice {invoice_number} saved, PDF generated, and email sent.")

    def _generate_pdf(self, invoice_id: int) -> str:
        details = self.invoice_service.get_invoice_details(invoice_id)
        if not details:
            raise ValueError("Invoice not found")

        settings = self.settings_service.get_settings()
        pdf_path = self.pdf_service.generate_invoice_pdf(details, settings)
        self.invoice_service.save_pdf_path(invoice_id, pdf_path)
        return pdf_path

    def _send_invoice(self, invoice_id: int) -> None:
        details = self.invoice_service.get_invoice_details(invoice_id)
        if not details:
            raise ValueError("Invoice not found")

        customer_email = (details.get("customer_email") or "").strip()
        if not customer_email:
            raise ValueError("Customer email is missing.")

        pdf_path = details.get("pdf_path", "").strip()
        if not pdf_path or not Path(pdf_path).exists():
            pdf_path = self._generate_pdf(invoice_id)

        settings = self.settings_service.get_settings()
        self.email_service.send_invoice_email(
            settings=settings,
            recipient_email=customer_email,
            customer_name=details.get("customer_name", "Customer"),
            invoice_number=details["invoice_number"],
            pdf_path=pdf_path,
            total_amount=float(details["total_amount"]),
            due_date=details["due_date"],
        )
        self.invoice_service.mark_sent(invoice_id)

    def _send_invoice_sms(self, invoice_id: int) -> list[str]:
        details = self.invoice_service.get_invoice_details(invoice_id)
        if not details:
            raise ValueError("Invoice not found")

        customer_phone = (details.get("customer_phone") or "").strip()
        if not customer_phone:
            raise ValueError("Customer phone is required. Update it in the Customers tab first.")

        settings = self.settings_service.get_settings()
        delivered_gateways = self.sms_service.send_invoice_text(
            settings=settings,
            recipient_phone=customer_phone,
            customer_name=details.get("customer_name", "Customer"),
            invoice_number=details["invoice_number"],
            total_amount=float(details["total_amount"]),
            due_date=details["due_date"],
        )
        self.invoice_service.mark_sent(invoice_id)
        return delivered_gateways

    def _load_invoices(self) -> None:
        rows = self.invoice_service.list_invoices()

        for item in self.invoices_tree.get_children():
            self.invoices_tree.delete(item)

        for row in rows:
            sent_at = row["sent_at"] or ""
            self.invoices_tree.insert(
                "",
                tk.END,
                values=(
                    row["id"],
                    row["invoice_number"],
                    row["customer_name"],
                    row["issue_date"],
                    row["due_date"],
                    format_usd(float(row["total_amount"])),
                    row["status"],
                    sent_at,
                ),
            )

    def _selected_invoice_id(self) -> int | None:
        selected = self.invoices_tree.selection()
        if not selected:
            messagebox.showwarning("Select", "Select an invoice first.")
            return None

        values = self.invoices_tree.item(selected[0], "values")
        return int(values[0])

    def _generate_pdf_for_selected(self) -> None:
        invoice_id = self._selected_invoice_id()
        if invoice_id is None:
            return

        try:
            pdf_path = self._generate_pdf(invoice_id)
        except Exception as exc:
            messagebox.showerror("PDF Error", str(exc))
            return

        self._load_invoices()
        messagebox.showinfo("PDF", f"PDF generated:\n{pdf_path}")

    def _send_selected_invoice(self) -> None:
        invoice_id = self._selected_invoice_id()
        if invoice_id is None:
            return

        try:
            self._send_invoice(invoice_id)
        except Exception as exc:
            messagebox.showerror("Email Error", str(exc))
            return

        self._load_invoices()
        messagebox.showinfo("Email", "Invoice email sent.")

    def _send_selected_sms(self) -> None:
        invoice_id = self._selected_invoice_id()
        if invoice_id is None:
            return

        try:
            delivered_gateways = self._send_invoice_sms(invoice_id)
        except Exception as exc:
            messagebox.showerror("SMS Error", str(exc))
            return

        self._load_invoices()
        messagebox.showinfo(
            "Text",
            (
                "Invoice text sent through carrier email gateways "
                f"(accepted by {len(delivered_gateways)} gateway(s))."
            ),
        )

    def _mark_selected_paid(self) -> None:
        invoice_id = self._selected_invoice_id()
        if invoice_id is None:
            return

        self.invoice_service.mark_paid(invoice_id)
        self._load_invoices()
        messagebox.showinfo("Invoice", "Invoice marked as paid.")

    def _delete_selected_invoice(self) -> None:
        selected = self.invoices_tree.selection()
        if not selected:
            messagebox.showwarning("Select", "Select an invoice first.")
            return

        values = self.invoices_tree.item(selected[0], "values")
        invoice_id = int(values[0])
        invoice_number = values[1]

        if not messagebox.askyesno(
            "Confirm",
            f"Delete invoice {invoice_number}? This cannot be undone.",
        ):
            return

        self.invoice_service.delete_invoice(invoice_id)
        self._load_invoices()
        messagebox.showinfo("Invoice", f"Invoice {invoice_number} deleted.")

    def _open_selected_pdf(self) -> None:
        invoice_id = self._selected_invoice_id()
        if invoice_id is None:
            return

        details = self.invoice_service.get_invoice_details(invoice_id)
        if not details:
            messagebox.showerror("Error", "Invoice not found.")
            return

        pdf_path = (details.get("pdf_path") or "").strip()
        if not pdf_path:
            messagebox.showwarning("Missing", "No PDF generated yet.")
            return

        pdf_file = Path(pdf_path)
        if not pdf_file.exists():
            messagebox.showwarning("Missing", "PDF file path exists in DB but file is missing.")
            return

        self._open_file(pdf_file)

    @staticmethod
    def _open_file(path: Path) -> None:
        try:
            if sys.platform == "win32":
                os.startfile(str(path))
            elif sys.platform == "darwin":
                subprocess.run(["open", str(path)], check=True)
            else:
                subprocess.run(["xdg-open", str(path)], check=True)
        except Exception as exc:
            messagebox.showerror("Open Failed", f"Could not open file:\n{exc}")

    @staticmethod
    def _parse_positive_number(value: str, field_name: str) -> float | None:
        try:
            number = float(value)
        except ValueError:
            messagebox.showwarning("Validation", f"{field_name} must be a number.")
            return None

        if number <= 0:
            messagebox.showwarning("Validation", f"{field_name} must be greater than zero.")
            return None

        return number

    @staticmethod
    def _parse_non_negative_number(value: str, field_name: str) -> float | None:
        try:
            number = float(value)
        except ValueError:
            messagebox.showwarning("Validation", f"{field_name} must be a number.")
            return None

        if number < 0:
            messagebox.showwarning("Validation", f"{field_name} cannot be negative.")
            return None

        return number

    @staticmethod
    def _valid_date(value: str) -> bool:
        try:
            datetime.strptime(value, DATE_FORMAT)
            return True
        except ValueError:
            return False
