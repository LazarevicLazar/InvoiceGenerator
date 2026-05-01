"""Main application window orchestrator."""

import tkinter as tk
from datetime import date
from tkinter import messagebox, ttk

from app.config import APP_NAME
from app.database import DatabaseManager
from app.services.calendar_service import CalendarService
from app.services.customer_service import CustomerService
from app.services.email_service import EmailService
from app.services.google_calendar_service import GoogleCalendarService
from app.services.invoice_service import InvoiceService
from app.services.pdf_service import PDFService
from app.services.sms_service import SMSService
from app.services.settings_service import SettingsService
from app.ui.dialogs import (
    open_cleaners_popup,
    open_google_popup,
    open_schedule_popup,
    load_cleaners,
    check_job_availability,
    schedule_job,
    update_selected_job_status,
    delete_selected_job,
    create_invoice_from_job,
    quick_rebook_job,
    load_jobs,
    reset_job_form,
)
from app.ui.invoice_actions import generate_pdf, send_invoice, send_invoice_sms
from app.ui.tabs.customers_tab import (
    build_customers_tab,
    add_customer,
    update_customer,
    delete_customer,
    delete_customer_with_invoices,
    clear_customer_form,
    on_customer_selected,
    load_customers,
)
from app.ui.tabs.invoice_tab import (
    build_invoice_tab,
    save_invoice,
)
from app.ui.tabs.calendar_tab import build_calendar_tab
from app.ui.tabs.invoices_tab import (
    build_invoices_tab,
    load_invoices,
)
from app.ui.tabs.dashboard_tab import (
    build_dashboard_tab,
    refresh_dashboard,
)
from app.ui.tabs.settings_tab import (
    build_settings_tab,
    load_settings_into_form,
)


class InvoiceGeneratorApp(tk.Tk):
    """Main application window."""

    def __init__(self, db: DatabaseManager) -> None:
        super().__init__()
        self.db = db

        # Initialize services
        self.customer_service = CustomerService(db)
        self.calendar_service = CalendarService(db)
        self.settings_service = SettingsService(db)
        self.invoice_service = InvoiceService(db)
        self.pdf_service = PDFService()
        self.email_service = EmailService()
        self.sms_service = SMSService()
        self.google_calendar_service: GoogleCalendarService | None = None
        self.google_connected = False

        # Initialize state variables
        self.google_calendar_lookup: dict[str, str] = {}
        self.customer_lookup: dict[str, int] = {}
        self.calendar_customer_lookup: dict[str, int] = {}
        self.cleaner_lookup: dict[str, int] = {}
        self.current_items: list[dict] = []
        self.calendar_jobs_cache: list[dict] = []
        self.calendar_anchor_date = date.today()
        self.customer_frequency_var = tk.StringVar(value="Single")

        # Initialize popup references
        self.cleaners_popup: tk.Toplevel | None = None
        self.google_popup: tk.Toplevel | None = None
        self.schedule_popup: tk.Toplevel | None = None

        # Configure window
        self.title(APP_NAME)
        self.geometry("1280x780")
        self.minsize(1100, 680)

        # Build UI
        self._build_ui()
        
        # Load initial data
        load_settings_into_form(self)
        self._refresh_google_status()
        load_customers(self)
        load_cleaners(self)
        load_jobs(self)
        load_invoices(self)
        reset_job_form(self)
        refresh_dashboard(self)

    def _build_ui(self) -> None:
        """Build the main UI with all tabs."""
        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill="both", expand=True, padx=10, pady=10)

        # Create tab frames
        self.customers_tab = ttk.Frame(self.notebook)
        self.invoice_tab = ttk.Frame(self.notebook)
        self.calendar_tab = ttk.Frame(self.notebook)
        self.invoices_tab = ttk.Frame(self.notebook)
        self.dashboard_tab = ttk.Frame(self.notebook)
        self.settings_tab = ttk.Frame(self.notebook)

        # Add tabs to notebook
        self.notebook.add(self.customers_tab, text="Customers")
        self.notebook.add(self.invoice_tab, text="Create Invoice")
        self.notebook.add(self.calendar_tab, text="Calendar")
        self.notebook.add(self.invoices_tab, text="Invoice History")
        self.notebook.add(self.dashboard_tab, text="Dashboard")
        self.notebook.add(self.settings_tab, text="Settings")

        # Build each tab
        build_customers_tab(self)
        build_invoice_tab(self)
        build_calendar_tab(self)
        build_invoices_tab(self)
        build_dashboard_tab(self)
        build_settings_tab(self)

    # ===== CUSTOMERS TAB =====
    def _add_customer(self) -> None:
        """Proxy to customers_tab.add_customer."""
        add_customer(self)

    def _update_customer(self) -> None:
        """Proxy to customers_tab.update_customer."""
        update_customer(self)

    def _delete_customer(self) -> None:
        """Proxy to customers_tab.delete_customer."""
        delete_customer(self)

    def _delete_customer_with_invoices(self) -> None:
        """Proxy to customers_tab.delete_customer_with_invoices."""
        delete_customer_with_invoices(self)

    def _clear_customer_form(self) -> None:
        """Proxy to customers_tab.clear_customer_form."""
        clear_customer_form(self)

    def _on_customer_selected(self, event) -> None:
        """Proxy to customers_tab.on_customer_selected."""
        on_customer_selected(self, event)

    def _load_customers(self) -> None:
        """Proxy to customers_tab.load_customers."""
        load_customers(self)

    # ===== INVOICE TAB =====
    def _reset_invoice_form(self) -> None:
        """Reset invoice form (imported from invoice_tab)."""
        from app.ui.tabs.invoice_tab import reset_invoice_form
        reset_invoice_form(self)

    def _save_invoice(self, mode: str) -> None:
        """Save invoice with specified mode."""
        save_invoice(self, mode)

    # ===== CALENDAR TAB =====
    def _open_cleaners_popup(self) -> None:
        """Proxy to dialogs.open_cleaners_popup."""
        open_cleaners_popup(self)

    def _open_google_popup(self) -> None:
        """Proxy to dialogs.open_google_popup."""
        open_google_popup(self)

    def _open_schedule_popup(self) -> None:
        """Proxy to dialogs.open_schedule_popup."""
        open_schedule_popup(self)

    def _load_cleaners(self) -> None:
        """Proxy to dialogs.load_cleaners."""
        load_cleaners(self)

    def _load_jobs(self) -> None:
        """Proxy to dialogs.load_jobs."""
        load_jobs(self)

    def _update_selected_job_status(self, status: str) -> None:
        """Proxy to dialogs.update_selected_job_status."""
        update_selected_job_status(self, status)

    def _delete_selected_job(self) -> None:
        """Proxy to dialogs.delete_selected_job."""
        delete_selected_job(self)

    def _create_invoice_from_job(self) -> None:
        """Proxy to dialogs.create_invoice_from_job."""
        create_invoice_from_job(self)

    def _quick_rebook_job(self) -> None:
        """Proxy to dialogs.quick_rebook_job."""
        quick_rebook_job(self)

    def _check_job_availability(self) -> None:
        """Proxy to dialogs.check_job_availability."""
        check_job_availability(self)

    def _refresh_google_status(self) -> None:
        """Refresh Google Calendar connection status."""
        from app.ui.dialogs import refresh_google_status
        refresh_google_status(self)

    # ===== INVOICES TAB =====
    def _load_invoices(self) -> None:
        """Proxy to invoices_tab.load_invoices."""
        load_invoices(self)

    # ===== INVOICE ACTIONS =====
    def _generate_pdf(self, invoice_id: int) -> str:
        """Generate PDF for an invoice."""
        return generate_pdf(self, invoice_id)

    def _send_invoice(self, invoice_id: int) -> None:
        """Send invoice via email."""
        send_invoice(self, invoice_id)

    def _send_invoice_sms(self, invoice_id: int) -> list[str]:
        """Send invoice via SMS."""
        return send_invoice_sms(self, invoice_id)

    def _refresh_live_views(self) -> None:
        """Refresh the UI sections that reflect live database values."""
        load_customers(self)
        load_cleaners(self)
        load_jobs(self)
        load_invoices(self)
        refresh_dashboard(self)
        self._refresh_google_status()
