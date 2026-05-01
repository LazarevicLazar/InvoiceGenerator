"""Invoice history and management tab."""

import smtplib
import tkinter as tk
from datetime import date as date_type
from email.message import EmailMessage
from pathlib import Path
from tkinter import messagebox, ttk

from app.services.pdf_service import format_usd
from app.ui.helpers import open_file


def build_invoices_tab(app) -> None:
    """Build the Invoice History tab UI."""
    container = ttk.Frame(app.invoices_tab)
    container.pack(fill="both", expand=True, padx=8, pady=8)
    container.columnconfigure(0, weight=1)
    container.rowconfigure(1, weight=1)

    banner_frame = ttk.Frame(container)
    banner_frame.grid(row=0, column=0, sticky="ew", pady=(0, 8))
    app.overdue_banner_label = ttk.Label(banner_frame, text="", foreground="red", font=("TkDefaultFont", 10, "bold"))
    app.overdue_banner_label.pack(side="left")

    app.invoices_tree = ttk.Treeview(
        container,
        columns=("id", "invoice_number", "customer", "issue", "due", "total", "status", "sent_at"),
        show="headings",
        height=22,
    )
    app.invoices_tree.heading("id", text="ID")
    app.invoices_tree.heading("invoice_number", text="Invoice #")
    app.invoices_tree.heading("customer", text="Customer")
    app.invoices_tree.heading("issue", text="Issue Date")
    app.invoices_tree.heading("due", text="Due Date")
    app.invoices_tree.heading("total", text="Total")
    app.invoices_tree.heading("status", text="Status")
    app.invoices_tree.heading("sent_at", text="Sent At")

    app.invoices_tree.column("id", width=60, anchor="center")
    app.invoices_tree.column("invoice_number", width=140)
    app.invoices_tree.column("customer", width=220)
    app.invoices_tree.column("issue", width=110)
    app.invoices_tree.column("due", width=110)
    app.invoices_tree.column("total", width=120, anchor="e")
    app.invoices_tree.column("status", width=100, anchor="center")
    app.invoices_tree.column("sent_at", width=170)

    app.invoices_tree.tag_configure("overdue", foreground="#d00000", background="#ffe0e0")

    app.invoices_tree.grid(row=1, column=0, sticky="nsew")
    scroll = ttk.Scrollbar(container, orient="vertical", command=app.invoices_tree.yview)
    app.invoices_tree.configure(yscrollcommand=scroll.set)
    scroll.grid(row=1, column=1, sticky="ns")

    actions = ttk.Frame(container)
    actions.grid(row=2, column=0, columnspan=2, sticky="ew", pady=(10, 0))
    actions.columnconfigure((0, 1, 2, 3, 4, 5, 6, 7), weight=1)

    ttk.Button(actions, text="Refresh", command=app._load_invoices).grid(row=0, column=0, sticky="ew", padx=4)
    ttk.Button(actions, text="Generate PDF", command=lambda: generate_pdf_for_selected(app)).grid(row=0, column=1, sticky="ew", padx=4)
    ttk.Button(actions, text="Send Email", command=lambda: send_selected_invoice(app)).grid(row=0, column=2, sticky="ew", padx=4)
    ttk.Button(actions, text="Mark Paid", command=lambda: mark_selected_paid(app)).grid(row=0, column=3, sticky="ew", padx=4)
    ttk.Button(actions, text="Open PDF", command=lambda: open_selected_pdf(app)).grid(row=0, column=4, sticky="ew", padx=4)
    ttk.Button(actions, text="Delete Invoice", command=lambda: delete_selected_invoice(app)).grid(row=0, column=5, sticky="ew", padx=4)
    ttk.Button(actions, text="Send Text", command=lambda: send_selected_sms(app)).grid(row=0, column=6, sticky="ew", padx=4)
    ttk.Button(actions, text="Send Overdue Reminder", command=lambda: send_overdue_reminder(app)).grid(row=0, column=7, sticky="ew", padx=4)


def load_invoices(app) -> None:
    """Load all invoices and update the tree."""
    rows = app.invoice_service.list_invoices()
    overdue_invoices = app.invoice_service.get_overdue_invoices()
    overdue_ids = {inv["id"] for inv in overdue_invoices}

    for item in app.invoices_tree.get_children():
        app.invoices_tree.delete(item)

    today = date_type.today().isoformat()
    for row in rows:
        sent_at = row["sent_at"] or ""
        is_overdue = row["id"] in overdue_ids
        tags = ("overdue",) if is_overdue else ()
        
        app.invoices_tree.insert(
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
            tags=tags,
        )

    overdue_count = len(overdue_invoices)
    if overdue_count > 0:
        app.overdue_banner_label.config(text=f"⚠️  {overdue_count} overdue invoice(s) — amounts are unpaid and past due date")
    else:
        app.overdue_banner_label.config(text="")


def _selected_invoice_id(app) -> int | None:
    """Get the selected invoice ID from the tree."""
    selected = app.invoices_tree.selection()
    if not selected:
        messagebox.showwarning("Select", "Select an invoice first.")
        return None

    values = app.invoices_tree.item(selected[0], "values")
    return int(values[0])


def generate_pdf_for_selected(app) -> None:
    """Generate PDF for selected invoice."""
    invoice_id = _selected_invoice_id(app)
    if invoice_id is None:
        return

    try:
        pdf_path = app._generate_pdf(invoice_id)
    except Exception as exc:
        messagebox.showerror("PDF Error", str(exc))
        return

    load_invoices(app)
    messagebox.showinfo("PDF", f"PDF generated:\n{pdf_path}")


def send_selected_invoice(app) -> None:
    """Send selected invoice via email."""
    invoice_id = _selected_invoice_id(app)
    if invoice_id is None:
        return

    try:
        app._send_invoice(invoice_id)
    except Exception as exc:
        messagebox.showerror("Email Error", str(exc))
        return

    load_invoices(app)
    messagebox.showinfo("Email", "Invoice email sent.")


def send_selected_sms(app) -> None:
    """Send selected invoice via SMS."""
    invoice_id = _selected_invoice_id(app)
    if invoice_id is None:
        return

    try:
        delivered_gateways = app._send_invoice_sms(invoice_id)
    except Exception as exc:
        messagebox.showerror("SMS Error", str(exc))
        return

    load_invoices(app)
    messagebox.showinfo(
        "Text",
        (
            "Invoice text sent through carrier email gateways "
            f"(accepted by {len(delivered_gateways)} gateway(s))."
        ),
    )


def send_overdue_reminder(app) -> None:
    """Send reminder emails for all overdue invoices."""
    overdue_invoices = app.invoice_service.get_overdue_invoices()
    if not overdue_invoices:
        messagebox.showinfo("Overdue Reminders", "No overdue invoices to send reminders for.")
        return

    if not messagebox.askyesno(
        "Send Overdue Reminders",
        f"Send reminder emails for {len(overdue_invoices)} overdue invoice(s)?"
    ):
        return

    settings = app.settings_service.get_settings()
    success_count = 0
    failed_count = 0

    for invoice in overdue_invoices:
        try:
            details = app.invoice_service.get_invoice_details(invoice["id"])
            if not details:
                failed_count += 1
                continue

            customer_email = (details.get("customer_email") or "").strip()
            if not customer_email:
                failed_count += 1
                continue

            invoice_number = details["invoice_number"]
            customer_name = details.get("customer_name", "Customer")
            total_amount = float(details["total_amount"])
            due_date = details["due_date"]

            email_template = settings.get("email_invoice_template", "").strip()
            if not email_template:
                email_template = (
                    "Hi {customer_name},\n\n"
                    "Please find attached invoice {invoice_number}.\n"
                    "Total Due: ${total_amount:,.2f}\n"
                    "Due Date: {due_date}\n\n"
                    "Thank you for choosing {business_name}.\n"
                )

            business_name = settings.get("business_name", "Your Cleaning Business").strip() or "Your Cleaning Business"
            reminder_body = email_template.format(
                customer_name=customer_name,
                invoice_number=invoice_number,
                total_amount=total_amount,
                due_date=due_date,
                business_name=business_name,
            )
            reminder_body = "REMINDER: This invoice is now OVERDUE.\n\n" + reminder_body

            smtp_server = settings.get("smtp_server", "").strip()
            smtp_port = int(settings.get("smtp_port", "587") or "587")
            smtp_username = settings.get("smtp_username", "").strip()
            smtp_password = settings.get("smtp_password", "")
            from_email = settings.get("smtp_from_email", "").strip() or settings.get("business_email", "").strip()
            use_tls = settings.get("smtp_use_tls", "1") == "1"

            if not smtp_server or not from_email:
                failed_count += 1
                continue

            message = EmailMessage()
            message["From"] = from_email
            message["To"] = customer_email
            message["Subject"] = f"REMINDER: Invoice {invoice_number} is Overdue"
            message.set_content(reminder_body)

            with smtplib.SMTP(smtp_server, smtp_port, timeout=30) as smtp:
                if use_tls:
                    smtp.starttls()
                if smtp_username:
                    smtp.login(smtp_username, smtp_password)
                smtp.send_message(message)

            success_count += 1
        except Exception as exc:
            print(f"Failed to send reminder for invoice {invoice.get('invoice_number')}: {exc}")
            failed_count += 1

    messagebox.showinfo(
        "Overdue Reminders Sent",
        f"Successfully sent {success_count} reminder(s).\n"
        f"Failed: {failed_count}."
    )
    load_invoices(app)


def mark_selected_paid(app) -> None:
    """Mark selected invoice as paid."""
    invoice_id = _selected_invoice_id(app)
    if invoice_id is None:
        return

    app.invoice_service.mark_paid(invoice_id)
    load_invoices(app)
    messagebox.showinfo("Invoice", "Invoice marked as paid.")


def delete_selected_invoice(app) -> None:
    """Delete selected invoice."""
    selected = app.invoices_tree.selection()
    if not selected:
        messagebox.showwarning("Select", "Select an invoice first.")
        return

    values = app.invoices_tree.item(selected[0], "values")
    invoice_id = int(values[0])
    invoice_number = values[1]

    if not messagebox.askyesno(
        "Confirm",
        f"Delete invoice {invoice_number}? This cannot be undone.",
    ):
        return

    app.invoice_service.delete_invoice(invoice_id)
    load_invoices(app)
    messagebox.showinfo("Invoice", f"Invoice {invoice_number} deleted.")


def open_selected_pdf(app) -> None:
    """Open PDF for selected invoice."""
    invoice_id = _selected_invoice_id(app)
    if invoice_id is None:
        return

    details = app.invoice_service.get_invoice_details(invoice_id)
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

    open_file(pdf_file)
