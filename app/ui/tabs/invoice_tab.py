"""Invoice creation tab for managing invoice items and generation."""

import tkinter as tk
from datetime import date, timedelta
from tkinter import messagebox, ttk

from app.config import DATE_FORMAT, DEFAULT_DUE_DAYS
from app.services.pdf_service import format_usd
from app.ui.helpers import parse_non_negative_number, valid_date


def build_invoice_tab(app) -> None:
    """Build the Create Invoice tab UI."""
    container = ttk.Frame(app.invoice_tab)
    container.pack(fill="both", expand=True, padx=8, pady=8)
    container.columnconfigure(0, weight=1)
    container.rowconfigure(1, weight=1)

    top = ttk.LabelFrame(container, text="Invoice Header")
    top.grid(row=0, column=0, sticky="ew")
    for i in range(6):
        top.columnconfigure(i, weight=1)

    app.invoice_customer_var = tk.StringVar()
    app.invoice_issue_date_var = tk.StringVar()
    app.invoice_due_date_var = tk.StringVar()
    app.invoice_tax_rate_var = tk.StringVar(value="0")

    ttk.Label(top, text="Customer").grid(row=0, column=0, sticky="w", padx=6, pady=(8, 4))
    app.invoice_customer_combo = ttk.Combobox(
        top,
        textvariable=app.invoice_customer_var,
        state="readonly",
        width=42,
    )
    app.invoice_customer_combo.grid(row=0, column=1, columnspan=2, sticky="ew", padx=6, pady=(8, 4))

    ttk.Label(top, text="Issue Date").grid(row=0, column=3, sticky="w", padx=6, pady=(8, 4))
    ttk.Entry(top, textvariable=app.invoice_issue_date_var).grid(row=0, column=4, sticky="ew", padx=6, pady=(8, 4))

    ttk.Label(top, text="Due Date").grid(row=0, column=5, sticky="w", padx=6, pady=(8, 4))
    ttk.Entry(top, textvariable=app.invoice_due_date_var).grid(row=0, column=5, sticky="e", padx=(80, 6), pady=(8, 4))

    ttk.Label(top, text="Tax %").grid(row=1, column=0, sticky="w", padx=6, pady=(4, 8))
    tax_entry = ttk.Entry(top, textvariable=app.invoice_tax_rate_var, width=10)
    tax_entry.grid(row=1, column=1, sticky="w", padx=6, pady=(4, 8))
    tax_entry.bind("<FocusOut>", lambda _event: refresh_invoice_item_table(app))

    ttk.Label(top, text="Notes").grid(row=1, column=2, sticky="w", padx=6, pady=(4, 8))
    app.invoice_notes_text = tk.Text(top, height=3, width=60)
    app.invoice_notes_text.grid(row=1, column=3, columnspan=3, sticky="ew", padx=6, pady=(4, 8))

    middle = ttk.Frame(container)
    middle.grid(row=1, column=0, sticky="nsew", pady=(10, 0))
    middle.columnconfigure(0, weight=1)
    middle.rowconfigure(1, weight=1)

    item_form = ttk.LabelFrame(middle, text="Add Service Item")
    item_form.grid(row=0, column=0, sticky="ew")
    for i in range(5):
        item_form.columnconfigure(i, weight=1)

    app.item_description_var = tk.StringVar()
    app.item_rate_var = tk.StringVar(value="0")

    ttk.Label(item_form, text="Description").grid(row=0, column=0, sticky="w", padx=6, pady=6)
    ttk.Entry(item_form, textvariable=app.item_description_var).grid(row=0, column=1, sticky="ew", padx=6, pady=6)

    ttk.Label(item_form, text="Rate (USD)").grid(row=0, column=2, sticky="w", padx=6, pady=6)
    ttk.Entry(item_form, textvariable=app.item_rate_var, width=12).grid(row=0, column=3, sticky="w", padx=6, pady=6)

    ttk.Button(item_form, text="Add Item", command=lambda: add_invoice_item(app)).grid(row=0, column=4, sticky="ew", padx=6, pady=6)

    items_frame = ttk.LabelFrame(middle, text="Invoice Items")
    items_frame.grid(row=1, column=0, sticky="nsew", pady=(10, 0))
    items_frame.columnconfigure(0, weight=1)
    items_frame.rowconfigure(0, weight=1)

    app.items_tree = ttk.Treeview(
        items_frame,
        columns=("description", "unit_price", "line_total"),
        show="headings",
        height=12,
    )
    app.items_tree.heading("description", text="Description")
    app.items_tree.heading("unit_price", text="Unit Price")
    app.items_tree.heading("line_total", text="Line Total")

    app.items_tree.column("description", width=620)
    app.items_tree.column("unit_price", width=160, anchor="e")
    app.items_tree.column("line_total", width=160, anchor="e")
    app.items_tree.grid(row=0, column=0, sticky="nsew")

    scroll = ttk.Scrollbar(items_frame, orient="vertical", command=app.items_tree.yview)
    app.items_tree.configure(yscrollcommand=scroll.set)
    scroll.grid(row=0, column=1, sticky="ns")

    action_row = ttk.Frame(items_frame)
    action_row.grid(row=1, column=0, columnspan=2, sticky="ew", padx=4, pady=8)
    action_row.columnconfigure((0, 1, 2, 3, 4, 5), weight=1)

    ttk.Button(action_row, text="Remove Selected Item", command=lambda: remove_selected_item(app)).grid(row=0, column=0, sticky="ew", padx=4)
    ttk.Button(action_row, text="Clear Items", command=lambda: clear_items(app)).grid(row=0, column=1, sticky="ew", padx=4)

    app.subtotal_label = ttk.Label(action_row, text="Subtotal: $0.00")
    app.subtotal_label.grid(row=0, column=2, sticky="e", padx=4)
    app.tax_label = ttk.Label(action_row, text="Tax: $0.00")
    app.tax_label.grid(row=0, column=3, sticky="e", padx=4)
    app.total_label = ttk.Label(action_row, text="Total: $0.00")
    app.total_label.grid(row=0, column=4, sticky="e", padx=4)

    ttk.Button(action_row, text="Reset Form", command=lambda: reset_invoice_form(app)).grid(row=0, column=5, sticky="ew", padx=4)

    bottom = ttk.Frame(container)
    bottom.grid(row=2, column=0, sticky="ew", pady=(12, 0))
    bottom.columnconfigure((0, 1, 2, 3), weight=1)

    ttk.Button(bottom, text="Save Draft", command=lambda: app._save_invoice("draft")).grid(row=0, column=0, sticky="ew", padx=4)
    ttk.Button(bottom, text="Save + Generate PDF", command=lambda: app._save_invoice("pdf")).grid(row=0, column=1, sticky="ew", padx=4)
    ttk.Button(bottom, text="Save + PDF + Email", command=lambda: app._save_invoice("email")).grid(row=0, column=2, sticky="ew", padx=4)
    ttk.Button(bottom, text="Save + Text", command=lambda: app._save_invoice("sms")).grid(row=0, column=3, sticky="ew", padx=4)


def add_invoice_item(app) -> None:
    """Add a service item to the invoice."""
    description = app.item_description_var.get().strip()
    if not description:
        messagebox.showwarning("Validation", "Service description is required.")
        return

    rate = parse_non_negative_number(app.item_rate_var.get(), "Rate")
    if rate is None:
        return

    app.current_items.append(
        {
            "description": description,
            "quantity": 1.0,
            "unit_price": rate,
        }
    )

    app.item_description_var.set("")
    app.item_rate_var.set("0")
    refresh_invoice_item_table(app)


def remove_selected_item(app) -> None:
    """Remove selected item from the invoice."""
    selection = app.items_tree.selection()
    if not selection:
        messagebox.showwarning("Select", "Select an item row first.")
        return

    index = app.items_tree.index(selection[0])
    if 0 <= index < len(app.current_items):
        del app.current_items[index]
    refresh_invoice_item_table(app)


def clear_items(app) -> None:
    """Clear all invoice items."""
    app.current_items = []
    refresh_invoice_item_table(app)


def refresh_invoice_item_table(app) -> None:
    """Refresh the invoice items table and recalculate totals."""
    for item in app.items_tree.get_children():
        app.items_tree.delete(item)

    for item in app.current_items:
        line_total = float(item["quantity"]) * float(item["unit_price"])
        app.items_tree.insert(
            "",
            tk.END,
            values=(
                item["description"],
                format_usd(float(item["unit_price"])),
                format_usd(line_total),
            ),
        )

    tax_rate = parse_non_negative_number(app.invoice_tax_rate_var.get(), "Tax")
    tax_rate = tax_rate if tax_rate is not None else 0

    totals = app.invoice_service.calculate_totals(app.current_items, tax_rate)
    app.subtotal_label.configure(text=f"Subtotal: {format_usd(float(totals.subtotal))}")
    app.tax_label.configure(text=f"Tax: {format_usd(float(totals.tax_amount))}")
    app.total_label.configure(text=f"Total: {format_usd(float(totals.total_amount))}")


def reset_invoice_form(app) -> None:
    """Reset invoice form to defaults."""
    today = date.today()
    app.invoice_issue_date_var.set(today.strftime(DATE_FORMAT))
    app.invoice_due_date_var.set((today + timedelta(days=DEFAULT_DUE_DAYS)).strftime(DATE_FORMAT))

    settings = app.settings_service.get_settings()
    app.invoice_tax_rate_var.set(settings.get("default_tax_rate", "0"))

    app.invoice_customer_var.set("")
    app.invoice_notes_text.delete("1.0", tk.END)

    app.item_description_var.set("")
    app.item_rate_var.set("0")

    app.current_items = []
    refresh_invoice_item_table(app)


def save_invoice(app, mode: str) -> None:
    """Save invoice and optionally generate PDF and/or send email/SMS."""
    customer_label = app.invoice_customer_var.get().strip()
    if customer_label not in app.customer_lookup:
        messagebox.showwarning("Validation", "Select a customer.")
        return

    if not app.current_items:
        messagebox.showwarning("Validation", "Add at least one service item.")
        return

    issue_date = app.invoice_issue_date_var.get().strip()
    due_date = app.invoice_due_date_var.get().strip()

    if not valid_date(issue_date):
        messagebox.showwarning("Validation", f"Issue date must match {DATE_FORMAT}.")
        return

    if not valid_date(due_date):
        messagebox.showwarning("Validation", f"Due date must match {DATE_FORMAT}.")
        return

    tax_rate = parse_non_negative_number(app.invoice_tax_rate_var.get(), "Tax")
    if tax_rate is None:
        return

    notes = app.invoice_notes_text.get("1.0", tk.END).strip()

    settings = app.settings_service.get_settings()
    numbering_scheme = settings.get("invoice_numbering_scheme", "by_year")

    saved = app.invoice_service.create_invoice(
        customer_id=app.customer_lookup[customer_label],
        issue_date=issue_date,
        due_date=due_date,
        tax_rate=tax_rate,
        notes=notes,
        items=app.current_items,
        status="draft",
        numbering_scheme=numbering_scheme,
    )

    invoice_id = int(saved["invoice_id"])
    invoice_number = saved["invoice_number"]
    sms_delivery_gateways: list[str] = []

    if mode in {"pdf", "email"}:
        app._generate_pdf(invoice_id)

    if mode == "email":
        app._send_invoice(invoice_id)
    elif mode == "sms":
        sms_delivery_gateways = app._send_invoice_sms(invoice_id)

    app._load_invoices()
    from app.ui.tabs.customers_tab import load_customers
    load_customers(app)
    app._refresh_live_views()
    reset_invoice_form(app)

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
