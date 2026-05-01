"""Customers tab for managing customer information."""

import tkinter as tk
from tkinter import messagebox, ttk
from datetime import date as date_type

from app.services.pdf_service import format_usd


def build_customers_tab(app) -> None:
    """Build the Customers tab UI."""
    main = ttk.Frame(app.customers_tab)
    main.pack(fill="both", expand=True, padx=8, pady=8)
    main.columnconfigure(0, weight=2)
    main.columnconfigure(1, weight=3)
    main.rowconfigure(0, weight=1)

    form = ttk.LabelFrame(main, text="Customer Details")
    form.grid(row=0, column=0, sticky="nsew", padx=(0, 8), pady=0)
    form.columnconfigure(1, weight=1)

    app.customer_id_var = tk.StringVar(value="")
    app.customer_name_var = tk.StringVar()
    app.customer_email_var = tk.StringVar()
    app.customer_phone_var = tk.StringVar()
    app.customer_address_var = tk.StringVar()
    app.customer_bedrooms_var = tk.StringVar()
    app.customer_bathrooms_var = tk.StringVar()
    app.customer_square_feet_var = tk.StringVar()
    app.customer_cleaning_duration_var = tk.StringVar()

    ttk.Label(form, text="Name").grid(row=0, column=0, sticky="w", padx=6, pady=(8, 4))
    ttk.Entry(form, textvariable=app.customer_name_var).grid(row=0, column=1, sticky="ew", padx=6, pady=(8, 4))

    ttk.Label(form, text="Email").grid(row=1, column=0, sticky="w", padx=6, pady=4)
    ttk.Entry(form, textvariable=app.customer_email_var).grid(row=1, column=1, sticky="ew", padx=6, pady=4)

    ttk.Label(form, text="Phone").grid(row=2, column=0, sticky="w", padx=6, pady=4)
    ttk.Entry(form, textvariable=app.customer_phone_var).grid(row=2, column=1, sticky="ew", padx=6, pady=4)

    job_frame = ttk.LabelFrame(form, text="Job Details")
    job_frame.grid(row=4, column=0, columnspan=2, sticky="ew", padx=6, pady=6)
    job_frame.columnconfigure((1, 3), weight=1)

    ttk.Label(job_frame, text="Beds").grid(row=0, column=0, sticky="w", padx=6, pady=4)
    ttk.Entry(job_frame, textvariable=app.customer_bedrooms_var).grid(row=0, column=1, sticky="ew", padx=6, pady=4)

    ttk.Label(job_frame, text="Baths").grid(row=0, column=2, sticky="w", padx=6, pady=4)
    ttk.Entry(job_frame, textvariable=app.customer_bathrooms_var).grid(row=0, column=3, sticky="ew", padx=6, pady=4)

    ttk.Label(job_frame, text="SqFt").grid(row=1, column=0, sticky="w", padx=6, pady=4)
    ttk.Entry(job_frame, textvariable=app.customer_square_feet_var).grid(row=1, column=1, sticky="ew", padx=6, pady=4)

    ttk.Label(job_frame, text="Cleaning Duration").grid(row=1, column=2, sticky="w", padx=6, pady=4)
    ttk.Entry(job_frame, textvariable=app.customer_cleaning_duration_var).grid(row=1, column=3, sticky="ew", padx=6, pady=4)

    ttk.Label(job_frame, text="Frequency").grid(row=2, column=2, sticky="w", padx=6, pady=4)
    frequency_combo = ttk.Combobox(
        job_frame,
        textvariable=app.customer_frequency_var,
        values=("Single", "Weekly", "Bi-Weekly", "Monthly", "Bi-Monthly"),
        state="readonly",
        width=14,
    )
    frequency_combo.grid(row=2, column=3, sticky="ew", padx=6, pady=4)

    ttk.Label(form, text="Address").grid(row=3, column=0, sticky="w", padx=6, pady=4)
    ttk.Entry(form, textvariable=app.customer_address_var).grid(row=3, column=1, sticky="ew", padx=6, pady=4)

    ttk.Label(form, text="Notes").grid(row=5, column=0, sticky="nw", padx=6, pady=4)
    app.customer_notes_text = tk.Text(form, height=5, width=40)
    app.customer_notes_text.grid(row=5, column=1, sticky="ew", padx=6, pady=4)

    button_row = ttk.Frame(form)
    button_row.grid(row=6, column=0, columnspan=2, sticky="ew", padx=6, pady=(10, 8))
    button_row.columnconfigure((0, 1, 2, 3, 4), weight=1)

    ttk.Button(button_row, text="Add", command=app._add_customer).grid(row=0, column=0, sticky="ew", padx=3)
    ttk.Button(button_row, text="Update", command=app._update_customer).grid(row=0, column=1, sticky="ew", padx=3)
    ttk.Button(button_row, text="Delete", command=app._delete_customer).grid(row=0, column=2, sticky="ew", padx=3)
    ttk.Button(button_row, text="Delete + Invoices", command=app._delete_customer_with_invoices).grid(row=0, column=3, sticky="ew", padx=3)
    ttk.Button(button_row, text="Clear", command=app._clear_customer_form).grid(row=0, column=4, sticky="ew", padx=3)

    table = ttk.LabelFrame(main, text="Customers")
    table.grid(row=0, column=1, sticky="nsew")
    table.columnconfigure(0, weight=1)
    table.rowconfigure(0, weight=1)

    columns = ("id", "name", "email", "phone", "address", "frequency", "last_clean", "balance")
    app.customers_tree = ttk.Treeview(table, columns=columns, show="headings", height=18)
    app.customers_tree.heading("id", text="ID")
    app.customers_tree.heading("name", text="Name")
    app.customers_tree.heading("email", text="Email")
    app.customers_tree.heading("phone", text="Phone")
    app.customers_tree.heading("address", text="Address")
    app.customers_tree.heading("frequency", text="Frequency")
    app.customers_tree.heading("last_clean", text="Last Clean")
    app.customers_tree.heading("balance", text="Unpaid Balance")

    app.customers_tree.column("id", width=45, anchor="center")
    app.customers_tree.column("name", width=160)
    app.customers_tree.column("email", width=200)
    app.customers_tree.column("phone", width=110)
    app.customers_tree.column("address", width=220)
    app.customers_tree.column("frequency", width=90, anchor="center")
    app.customers_tree.column("last_clean", width=100, anchor="center")
    app.customers_tree.column("balance", width=120, anchor="e")

    app.customers_tree.tag_configure("Single", background="#ffffff")
    app.customers_tree.tag_configure("Weekly", background="#c8e6c9")
    app.customers_tree.tag_configure("Bi-Weekly", background="#bbdefb")
    app.customers_tree.tag_configure("Monthly", background="#fff9c4")
    app.customers_tree.tag_configure("Bi-Monthly", background="#e1bee7")

    app.customers_tree.grid(row=0, column=0, sticky="nsew")
    app.customers_tree.bind("<<TreeviewSelect>>", app._on_customer_selected)

    scrollbar = ttk.Scrollbar(table, orient="vertical", command=app.customers_tree.yview)
    app.customers_tree.configure(yscrollcommand=scrollbar.set)
    scrollbar.grid(row=0, column=1, sticky="ns")


def add_customer(app) -> None:
    """Add a new customer."""
    payload = _get_customer_form_payload(app)
    if not payload:
        return

    app.customer_service.create_customer(**payload)
    load_customers(app)
    app._refresh_live_views()
    clear_customer_form(app)
    messagebox.showinfo("Customer", "Customer added.")


def update_customer(app) -> None:
    """Update the selected customer."""
    if not app.customer_id_var.get().strip():
        messagebox.showwarning("Select", "Select a customer first.")
        return

    payload = _get_customer_form_payload(app)
    if not payload:
        return

    app.customer_service.update_customer(int(app.customer_id_var.get()), **payload)
    load_customers(app)
    app._refresh_live_views()
    messagebox.showinfo("Customer", "Customer updated.")


def delete_customer(app) -> None:
    """Delete the selected customer."""
    if not app.customer_id_var.get().strip():
        messagebox.showwarning("Select", "Select a customer first.")
        return

    if not messagebox.askyesno("Confirm", "Delete selected customer?"):
        return

    try:
        app.customer_service.delete_customer(int(app.customer_id_var.get()))
    except ValueError as exc:
        messagebox.showwarning("Delete Blocked", str(exc))
        return
    except Exception as exc:
        messagebox.showerror("Delete Failed", f"Could not delete customer: {exc}")
        return

    load_customers(app)
    app._refresh_live_views()
    clear_customer_form(app)


def delete_customer_with_invoices(app) -> None:
    """Delete customer and all associated invoices."""
    if not app.customer_id_var.get().strip():
        messagebox.showwarning("Select", "Select a customer first.")
        return

    from tkinter import simpledialog
    
    customer_id = int(app.customer_id_var.get())
    customer_name = app.customer_name_var.get().strip() or f"Customer {customer_id}"
    invoice_count = app.customer_service.count_customer_invoices(customer_id)

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
        parent=app,
    )
    if confirmation != "DELETE":
        messagebox.showinfo("Cancelled", "Deletion cancelled.")
        return

    try:
        removed_count = app.customer_service.delete_customer_with_invoices(customer_id)
    except Exception as exc:
        messagebox.showerror("Delete Failed", f"Could not delete customer and invoices: {exc}")
        return

    load_customers(app)
    app._load_invoices()
    app._refresh_live_views()
    clear_customer_form(app)
    messagebox.showinfo(
        "Deleted",
        f"Deleted customer '{customer_name}' and {removed_count} linked invoice(s).",
    )


def clear_customer_form(app) -> None:
    """Clear all customer form fields."""
    app.customer_id_var.set("")
    app.customer_name_var.set("")
    app.customer_email_var.set("")
    app.customer_phone_var.set("")
    app.customer_bedrooms_var.set("")
    app.customer_bathrooms_var.set("")
    app.customer_square_feet_var.set("")
    app.customer_cleaning_duration_var.set("")
    app.customer_address_var.set("")
    app.customer_frequency_var.set("Single")
    app.customer_notes_text.delete("1.0", tk.END)


def on_customer_selected(app, _event) -> None:
    """Populate form when a customer is selected in the tree."""
    selection = app.customers_tree.selection()
    if not selection:
        return

    row_values = app.customers_tree.item(selection[0], "values")
    customer_id = int(row_values[0])
    customer = app.customer_service.get_customer(customer_id)
    if not customer:
        return

    app.customer_id_var.set(str(customer["id"]))
    app.customer_name_var.set(customer["name"])
    app.customer_email_var.set(customer["email"])
    app.customer_phone_var.set(customer["phone"])
    app.customer_bedrooms_var.set(customer["bedrooms"])
    app.customer_bathrooms_var.set(customer["bathrooms"])
    app.customer_square_feet_var.set(customer["square_feet"])
    app.customer_cleaning_duration_var.set(customer["cleaning_duration"])

    app.customer_address_var.set(customer["address"])
    app.customer_frequency_var.set(customer.get("frequency") or "Single")
    app.customer_notes_text.delete("1.0", tk.END)
    app.customer_notes_text.insert("1.0", customer["notes"])


def load_customers(app) -> None:
    """Load all customers and populate the tree, balances, and combo boxes."""
    rows = app.customer_service.list_customers()
    customer_balances = app.invoice_service.get_all_customer_balances()

    for item in app.customers_tree.get_children():
        app.customers_tree.delete(item)

    customer_labels: list[str] = []
    app.customer_lookup.clear()
    app.calendar_customer_lookup.clear()

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

        balance = customer_balances.get(row["id"], 0.0)
        balance_str = format_usd(balance) if balance > 0 else "—"

        app.customers_tree.insert(
            "",
            tk.END,
            values=(row["id"], row["name"], row["email"], row["phone"], one_line_address, frequency, last_clean, balance_str),
            tags=(frequency,),
        )
        label = f"{row['id']} - {row['name']} <{row['email']}>"
        customer_labels.append(label)
        app.customer_lookup[label] = row["id"]
        app.calendar_customer_lookup[label] = row["id"]

    app.invoice_customer_combo["values"] = customer_labels
    if hasattr(app, "job_customer_combo") and app.job_customer_combo.winfo_exists():
        app.job_customer_combo["values"] = customer_labels


def _get_customer_form_payload(app) -> dict | None:
    """Get and validate customer form data."""
    name = app.customer_name_var.get().strip()
    email = app.customer_email_var.get().strip()
    phone = app.customer_phone_var.get().strip()
    bedrooms = app.customer_bedrooms_var.get().strip()
    bathrooms = app.customer_bathrooms_var.get().strip()
    square_feet = app.customer_square_feet_var.get().strip()
    cleaning_duration = app.customer_cleaning_duration_var.get().strip()
    frequency = app.customer_frequency_var.get().strip()
    address = app.customer_address_var.get().strip()
    notes = app.customer_notes_text.get("1.0", tk.END).strip()

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
