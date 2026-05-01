"""Settings tab for application configuration."""

import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import re


def build_settings_tab(app) -> None:
    """Build the Settings tab UI."""
    container = ttk.Frame(app.settings_tab)
    container.pack(fill="both", expand=True, padx=8, pady=8)
    container.columnconfigure(1, weight=1)

    app.settings_vars: dict[str, tk.StringVar] = {
        "business_name": tk.StringVar(),
        "business_email": tk.StringVar(),
        "business_phone": tk.StringVar(),
        "business_logo_path": tk.StringVar(),
        "default_tax_rate": tk.StringVar(value="0"),
        "invoice_numbering_scheme": tk.StringVar(value="by_year"),
        "smtp_server": tk.StringVar(value="smtp.gmail.com"),
        "smtp_port": tk.StringVar(value="587"),
        "smtp_username": tk.StringVar(),
        "smtp_password": tk.StringVar(),
        "smtp_from_email": tk.StringVar(),
        "google_credentials_file": tk.StringVar(),
        "google_token_file": tk.StringVar(),
    }
    app.smtp_use_tls_var = tk.IntVar(value=1)

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
        ttk.Entry(container, textvariable=app.settings_vars[key], show=show).grid(
            row=row,
            column=1,
            sticky="ew",
            padx=6,
            pady=4,
        )
        row += 1

    ttk.Label(container, text="Invoice Numbering Scheme").grid(row=row, column=0, sticky="w", padx=6, pady=4)
    scheme_combo = ttk.Combobox(
        container,
        textvariable=app.settings_vars["invoice_numbering_scheme"],
        values=("by_year", "by_customer"),
        state="readonly",
        width=20,
    )
    scheme_combo.grid(row=row, column=1, sticky="w", padx=6, pady=4)
    row += 1

    ttk.Label(container, text="Business Logo Path").grid(row=row, column=0, sticky="w", padx=6, pady=4)
    logo_wrap = ttk.Frame(container)
    logo_wrap.grid(row=row, column=1, sticky="ew", padx=6, pady=4)
    logo_wrap.columnconfigure(0, weight=1)
    ttk.Entry(logo_wrap, textvariable=app.settings_vars["business_logo_path"]).grid(row=0, column=0, sticky="ew")
    ttk.Button(logo_wrap, text="Browse", command=lambda: browse_logo_path(app)).grid(row=0, column=1, padx=(6, 0))
    row += 1

    ttk.Label(container, text="Business Address").grid(row=row, column=0, sticky="nw", padx=6, pady=4)
    app.business_address_text = tk.Text(container, height=4, width=50)
    app.business_address_text.grid(row=row, column=1, sticky="ew", padx=6, pady=4)
    row += 1

    ttk.Label(container, text="Payment Instructions").grid(row=row, column=0, sticky="nw", padx=6, pady=4)
    app.payment_instructions_text = tk.Text(container, height=4, width=50)
    app.payment_instructions_text.grid(row=row, column=1, sticky="ew", padx=6, pady=4)
    row += 1

    ttk.Label(container, text="Invoice Email Template").grid(row=row, column=0, sticky="nw", padx=6, pady=4)
    app.email_template_text = tk.Text(container, height=6, width=50)
    app.email_template_text.grid(row=row, column=1, sticky="ew", padx=6, pady=4)
    row += 1

    ttk.Checkbutton(
        container,
        text="Use TLS for SMTP",
        variable=app.smtp_use_tls_var,
        onvalue=1,
        offvalue=0,
    ).grid(row=row, column=1, sticky="w", padx=6, pady=4)
    row += 1

    actions = ttk.Frame(container)
    actions.grid(row=row, column=0, columnspan=2, sticky="ew", padx=6, pady=(10, 0))
    actions.columnconfigure((0, 1), weight=1)

    ttk.Button(actions, text="Save Settings", command=lambda: save_settings(app)).grid(row=0, column=0, sticky="ew", padx=4)
    ttk.Button(actions, text="Reload", command=lambda: load_settings_into_form(app)).grid(row=0, column=1, sticky="ew", padx=4)


def load_settings_into_form(app) -> None:
    """Load settings from database into form."""
    settings = app.settings_service.get_settings()
    for key, var in app.settings_vars.items():
        var.set(settings.get(key, ""))

    app.smtp_use_tls_var.set(1 if settings.get("smtp_use_tls", "1") == "1" else 0)

    app.business_address_text.delete("1.0", tk.END)
    app.business_address_text.insert("1.0", settings.get("business_address", ""))

    app.payment_instructions_text.delete("1.0", tk.END)
    app.payment_instructions_text.insert("1.0", settings.get("payment_instructions", ""))

    app.email_template_text.delete("1.0", tk.END)
    app.email_template_text.insert("1.0", settings.get("email_invoice_template", ""))

    if not app.invoice_tax_rate_var.get().strip():
        app.invoice_tax_rate_var.set(settings.get("default_tax_rate", "0"))


def _validate_email(email: str) -> bool:
    """Validate email format using regex."""
    pattern = r'^[^\s@]+@[^\s@]+\.[^\s@]+$'
    return re.match(pattern, email) is not None


def save_settings(app) -> None:
    """Save settings to database."""
    # Validate SMTP port
    port_str = app.settings_vars["smtp_port"].get().strip()
    if port_str:
        try:
            port = int(port_str)
            if port < 1 or port > 65535:
                messagebox.showwarning("Validation", "SMTP port must be between 1 and 65535.")
                return
        except ValueError:
            messagebox.showwarning("Validation", "SMTP port must be a valid number.")
            return
    
    # Validate email fields
    business_email = app.settings_vars["business_email"].get().strip()
    smtp_from_email = app.settings_vars["smtp_from_email"].get().strip()
    
    if business_email and not _validate_email(business_email):
        messagebox.showwarning("Validation", "Business email format is invalid. Use format: user@example.com")
        return
    
    if smtp_from_email and not _validate_email(smtp_from_email):
        messagebox.showwarning("Validation", "SMTP from email format is invalid. Use format: user@example.com")
        return
    
    updates = {key: var.get().strip() for key, var in app.settings_vars.items()}
    updates["business_address"] = app.business_address_text.get("1.0", tk.END).strip()
    updates["payment_instructions"] = app.payment_instructions_text.get("1.0", tk.END).strip()
    updates["email_invoice_template"] = app.email_template_text.get("1.0", tk.END).strip()
    updates["smtp_use_tls"] = "1" if app.smtp_use_tls_var.get() else "0"

    app.settings_service.save_settings(updates)
    app.google_calendar_service = None
    app.google_connected = False
    app._refresh_google_status()
    messagebox.showinfo("Saved", "Settings saved.")


def browse_logo_path(app) -> None:
    """Open file dialog to browse for logo image."""
    filename = filedialog.askopenfilename(
        title="Select logo image",
        filetypes=[("Image files", "*.png *.jpg *.jpeg *.gif")],
    )
    if filename:
        app.settings_vars["business_logo_path"].set(filename)
