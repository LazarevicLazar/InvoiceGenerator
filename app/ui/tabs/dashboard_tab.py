"""Dashboard tab for revenue and business metrics."""

import tkinter as tk
from datetime import date as date_type
from tkinter import ttk

from app.services.pdf_service import format_usd


def build_dashboard_tab(app) -> None:
    """Build the Dashboard tab UI."""
    container = ttk.Frame(app.dashboard_tab)
    container.pack(fill="both", expand=True, padx=8, pady=8)
    container.columnconfigure(0, weight=1)
    container.rowconfigure(2, weight=1)

    summary_frame = ttk.LabelFrame(container, text="Summary Metrics")
    summary_frame.grid(row=0, column=0, sticky="ew", pady=(0, 10))
    summary_frame.columnconfigure((0, 1, 2, 3), weight=1)

    app.dashboard_this_month_var = tk.StringVar(value="$0.00")
    app.dashboard_this_year_var = tk.StringVar(value="$0.00")
    app.dashboard_unpaid_var = tk.StringVar(value="$0.00")

    ttk.Label(summary_frame, text="This Month:", font=("TkDefaultFont", 10, "bold")).grid(row=0, column=0, sticky="w", padx=6, pady=8)
    ttk.Label(summary_frame, textvariable=app.dashboard_this_month_var, font=("TkDefaultFont", 12, "bold"), foreground="green").grid(row=0, column=1, sticky="w", padx=6, pady=8)

    ttk.Label(summary_frame, text="This Year:", font=("TkDefaultFont", 10, "bold")).grid(row=0, column=2, sticky="w", padx=6, pady=8)
    ttk.Label(summary_frame, textvariable=app.dashboard_this_year_var, font=("TkDefaultFont", 12, "bold"), foreground="green").grid(row=0, column=3, sticky="w", padx=6, pady=8)

    ttk.Label(summary_frame, text="Unpaid Total:", font=("TkDefaultFont", 10, "bold")).grid(row=1, column=0, sticky="w", padx=6, pady=8)
    ttk.Label(summary_frame, textvariable=app.dashboard_unpaid_var, font=("TkDefaultFont", 12, "bold"), foreground="red").grid(row=1, column=1, sticky="w", padx=6, pady=8)

    ttk.Button(summary_frame, text="Refresh Dashboard", command=lambda: refresh_dashboard(app)).grid(row=1, column=3, sticky="e", padx=6, pady=8)

    breakdown_frame = ttk.Frame(container)
    breakdown_frame.grid(row=2, column=0, sticky="nsew")
    breakdown_frame.columnconfigure((0, 1), weight=1)
    breakdown_frame.rowconfigure(0, weight=1)

    customer_frame = ttk.LabelFrame(breakdown_frame, text="Revenue by Customer (This Year)")
    customer_frame.grid(row=0, column=0, sticky="nsew", padx=(0, 4))
    customer_frame.columnconfigure(0, weight=1)
    customer_frame.rowconfigure(0, weight=1)

    app.dashboard_customer_tree = ttk.Treeview(customer_frame, columns=("name", "revenue"), show="headings", height=15)
    app.dashboard_customer_tree.heading("name", text="Customer")
    app.dashboard_customer_tree.heading("revenue", text="Revenue")
    app.dashboard_customer_tree.column("name", width=200)
    app.dashboard_customer_tree.column("revenue", width=100, anchor="e")
    app.dashboard_customer_tree.grid(row=0, column=0, sticky="nsew")

    customer_scroll = ttk.Scrollbar(customer_frame, orient="vertical", command=app.dashboard_customer_tree.yview)
    app.dashboard_customer_tree.configure(yscrollcommand=customer_scroll.set)
    customer_scroll.grid(row=0, column=1, sticky="ns")

    cleaner_frame = ttk.LabelFrame(breakdown_frame, text="Revenue by Cleaner (This Year)")
    cleaner_frame.grid(row=0, column=1, sticky="nsew", padx=(4, 0))
    cleaner_frame.columnconfigure(0, weight=1)
    cleaner_frame.rowconfigure(0, weight=1)

    app.dashboard_cleaner_tree = ttk.Treeview(cleaner_frame, columns=("name", "revenue"), show="headings", height=15)
    app.dashboard_cleaner_tree.heading("name", text="Cleaner")
    app.dashboard_cleaner_tree.heading("revenue", text="Revenue")
    app.dashboard_cleaner_tree.column("name", width=200)
    app.dashboard_cleaner_tree.column("revenue", width=100, anchor="e")
    app.dashboard_cleaner_tree.grid(row=0, column=0, sticky="nsew")

    cleaner_scroll = ttk.Scrollbar(cleaner_frame, orient="vertical", command=app.dashboard_cleaner_tree.yview)
    app.dashboard_cleaner_tree.configure(yscrollcommand=cleaner_scroll.set)
    cleaner_scroll.grid(row=0, column=1, sticky="ns")


def refresh_dashboard(app) -> None:
    """Refresh all dashboard metrics and breakdowns."""
    today = date_type.today()
    current_month = today.month
    current_year = today.year

    month_revenue = app.invoice_service.get_revenue_by_period(current_year, current_month)
    year_revenue = app.invoice_service.get_revenue_by_period(current_year)
    unpaid_total = app.invoice_service.get_total_unpaid()

    app.dashboard_this_month_var.set(format_usd(month_revenue))
    app.dashboard_this_year_var.set(format_usd(year_revenue))
    app.dashboard_unpaid_var.set(format_usd(unpaid_total))

    customer_breakdown = app.invoice_service.get_revenue_by_customer(current_year)
    cleaner_breakdown = app.invoice_service.get_revenue_by_cleaner(current_year)

    for item in app.dashboard_customer_tree.get_children():
        app.dashboard_customer_tree.delete(item)
    for customer in customer_breakdown:
        app.dashboard_customer_tree.insert(
            "",
            tk.END,
            values=(customer["name"], format_usd(float(customer["revenue"]))),
        )

    for item in app.dashboard_cleaner_tree.get_children():
        app.dashboard_cleaner_tree.delete(item)
    for cleaner in cleaner_breakdown:
        app.dashboard_cleaner_tree.insert(
            "",
            tk.END,
            values=(cleaner["name"], format_usd(float(cleaner["revenue"]))),
        )
