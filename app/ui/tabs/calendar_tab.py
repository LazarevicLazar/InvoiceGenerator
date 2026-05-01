"""Calendar and job scheduling tab."""

import tkinter as tk
from datetime import date, datetime, timedelta
from tkinter import messagebox, simpledialog, ttk

from app.ui.constants import CALENDAR_DATETIME_FORMAT


def build_calendar_tab(app) -> None:
    """Build the Calendar/Jobs tab UI."""
    app.cleaner_id_var = tk.StringVar(value="")
    app.cleaner_name_var = tk.StringVar()
    app.cleaner_phone_var = tk.StringVar()
    app.cleaner_email_var = tk.StringVar()
    app.cleaner_google_calendar_id_var = tk.StringVar()
    app.auto_create_cleaner_calendar_var = tk.IntVar(value=1)
    app.google_status_var = tk.StringVar(value="Status: Not connected")
    app.google_calendars_combo_var = tk.StringVar()
    app.google_target_cleaner_var = tk.StringVar()

    app.job_customer_var = tk.StringVar()
    app.job_cleaner_var = tk.StringVar()
    app.job_title_var = tk.StringVar()
    app.job_start_var = tk.StringVar()
    app.job_end_var = tk.StringVar()
    app.job_location_var = tk.StringVar()

    app.calendar_view_var = tk.StringVar(value="Week")
    app.calendar_range_var = tk.StringVar(value="")

    container = ttk.Frame(app.calendar_tab)
    container.pack(fill="both", expand=True, padx=8, pady=8)
    container.columnconfigure(0, weight=1)
    container.rowconfigure(2, weight=1)

    top_actions = ttk.Frame(container)
    top_actions.grid(row=0, column=0, sticky="ew", pady=(0, 8))
    top_actions.columnconfigure((0, 1, 2, 3), weight=1)
    ttk.Button(top_actions, text="Manage Cleaners", command=app._open_cleaners_popup).grid(row=0, column=0, sticky="ew", padx=4)
    ttk.Button(top_actions, text="Google Calendar", command=app._open_google_popup).grid(row=0, column=1, sticky="ew", padx=4)
    ttk.Button(top_actions, text="Schedule Job", command=app._open_schedule_popup).grid(row=0, column=2, sticky="ew", padx=4)
    ttk.Button(top_actions, text="Refresh", command=app._load_jobs).grid(row=0, column=3, sticky="ew", padx=4)

    nav = ttk.Frame(container)
    nav.grid(row=1, column=0, sticky="ew", pady=(0, 8))
    nav.columnconfigure(4, weight=1)
    ttk.Button(nav, text="<", width=4, command=lambda: move_calendar_range(app, -1)).grid(row=0, column=0, padx=(0, 4))
    ttk.Button(nav, text="Today", command=lambda: set_calendar_today(app)).grid(row=0, column=1, padx=4)
    ttk.Button(nav, text=">", width=4, command=lambda: move_calendar_range(app, 1)).grid(row=0, column=2, padx=4)

    view_combo = ttk.Combobox(
        nav,
        textvariable=app.calendar_view_var,
        values=("Day", "Week", "Month", "Year"),
        state="readonly",
        width=10,
    )
    view_combo.grid(row=0, column=3, padx=(10, 8))
    view_combo.bind("<<ComboboxSelected>>", lambda e: on_calendar_view_changed(app, e))
    ttk.Label(nav, textvariable=app.calendar_range_var).grid(row=0, column=4, sticky="w")

    jobs_box = ttk.LabelFrame(container, text="Scheduled Jobs")
    jobs_box.grid(row=2, column=0, sticky="nsew")
    jobs_box.columnconfigure(0, weight=1)
    jobs_box.rowconfigure(0, weight=1)

    app.jobs_tree = ttk.Treeview(jobs_box, show="headings", height=18)
    app.jobs_tree.grid(row=0, column=0, sticky="nsew")

    jobs_scroll = ttk.Scrollbar(jobs_box, orient="vertical", command=app.jobs_tree.yview)
    app.jobs_tree.configure(yscrollcommand=jobs_scroll.set)
    jobs_scroll.grid(row=0, column=1, sticky="ns")

    jobs_actions = ttk.Frame(jobs_box)
    jobs_actions.grid(row=1, column=0, columnspan=2, sticky="ew", padx=6, pady=(8, 8))
    jobs_actions.columnconfigure((0, 1, 2, 3, 4, 5), weight=1)
    ttk.Button(jobs_actions, text="Mark In Progress", command=lambda: app._update_selected_job_status("in-progress")).grid(row=0, column=0, sticky="ew", padx=3)
    ttk.Button(jobs_actions, text="Mark Done", command=lambda: app._update_selected_job_status("completed")).grid(row=0, column=1, sticky="ew", padx=3)
    ttk.Button(jobs_actions, text="Cancel Job", command=lambda: app._update_selected_job_status("cancelled")).grid(row=0, column=2, sticky="ew", padx=3)
    ttk.Button(jobs_actions, text="Delete Job", command=app._delete_selected_job).grid(row=0, column=3, sticky="ew", padx=3)
    ttk.Button(jobs_actions, text="Create Invoice From Job", command=app._create_invoice_from_job).grid(row=0, column=4, sticky="ew", padx=3)
    ttk.Button(jobs_actions, text="Quick Rebook", command=app._quick_rebook_job).grid(row=0, column=5, sticky="ew", padx=3)

    configure_calendar_columns(app)
    update_calendar_range_label(app)


def on_calendar_view_changed(app, _event) -> None:
    """Handle calendar view change."""
    update_calendar_range_label(app)
    refresh_calendar_view(app)


def set_calendar_today(app) -> None:
    """Set calendar to today's date."""
    app.calendar_anchor_date = date.today()
    update_calendar_range_label(app)
    refresh_calendar_view(app)


def move_calendar_range(app, step: int) -> None:
    """Move calendar range forward or backward."""
    mode = app.calendar_view_var.get()
    if mode == "Day":
        app.calendar_anchor_date = app.calendar_anchor_date + timedelta(days=step)
    elif mode == "Week":
        app.calendar_anchor_date = app.calendar_anchor_date + timedelta(days=7 * step)
    elif mode == "Month":
        app.calendar_anchor_date = _shift_months(app.calendar_anchor_date, step)
    else:
        try:
            app.calendar_anchor_date = app.calendar_anchor_date.replace(
                year=app.calendar_anchor_date.year + step
            )
        except ValueError:
            app.calendar_anchor_date = app.calendar_anchor_date.replace(
                month=2,
                day=28,
                year=app.calendar_anchor_date.year + step,
            )

    update_calendar_range_label(app)
    refresh_calendar_view(app)


def _shift_months(base_date: date, month_delta: int) -> date:
    """Shift a date by a number of months."""
    total_months = (base_date.year * 12 + (base_date.month - 1)) + month_delta
    target_year = total_months // 12
    target_month = (total_months % 12) + 1
    target_day = min(base_date.day, 28)
    return date(target_year, target_month, target_day)


def _get_calendar_window(app) -> tuple[date, date]:
    """Get the start and end dates for the current calendar view."""
    mode = app.calendar_view_var.get()

    if mode == "Day":
        return app.calendar_anchor_date, app.calendar_anchor_date

    if mode == "Week":
        start = app.calendar_anchor_date - timedelta(days=app.calendar_anchor_date.weekday())
        end = start + timedelta(days=6)
        return start, end

    if mode == "Month":
        start = app.calendar_anchor_date.replace(day=1)
        next_month = (start.replace(day=28) + timedelta(days=4)).replace(day=1)
        end = next_month - timedelta(days=1)
        return start, end

    start = date(app.calendar_anchor_date.year, 1, 1)
    end = date(app.calendar_anchor_date.year, 12, 31)
    return start, end


def update_calendar_range_label(app) -> None:
    """Update the calendar range label."""
    start, end = _get_calendar_window(app)
    if start == end:
        label = start.strftime("%A, %B %d, %Y")
    elif start.year == end.year:
        label = f"{start.strftime('%b %d')} - {end.strftime('%b %d, %Y')}"
    else:
        label = f"{start.strftime('%b %d, %Y')} - {end.strftime('%b %d, %Y')}"
    app.calendar_range_var.set(label)


def configure_calendar_columns(app) -> None:
    """Configure tree columns based on current view mode."""
    mode = app.calendar_view_var.get()
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

    app.jobs_tree["columns"] = columns
    for key in columns:
        app.jobs_tree.heading(key, text=headings[key])
        anchor = "center" if key in {"status", "date", "time", "month"} else "w"
        app.jobs_tree.column(key, width=widths[key], anchor=anchor)

    app.jobs_tree.column("id", width=0, stretch=False)


def refresh_calendar_view(app) -> None:
    """Refresh the calendar view with current jobs."""
    if not hasattr(app, "jobs_tree"):
        return

    configure_calendar_columns(app)

    for item in app.jobs_tree.get_children():
        app.jobs_tree.delete(item)

    start_date, end_date = _get_calendar_window(app)
    mode = app.calendar_view_var.get()

    for row in app.calendar_jobs_cache:
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

        app.jobs_tree.insert("", tk.END, values=values)
