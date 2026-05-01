from __future__ import annotations

import tkinter as tk
from datetime import date, datetime, timedelta
from tkinter import messagebox, simpledialog, ttk

from app.services.google_calendar_service import GoogleCalendarService
from app.ui.constants import CALENDAR_DATETIME_FORMAT


def open_cleaners_popup(app) -> None:
    if app.cleaners_popup and app.cleaners_popup.winfo_exists():
        app.cleaners_popup.lift()
        app.cleaners_popup.focus_force()
        return

    popup = tk.Toplevel(app)
    popup.title("Manage Cleaners")
    popup.geometry("900x620")
    popup.minsize(760, 520)
    popup.transient(app)
    popup.protocol("WM_DELETE_WINDOW", lambda: _close_cleaners_popup(app))
    app.cleaners_popup = popup

    container = ttk.Frame(popup)
    container.pack(fill="both", expand=True, padx=10, pady=10)
    container.columnconfigure(0, weight=1)
    container.rowconfigure(1, weight=1)

    cleaner_box = ttk.LabelFrame(container, text="Cleaner")
    cleaner_box.grid(row=0, column=0, sticky="ew", pady=(0, 8))
    cleaner_box.columnconfigure(1, weight=1)

    ttk.Label(cleaner_box, text="Name").grid(row=0, column=0, sticky="w", padx=6, pady=(8, 4))
    ttk.Entry(cleaner_box, textvariable=app.cleaner_name_var).grid(row=0, column=1, sticky="ew", padx=6, pady=(8, 4))

    ttk.Label(cleaner_box, text="Phone").grid(row=1, column=0, sticky="w", padx=6, pady=4)
    ttk.Entry(cleaner_box, textvariable=app.cleaner_phone_var).grid(row=1, column=1, sticky="ew", padx=6, pady=4)

    ttk.Label(cleaner_box, text="Email").grid(row=2, column=0, sticky="w", padx=6, pady=4)
    ttk.Entry(cleaner_box, textvariable=app.cleaner_email_var).grid(row=2, column=1, sticky="ew", padx=6, pady=4)

    ttk.Label(cleaner_box, text="Google Calendar ID").grid(row=3, column=0, sticky="w", padx=6, pady=4)
    ttk.Entry(cleaner_box, textvariable=app.cleaner_google_calendar_id_var).grid(row=3, column=1, sticky="ew", padx=6, pady=4)

    ttk.Checkbutton(
        cleaner_box,
        text="Auto-create Google calendar on add when empty",
        variable=app.auto_create_cleaner_calendar_var,
        onvalue=1,
        offvalue=0,
    ).grid(row=4, column=0, columnspan=2, sticky="w", padx=6, pady=(2, 4))

    ttk.Label(cleaner_box, text="Notes").grid(row=5, column=0, sticky="nw", padx=6, pady=4)
    app.cleaner_notes_text = tk.Text(cleaner_box, height=4, width=34)
    app.cleaner_notes_text.grid(row=5, column=1, sticky="ew", padx=6, pady=4)

    cleaner_btns = ttk.Frame(cleaner_box)
    cleaner_btns.grid(row=6, column=0, columnspan=2, sticky="ew", padx=6, pady=(8, 8))
    cleaner_btns.columnconfigure((0, 1, 2, 3, 4), weight=1)
    ttk.Button(cleaner_btns, text="Add Cleaner", command=lambda: add_cleaner(app)).grid(row=0, column=0, sticky="ew", padx=3)
    ttk.Button(cleaner_btns, text="Update Cleaner", command=lambda: update_cleaner(app)).grid(row=0, column=1, sticky="ew", padx=3)
    ttk.Button(cleaner_btns, text="Delete Cleaner", command=lambda: delete_cleaner(app)).grid(row=0, column=2, sticky="ew", padx=3)
    ttk.Button(cleaner_btns, text="Clear", command=lambda: clear_cleaner_form(app)).grid(row=0, column=3, sticky="ew", padx=3)
    ttk.Button(cleaner_btns, text="Close", command=lambda: _close_cleaners_popup(app)).grid(row=0, column=4, sticky="ew", padx=3)

    cleaner_table = ttk.LabelFrame(container, text="Cleaner List")
    cleaner_table.grid(row=1, column=0, sticky="nsew")
    cleaner_table.columnconfigure(0, weight=1)
    cleaner_table.rowconfigure(0, weight=1)

    app.cleaners_tree = ttk.Treeview(
        cleaner_table,
        columns=("id", "name", "phone", "email", "calendar"),
        show="headings",
        height=14,
    )
    app.cleaners_tree.heading("id", text="ID")
    app.cleaners_tree.heading("name", text="Name")
    app.cleaners_tree.heading("phone", text="Phone")
    app.cleaners_tree.heading("email", text="Email")
    app.cleaners_tree.heading("calendar", text="Calendar ID")
    app.cleaners_tree.column("id", width=55, anchor="center")
    app.cleaners_tree.column("name", width=170)
    app.cleaners_tree.column("phone", width=120)
    app.cleaners_tree.column("email", width=220)
    app.cleaners_tree.column("calendar", width=290)
    app.cleaners_tree.grid(row=0, column=0, sticky="nsew")
    app.cleaners_tree.bind("<<TreeviewSelect>>", lambda event: on_cleaner_selected(app, event))

    cleaner_scroll = ttk.Scrollbar(cleaner_table, orient="vertical", command=app.cleaners_tree.yview)
    app.cleaners_tree.configure(yscrollcommand=cleaner_scroll.set)
    cleaner_scroll.grid(row=0, column=1, sticky="ns")

    load_cleaners(app)
    clear_cleaner_form(app)


def _close_cleaners_popup(app) -> None:
    if app.cleaners_popup and app.cleaners_popup.winfo_exists():
        app.cleaners_popup.destroy()
    app.cleaners_popup = None


def open_google_popup(app) -> None:
    if app.google_popup and app.google_popup.winfo_exists():
        app.google_popup.lift()
        app.google_popup.focus_force()
        return

    popup = tk.Toplevel(app)
    popup.title("Google Calendar")
    popup.geometry("760x280")
    popup.minsize(680, 240)
    popup.transient(app)
    popup.protocol("WM_DELETE_WINDOW", lambda: _close_google_popup(app))
    app.google_popup = popup

    container = ttk.Frame(popup)
    container.pack(fill="both", expand=True, padx=10, pady=10)
    container.columnconfigure(1, weight=1)

    ttk.Label(container, textvariable=app.google_status_var).grid(
        row=0,
        column=0,
        columnspan=2,
        sticky="w",
        pady=(0, 8),
    )

    ttk.Label(container, text="Google Calendar").grid(row=1, column=0, sticky="w", padx=(0, 6), pady=4)
    app.google_calendars_combo = ttk.Combobox(
        container,
        textvariable=app.google_calendars_combo_var,
        state="readonly",
    )
    app.google_calendars_combo.grid(row=1, column=1, sticky="ew", pady=4)

    ttk.Label(container, text="Assign To Cleaner").grid(row=2, column=0, sticky="w", padx=(0, 6), pady=4)
    app.google_target_cleaner_combo = ttk.Combobox(
        container,
        textvariable=app.google_target_cleaner_var,
        state="readonly",
    )
    app.google_target_cleaner_combo.grid(row=2, column=1, sticky="ew", pady=4)

    btns = ttk.Frame(container)
    btns.grid(row=3, column=0, columnspan=2, sticky="ew", pady=(10, 0))
    btns.columnconfigure((0, 1, 2, 3), weight=1)
    ttk.Button(btns, text="Connect Google", command=lambda: _connect_google_calendar(app)).grid(row=0, column=0, sticky="ew", padx=3)
    ttk.Button(btns, text="Load Calendars", command=lambda: _load_google_calendars(app)).grid(row=0, column=1, sticky="ew", padx=3)
    ttk.Button(btns, text="Assign To Cleaner", command=lambda: _apply_google_calendar_to_cleaner(app)).grid(row=0, column=2, sticky="ew", padx=3)
    ttk.Button(btns, text="Close", command=lambda: _close_google_popup(app)).grid(row=0, column=3, sticky="ew", padx=3)

    load_cleaners(app)
    refresh_google_status(app)


def _close_google_popup(app) -> None:
    if app.google_popup and app.google_popup.winfo_exists():
        app.google_popup.destroy()
    app.google_popup = None


def open_schedule_popup(app) -> None:
    if app.schedule_popup and app.schedule_popup.winfo_exists():
        app.schedule_popup.lift()
        app.schedule_popup.focus_force()
        return

    popup = tk.Toplevel(app)
    popup.title("Schedule Job")
    popup.geometry("900x430")
    popup.minsize(820, 380)
    popup.transient(app)
    popup.protocol("WM_DELETE_WINDOW", lambda: _close_schedule_popup(app))
    app.schedule_popup = popup

    container = ttk.Frame(popup)
    container.pack(fill="both", expand=True, padx=10, pady=10)
    for idx in range(4):
        container.columnconfigure(idx, weight=1)

    ttk.Label(container, text="Customer").grid(row=0, column=0, sticky="w", padx=6, pady=(4, 4))
    app.job_customer_combo = ttk.Combobox(container, textvariable=app.job_customer_var, state="readonly")
    app.job_customer_combo.grid(row=0, column=1, sticky="ew", padx=6, pady=(4, 4))

    ttk.Label(container, text="Cleaner").grid(row=0, column=2, sticky="w", padx=6, pady=(4, 4))
    app.job_cleaner_combo = ttk.Combobox(container, textvariable=app.job_cleaner_var, state="readonly")
    app.job_cleaner_combo.grid(row=0, column=3, sticky="ew", padx=6, pady=(4, 4))

    ttk.Label(container, text="Job Title").grid(row=1, column=0, sticky="w", padx=6, pady=4)
    ttk.Entry(container, textvariable=app.job_title_var).grid(row=1, column=1, sticky="ew", padx=6, pady=4)

    ttk.Label(container, text="Location").grid(row=1, column=2, sticky="w", padx=6, pady=4)
    ttk.Entry(container, textvariable=app.job_location_var).grid(row=1, column=3, sticky="ew", padx=6, pady=4)

    ttk.Label(container, text="Start (YYYY-MM-DD HH:MM)").grid(row=2, column=0, sticky="w", padx=6, pady=4)
    ttk.Entry(container, textvariable=app.job_start_var).grid(row=2, column=1, sticky="ew", padx=6, pady=4)

    ttk.Label(container, text="End (YYYY-MM-DD HH:MM)").grid(row=2, column=2, sticky="w", padx=6, pady=4)
    ttk.Entry(container, textvariable=app.job_end_var).grid(row=2, column=3, sticky="ew", padx=6, pady=4)

    ttk.Label(container, text="Notes").grid(row=3, column=0, sticky="nw", padx=6, pady=4)
    app.job_notes_text = tk.Text(container, height=6, width=50)
    app.job_notes_text.grid(row=3, column=1, columnspan=3, sticky="ew", padx=6, pady=4)

    btns = ttk.Frame(container)
    btns.grid(row=4, column=0, columnspan=4, sticky="ew", padx=6, pady=(10, 0))
    btns.columnconfigure((0, 1, 2, 3), weight=1)
    ttk.Button(btns, text="Check Availability", command=lambda: check_job_availability(app)).grid(row=0, column=0, sticky="ew", padx=3)
    ttk.Button(btns, text="Schedule Job", command=lambda: schedule_job(app)).grid(row=0, column=1, sticky="ew", padx=3)
    ttk.Button(btns, text="Reset", command=lambda: reset_job_form(app)).grid(row=0, column=2, sticky="ew", padx=3)
    ttk.Button(btns, text="Close", command=lambda: _close_schedule_popup(app)).grid(row=0, column=3, sticky="ew", padx=3)

    app._load_customers()
    load_cleaners(app)
    reset_job_form(app)


def _close_schedule_popup(app) -> None:
    if app.schedule_popup and app.schedule_popup.winfo_exists():
        app.schedule_popup.destroy()
    app.schedule_popup = None


def refresh_google_status(app) -> None:
    try:
        service = _get_google_calendar_service(app)
        if app.google_connected:
            status = "Status: Connected"
        elif service.is_configured():
            status = "Status: Configured (connect required)"
        else:
            status = "Status: Credentials file missing"
    except Exception as exc:
        status = f"Status: Not configured ({exc})"

    if hasattr(app, "google_status_var"):
        app.google_status_var.set(status)


def _get_google_calendar_service(app) -> GoogleCalendarService:
    settings = app.settings_service.get_settings()
    credentials_file = settings.get("google_credentials_file", "").strip()
    token_file = settings.get("google_token_file", "").strip()

    if not credentials_file or not token_file:
        raise ValueError("Set Google credentials and token file paths in Settings first.")

    if (
        app.google_calendar_service is None
        or app.google_calendar_service.credentials_file != credentials_file
        or app.google_calendar_service.token_file != token_file
    ):
        app.google_calendar_service = GoogleCalendarService(credentials_file, token_file)
        app.google_connected = False

    return app.google_calendar_service


def _connect_google_calendar(app) -> None:
    try:
        service = _get_google_calendar_service(app)
        service.connect()
        app.google_connected = True
        refresh_google_status(app)
    except Exception as exc:
        messagebox.showerror("Google Calendar", str(exc))
        return

    _load_google_calendars(app)


def _load_google_calendars(app) -> None:
    try:
        service = _get_google_calendar_service(app)
        calendars = service.list_calendars()
        app.google_connected = True
        refresh_google_status(app)
    except Exception as exc:
        messagebox.showerror("Google Calendar", str(exc))
        return

    labels: list[str] = []
    app.google_calendar_lookup.clear()

    for cal in calendars:
        cal_id = cal.get("id", "").strip()
        summary = cal.get("summary", "Untitled")
        if not cal_id:
            continue
        label = f"{summary} ({cal_id})"
        labels.append(label)
        app.google_calendar_lookup[label] = cal_id

    if hasattr(app, "google_calendars_combo") and app.google_calendars_combo.winfo_exists():
        app.google_calendars_combo["values"] = labels
    if labels:
        app.google_calendars_combo_var.set(labels[0])
        messagebox.showinfo("Google Calendar", f"Loaded {len(labels)} calendar(s).")
    else:
        messagebox.showwarning("Google Calendar", "No calendars were returned by Google.")


def _apply_google_calendar_to_cleaner(app) -> None:
    selected = app.google_calendars_combo_var.get().strip()
    if selected not in app.google_calendar_lookup:
        messagebox.showwarning("Google Calendar", "Select a calendar first.")
        return

    cleaner_label = app.google_target_cleaner_var.get().strip()
    if cleaner_label not in app.cleaner_lookup:
        messagebox.showwarning("Google Calendar", "Select a cleaner to assign this calendar to.")
        return

    cleaner_id = app.cleaner_lookup[cleaner_label]
    cleaner = app.calendar_service.get_cleaner(cleaner_id)
    if not cleaner:
        messagebox.showwarning("Google Calendar", "Selected cleaner was not found.")
        return

    calendar_id = app.google_calendar_lookup[selected]
    app.calendar_service.update_cleaner(
        cleaner_id=cleaner_id,
        name=cleaner["name"],
        phone=cleaner["phone"],
        email=cleaner["email"],
        notes=cleaner["notes"],
        google_calendar_id=calendar_id,
    )

    app.cleaner_google_calendar_id_var.set(calendar_id)
    load_cleaners(app)
    messagebox.showinfo("Google Calendar", "Calendar assigned to cleaner.")


def load_cleaners(app) -> None:
    rows = app.calendar_service.list_cleaners()

    if hasattr(app, "cleaners_tree") and app.cleaners_tree.winfo_exists():
        for item in app.cleaners_tree.get_children():
            app.cleaners_tree.delete(item)

    cleaner_labels: list[str] = []
    app.cleaner_lookup.clear()

    for row in rows:
        if hasattr(app, "cleaners_tree") and app.cleaners_tree.winfo_exists():
            app.cleaners_tree.insert(
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
        app.cleaner_lookup[label] = row["id"]

    if hasattr(app, "job_cleaner_combo") and app.job_cleaner_combo.winfo_exists():
        app.job_cleaner_combo["values"] = cleaner_labels

    if hasattr(app, "google_target_cleaner_combo") and app.google_target_cleaner_combo.winfo_exists():
        app.google_target_cleaner_combo["values"] = cleaner_labels


def clear_cleaner_form(app) -> None:
    app.cleaner_id_var.set("")
    app.cleaner_name_var.set("")
    app.cleaner_phone_var.set("")
    app.cleaner_email_var.set("")
    app.cleaner_google_calendar_id_var.set("")
    if hasattr(app, "cleaner_notes_text") and app.cleaner_notes_text.winfo_exists():
        app.cleaner_notes_text.delete("1.0", tk.END)


def on_cleaner_selected(app, _event) -> None:
    selection = app.cleaners_tree.selection()
    if not selection:
        return

    values = app.cleaners_tree.item(selection[0], "values")
    cleaner_id = int(values[0])
    cleaner = app.calendar_service.get_cleaner(cleaner_id)
    if not cleaner:
        return

    app.cleaner_id_var.set(str(cleaner["id"]))
    app.cleaner_name_var.set(cleaner["name"])
    app.cleaner_phone_var.set(cleaner["phone"])
    app.cleaner_email_var.set(cleaner["email"])
    app.cleaner_google_calendar_id_var.set(cleaner["google_calendar_id"])
    if hasattr(app, "cleaner_notes_text") and app.cleaner_notes_text.winfo_exists():
        app.cleaner_notes_text.delete("1.0", tk.END)
        app.cleaner_notes_text.insert("1.0", cleaner["notes"])


def add_cleaner(app) -> None:
    name = app.cleaner_name_var.get().strip()
    if not name:
        messagebox.showwarning("Validation", "Cleaner name is required.")
        return

    google_calendar_id = app.cleaner_google_calendar_id_var.get().strip()
    if not google_calendar_id and app.auto_create_cleaner_calendar_var.get() == 1:
        try:
            service = _get_google_calendar_service(app)
            created = service.create_calendar(
                summary=f"Cleaner - {name}",
                description=(
                    "Auto-created by Cleaning Invoice Generator "
                    f"for cleaner '{name}'."
                ),
            )
            google_calendar_id = created.get("id", "").strip()
            app.cleaner_google_calendar_id_var.set(google_calendar_id)
            app.google_connected = True
            refresh_google_status(app)
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

    app.calendar_service.create_cleaner(
        name=name,
        phone=app.cleaner_phone_var.get().strip(),
        email=app.cleaner_email_var.get().strip(),
        notes=app.cleaner_notes_text.get("1.0", tk.END).strip(),
        google_calendar_id=google_calendar_id,
    )
    load_cleaners(app)
    app._refresh_live_views()
    clear_cleaner_form(app)
    if google_calendar_id:
        messagebox.showinfo("Cleaner", "Cleaner added and linked to Google Calendar.")
    else:
        messagebox.showinfo("Cleaner", "Cleaner added.")


def update_cleaner(app) -> None:
    if not app.cleaner_id_var.get().strip():
        messagebox.showwarning("Select", "Select a cleaner first.")
        return

    name = app.cleaner_name_var.get().strip()
    if not name:
        messagebox.showwarning("Validation", "Cleaner name is required.")
        return

    app.calendar_service.update_cleaner(
        cleaner_id=int(app.cleaner_id_var.get()),
        name=name,
        phone=app.cleaner_phone_var.get().strip(),
        email=app.cleaner_email_var.get().strip(),
        notes=app.cleaner_notes_text.get("1.0", tk.END).strip(),
        google_calendar_id=app.cleaner_google_calendar_id_var.get().strip(),
    )
    load_cleaners(app)
    app._refresh_live_views()
    messagebox.showinfo("Cleaner", "Cleaner updated.")


def delete_cleaner(app) -> None:
    if not app.cleaner_id_var.get().strip():
        messagebox.showwarning("Select", "Select a cleaner first.")
        return

    if not messagebox.askyesno("Confirm", "Delete selected cleaner?"):
        return

    cleaner_id = int(app.cleaner_id_var.get())
    cleaner = app.calendar_service.get_cleaner(cleaner_id)

    try:
        google_calendar_id = (cleaner or {}).get("google_calendar_id", "").strip()
        if google_calendar_id:
            active_with_same_calendar = [
                row
                for row in app.calendar_service.list_cleaners()
                if int(row["id"]) != cleaner_id and row.get("google_calendar_id", "").strip() == google_calendar_id
            ]

            if active_with_same_calendar:
                messagebox.showwarning(
                    "Google Calendar",
                    "This calendar is assigned to another cleaner, so it was not deleted.",
                )
            else:
                try:
                    service = _get_google_calendar_service(app)
                    service.delete_calendar(google_calendar_id)
                    app.google_connected = True
                    refresh_google_status(app)
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

        app.calendar_service.delete_cleaner(cleaner_id)
    except ValueError as exc:
        messagebox.showwarning("Delete Blocked", str(exc))
        return

    load_cleaners(app)
    app._refresh_live_views()
    clear_cleaner_form(app)
    messagebox.showinfo("Cleaner", "Cleaner deleted.")


def reset_job_form(app) -> None:
    now = datetime.now().replace(minute=0, second=0, microsecond=0)
    app.job_customer_var.set("")
    app.job_cleaner_var.set("")
    app.job_title_var.set("")
    app.job_location_var.set("")
    app.job_start_var.set(now.strftime(CALENDAR_DATETIME_FORMAT))
    app.job_end_var.set((now + timedelta(hours=2)).strftime(CALENDAR_DATETIME_FORMAT))
    if hasattr(app, "job_notes_text") and app.job_notes_text.winfo_exists():
        app.job_notes_text.delete("1.0", tk.END)


def _selected_job_id(app) -> int | None:
    selection = app.jobs_tree.selection()
    if not selection:
        messagebox.showwarning("Select", "Select a job first.")
        return None

    values = app.jobs_tree.item(selection[0], "values")
    return int(values[0])


def load_jobs(app) -> None:
    app.calendar_jobs_cache = app.calendar_service.list_jobs()
    app._refresh_calendar_view()


def _validate_job_form(app) -> tuple[int, int, str, str, str, str, str]:
    if not hasattr(app, "job_notes_text") or not app.job_notes_text.winfo_exists():
        raise ValueError("Open the Schedule Job popup first.")

    customer_label = app.job_customer_var.get().strip()
    cleaner_label = app.job_cleaner_var.get().strip()

    if customer_label not in app.customer_lookup:
        raise ValueError("Select a customer.")
    if cleaner_label not in app.cleaner_lookup:
        raise ValueError("Select a cleaner.")

    title = app.job_title_var.get().strip()
    if not title:
        raise ValueError("Job title is required.")

    start_at = app.job_start_var.get().strip()
    end_at = app.job_end_var.get().strip()

    try:
        start_dt = datetime.strptime(start_at, CALENDAR_DATETIME_FORMAT)
        end_dt = datetime.strptime(end_at, CALENDAR_DATETIME_FORMAT)
    except ValueError as exc:
        raise ValueError("Start/End must use format YYYY-MM-DD HH:MM.") from exc

    if end_dt <= start_dt:
        raise ValueError("End time must be after start time.")

    location = app.job_location_var.get().strip()
    notes = app.job_notes_text.get("1.0", tk.END).strip()

    return (
        app.customer_lookup[customer_label],
        app.cleaner_lookup[cleaner_label],
        title,
        start_at,
        end_at,
        location,
        notes,
    )


def check_job_availability(app) -> None:
    try:
        _, cleaner_id, _, start_at, end_at, _, _ = _validate_job_form(app)
    except ValueError as exc:
        messagebox.showwarning("Validation", str(exc))
        return

    conflicts = app.calendar_service.check_cleaner_availability(
        cleaner_id=cleaner_id,
        start_at=start_at,
        end_at=end_at,
    )

    if not conflicts:
        cleaner = app.calendar_service.get_cleaner(cleaner_id)
        google_calendar_id = (cleaner or {}).get("google_calendar_id", "").strip()
        if google_calendar_id:
            try:
                service = _get_google_calendar_service(app)
                is_busy = service.check_time_busy(
                    google_calendar_id,
                    start_at,
                    end_at,
                    cleaner_id=cleaner_id,
                )
                app.google_connected = True
                refresh_google_status(app)
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


def schedule_job(app) -> None:
    try:
        customer_id, cleaner_id, title, start_at, end_at, location, notes = _validate_job_form(app)
        cleaner = app.calendar_service.get_cleaner(cleaner_id)
        google_calendar_id = (cleaner or {}).get("google_calendar_id", "").strip()
        if google_calendar_id:
            service = _get_google_calendar_service(app)
            is_busy = service.check_time_busy(
                google_calendar_id,
                start_at,
                end_at,
                cleaner_id=cleaner_id,
            )
            app.google_connected = True
            refresh_google_status(app)
            if is_busy:
                raise ValueError(
                    "Schedule blocked: cleaner has a conflict in Google Calendar for that time range."
                )

        job_id = app.calendar_service.create_job(
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
            service = _get_google_calendar_service(app)
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
                app.calendar_service.set_job_google_event_id(job_id, google_event_id)
            app.google_connected = True
            refresh_google_status(app)
        except Exception as exc:
            messagebox.showwarning(
                "Google Sync",
                f"Job was scheduled locally, but Google sync failed: {exc}",
            )

    load_jobs(app)
    app._refresh_live_views()
    messagebox.showinfo("Job", "Job scheduled successfully.")


def update_selected_job_status(app, status: str) -> None:
    job_id = _selected_job_id(app)
    if job_id is None:
        return

    if status == "cancelled":
        job = app.calendar_service.get_job(job_id)
        if job:
            google_event_id = (job.get("google_event_id") or "").strip()
            google_calendar_id = (job.get("google_calendar_id") or "").strip()
            if google_event_id and google_calendar_id:
                try:
                    service = _get_google_calendar_service(app)
                    service.delete_event(google_calendar_id, google_event_id)
                    app.calendar_service.set_job_google_event_id(job_id, "")
                    app.google_connected = True
                    refresh_google_status(app)
                except Exception:
                    pass

    app.calendar_service.update_job_status(job_id, status)
    load_jobs(app)
    app._refresh_live_views()
    messagebox.showinfo("Job", f"Job updated to '{status}'.")


def delete_selected_job(app) -> None:
    job_id = _selected_job_id(app)
    if job_id is None:
        return

    if not messagebox.askyesno("Confirm", "Delete selected job?"):
        return

    job = app.calendar_service.get_job(job_id)
    if job:
        google_event_id = (job.get("google_event_id") or "").strip()
        google_calendar_id = (job.get("google_calendar_id") or "").strip()
        if google_event_id and google_calendar_id:
            try:
                service = _get_google_calendar_service(app)
                service.delete_event(google_calendar_id, google_event_id)
                app.google_connected = True
                refresh_google_status(app)
            except Exception:
                pass

    app.calendar_service.delete_job(job_id)
    load_jobs(app)
    app._refresh_live_views()
    messagebox.showinfo("Job", "Job deleted.")


def create_invoice_from_job(app) -> None:
    job_id = _selected_job_id(app)
    if job_id is None:
        return

    job = app.calendar_service.get_job(job_id)
    if not job:
        messagebox.showerror("Error", "Job not found.")
        return

    customer_id = job["customer_id"]
    actual_label = None
    for label, cid in app.customer_lookup.items():
        if cid == customer_id:
            actual_label = label
            break

    if not actual_label:
        messagebox.showwarning("Error", "Customer not found in system.")
        return

    rate_str = simpledialog.askstring(
        "Service Rate",
        f"Enter the hourly or flat rate for this service:\n(Job: {job['title']})",
        initialvalue="0",
    )
    if rate_str is None:
        return

    try:
        rate = float(rate_str)
        if rate < 0:
            raise ValueError
    except ValueError:
        messagebox.showwarning("Validation", "Rate must be a non-negative number.")
        return

    app._reset_invoice_form()
    app.invoice_customer_var.set(actual_label)
    app.item_description_var.set(job["title"])
    app.item_rate_var.set(str(rate))
    app.notebook.select(app.invoice_tab)

    messagebox.showinfo("Invoice Created", "Switch to Create Invoice tab to complete the invoice.")


def quick_rebook_job(app) -> None:
    job_id = _selected_job_id(app)
    if job_id is None:
        return

    job = app.calendar_service.get_job(job_id)
    if not job:
        messagebox.showerror("Error", "Job not found.")
        return

    customer_id = job["customer_id"]
    cleaner_id = job["cleaner_id"]
    customer_label = None
    cleaner_label = None

    for label, cid in app.customer_lookup.items():
        if cid == customer_id:
            customer_label = label
            break

    for label, clid in app.cleaner_lookup.items():
        if clid == cleaner_id:
            cleaner_label = label
            break

    if not customer_label or not cleaner_label:
        messagebox.showwarning("Error", "Customer or cleaner not found.")
        return

    if not app.schedule_popup or not app.schedule_popup.winfo_exists():
        open_schedule_popup(app)

    app.job_customer_var.set(customer_label)
    app.job_cleaner_var.set(cleaner_label)
    app.job_title_var.set(job["title"])
    app.job_location_var.set(job.get("location", ""))

    try:
        start_dt = datetime.strptime(job["start_at"], CALENDAR_DATETIME_FORMAT)
        end_dt = datetime.strptime(job["end_at"], CALENDAR_DATETIME_FORMAT)
        new_start = start_dt + timedelta(days=7)
        new_end = end_dt + timedelta(days=7)
        app.job_start_var.set(new_start.strftime(CALENDAR_DATETIME_FORMAT))
        app.job_end_var.set(new_end.strftime(CALENDAR_DATETIME_FORMAT))
    except ValueError:
        now = datetime.now().replace(minute=0, second=0, microsecond=0)
        app.job_start_var.set(now.strftime(CALENDAR_DATETIME_FORMAT))
        app.job_end_var.set((now + timedelta(hours=2)).strftime(CALENDAR_DATETIME_FORMAT))

    if hasattr(app, "job_notes_text") and app.job_notes_text.winfo_exists():
        app.job_notes_text.delete("1.0", tk.END)
        app.job_notes_text.insert("1.0", f"Rebook of {job['title']} on {job['start_at'].split()[0]}")

    if app.schedule_popup and app.schedule_popup.winfo_exists():
        app.schedule_popup.lift()
        app.schedule_popup.focus_force()

    messagebox.showinfo("Quick Rebook", "Job form pre-filled. Adjust dates/times as needed and click Schedule Job.")