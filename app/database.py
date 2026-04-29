import sqlite3
from typing import Any, Iterable

from app.config import DB_PATH, DEFAULT_SETTINGS, ensure_runtime_directories

SCHEMA_SQL = """
CREATE TABLE IF NOT EXISTS app_settings (
    key TEXT PRIMARY KEY,
    value TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS customers (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    name TEXT NOT NULL,
    email TEXT NOT NULL,
    phone TEXT NOT NULL DEFAULT '',
    sms_gateway TEXT NOT NULL DEFAULT '',
    bedrooms TEXT NOT NULL DEFAULT '',
    bathrooms TEXT NOT NULL DEFAULT '',
    square_feet TEXT NOT NULL DEFAULT '',
    cleaning_duration TEXT NOT NULL DEFAULT '',
    address TEXT NOT NULL,
    notes TEXT NOT NULL DEFAULT '',
    created_at TEXT NOT NULL,
    updated_at TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS invoices (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    invoice_number TEXT NOT NULL UNIQUE,
    customer_id INTEGER NOT NULL,
    issue_date TEXT NOT NULL,
    due_date TEXT NOT NULL,
    status TEXT NOT NULL,
    subtotal REAL NOT NULL,
    tax_rate REAL NOT NULL,
    tax_amount REAL NOT NULL,
    total_amount REAL NOT NULL,
    notes TEXT NOT NULL DEFAULT '',
    pdf_path TEXT NOT NULL DEFAULT '',
    sent_at TEXT,
    paid_at TEXT,
    created_at TEXT NOT NULL,
    updated_at TEXT NOT NULL,
    FOREIGN KEY(customer_id) REFERENCES customers(id)
);

CREATE TABLE IF NOT EXISTS invoice_items (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    invoice_id INTEGER NOT NULL,
    description TEXT NOT NULL,
    quantity REAL NOT NULL,
    unit_price REAL NOT NULL,
    line_total REAL NOT NULL,
    display_order INTEGER NOT NULL,
    FOREIGN KEY(invoice_id) REFERENCES invoices(id) ON DELETE CASCADE
);

CREATE TABLE IF NOT EXISTS cleaners (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    name TEXT NOT NULL,
    phone TEXT NOT NULL DEFAULT '',
    email TEXT NOT NULL DEFAULT '',
    google_calendar_id TEXT NOT NULL DEFAULT '',
    notes TEXT NOT NULL DEFAULT '',
    active INTEGER NOT NULL DEFAULT 1,
    created_at TEXT NOT NULL,
    updated_at TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS jobs (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    customer_id INTEGER NOT NULL,
    cleaner_id INTEGER NOT NULL,
    title TEXT NOT NULL,
    start_at TEXT NOT NULL,
    end_at TEXT NOT NULL,
    location TEXT NOT NULL DEFAULT '',
    notes TEXT NOT NULL DEFAULT '',
    status TEXT NOT NULL DEFAULT 'scheduled',
    google_event_id TEXT NOT NULL DEFAULT '',
    created_at TEXT NOT NULL,
    updated_at TEXT NOT NULL,
    FOREIGN KEY(customer_id) REFERENCES customers(id),
    FOREIGN KEY(cleaner_id) REFERENCES cleaners(id)
);

CREATE INDEX IF NOT EXISTS idx_invoices_customer_id ON invoices(customer_id);
CREATE INDEX IF NOT EXISTS idx_invoice_items_invoice_id ON invoice_items(invoice_id);
CREATE INDEX IF NOT EXISTS idx_jobs_cleaner_time ON jobs(cleaner_id, start_at, end_at);
CREATE INDEX IF NOT EXISTS idx_jobs_customer_id ON jobs(customer_id);
"""


class DatabaseManager:
    def __init__(self, db_path: str | None = None) -> None:
        self.db_path = db_path or str(DB_PATH)

    def initialize(self) -> None:
        ensure_runtime_directories()
        with self.get_connection() as connection:
            connection.executescript(SCHEMA_SQL)
            self._run_migrations(connection)
            self._seed_default_settings(connection)
            connection.commit()

    def get_connection(self) -> sqlite3.Connection:
        connection = sqlite3.connect(self.db_path)
        connection.row_factory = sqlite3.Row
        connection.execute("PRAGMA foreign_keys = ON")
        return connection

    def fetchone(self, query: str, params: Iterable[Any] = ()) -> sqlite3.Row | None:
        with self.get_connection() as connection:
            cursor = connection.execute(query, tuple(params))
            return cursor.fetchone()

    def fetchall(self, query: str, params: Iterable[Any] = ()) -> list[sqlite3.Row]:
        with self.get_connection() as connection:
            cursor = connection.execute(query, tuple(params))
            return cursor.fetchall()

    def execute(self, query: str, params: Iterable[Any] = ()) -> int:
        with self.get_connection() as connection:
            cursor = connection.execute(query, tuple(params))
            connection.commit()
            return int(cursor.lastrowid)

    def executemany(self, query: str, seq_of_params: Iterable[Iterable[Any]]) -> None:
        with self.get_connection() as connection:
            connection.executemany(query, seq_of_params)
            connection.commit()

    def _seed_default_settings(self, connection: sqlite3.Connection) -> None:
        rows = connection.execute("SELECT key FROM app_settings").fetchall()
        existing_keys = {row[0] for row in rows}
        for key, value in DEFAULT_SETTINGS.items():
            if key not in existing_keys:
                connection.execute(
                    "INSERT INTO app_settings (key, value) VALUES (?, ?)",
                    (key, value),
                )

    def _run_migrations(self, connection: sqlite3.Connection) -> None:
        self._ensure_column(connection, "customers", "phone", "TEXT NOT NULL DEFAULT ''")
        self._ensure_column(connection, "customers", "sms_gateway", "TEXT NOT NULL DEFAULT ''")
        self._ensure_column(connection, "customers", "bedrooms", "TEXT NOT NULL DEFAULT ''")
        self._ensure_column(connection, "customers", "bathrooms", "TEXT NOT NULL DEFAULT ''")
        self._ensure_column(connection, "customers", "square_feet", "TEXT NOT NULL DEFAULT ''")
        self._ensure_column(connection, "customers", "cleaning_duration", "TEXT NOT NULL DEFAULT ''")
        self._ensure_column(connection, "cleaners", "google_calendar_id", "TEXT NOT NULL DEFAULT ''")
        self._ensure_column(connection, "jobs", "google_event_id", "TEXT NOT NULL DEFAULT ''")
        self._ensure_column(connection, "customers", "frequency", "TEXT NOT NULL DEFAULT 'Single'")

    def _ensure_column(
        self,
        connection: sqlite3.Connection,
        table_name: str,
        column_name: str,
        column_definition: str,
    ) -> None:
        rows = connection.execute(f"PRAGMA table_info({table_name})").fetchall()
        existing_columns = {row["name"] for row in rows}
        if column_name in existing_columns:
            return

        connection.execute(
            f"ALTER TABLE {table_name} ADD COLUMN {column_name} {column_definition}"
        )
