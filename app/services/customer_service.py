from datetime import datetime

from app.database import DatabaseManager


class CustomerService:
    def __init__(self, db: DatabaseManager) -> None:
        self.db = db

    def list_customers(self) -> list[dict]:
        rows = self.db.fetchall(
            """
            SELECT
                c.id,
                c.name,
                c.email,
                c.phone,
                c.sms_gateway,
                c.bedrooms,
                c.bathrooms,
                c.square_feet,
                c.cleaning_duration,
                c.frequency,
                c.address,
                c.notes,
                MAX(i.issue_date) AS last_invoice_date
            FROM customers c
            LEFT JOIN invoices i ON i.customer_id = c.id
            GROUP BY c.id
            ORDER BY c.name COLLATE NOCASE ASC
            """
        )
        return [dict(row) for row in rows]

    def create_customer(
        self,
        name: str,
        email: str,
        phone: str,
        bedrooms: str,
        bathrooms: str,
        square_feet: str,
        cleaning_duration: str,
        address: str,
        notes: str = "",
        sms_gateway: str = "",
        frequency: str = "Single",
    ) -> int:
        now = datetime.now().isoformat(timespec="seconds")
        return self.db.execute(
            """
            INSERT INTO customers (
                name, email, phone, sms_gateway,
                bedrooms, bathrooms, square_feet, cleaning_duration,
                frequency, address, notes, created_at, updated_at
            )
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                name.strip(),
                email.strip(),
                phone.strip(),
                sms_gateway.strip().lower().lstrip("@"),
                bedrooms.strip(),
                bathrooms.strip(),
                square_feet.strip(),
                cleaning_duration.strip(),
                frequency.strip(),
                address.strip(),
                notes.strip(),
                now,
                now,
            ),
        )

    def update_customer(
        self,
        customer_id: int,
        name: str,
        email: str,
        phone: str,
        bedrooms: str,
        bathrooms: str,
        square_feet: str,
        cleaning_duration: str,
        address: str,
        notes: str = "",
        sms_gateway: str = "",
        frequency: str = "Single",
    ) -> None:
        now = datetime.now().isoformat(timespec="seconds")
        self.db.execute(
            """
            UPDATE customers
            SET
                name = ?,
                email = ?,
                phone = ?,
                sms_gateway = ?,
                bedrooms = ?,
                bathrooms = ?,
                square_feet = ?,
                cleaning_duration = ?,
                frequency = ?,
                address = ?,
                notes = ?,
                updated_at = ?
            WHERE id = ?
            """,
            (
                name.strip(),
                email.strip(),
                phone.strip(),
                sms_gateway.strip().lower().lstrip("@"),
                bedrooms.strip(),
                bathrooms.strip(),
                square_feet.strip(),
                cleaning_duration.strip(),
                frequency.strip(),
                address.strip(),
                notes.strip(),
                now,
                customer_id,
            ),
        )

    def delete_customer(self, customer_id: int) -> None:
        invoice_count = self.count_customer_invoices(customer_id)
        if invoice_count > 0:
            raise ValueError(
                "Cannot delete this customer because linked invoices exist "
                f"({invoice_count}). Use Delete + Invoices or delete invoices first in Invoice History."
            )
        self.db.execute("DELETE FROM customers WHERE id = ?", (customer_id,))

    def delete_customer_with_invoices(self, customer_id: int) -> int:
        invoice_count = self.count_customer_invoices(customer_id)
        with self.db.get_connection() as connection:
            connection.execute("DELETE FROM invoices WHERE customer_id = ?", (customer_id,))
            connection.execute("DELETE FROM customers WHERE id = ?", (customer_id,))
            connection.commit()
        return invoice_count

    def count_customer_invoices(self, customer_id: int) -> int:
        row = self.db.fetchone(
            "SELECT COUNT(*) AS invoice_count FROM invoices WHERE customer_id = ?",
            (customer_id,),
        )
        return int(row["invoice_count"]) if row else 0

    def get_customer(self, customer_id: int) -> dict | None:
        row = self.db.fetchone(
            """
            SELECT
                id, name, email, phone, sms_gateway,
                bedrooms, bathrooms, square_feet, cleaning_duration,
                frequency, address, notes
            FROM customers
            WHERE id = ?
            """,
            (customer_id,),
        )
        return dict(row) if row else None
