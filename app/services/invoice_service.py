from datetime import date, datetime
from decimal import Decimal, ROUND_HALF_UP

from app.config import DATE_FORMAT
from app.database import DatabaseManager
from app.models import InvoiceTotals


class InvoiceService:
    def __init__(self, db: DatabaseManager) -> None:
        self.db = db

    @staticmethod
    def _money(value: Decimal) -> Decimal:
        return value.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)

    @staticmethod
    def _to_decimal(value: float | str | Decimal) -> Decimal:
        return Decimal(str(value))

    def calculate_totals(self, items: list[dict], tax_rate: float | str) -> InvoiceTotals:
        subtotal = Decimal("0")
        for item in items:
            quantity = self._to_decimal(item["quantity"])
            unit_price = self._to_decimal(item["unit_price"])
            subtotal += quantity * unit_price

        subtotal = self._money(subtotal)
        tax_rate_dec = self._to_decimal(tax_rate)
        tax_amount = self._money(subtotal * (tax_rate_dec / Decimal("100")))
        total_amount = self._money(subtotal + tax_amount)
        return InvoiceTotals(subtotal=subtotal, tax_amount=tax_amount, total_amount=total_amount)

    def generate_invoice_number(self, issue_date_text: str) -> str:
        try:
            year = datetime.strptime(issue_date_text, DATE_FORMAT).year
        except ValueError:
            year = date.today().year

        prefix = f"INV-{year}-"
        row = self.db.fetchone(
            """
            SELECT invoice_number
            FROM invoices
            WHERE invoice_number LIKE ?
            ORDER BY id DESC
            LIMIT 1
            """,
            (f"{prefix}%",),
        )

        if not row:
            return f"{prefix}0001"

        current = row["invoice_number"].replace(prefix, "")
        next_num = int(current) + 1 if current.isdigit() else 1
        return f"{prefix}{next_num:04d}"

    def create_invoice(
        self,
        customer_id: int,
        issue_date: str,
        due_date: str,
        tax_rate: float | str,
        notes: str,
        items: list[dict],
        status: str = "draft",
    ) -> dict:
        totals = self.calculate_totals(items, tax_rate)
        invoice_number = self.generate_invoice_number(issue_date)
        now = datetime.now().isoformat(timespec="seconds")

        with self.db.get_connection() as connection:
            cursor = connection.execute(
                """
                INSERT INTO invoices (
                    invoice_number, customer_id, issue_date, due_date, status,
                    subtotal, tax_rate, tax_amount, total_amount, notes,
                    created_at, updated_at
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    invoice_number,
                    customer_id,
                    issue_date,
                    due_date,
                    status,
                    float(totals.subtotal),
                    float(tax_rate),
                    float(totals.tax_amount),
                    float(totals.total_amount),
                    notes.strip(),
                    now,
                    now,
                ),
            )
            invoice_id = int(cursor.lastrowid)

            for index, item in enumerate(items):
                quantity = self._to_decimal(item["quantity"])
                unit_price = self._to_decimal(item["unit_price"])
                line_total = self._money(quantity * unit_price)
                connection.execute(
                    """
                    INSERT INTO invoice_items (
                        invoice_id, description, quantity, unit_price, line_total, display_order
                    ) VALUES (?, ?, ?, ?, ?, ?)
                    """,
                    (
                        invoice_id,
                        item["description"].strip(),
                        float(quantity),
                        float(unit_price),
                        float(line_total),
                        index,
                    ),
                )

            connection.commit()

        return {
            "invoice_id": invoice_id,
            "invoice_number": invoice_number,
            "subtotal": float(totals.subtotal),
            "tax_amount": float(totals.tax_amount),
            "total_amount": float(totals.total_amount),
        }

    def list_invoices(self) -> list[dict]:
        rows = self.db.fetchall(
            """
            SELECT
                invoices.id,
                invoices.invoice_number,
                invoices.issue_date,
                invoices.due_date,
                invoices.status,
                invoices.total_amount,
                invoices.pdf_path,
                invoices.sent_at,
                customers.name AS customer_name,
                customers.email AS customer_email
            FROM invoices
            JOIN customers ON customers.id = invoices.customer_id
            ORDER BY invoices.id DESC
            """
        )
        return [dict(row) for row in rows]

    def get_invoice_details(self, invoice_id: int) -> dict | None:
        invoice_row = self.db.fetchone(
            """
            SELECT
                invoices.id,
                invoices.invoice_number,
                invoices.issue_date,
                invoices.due_date,
                invoices.status,
                invoices.subtotal,
                invoices.tax_rate,
                invoices.tax_amount,
                invoices.total_amount,
                invoices.notes,
                invoices.pdf_path,
                invoices.sent_at,
                invoices.paid_at,
                customers.id AS customer_id,
                customers.name AS customer_name,
                customers.email AS customer_email,
                customers.phone AS customer_phone,
                customers.address AS customer_address,
                customers.notes AS customer_notes
            FROM invoices
            JOIN customers ON customers.id = invoices.customer_id
            WHERE invoices.id = ?
            """,
            (invoice_id,),
        )

        if not invoice_row:
            return None

        item_rows = self.db.fetchall(
            """
            SELECT description, quantity, unit_price, line_total
            FROM invoice_items
            WHERE invoice_id = ?
            ORDER BY display_order ASC
            """,
            (invoice_id,),
        )

        details = dict(invoice_row)
        details["items"] = [dict(row) for row in item_rows]
        return details

    def save_pdf_path(self, invoice_id: int, pdf_path: str) -> None:
        now = datetime.now().isoformat(timespec="seconds")
        self.db.execute(
            "UPDATE invoices SET pdf_path = ?, updated_at = ? WHERE id = ?",
            (pdf_path, now, invoice_id),
        )

    def mark_sent(self, invoice_id: int) -> None:
        now = datetime.now().isoformat(timespec="seconds")
        self.db.execute(
            """
            UPDATE invoices
            SET status = 'sent', sent_at = ?, updated_at = ?
            WHERE id = ?
            """,
            (now, now, invoice_id),
        )

    def mark_paid(self, invoice_id: int) -> None:
        now = datetime.now().isoformat(timespec="seconds")
        self.db.execute(
            """
            UPDATE invoices
            SET status = 'paid', paid_at = ?, updated_at = ?
            WHERE id = ?
            """,
            (now, now, invoice_id),
        )

    def delete_invoice(self, invoice_id: int) -> None:
        self.db.execute("DELETE FROM invoices WHERE id = ?", (invoice_id,))
