from datetime import datetime
from pathlib import Path

from reportlab.lib import colors
from reportlab.lib.pagesizes import LETTER
from reportlab.lib.units import inch
from reportlab.pdfgen import canvas

from app.config import DATA_DIR, OUTPUT_DIR

IMAGE_EXTENSIONS = {".png", ".jpg", ".jpeg", ".gif", ".webp"}


def format_usd(amount: float) -> str:
    return f"${amount:,.2f}"


class PDFService:
    def __init__(self, output_dir: Path | None = None) -> None:
        self.output_dir = output_dir or OUTPUT_DIR
        self.output_dir.mkdir(parents=True, exist_ok=True)

    @staticmethod
    def _resolve_image_path(explicit_path: str, keyword: str) -> str | None:
        candidate = explicit_path.strip()
        if candidate:
            path = Path(candidate)
            if path.exists() and path.is_file():
                return str(path)

        needle = keyword.lower().strip()
        if not DATA_DIR.exists():
            return None

        for path in sorted(DATA_DIR.iterdir()):
            if not path.is_file():
                continue
            if path.suffix.lower() not in IMAGE_EXTENSIONS:
                continue
            if needle in path.stem.lower():
                return str(path)

        return None

    def generate_invoice_pdf(self, invoice: dict, business: dict[str, str]) -> str:
        safe_invoice_number = invoice["invoice_number"].replace("/", "-")
        filename = f"{safe_invoice_number}.pdf"
        file_path = self.output_dir / filename

        c = canvas.Canvas(str(file_path), pagesize=LETTER)
        width, height = LETTER
        margin_x = 0.75 * inch
        current_y = height - 0.75 * inch
        title_y = current_y

        logo_path = self._resolve_image_path(business.get("business_logo_path", ""), "logo")
        qr_path = self._resolve_image_path(business.get("business_qr_path", ""), "qr")

        if logo_path:
            logo_width = 2.8 * inch
            logo_height = 1.35 * inch
            logo_y = current_y - logo_height
            c.drawImage(
                logo_path,
                margin_x * 0.4,
                logo_y + 0.3 * inch,
                width=logo_width,
                height=logo_height,
                preserveAspectRatio=True,
                mask="auto",
            )
            # Reserve vertical space so header text never overlaps the logo.
            current_y = logo_y - 0.14 * inch

        c.setFont("Helvetica-Bold", 18)
        c.drawRightString(width - margin_x, title_y, "INVOICE")
        current_y = min(current_y, title_y - 0.35 * inch)

        business_name = business.get("business_name") or ""
        c.setFont("Helvetica-Bold", 12)
        c.drawString(margin_x, current_y, business_name)

        c.setFont("Helvetica", 10)
        c.drawRightString(width - margin_x, current_y, f"Invoice #: {invoice['invoice_number']}")
        current_y -= 0.2 * inch

        if business.get("business_email"):
            c.drawString(margin_x, current_y, f"Email: {business['business_email']}")
        c.drawRightString(width - margin_x, current_y, f"Issue Date: {invoice['issue_date']}")
        current_y -= 0.18 * inch

        if business.get("business_phone"):
            c.drawString(margin_x, current_y, f"Phone: {business['business_phone']}")
        c.drawRightString(width - margin_x, current_y, f"Due Date: {invoice['due_date']}")
        current_y -= 0.2 * inch

        business_address = business.get("business_address", "").strip()
        if business_address:
            for line in business_address.splitlines()[:3]:
                c.drawString(margin_x, current_y, line)
                current_y -= 0.16 * inch

        current_y -= 0.15 * inch
        c.setFont("Helvetica-Bold", 11)
        c.drawString(margin_x, current_y, "Bill To")
        current_y -= 0.18 * inch

        c.setFont("Helvetica", 10)
        c.drawString(margin_x, current_y, invoice["customer_name"])
        current_y -= 0.16 * inch

        customer_email = invoice.get("customer_email", "")
        if customer_email:
            c.drawString(margin_x, current_y, customer_email)
            current_y -= 0.16 * inch

        customer_address = (invoice.get("customer_address") or "").strip()
        for line in customer_address.splitlines()[:4]:
            c.drawString(margin_x, current_y, line)
            current_y -= 0.16 * inch

        current_y -= 0.2 * inch

        table_left = margin_x
        table_right = width - margin_x
        table_width = table_right - table_left

        col_description = table_left
        col_rate = table_left + table_width * 0.68
        col_total = table_left + table_width * 0.82

        c.setFillColor(colors.lightgrey)
        c.rect(table_left, current_y - 0.2 * inch, table_width, 0.22 * inch, fill=1, stroke=0)
        c.setFillColor(colors.black)

        c.setFont("Helvetica-Bold", 10)
        c.drawString(col_description + 2, current_y - 0.13 * inch, "Service Description")
        c.drawString(col_rate + 40, current_y - 0.13 * inch, "Rate")
        c.drawString(col_total + 35, current_y - 0.13 * inch, "Line Total")

        current_y -= 0.26 * inch
        c.setFont("Helvetica", 10)

        for item in invoice["items"]:
            if current_y < 1.6 * inch:
                c.showPage()
                current_y = height - 1.0 * inch
                c.setFont("Helvetica", 10)

            description = str(item["description"])[:58]
            c.drawString(col_description + 2, current_y - 0.12 * inch, description)
            c.drawRightString(col_total - 8, current_y - 0.12 * inch, format_usd(float(item["unit_price"])))
            c.drawRightString(table_right - 8, current_y - 0.12 * inch, format_usd(float(item["line_total"])))
            current_y -= 0.19 * inch

        current_y -= 0.12 * inch
        c.line(table_left, current_y, table_right, current_y)

        current_y -= 0.24 * inch
        totals_left = table_right - 2.35 * inch
        totals_value_x = table_right - 8
        c.setFont("Helvetica", 10)
        c.drawString(totals_left, current_y, "Subtotal")
        c.drawRightString(totals_value_x, current_y, format_usd(float(invoice["subtotal"])))

        current_y -= 0.18 * inch
        c.drawString(totals_left, current_y, f"Tax ({float(invoice['tax_rate']):.2f}%)")
        c.drawRightString(totals_value_x, current_y, format_usd(float(invoice["tax_amount"])))

        current_y -= 0.2 * inch
        c.setFont("Helvetica-Bold", 11)
        c.drawString(totals_left, current_y, "Total Due")
        c.drawRightString(totals_value_x, current_y, format_usd(float(invoice["total_amount"])))

        current_y -= 0.3 * inch
        c.setFont("Helvetica", 10)
        payment_instructions = business.get("payment_instructions", "").strip()
        if payment_instructions:
            c.setFont("Helvetica-Bold", 10)
            c.drawString(table_left, current_y, "Payment Instructions")
            current_y -= 0.18 * inch
            c.setFont("Helvetica", 10)
            for line in payment_instructions.splitlines()[:4]:
                c.drawString(table_left, current_y, line)
                current_y -= 0.16 * inch

        if invoice.get("notes"):
            current_y -= 0.12 * inch
            c.setFont("Helvetica-Bold", 10)
            c.drawString(table_left, current_y, "Notes")
            current_y -= 0.18 * inch
            c.setFont("Helvetica", 10)
            for line in str(invoice["notes"]).splitlines()[:4]:
                c.drawString(table_left, current_y, line)
                current_y -= 0.16 * inch

        if qr_path:
            qr_size = 1.75 * inch
            min_bottom_margin = 0.7 * inch
            required_height = qr_size + 0.28 * inch
            if current_y - required_height < min_bottom_margin:
                c.showPage()
                current_y = height - 1.0 * inch

            qr_x = (width - qr_size) / 2
            qr_y = current_y - qr_size
            c.drawImage(
                qr_path,
                qr_x,
                qr_y,
                width=qr_size,
                height=qr_size,
                preserveAspectRatio=True,
                mask="auto",
            )
            c.setFont("Helvetica-Bold", 9)
            c.drawCentredString(width / 2, qr_y - 0.12 * inch, "Scan to pay")

        c.save()
        return str(file_path)
