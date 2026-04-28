import smtplib
from email.message import EmailMessage
from pathlib import Path


class EmailService:
    def send_invoice_email(
        self,
        settings: dict[str, str],
        recipient_email: str,
        customer_name: str,
        invoice_number: str,
        pdf_path: str,
        total_amount: float,
        due_date: str,
    ) -> None:
        smtp_server = settings.get("smtp_server", "").strip()
        smtp_port = int(settings.get("smtp_port", "587") or "587")
        smtp_username = settings.get("smtp_username", "").strip()
        smtp_password = settings.get("smtp_password", "")
        from_email = settings.get("smtp_from_email", "").strip() or settings.get("business_email", "").strip()
        use_tls = settings.get("smtp_use_tls", "1") == "1"

        if not smtp_server:
            raise ValueError("SMTP server is required in Settings.")
        if not from_email:
            raise ValueError("Sender email is required in Settings.")

        pdf_file = Path(pdf_path)
        if not pdf_file.exists():
            raise FileNotFoundError(f"Invoice PDF not found: {pdf_path}")

        business_name = settings.get("business_name", "Your Cleaning Business").strip() or "Your Cleaning Business"

        message = EmailMessage()
        message["From"] = from_email
        message["To"] = recipient_email
        message["Subject"] = f"Invoice {invoice_number} from {business_name}"

        message_body = (
            f"Hi {customer_name},\n\n"
            f"Please find attached invoice {invoice_number}.\n"
            f"Total Due: ${total_amount:,.2f}\n"
            f"Due Date: {due_date}\n\n"
            f"Thank you for choosing {business_name}.\n"
        )
        message.set_content(message_body)

        pdf_bytes = pdf_file.read_bytes()
        message.add_attachment(
            pdf_bytes,
            maintype="application",
            subtype="pdf",
            filename=pdf_file.name,
        )

        with smtplib.SMTP(smtp_server, smtp_port, timeout=30) as smtp:
            if use_tls:
                smtp.starttls()
            if smtp_username:
                smtp.login(smtp_username, smtp_password)
            smtp.send_message(message)
