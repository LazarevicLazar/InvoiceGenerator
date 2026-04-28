import smtplib
from email.message import EmailMessage


DEFAULT_SMS_GATEWAYS = (
    "vtext.com",
    "txt.att.net",
    "tmomail.net",
    "messaging.sprintpcs.com",
    "mms.att.net",
)


class SMSService:
    @staticmethod
    def _normalize_phone(phone_number: str) -> str:
        digits = "".join(ch for ch in phone_number if ch.isdigit())
        if len(digits) == 11 and digits.startswith("1"):
            digits = digits[1:]

        if len(digits) != 10:
            raise ValueError("Customer phone must contain 10 US digits for SMS gateway delivery.")

        return digits

    def send_invoice_text(
        self,
        settings: dict[str, str],
        recipient_phone: str,
        customer_name: str,
        invoice_number: str,
        total_amount: float,
        due_date: str,
    ) -> list[str]:
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

        phone_digits = self._normalize_phone(recipient_phone)
        body = (
            f"Hi {customer_name}, invoice {invoice_number} is ready. "
            f"Total ${total_amount:,.2f}. Due {due_date}."
        )

        delivered_gateways: list[str] = []
        delivery_errors: list[str] = []

        with smtplib.SMTP(smtp_server, smtp_port, timeout=30) as smtp:
            if use_tls:
                smtp.starttls()
            if smtp_username:
                smtp.login(smtp_username, smtp_password)

            for gateway in DEFAULT_SMS_GATEWAYS:
                sms_address = f"{phone_digits}@{gateway}"
                message = EmailMessage()
                message["From"] = from_email
                message["To"] = sms_address
                message["Subject"] = f"Invoice {invoice_number}"
                message.set_content(body)
                try:
                    smtp.send_message(message)
                    delivered_gateways.append(gateway)
                except Exception as exc:
                    delivery_errors.append(f"{gateway}: {exc}")

        if not delivered_gateways:
            joined_errors = "; ".join(delivery_errors[:3])
            raise ValueError(
                "Failed to deliver SMS through known carrier gateways. "
                f"Sample errors: {joined_errors}"
            )

        return delivered_gateways
