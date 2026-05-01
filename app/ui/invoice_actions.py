from __future__ import annotations

from pathlib import Path


def generate_pdf(app, invoice_id: int) -> str:
    details = app.invoice_service.get_invoice_details(invoice_id)
    if not details:
        raise ValueError("Invoice not found")

    settings = app.settings_service.get_settings()
    pdf_path = app.pdf_service.generate_invoice_pdf(details, settings)
    app.invoice_service.save_pdf_path(invoice_id, pdf_path)
    return pdf_path


def send_invoice(app, invoice_id: int) -> None:
    details = app.invoice_service.get_invoice_details(invoice_id)
    if not details:
        raise ValueError("Invoice not found")

    customer_email = (details.get("customer_email") or "").strip()
    if not customer_email:
        raise ValueError("Customer email is missing.")

    pdf_path = details.get("pdf_path", "").strip()
    if not pdf_path or not Path(pdf_path).exists():
        pdf_path = generate_pdf(app, invoice_id)

    settings = app.settings_service.get_settings()
    app.email_service.send_invoice_email(
        settings=settings,
        recipient_email=customer_email,
        customer_name=details.get("customer_name", "Customer"),
        invoice_number=details["invoice_number"],
        pdf_path=pdf_path,
        total_amount=float(details["total_amount"]),
        due_date=details["due_date"],
    )
    app.invoice_service.mark_sent(invoice_id)


def send_invoice_sms(app, invoice_id: int) -> list[str]:
    details = app.invoice_service.get_invoice_details(invoice_id)
    if not details:
        raise ValueError("Invoice not found")

    customer_phone = (details.get("customer_phone") or "").strip()
    if not customer_phone:
        raise ValueError("Customer phone is required. Update it in the Customers tab first.")

    settings = app.settings_service.get_settings()
    delivered_gateways = app.sms_service.send_invoice_text(
        settings=settings,
        recipient_phone=customer_phone,
        customer_name=details.get("customer_name", "Customer"),
        invoice_number=details["invoice_number"],
        total_amount=float(details["total_amount"]),
        due_date=details["due_date"],
    )
    app.invoice_service.mark_sent(invoice_id)
    return delivered_gateways