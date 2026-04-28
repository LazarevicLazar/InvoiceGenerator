# Cleaning Invoice Generator (Python Desktop App)

Single-user desktop invoice app for a service business (cleaning-focused):

- Customer management
- Service invoice creation in USD
- PDF invoice generation
- Email invoice sending
- Free text reminders via carrier email-to-SMS gateways
- Calendar tab for cleaner scheduling, availability checks, and Google Calendar sync
- Invoice history with draft/sent/paid status

## Tech Stack

- Python 3.11+
- Tkinter (desktop UI)
- SQLite (local storage)
- ReportLab (PDF generation)
- smtplib (email sending)
- Google Calendar API (optional integration)

## Quick Start

1. Create and activate a virtual environment.
2. Install dependencies.
3. Run the app.

```bash
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
python run.py
```

## First-Time Setup

1. Open the **Settings** tab.
2. Fill in:
   - Business Name, Email, Phone, Address
   - Optional logo image path
   - Default Tax Rate
   - SMTP details for your email provider
3. Save settings.

## Free Text Message Option (Carrier Email Gateway)

You can send invoice text reminders for free using carrier email-to-SMS gateways.
This uses the same SMTP settings you already configured for email.

1. In **Customers**, add:
   - Phone (US number)
2. The app automatically tries these gateway domains on send:
   - `vtext.com`
   - `txt.att.net`
   - `tmomail.net`
   - `messaging.sprintpcs.com`
   - `mms.att.net`
3. Send text from either:
   - **Create Invoice** -> `Save + Text`
   - **Invoice History** -> `Send Text`

Notes:

- Delivery is best-effort and can be filtered by carriers.
- Text messages do not carry PDF attachments.

## Google Calendar Integration (Optional)

You can connect cleaner schedules with Google Calendar.

1. In Google Cloud Console:
   - Enable the **Google Calendar API**.
   - Create an **OAuth Client ID** for a Desktop app.
   - Download the credentials JSON file.
2. Save that JSON file in your project, for example:
   - `data/google_client_secret.json`
3. In **Settings**, set:
   - `Google Credentials File` -> path to credentials JSON
   - `Google Token File` -> where auth token will be stored (default: `data/google_token.json`)
4. In **Calendar** tab:
   - Click `Connect Google` and complete the OAuth browser flow.
   - Click `Load Calendars`.
   - Select a calendar and click `Use On Cleaner`.
   - Save or update the cleaner.

Behavior:

- Availability checks can verify both local jobs and Google busy times.
- Scheduling a job can create a Google Calendar event for that cleaner.
- Deleting or cancelling a job removes the linked Google event.
- If Google sync fails, local scheduling still works and you get a warning.

## Daily Workflow

1. Add customers in **Customers** tab.
2. Add cleaners and schedule jobs in **Calendar** tab:
   - Create cleaner profiles
   - (Optional) assign each cleaner a Google Calendar ID
   - Pick customer + cleaner + start/end time
   - Check cleaner availability
   - Schedule and manage job status
3. Create invoice in **Create Invoice** tab:
   - Select customer
   - Add service line items
   - Save draft, or save+PDF, or save+PDF+email
4. Track invoices in **Invoice History** tab:
   - Generate PDF
   - Send email
   - Send text
   - Mark paid
   - Open generated PDF

## Data and Output

- SQLite database: `data/invoice_generator.db`
- Generated PDFs: `output/`

## Notes

- Currency is fixed to USD.
- No login/auth (single-user app).
- No payment gateway integration.
- No cloud/mobile sync.
