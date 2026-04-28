from pathlib import Path

APP_NAME = "Cleaning Invoice Generator"
BASE_DIR = Path(__file__).resolve().parent.parent
DATA_DIR = BASE_DIR / "data"
OUTPUT_DIR = BASE_DIR / "output"
DB_PATH = DATA_DIR / "invoice_generator.db"
DATE_FORMAT = "%Y-%m-%d"
DEFAULT_DUE_DAYS = 14

DEFAULT_SETTINGS = {
    "business_name": "",
    "business_email": "",
    "business_phone": "",
    "business_address": "",
    "business_logo_path": "",
    "payment_instructions": "",
    "default_tax_rate": "0",
    "smtp_server": "smtp.gmail.com",
    "smtp_port": "587",
    "smtp_username": "",
    "smtp_password": "",
    "smtp_from_email": "",
    "smtp_use_tls": "1",
    "google_credentials_file": str(DATA_DIR / "google_client_secret.json"),
    "google_token_file": str(DATA_DIR / "google_token.json"),
}


def ensure_runtime_directories() -> None:
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
