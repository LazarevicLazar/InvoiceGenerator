from app.config import ensure_runtime_directories
from app.database import DatabaseManager
from app.ui.app_window import InvoiceGeneratorApp

def main() -> None:
    ensure_runtime_directories()

    db = DatabaseManager()
    db.initialize()

    app = InvoiceGeneratorApp(db)
    app.mainloop()
