from app.config import ensure_runtime_directories
from app.database import DatabaseManager
from app.ui.app_window_new import InvoiceGeneratorApp

def main() -> None:
    ensure_runtime_directories()

    db = DatabaseManager()
    db.initialize()

    app = InvoiceGeneratorApp(db)
    try:
        app.iconbitmap("data/icon.ico")
    except Exception:
        pass
    app.mainloop()
