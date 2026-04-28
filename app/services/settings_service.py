from app.config import DEFAULT_SETTINGS
from app.database import DatabaseManager


class SettingsService:
    def __init__(self, db: DatabaseManager) -> None:
        self.db = db

    def get_settings(self) -> dict[str, str]:
        rows = self.db.fetchall("SELECT key, value FROM app_settings")
        settings = {row["key"]: row["value"] for row in rows}

        for key, value in DEFAULT_SETTINGS.items():
            settings.setdefault(key, value)

        return settings

    def save_settings(self, updates: dict[str, str]) -> None:
        with self.db.get_connection() as connection:
            for key, value in updates.items():
                connection.execute(
                    """
                    INSERT INTO app_settings (key, value)
                    VALUES (?, ?)
                    ON CONFLICT(key) DO UPDATE SET value = excluded.value
                    """,
                    (key, value),
                )
            connection.commit()
