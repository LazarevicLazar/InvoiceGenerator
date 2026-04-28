from datetime import datetime

from app.database import DatabaseManager

ACTIVE_JOB_STATUSES = ("scheduled", "in-progress")


class CalendarService:
    def __init__(self, db: DatabaseManager) -> None:
        self.db = db

    def list_cleaners(self) -> list[dict]:
        rows = self.db.fetchall(
            """
            SELECT id, name, phone, email, google_calendar_id, notes, active
            FROM cleaners
            WHERE active = 1
            ORDER BY name COLLATE NOCASE ASC
            """
        )
        return [dict(row) for row in rows]

    def get_cleaner(self, cleaner_id: int) -> dict | None:
        row = self.db.fetchone(
            """
            SELECT id, name, phone, email, google_calendar_id, notes, active
            FROM cleaners
            WHERE id = ?
            """,
            (cleaner_id,),
        )
        return dict(row) if row else None

    def create_cleaner(
        self,
        name: str,
        phone: str = "",
        email: str = "",
        notes: str = "",
        google_calendar_id: str = "",
    ) -> int:
        now = datetime.now().isoformat(timespec="seconds")
        return self.db.execute(
            """
            INSERT INTO cleaners (
                name, phone, email, google_calendar_id, notes, active, created_at, updated_at
            )
            VALUES (?, ?, ?, ?, ?, 1, ?, ?)
            """,
            (
                name.strip(),
                phone.strip(),
                email.strip(),
                google_calendar_id.strip(),
                notes.strip(),
                now,
                now,
            ),
        )

    def update_cleaner(
        self,
        cleaner_id: int,
        name: str,
        phone: str = "",
        email: str = "",
        notes: str = "",
        google_calendar_id: str = "",
    ) -> None:
        now = datetime.now().isoformat(timespec="seconds")
        self.db.execute(
            """
            UPDATE cleaners
            SET name = ?, phone = ?, email = ?, google_calendar_id = ?, notes = ?, updated_at = ?
            WHERE id = ?
            """,
            (
                name.strip(),
                phone.strip(),
                email.strip(),
                google_calendar_id.strip(),
                notes.strip(),
                now,
                cleaner_id,
            ),
        )

    def delete_cleaner(self, cleaner_id: int) -> None:
        active_job_count = self.count_cleaner_active_jobs(cleaner_id)
        if active_job_count > 0:
            raise ValueError(
                "Cannot delete cleaner with active scheduled jobs. "
                f"Active jobs found: {active_job_count}."
            )

        # Soft delete to preserve historical job records.
        self.db.execute("UPDATE cleaners SET active = 0 WHERE id = ?", (cleaner_id,))

    def count_cleaner_active_jobs(self, cleaner_id: int) -> int:
        placeholders = ",".join("?" for _ in ACTIVE_JOB_STATUSES)
        row = self.db.fetchone(
            f"""
            SELECT COUNT(*) AS active_job_count
            FROM jobs
            WHERE cleaner_id = ?
              AND status IN ({placeholders})
            """,
            (cleaner_id, *ACTIVE_JOB_STATUSES),
        )
        return int(row["active_job_count"]) if row else 0

    def list_jobs(self) -> list[dict]:
        rows = self.db.fetchall(
            """
            SELECT
                jobs.id,
                jobs.customer_id,
                jobs.cleaner_id,
                jobs.title,
                jobs.start_at,
                jobs.end_at,
                jobs.location,
                jobs.notes,
                jobs.status,
                jobs.google_event_id,
                customers.name AS customer_name,
                cleaners.name AS cleaner_name
            FROM jobs
            JOIN customers ON customers.id = jobs.customer_id
            JOIN cleaners ON cleaners.id = jobs.cleaner_id
            ORDER BY jobs.start_at ASC
            """
        )
        return [dict(row) for row in rows]

    def check_cleaner_availability(
        self,
        cleaner_id: int,
        start_at: str,
        end_at: str,
        exclude_job_id: int | None = None,
    ) -> list[dict]:
        params: list = [cleaner_id, *ACTIVE_JOB_STATUSES, end_at, start_at]
        exclude_sql = ""
        if exclude_job_id is not None:
            exclude_sql = "AND id != ?"
            params.append(exclude_job_id)

        placeholders = ",".join("?" for _ in ACTIVE_JOB_STATUSES)
        rows = self.db.fetchall(
            f"""
            SELECT id, title, start_at, end_at, status
            FROM jobs
            WHERE cleaner_id = ?
              AND status IN ({placeholders})
              AND start_at < ?
              AND end_at > ?
              {exclude_sql}
            ORDER BY start_at ASC
            """,
            tuple(params),
        )
        return [dict(row) for row in rows]

    def create_job(
        self,
        customer_id: int,
        cleaner_id: int,
        title: str,
        start_at: str,
        end_at: str,
        location: str = "",
        notes: str = "",
    ) -> int:
        conflicts = self.check_cleaner_availability(cleaner_id, start_at, end_at)
        if conflicts:
            raise ValueError("Cleaner is not available in the selected time range.")

        now = datetime.now().isoformat(timespec="seconds")
        return self.db.execute(
            """
            INSERT INTO jobs (
                customer_id, cleaner_id, title, start_at, end_at,
                location, notes, status, google_event_id, created_at, updated_at
            ) VALUES (?, ?, ?, ?, ?, ?, ?, 'scheduled', '', ?, ?)
            """,
            (
                customer_id,
                cleaner_id,
                title.strip(),
                start_at,
                end_at,
                location.strip(),
                notes.strip(),
                now,
                now,
            ),
        )

    def set_job_google_event_id(self, job_id: int, google_event_id: str) -> None:
        now = datetime.now().isoformat(timespec="seconds")
        self.db.execute(
            "UPDATE jobs SET google_event_id = ?, updated_at = ? WHERE id = ?",
            (google_event_id.strip(), now, job_id),
        )

    def get_job(self, job_id: int) -> dict | None:
        row = self.db.fetchone(
            """
            SELECT
                jobs.id,
                jobs.customer_id,
                jobs.cleaner_id,
                jobs.title,
                jobs.start_at,
                jobs.end_at,
                jobs.location,
                jobs.notes,
                jobs.status,
                jobs.google_event_id,
                cleaners.google_calendar_id,
                customers.name AS customer_name,
                cleaners.name AS cleaner_name
            FROM jobs
            JOIN customers ON customers.id = jobs.customer_id
            JOIN cleaners ON cleaners.id = jobs.cleaner_id
            WHERE jobs.id = ?
            """,
            (job_id,),
        )
        return dict(row) if row else None

    def update_job_status(self, job_id: int, status: str) -> None:
        now = datetime.now().isoformat(timespec="seconds")
        self.db.execute(
            "UPDATE jobs SET status = ?, updated_at = ? WHERE id = ?",
            (status, now, job_id),
        )

    def delete_job(self, job_id: int) -> None:
        self.db.execute("DELETE FROM jobs WHERE id = ?", (job_id,))
