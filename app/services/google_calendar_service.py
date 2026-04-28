from __future__ import annotations

import json
from datetime import datetime
from pathlib import Path
from typing import Any


class GoogleCalendarService:
    def __init__(self, credentials_file: str, token_file: str) -> None:
        self.credentials_file = credentials_file
        self.token_file = token_file
        self._service = None

    def is_configured(self) -> bool:
        return Path(self.credentials_file).exists()

    def connect(self) -> None:
        self._service = self._build_service(interactive=True)

    def list_calendars(self) -> list[dict[str, str]]:
        service = self._build_service(interactive=False)
        result = service.calendarList().list().execute()
        calendars = result.get("items", [])
        return [
            {
                "id": item.get("id", ""),
                "summary": item.get("summary", ""),
            }
            for item in calendars
        ]

    def check_time_busy(
        self,
        calendar_id: str,
        start_at: str,
        end_at: str,
        cleaner_id: int | None = None,
    ) -> bool:
        service = self._build_service(interactive=False)
        result = (
            service.events()
            .list(
                calendarId=calendar_id,
                timeMin=self._to_rfc3339(start_at),
                timeMax=self._to_rfc3339(end_at),
                singleEvents=True,
                orderBy="startTime",
                maxResults=250,
            )
            .execute()
        )
        events = result.get("items", [])

        for event in events:
            if event.get("status") == "cancelled":
                continue
            if event.get("transparency") == "transparent":
                continue
            if self._event_blocks_cleaner(event, cleaner_id):
                return True

        return False

    def create_event(
        self,
        calendar_id: str,
        title: str,
        start_at: str,
        end_at: str,
        location: str = "",
        description: str = "",
        cleaner_id: int | None = None,
    ) -> str:
        service = self._build_service(interactive=False)
        private_props: dict[str, str] = {
            "source_app": "invoice_generator",
        }
        if cleaner_id is not None:
            private_props["cleaner_id"] = str(cleaner_id)

        event_description = description.strip()
        if cleaner_id is not None:
            marker = f"Cleaner ID: {cleaner_id}"
            if marker not in event_description:
                event_description = f"{event_description}\n\n{marker}".strip()

        event_body: dict[str, Any] = {
            "summary": title,
            "location": location,
            "description": event_description,
            "start": {"dateTime": self._to_rfc3339(start_at)},
            "end": {"dateTime": self._to_rfc3339(end_at)},
            "extendedProperties": {
                "private": private_props,
            },
        }
        event = service.events().insert(calendarId=calendar_id, body=event_body).execute()
        return event.get("id", "")

    def create_calendar(self, summary: str, description: str = "") -> dict[str, str]:
        service = self._build_service(interactive=False)
        body: dict[str, Any] = {
            "summary": summary.strip(),
        }
        if description.strip():
            body["description"] = description.strip()

        calendar = service.calendars().insert(body=body).execute()
        return {
            "id": calendar.get("id", ""),
            "summary": calendar.get("summary", summary.strip()),
        }

    def delete_event(self, calendar_id: str, event_id: str) -> None:
        service = self._build_service(interactive=False)
        service.events().delete(calendarId=calendar_id, eventId=event_id).execute()

    def delete_calendar(self, calendar_id: str) -> None:
        service = self._build_service(interactive=False)
        service.calendars().delete(calendarId=calendar_id).execute()

    def _build_service(self, interactive: bool):
        if self._service is not None:
            return self._service

        credentials = self._load_credentials(interactive=interactive)

        try:
            from googleapiclient.discovery import build
        except ImportError as exc:
            raise RuntimeError(
                "Google Calendar packages are not installed. "
                "Install dependencies from requirements.txt."
            ) from exc

        self._service = build("calendar", "v3", credentials=credentials)
        return self._service

    def _load_credentials(self, interactive: bool):
        scopes = ["https://www.googleapis.com/auth/calendar"]

        try:
            from google.auth.transport.requests import Request
            from google.oauth2.credentials import Credentials
            from google_auth_oauthlib.flow import InstalledAppFlow
        except ImportError as exc:
            raise RuntimeError(
                "Google Calendar packages are not installed. "
                "Install dependencies from requirements.txt."
            ) from exc

        creds = None
        token_path = Path(self.token_file)
        if token_path.exists():
            try:
                # Treat empty/corrupt token files as missing so connect flow can recover.
                if token_path.stat().st_size > 0:
                    creds = Credentials.from_authorized_user_file(str(token_path), scopes)
            except (ValueError, json.JSONDecodeError):
                creds = None

        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        elif not creds or not creds.valid:
            credentials_path = Path(self.credentials_file)
            if not credentials_path.exists():
                raise RuntimeError(
                    "Google credentials file not found. "
                    f"Expected at: {self.credentials_file}"
                )
            if not interactive:
                raise RuntimeError(
                    "Google Calendar is not connected yet. Click Connect Google first."
                )

            flow = InstalledAppFlow.from_client_secrets_file(str(credentials_path), scopes)
            creds = flow.run_local_server(port=0)

        token_path.parent.mkdir(parents=True, exist_ok=True)
        token_path.write_text(creds.to_json(), encoding="utf-8")
        return creds

    @staticmethod
    def _to_rfc3339(value: str) -> str:
        local_tz = datetime.now().astimezone().tzinfo
        parsed = datetime.strptime(value, "%Y-%m-%d %H:%M")
        aware = parsed.replace(tzinfo=local_tz)
        return aware.isoformat()

    @staticmethod
    def _event_blocks_cleaner(event: dict[str, Any], cleaner_id: int | None) -> bool:
        private_props = event.get("extendedProperties", {}).get("private", {})
        source_app = str(private_props.get("source_app", "")).strip()
        event_cleaner_id = str(private_props.get("cleaner_id", "")).strip()

        # Events created by this app should only block the matching cleaner.
        if source_app == "invoice_generator":
            if cleaner_id is None:
                return True
            return event_cleaner_id == str(cleaner_id)

        # Backward compatibility for older app-generated events.
        description = str(event.get("description", ""))
        if "Scheduled from Cleaning Invoice Generator" in description:
            if cleaner_id is None:
                return True
            marker = f"Cleaner ID: {cleaner_id}"
            if marker in description:
                return True
            # Legacy app events without cleaner markers should not block other cleaners
            # when several cleaners share one Google calendar.
            return False

        # External calendar events continue to block bookings.
        return True
