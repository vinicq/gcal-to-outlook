"""Google Calendar client.

Handles OAuth login (opens the browser once, persists the token) and fetches
events incrementally using syncToken: after the initial full load, each call
returns only what changed (created, edited, or cancelled).

Supports both read and write operations (create, update, delete).
"""

import os
from pathlib import Path

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

SCOPES = ["https://www.googleapis.com/auth/calendar"]


class GoogleClient:
    def __init__(self, credentials_file: str, token_file: str, calendar_id: str = "primary"):
        self.credentials_file = str(Path(credentials_file).expanduser())
        self.token_file = str(Path(token_file).expanduser())
        self.calendar_id = calendar_id
        self.service = None

    def authenticate(self):
        creds = None
        if os.path.exists(self.token_file):
            creds = Credentials.from_authorized_user_file(self.token_file, SCOPES)
        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
            else:
                flow = InstalledAppFlow.from_client_secrets_file(
                    self.credentials_file, SCOPES
                )
                creds = flow.run_local_server(port=0)
            with open(self.token_file, "w", encoding="utf-8") as f:
                f.write(creds.to_json())
        self.service = build("calendar", "v3", credentials=creds, cache_discovery=False)

    def list_changes(self, sync_token: str | None, time_min: str, time_max: str):
        """Return (event_list, new_sync_token).

        If sync_token is None, performs a full load over the [time_min, time_max] window.
        If the sync_token has expired (HTTP 410), signals the caller to redo the full load
        by returning ("EXPIRED", None).
        """
        events = []
        page_token = None
        new_sync_token = None

        while True:
            params = {
                "calendarId": self.calendar_id,
                "singleEvents": True,
                "showDeleted": True,
                "maxResults": 250,
                "pageToken": page_token,
            }
            if sync_token:
                params["syncToken"] = sync_token
            else:
                params["timeMin"] = time_min
                params["timeMax"] = time_max
                params["orderBy"] = "startTime"

            try:
                resp = self.service.events().list(**params).execute()
            except HttpError as e:
                if e.resp.status == 410:
                    return "EXPIRED", None
                raise

            events.extend(resp.get("items", []))
            page_token = resp.get("nextPageToken")
            if not page_token:
                new_sync_token = resp.get("nextSyncToken")
                break

        return events, new_sync_token

    def create_event(self, payload: dict) -> str:
        """Create an event in Google Calendar and return the new event id."""
        try:
            event = self.service.events().insert(
                calendarId=self.calendar_id, body=payload
            ).execute()
        except HttpError as e:
            raise RuntimeError(f"Error creating event in Google Calendar: {e}") from e
        return event["id"]

    def update_event(self, event_id: str, payload: dict) -> None:
        """Partially update an existing event via PATCH."""
        try:
            self.service.events().patch(
                calendarId=self.calendar_id, eventId=event_id, body=payload
            ).execute()
        except HttpError as e:
            raise RuntimeError(
                f"Error updating event {event_id} in Google Calendar: {e}"
            ) from e

    def delete_event(self, event_id: str) -> None:
        """Remove an event from Google Calendar. 404 is treated as success."""
        try:
            self.service.events().delete(
                calendarId=self.calendar_id, eventId=event_id
            ).execute()
        except HttpError as e:
            if e.resp.status in (404, 410):
                return
            raise RuntimeError(
                f"Error deleting event {event_id} from Google Calendar: {e}"
            ) from e
