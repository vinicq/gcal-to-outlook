"""Microsoft Graph client.

Authenticates via delegated OAuth (MSAL) with a persistent on-disk token cache,
so the browser login is only required once. Creates, updates, and deletes events
in the Outlook/M365 calendar (the same calendar Teams displays).

Supports delta queries for incremental change reads (MS Graph delta sync).
"""

import os
import time
from pathlib import Path

import msal
import requests

GRAPH = "https://graph.microsoft.com/v1.0"
SCOPES = ["Calendars.ReadWrite"]

DELTA_SELECT = "id,subject,start,end,body,location,isCancelled,isAllDay,attendees,organizer"


class MicrosoftClient:
    def __init__(self, client_id: str, tenant_id: str, token_cache_file: str,
                 calendar_id: str = ""):
        self.client_id = client_id
        self.tenant_id = tenant_id
        self.token_cache_file = str(Path(token_cache_file).expanduser())
        self.calendar_id = calendar_id  # empty = default calendar
        self._token = None

    def _build_app(self):
        cache = msal.SerializableTokenCache()
        if os.path.exists(self.token_cache_file):
            cache.deserialize(open(self.token_cache_file, encoding="utf-8").read())
        app = msal.PublicClientApplication(
            self.client_id,
            authority=f"https://login.microsoftonline.com/{self.tenant_id}",
            token_cache=cache,
        )
        return app, cache

    def _save_cache(self, cache):
        if cache.has_state_changed:
            with open(self.token_cache_file, "w", encoding="utf-8") as f:
                f.write(cache.serialize())

    def authenticate(self):
        app, cache = self._build_app()
        result = None
        accounts = app.get_accounts()
        if accounts:
            result = app.acquire_token_silent(SCOPES, account=accounts[0])
        if not result:
            # Opens the browser and captures the token via a temporary local server
            # (same experience as Google login).
            # Requires http://localhost redirect registered in the Azure app.
            print("\n=== MICROSOFT LOGIN ===")
            print("Opening browser for login...")
            print("=======================\n")
            result = app.acquire_token_interactive(scopes=SCOPES)
        if "access_token" not in result:
            raise RuntimeError(f"Microsoft login failed: {result.get('error_description')}")
        self._token = result["access_token"]
        self._save_cache(cache)

    def _refresh_if_needed(self):
        app, cache = self._build_app()
        accounts = app.get_accounts()
        if accounts:
            result = app.acquire_token_silent(SCOPES, account=accounts[0])
            if result and "access_token" in result:
                self._token = result["access_token"]
                self._save_cache(cache)

    def _headers(self):
        return {
            "Authorization": f"Bearer {self._token}",
            "Content-Type": "application/json",
        }

    def _events_base(self):
        if self.calendar_id:
            return f"{GRAPH}/me/calendars/{self.calendar_id}/events"
        return f"{GRAPH}/me/events"

    def _request(self, method, url, **kwargs):
        for attempt in range(4):  # up to 3 retries after the first failure
            resp = requests.request(method, url, headers=self._headers(), timeout=30, **kwargs)

            if resp.status_code == 401:
                self._refresh_if_needed()
                resp = requests.request(method, url, headers=self._headers(), timeout=30, **kwargs)

            if resp.status_code == 429:
                # Rate limited: honor the Retry-After value returned by the API.
                if attempt == 3:
                    return resp
                wait = int(resp.headers.get("Retry-After", 2 ** attempt))
                time.sleep(wait)
            elif resp.status_code == 503:
                # Service unavailable: exponential backoff (1s, 2s, 4s).
                if attempt == 3:
                    return resp
                time.sleep(2 ** attempt)
            else:
                return resp

        return resp

    def create_event(self, payload: dict) -> str | None:
        resp = self._request("POST", self._events_base(), json=payload)
        if resp.status_code in (200, 201):
            return resp.json().get("id")
        raise RuntimeError(f"Error creating event ({resp.status_code}): {resp.text}")

    def update_event(self, ms_id: str, payload: dict):
        url = f"{self._events_base()}/{ms_id}"
        resp = self._request("PATCH", url, json=payload)
        if resp.status_code not in (200, 201):
            raise RuntimeError(f"Error updating event ({resp.status_code}): {resp.text}")

    def delete_event(self, ms_id: str):
        url = f"{self._events_base()}/{ms_id}"
        resp = self._request("DELETE", url)
        # 404 = already gone; treat as success.
        if resp.status_code not in (200, 204, 404):
            raise RuntimeError(f"Error deleting event ({resp.status_code}): {resp.text}")

    def find_by_sync_id(self, google_id: str) -> str | None:
        """Return the id of an event previously created for this Google id.

        Searches event bodies for the '[sync-id: <google_id>]' marker so the
        sync can adopt its own events after a local DB loss instead of creating
        duplicates. Uses Graph $search, which covers the body field.
        Returns None on no match or if the search is unavailable.
        """
        marker = f"[sync-id: {google_id}]"
        try:
            resp = self._request(
                "GET", self._events_base(),
                params={"$search": f'"{marker}"', "$select": "id,body", "$top": "25"},
            )
            if resp.status_code != 200:
                return None
            for ev in resp.json().get("value", []):
                if marker in (ev.get("body", {}).get("content", "") or ""):
                    return ev.get("id")
        except Exception:
            pass
        return None

    def list_changes(
        self,
        delta_token: str | None = None,
        time_min: str | None = None,
        time_max: str | None = None,
    ) -> tuple:
        """Incremental read via MS Graph calendarView/delta.

        Parameters
        ----------
        delta_token:
            Full deltaLink returned by a previous call.
            If None, performs an initial load over the [time_min, time_max] window.
        time_min / time_max:
            ISO-8601 bounds for the sync window. Required for the initial load
            (delta_token=None). The calendarView/delta endpoint requires
            startDateTime and endDateTime; /events/delta silently ignored those
            filters.

        Returns
        -------
        ('EXPIRED', None)
            The delta_token has expired (HTTP 410). The caller must discard the
            token and redo the initial load.
        (event_list, new_delta_link)
            Deleted events carry '@removed': {'reason': 'deleted'}.
            new_delta_link is the opaque URL for the next incremental call.
        """
        if delta_token:
            url = delta_token
            params = {}
        else:
            # calendarView/delta correctly handles startDateTime/endDateTime;
            # /events/delta silently ignored $filter (endpoint bug).
            if self.calendar_id:
                url = f"{GRAPH}/me/calendars/{self.calendar_id}/calendarView/delta"
            else:
                url = f"{GRAPH}/me/calendarView/delta"
            params: dict = {"$select": DELTA_SELECT}
            if time_min:
                params["startDateTime"] = time_min
            if time_max:
                params["endDateTime"] = time_max

        events: list = []
        delta_link: str | None = None

        while url:
            resp = self._request("GET", url, params=params if params else None)

            # 410 Gone: token expired, caller must redo full load.
            if resp.status_code == 410:
                return ("EXPIRED", None)

            if resp.status_code not in (200,):
                raise RuntimeError(
                    f"Delta query error ({resp.status_code}): {resp.text}"
                )

            data = resp.json()
            events.extend(data.get("value", []))

            # After the first page, pagination params are embedded in nextLink.
            params = {}

            next_link = data.get("@odata.nextLink")
            delta_link_raw = data.get("@odata.deltaLink")

            if next_link:
                url = next_link
            elif delta_link_raw:
                delta_link = delta_link_raw
                url = None
            else:
                # Response has neither nextLink nor deltaLink: unexpected end of pagination.
                url = None

        return (events, delta_link)
