"""Outlook client via COM automation.

Accesses the Outlook desktop calendar directly, without the Microsoft Graph API
and without requiring admin consent.

Requires:
  - Outlook desktop installed (part of Microsoft 365 / Office)
  - Organizational account signed in to Outlook
  - pywin32  (pip install pywin32)

Interface identical to MicrosoftClient for transparent substitution.
"""

from datetime import datetime
from datetime import timezone as stdlib_tz

import win32com.client

OL_APPOINTMENT_ITEM = 1
OL_FOLDER_CALENDAR = 9


class OutlookComClient:
    """Creates, updates, and deletes events in Outlook via COM.

    Outlook automatically syncs with the Exchange/M365 server,
    making events appear in Teams without any API calls.

    Args:
        outlook_account: email of the Exchange account to use (e.g. 'user@org.com').
                         If None, uses the first Exchange mailbox found.
    """

    def __init__(self, outlook_account: str | None = None):
        self._app = None
        self._ns = None
        self._calendar = None
        self._account_filter = (outlook_account or "").strip().lower()

    def _connect(self):
        if self._app is not None:
            return
        self._app = win32com.client.Dispatch("Outlook.Application")
        self._ns = self._app.GetNamespace("MAPI")
        self._calendar = self._find_calendar()

    def _find_calendar(self):
        """Finds the correct Exchange calendar (the one that syncs with Teams).

        Uses only DisplayName and ExchangeStoreType - never accesses SmtpAddress
        or CurrentUser.Address, which trigger the Outlook Object Model Guard popup.

        Priority:
          1. Any store whose DisplayName matches outlook_account (exact account match)
          2. First store with ExchangeStoreType == 1 (primary Exchange mailbox)
          3. GetDefaultFolder(9) as fallback
        """
        try:
            exchange_candidate = None
            for store in self._ns.Stores:
                try:
                    display = getattr(store, "DisplayName", "").lower()
                    stype = getattr(store, "ExchangeStoreType", 0)

                    # Account filter wins regardless of store type
                    if self._account_filter and (
                        self._account_filter in display or display in self._account_filter
                    ):
                        return store.GetDefaultFolder(OL_FOLDER_CALENDAR)

                    # Keep first primary Exchange mailbox as fallback
                    if stype == 1 and exchange_candidate is None:
                        exchange_candidate = store.GetDefaultFolder(OL_FOLDER_CALENDAR)
                except Exception:
                    continue

            if exchange_candidate:
                return exchange_candidate
        except Exception:
            pass

        return self._ns.GetDefaultFolder(OL_FOLDER_CALENDAR)

    def authenticate(self):
        """Verifies Outlook access. No browser, no OAuth, no admin required."""
        self._connect()
        try:
            _ = self._calendar.Items.Count
            folder_name = getattr(self._calendar, "Name", "Calendar")
            try:
                store_name = self._calendar.Store.DisplayName
            except Exception:
                store_name = ""
            label = f"{store_name} > {folder_name}" if store_name else folder_name
            print("\n=== MICROSOFT (Outlook COM) ===")
            print(f"Calendar: {label}")
            print("No admin authorization required.")
            print("===============================\n")
        except Exception as e:
            raise RuntimeError(
                f"Could not access the Outlook calendar: {e}\n"
                "Make sure Outlook is installed and your account is signed in."
            ) from e

    def create_event(self, payload: dict) -> str | None:
        self._connect()
        appt = self._app.CreateItem(OL_APPOINTMENT_ITEM)
        self._apply(appt, payload)
        appt.Save()
        return appt.EntryID

    def update_event(self, entry_id: str, payload: dict) -> str:
        """Update an existing event and return its EntryID.

        If the stored EntryID is no longer valid (event was deleted manually),
        recreates the event and returns the new EntryID so the store can be
        updated - preventing perpetual duplicates on the next sync cycle.
        """
        self._connect()
        try:
            appt = self._ns.GetItemFromID(entry_id)
        except Exception:
            appt = self._app.CreateItem(OL_APPOINTMENT_ITEM)
        self._apply(appt, payload)
        appt.Save()
        return appt.EntryID

    def delete_event(self, entry_id: str):
        self._connect()
        try:
            appt = self._ns.GetItemFromID(entry_id)
            appt.Delete()
        except Exception:
            pass  # already deleted or invalid ID

    def _apply(self, appt, payload: dict):
        appt.Subject = payload.get("subject", "(no title)")
        appt.Body = payload.get("body", {}).get("content", "")

        loc = payload.get("location", {}).get("displayName", "")
        if loc:
            appt.Location = loc

        if payload.get("isAllDay"):
            appt.AllDayEvent = True
            appt.Start = payload["start"]["dateTime"][:10]
            appt.End = payload["end"]["dateTime"][:10]
        else:
            appt.AllDayEvent = False
            appt.Start = _to_local_str(
                payload["start"]["dateTime"],
                payload["start"].get("timeZone", "UTC"),
            )
            appt.End = _to_local_str(
                payload["end"]["dateTime"],
                payload["end"].get("timeZone", "UTC"),
            )


def _to_local_str(dt_str: str, tz_name: str) -> str:
    """Converts ISO 8601 + timezone to a local-time string for the current machine.

    Outlook interprets datetime strings as local time, so conversion must happen
    before assignment. Avoids relying on pywintypes.Time with IANA zones, which
    fails on Windows without tzdata.
    """
    s = dt_str.rstrip("Z")
    for sep in ("+", "-"):
        idx = s.rfind(sep, 10)
        if idx != -1:
            s = s[:idx]
            break
    dt = datetime.fromisoformat(s)

    tz_upper = tz_name.strip().upper()
    if tz_upper in ("UTC", "GMT", "ETC/UTC", "UTC+00:00"):
        dt = dt.replace(tzinfo=stdlib_tz.utc)
    else:
        try:
            from zoneinfo import ZoneInfo
            dt = dt.replace(tzinfo=ZoneInfo(tz_name))
        except Exception:
            dt = dt.replace(tzinfo=stdlib_tz.utc)

    local_dt = dt.astimezone().replace(tzinfo=None)
    return local_dt.strftime("%m/%d/%Y %H:%M:%S")
