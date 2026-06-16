"""Converts Google Calendar events into Microsoft Graph / Outlook payloads.

The sync is one-way (Google -> Microsoft), so only that direction is implemented.
"""


# ---------------------------------------------------------------------------
# Internal utilities
# ---------------------------------------------------------------------------

def _has_datetime(g_event: dict) -> bool:
    return "dateTime" in g_event.get("start", {})


# ---------------------------------------------------------------------------
# Direction Google -> Microsoft
# ---------------------------------------------------------------------------

def _google_attendees_to_ms(g_event: dict) -> list:
    """Converts Google Calendar attendees to MS Graph format.

    Excludes the organizer - MS already treats them as the event creator.
    """
    organizer_email = g_event.get("organizer", {}).get("email", "")
    result = []
    for a in g_event.get("attendees", []):
        email = a.get("email", "")
        if not email or email == organizer_email:
            continue
        result.append({
            "emailAddress": {
                "address": email,
                "name": a.get("displayName", ""),
            },
            "type": "optional" if a.get("optional") else "required",
        })
    return result


def google_to_ms(g_event: dict, default_tz: str, tag_prefix: str) -> dict:
    summary = g_event.get("summary", "(no title)")
    subject = f"{tag_prefix} {summary}".strip() if tag_prefix else summary

    google_id = g_event.get("id", "")
    description = g_event.get("description", "") or ""
    # tracking marker in the body, useful for auditing and future deletion support
    body_content = f"{description}\n\n[sync-id: {google_id}]".strip()

    payload = {
        "subject": subject,
        "body": {"contentType": "text", "content": body_content},
    }

    location = g_event.get("location")
    if location:
        payload["location"] = {"displayName": location}

    attendees = _google_attendees_to_ms(g_event)
    if attendees:
        payload["attendees"] = attendees

    if _has_datetime(g_event):
        start = g_event["start"]["dateTime"]
        end = g_event["end"]["dateTime"]
        start_tz = g_event["start"].get("timeZone", default_tz)
        end_tz = g_event["end"].get("timeZone", default_tz)
        payload["isAllDay"] = False
        payload["start"] = {"dateTime": start, "timeZone": start_tz}
        payload["end"] = {"dateTime": end, "timeZone": end_tz}
    else:
        # all-day event: Google uses "date" (YYYY-MM-DD)
        start_date = g_event["start"]["date"]
        end_date = g_event["end"]["date"]
        payload["isAllDay"] = True
        payload["start"] = {"dateTime": f"{start_date}T00:00:00", "timeZone": default_tz}
        payload["end"] = {"dateTime": f"{end_date}T00:00:00", "timeZone": default_tz}

    return payload


def is_cancelled(g_event: dict) -> bool:
    """In incremental sync, deleted events are returned with status 'cancelled'."""
    return g_event.get("status") == "cancelled"
