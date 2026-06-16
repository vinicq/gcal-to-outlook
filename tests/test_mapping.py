"""Behavior tests for mapping.google_to_ms / is_cancelled.

Oracle for every expected value: the module docstring and the documented
Google -> Microsoft Graph contract, not the production code's own output.
"""

from mapping import google_to_ms, is_cancelled


def _timed_event():
    return {
        "id": "evt-timed-1",
        "summary": "Sprint planning",
        "description": "Discuss the backlog",
        "location": "Room 401",
        "status": "confirmed",
        "start": {"dateTime": "2026-07-01T09:00:00-03:00", "timeZone": "America/Sao_Paulo"},
        "end": {"dateTime": "2026-07-01T10:00:00-03:00", "timeZone": "America/Sao_Paulo"},
    }


def test_timed_event_maps_subject_dates_body_and_location():
    payload = google_to_ms(_timed_event(), default_tz="UTC", tag_prefix="[GCAL]")

    # tag prefix is prepended to the summary
    assert payload["subject"] == "[GCAL] Sprint planning"
    # a timed event is not all-day
    assert payload["isAllDay"] is False
    # start/end dateTime and timeZone are carried through unchanged
    assert payload["start"] == {
        "dateTime": "2026-07-01T09:00:00-03:00",
        "timeZone": "America/Sao_Paulo",
    }
    assert payload["end"] == {
        "dateTime": "2026-07-01T10:00:00-03:00",
        "timeZone": "America/Sao_Paulo",
    }
    # the sync-id marker is appended to the body, after the description
    assert payload["body"]["contentType"] == "text"
    assert payload["body"]["content"] == "Discuss the backlog\n\n[sync-id: evt-timed-1]"
    # location maps to location.displayName
    assert payload["location"] == {"displayName": "Room 401"}


def test_empty_tag_prefix_keeps_raw_summary_without_leading_space():
    event = _timed_event()
    payload = google_to_ms(event, default_tz="UTC", tag_prefix="")

    # with no prefix, the subject is exactly the summary (no leading space)
    assert payload["subject"] == "Sprint planning"


def test_missing_timezone_falls_back_to_default_tz_for_timed_event():
    event = _timed_event()
    del event["start"]["timeZone"]
    del event["end"]["timeZone"]

    payload = google_to_ms(event, default_tz="Europe/Berlin", tag_prefix="")

    assert payload["start"]["timeZone"] == "Europe/Berlin"
    assert payload["end"]["timeZone"] == "Europe/Berlin"


def test_all_day_event_uses_midnight_and_default_tz():
    event = {
        "id": "evt-allday-1",
        "summary": "Company holiday",
        "status": "confirmed",
        "start": {"date": "2026-07-04"},
        "end": {"date": "2026-07-05"},
    }

    payload = google_to_ms(event, default_tz="America/New_York", tag_prefix="")

    assert payload["isAllDay"] is True
    # Google all-day "date" becomes a midnight dateTime tagged with default_tz
    assert payload["start"] == {
        "dateTime": "2026-07-04T00:00:00",
        "timeZone": "America/New_York",
    }
    assert payload["end"] == {
        "dateTime": "2026-07-05T00:00:00",
        "timeZone": "America/New_York",
    }


def test_organizer_is_excluded_and_optional_flag_maps_to_type():
    event = _timed_event()
    event["organizer"] = {"email": "boss@org.com"}
    event["attendees"] = [
        {"email": "boss@org.com", "displayName": "The Boss"},
        {"email": "dev@org.com", "displayName": "Dev One"},
        {"email": "guest@org.com", "displayName": "Guest", "optional": True},
    ]

    payload = google_to_ms(event, default_tz="UTC", tag_prefix="")
    attendees = payload["attendees"]

    addresses = [a["emailAddress"]["address"] for a in attendees]
    # the organizer must not appear among the attendees
    assert "boss@org.com" not in addresses
    assert addresses == ["dev@org.com", "guest@org.com"]

    by_email = {a["emailAddress"]["address"]: a for a in attendees}
    # a non-optional attendee is required; an optional one is optional
    assert by_email["dev@org.com"]["type"] == "required"
    assert by_email["dev@org.com"]["emailAddress"]["name"] == "Dev One"
    assert by_email["guest@org.com"]["type"] == "optional"


def test_no_attendees_key_when_only_organizer_present():
    event = _timed_event()
    event["organizer"] = {"email": "solo@org.com"}
    event["attendees"] = [{"email": "solo@org.com"}]

    payload = google_to_ms(event, default_tz="UTC", tag_prefix="")

    # after excluding the organizer there is nobody left, so no attendees key
    assert "attendees" not in payload


def test_is_cancelled_true_only_for_cancelled_status():
    assert is_cancelled({"status": "cancelled"}) is True
    assert is_cancelled({"status": "confirmed"}) is False
    assert is_cancelled({}) is False
