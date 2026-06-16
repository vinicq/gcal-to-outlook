"""Behavior tests for ms_client_com._to_local_str (a pure function).

Oracle: the function docstring. It converts an ISO-8601 string plus a timezone
name into a *local-time* string for the current machine, formatted as
"%m/%d/%Y %H:%M:%S". The class methods touch live Outlook COM and are not
unit-tested here (see tests/README.md).

To stay independent of the CI machine's timezone (J2), we never hard-code a
local string. Instead we parse the returned string back as a naive local time,
re-attach the machine's local zone, and compare the UTC instant against an
instant we compute independently from the documented conversion rule.
"""

from datetime import datetime, timezone

import pytest

# ms_client_com imports win32com at module load. Skip the whole module when
# pywin32 is unavailable (e.g. non-Windows CI), since the function under test
# cannot even be imported without it.
pytest.importorskip("win32com")

from ms_client_com import _to_local_str  # noqa: E402

OUT_FORMAT = "%m/%d/%Y %H:%M:%S"


def _output_as_utc(local_str: str) -> datetime:
    """Reverse the production conversion: read the local-time output back as a
    UTC instant, using the running machine's own local zone."""
    naive_local = datetime.strptime(local_str, OUT_FORMAT)
    local_zone = datetime.now().astimezone().tzinfo
    return naive_local.replace(tzinfo=local_zone).astimezone(timezone.utc)


def test_format_matches_contract():
    out = _to_local_str("2026-07-01T09:00:00", "UTC")

    # must parse exactly with the documented format; a wrong format raises here
    parsed = datetime.strptime(out, OUT_FORMAT)
    assert parsed.year == 2026


def test_utc_z_suffix_is_treated_as_utc():
    # "2026-07-01T12:00:00Z" is noon UTC; independent oracle: that instant.
    out = _to_local_str("2026-07-01T12:00:00Z", "UTC")

    expected = datetime(2026, 7, 1, 12, 0, 0, tzinfo=timezone.utc)
    assert _output_as_utc(out) == expected


def test_named_utc_zone_is_treated_as_utc():
    out = _to_local_str("2026-07-01T12:00:00", "UTC")

    expected = datetime(2026, 7, 1, 12, 0, 0, tzinfo=timezone.utc)
    assert _output_as_utc(out) == expected


def test_numeric_offset_is_stripped_then_zone_applied():
    # Documented behavior (code contract): the +/- offset is REMOVED from the
    # string, then the tz_name is applied. So "...T12:00:00-05:00" with tz_name
    # "UTC" is interpreted as 12:00 UTC, NOT 17:00 UTC. The offset does not move
    # the wall-clock time; only tz_name decides the zone.
    out_with_offset = _to_local_str("2026-07-01T12:00:00-05:00", "UTC")
    out_plain = _to_local_str("2026-07-01T12:00:00", "UTC")

    # both must resolve to the same instant: 12:00 UTC
    expected = datetime(2026, 7, 1, 12, 0, 0, tzinfo=timezone.utc)
    assert _output_as_utc(out_with_offset) == expected
    assert _output_as_utc(out_with_offset) == _output_as_utc(out_plain)


def test_iana_zone_is_applied_to_naive_wall_clock():
    # America/New_York is UTC-4 on 2026-07-01 (EDT). 12:00 wall clock there is
    # 16:00 UTC. Independent oracle: standard EDT offset for that date.
    out = _to_local_str("2026-07-01T12:00:00", "America/New_York")

    expected = datetime(2026, 7, 1, 16, 0, 0, tzinfo=timezone.utc)
    assert _output_as_utc(out) == expected
