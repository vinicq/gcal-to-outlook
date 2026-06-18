"""Behavior tests for GoogleClient.authenticate branching.

authenticate IS the unit under test, so it is never mocked. Only the OAuth
edges are faked: the saved-credentials loader, the installed-app flow, the
token refresh, and the service builder. Each fake is injected at the
google_client module boundary.

Oracle: the docstring contract on authenticate — a valid token builds the
service with no flow; an expired token with a refresh token refreshes silently
and persists; a dead refresh token falls back (interactive) or raises
(background); allow_interactive=False never opens the browser; the interactive
flow forces offline access + consent so a refresh_token is always returned.
"""

import pytest
from google.auth.exceptions import RefreshError

import google_client
from google_client import GoogleClient


class FakeCreds:
    def __init__(self, valid=True, expired=False, refresh_token="rt",
                 refresh_error=False):
        self.valid = valid
        self.expired = expired
        self.refresh_token = refresh_token
        self._refresh_error = refresh_error
        self.refreshed = False

    def refresh(self, _request):
        if self._refresh_error:
            raise RefreshError("invalid_grant")
        self.refreshed = True
        self.valid = True
        self.expired = False

    def to_json(self):
        return '{"token": "x"}'


def _patch_build(monkeypatch):
    built = {"creds": None}
    monkeypatch.setattr(
        google_client, "build",
        lambda *a, **k: built.__setitem__("creds", k.get("credentials")) or "service",
    )
    return built


def _client(tmp_path, token_exists):
    token = tmp_path / "token.json"
    if token_exists:
        token.write_text("{}", encoding="utf-8")
    return GoogleClient(
        credentials_file=str(tmp_path / "creds.json"),
        token_file=str(token),
    )


def test_valid_token_builds_service_without_flow(tmp_path, monkeypatch):
    g = _client(tmp_path, token_exists=True)
    monkeypatch.setattr(
        google_client.Credentials, "from_authorized_user_file",
        staticmethod(lambda *a, **k: FakeCreds(valid=True)),
    )
    _patch_build(monkeypatch)
    # If the flow were touched it would error: no client_secrets file is mocked.
    g.authenticate(allow_interactive=False)
    assert g.service == "service"


def test_expired_token_refreshes_and_persists(tmp_path, monkeypatch):
    g = _client(tmp_path, token_exists=True)
    creds = FakeCreds(valid=False, expired=True, refresh_token="rt")
    monkeypatch.setattr(
        google_client.Credentials, "from_authorized_user_file",
        staticmethod(lambda *a, **k: creds),
    )
    _patch_build(monkeypatch)
    g.authenticate(allow_interactive=False)
    assert creds.refreshed is True
    # The refreshed credentials are written back so the next launch reuses them.
    assert (tmp_path / "token.json").read_text(encoding="utf-8") == '{"token": "x"}'


def test_background_without_token_raises_and_never_opens_browser(tmp_path, monkeypatch):
    g = _client(tmp_path, token_exists=False)

    def _boom(*a, **k):
        raise AssertionError("interactive flow must not run in background mode")

    monkeypatch.setattr(
        google_client.InstalledAppFlow, "from_client_secrets_file",
        staticmethod(_boom),
    )
    with pytest.raises(RuntimeError, match="authorization required"):
        g.authenticate(allow_interactive=False)


def test_background_dead_refresh_token_raises(tmp_path, monkeypatch):
    g = _client(tmp_path, token_exists=True)
    creds = FakeCreds(valid=False, expired=True, refresh_token="rt",
                      refresh_error=True)
    monkeypatch.setattr(
        google_client.Credentials, "from_authorized_user_file",
        staticmethod(lambda *a, **k: creds),
    )
    with pytest.raises(RuntimeError, match="authorization required"):
        g.authenticate(allow_interactive=False)


def test_interactive_flow_forces_offline_consent(tmp_path, monkeypatch):
    g = _client(tmp_path, token_exists=False)
    captured = {}

    class FakeFlow:
        def run_local_server(self, **kwargs):
            captured.update(kwargs)
            return FakeCreds(valid=True)

    monkeypatch.setattr(
        google_client.InstalledAppFlow, "from_client_secrets_file",
        staticmethod(lambda *a, **k: FakeFlow()),
    )
    _patch_build(monkeypatch)
    g.authenticate(allow_interactive=True)
    # access_type=offline + prompt=consent guarantee Google returns a
    # refresh_token, the fix for re-authorization on every boot.
    assert captured.get("access_type") == "offline"
    assert captured.get("prompt") == "consent"
    assert (tmp_path / "token.json").exists()
