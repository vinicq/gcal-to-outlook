"""Behavior tests for MicrosoftClient (Graph) authenticate gating.

authenticate IS the unit under test. Only the MSAL app is faked: a stub whose
get_accounts() and acquire_token_* are scripted, injected by replacing
_build_app. The Store and network are never touched.

Oracle: the docstring contract — allow_interactive=False with no cached account
raises instead of opening a browser; the interactive default still runs the
browser flow.
"""

import pytest

from ms_client import MicrosoftClient


class FakeApp:
    def __init__(self, accounts=None, silent=None):
        self._accounts = accounts or []
        self._silent = silent
        self.interactive_called = False

    def get_accounts(self):
        return self._accounts

    def acquire_token_silent(self, scopes, account=None):
        return self._silent

    def acquire_token_interactive(self, scopes=None):
        self.interactive_called = True
        return {"access_token": "tok-interactive"}


def _client(tmp_path):
    return MicrosoftClient(
        client_id="cid", tenant_id="tid",
        token_cache_file=str(tmp_path / "cache.bin"), calendar_id="",
    )


def test_background_without_account_raises_and_no_browser(tmp_path, monkeypatch):
    g = _client(tmp_path)
    app = FakeApp(accounts=[])
    monkeypatch.setattr(g, "_build_app", lambda: (app, object()))
    with pytest.raises(RuntimeError, match="authorization required"):
        g.authenticate(allow_interactive=False)
    assert app.interactive_called is False


def test_background_stale_account_silent_fails_raises_and_no_browser(tmp_path, monkeypatch):
    # The common unattended-failure case: a cached account exists but its token
    # can no longer be refreshed silently (acquire_token_silent -> None). This
    # must still raise rather than fall through to the browser.
    g = _client(tmp_path)
    app = FakeApp(accounts=["acct"], silent=None)
    monkeypatch.setattr(g, "_build_app", lambda: (app, object()))
    with pytest.raises(RuntimeError, match="authorization required"):
        g.authenticate(allow_interactive=False)
    assert app.interactive_called is False


def test_cached_account_authenticates_silently(tmp_path, monkeypatch):
    g = _client(tmp_path)
    app = FakeApp(accounts=["acct"], silent={"access_token": "tok-silent"})
    monkeypatch.setattr(g, "_build_app", lambda: (app, object()))
    monkeypatch.setattr(g, "_save_cache", lambda cache: None)
    g.authenticate(allow_interactive=False)
    assert g._token == "tok-silent"
    assert app.interactive_called is False


def test_interactive_default_runs_browser_flow(tmp_path, monkeypatch):
    g = _client(tmp_path)
    app = FakeApp(accounts=[])
    monkeypatch.setattr(g, "_build_app", lambda: (app, object()))
    monkeypatch.setattr(g, "_save_cache", lambda cache: None)
    g.authenticate(allow_interactive=True)
    assert app.interactive_called is True
    assert g._token == "tok-interactive"
