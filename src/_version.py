"""Single source of truth for the app version.

Keep this in sync with MyAppVersion in installer.iss when cutting a release.
Surfaced in the monitor UI and compared against the latest GitHub release to
tell the user when an update is available.
"""

__version__ = "1.0.5"

# Public repo used for the "new version available" check.
GITHUB_REPO = "vinicq/gcal-to-outlook"
RELEASES_URL = f"https://github.com/{GITHUB_REPO}/releases/latest"
