# Setup Guide

[Português](SETUP.pt-BR.md)

This guide is detailed on purpose so a non-technical colleague can follow it.

The Teams calendar is the same calendar as your Microsoft 365 mailbox. So to make
your Google Calendar events show up in Teams, the app has to write them into that
mailbox calendar. There are only two ways to do that, and this guide covers both.

## Before you start

- Windows 10 or 11.
- Your institutional account that exists on both sides, for example
  `you@your-domain` on Google (calendar/email) and the same address on Microsoft
  365 (Teams).
- A few minutes. The first sync can take 1 to 2 minutes.

## Which method should I use?

| | Method A: admin authorization | Method B: Outlook Classic |
|---|---|---|
| Outlook installed on the PC | Not needed | Required (Classic) |
| Needs IT admin | Yes, once for everyone | No |
| Best for | Many colleagues, no Outlook | A PC you control |

- You can install Outlook Classic on your PC: use **Method B**. You do not need IT.
- You want several colleagues to use it and they do not have Outlook: ask IT for
  **Method A**. One approval covers everyone, so nobody has to keep requesting it.

---

# Method A: with admin authorization (Microsoft Graph)

The app writes to your Microsoft 365 calendar through the Microsoft Graph API, so
no Outlook is needed on the PC. The university tenant blocks each user from
approving this alone, so it needs ONE approval from an IT admin for the whole
organization. After that, nobody is prompted again.

## Part 1 - For the IT admin (done once for the whole org)

1. Go to the Microsoft Entra admin center: https://entra.microsoft.com
2. Open **Identity → Applications → App registrations → New registration**.
3. Name it (for example `GCalSync`). Under **Supported account types** choose
   **Accounts in this organizational directory only (single tenant)**. Click
   **Register**.
4. Open the new app → **Authentication** → **Add a platform** → **Mobile and
   desktop applications** → check `https://login.microsoftonline.com/common/oauth2/nativeclient`
   and add the redirect URI `http://localhost`. Set **Allow public client flows**
   to **Yes**. Save.
5. Open **API permissions → Add a permission → Microsoft Graph → Delegated
   permissions →** search and add **`Calendars.ReadWrite`**.
6. Click **Grant admin consent for [organization]** and confirm. This is the one
   approval that covers every user.
7. From the app **Overview**, copy the **Application (client) ID** and the
   **Directory (tenant) ID**. Send both to the people who will use the app.

## Part 2 - For each user

1. Install GCalSync (download `GCalSync-Setup.exe` from the
   [Releases page](https://github.com/vinicq/gcal-to-outlook/releases) and run it).
2. The setup wizard opens. If it does not find Outlook, it asks for the Graph
   details. Enter the **client ID** and **tenant ID** the admin sent. (You can
   also edit `config.json` in `%LOCALAPPDATA%\GCalSync` directly:)
   ```json
   "microsoft": {
     "mode": "graph",
     "client_id": "CLIENT-ID-FROM-ADMIN",
     "tenant_id": "TENANT-ID-FROM-ADMIN"
   }
   ```
3. The browser opens twice: once to sign in to **Google** (approve calendar
   access), once to sign in to **Microsoft**. No passwords are stored by the app.
4. The monitor window shows both accounts as **Connected**. Click **Sync now**.

## Verify

Open Teams → Calendar. Your Google events appear within a minute or two. Tick
**Start automatically when Windows starts** to keep it running in the background.

---

# Method B: with Outlook Classic (no admin needed)

The app writes events into the Outlook calendar through the local Outlook program.
Teams shows them because Outlook holds your Microsoft 365 account. This works on a
single PC without involving IT.

> **You must use Outlook Classic, not "new Outlook".**
> The "new Outlook" is a different, web-based app. It does not give GCalSync the
> local access it needs, so the app will not work with it. Every step below assumes
> Outlook Classic.

## Step 0 - Check which Outlook you have

Open Outlook. Look at the top-right corner of the window:

- If you see a toggle labelled **"New Outlook"** that is **on**, turn it **off**.
  Outlook restarts in Classic mode.
- Classic Outlook has the old ribbon and a **File** menu in the top-left. The new
  Outlook looks simpler and has the toggle on.

If you do not have Outlook Classic at all, install it (next step).

## Step 1 - Install Outlook Classic

- Instructions (Microsoft): https://support.microsoft.com/en-US/Outlook/install-or-reinstall-classic-outlook-on-a-windows-pc
- Direct installer (English): https://go.microsoft.com/fwlink/?linkid=2276500&clcid=0x409
- Direct installer (Português - Brasil): https://go.microsoft.com/fwlink/?linkid=2276500&clcid=0x416

Run the installer, then open **Outlook (classic)** from the Start menu. On first
launch it may ask you to create a profile; accept the default.

## Step 2 - Add your Microsoft 365 account (required)

This is the account whose calendar Teams reads. It must be added for the app to
work.

1. In Outlook Classic, go to **File → Add Account** (top-left).
2. Type your institutional address (for example `you@your-domain`) and click
   **Connect**.
3. Enter your **Microsoft** password and complete any sign-in / MFA prompts.
4. Outlook detects the Exchange / Microsoft 365 mailbox automatically and finishes.
5. Wait until the mailbox finishes loading and its calendar appears in the
   **Calendar** view. In **File → Account Information** it shows as
   **Microsoft Exchange**.

## Step 3 - (Optional) Add your Google account

You do NOT need this for the app. GCalSync reads your Google Calendar through the
Google API during its own login, not through Outlook. Add the Google account only
if you also want your Google email inside Outlook.

1. Go to **File → Add Account** again.
2. Type the Google address and click **Connect**. Outlook adds it as
   **IMAP/SMTP**; the Google sign-in page opens, log in and approve.
3. Finish. (IMAP brings email only, not the Google calendar - that is expected.)

## Step 4 - Confirm you are in Classic Outlook

Check the top-right corner again. If a **"New Outlook"** toggle is on, turn it
**off**. The app only works with Classic.

## Step 5 - Install GCalSync

1. Download `GCalSync-Setup.exe` from the
   [Releases page](https://github.com/vinicq/gcal-to-outlook/releases).
2. Run it. Windows SmartScreen may warn about an unrecognized publisher: click
   **More info → Run anyway**. It installs per user, no admin needed.

## Step 6 - First run and first sync

1. The setup wizard opens. It detects Outlook automatically and selects
   **Outlook COM** mode.
2. It asks which Outlook account to sync into. Enter your Microsoft 365 address
   (the one from Step 2).
3. The browser opens once to sign in to **Google**. Approve calendar access.
4. The first sync runs. It can take 1 to 2 minutes. Outlook must be open (or it
   opens automatically).

## Step 7 - Verify in Teams

Open Teams → Calendar. Your Google events appear within a minute or two (Teams and
Outlook share the same Microsoft 365 calendar). They also appear in Outlook.

## Step 8 - Keep it running

In the monitor window, tick **Start automatically when Windows starts**. The sync
then runs silently in the background on every login.

---

# Troubleshooting

### Outlook popup: "A program is trying to access email address information"

This appears when Windows reports that your antivirus is off, expired, or not
reporting its status. It is an Outlook security setting, not the app. To stop it:

1. Close Outlook. Reopen it as Administrator (right-click the Outlook shortcut →
   **Run as administrator**).
2. Go to **File → Options → Trust Center → Trust Center Settings → Programmatic
   Access** and choose **Never warn me about suspicious activity**.
3. Restart Outlook normally. (The option is greyed out unless Outlook was started
   as Administrator.)

Or make sure one antivirus is active and up to date in Windows Security; then the
prompt stops on its own. Code-signing the app does not remove this prompt.

### "invalid_client: The provided client secret is invalid"

The Google OAuth client secret was reset or revoked. Generate a new secret for the
Desktop OAuth client in Google Cloud Console, download the updated
`google_credentials.json`, delete `google_token.json`, and log in again.

### Events do not appear in Teams

- Confirm you added the **Microsoft 365** account in Outlook (Step 2), not only
  the Google one. Teams shows only that mailbox calendar.
- Confirm you are in **Classic** Outlook, not new Outlook.
- Wait a couple of minutes; Teams can lag behind Outlook.

### The wizard says Outlook was not found

You are likely on "new Outlook" or Outlook Classic is not installed. Do Step 0 and
Step 1.

---

# FAQ

**Does this sync Teams events back to Google?** No. Sync is one-way, Google to
Microsoft only.

**Is my data sent anywhere?** No. The app runs locally. Your Google token stays on
your machine. The app reads Google and writes to your Microsoft 365 calendar.

**Do I need Microsoft 365 desktop / a paid license for Method B?** You need Outlook
Classic with your Microsoft 365 account signed in. Follow the Microsoft install
link in Step 1.
