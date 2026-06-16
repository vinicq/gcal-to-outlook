# Security Policy

## Supported versions

| Version | Supported |
|---------|-----------|
| 1.0.x   | Yes       |
| < 1.0   | No        |

Only the latest 1.0.x release receives security fixes. Older versions should
upgrade before reporting a problem.

## Reporting a vulnerability

Do not open a public issue for a security problem.

Report it privately one of two ways:

- Open a [GitHub Security Advisory](https://github.com/vinicq/gcal-to-outlook/security/advisories/new)
  on this repository, or
- Email the maintainer at `vinicq@gmail.com` with the subject prefix
  `[gcal-to-outlook security]`.

Include the version, your OS, the steps to reproduce, and what an attacker could
do with the flaw. Expect an acknowledgment within a few days. Once a fix ships,
you will be credited unless you ask otherwise.

## Handling of credentials and tokens

This tool stores all secrets locally on your machine. Nothing is uploaded to any
server the maintainer controls. The sensitive files are:

- `google_credentials.json` - your Google Cloud OAuth client
- `google_token.json` - your Google access and refresh tokens
- `ms_token_cache.bin` - the MSAL token cache (Graph mode)
- `config.json` - your account settings

Never attach any of these files to an issue, a pull request, a discussion, or a
log paste. They grant access to your calendar accounts. When sharing `sync.log`
for a bug report, redact any token strings, client IDs, and account addresses
first.

If you accidentally exposed `google_token.json` or `ms_token_cache.bin`, revoke
the session: remove the app's access from your
[Google account permissions](https://myaccount.google.com/permissions) and, for
Graph mode, from your Microsoft Entra app, then delete the local token files and
log in again.
