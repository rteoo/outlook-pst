# Security

Report security issues through GitHub's private vulnerability reporting when
available. Do not put real mail, credentials, personal addresses, or account
names in public issues. If private reporting is unavailable, open a minimal
issue asking for a private reporting channel without disclosing exploit details.

This utility processes local, potentially untrusted mail archives through
native libpff code. Use maintained dependencies and a disposable environment for
untrusted archives. The Python regression suite does not audit native libraries.

Archive exports preserve potentially active HTML and attachments. Viewing or
opening them in another application is a separate trust decision. The CLI does
not render HTML, execute attachments, send messages, or upload mail.

CSV presentation fields neutralize common spreadsheet formula prefixes. Original
message bodies, headers, attachments, and JSON output retain their contents.
Keep exports in a private directory under your control; this utility does not
defend against another process concurrently replacing its directory tree.

Live commands require Classic Outlook on Windows. Edit, move, and soft-delete
commands default to previews; `--apply` changes selected mail. Exchange-backed
changes can synchronize to the server. Attaching a PST changes the local profile
and can modify the archive. Use `--via pff` for strictly offline evidence reads.
