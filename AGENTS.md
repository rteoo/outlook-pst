# Project instructions

This repository contains a standalone local mail CLI and a skills plugin.
The canonical implementation is `skills/outlook-pst/scripts/outlook_pst.py`.
Root `plugin.json` and `.codex-plugin/plugin.json` own package metadata; keep
their identity, version, license, and interface synchronized.

## Commands

Verified on Windows with Python 3.14:

```powershell
python -B -W error::ResourceWarning -m unittest discover -s skills/outlook-pst/tests -v
python -B skills/outlook-pst/scripts/outlook_pst.py --help
python -B skills/outlook-pst/scripts/build_plugin.py --out dist/plugin-0.1.0
```

Optional static checks use an existing Ruff installation:
`python -m ruff check --no-cache skills/outlook-pst`.

Tests use fake libpff and Outlook backends. They do not prove live mail access.
Every test removes its files and directories, including on failure.
Preserve the CLI's existing commands and default dry-run behavior.

## Git and privacy

The initial scaffold commit may be made directly on `main`; later changes use
task branches and review. Publication still requires user authorization.
Stage explicit paths and inspect the complete outgoing tree and metadata.
Never commit real mail, attachments, account names, local machine paths,
credentials, exports, private fixtures, or private source history.
Use a repository-local GitHub noreply identity. Test fixtures use fictional
names and reserved example domains. Generated packages stay in ignored `dist/`.

The package builder allows only listed source files and refuses non-empty output.
Do not include caches or dependencies by copying entire source directories.
Mailbox and profile changes require approval of exact targets. Live regression
work uses an approved scratch PST, never a real mailbox or an evidence original.
