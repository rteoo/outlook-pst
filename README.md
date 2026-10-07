# Outlook PST/OST

<p align="center">
  <img src="assets/outlook-pst-icon.png" width="128" alt="Outlook PST mail archive icon">
</p>

<p align="center">
  Find the email you need. Export it with confidence. Keep your archives local.
</p>

<p align="center">
  <a href="https://github.com/rteoo/outlook-pst/actions/workflows/ci.yml"><img src="https://github.com/rteoo/outlook-pst/actions/workflows/ci.yml/badge.svg" alt="CI status"></a>
  <a href="docs/review.md"><img src="https://img.shields.io/badge/status-early%20release-amber" alt="Early release"></a>
  <a href="https://www.python.org/"><img src="https://img.shields.io/badge/Python-3.10%2B-blue" alt="Python 3.10 or later"></a>
  <a href="LICENSE"><img src="https://img.shields.io/badge/license-MIT-blue.svg" alt="MIT license"></a>
</p>

Have an old `.pst` archive, an `.ost` cache, or mail in Classic Outlook?
Outlook PST helps you browse folders, find messages by sender or date, and
export the results. On Windows, you can also choose a mailbox and preview
changes before applying them.

Use it from the command line or with **Codex, Claude Code, Cursor, or OpenClaw**.
Archive commands work with files on your computer and need no mail-service
sign-in. One plugin package includes the shared skill, Python CLI, and project
artwork.
This is an **early release**; see [platform status and limitations](#platform-status-and-limitations)
for what has been verified.

## Highlights

- **Find messages quickly:** filter by folder, sender, recipient, subject,
  body text, date range, or attachments.
- **Read local archives:** inspect PST/OST files through libpff on Windows,
  macOS, and Linux.
- **Choose your mailbox:** list stores in Classic Outlook and select the one
  you want by name or data-file path.
- **Export useful files:** rebuild `.eml` messages, export bodies and
  attachments into folders, or save live Outlook items as `.msg`.
- **Check exported-file integrity:** archive exports include a CSV manifest
  with file sizes, MD5, and SHA-256 hashes.
- **Preview mailbox changes:** move mail, edit text, mark read/unread, add
  categories, or move items to Deleted Items. Item changes require `--apply`.
- **Use it with your agent:** one package for Codex, Claude Code, Cursor, and
  OpenClaw, with the same Python CLI and local archive access.

## Download the plugin

Choose the archive for your installation method from
[release v0.1.0](https://github.com/rteoo/outlook-pst/releases/tag/v0.1.0):

| Installation method | Download |
| --- | --- |
| **Add plugin / Upload local plugin dialog** | [Standalone plugin ZIP](https://github.com/rteoo/outlook-pst/releases/download/v0.1.0/outlook-pst-plugin-v0.1.0.zip) · [SHA-256](https://github.com/rteoo/outlook-pst/releases/download/v0.1.0/outlook-pst-plugin-v0.1.0.zip.sha256) |
| **CLI marketplace registration** | [Marketplace bundle ZIP](https://github.com/rteoo/outlook-pst/releases/download/v0.1.0/outlook-pst-v0.1.0.zip) · [SHA-256](https://github.com/rteoo/outlook-pst/releases/download/v0.1.0/outlook-pst-v0.1.0.zip.sha256) |

For a file-upload dialog, select **`outlook-pst-plugin-v0.1.0.zip`**. It contains
one `outlook-pst/` directory with `plugin.json`, `.claude-plugin/plugin.json`,
`.codex-plugin/plugin.json`, the skill, icon, and documentation. The marketplace
bundle adds catalogs around that directory and is intended for extracted CLI
installation; direct-upload dialogs cannot use that outer marketplace layout.

**Version note:** the GitHub release is named `v0.1.0`; its tagged source and
packaged agent manifests declare plugin version `0.2.0`.

To install from the marketplace bundle, with Python and uv already available:

```powershell
Expand-Archive ./outlook-pst-v0.1.0.zip -DestinationPath .
cd outlook-pst-v0.1.0
codex plugin marketplace add .
codex plugin add outlook-pst@outlook-pst-local
```

For other runtimes, use the extracted `outlook-pst/` directory:

- Claude Code: `claude --plugin-dir ./outlook-pst`.
- Cursor Agent CLI: `cursor --plugin-dir ./outlook-pst`.
- OpenClaw: on the Gateway host, `openclaw plugins install ./outlook-pst`.

Read [the installation guide](docs/plugin-installation.md) for requirements,
persistent installation, IDE setup, and verification limits. Archive dependencies
still need to be available or approved before downloading. Installing the plugin
in a cloud workspace does not grant access to mail files on your computer.

## Quick start

You need **Python 3.10+**. The examples use an existing **uv** installation to
resolve dependencies per run. Archive access needs `libpff-python`; live Outlook
access also needs `pywin32`, Windows, and **Classic Outlook**. Review dependency
downloads before running them. Help, packaging, and tests need only Python's
standard library.

Clone the repository, then check the available commands:

```powershell
git clone https://github.com/rteoo/outlook-pst.git
cd outlook-pst
$scriptPath = '.\skills\outlook-pst\scripts\outlook_pst.py'
python -B $scriptPath --help
```

### Browse an archive

Replace the example path with your archive. For evidence work, use a copy
taken while Outlook was closed; `--via pff` keeps access strictly offline.

```powershell
uv run --with libpff-python python $scriptPath tree 'C:\mail\archive.pst' --via pff
```

Find the first 20 messages about invoices:

```powershell
uv run --with libpff-python python $scriptPath list 'C:\mail\archive.pst' --via pff --subject invoice --limit 20 --format json
```

Read a message using an ID returned by `list`:

```powershell
uv run --with libpff-python python $scriptPath show 'C:\mail\archive.pst' 2097188 --via pff
```

JSON listings contain one object per line, making them easy to pipe into other
tools. On macOS/Linux, set `S=./skills/outlook-pst/scripts/outlook_pst.py` and
use `"$S"` in place of `$scriptPath`, with your own archive paths.

## Choose a mailbox or account

For mail already available in Classic Outlook, first list its attached stores:

```powershell
uv run --with pywin32 python $scriptPath outlook stores
```

Then pass the **exact displayed mailbox name** to `--store`:

```powershell
uv run --with pywin32 python $scriptPath outlook list --store 'Personal Mailbox' --folder @inbox --limit 20 --format json
```

`--store` also accepts a PST/OST filename or full path. Ambiguous matches stop
with an error. Selection is by Outlook store; an email address works only if it
matches that store's displayed name. There is no interactive account picker.
Offline commands select the archive file directly.

Common folders have shortcuts: `@inbox`, `@sent`, `@drafts`, `@deleted`, `@junk`,
and `@outbox`. Use `--recursive` to include subfolders. For a cross-mailbox move,
`--to-store` selects the destination.

## Export the messages you need

Choose a **new or empty output directory** and export matching messages as EML:

```powershell
uv run --with libpff-python python $scriptPath export 'C:\mail\archive.pst' 'C:\mail\export' --via pff --subject invoice --format eml
```

| Format | What you get |
| --- | --- |
| `--format eml` | Reconstructed mail messages with transport headers, bodies, and attachments |
| `--format dir` | A folder per message containing its summary, recipients, headers, bodies, and attachments |
| `outlook export --store S OUT` | Live Outlook items saved as `.msg` files |

Archive exports write `manifest.csv` with the original folder path, message ID,
file path, size, MD5, and SHA-256. Live `.msg` exports avoid filename collisions
but do not generate that manifest.

Hashes check exported-file integrity. They do not authenticate a message;
reconstructed EML is not a byte-identical copy. Keep the original archive when
preserving evidence or checking mail signatures.

## Preview changes before applying them

Start with a dry run. This example previews marking up to ten matching messages
as read:

```powershell
uv run --with pywin32 python $scriptPath outlook edit --store 'Personal Mailbox' --folder @inbox --subject invoice --limit 10 --mark read
```

Review the `would:` lines, including each item's EntryID. Once the exact items
are approved, use their IDs and add `--apply`:

```powershell
uv run --with pywin32 python $scriptPath outlook edit --store 'Personal Mailbox' --id 'ENTRY_ID_FROM_LIST' --mark read --apply
```

Repeat `--id` for more items. Explicit IDs bypass folder and message filters;
duplicate IDs and non-mail items are rejected. Filter selections are evaluated
again on every invocation, so IDs help keep approval tied to particular items.

| Action | Options |
| --- | --- |
| Move messages | `outlook move --to-folder 'Archive'`; optional `--to-store` and `--create` |
| Change text | `outlook edit --set-subject 'New subject'` or `--replace OLD NEW` |
| Update read state | `outlook edit --mark read` or `--mark unread` |
| Add a category | `outlook edit --add-category 'Reviewed'` |
| Move to Deleted Items | `outlook delete`; refuses items already there or in its descendants |

All these commands require `--store` and default to a preview. Destination
folders are created only after a valid, nonempty selection with `--apply`.
Replacement previews follow command order; HTML replacements edit raw HTML,
and RTF bodies are skipped with a warning to preserve formatting.

**Exchange-backed changes can sync to the server and other devices.** Changes
are not transactional: if a later operation fails, earlier ones may remain
applied. An `applied:` line is printed only after Outlook reports success.

## Use it with your agent

**One package, four agent runtimes.** Choose a runtime below; the same skill
and Python CLI run locally in each documented layout.

Build version **0.2.0** from a source checkout:

```powershell
python -B skills/outlook-pst/scripts/build_plugin.py --out dist/plugin-0.2.0
```

The output contains `outlook-pst/`, a ZIP, and separate marketplace catalogs for
Codex, Claude Code, and Cursor. The package includes this README, the icon,
installation instructions, and the shared skill. OpenClaw imports it as a bundle.

| Runtime | Package format | Verification status |
| --- | --- | --- |
| **Codex** | Agent Plugins manifest plus compatibility overlay | Installed and enabled in an isolated configuration |
| **Claude Code** | Claude manifest pointing to the shared skill | Installed; component inventory detected the skill |
| **Cursor** | Portable manifest plus Cursor metadata and logo | CLI loading option confirmed; skill loading unverified |
| **OpenClaw** | Imports the package as a compatible bundle | Documented layout; Gateway loading unverified |

### Install with Codex

From the source checkout, register the build output and install the plugin:

```powershell
codex plugin marketplace add ./dist/plugin-0.2.0
codex plugin add outlook-pst@outlook-pst-local
```

Start a new chat and select **Outlook PST/OST**. The repository also includes a
marketplace catalog, so `codex plugin marketplace add .` can register the checkout
without building. Register one source under the `outlook-pst-local` name.

### Try with Claude Code

Load the package for one session:

```powershell
claude --plugin-dir ./dist/plugin-0.2.0/outlook-pst
```

Invoke `/outlook-pst:outlook-pst`. For persistent installation, follow the
[Claude Code instructions](docs/plugin-installation.md#claude-code).

### Use Cursor or OpenClaw

Cursor's Agent CLI accepts `cursor --plugin-dir ./dist/plugin-0.2.0/outlook-pst`.
For the Cursor IDE, follow the [local installation steps](docs/plugin-installation.md#cursor).
On your OpenClaw Gateway host, install the unpacked package with
`openclaw plugins install ./outlook-pst`, then inspect it with
`openclaw plugins inspect outlook-pst`. See [the OpenClaw guide](docs/plugin-installation.md#openclaw)
for detection and session details.

See **[the installation guide](docs/plugin-installation.md)** for full commands,
reload steps, format details, and verification limits. Installer commands change
local runtime configuration; building the package does not install anything.
The plugin requires an execution host with access to your archives. Installing
it in a cloud session does not grant access to files on your computer.

Try prompts such as:

> Find messages about invoices in my local PST archive from January onward.
>
> Export the matching messages to my chosen local folder with a hash manifest.
>
> Show me a preview of marking these Outlook messages as read.

The builder includes only allowlisted files and refuses nonempty output
directories. No mail, exports, credentials, or dependencies are packaged.

## Command reference

| Command | Purpose |
| --- | --- |
| `tree FILE` | Browse folders and message counts |
| `list FILE` | Find messages; output table, JSON Lines, or CSV |
| `show FILE ID` | Read one message as JSON; add `--html` for its HTML body |
| `export FILE OUT` | Export selected archive messages and a hash manifest |
| `outlook stores` / `outlook list --store S` | Discover mailboxes and list live messages |
| `outlook move` / `edit` / `delete` | Preview or apply changes to selected MailItems |
| `outlook export --store S OUT` | Save live items as `.msg` |
| `outlook attach PST` / `detach PST` | Attach or detach a PST from the Outlook profile |
| `outlook open` | Launch Classic Outlook; supports `--profile`, `--select`, `--msg`, and `--safe` |

Shared filters: `--from`, `--to`, `--subject`, `--text`, `--since`, `--until`,
`--has-attachments`, and positive `--limit`. Dates use local calendar days.
For archive commands, `--folder` matches part of a path; for live commands,
it selects an exact folder path or one of the shortcuts above.

## Data safety and privacy

Mail and exports stay in the local destinations you choose; the CLI does not
send messages or upload mail. Keep exported content private. HTML and
attachments can contain active content, which the CLI does not render or execute.

`--via auto` first tries libpff and, when locked, can read a store already attached
to Outlook. `--via outlook` may temporarily attach a PST; attachment can modify
the archive. Attaching, detaching, and creating PSTs require separate approval.
Use `--via pff` for strict offline evidence access.

Human-facing output escapes terminal controls. CSV presentation fields neutralize
common spreadsheet formula prefixes; JSON and exported evidence preserve their
values. Warnings return exit code 1 and mean results may be incomplete. Report
them before relying on an export. See [SECURITY.md](SECURITY.md).

## Platform status and limitations

| Mode | Requirements |
| --- | --- |
| Archive reads and exports | Windows, macOS, or Linux; Python 3.10+ and `libpff-python` |
| Live Outlook and locked-file access | Windows, `pywin32`, and Classic Outlook; New Outlook has no compatible COM API |
| Help, tests, and plugin packaging | Python 3.10+ standard library |

Tests use synthetic libpff and COM backends. CI covers Windows, Linux, and macOS
on Python 3.10 and 3.14. Native libpff loading and invalid-file rejection were
checked on Windows; **successful real-archive reads, live Outlook changes, and
agent session execution remain unverified**. Codex and Claude marketplace
registration and installation passed in isolated temporary configuration
directories; Claude detected the skill, and Codex listed it as installed and
enabled. See [plugin verification status](docs/plugin-installation.md#verification-status)
and [the review](docs/review.md).

Large live mailboxes can be slow to scan. MIME assembly holds attachments in
memory, and RTF extraction is intentionally focused rather than a full RTF
interpreter. Malformed bodies and unreadable attachments produce warnings while
later readable items continue. Rebuilt EML currently omits inline attachment
Content-ID/related MIME metadata, and sanitized folder names can share an output
directory; original paths remain in the manifest.

## Develop and build

Source and tests live under [skills/outlook-pst](skills/outlook-pst). From a
source checkout, run:

```powershell
python -B -W error::ResourceWarning -m unittest discover -s skills/outlook-pst/tests -v
python -m ruff check --no-cache skills/outlook-pst
python -B skills/outlook-pst/scripts/build_plugin.py --out dist/plugin-0.2.0
```

Ruff is optional and must already be installed. Tests use temporary directories
that clean up on failure and never access a mailbox. Native and live validation
require a separately approved scratch archive.

## License

Outlook PST is released under the [MIT License](LICENSE).
