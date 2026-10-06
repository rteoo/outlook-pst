# Outlook PST/OST

Read and search local Outlook archives, export messages with hash manifests,
and preview changes to mail in Classic Outlook. Includes a Python CLI and a
local skills plugin for Codex. No mail service account or remote server is needed.

This is an early release. Automated tests use fake backends; native libpff
loading and invalid-file handling have been checked on Windows. Successful
reads of real archives and live Outlook mutations are not verified by this
release's evidence. See [the review](docs/review.md) for fixes and limits.

## Requirements

Python 3.10 or later. Archive commands use `libpff-python` (imported as `pypff`)
on Windows, Linux, or macOS. Live commands need Windows, `pywin32`, and Classic
Outlook. New Outlook has no compatible COM API.

The examples use an existing `uv` installation to resolve dependencies per run.
They do not modify a project's dependencies or install packages into host Python.
Review dependency downloads before running them. CLI help, plugin packaging,
and all regression tests use only Python's standard library.

## Start with the CLI

Clone this repository and work from its root:

```powershell
git clone https://github.com/rteoo/outlook-pst.git
cd outlook-pst
$scriptPath = '.\skills\outlook-pst\scripts\outlook_pst.py'
python -B $scriptPath --help
uv run --with libpff-python python $scriptPath tree 'C:\mail\archive.pst' --via pff
uv run --with libpff-python python $scriptPath list 'C:\mail\archive.pst' --via pff --subject invoice --since 2026-01-01 --limit 20 --format json
uv run --with libpff-python python $scriptPath show 'C:\mail\archive.pst' 2097188 --via pff
uv run --with libpff-python python $scriptPath export 'C:\mail\archive.pst' 'C:\mail\export' --via pff --subject invoice --format eml
```

On macOS/Linux use `S=./skills/outlook-pst/scripts/outlook_pst.py` and pass `"$S"`
with your local archive and output paths. Add `--with pywin32` on Windows when
using Outlook-backed access. JSON listings emit one object per line, not an array.

## Command map

| Command | Behavior |
| --- | --- |
| `tree FILE` | Folder structure and counts |
| `list FILE` | Filter and list messages as table, JSON Lines, or CSV |
| `show FILE ID` | Read one message as JSON; `--html` selects its HTML body |
| `export FILE OUT --format eml` | Rebuild `.eml` messages with transport headers |
| `export FILE OUT --format dir` | Bodies, headers, recipients, attachments, and hash manifest |
| `outlook stores` / `outlook list --store S` | Inspect attached stores and MailItems |
| `outlook move` / `edit` / `delete` | Preview selected changes; apply only with `--apply` |
| `outlook export --store S OUT` | Save selected items as `.msg` |
| `outlook attach PST` / `detach PST` | Change the Outlook profile |
| `outlook open` | Launch Classic Outlook with supported switches |

Selection filters include `--folder`, `--from`, `--to`, `--subject`, `--text`,
`--since`, `--until`, `--has-attachments`, and positive `--limit`. Dates use local
calendar days. Live commands support `@inbox`, `@sent`, `@drafts`, `@deleted`,
`@junk`, `@outbox`, `--recursive`, and explicit `--id`. With `--id`, folder and
message filters are bypassed; duplicate IDs and non-MailItems are rejected.

## Outlook previews

```powershell
uv run --with pywin32 python $scriptPath outlook stores
uv run --with pywin32 python $scriptPath outlook edit --store 'archive.pst' --folder @inbox --subject invoice --limit 10 --mark read
```

Inspect the `would:` lines and obtain approval for those exact items before
rerunning with `--apply`. Exchange-backed changes can synchronize to the server
and other devices. Selection is reevaluated on every invocation; use explicit
IDs when you need to bind approval to particular items. `delete` refuses items
in Deleted Items and its descendants, where deletion could be permanent.

`--via auto` tries libpff, then an already attached Outlook store if locked.
`--via outlook` may attach a detached PST temporarily; attachment can modify
the archive even for a read command. Use a closed-Outlook copy with `--via pff`
for evidence work. Attaching, detaching, and creating PSTs need separate approval.

## Exports and privacy

Archive export requires a new or empty output directory. `manifest.csv` records
file size, MD5, and SHA-256; hashes check exported-file integrity, not authenticity.
Exported MIME is reconstructed and is not a byte-identical source copy. Keep
the original archive when preserving evidence or checking mail signatures.
Live `.msg` exports avoid filename collisions and do not produce this manifest.

CSV presentation fields prefix common formula-like values with an apostrophe.
JSON, message bodies, headers, and attachments retain their original content.
HTML and attachments may contain active content: the CLI does not render or
execute them. Keep exports private. Warnings return exit code 1 and indicate
incomplete results. See [SECURITY.md](SECURITY.md).

## Build and install the local plugin

```powershell
python -B skills/outlook-pst/scripts/build_plugin.py --out dist/plugin-0.1.0
codex plugin marketplace add ./dist/plugin-0.1.0
codex plugin add outlook-pst@outlook-pst-local
```

Then start a new chat and select **Outlook PST/OST**. The builder emits a
self-contained plugin, a ZIP with its hidden compatibility manifest, and a
separate local marketplace catalog. The root portable manifest and Codex
overlay are checked for agreement. Packaging copies only an explicit allowlist,
includes the MIT license, and refuses non-empty output directories.

Installation changes local Codex configuration; it is not performed by the
builder. This plugin needs a local execution host for files and processes.
Publishing or saving it to an account does not provide remote file access.
See the [OpenAI packaging guide](https://developers.openai.com/plugins/build/plugins).

## Development

```powershell
python -B -W error::ResourceWarning -m unittest discover -s skills/outlook-pst/tests -v
```

All tests use temporary directories that clean up on failure. COM tests use
pure Python fakes and run on every OS without Outlook or pywin32. CI runs the
suite and plugin build on Windows, Linux, and macOS with Python 3.10 and 3.14.
Native library and successful live Outlook behavior require separate validation.

MIT licensed. Native dependencies have their own licenses.
