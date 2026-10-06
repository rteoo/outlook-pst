---
name: "outlook-pst"
description: "Read, search, export, or edit Outlook Classic .pst/.ost mail: list and filter messages, dump .eml or pffexport-style folders with hashes, move/edit/categorize items in a running Outlook."
---

# Outlook PST/OST

Use this skill when the task involves an Outlook data file (`.pst` archive, `.ost`
offline cache) or mail held in Classic Outlook: find messages, read one, export
evidence, or change items. Mail in Gmail, IMAP, or Microsoft 365 reached through
an API or connector is out of scope; use that connector instead.

## Host notes

- File reads (`tree`, `list`, `show`, `export`) work on any OS through libpff.
- Live commands (`outlook ...`) and reads of files Outlook holds open need
  Windows with Classic Outlook and pywin32. The new Outlook has no COM API.
- Dependencies are supplied per run with `uv`. Nothing is installed into the
  host Python. Off Windows, drop `--with pywin32`.

```bash
S="/absolute/path/to/this/skill/scripts/outlook_pst.py"
uv run --with libpff-python --with pywin32 python "$S" --help
```

Resolve `scripts/outlook_pst.py` relative to this `SKILL.md`, including when
loaded from an installed plugin. Do not assume a vault checkout exists.
On Windows use PowerShell:

```powershell
$scriptPath = 'C:\absolute\path\to\this\skill\scripts\outlook_pst.py'
uv run --with libpff-python --with pywin32 python $scriptPath --help
```

Check for `uv` and the required dependencies before executing. Do not install
or upgrade host tools or download missing dependencies without authorization.
Attaching or detaching PSTs changes the Outlook profile and also needs approval.
Mail content is private: keep exports in the user's approved local destination
and do not upload them.

## Pick the mode

| Need | Command | Changes anything? |
|---|---|---|
| Folder structure and counts | `tree FILE` | no |
| Find messages | `list FILE [filters] [--format table\|json\|csv]` | no |
| Read one message | `show FILE ID [--html]` | no |
| Evidence or archive export | `export FILE OUT --format eml\|dir` | no (writes to `OUT` only) |
| Stores in the running Outlook | `outlook stores` | no |
| Change items | `outlook move\|edit\|delete ... [--apply]` | only with `--apply` |
| Save live items as `.msg` | `outlook export --store S OUT` | no |
| Mount or unmount a PST | `outlook attach PST [--create]`, `outlook detach PST` | Outlook profile |
| Launch Outlook | `outlook open --select outlook:inbox \| --msg F.msg \| --profile P \| --safe` | no |

Filters, shared by `list`, `export` and the live commands: `--folder` (a
substring of the folder path for file reads; an exact path for live commands),
`--from`, `--to`, `--subject`, `--text` (searches the body), `--since`/`--until`
YYYY-MM-DD, `--has-attachments`, `--limit`. Live commands take `--store`
(a display name, path, or file name from `outlook stores`). `--folder` also
accepts `@inbox`, `@sent`, `@drafts`, `@deleted`, `@junk` and `@outbox`, which
resolve localized folder names; add `--recursive` to include subfolders, or
select items with `--id`.

## Reading files, including while Outlook runs

`--via auto` is the default:

1. libpff opens the file directly.
2. If the file is locked and the running Outlook has it open as a store, the
   script reads that store through Outlook and says so on stderr.
3. If the file is locked but Outlook does not have it attached, the script
   stops. Outlook keeps a detached PST locked until Outlook exits. Close Outlook,
   or pass `--via outlook`, which attaches the PST for the read and detaches it
   afterwards.

Attaching lets Outlook write to the PST: it creates folders such as
`Deleted Items` on attach. For forensic work, read a copy taken while Outlook
was closed, and never pass `--via outlook` on an evidence original. Use
`--via pff` to forbid the Outlook fallback.

The two backends expose a common message view. Validate parity on a scratch
archive before relying on equivalent content; fake-backend tests do not prove
live Outlook or archive compatibility. Differences to expect:

- The top folder is `Top of Outlook data file/...` through libpff and the store
  display name (`Outlook Data File/...`) through Outlook. Filter on subfolder
  names.
- PST ids are libpff node ids on both paths. An `.ost` read through Outlook
  uses its long Exchange EntryID.
- The Outlook path cannot see orphaned items or hidden search folders. It writes
  `Body.txt` + `Body.html` where libpff writes `Body.html` + `Body.rtf`.

## Changing items safely

1. Run the live command without `--apply`. It prints one `would: ...` line per
   item and changes nothing.
   Each line includes the EntryID so approval can identify the exact item.
2. Show the plan to the user. On a real mailbox, get explicit approval before
   re-running with `--apply`. Edits to an `.ost` store sync to the Exchange
   server and reach every device.
3. `edit` takes `--set-subject`, `--replace OLD NEW` (repeatable; HTML bodies are
   edited as raw HTML), `--mark read|unread`, and `--add-category`. RTF bodies
   are skipped with a warning rather than flattened.
4. `delete` only moves to Deleted Items. It refuses items already there, where
   a delete would be permanent.
5. `move --create` makes missing destination folders, but only with `--apply`.
   It validates a nonempty source selection before creating destination folders.
6. Replacement previews simulate `--set-subject` followed by all replacements in
   command order. `applied:` is printed only after Outlook reports success.

## Reading the output

- `list` tables are `id`, local time, `@` for attachments, sender, subject, and
  `[folder]`. JSON and CSV carry the same fields plus `to`, `cc`, and
  `message_id`.
- `export` refuses a non-empty output directory. `OUT/manifest.csv` lists every
  file written, with MD5 and SHA-256. `--format dir` writes one `Message<id>/`
  folder per message, with `Message.txt`, `Recipients.txt`, any
  `InternetHeaders.txt`, the bodies, and `Attachments/`. `--format eml` keeps a
  received message's original transport headers and rebuilds the MIME parts.
- Exit code 1 with `N warning(s)` on stderr means output is incomplete: an
  unreadable message or attachment, or an edit skipped. Report it; never
  present the run as complete.
- `--limit` must be positive; reversed date ranges and empty replacement text
  are rejected before accessing Outlook. Live item operations process MailItems
  only, and explicit selections reject duplicate IDs and non-mail items.
- CSV presentation fields prefix formula-like values with an apostrophe for
  spreadsheet safety. JSON and exported message contents retain their values.
  Exports rebuild MIME and are not byte-identical forensic copies of the source.
- Human-facing fields escape terminal control characters. JSON and evidence
  values retain their content. Malformed RTF and unreadable attachment metadata
  generate warnings while later readable items continue.
- Rebuilt EML does not retain inline attachment Content-ID/related MIME metadata.
  Sanitized folder names can collide; use the original paths in the manifest.

## Traps found live

- pypff decodes 8-bit string properties as cp1252, but Outlook 16 writes them in
  the message codepage, usually UTF-8. The script decodes them itself.
- Outlook often stores the body as compressed RTF only, either HTML wrapped in
  RTF (`\fromhtml1`) or plain RTF. The script decodes both, so `--text` and
  exports still see the body.
- An unresolved recipient keeps the typed address as its display name. The
  script reports it as the address.
- Never create test mail with COM `Folder.Items.Add` on a PST: the draft is saved
  to the default mailbox's Drafts and syncs to the server. Build fixtures with
  `CreateItem` → `SaveAs(.msg)` → `Namespace.OpenSharedItem` → `Move`, and check
  the Drafts count before and after.
- Outlook started through COM exits when the script releases it. A test that
  needs Outlook running must start Outlook itself (`outlook open`).
- `PR_TRANSPORT_MESSAGE_HEADERS` cannot be set through the Outlook object model,
  so received-mail fixtures cannot be built through COM.

## Validation

```bash
python -B -W error::ResourceWarning /absolute/path/to/this/skill/tests/test_outlook_pst.py
```

The suite uses fakes for libpff and COM and never touches Outlook. Before
claiming a live path works, exercise it on a scratch PST, not a real mailbox.

## Report back

State which read path ran (libpff or Outlook), the message count, the filters
used, any warnings, and for exports the output directory and manifest. For live
changes, quote the dry-run plan, then the `applied:` lines after approval.
