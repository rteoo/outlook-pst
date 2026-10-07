# Install the Outlook PST/OST plugin

One local package shares `skills/outlook-pst/SKILL.md` and its Python CLI across
Codex, Claude Code, Cursor, and OpenClaw. It has no MCP server, hooks, account
connector, or background service. Every mail operation still runs on an execution
host that can access your chosen archive or Classic Outlook.

## Build once

From a source checkout, with an existing Python 3.10+ installation:

```powershell
python -B skills/outlook-pst/scripts/build_plugin.py --out dist/plugin-0.2.0
```

Use a new or empty output directory. The output layout is:

```text
plugin-0.2.0/
  .agents/plugins/marketplace.json
  .claude-plugin/marketplace.json
  .cursor-plugin/marketplace.json
  outlook-pst/
    plugin.json
    .codex-plugin/plugin.json
    .claude-plugin/plugin.json
    .cursor-plugin/plugin.json
    README.md
    LICENSE
    SECURITY.md
    assets/
    docs/
    skills/outlook-pst/
  outlook-pst-0.2.0.zip
```

Marketplace catalogs sit beside the plugin and are not inside the ZIP. Register
the build output directory as a marketplace, or load/install `outlook-pst/`
directly using the runtime's supported command. Repository catalogs instead point
to the repository root. The ZIP contains one plugin directory, including hidden
manifests. Copy the unpacked package to the execution host for OpenClaw or remote
agent sessions; do not copy mail or credentials with it.

Python and uv must already be available on the execution host. Archive reads need
`libpff-python`; live mail needs Windows, `pywin32`, and Classic Outlook. Check and
approve missing dependencies before downloading them. New Outlook does not support
the COM API. The plugin never installs dependencies automatically.

## Archive upload dialogs

For **Add plugin** or **Upload local plugin**, use the release asset
`outlook-pst-plugin-v0.1.0.zip`. This standalone archive contains one
`outlook-pst/` directory with the root manifest and hidden client manifests.
The builder's `outlook-pst-0.2.0.zip` has the same standalone layout.

The separate release asset `outlook-pst-v0.1.0.zip` is a marketplace bundle. It
adds an outer directory containing catalogs and the plugin directory beneath it.
Use it after extraction for CLI marketplace registration. Uploading that outer
bundle as a plugin hides `.claude-plugin/plugin.json` one directory too deep and
can produce a missing-manifest error.

Both release assets are built from the unchanged v0.1.0 tag, whose plugin
metadata declares version 0.2.0. See the README for direct download links.

## Codex

From the source checkout:

```powershell
codex plugin marketplace add ./dist/plugin-0.2.0
codex plugin add outlook-pst@outlook-pst-local
codex plugin list --marketplace outlook-pst-local --json
```

Start a new chat and select **Outlook PST/OST**. To use repository catalogs without
building, register `.` instead of the build output directory. Do not register both
copies under the same marketplace name.

Root `plugin.json` follows Agent Plugins 1.0. OpenAI presentation metadata lives
in `extensions.com.openai`, including `composerIcon` and `logo`. When present,
that object replaces the whole compatibility overlay for OpenAI-specific settings;
the two do not merge. The `.codex-plugin/plugin.json` overlay preserves older
client compatibility with the same metadata.

Reference: [OpenAI packaging and marketplaces](https://developers.openai.com/plugins/build/plugins).

## Claude Code

For one session, without registering a marketplace:

```powershell
claude plugin validate ./dist/plugin-0.2.0/outlook-pst --strict
claude --plugin-dir ./dist/plugin-0.2.0/outlook-pst
```

Invoke `/outlook-pst:outlook-pst`. For persistent installation:

```powershell
claude plugin marketplace add ./dist/plugin-0.2.0
claude plugin install outlook-pst@outlook-pst-local
claude plugin details outlook-pst
```

Restart the session or reload plugins after installation. Check that the component
inventory contains the `outlook-pst` skill. To install directly from this repository,
use `claude plugin marketplace add .` from the checkout. A published GitHub version
can be registered with `claude plugin marketplace add rteoo/outlook-pst` once the
catalog changes have reached the desired remote revision.

`.claude-plugin/plugin.json` supplies identity, version, and the skill root. Claude
also supports default skill layouts without a manifest, but explicit metadata makes
versioned distribution predictable. This package has no Claude-specific hooks or
MCP configuration.

References: [manifest](https://code.claude.com/docs/en/plugins-reference),
[marketplace installation](https://code.claude.com/docs/en/plugin-marketplaces).

## Cursor

The Cursor Agent CLI can load the package for one session:

```powershell
cursor --plugin-dir ./dist/plugin-0.2.0/outlook-pst
```

For local development in the Cursor IDE, copy the **contents of the packaged
`outlook-pst/` directory** into a new `outlook-pst` folder under
`~/.cursor/plugins/local/`. On Windows this is
`$env:USERPROFILE/.cursor/plugins/local/outlook-pst`. Do not overwrite an existing
installation. Restart Cursor or run **Developer: Reload Window**, then open
**Customize** and confirm that the skill is available.

Local imports can be restricted by organization policy. An installed marketplace
plugin with the same name takes precedence over a local copy. A symlink pointing
outside the local plugins directory is skipped, so use a contained copy.

Root `plugin.json` is portable. The small `.cursor-plugin/plugin.json` manifest adds
Cursor listing artwork and references the same skill root. Cursor's repository
catalog is `.cursor-plugin/marketplace.json`; team marketplace imports require the
appropriate plan and access. Public Marketplace listing requires a separate review
and submission; publishing a GitHub repository does not create that listing.

References: [installation and local testing](https://cursor.com/docs/plugins),
[manifest and marketplace format](https://cursor.com/docs/reference/plugins).

## OpenClaw

On the host that runs your OpenClaw Gateway, using the unpacked package path:

```sh
openclaw plugins install ./outlook-pst
openclaw plugins list
openclaw plugins inspect outlook-pst
```

Look for `Format: bundle`, the detected bundle format, and the skill root. Start a
new session to use the mapped skill. Supported reload behavior depends on the
installed OpenClaw version; follow that host's documentation and inspect the
running Gateway before assuming it loaded the skill.

OpenClaw supports Agent Plugins, Codex, Claude, and Cursor bundles. Native manifests
win detection first, then client-specific bundle markers, then root `plugin.json`.
Our existing Codex overlay takes precedence and maps the shared `skills/` root.
No `openclaw.plugin.json` or native JavaScript adapter is needed for this skill-only
bundle. Cold manifest inspection proves detection, not successful Gateway loading.

Only supported bundle features are mapped; generic Claude/Cursor JSON automation
is not equivalent to OpenClaw hook packs. This package includes neither.

Reference: [OpenClaw bundles](https://docs.openclaw.ai/plugins/bundles).

## Verify safely

From the installed plugin root, before opening any archive:

```powershell
python -B skills/outlook-pst/scripts/outlook_pst.py --help
python -B -W error::ResourceWarning -m unittest discover -s skills/outlook-pst/tests
```

These checks use synthetic backends and never access a mailbox. They verify the
packaged CLI independently of the original checkout. Confirm the skill appears in
the agent's inventory before using it. Successful manifest validation, detection,
and CLI help do not prove native archive access or live Outlook operations.

Mailbox changes and profile attachment/detachment still require approval of exact
targets. Use an approved scratch PST for live validation. Keep mail and exports in
approved private locations. Runtime installation changes local configuration and
does not submit the plugin to public directories.

## Verification status

Checked on Windows on 2026-10-07:

| Check | Result |
| --- | --- |
| Offline regression suite, Python 3.14 | 54 tests passed, including relocated CLI execution and packaged tests |
| Ruff | Passed |
| Agent Plugins 1.0 schema | Repository and packaged root manifests passed |
| Claude Code 2.1.287 | Strict package/catalog validation passed; isolated marketplace registration and installation passed; component inventory reports exactly one `outlook-pst` skill |
| Codex 0.159.0 CLI | Isolated marketplace registration and installation passed; plugin inventory reports version 0.2.0 installed and enabled; session execution remains unverified |
| Cursor Agent CLI 2026.08.11-e8db854 | Session-only `--plugin-dir` option confirmed; skill loading and IDE installation remain unverified |
| OpenClaw | No local executable available; bundle detection and Gateway loading remain unverified |

The approved Codex/Claude installation checks used temporary configuration
directories through `CODEX_HOME` and `CLAUDE_CONFIG_DIR`. Both directories were
removed afterward. No model session, archive, or mailbox was opened. Codex declined
to create PATH helper aliases under its temporary home, while installation and
inventory checks completed successfully.

Session execution, Cursor loading, and OpenClaw Gateway loading still need separate
verification. Package validation does not authorize changes to existing runtime
configuration, mailbox access, or public submission.
