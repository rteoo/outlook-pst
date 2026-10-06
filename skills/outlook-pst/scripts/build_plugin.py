#!/usr/bin/env python3
"""Build the local Outlook PST/OST skills plugin from repository source.

Usage: python -B skills/outlook-pst/scripts/build_plugin.py --out dist/plugin-0.1.0
Cron: none; manual packaging only.
Dependencies: Python 3.10+ standard library; no mail backends needed.
Output: a self-contained plugin, ZIP, and separate local marketplace catalog.
No installation, upload, host configuration changes, or mailbox access.
"""

from __future__ import annotations

import argparse
import hashlib
import json
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile, ZipInfo

SOURCE = Path(__file__).resolve().parents[1]
PROJECT_ROOT = SOURCE.parents[1]
MANIFEST = json.loads((PROJECT_ROOT / "plugin.json").read_bytes())
NAME = MANIFEST["name"]
VERSION = MANIFEST["version"]
# Explicit allowlist: adding mail fixtures or exports cannot include them accidentally.
SOURCE_FILES = ("SKILL.md", "scripts/outlook_pst.py", "tests/test_outlook_pst.py", "tests/test_release_safety.py")
README = """# Outlook PST/OST

A local skills plugin for archive search, message reading, evidence export,
and approved Classic Outlook item changes. This package contains instructions,
a Python CLI, and fake-backend regression tests. It has no MCP server or hooks.

## Requirements

- Python 3.10+ and uv, already available on the execution host.
- libpff-python for archive reads; pywin32 and Classic Outlook on Windows for
  live access or reads of attached files that Outlook holds open.
- New Outlook does not provide the required COM API.

Dependencies are resolved per run by uv using the commands in SKILL.md; the
plugin does not install them. Obtain authorization for missing dependencies.
The fake-backend tests and CLI --help need only Python's standard library.

## Install locally

The builder writes a separate `.agents/plugins/marketplace.json` beside the
plugin folder. Register the builder's output directory in PowerShell:

```powershell
codex plugin marketplace add 'C:\\absolute\\path\\to\\build-output'
codex plugin add outlook-pst@outlook-pst-local
```

Then start a new chat and select Outlook PST/OST. Registration and installation
change the user's local Codex configuration; run them only when authorized.
No account upload or publication is required for this local package.

## Verify after installation

Resolve the installed skill's directory, then run:

```powershell
python -B -W error::ResourceWarning -m unittest discover -s .\\skills\\outlook-pst\\tests
python -B .\\skills\\outlook-pst\\scripts\\outlook_pst.py --help
```

Run those commands from the plugin root. They never access a mailbox.
For live validation use an approved scratch PST. Mailbox changes require a
dry-run preview and approval of the exact items; Exchange edits sync to devices.
Attaching or detaching a PST changes the Outlook profile and needs approval.
Use `--via pff` for strict read-only evidence access; Outlook attachment can
modify a PST. Keep mail and exports private in an approved local destination.

## Package format

`plugin.json` follows Agent Plugins 1.0; `.codex-plugin/plugin.json` provides
the compatibility overlay. Both use the same identity, version, and listing.
Only explicitly listed source files are packaged; no mail, attachments,
credentials, caches, dependency directories, or machine-specific paths.

Format and local installation reference:
https://developers.openai.com/plugins/build/plugins
"""


def json_bytes(value: dict) -> bytes:
    return (json.dumps(value, indent=2, ensure_ascii=False) + "\n").encode("utf-8")


def package_files(source: Path) -> dict[str, bytes]:
    files = {"README.md": README.encode("utf-8")}
    for relative in ("plugin.json", ".codex-plugin/plugin.json", "LICENSE"):
        path = PROJECT_ROOT / relative
        if not path.is_file() or path.is_symlink() or path.resolve() != PROJECT_ROOT / relative:
            raise ValueError(f"metadata must be a contained regular file: {relative}")
        files[relative] = path.read_bytes()
    portable = json.loads(files["plugin.json"])
    overlay = json.loads(files[".codex-plugin/plugin.json"])
    for key in ("name", "version", "description", "author", "license", "repository"):
        if portable.get(key) != overlay.get(key):
            raise ValueError(f"manifest mismatch: {key}")
    interface = portable["extensions"]["com.openai"]["interface"]
    if interface != overlay["interface"] or len(interface["shortDescription"]) > 30:
        raise ValueError("plugin listing mismatch or subtitle longer than 30 characters")
    if portable["name"] != NAME or portable["version"] != VERSION:
        raise ValueError("plugin metadata changed during packaging; run the builder again")
    for relative in SOURCE_FILES:
        path = source / relative
        if not path.is_file() or path.is_symlink() or path.resolve() != source.resolve() / relative:
            raise ValueError(f"source must be a contained regular file: {relative}")
        files[f"skills/{NAME}/{relative}"] = path.read_bytes()
    return files


def build(out: Path, source: Path = SOURCE) -> dict:
    files = package_files(source)  # Validate all inputs before creating output.
    out = out.resolve()
    if out == PROJECT_ROOT or out == SOURCE or SOURCE in out.parents:
        raise ValueError("output must be outside the skill source and cannot be the project root")
    if out.exists() and (not out.is_dir() or any(out.iterdir())):
        raise ValueError("output must be a new or empty directory")
    out.mkdir(parents=True, exist_ok=True)
    root = out / NAME
    for relative, payload in files.items():
        path = root / relative
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_bytes(payload)
    catalog = out / ".agents" / "plugins" / "marketplace.json"
    catalog.parent.mkdir(parents=True, exist_ok=True)
    catalog.write_bytes(json_bytes({
        "name": "outlook-pst-local",
        "interface": {"displayName": "Local Outlook plugins"},
        "plugins": [{
            "name": NAME,
            "source": {"source": "local", "path": f"./{NAME}"},
            "policy": {"installation": "AVAILABLE", "authentication": "ON_INSTALL"},
            "category": "Productivity",
        }],
    }))
    archive = out / f"{NAME}-{VERSION}.zip"
    with ZipFile(archive, "x", compression=ZIP_DEFLATED) as bundle:
        for relative, payload in sorted(files.items()):
            entry = ZipInfo(f"{NAME}/{relative}", date_time=(2026, 10, 6, 0, 0, 0))
            entry.compress_type = ZIP_DEFLATED
            entry.external_attr = 0o100644 << 16
            bundle.writestr(entry, payload)
    with ZipFile(archive) as bundle:
        if bundle.testzip() is not None:
            raise ValueError("archive integrity check failed")
    return {"plugin": str(root), "archive": str(archive), "marketplace": str(catalog),
            "files": len(files), "sha256": hashlib.sha256(archive.read_bytes()).hexdigest()}


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--out", required=True, type=Path, help="new or empty directory, usually ignored dist/")
    args = parser.parse_args()
    try:
        result = build(args.out)
    except (OSError, ValueError) as exc:
        parser.exit(1, f"packaging failed: {exc}\n")
    print(json.dumps(result, indent=2))


if __name__ == "__main__":
    main()
