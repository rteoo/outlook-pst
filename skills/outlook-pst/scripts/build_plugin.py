#!/usr/bin/env python3
"""Build the local Outlook PST/OST skills plugin from repository source.

Usage: python -B skills/outlook-pst/scripts/build_plugin.py --out dist/plugin-0.2.0
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
# Package documentation and assets explicitly; never copy whole directories.
PACKAGE_FILES = (
    "plugin.json", ".codex-plugin/plugin.json", ".claude-plugin/plugin.json",
    ".cursor-plugin/plugin.json", "README.md", "LICENSE", "SECURITY.md",
    "docs/review.md", "docs/plugin-installation.md",
    "assets/outlook-pst-icon.png", "assets/README.md",
)
CATALOG_PATHS = (
    ".agents/plugins/marketplace.json",
    ".claude-plugin/marketplace.json",
    ".cursor-plugin/marketplace.json",
)



def json_bytes(value: dict) -> bytes:
    return (json.dumps(value, indent=2, ensure_ascii=False) + "\n").encode("utf-8")


def package_files(source: Path) -> dict[str, bytes]:
    files = {}
    for relative in PACKAGE_FILES:
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
    for client in ("claude", "cursor"):
        metadata = json.loads(files[f".{client}-plugin/plugin.json"])
        for key in ("name", "version", "description", "author", "license", "repository"):
            if portable.get(key) != metadata.get(key):
                raise ValueError(f"{client} manifest mismatch: {key}")
        if metadata.get("skills") != "./skills/":
            raise ValueError(f"{client} skills must resolve to ./skills/")
    for value in (interface.get("composerIcon"), interface.get("logo"),
                  json.loads(files[".cursor-plugin/plugin.json"]).get("logo")):
        if not isinstance(value, str) or value not in (
            "./assets/outlook-pst-icon.png", "assets/outlook-pst-icon.png"
        ) or value.removeprefix("./") not in files:
            raise ValueError("plugin icon must reference the packaged asset")
    if portable["name"] != NAME or portable["version"] != VERSION:
        raise ValueError("plugin metadata changed during packaging; run the builder again")
    for relative in SOURCE_FILES:
        path = source / relative
        if not path.is_file() or path.is_symlink() or path.resolve() != source.resolve() / relative:
            raise ValueError(f"source must be a contained regular file: {relative}")
        files[f"skills/{NAME}/{relative}"] = path.read_bytes()
    return files


def marketplace_files(plugin_path: str) -> dict[str, bytes]:
    catalogs = {}
    for relative in CATALOG_PATHS:
        path = PROJECT_ROOT / relative
        if not path.is_file() or path.is_symlink() or path.resolve() != PROJECT_ROOT / relative:
            raise ValueError(f"catalog must be a contained regular file: {relative}")
        catalog = json.loads(path.read_bytes())
        entries = catalog.get("plugins", [])
        if catalog.get("name") != "outlook-pst-local" or len(entries) != 1:
            raise ValueError(f"invalid single-plugin catalog: {relative}")
        entry = entries[0]
        if entry.get("name") != NAME:
            raise ValueError(f"catalog plugin mismatch: {relative}")
        if relative.startswith(".agents/"):
            if entry.get("source") != {"source": "local", "path": "./"}:
                raise ValueError("repository catalog must point to its root")
            entry["source"]["path"] = plugin_path
        else:
            if entry.get("source") != "./" or entry.get("version") != VERSION:
                raise ValueError(f"catalog source or version mismatch: {relative}")
            entry["source"] = plugin_path
        catalogs[relative] = json_bytes(catalog)
    return catalogs


def build(out: Path, source: Path = SOURCE) -> dict:
    files = package_files(source)  # Validate all inputs before creating output.
    catalogs = marketplace_files(f"./{NAME}")
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
    for relative, payload in catalogs.items():
        path = out / relative
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_bytes(payload)
    catalog = out / CATALOG_PATHS[0]
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
            "marketplaces": {client: str(out / relative)
                             for client, relative in zip(("codex", "claude", "cursor"), CATALOG_PATHS, strict=True)},
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
