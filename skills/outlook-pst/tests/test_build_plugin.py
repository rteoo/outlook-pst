"""Packaging regressions; all generated artifacts use cleaned temporary dirs."""

from __future__ import annotations

import importlib.util
import json
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch
from zipfile import ZipFile

SOURCE = Path(__file__).resolve().parents[1]
spec = importlib.util.spec_from_file_location("build_plugin", SOURCE / "scripts" / "build_plugin.py")
builder = importlib.util.module_from_spec(spec)
spec.loader.exec_module(builder)


class PluginBuildTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory(prefix="outlook-pst-plugin-test-")
        self.addCleanup(self.tmp.cleanup)
        self.out = Path(self.tmp.name) / "build"

    def test_relocated_package_runs_cli_and_existing_regressions(self):
        result = builder.build(self.out)
        root = Path(result["plugin"])
        for relative in ("scripts/outlook_pst.py", "tests/test_outlook_pst.py"):
            self.assertEqual((root / "skills" / builder.NAME / relative).read_bytes(),
                             (SOURCE / relative).read_bytes())
        script = root / "skills" / builder.NAME / "scripts" / "outlook_pst.py"
        help_result = subprocess.run([sys.executable, "-B", str(script), "--help"],
                                     cwd=self.tmp.name, capture_output=True, text=True, timeout=30)
        self.assertEqual(help_result.returncode, 0, help_result.stderr)
        self.assertIn("outlook", help_result.stdout)
        test_path = root / "skills" / builder.NAME / "tests"
        tests = subprocess.run([sys.executable, "-B", "-W", "error::ResourceWarning", "-m", "unittest",
                                "discover", "-s", str(test_path)],
                               cwd=self.tmp.name, capture_output=True, text=True, timeout=30)
        self.assertEqual(tests.returncode, 0, tests.stderr)
        self.assertIn("OK", tests.stderr)

    def test_archive_contains_only_allowlisted_files_and_hidden_overlay(self):
        result = builder.build(self.out)
        with ZipFile(result["archive"]) as archive:
            expected = {f"{builder.NAME}/{p}" for p in builder.package_files(SOURCE)}
            self.assertEqual(set(archive.namelist()), expected)
            self.assertIsNone(archive.testzip())
            self.assertIn("outlook-pst/.codex-plugin/plugin.json", archive.namelist())
            self.assertEqual(len(archive.namelist()), 15)
            for relative in (".claude-plugin/plugin.json", ".cursor-plugin/plugin.json",
                             "docs/plugin-installation.md", "assets/outlook-pst-icon.png",
                             "README.md", "SECURITY.md", "docs/review.md"):
                self.assertIn(f"outlook-pst/{relative}", archive.namelist())
            self.assertFalse(any(name.endswith("marketplace.json") for name in archive.namelist()))
        # Unrelated source files, including sensitive artifacts, never get copied.
        source = Path(self.tmp.name) / "source"
        for relative in builder.SOURCE_FILES:
            path = source / relative
            path.parent.mkdir(parents=True, exist_ok=True)
            path.write_bytes((SOURCE / relative).read_bytes())
        (source / "private.pst").write_bytes(b"not-a-real-mail-file")
        (source / ".env").write_bytes(b"FAKE_FIXTURE=do-not-package")
        self.assertEqual(set(builder.package_files(source)), set(builder.package_files(SOURCE)))

    def test_manifests_agree_and_marketplace_resolves_plugin(self):
        result = builder.build(self.out)
        root = Path(result["plugin"])
        portable = json.loads((root / "plugin.json").read_bytes())
        overlay = json.loads((root / ".codex-plugin" / "plugin.json").read_bytes())
        for key in ("name", "version", "description", "author"):
            self.assertEqual(portable[key], overlay[key])
        interface = portable["extensions"]["com.openai"]["interface"]
        self.assertEqual(interface, overlay["interface"])
        self.assertLessEqual(len(interface["shortDescription"]), 30)
        self.assertNotIn("skills", portable)
        self.assertEqual(overlay["skills"], "./skills/")
        catalog = json.loads(Path(result["marketplace"]).read_bytes())
        self.assertEqual((self.out / catalog["plugins"][0]["source"]["path"]).resolve(), root)
        skill = (root / "skills" / builder.NAME / "SKILL.md").read_text(encoding="utf-8")
        self.assertNotIn("teo-box", skill)
        self.assertIn("relative to this `SKILL.md`", skill)
        self.assertIn("--apply", skill)

    def test_branding_documentation_and_all_client_catalogs_survive_relocation(self):
        result = builder.build(self.out)
        root = Path(result["plugin"])
        self.assertEqual((root / "README.md").read_bytes(),
                         (builder.PROJECT_ROOT / "README.md").read_bytes())
        icon = (root / "assets/outlook-pst-icon.png").read_bytes()
        self.assertEqual(icon[:8], b"\x89PNG\r\n\x1a\n")
        self.assertEqual(icon, (builder.PROJECT_ROOT / "assets/outlook-pst-icon.png").read_bytes())
        self.assertEqual(set(result["marketplaces"]), {"codex", "claude", "cursor"})
        for client, path in result["marketplaces"].items():
            catalog = json.loads(Path(path).read_bytes())
            entry = catalog["plugins"][0]
            source = entry["source"]["path"] if client == "codex" else entry["source"]
            self.assertEqual((self.out / source).resolve(), root)
            if client != "codex":
                metadata = json.loads((root / f".{client}-plugin/plugin.json").read_bytes())
                self.assertEqual(entry["version"], metadata["version"])
                self.assertTrue((root / metadata["skills"] / "outlook-pst/SKILL.md").is_file())

    def test_metadata_drift_and_missing_artwork_fail_before_output(self):
        project = (Path(self.tmp.name) / "project").resolve()
        for relative in (*builder.PACKAGE_FILES, *builder.CATALOG_PATHS):
            path = project / relative
            path.parent.mkdir(parents=True, exist_ok=True)
            path.write_bytes((builder.PROJECT_ROOT / relative).read_bytes())
        metadata_path = project / ".claude-plugin/plugin.json"
        metadata = json.loads(metadata_path.read_bytes())
        metadata["version"] = "999.0.0"
        metadata_path.write_text(json.dumps(metadata), encoding="utf-8")
        with patch.object(builder, "PROJECT_ROOT", project):
            with self.assertRaisesRegex(ValueError, "claude manifest mismatch: version"):
                builder.build(self.out)
        self.assertFalse(self.out.exists())
        metadata["version"] = builder.VERSION
        metadata_path.write_text(json.dumps(metadata), encoding="utf-8")
        with patch.object(builder, "PROJECT_ROOT", project), patch.object(
            builder, "PACKAGE_FILES", tuple(p for p in builder.PACKAGE_FILES if not p.endswith(".png"))
        ):
            with self.assertRaisesRegex(ValueError, "packaged asset"):
                builder.build(self.out)
        self.assertFalse(self.out.exists())

    def test_catalog_drift_fails_before_output(self):
        catalog = json.loads((builder.PROJECT_ROOT / builder.CATALOG_PATHS[1]).read_bytes())
        catalog["plugins"][0]["version"] = "999.0.0"
        original = Path.read_bytes
        def read_bytes(path):
            if path == builder.PROJECT_ROOT / builder.CATALOG_PATHS[1]:
                return json.dumps(catalog).encode("utf-8")
            return original(path)
        with patch.object(Path, "read_bytes", read_bytes):
            with self.assertRaisesRegex(ValueError, "catalog source or version mismatch"):
                builder.build(self.out)
        self.assertFalse(self.out.exists())

    def test_refuses_nonempty_output_without_changing_existing_data(self):
        builder.build(self.out)
        archive = self.out / f"outlook-pst-{builder.VERSION}.zip"
        original = archive.read_bytes()
        with self.assertRaisesRegex(ValueError, "new or empty"):
            builder.build(self.out)
        self.assertEqual(archive.read_bytes(), original)

    def test_refuses_output_in_source_and_missing_source_before_writing(self):
        with self.assertRaisesRegex(ValueError, "outside the skill source"):
            builder.build(SOURCE / "generated-plugin")
        self.assertFalse((SOURCE / "generated-plugin").exists())
        with self.assertRaisesRegex(ValueError, "contained regular file"):
            builder.build(self.out, Path(self.tmp.name) / "missing-source")
        self.assertFalse(self.out.exists())

    def test_repeated_builds_have_identical_archive_bytes(self):
        first = builder.build(self.out)
        second = builder.build(Path(self.tmp.name) / "other-build")
        self.assertEqual(first["sha256"], second["sha256"])


if __name__ == "__main__":
    unittest.main()
