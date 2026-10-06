"""Packaging regressions; all generated artifacts use cleaned temporary dirs."""

from __future__ import annotations

import importlib.util
import json
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path
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
            self.assertEqual(len(archive.namelist()), 8)
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

    def test_refuses_nonempty_output_without_changing_existing_data(self):
        builder.build(self.out)
        archive = self.out / "outlook-pst-0.1.0.zip"
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
