"""Release regressions for bounded selection, exports, and resource cleanup."""

from __future__ import annotations

import csv
import io
import sys
import tempfile
import types
import unittest
from contextlib import redirect_stderr, redirect_stdout
from pathlib import Path
from unittest.mock import Mock, patch

import test_outlook_pst as fixtures

op = fixtures.op


class ReleaseSafetyTests(unittest.TestCase):
    def setUp(self):
        op.warnings_seen = 0
        self.tmp = tempfile.TemporaryDirectory(prefix="outlook-pst-release-test-")
        self.addCleanup(self.tmp.cleanup)
        self.pst = Path(self.tmp.name) / "sample.pst"
        self.pst.write_bytes(b"!BDN")

    def test_rejects_nonpositive_limits_before_opening_any_store(self):
        for command in (["list", str(self.pst)],
                        ["outlook", "delete", "--store", "example.pst", "--apply"]):
            for value in ("0", "-1"):
                with self.subTest(command=command, value=value), \
                     patch.object(op, "open_source") as source, \
                     patch.object(op, "outlook_namespace") as namespace, \
                     redirect_stderr(io.StringIO()), self.assertRaises(SystemExit):
                    op.main([*command, "--limit", value])
                source.assert_not_called()
                namespace.assert_not_called()

    def test_rejects_reversed_dates_and_empty_replacement(self):
        for args in (["list", str(self.pst), "--since", "2026-10-02", "--until", "2026-10-01"],
                     ["outlook", "edit", "--store", "example.pst", "--replace", "", "new", "--apply"]):
            with patch.object(op, "open_source") as source, \
                 patch.object(op, "outlook_namespace") as namespace, \
                 redirect_stderr(io.StringIO()), self.assertRaises(SystemExit):
                op.main(args)
            source.assert_not_called()
            namespace.assert_not_called()

    def test_pff_handle_closes_on_success_and_on_consumer_error(self):
        for fails in (False, True):
            pff = fixtures.fixture_file()
            pff.close = Mock()
            with patch.object(op, "open_pff", return_value=pff):
                try:
                    with op.open_source(self.pst, "pff"):
                        if fails:
                            raise RuntimeError("consumer failed")
                except RuntimeError:
                    self.assertTrue(fails)
            pff.close.assert_called_once_with()

    def test_failed_pff_open_preserves_error_without_closing_an_unopened_handle(self):
        pff = types.SimpleNamespace(open=Mock(side_effect=OSError("bad file")), close=Mock())
        with patch.dict(sys.modules, {"pypff": types.SimpleNamespace(file=lambda: pff)}), \
             self.assertRaises(OSError):
            op.open_pff(self.pst)
        pff.close.assert_not_called()  # Native libpff rejects close() after open() failed.

    def test_formula_text_is_neutralized_in_list_and_manifest_csv(self):
        pff = fixtures.fixture_file()
        pff.close = Mock()
        message = pff.root_folder.folders[0].folders[0].messages[0]
        message.subject = '=HYPERLINK("https://example.com", "click")'
        with patch.object(op, "open_pff", return_value=pff), \
             redirect_stdout(out := io.StringIO()), redirect_stderr(io.StringIO()):
            self.assertEqual(op.main(["list", str(self.pst), "--format", "csv", "--limit", "1"]), 0)
        self.assertTrue(next(csv.DictReader(io.StringIO(out.getvalue())))["subject"].startswith("'="))
        target = Path(self.tmp.name) / "export"
        with patch.object(op, "open_pff", return_value=pff), \
             redirect_stdout(io.StringIO()), redirect_stderr(io.StringIO()):
            self.assertEqual(op.main(["export", str(self.pst), str(target), "--limit", "1"]), 0)
        with (target / "manifest.csv").open(encoding="utf-8") as handle:
            self.assertTrue(next(csv.DictReader(handle))["subject"].startswith("'="))
        summary = next(target.rglob("Message.txt")).read_text(encoding="utf-8")
        self.assertIn(message.subject, summary)  # Evidence itself is not altered.

    def test_rtf_surrogate_pairs_export_as_unicode_without_encoding_errors(self):
        self.assertEqual(op.rtf_to_body(rb"{\rtf1\ansi\u-10179?\u-8704?}"), ("text", "😀"))
        self.assertEqual(op.rtf_to_body(rb"{\rtf1\ansi\u-10179?}"), ("text", "�"))

    def test_warning_state_is_reset_for_each_cli_invocation(self):
        op.warnings_seen = 1
        with patch.object(op, "cmd_stores"), redirect_stderr(io.StringIO()):
            self.assertEqual(op.main(["outlook", "stores"]), 0)

    def test_delete_guard_checks_folder_boundary(self):
        deleted = types.SimpleNamespace(FolderPath=r"\\Example\Deleted Items")
        parent = types.SimpleNamespace(FolderPath=r"\\Example\Deleted Items archive",
                                       Store=types.SimpleNamespace(GetDefaultFolder=lambda _: deleted))
        item = types.SimpleNamespace(Parent=parent, Subject="Example")
        plans = []
        with patch.object(op, "run_mutation", side_effect=lambda args, describe, act: plans.append(describe(item))):
            op.cmd_delete(types.SimpleNamespace())
        self.assertEqual(plans, ["move to Deleted Items"])

    def test_explicit_ids_reject_non_mail_and_duplicates_before_mutation(self):
        store = types.SimpleNamespace(StoreID="store-id")
        item = types.SimpleNamespace(Class=40)  # ContactItem, not MailItem.
        namespace = types.SimpleNamespace(GetItemFromID=Mock(return_value=item))
        args = types.SimpleNamespace(store="example.pst", id=["A"])
        with patch.object(op, "find_store", return_value=store), self.assertRaises(SystemExit):
            op.select_com(namespace, args)
        args.id = ["A", "a"]
        namespace.GetItemFromID.reset_mock()
        with patch.object(op, "find_store", return_value=store), self.assertRaises(SystemExit):
            op.select_com(namespace, args)
        namespace.GetItemFromID.assert_not_called()

    def test_attach_rejects_non_pst_without_loading_outlook(self):
        with patch.object(op, "outlook_namespace") as namespace, self.assertRaises(SystemExit):
            op.cmd_attach(types.SimpleNamespace(pst=Path(self.tmp.name) / "sample.ost", create=True))
        namespace.assert_not_called()

    def test_live_preview_does_not_call_the_mutation(self):
        item = types.SimpleNamespace(Subject="Example")
        record = types.SimpleNamespace(subject="Example", folder="Inbox")
        act = Mock()
        with patch.object(op, "outlook_namespace"), \
             patch.object(op, "select_com", return_value=[(record, item)]), \
             redirect_stdout(out := io.StringIO()), redirect_stderr(io.StringIO()):
            op.run_mutation(types.SimpleNamespace(apply=False), lambda _: "mark read", act)
        self.assertIn("would: mark read", out.getvalue())
        act.assert_not_called()

    def test_live_apply_calls_the_mutation_once_per_selected_item(self):
        item = types.SimpleNamespace(Subject="Example")
        record = types.SimpleNamespace(subject="Example", folder="Inbox")
        act = Mock()
        with patch.object(op, "outlook_namespace"), \
             patch.object(op, "select_com", return_value=[(record, item)]), \
             redirect_stdout(out := io.StringIO()):
            op.run_mutation(types.SimpleNamespace(apply=True), lambda _: "mark read", act)
        act.assert_called_once_with(item)
        self.assertIn("applied: mark read", out.getvalue())

    def test_deleted_items_and_descendants_are_never_deleted(self):
        deleted = types.SimpleNamespace(FolderPath=r"\\Example\Deleted Items")
        for folder_path in (deleted.FolderPath, deleted.FolderPath + r"\Child"):
            item = types.SimpleNamespace(Subject="Example", Delete=Mock(),
                                         Parent=types.SimpleNamespace(FolderPath=folder_path,
                                         Store=types.SimpleNamespace(GetDefaultFolder=lambda _: deleted)))
            record = types.SimpleNamespace(subject="Example", folder=folder_path)
            with patch.object(op, "outlook_namespace"), \
                 patch.object(op, "select_com", return_value=[(record, item)]), \
                 redirect_stdout(io.StringIO()), redirect_stderr(io.StringIO()):
                op.cmd_delete(types.SimpleNamespace(apply=True))
            item.Delete.assert_not_called()

    def test_com_source_ignores_non_mail_in_iteration_and_counts(self):
        folder = types.SimpleNamespace(Name="Example", Folders=[],
                    Items=[types.SimpleNamespace(Class=40), types.SimpleNamespace(Class=op.OL_MAIL)])
        source = op.ComSource(types.SimpleNamespace(GetRootFolder=lambda: folder))
        with patch.object(op, "ComMessage", return_value="mail"):
            self.assertEqual(list(source.messages()), ["mail"])
        self.assertEqual(list(source.folders()), [(0, "Example", 1)])

    def test_csv_prefixes_are_neutralized_but_regular_text_stays_exact(self):
        for value in ("=1+1", "+cmd", "-cmd", "@SUM(A1)", "  =1", "\tplain", "\rplain", "\nplain"):
            self.assertEqual(op.csv_safe_row({"value": value})["value"], "'" + value)
        self.assertEqual(op.csv_safe_row({"value": "Invoice", "count": 2}), {"value": "Invoice", "count": 2})

    def test_export_rejects_file_destination_without_reading_mail(self):
        target = Path(self.tmp.name) / "keep.txt"
        target.write_bytes(b"keep")
        with patch.object(op, "open_source") as source, self.assertRaises(SystemExit):
            op.cmd_export(types.SimpleNamespace(out=target))
        self.assertEqual(target.read_bytes(), b"keep")
        source.assert_not_called()


if __name__ == "__main__":
    unittest.main()
