"""Release regressions for bounded selection, exports, and resource cleanup."""

from __future__ import annotations

import csv
import io
import json
import sys
import tempfile
import types
import unittest
from contextlib import redirect_stderr, redirect_stdout
from pathlib import Path
from unittest.mock import Mock, PropertyMock, patch

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
        record = types.SimpleNamespace(id="A", subject="Example", folder="Inbox")
        act = Mock()
        with patch.object(op, "outlook_namespace"), \
             patch.object(op, "select_com", return_value=[(record, item)]), \
             redirect_stdout(out := io.StringIO()), redirect_stderr(io.StringIO()):
            op.run_mutation(types.SimpleNamespace(apply=False), lambda _: "mark read", act)
        self.assertIn("would: mark read", out.getvalue())
        act.assert_not_called()

    def test_live_apply_calls_the_mutation_once_per_selected_item(self):
        item = types.SimpleNamespace(Subject="Example")
        record = types.SimpleNamespace(id="A", subject="Example", folder="Inbox")
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
            record = types.SimpleNamespace(id="A", subject="Example", folder=folder_path)
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

    def test_failed_mutation_logs_only_completed_items(self):
        selection = [(types.SimpleNamespace(id=value, subject="Example", folder="Inbox"), value)
                     for value in ("A", "B")]
        for fail_at in (1, 2):
            effects = [RuntimeError("failed")] if fail_at == 1 else [None, RuntimeError("failed")]
            with patch.object(op, "outlook_namespace"), \
                 patch.object(op, "select_com", return_value=selection), \
                 redirect_stdout(out := io.StringIO()), self.assertRaises(RuntimeError):
                op.run_mutation(types.SimpleNamespace(apply=True), lambda _: "mark read", Mock(side_effect=effects))
            self.assertEqual(out.getvalue().count("applied:"), fail_at - 1)
            self.assertNotIn("[id=B]", out.getvalue())

    def test_edit_preview_and_apply_share_sequential_replacements(self):
        for set_subject, expected in ((None, "C"), ("A A", "C C"), ("", "")):
            for apply in (False, True):
                item = types.SimpleNamespace(Subject="A", Body="A", BodyFormat=1, Save=Mock())
                record = types.SimpleNamespace(id="A", subject="A", folder="Inbox")
                args = types.SimpleNamespace(apply=apply, set_subject=set_subject,
                        replace=[("A", "B"), ("B", "C")], mark=None, add_category=None)
                with patch.object(op, "outlook_namespace"), \
                     patch.object(op, "select_com", return_value=[(record, item)]), \
                     redirect_stdout(out := io.StringIO()), redirect_stderr(io.StringIO()):
                    op.cmd_edit(args)
                hits = 3 if set_subject == "A A" else 1 if set_subject == "" else 2
                self.assertIn(f"replace 'A'->'B' x{hits}", out.getvalue())
                self.assertIn(f"replace 'B'->'C' x{hits}", out.getvalue())
                self.assertIn("[id=A]", out.getvalue())
                self.assertEqual(item.Subject, expected if apply else "A")
                self.assertEqual(item.Body, "C" if apply else "A")
                self.assertEqual(item.Save.call_count, int(apply))

    def test_move_validates_selection_before_creating_destination(self):
        args = types.SimpleNamespace(apply=True, create=True, store="example.pst",
                                     to_store=None, to_folder="New", id=["missing"])
        with patch.object(op, "outlook_namespace"), \
             patch.object(op, "find_store"), \
             patch.object(op, "select_com", side_effect=SystemExit("invalid selection")), \
             patch.object(op, "find_folder") as folder, self.assertRaises(SystemExit):
            op.cmd_move(args)
        folder.assert_not_called()
        with patch.object(op, "outlook_namespace"), \
             patch.object(op, "find_store"), \
             patch.object(op, "select_com", return_value=[]), \
             patch.object(op, "find_folder") as folder:
            op.cmd_move(args)
        folder.assert_not_called()

    def test_move_materializes_valid_selection_once_then_moves(self):
        item = types.SimpleNamespace(Move=Mock())
        record = types.SimpleNamespace(id="A", subject="Example", folder="Inbox")
        destination = types.SimpleNamespace(FolderPath="New")
        args = types.SimpleNamespace(apply=True, create=True, store="example.pst",
                                     to_store=None, to_folder="New")
        with patch.object(op, "outlook_namespace"), \
             patch.object(op, "select_com", return_value=[(record, item)]) as select, \
             patch.object(op, "find_store"), \
             patch.object(op, "find_folder", return_value=destination) as folder, \
             redirect_stdout(io.StringIO()):
            op.cmd_move(args)
        select.assert_called_once()
        self.assertTrue(folder.call_args.kwargs["create"])
        item.Move.assert_called_once_with(destination)

    def test_terminal_controls_are_escaped_in_human_output_but_json_stays_exact(self):
        value = "invoice\x1b]0;fake\x07\r\n\t\u202eexe\u2028\u2029"
        record = op.Record("A", value, value, value, "", "", None, 0)
        with redirect_stdout(out := io.StringIO()), redirect_stderr(io.StringIO()):
            op.print_records(iter([record]), "table")
        self.assertNotIn("\x1b", out.getvalue())
        self.assertNotIn("\u202e", out.getvalue())
        self.assertNotIn("\u2028", out.getvalue())
        self.assertNotIn("\u2029", out.getvalue())
        self.assertIn(r"\x1b]0;fake\x07\x0d\x0a\x09\u202eexe", out.getvalue())
        with redirect_stdout(out := io.StringIO()), redirect_stderr(io.StringIO()):
            op.print_records(iter([record]), "json")
        self.assertEqual(json.loads(out.getvalue())["subject"], value)
        with redirect_stderr(err := io.StringIO()):
            op.warn(value)
        self.assertNotIn("\x1b", err.getvalue())
        self.assertEqual(len(err.getvalue().splitlines()), 1)
        with patch.object(op, "outlook_namespace"), \
             patch.object(op, "select_com", return_value=[(record, object())]), \
             redirect_stdout(out := io.StringIO()), redirect_stderr(io.StringIO()):
            op.run_mutation(types.SimpleNamespace(apply=False), lambda _: value, Mock())
        self.assertNotIn("\x1b", out.getvalue())
        self.assertNotIn("\u202e", out.getvalue())
        self.assertEqual(len(out.getvalue().splitlines()), 1)

    def test_rtf_unicode_fallback_consumes_escaped_characters(self):
        for fallback in (rb"\{", rb"\}", rb"\\", rb"\~", rb"\_", rb"\-", rb"\'3f", b"?", rb"\tab ", rb"\b "):
            with self.subTest(fallback=fallback):
                self.assertEqual(op.rtf_to_body(rb"{\rtf1\ansi\uc1\u233" + fallback + b"X}"), ("text", "éX"))
        self.assertEqual(op.rtf_to_body(rb"{\rtf1\ansi\uc2\u233\{?X}"), ("text", "éX"))
        self.assertEqual(op.rtf_to_body(rb"{\rtf1\ansi{\uc1\u233}X}"), ("text", "éX"))
        with self.assertRaises(op.RtfParseError):
            op.rtf_to_body(rb"{\rtf1\ansi\uc1\u233\bin1 ?X}")

    def test_malformed_rtf_is_a_controlled_parse_failure(self):
        for control in (rb"\u hello", rb"\uc hello", rb"\uc-1 hello", b"\\u" + b"1" * 5000):
            with self.subTest(control=control[:30]), self.assertRaises(op.RtfParseError):
                op.rtf_to_body(rb"{\rtf1\ansi" + control + b"}")

    def test_malformed_rtf_does_not_stop_search_or_export_of_later_mail(self):
        pff = fixtures.fixture_file()
        message = pff.root_folder.folders[0].folders[0].messages[0]
        message.plain_text_body = b""
        message.rtf_body = rb"{\rtf1\ansi\u hello}"
        with patch.object(op, "open_pff", return_value=pff), \
             redirect_stdout(out := io.StringIO()), redirect_stderr(err := io.StringIO()):
            self.assertEqual(op.main(["list", str(self.pst), "--text", "archive", "--format", "json"]), 1)
        self.assertIn('"id": "303"', out.getvalue())
        self.assertIn("warning:", err.getvalue())
        for fmt in ("dir", "eml"):
            target = Path(self.tmp.name) / fmt
            with patch.object(op, "open_pff", return_value=pff), \
                 redirect_stdout(io.StringIO()), redirect_stderr(err := io.StringIO()):
                self.assertEqual(op.main(["export", str(self.pst), str(target), "--format", fmt]), 1)
            self.assertIn("exported 2 message(s)", err.getvalue())
            self.assertFalse(any("202" in path.name for path in target.rglob("*")))
            self.assertTrue(any("303" in path.name for path in target.rglob("*")))

    def test_attachment_metadata_failure_preserves_later_attachments(self):
        message = fixtures.fixture_file().root_folder.folders[0].folders[1].messages[0]
        readable = message.attachments[1]
        message.get_attachment = Mock(side_effect=[OSError("broken metadata"), readable, message.attachments[2]])
        view = op.PffMessage("Sent", message)
        with redirect_stderr(err := io.StringIO()):
            attachments = list(view.attachments())
        self.assertEqual(attachments[0], ("attachment-1", None))
        self.assertEqual(attachments[1], ("CON.txt", b"reserved"))
        self.assertEqual(attachments[2], ("proposta.pdf", b"second copy"))
        self.assertIn("attachment #1", err.getvalue())

    def test_edits_without_replacements_do_not_read_body(self):
        class Item:
            Subject = "Example"
            Categories = ""
            Save = Mock()

        for edit in ({"mark": "read"}, {"set_subject": "Changed"}, {"add_category": "Reviewed"}):
            item = Item()
            args = types.SimpleNamespace(apply=True, set_subject=None, replace=None,
                                         mark=None, add_category=None)
            args.__dict__.update(edit)
            record = types.SimpleNamespace(id="A", subject="Example", folder="Inbox")
            with patch.object(Item, "Body", new_callable=PropertyMock, create=True,
                              side_effect=RuntimeError("body inaccessible")) as body, \
                 patch.object(op, "outlook_namespace"), \
                 patch.object(op, "select_com", return_value=[(record, item)]), \
                 redirect_stdout(io.StringIO()):
                op.cmd_edit(args)
            body.assert_not_called()


if __name__ == "__main__":
    unittest.main()
