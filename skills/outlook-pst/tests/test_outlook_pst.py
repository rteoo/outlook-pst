"""Regressions for outlook_pst.py using duck-typed libpff and Outlook COM fakes.

The live Outlook COM backend is exercised manually against a scratch PST; these
tests cover selection, export fidelity, and failure reporting without Outlook.
"""

from __future__ import annotations

import csv
import hashlib
import importlib.util
import io
import json
import sys
import tempfile
import types
import unittest
from contextlib import redirect_stderr, redirect_stdout
from datetime import datetime, timezone
from email import policy
from email.parser import BytesParser
from pathlib import Path
from unittest.mock import Mock, patch

SCRIPT = Path(__file__).resolve().parents[1] / "scripts" / "outlook_pst.py"
spec = importlib.util.spec_from_file_location("outlook_pst", SCRIPT)
op = importlib.util.module_from_spec(spec)
sys.modules["outlook_pst"] = op
spec.loader.exec_module(op)


class Entry:
    """int -> PT_LONG, str -> PT_UNICODE, bytes -> PT_STRING8 that pypff would mis-decode as cp1252."""

    def __init__(self, value):
        self.value_type = {int: 0x3, str: 0x1F, bytes: 0x1E}[type(value)]
        self.data = value if isinstance(value, bytes) else None
        self.data_as_integer = value if isinstance(value, int) else None
        self.data_as_string = value if isinstance(value, str) else value.decode("cp1252", "replace") if self.data else None


class RecordSet:
    def __init__(self, props: dict):
        self.props = props

    def get_entry_by_type(self, tag):
        return Entry(self.props[tag]) if tag in self.props else None


class Attachment:
    def __init__(self, name, payload, broken=False):
        self.long_filename, self.payload, self.broken = name, payload, broken
        self.size = len(payload)

    def read_buffer(self, size):
        if self.broken:
            raise OSError("pypff_attachment_read_buffer: unable to read")
        return self.payload[:size]


class Message:
    def __init__(self, identifier, subject, sender, when, plain=b"", html=None, headers=None,
                 props=None, recipients=(), attachments=()):
        self.identifier, self.subject, self.sender_name = identifier, subject, sender
        self.delivery_time = when
        self.client_submit_time = self.creation_time = None
        self.plain_text_body, self.html_body, self.rtf_body = plain, html, None
        self.transport_headers = headers
        self.record_sets = [RecordSet(props or {})]
        self.recipients = types.SimpleNamespace(record_sets=[RecordSet(r) for r in recipients]) if recipients else None
        self.attachments = list(attachments)
        self.number_of_attachments = len(self.attachments)

    def get_attachment(self, index):
        return self.attachments[index]


class Folder:
    def __init__(self, name, messages=(), folders=()):
        self.name, self.messages, self.folders = name, list(messages), list(folders)
        self.number_of_sub_messages = len(self.messages)
        self.number_of_sub_folders = len(self.folders)

    def get_sub_message(self, index):
        return self.messages[index]

    def get_sub_folder(self, index):
        return self.folders[index]


RECEIVED_HEADERS = (
    "Received: from mx.example.com by mail.example.org\r\n"
    "From: Bob <bob@example.org>\r\nTo: Alex <ana@example.com>\r\n"
    "Subject: Q3 report\r\nMessage-ID: <q3@example.org>\r\n"
    "Date: Mon, 05 Oct 2026 12:00:00 +0000\r\n"
    "MIME-Version: 1.0\r\nContent-Type: text/plain; charset=us-ascii\r\n\r\n"
)


def fixture_file():
    sent = Message(
        101, "Proposta – revisão", "René", datetime(2026, 9, 1, 15, 0),
        plain="Olá, segue a proposta.".encode("utf-8"),
        html="<p>Olá, segue a <b>proposta</b>.</p>".encode("cp1252"),
        props={op.PR_SENDER_SMTP_ADDRESS: "rene@example.com", op.PR_DISPLAY_TO: "Alex",
               op.PR_INTERNET_CPID: 1252, op.PR_INTERNET_MESSAGE_ID: "<p101@example.com>"},
        recipients=[{op.PR_RECIPIENT_TYPE: 1, op.PR_DISPLAY_NAME: "Alex", op.PR_SMTP_ADDRESS: "ana@example.com"},
                    {op.PR_RECIPIENT_TYPE: 2, op.PR_DISPLAY_NAME: "José", op.PR_EMAIL_ADDRESS: "jose@example.com"}],
        attachments=[Attachment("proposta.pdf", b"%PDF-1.7 fake"), Attachment("CON.txt", b"reserved"),
                     Attachment("proposta.pdf", b"second copy")],
    )
    received = Message(
        202, "Q3 report", "Bob", datetime(2026, 10, 5, 12, 0),
        plain=b"Numbers attached. invoice-7781", headers=RECEIVED_HEADERS,
        props={op.PR_SENDER_SMTP_ADDRESS: "bob@example.org", op.PR_DISPLAY_TO: "Alex"},
    )
    old = Message(303, "Old note", "Carl", datetime(2025, 1, 2, 9, 0), plain=b"archive")
    top = Folder("Top of Personal Folders", folders=[
        Folder("Inbox", [received, old], folders=[Folder("Clients", [])]),
        Folder("Sent Items", [sent]),
    ])
    return types.SimpleNamespace(root_folder=Folder(None, folders=[top]), number_of_orphan_items=0, close=Mock())


class OutlookPstTests(unittest.TestCase):
    def setUp(self):
        op.warnings_seen = 0
        self.tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)
        self.pst = str(Path(self.tmp.name) / "x.pst")
        Path(self.pst).write_bytes(b"!BDN")
        self.pff = fixture_file()
        patcher = patch.object(op, "open_pff", return_value=self.pff)
        patcher.start()
        self.addCleanup(patcher.stop)

    def run_cli(self, *argv):
        out, err = io.StringIO(), io.StringIO()
        with redirect_stdout(out), redirect_stderr(err):
            code = op.main(list(argv))
        return code, out.getvalue(), err.getvalue()

    def listed_ids(self, *filters):
        code, out, _ = self.run_cli("list", self.pst, "--format", "json", *filters)
        self.assertEqual(code, 0)
        return [json.loads(line)["id"] for line in out.splitlines()]

    def test_filters_select_by_folder_sender_dates_body_and_attachments(self):
        self.assertEqual(self.listed_ids(), ["202", "303", "101"])
        self.assertEqual(self.listed_ids("--folder", "inbox"), ["202", "303"])
        self.assertEqual(self.listed_ids("--from", "EXAMPLE.ORG"), ["202"])
        self.assertEqual(self.listed_ids("--since", "2026-01-01", "--until", "2026-09-30"), ["101"])
        self.assertEqual(self.listed_ids("--text", "invoice-7781"), ["202"])
        self.assertEqual(self.listed_ids("--has-attachments"), ["101"])
        self.assertEqual(self.listed_ids("--to", "alex", "--limit", "1"), ["202"])

    def test_list_reports_folder_path_and_combined_sender(self):
        _, out, _ = self.run_cli("list", self.pst, "--format", "json", "--subject", "proposta")
        row = json.loads(out)
        self.assertEqual(row["folder"], "Top of Personal Folders/Sent Items")
        self.assertEqual(row["sender"], "René <rene@example.com>")
        self.assertEqual(row["date"], "2026-09-01T15:00:00+00:00")

    def test_eml_export_round_trips_bodies_recipients_and_attachments(self):
        out = Path(self.tmp.name) / "eml"
        code, _, _ = self.run_cli("export", self.pst, str(out), "--format", "eml", "--folder", "sent")
        self.assertEqual(code, 0)
        eml = out / "Top of Personal Folders" / "Sent Items" / "101.eml"
        msg = BytesParser(policy=policy.default).parsebytes(eml.read_bytes())
        self.assertEqual(msg["Subject"], "Proposta – revisão")
        self.assertEqual(msg["From"], "René <rene@example.com>")
        self.assertEqual(msg["To"], "Alex <ana@example.com>")
        self.assertEqual(msg["Cc"], "José <jose@example.com>")
        self.assertEqual(msg["Message-ID"], "<p101@example.com>")
        self.assertIn("segue a proposta", msg.get_body(("plain",)).get_content())
        self.assertIn("<b>proposta</b>", msg.get_body(("html",)).get_content())
        payloads = {part.get_filename(): part.get_payload(decode=True) for part in msg.iter_attachments()}
        self.assertEqual(payloads["proposta.pdf"], b"second copy")  # last same-name part wins in the dict
        self.assertEqual(payloads["CON.txt"], b"reserved")

        with (out / "manifest.csv").open(encoding="utf-8") as handle:
            rows = list(csv.DictReader(handle))
        self.assertEqual(len(rows), 1)
        self.assertEqual(rows[0]["sha256"], hashlib.sha256(eml.read_bytes()).hexdigest())

    def test_eml_export_keeps_original_transport_headers_but_rebuilds_mime(self):
        out = Path(self.tmp.name) / "eml"
        self.run_cli("export", self.pst, str(out), "--format", "eml", "--subject", "Q3")
        raw = (out / "Top of Personal Folders" / "Inbox" / "202.eml").read_bytes()
        msg = BytesParser(policy=policy.default).parsebytes(raw)
        self.assertEqual(msg["Received"], "from mx.example.com by mail.example.org")
        self.assertEqual(msg["Message-ID"], "<q3@example.org>")
        self.assertEqual(msg.get_all("Content-Type"), ['text/plain; charset="utf-8"'])
        self.assertIn("invoice-7781", msg.get_content())

    def test_dir_export_matches_pffexport_layout_and_hashes_every_file(self):
        out = Path(self.tmp.name) / "dir"
        code, _, _ = self.run_cli("export", self.pst, str(out), "--subject", "proposta")
        self.assertEqual(code, 0)
        msg_dir = out / "Top of Personal Folders" / "Sent Items" / "Message101"
        self.assertIn("subject: Proposta – revisão", (msg_dir / "Message.txt").read_text(encoding="utf-8"))
        self.assertEqual((msg_dir / "Recipients.txt").read_text(encoding="utf-8"),
                         "to: Alex <ana@example.com>\ncc: José <jose@example.com>\n")
        self.assertEqual((msg_dir / "Body.html").read_text(encoding="utf-8"), "<p>Olá, segue a <b>proposta</b>.</p>")
        attachments = sorted(p.name for p in (msg_dir / "Attachments").iterdir())
        self.assertEqual(attachments, ["_CON.txt", "proposta (2).pdf", "proposta.pdf"])

        with (out / "manifest.csv").open(encoding="utf-8") as handle:
            rows = list(csv.DictReader(handle))
        for row in rows:
            self.assertEqual(row["md5"], hashlib.md5((out / row["path"]).read_bytes()).hexdigest())
        self.assertEqual(sum(row["kind"] == "attachment" for row in rows), 3)

    def test_export_refuses_a_non_empty_directory(self):
        out = Path(self.tmp.name) / "busy"
        out.mkdir()
        (out / "keep.txt").write_text("mine")
        with self.assertRaises(SystemExit) as caught:
            self.run_cli("export", self.pst, str(out))
        self.assertIn("not empty", str(caught.exception))
        self.assertEqual([p.name for p in out.iterdir()], ["keep.txt"])

    def test_outlook_written_utf8_string8_and_html_only_body(self):
        # Shape observed in a PST written by Outlook 16: PT_STRING8 holding utf-8 with a subject-prefix
        # marker and trailing bytes, no PR_BODY, and an unresolved recipient.
        draft = Message(
            404, "Invoice 1 â€“ aÃ§Ã£o", "", datetime(2026, 10, 6, 10, 57),
            plain=None, html="<html><head><style>p{}</style></head><body><p>Hello</p><p>token-9</p></body></html>".encode(),
            props={op.PR_SUBJECT: b"\x01\x01Invoice 1 \xe2\x80\x93 a\xc3\xa7\xc3\xa3o\x00\x00\x008\x00",
                   op.PR_INTERNET_CPID: 65001, op.PR_DISPLAY_TO: b"alice1@example.com"},
            recipients=[{op.PR_RECIPIENT_TYPE: 1, op.PR_DISPLAY_NAME: b"alice1@example.com", op.PR_EMAIL_ADDRESS: b""}],
        )
        pff = types.SimpleNamespace(root_folder=Folder(None, [draft]), number_of_orphan_items=0, close=Mock())
        with patch.object(op, "open_pff", return_value=pff):
            _, out, _ = self.run_cli("list", self.pst, "--format", "json", "--text", "token-9")
            self.assertEqual(json.loads(out)["subject"], "Invoice 1 – ação")
            _, out, _ = self.run_cli("show", self.pst, "404")
            shown = json.loads(out)
            self.assertEqual(shown["body"], "Hello\ntoken-9")
            self.assertEqual(shown["recipients"], [{"kind": "to", "name": "", "address": "alice1@example.com"}])
            eml_dir = Path(self.tmp.name) / "eml"
            self.run_cli("export", self.pst, str(eml_dir), "--format", "eml")
        msg = BytesParser(policy=policy.default).parsebytes((eml_dir / "404.eml").read_bytes())
        self.assertEqual(msg["Subject"], "Invoice 1 – ação")
        self.assertEqual(msg["To"], "alice1@example.com")

    def test_rtf_only_message_is_decoded_searched_and_exported_with_its_rtf(self):
        # Outlook stores mail imported from .msg as compressed RTF only: HTML encapsulated per [MS-OXRTFEX].
        rtf = (rb"{\rtf1\ansi\ansicpg65001\fromhtml1 \fbidis \deff0{\fonttbl{\f0\fswiss Arial;}}" b"\r\n"
               rb"{\*\htmltag19 <html>}{\*\htmltag64 <p>}\htmlrtf {\htmlrtf0 Ol\'c3\'a1 token-rtf\htmlrtf\par}"
               rb"\htmlrtf0" b"\r\n" rb"\htmlrtf \par" b"\r\n" rb"\htmlrtf0 {\*\htmltag72 </p>}{\*\htmltag27 </html>}}" b"\0")
        msg = Message(505, "RTF only", "Alex", datetime(2026, 10, 6, 11, 0), plain=None)
        msg.rtf_body = rtf
        pff = types.SimpleNamespace(root_folder=Folder(None, [msg]), number_of_orphan_items=0, close=Mock())
        out = Path(self.tmp.name) / "dir"
        with patch.object(op, "open_pff", return_value=pff):
            _, listed, _ = self.run_cli("list", self.pst, "--format", "json", "--text", "olá token-rtf")
            self.run_cli("export", self.pst, str(out))
        self.assertEqual(json.loads(listed)["id"], "505")
        self.assertEqual((out / "Message505" / "Body.html").read_text(encoding="utf-8"),
                         "<html><p>Olá token-rtf</p></html>")
        self.assertEqual((out / "Message505" / "Body.rtf").read_bytes(), rtf)

    def test_plain_rtf_becomes_text(self):
        rtf = rb"{\rtf1\ansi\ansicpg1252{\fonttbl{\f0 Arial;}}\f0 Ol\'e1 \u8211? mundo\par segunda {\*\generator x;}linha\tab fim\}}"
        self.assertEqual(op.rtf_to_body(rtf), ("text", "Olá – mundo\nsegunda linha\tfim}"))

    def test_unreadable_attachment_is_reported_and_fails_the_run(self):
        self.pff.root_folder.folders[0].folders[1].messages[0].attachments[0].broken = True
        out = Path(self.tmp.name) / "dir"
        code, _, err = self.run_cli("export", self.pst, str(out), "--subject", "proposta")
        self.assertEqual(code, 1)
        self.assertIn("'proposta.pdf' not readable", err)
        names = sorted(p.name for p in (out / "Top of Personal Folders" / "Sent Items" / "Message101" / "Attachments").iterdir())
        self.assertEqual(names, ["_CON.txt", "proposta.pdf"])


class StandaloneTests(unittest.TestCase):
    def setUp(self):
        op.warnings_seen = 0

    def test_libpff_lock_error_becomes_file_locked(self):
        class LockedFile:
            close = Mock()
            def open(self, path):
                raise OSError("unable to read from file with error: another process has locked a portion of the file.")

        with patch.dict(sys.modules, {"pypff": types.SimpleNamespace(file=LockedFile)}), \
             self.assertRaises(op.FileLocked):
            op.open_pff(Path("x.pst"))

    def test_safe_name_strips_path_syntax_and_reserved_names(self):
        self.assertEqual(op.safe_name('a/b:c*?.txt'), "a_b_c__.txt")
        self.assertEqual(op.safe_name("nul"), "_nul")
        self.assertEqual(op.safe_name(" ... "), "_")

    def test_open_maps_options_to_outlook_switches(self):
        with patch.object(op, "outlook_exe", return_value="OUTLOOK.EXE"), \
             patch.object(op.subprocess, "Popen") as popen, redirect_stdout(io.StringIO()):
            code = op.main(["outlook", "open", "--profile", "Work", "--safe", "--select", "outlook:inbox"])
        self.assertEqual(code, 0)
        popen.assert_called_once_with(["OUTLOOK.EXE", "/profile", "Work", "/safe", "/recycle", "/select", "outlook:inbox"])


class FakeComError(Exception):
    """Shape needed by the fake Outlook property accessor; no native COM dependency."""


pywintypes = types.SimpleNamespace(com_error=FakeComError)


class ComCollection(list):
    @property
    def Count(self):  # noqa: N802 - Outlook object model name
        return len(self)


class Accessor:
    def __init__(self, props):
        self.props = props

    def GetProperty(self, uri):  # noqa: N802
        tag = uri.rsplit("0x", 1)[1]
        if tag not in self.props:
            raise pywintypes.com_error(-2147221233, "property not found", None, None)
        return self.props[tag]


class ComAttachment:
    def __init__(self, name, payload):
        self.FileName, self.DisplayName, self.payload = name, name, payload

    def SaveAsFile(self, path):  # noqa: N802
        Path(path).write_bytes(self.payload)


def pst_entry_id(node_id: int) -> str:
    return (b"\0" * 4 + b"\x11" * 16 + node_id.to_bytes(4, "little")).hex().upper()


class OutlookHeldFileTests(unittest.TestCase):
    """The file is locked by Outlook, so file commands read the store through COM."""

    def setUp(self):
        op.warnings_seen = 0
        modules = patch.dict(sys.modules, {"pywintypes": pywintypes})
        modules.start()
        self.addCleanup(modules.stop)
        self.tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)
        self.path = Path(self.tmp.name) / "archive.pst"
        self.path.write_bytes(b"!BDN")
        item = types.SimpleNamespace(
            Class=op.OL_MAIL, EntryID=pst_entry_id(2097220), Subject="Invoice 1 – ação", SenderName="Bob",
            SenderEmailAddress="bob@example.org", SenderEmailType="SMTP", To="alice1@example.com", CC="",
            ReceivedTime=datetime(2026, 10, 6, 7, 57, 48, tzinfo=timezone.utc),  # pywin32 tags local time as UTC
            Attachments=ComCollection([ComAttachment("note.txt", b"payload\r\n")]),
            Body="Hello body 1\r\nsecret-token-1", HTMLBody="<p>Hello body 1</p>", BodyFormat=2,
            Recipients=[types.SimpleNamespace(Type=1, Name="alice1@example.com", Address="alice1@example.com",
                                              PropertyAccessor=Accessor({}))],
            PropertyAccessor=Accessor({"007D001F": RECEIVED_HEADERS, "1035001F": "<q3@example.org>"}),
        )
        inbox = types.SimpleNamespace(Name="Inbox Test", Folders=[], Items=ComCollection([item]))
        self.root = types.SimpleNamespace(Name="Outlook Data File", Folders=[inbox], Items=ComCollection())
        self.store = types.SimpleNamespace(FilePath=str(self.path.resolve()), GetRootFolder=lambda: self.root)
        self.ns = types.SimpleNamespace(Stores=[self.store], RemoveStore=Mock(),
                                        AddStoreEx=Mock(side_effect=lambda *_: self.ns.Stores.append(self.store)))
        for target, value in ((op, "open_pff"), (op, "outlook_namespace"), (op.sys, "platform")):
            replacement = {"open_pff": Mock(side_effect=op.FileLocked(f"{self.path} is locked")),
                           "outlook_namespace": Mock(return_value=self.ns), "platform": "win32"}[value]
            patcher = patch.object(target, value, replacement)
            patcher.start()
            self.addCleanup(patcher.stop)

    def run_cli(self, *argv):
        out, err = io.StringIO(), io.StringIO()
        with redirect_stdout(out), redirect_stderr(err):
            code = op.main(list(argv))
        return code, out.getvalue(), err.getvalue()

    def test_auto_reads_a_store_outlook_has_open_without_attaching(self):
        code, out, err = self.run_cli("list", str(self.path), "--format", "json", "--text", "secret-token-1")
        self.assertEqual(code, 0)
        row = json.loads(out)
        self.assertEqual(row["id"], "2097220")  # same id libpff reports for this node
        self.assertEqual(row["folder"], "Outlook Data File/Inbox Test")
        self.assertEqual(row["sender"], "Bob <bob@example.org>")
        self.assertEqual(row["message_id"], "<q3@example.org>")
        self.assertTrue(row["date"].startswith("2026-10-06T07:57:48"))
        self.assertIn("reading it through Outlook", err)
        self.ns.AddStoreEx.assert_not_called()

    def test_eml_export_through_outlook_keeps_headers_bodies_and_attachments(self):
        out = Path(self.tmp.name) / "eml"
        code, _, _ = self.run_cli("export", str(self.path), str(out), "--format", "eml")
        self.assertEqual(code, 0)
        msg = BytesParser(policy=policy.default).parsebytes((out / "Outlook Data File" / "Inbox Test" / "2097220.eml").read_bytes())
        self.assertEqual(msg["Received"], "from mx.example.com by mail.example.org")
        self.assertIn("secret-token-1", msg.get_body(("plain",)).get_content())
        self.assertIn("<p>Hello body 1</p>", msg.get_body(("html",)).get_content())
        self.assertEqual([(a.get_filename(), a.get_payload(decode=True)) for a in msg.iter_attachments()],
                         [("note.txt", b"payload\r\n")])

    def test_show_and_tree_through_outlook(self):
        _, out, _ = self.run_cli("show", str(self.path), "2097220")
        shown = json.loads(out)
        self.assertEqual(shown["recipients"], [{"kind": "to", "name": "", "address": "alice1@example.com"}])
        self.assertEqual(shown["attachment_names"], ["note.txt"])
        _, out, _ = self.run_cli("tree", str(self.path))
        self.assertEqual(out, "Outlook Data File  [0]\n  Inbox Test  [1]\n")

    def test_auto_refuses_to_attach_a_locked_file_outlook_does_not_have_open(self):
        self.ns.Stores.clear()
        with self.assertRaises(SystemExit) as caught:
            self.run_cli("list", str(self.path))
        self.assertIn("--via outlook", str(caught.exception))
        self.ns.AddStoreEx.assert_not_called()

    def test_via_outlook_attaches_for_the_read_and_detaches_after(self):
        self.ns.Stores.clear()
        code, out, err = self.run_cli("list", str(self.path), "--via", "outlook", "--format", "json")
        self.assertEqual(code, 0)
        self.assertEqual(json.loads(out)["id"], "2097220")
        self.ns.AddStoreEx.assert_called_once_with(str(self.path.resolve()), op.OL_STORE_UNICODE)
        self.ns.RemoveStore.assert_called_once_with(self.root)
        self.assertIn("detached afterwards", err)

    def test_via_pff_reports_the_lock_instead_of_using_outlook(self):
        with self.assertRaises(SystemExit) as caught:
            self.run_cli("list", str(self.path), "--via", "pff")
        self.assertIn("is locked", str(caught.exception))
        op.outlook_namespace.assert_not_called()


if __name__ == "__main__":
    unittest.main()
