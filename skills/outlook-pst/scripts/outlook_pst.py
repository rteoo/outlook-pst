#!/usr/bin/env python3
"""Read, search, export, and manipulate Outlook Classic .pst/.ost mail.

Purpose: two backends over Outlook data files.
  * Read-only file commands (tree, list, show, export): parse any .pst/.ost
    with libpff, including orphaned OSTs and archives from other machines; no
    Outlook needed, any OS. When the running Outlook holds the file open, they
    read the same store through Outlook instead. Export writes pffexport-style
    folders or .eml files plus an MD5/SHA-256 manifest.
  * Live commands (`outlook ...`, Windows + Classic Outlook over COM): list
    stores, attach/detach PSTs, and move, edit, categorize, soft-delete, or
    save items as .msg. Every mutation is a dry run unless --apply is given.
    Changes to an Exchange-backed store (.ost) sync to the server.

Usage (S = skills/outlook-pst/scripts/outlook_pst.py):
  uv run --with libpff-python --with pywin32 python S tree FILE.pst
  uv run --with libpff-python --with pywin32 python S list FILE.pst --from acme --since 2026-01-01
  uv run --with libpff-python --with pywin32 python S show FILE.pst 2097188
  uv run --with libpff-python --with pywin32 python S export FILE.pst OUT_DIR --format eml
  uv run --with pywin32 python S outlook stores
  uv run --with pywin32 python S outlook list --store archive.pst --folder @inbox --subject invoice
  uv run --with pywin32 python S outlook edit --store archive.pst --folder @inbox --subject invoice --replace ACME Acme --apply
  uv run --with pywin32 python S outlook open --select outlook:inbox
Cron: none; manual utility only.
Dependencies: libpff-python (module pypff) for file reads; pywin32 and
Classic Outlook for live commands and for reading files Outlook holds open
(the new Outlook has no COM API). Omit pywin32 off Windows.
"""

from __future__ import annotations

import argparse
import csv
import hashlib
import json
import mimetypes
import os
import re
import subprocess
import sys
import tempfile
from collections.abc import Callable, Iterator
from contextlib import contextmanager
from dataclasses import asdict, dataclass, field
from datetime import date, datetime, time, timezone
from email import encoders
from email.header import Header
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.parser import HeaderParser
from email.utils import format_datetime, formataddr
from functools import cached_property
from html.parser import HTMLParser
from pathlib import Path

# MAPI property tags read from libpff record sets.
PR_SUBJECT = 0x0037
PR_SENDER_NAME = 0x0C1A
PR_DISPLAY_TO = 0x0E04
PR_DISPLAY_CC = 0x0E03
PR_SENDER_EMAIL_ADDRESS = 0x0C1F
PR_SENDER_SMTP_ADDRESS = 0x5D01
PR_INTERNET_MESSAGE_ID = 0x1035
PR_INTERNET_CPID = 0x3FDE
PR_MESSAGE_CODEPAGE = 0x3FFD
PR_RECIPIENT_TYPE = 0x0C15
PR_DISPLAY_NAME = 0x3001
PR_EMAIL_ADDRESS = 0x3003
PR_SMTP_ADDRESS = 0x39FE
PT_STRING8 = 0x1E
RECIPIENT_KINDS = {1: "to", 2: "cc", 3: "bcc"}

# Outlook object-model constants.
OL_STORE_UNICODE = 3
OL_MSG_UNICODE = 9
OL_NOT_EXCHANGE = 3
OL_FORMAT_RICH_TEXT = 3
OL_FORMAT_HTML = 2
OL_MAIL = 43
DEFAULT_FOLDERS = {"@deleted": 3, "@outbox": 4, "@sent": 5, "@inbox": 6, "@drafts": 16, "@junk": 23}

WINDOWS_RESERVED = {"CON", "PRN", "AUX", "NUL", *(f"COM{i}" for i in range(1, 10)), *(f"LPT{i}" for i in range(1, 10))}

warnings_seen = 0


def warn(message: str) -> None:
    global warnings_seen
    warnings_seen += 1
    print(f"warning: {message}", file=sys.stderr)


@dataclass
class Record:
    id: str
    folder: str
    subject: str
    sender: str
    to: str
    cc: str
    date: datetime | None
    attachments: int
    extra: dict = field(default_factory=dict)

    def as_row(self) -> dict:
        row = asdict(self)
        row["date"] = self.date.isoformat(timespec="seconds") if self.date else ""
        row.update(row.pop("extra"))
        return row


# --------------------------------------------------------------------------
# Selection filters shared by both backends
# --------------------------------------------------------------------------


def positive_int(value: str) -> int:
    number = int(value)
    if number < 1:
        raise argparse.ArgumentTypeError("must be a positive integer")
    return number


def csv_safe_row(row: dict) -> dict:
    """Neutralize formula prefixes in spreadsheet-facing CSV, keeping evidence files exact."""
    def cell(value):
        if isinstance(value, str) and value.lstrip().startswith(("=", "+", "-", "@")):
            return "'" + value
        if isinstance(value, str) and value.startswith(("\t", "\r", "\n")):
            return "'" + value
        return value
    return {key: cell(value) for key, value in row.items()}


def add_filter_args(parser: argparse.ArgumentParser) -> None:
    parser.add_argument("--from", dest="sender", help="sender name/address contains (case-insensitive)")
    parser.add_argument("--to", help="To/Cc display contains")
    parser.add_argument("--subject", help="subject contains")
    parser.add_argument("--text", help="plain-text body contains (slower: reads every body)")
    parser.add_argument("--since", type=date.fromisoformat, help="on or after YYYY-MM-DD (local time)")
    parser.add_argument("--until", type=date.fromisoformat, help="on or before YYYY-MM-DD (local time)")
    parser.add_argument("--has-attachments", action="store_true")
    parser.add_argument("--limit", type=positive_int, help="stop after N matches (positive integer)")


def matches(rec: Record, args: argparse.Namespace, body: Callable[[], str]) -> bool:
    def contains(haystack: str, needle: str | None) -> bool:
        return needle is None or needle.casefold() in (haystack or "").casefold()

    if not (contains(rec.sender, args.sender) and contains(rec.subject, args.subject)):
        return False
    if args.to is not None and not contains(f"{rec.to} {rec.cc}", args.to):
        return False
    if args.has_attachments and not rec.attachments:
        return False
    if args.since or args.until:
        if rec.date is None:
            return False
        local = rec.date.astimezone()
        if args.since and local < datetime.combine(args.since, time.min).astimezone():
            return False
        if args.until and local > datetime.combine(args.until, time.max).astimezone():
            return False
    return contains(body(), args.text) if args.text is not None else True


def limited(items: Iterator, limit: int | None) -> Iterator:
    for count, item in enumerate(items, 1):
        yield item
        if limit and count >= limit:
            return


def print_records(records: Iterator[Record], fmt: str) -> int:
    count = 0
    if fmt == "csv":
        writer = None
        for rec in records:
            row = rec.as_row()
            if writer is None:
                writer = csv.DictWriter(sys.stdout, fieldnames=list(row), lineterminator="\n")
                writer.writeheader()
            writer.writerow(csv_safe_row(row))
            count += 1
    else:
        for rec in records:
            if fmt == "json":
                print(json.dumps(rec.as_row(), ensure_ascii=False))
            else:
                when = rec.date.astimezone().strftime("%Y-%m-%d %H:%M") if rec.date else "-" * 16
                clip = "@" if rec.attachments else " "
                print(f"{rec.id}\t{when}\t{clip}\t{rec.sender[:30]:30}\t{rec.subject[:70]}\t[{rec.folder}]")
            count += 1
    print(f"{count} message(s)", file=sys.stderr)
    return count


# --------------------------------------------------------------------------
# Offline backend: libpff
# --------------------------------------------------------------------------


class FileLocked(Exception):
    pass


def open_pff(path: Path):
    try:
        import pypff
    except ImportError:
        sys.exit("pypff is missing: run with `uv run --with libpff-python --with pywin32 python outlook_pst.py ...`")
    pff = pypff.file()
    try:
        pff.open(str(path))
    except OSError as exc:
        if "locked" in str(exc):
            raise FileLocked(f"{path} is locked by another process") from exc
        raise
    return pff


def record_value(record_set, tag: int, codepage: int | None = None):
    """Return a string or integer MAPI property from a libpff record set, or None when absent."""
    entry = record_set.get_entry_by_type(tag)
    if entry is None:
        return None
    if entry.value_type in (0x2, 0x3, 0x14):
        return entry.data_as_integer
    if entry.value_type == PT_STRING8:
        # pypff decodes 8-bit strings with a fixed cp1252; Outlook writes them in the message codepage (often utf-8).
        return decode(bytes(entry.data or b"").split(b"\0", 1)[0], codepage)
    return entry.data_as_string


def pff_entry(item, tag: int, codepage: int | None = None):
    for record_set in item.record_sets:
        value = record_value(record_set, tag, codepage)
        if value is not None:
            return value
    return None


def message_codepage(msg) -> int | None:
    return pff_entry(msg, PR_MESSAGE_CODEPAGE) or pff_entry(msg, PR_INTERNET_CPID)


def decode(data: bytes | None, codepage: int | None = None) -> str:
    if not data:
        return ""
    declared = ["utf-8" if codepage == 65001 else f"cp{codepage}"] if codepage else []
    for encoding in [*declared, "utf-8"]:
        try:
            return data.decode(encoding)
        except (LookupError, UnicodeDecodeError):
            continue
    # Undeclared non-utf-8 Outlook text is Windows-1252 in practice; U+FFFD marks any residual loss visibly.
    return data.decode("cp1252", errors="replace")


class _TextExtractor(HTMLParser):
    BLOCKS = frozenset({"p", "div", "br", "tr", "li", "h1", "h2", "h3", "h4", "h5", "h6", "table", "blockquote"})
    HIDDEN = frozenset({"head", "script", "style"})

    def __init__(self) -> None:
        super().__init__()
        self.parts: list[str] = []
        self.hidden = 0

    def handle_starttag(self, tag, attrs) -> None:
        if tag in self.HIDDEN:
            self.hidden += 1
        elif tag in self.BLOCKS:
            self.parts.append("\n")

    def handle_endtag(self, tag) -> None:
        if tag in self.HIDDEN and self.hidden:
            self.hidden -= 1

    def handle_data(self, data) -> None:
        if not self.hidden:
            self.parts.append(data)


RTF_TOKEN = re.compile(rb"\\([a-zA-Z]+)(-?\d+)? ?|\\'([0-9a-fA-F]{2})|\\(.)|([{}])|([\r\n]+)|([^\\{}\r\n]+)", re.DOTALL)
RTF_SKIPPED_DESTINATIONS = {b"fonttbl", b"colortbl", b"stylesheet", b"info", b"pict", b"object", b"header", b"footer"}


def rtf_to_body(rtf: bytes) -> tuple[str, str]:
    """Return ("html", ...) for HTML-encapsulated RTF ([MS-OXRTFEX] \\fromhtml1), else ("text", ...).

    Outlook stores most rich mail as compressed RTF only, so this is the body when PR_BODY/PR_HTML are absent.
    """
    rtf = rtf.rstrip(b"\0")  # libpff hands back the stored buffer with its NUL terminator
    codepage = int(m.group(1)) if (m := re.search(rb"\\ansicpg(\d+)", rtf[:400])) else 1252
    is_html = b"\\fromhtml1" in rtf[:400]
    out: list[str] = []
    pending = bytearray()  # \'hh bytes are decoded together so multi-byte codepages survive
    # Per group: [suppressed by \htmlrtf, skipped destination, inside \htmltag, \uc skip count]
    stack = [[False, False, False, 1]]
    skip_chars = 0

    def flush() -> None:
        if pending:
            out.append(decode(bytes(pending), codepage))
            pending.clear()

    def emitting() -> bool:
        suppressed, skipped, in_tag, _ = stack[-1]
        return not skipped and (in_tag or not suppressed)

    for token in RTF_TOKEN.finditer(rtf):
        word, arg, hex_byte, symbol, brace, _newline, text = token.groups()  # raw newlines carry no text in RTF
        if hex_byte:
            if skip_chars:
                skip_chars -= 1
            elif emitting():
                pending.append(int(hex_byte, 16))
            continue
        flush()
        if brace == b"{":
            stack.append(list(stack[-1]))
        elif brace == b"}":
            if len(stack) > 1:
                stack.pop()
        elif symbol:
            if symbol == b"*":
                stack[-1][1] = True  # unknown destination unless the next word is \htmltag
            elif emitting() and symbol in b"\\{}":
                out.append(symbol.decode())
        elif word:
            state = stack[-1]
            if word == b"htmltag":
                state[1], state[2] = False, True
            elif word in RTF_SKIPPED_DESTINATIONS:
                state[1] = True
            elif word == b"htmlrtf":
                state[0] = arg != b"0"
            elif word == b"uc":
                state[3] = int(arg or 1)
            elif word == b"u" and emitting():
                out.append(chr(int(arg) % 65536))
                skip_chars = state[3]
            elif word in (b"par", b"line") and emitting():
                out.append("\r\n" if is_html else "\n")
            elif word == b"tab" and emitting():
                out.append("\t")
        elif text:
            chunk = text
            if skip_chars:
                chunk, skip_chars = chunk[skip_chars:], max(0, skip_chars - len(chunk))
            if emitting():
                pending.extend(chunk)
    flush()
    # RTF encodes supplementary characters as signed UTF-16 surrogate pairs.
    body = "".join(out).encode("utf-16-le", "surrogatepass").decode("utf-16-le", "replace")
    return ("html", body) if is_html else ("text", body.strip())


def html_to_text(html: str) -> str:
    parser = _TextExtractor()
    parser.feed(html)
    parser.close()
    lines = (" ".join(line.split()) for line in "".join(parser.parts).splitlines())
    return re.sub(r"\n{3,}", "\n\n", "\n".join(lines)).strip()


def walk_pff(folder, path: tuple[str, ...] = ()) -> Iterator[tuple[str, object]]:
    here = path + (folder.name,) if folder.name else path
    folder_path = "/".join(here)
    for index in range(folder.number_of_sub_messages):
        try:
            yield folder_path, folder.get_sub_message(index)
        except OSError as exc:
            warn(f"unreadable message #{index} in {folder_path!r}: {exc}")
    for index in range(folder.number_of_sub_folders):
        try:
            sub = folder.get_sub_folder(index)
        except OSError as exc:
            warn(f"unreadable subfolder #{index} in {folder_path!r}: {exc}")
            continue
        yield from walk_pff(sub, here)


def format_sender(name: str, address: str) -> str:
    if address and address.casefold() != name.casefold():
        return f"{name} <{address}>" if name else address
    return name


def make_recipient(kind: str, name: str, address: str) -> dict:
    if (not address or address.casefold() == name.casefold()) and "@" in name:
        name, address = "", name  # an unresolved recipient keeps the typed address as its display name
    return {"kind": kind, "name": name, "address": address}


class MessageView:
    """Backend-neutral read view of one message: PffMessage (libpff) or ComMessage (running Outlook)."""

    id: str
    folder: str

    def text(self) -> str:
        """Plain-text body, derived from the HTML body when the store kept only HTML."""
        plain, html = self.bodies()
        return plain or html_to_text(html)


class PffMessage(MessageView):
    def __init__(self, folder: str, msg) -> None:
        self.folder, self.msg, self.id = folder, msg, str(msg.identifier)

    @cached_property
    def codepage(self) -> int | None:
        return message_codepage(self.msg)

    def sender(self) -> tuple[str, str]:
        msg, cp = self.msg, self.codepage
        name = pff_entry(msg, PR_SENDER_NAME, cp) or msg.sender_name or ""
        return name, pff_entry(msg, PR_SENDER_SMTP_ADDRESS, cp) or pff_entry(msg, PR_SENDER_EMAIL_ADDRESS, cp) or ""

    def record(self) -> Record:
        msg, cp = self.msg, self.codepage
        subject = pff_entry(msg, PR_SUBJECT, cp) or msg.subject or ""
        when = msg.delivery_time or msg.client_submit_time or msg.creation_time
        return Record(
            id=self.id,
            folder=self.folder,
            # A stored subject may start with 0x01 plus one character encoding the "RE: "-style prefix length.
            subject=subject[2:] if subject.startswith("\x01") else subject,
            sender=format_sender(*self.sender()),
            to=pff_entry(msg, PR_DISPLAY_TO, cp) or "",
            cc=pff_entry(msg, PR_DISPLAY_CC, cp) or "",
            date=when.replace(tzinfo=timezone.utc) if when else None,  # libpff returns naive UTC
            attachments=msg.number_of_attachments,
            extra={"message_id": pff_entry(msg, PR_INTERNET_MESSAGE_ID, cp) or ""},
        )

    def bodies(self) -> tuple[str, str]:
        plain = decode(self.msg.plain_text_body)
        html = decode(self.msg.html_body, pff_entry(self.msg, PR_INTERNET_CPID))
        if not (plain or html) and self.rtf():
            kind, body = rtf_to_body(self.rtf())
            return ("", body) if kind == "html" else (body, "")
        return plain, html

    def rtf(self) -> bytes:
        return self.msg.rtf_body or b""

    def headers(self) -> str:
        return self.msg.transport_headers or ""

    def recipients(self) -> list[dict]:
        if self.msg.recipients is None:
            return []
        cp = self.codepage
        return [
            make_recipient(
                RECIPIENT_KINDS.get(record_value(rs, PR_RECIPIENT_TYPE), "to"),
                record_value(rs, PR_DISPLAY_NAME, cp) or "",
                record_value(rs, PR_SMTP_ADDRESS, cp) or record_value(rs, PR_EMAIL_ADDRESS, cp) or "",
            )
            for rs in self.msg.recipients.record_sets
        ]

    def attachment_names(self) -> list[str]:
        return [self.msg.get_attachment(i).long_filename or f"attachment-{i + 1}" for i in range(self.msg.number_of_attachments)]

    def attachments(self) -> Iterator[tuple[str, bytes | None]]:
        """Yield (filename, payload); payload is None when libpff cannot read it (e.g. an embedded message)."""
        for index, name in enumerate(self.attachment_names()):
            try:
                attachment = self.msg.get_attachment(index)
                yield name, attachment.read_buffer(attachment.size) if attachment.size else b""
            except OSError as exc:
                warn(f"message {self.id}: attachment {name!r} not readable ({exc})")
                yield name, None


class PffSource:
    def __init__(self, pff) -> None:
        self.pff = pff

    def messages(self) -> Iterator[MessageView]:
        for folder_path, msg in walk_pff(self.pff.root_folder):
            yield PffMessage(folder_path, msg)

    def folders(self, folder=None, depth: int = 0) -> Iterator[tuple[int, str, int]]:
        folder = self.pff.root_folder if folder is None else folder
        yield depth, folder.name or "(root)", folder.number_of_sub_messages
        for index in range(folder.number_of_sub_folders):
            yield from self.folders(folder.get_sub_folder(index), depth + 1)

    def note(self) -> str:
        return f"orphan items: {self.pff.number_of_orphan_items}"


@contextmanager
def open_source(path: Path, via: str) -> Iterator[PffSource | ComSource]:
    """Read with libpff; when Outlook holds the file, read the store through the running Outlook.

    Outlook writes to a PST when it attaches one, so `auto` only reads stores Outlook already has open.
    Re-attaching a detached-but-locked PST takes an explicit `--via outlook`.
    """
    path = path.resolve()
    if not path.is_file():
        sys.exit(f"not a file: {path}")
    pff = None
    if via != "outlook":
        try:
            pff = open_pff(path)
        except FileLocked as exc:
            if via == "pff" or sys.platform != "win32":
                sys.exit(f"{exc}: close the program holding it, or read a copy taken while it was closed")
    if pff is not None:
        try:
            yield PffSource(pff)
        finally:
            pff.close()
        return

    ns = outlook_namespace()
    store = store_for_path(ns, path)
    if store is None and via == "auto":
        sys.exit(
            f"{path} is locked but not open in this Outlook profile (Outlook keeps a detached PST locked until it exits). "
            "Close Outlook, or pass --via outlook to attach it temporarily; attaching lets Outlook write to the file."
        )
    attached = store is None
    if attached:
        if path.suffix.lower() != ".pst":
            sys.exit(f"only a .pst can be attached; {path.name} must belong to the running Outlook profile")
        ns.AddStoreEx(str(path), OL_STORE_UNICODE)
        store = store_for_path(ns, path)
        print(f"attached {path.name} to Outlook for this read; it is detached afterwards", file=sys.stderr)
    else:
        print(f"{path.name} is open in Outlook; reading it through Outlook", file=sys.stderr)
    try:
        yield ComSource(store)
    finally:
        if attached:
            ns.RemoveStore(store.GetRootFolder())


def select(source, args: argparse.Namespace) -> Iterator[tuple[Record, MessageView]]:
    def generate():
        for view in source.messages():
            if args.folder and args.folder.casefold() not in view.folder.casefold():
                continue
            try:
                rec = view.record()
                if matches(rec, args, view.text):
                    yield rec, view
            except OSError as exc:
                warn(f"skipped message in {view.folder!r}: {exc}")
    return limited(generate(), args.limit)


def cmd_tree(args: argparse.Namespace) -> None:
    with open_source(args.file, args.via) as source:
        total = 0
        for depth, name, count in source.folders():
            print(f"{'  ' * depth}{name}  [{count}]")
            total += count
        print(f"{total} message(s); {source.note()}", file=sys.stderr)


def cmd_list(args: argparse.Namespace) -> None:
    with open_source(args.file, args.via) as source:
        print_records((rec for rec, _ in select(source, args)), args.format)


def cmd_show(args: argparse.Namespace) -> None:
    with open_source(args.file, args.via) as source:
        view = next((v for v in source.messages() if v.id == args.id), None)
        if view is None:
            sys.exit(f"no message with id {args.id}")
        print(json.dumps({
            **view.record().as_row(),
            "recipients": view.recipients(),
            "attachment_names": view.attachment_names(),
            "headers": view.headers(),
            "body": view.bodies()[1] if args.html else view.text(),
        }, ensure_ascii=False, indent=2))


def safe_name(name: str, limit: int = 80) -> str:
    cleaned = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", name).strip().rstrip(". ")[:limit].rstrip(". ")
    if not cleaned:
        return "_"
    return f"_{cleaned}" if cleaned.split(".")[0].upper() in WINDOWS_RESERVED else cleaned


def unique_path(path: Path) -> Path:
    candidate, n = path, 1
    while candidate.exists() or candidate.is_symlink():
        n += 1
        candidate = path.with_name(f"{path.stem} ({n}){path.suffix}")
    return candidate


def header_value(value: str) -> str | Header:
    return value if value.isascii() else Header(value, "utf-8")


def build_eml(rec: Record, view: MessageView) -> bytes:
    # ceiling: MIME assembly holds attachments in RAM; use streaming for messages exceeding available memory.
    plain, html = view.bodies()
    attachments = list(view.attachments())
    body = MIMEMultipart("alternative") if plain and html else None
    if body:
        body.attach(MIMEText(plain, "plain", "utf-8"))
        body.attach(MIMEText(html, "html", "utf-8"))
    else:
        body = MIMEText(html, "html", "utf-8") if html else MIMEText(plain, "plain", "utf-8")
    if attachments:
        root = MIMEMultipart("mixed")
        root.attach(body)
        for name, payload in attachments:
            if payload is None:
                continue
            maintype, _, subtype = (mimetypes.guess_type(name)[0] or "application/octet-stream").partition("/")
            part = MIMEBase(maintype, subtype)
            part.set_payload(payload)
            encoders.encode_base64(part)
            part.add_header("Content-Disposition", "attachment", filename=("utf-8", "", name) if not name.isascii() else name)
            root.attach(part)
    else:
        root = body

    original = view.headers()
    if original.strip():
        # Received mail keeps its original routing headers; the MIME structure is rebuilt above.
        for key, value in HeaderParser().parsestr(original, headersonly=True).items():
            if not key.lower().startswith("content-") and key.lower() != "mime-version":
                root[key] = value
    else:
        recipients = view.recipients()
        for kind in ("to", "cc", "bcc"):
            values = [formataddr((r["name"], r["address"]), charset="utf-8") for r in recipients if r["kind"] == kind]
            if values:
                root[kind.capitalize()] = ", ".join(values)
        if any(view.sender()):
            root["From"] = formataddr(view.sender(), charset="utf-8")
        if rec.subject:
            root["Subject"] = header_value(rec.subject)
        if rec.date:
            root["Date"] = format_datetime(rec.date)
        if rec.extra.get("message_id"):
            root["Message-ID"] = rec.extra["message_id"]
    return root.as_bytes()


def write_file(path: Path, data: bytes, manifest: csv.DictWriter, rec: Record, kind: str, out: Path) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path = unique_path(path)
    with path.open("xb") as handle:
        handle.write(data)
    manifest.writerow(csv_safe_row({
        "kind": kind, "message_id": rec.id, "folder": rec.folder, "date": rec.as_row()["date"],
        "from": rec.sender, "subject": rec.subject, "path": path.relative_to(out).as_posix(),
        "size": len(data), "md5": hashlib.md5(data).hexdigest(), "sha256": hashlib.sha256(data).hexdigest(),
    }))


def cmd_export(args: argparse.Namespace) -> None:
    out = args.out.resolve()
    if out.exists():
        if not out.is_dir():
            sys.exit(f"output must be a directory: {out}")
        if any(out.iterdir()):
            sys.exit(f"output directory is not empty: {out}")
    fields = ["kind", "message_id", "folder", "date", "from", "subject", "path", "size", "md5", "sha256"]
    count = 0
    with open_source(args.file, args.via) as source:
        out.mkdir(parents=True, exist_ok=True)
        with (out / "manifest.csv").open("x", newline="", encoding="utf-8") as handle:
            manifest = csv.DictWriter(handle, fieldnames=fields)
            manifest.writeheader()
            for rec, view in select(source, args):
                folder_dir = out.joinpath(*(safe_name(part) for part in rec.folder.split("/") if part))
                if args.format == "eml":
                    write_file(folder_dir / f"{rec.id}.eml", build_eml(rec, view), manifest, rec, "eml", out)
                else:
                    write_message_dir(folder_dir / f"Message{rec.id}", rec, view, manifest, out)
                count += 1
    print(f"exported {count} message(s) to {out}", file=sys.stderr)


def write_message_dir(target: Path, rec: Record, view: MessageView, manifest: csv.DictWriter, out: Path) -> None:
    """pffexport-style layout: one folder per message with headers, bodies, recipients, attachments."""
    summary = "\n".join(f"{k}: {v}" for k, v in rec.as_row().items()) + "\n"
    write_file(target / "Message.txt", summary.encode("utf-8"), manifest, rec, "summary", out)
    recipients = "".join(f"{r['kind']}: {r['name']} <{r['address']}>\n" for r in view.recipients())
    write_file(target / "Recipients.txt", recipients.encode("utf-8"), manifest, rec, "recipients", out)
    if view.headers():
        write_file(target / "InternetHeaders.txt", view.headers().encode("utf-8"), manifest, rec, "headers", out)
    plain, html = view.bodies()
    if plain:
        write_file(target / "Body.txt", plain.encode("utf-8"), manifest, rec, "body", out)
    if html:
        write_file(target / "Body.html", html.encode("utf-8"), manifest, rec, "body", out)
    if view.rtf():  # kept as returned by the backend; text/HTML may be decoded from it
        write_file(target / "Body.rtf", view.rtf(), manifest, rec, "body", out)
    for name, payload in view.attachments():
        if payload is not None:
            write_file(target / "Attachments" / safe_name(name, 120), payload, manifest, rec, "attachment", out)


# --------------------------------------------------------------------------
# Live backend: Outlook COM
# --------------------------------------------------------------------------


def outlook_namespace():
    if sys.platform != "win32":
        sys.exit("outlook commands need Windows with Classic Outlook")
    try:
        import win32com.client
    except ImportError:
        sys.exit("pywin32 is missing: run with `uv run --with libpff-python --with pywin32 python outlook_pst.py ...`")
    return win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")


def find_store(ns, key: str):
    wanted = key.casefold()
    absolute = os.path.abspath(key).casefold()
    found = [
        store for store in ns.Stores
        if store.DisplayName.casefold() == wanted
        or (store.FilePath and (store.FilePath.casefold() == absolute or Path(store.FilePath).name.casefold() == wanted))
    ]
    if len(found) != 1:
        sys.exit(f"{len(found)} stores match {key!r}; run `outlook stores` and pass the exact path")
    return found[0]


def find_folder(store, path: str, create: bool = False):
    parts = [p for p in re.split(r"[\\/]", path) if p]
    if parts and parts[0].lower() in DEFAULT_FOLDERS:
        folder = store.GetDefaultFolder(DEFAULT_FOLDERS[parts.pop(0).lower()])
    else:
        folder = store.GetRootFolder()
    for part in parts:
        child = next((f for f in folder.Folders if f.Name.casefold() == part.casefold()), None)
        if child is None:
            if not create:
                names = ", ".join(f.Name for f in folder.Folders)
                sys.exit(f"no folder {part!r} under {folder.FolderPath}; children: {names}")
            child = folder.Folders.Add(part)
        folder = child
    return folder


def com_date(item) -> datetime | None:
    value = getattr(item, "ReceivedTime", None) or getattr(item, "CreationTime", None)
    # pywin32 tags Outlook's local wall-clock time as UTC; drop the tag and read it as local.
    return value.replace(tzinfo=None).astimezone() if value else None


def com_sender(item) -> tuple[str, str]:
    address = getattr(item, "SenderEmailAddress", "") or ""
    if getattr(item, "SenderEmailType", "") == "EX":
        user = item.Sender.GetExchangeUser() if item.Sender else None
        address = user.PrimarySmtpAddress if user else address
    return getattr(item, "SenderName", "") or "", address


def com_property(obj, tag: str) -> str:
    """Read a MAPI string property (tag like '007D001F'); '' when the item does not carry it."""
    import pywintypes

    try:
        return obj.PropertyAccessor.GetProperty(f"http://schemas.microsoft.com/mapi/proptag/0x{tag}") or ""
    except pywintypes.com_error:
        return ""


def com_record(item, record_id: str | None = None, folder: str | None = None) -> Record:
    return Record(
        id=record_id or item.EntryID,
        folder=folder or item.Parent.FolderPath,
        subject=getattr(item, "Subject", "") or "",
        sender=format_sender(*com_sender(item)),
        to=getattr(item, "To", "") or "",
        cc=getattr(item, "CC", "") or "",
        date=com_date(item),
        attachments=item.Attachments.Count,
        extra={"message_id": com_property(item, "1035001F")},
    )


class ComMessage(MessageView):
    def __init__(self, folder: str, item) -> None:
        self.folder, self.item = folder, item
        raw = bytes.fromhex(item.EntryID)
        # A PST entry ID is flags(4) + provider UID(16) + node ID(4); the node ID is libpff's message
        # identifier, so ids match across backends. Exchange entry IDs have no such mapping.
        self.id = str(int.from_bytes(raw[20:], "little")) if len(raw) == 24 else item.EntryID

    def sender(self) -> tuple[str, str]:
        return com_sender(self.item)

    def record(self) -> Record:
        return com_record(self.item, self.id, self.folder)

    def bodies(self) -> tuple[str, str]:
        item = self.item
        html = item.HTMLBody if getattr(item, "BodyFormat", None) == OL_FORMAT_HTML else ""
        return getattr(item, "Body", "") or "", html or ""

    def rtf(self) -> bytes:
        return bytes(self.item.RTFBody) if getattr(self.item, "BodyFormat", None) == OL_FORMAT_RICH_TEXT else b""

    def headers(self) -> str:
        return com_property(self.item, "007D001F")

    def recipients(self) -> list[dict]:
        return [
            make_recipient(RECIPIENT_KINDS.get(r.Type, "to"), r.Name or "", com_property(r, "39FE001F") or r.Address or "")
            for r in getattr(self.item, "Recipients", [])
        ]

    def attachment_names(self) -> list[str]:
        return [a.FileName or a.DisplayName or f"attachment-{i}" for i, a in enumerate(self.item.Attachments, 1)]

    def attachments(self) -> Iterator[tuple[str, bytes | None]]:
        import pywintypes

        names = self.attachment_names()
        with tempfile.TemporaryDirectory() as tmp:
            for index, (name, attachment) in enumerate(zip(names, self.item.Attachments, strict=True), 1):
                target = Path(tmp) / str(index)
                try:
                    attachment.SaveAsFile(str(target))
                    payload = target.read_bytes()
                except pywintypes.com_error as exc:
                    warn(f"message {self.id}: attachment {name!r} not readable ({exc})")
                    payload = None
                yield name, payload


class ComSource:
    def __init__(self, store) -> None:
        self.store = store

    def _walk(self, folder, path: str = "", depth: int = 0):
        here = f"{path}/{folder.Name}" if path else folder.Name
        yield depth, here, folder
        for sub in folder.Folders:
            yield from self._walk(sub, here, depth + 1)

    def messages(self) -> Iterator[MessageView]:
        # ceiling: one COM round-trip per property; fine for archives, slow for 100k+ item mailboxes
        for _, path, folder in self._walk(self.store.GetRootFolder()):
            for item in folder.Items:
                if getattr(item, "Class", None) == OL_MAIL:
                    yield ComMessage(path, item)

    def folders(self) -> Iterator[tuple[int, str, int]]:
        for depth, _, folder in self._walk(self.store.GetRootFolder()):
            count = sum(getattr(item, "Class", None) == OL_MAIL for item in folder.Items)
            yield depth, folder.Name, count

    def note(self) -> str:
        return "read through the running Outlook; orphan and hidden items are not visible"


def store_for_path(ns, path: Path):
    wanted = str(path).casefold()
    return next((s for s in ns.Stores if (s.FilePath or "").casefold() == wanted), None)


def walk_com(folder, recursive: bool) -> Iterator:
    # ceiling: client-side scan of every item; switch to Items.Restrict/AdvancedSearch if folders hit 100k+ items
    yield from folder.Items
    if recursive:
        for sub in folder.Folders:
            yield from walk_com(sub, recursive)


def select_com(ns, args: argparse.Namespace) -> list[tuple[Record, object]]:
    store = find_store(ns, args.store)
    if args.id:
        if len({entry_id.casefold() for entry_id in args.id}) != len(args.id):
            sys.exit("duplicate --id values are not allowed")
        items = [ns.GetItemFromID(entry_id, store.StoreID) for entry_id in args.id]
        if any(getattr(item, "Class", None) != OL_MAIL for item in items):
            sys.exit("selected --id must identify a MailItem; non-mail items cannot be changed")
        return [(com_record(item), item) for item in items]
    folder = find_folder(store, args.folder)
    selected = (
        (rec, item)
        for item in walk_com(folder, args.recursive)
        if getattr(item, "Class", None) == OL_MAIL
        for rec in [com_record(item)]
        if matches(rec, args, lambda item=item: getattr(item, "Body", "") or "")
    )
    return list(limited(selected, args.limit))  # materialize before mutating the collection


def cmd_stores(args: argparse.Namespace) -> None:
    for store in outlook_namespace().Stores:
        kind = "pst" if store.ExchangeStoreType == OL_NOT_EXCHANGE and store.FilePath else "exchange"
        print(f"{store.DisplayName}\t{kind}\t{store.FilePath or '-'}")


def cmd_attach(args: argparse.Namespace) -> None:
    path = args.pst.resolve()
    if path.suffix.lower() != ".pst":
        sys.exit("only .pst files can be attached")
    if not path.exists() and not args.create:
        sys.exit(f"{path} does not exist; pass --create to make a new PST")
    ns = outlook_namespace()
    ns.AddStoreEx(str(path), OL_STORE_UNICODE)
    print(f"attached {find_store(ns, str(path)).DisplayName}: {path}")


def cmd_detach(args: argparse.Namespace) -> None:
    ns = outlook_namespace()
    store = find_store(ns, str(args.pst.resolve()))
    if store.ExchangeStoreType != OL_NOT_EXCHANGE or not store.FilePath.lower().endswith(".pst"):
        sys.exit(f"refusing to detach a non-PST store: {store.DisplayName}")
    ns.RemoveStore(store.GetRootFolder())
    print(f"detached {store.FilePath} (Outlook may keep the file locked until it exits)")


def cmd_outlook_list(args: argparse.Namespace) -> None:
    print_records((rec for rec, _ in select_com(outlook_namespace(), args)), args.format)


def run_mutation(args: argparse.Namespace, describe: Callable[[object], str], act: Callable[[object], None]) -> None:
    selection = select_com(outlook_namespace(), args)
    verb = "applied" if args.apply else "would"
    for rec, item in selection:
        plan = describe(item)
        if not plan:
            continue
        print(f"{verb}: {plan} | {rec.subject[:60]} [{rec.folder}]")
        if args.apply:
            act(item)
    if not args.apply:
        print(f"dry run over {len(selection)} item(s); re-run with --apply to change them", file=sys.stderr)


def cmd_move(args: argparse.Namespace) -> None:
    store = find_store(outlook_namespace(), args.to_store or args.store)
    if args.apply or not args.create:
        target = find_folder(store, args.to_folder, create=args.create)
        label = target.FolderPath
    else:
        target, label = None, f"{store.DisplayName}/{args.to_folder} (created if missing)"
    run_mutation(args, lambda item: f"move to {label}", lambda item: item.Move(target))


def cmd_edit(args: argparse.Namespace) -> None:
    if not (args.set_subject or args.replace or args.mark or args.add_category):
        sys.exit("nothing to edit: pass --set-subject, --replace, --mark, or --add-category")

    def body_field(item) -> str | None:
        body_format = getattr(item, "BodyFormat", None)
        if body_format == OL_FORMAT_RICH_TEXT:
            return None  # rewriting Body would silently drop the RTF formatting
        return "HTMLBody" if body_format == OL_FORMAT_HTML else "Body"

    def describe(item) -> str:
        changes = []
        if args.set_subject:
            changes.append(f"subject -> {args.set_subject!r}")
        for old, new in args.replace or []:
            hits = (item.Subject or "").count(old)
            field_name = body_field(item)
            if field_name is None and old in (item.Body or ""):
                warn(f"RTF body not edited: {item.Subject!r}")
            elif field_name:
                hits += (getattr(item, field_name) or "").count(old)
            if hits:
                changes.append(f"replace {old!r}->{new!r} x{hits}")
        if args.mark:
            changes.append(f"mark {args.mark}")
        if args.add_category:
            changes.append(f"category +{args.add_category}")
        return ", ".join(changes)

    def act(item) -> None:
        if args.set_subject:
            item.Subject = args.set_subject
        for old, new in args.replace or []:
            if old in (item.Subject or ""):
                item.Subject = item.Subject.replace(old, new)
            field_name = body_field(item)
            if field_name and old in (getattr(item, field_name) or ""):
                setattr(item, field_name, getattr(item, field_name).replace(old, new))
        if args.mark:
            item.UnRead = args.mark == "unread"
        if args.add_category:
            current = [c.strip() for c in (item.Categories or "").split(",") if c.strip()]
            if args.add_category not in current:
                item.Categories = ", ".join([*current, args.add_category])
        item.Save()

    run_mutation(args, describe, act)


def cmd_delete(args: argparse.Namespace) -> None:
    def describe(item) -> str:
        deleted = item.Parent.Store.GetDefaultFolder(DEFAULT_FOLDERS["@deleted"])
        current = item.Parent.FolderPath.casefold()
        deleted_path = deleted.FolderPath.casefold()
        if current == deleted_path or current.startswith(deleted_path + "\\"):
            warn(f"skipped (already in Deleted Items, deleting would be permanent): {item.Subject!r}")
            return ""
        return "move to Deleted Items"

    run_mutation(args, describe, lambda item: item.Delete())


def cmd_outlook_export(args: argparse.Namespace) -> None:
    out = args.out.resolve()
    out.mkdir(parents=True, exist_ok=True)
    selection = select_com(outlook_namespace(), args)
    for rec, item in selection:
        stamp = rec.date.strftime("%Y-%m-%d_%H%M") if rec.date else "undated"
        path = unique_path(out / f"{stamp} {safe_name(rec.subject or 'no subject', 80)}.msg")
        item.SaveAs(str(path), OL_MSG_UNICODE)
        print(path)
    print(f"saved {len(selection)} message(s) to {out}", file=sys.stderr)


def outlook_exe() -> str:
    import winreg

    for hive in (winreg.HKEY_LOCAL_MACHINE, winreg.HKEY_CURRENT_USER):
        try:
            with winreg.OpenKey(hive, r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\OUTLOOK.EXE") as key:
                return winreg.QueryValue(key, None)
        except OSError:
            continue
    sys.exit("OUTLOOK.EXE is not registered under App Paths; is Classic Outlook installed?")


def cmd_open(args: argparse.Namespace) -> None:
    """Launch Outlook with its documented command-line switches."""
    command = [outlook_exe()]
    if args.profile:
        command += ["/profile", args.profile]
    if args.safe:
        command.append("/safe")
    if args.msg:
        command += ["/f", str(args.msg.resolve())]
    if args.select:
        command += ["/recycle", "/select", args.select]
    subprocess.Popen(command)
    print(" ".join(command))


# --------------------------------------------------------------------------
# CLI
# --------------------------------------------------------------------------


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = parser.add_subparsers(dest="command", required=True)

    def file_parser(name: str, help_text: str, func) -> argparse.ArgumentParser:
        p = sub.add_parser(name, help=help_text)
        p.add_argument("file", type=Path)
        p.add_argument("--via", choices=("auto", "pff", "outlook"), default="auto",
                       help="auto = libpff, or the running Outlook when it holds the file open; "
                            "outlook also attaches a locked, detached PST for the read")
        p.set_defaults(func=func)
        return p

    file_parser("tree", "read-only: folder tree with message counts", cmd_tree)

    for name, func, help_text in (
        ("list", cmd_list, "read-only: list matching messages"),
        ("export", cmd_export, "read-only: export matching messages with a hash manifest"),
    ):
        p = file_parser(name, help_text, func)
        if name == "export":
            p.add_argument("out", type=Path, help="new or empty output directory")
            p.add_argument("--format", choices=("dir", "eml"), default="dir",
                           help="dir = pffexport-style folder per message; eml = one .eml per message")
        else:
            p.add_argument("--format", choices=("table", "json", "csv"), default="table")
        p.add_argument("--folder", help="folder path contains (case-insensitive)")
        add_filter_args(p)

    p = file_parser("show", "read-only: one message as JSON", cmd_show)
    p.add_argument("id", help="message id from `list`")
    p.add_argument("--html", action="store_true", help="print the HTML body instead of plain text")

    live = sub.add_parser("outlook", help="live: Classic Outlook via COM (Windows)").add_subparsers(dest="action", required=True)
    live.add_parser("stores", help="list mailboxes and attached data files").set_defaults(func=cmd_stores)

    p = live.add_parser("attach", help="attach a PST to the current profile")
    p.add_argument("pst", type=Path)
    p.add_argument("--create", action="store_true", help="create the PST if it does not exist")
    p.set_defaults(func=cmd_attach)

    p = live.add_parser("detach", help="remove a PST from the profile (the file is kept)")
    p.add_argument("pst", type=Path)
    p.set_defaults(func=cmd_detach)

    def selection_parser(name: str, help_text: str, func, mutates: bool) -> argparse.ArgumentParser:
        p = live.add_parser(name, help=help_text)
        p.add_argument("--store", required=True, help="store display name, PST/OST path, or file name (see `stores`)")
        p.add_argument("--folder", default="", help="folder path from the store root; may start with @inbox, @sent, @drafts, @deleted, @junk, @outbox")
        p.add_argument("--recursive", action="store_true", help="include subfolders")
        p.add_argument("--id", action="append", help="EntryID from `outlook list` (repeatable; overrides folder filters)")
        add_filter_args(p)
        if mutates:
            p.add_argument("--apply", action="store_true", help="perform the change (default: dry run)")
        p.set_defaults(func=func)
        return p

    selection_parser("list", "list matching items", cmd_outlook_list, False).add_argument(
        "--format", choices=("table", "json", "csv"), default="table")

    p = selection_parser("move", "move matching items to another folder", cmd_move, True)
    p.add_argument("--to-folder", required=True)
    p.add_argument("--to-store", help="destination store (default: --store)")
    p.add_argument("--create", action="store_true", help="create missing destination folders")

    p = selection_parser("edit", "edit subject/body, read state, or categories", cmd_edit, True)
    p.add_argument("--set-subject")
    p.add_argument("--replace", nargs=2, action="append", metavar=("OLD", "NEW"),
                   help="replace text in subject and body (repeatable; HTML bodies are edited as raw HTML)")
    p.add_argument("--mark", choices=("read", "unread"))
    p.add_argument("--add-category")

    selection_parser("delete", "move matching items to Deleted Items (never permanent)", cmd_delete, True)

    p = selection_parser("export", "save matching items as .msg files", cmd_outlook_export, False)
    p.add_argument("out", type=Path)

    p = live.add_parser("open", help="launch Outlook with command-line switches")
    p.add_argument("--select", help="folder to show, e.g. outlook:inbox or outlook:calendar (/select)")
    p.add_argument("--msg", type=Path, help="open a .msg file (/f)")
    p.add_argument("--profile", help="load this mail profile (/profile)")
    p.add_argument("--safe", action="store_true", help="start without add-ins or customizations (/safe)")
    p.set_defaults(func=cmd_open)
    return parser


def main(argv: list[str] | None = None) -> int:
    global warnings_seen
    warnings_seen = 0
    for stream in (sys.stdout, sys.stderr):
        if hasattr(stream, "reconfigure"):  # Windows consoles default to cp1252 and choke on subjects
            stream.reconfigure(encoding="utf-8")
    parser = build_parser()
    args = parser.parse_args(argv)
    since, until = getattr(args, "since", None), getattr(args, "until", None)
    if since and until and since > until:
        parser.error("--since must not be after --until")
    if any(old == "" for old, _ in getattr(args, "replace", None) or []):
        parser.error("--replace OLD must not be empty")
    args.func(args)
    if warnings_seen:
        print(f"{warnings_seen} warning(s); output may be incomplete", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
