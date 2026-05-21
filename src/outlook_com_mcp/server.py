#!/usr/bin/env python3
"""
MCP Server — Microsoft Outlook desktop via COM Automation.

Exposes ~19 tools across mail (read/write), calendar, rules, and contacts (GAL).
Drives the user's running Outlook Classic client through pywin32; inherits the
already-authenticated session, requires no tenant permissions, no app registration.

Prerequisite: Outlook (M365 or Office 2019+) running on a Windows host.
Launched over stdio by the MCP client (Claude Code, Cursor, Gemini CLI, etc.).
"""

import json
import logging
import os
import sys
from datetime import datetime, timezone
from logging.handlers import RotatingFileHandler
from pathlib import Path

import pythoncom
import pywintypes
import win32com.client
from mcp.server.fastmcp import FastMCP

# =============================================================================
# Logging
# =============================================================================

LOG_DIR = Path(os.environ.get("LOCALAPPDATA", os.path.expanduser("~"))) / "outlook-com-mcp"
LOG_DIR.mkdir(parents=True, exist_ok=True)
LOG_PATH = LOG_DIR / "server.log"

logger = logging.getLogger("outlook-com")
logger.setLevel(logging.INFO)
handler = RotatingFileHandler(LOG_PATH, maxBytes=5 * 1024 * 1024, backupCount=5, encoding="utf-8")
handler.setFormatter(logging.Formatter("%(asctime)s %(levelname)s %(message)s"))
logger.addHandler(handler)
logger.info("=== outlook-com MCP server start (pid=%d) ===", os.getpid())

# =============================================================================
# MAPI / OOM constants
# =============================================================================

OL_FOLDER_INBOX = 6
OL_FOLDER_SENT = 5
OL_FOLDER_DRAFTS = 16
OL_FOLDER_CALENDAR = 9
OL_FOLDER_CONTACTS = 10

OL_MAIL_ITEM = 0  # CreateItem
OL_APPOINTMENT_ITEM = 1
OL_MEETING_ITEM_CLASS = 53  # AppointmentItem
OL_MEETING_REQUEST_CLASS = 55  # MeetingItem (request in inbox)

OL_FORMAT_PLAIN = 1
OL_FORMAT_HTML = 2

OL_RESPONSE_NONE = 0
OL_RESPONSE_ACCEPTED = 3
OL_RESPONSE_TENTATIVE = 2
OL_RESPONSE_DECLINED = 4

OL_RECIP_REQUIRED = 1
OL_RECIP_OPTIONAL = 2

MAX_BODY_BYTES = 50_000  # body truncated to this many bytes by default
DEFAULT_LIMIT = 20
MAX_LIMIT = 1000

# =============================================================================
# Write guardrails
# =============================================================================
# - create_draft always creates a visible draft in Outlook (never auto-sends)
# - send_draft is a separate explicit two-step call
# - reply_mail accepts send=True ONLY if OUTLOOK_MCP_ALLOW_SEND=1
# - recipient domain allowlist (default: empty, no restriction).
#   In a corporate context, setting OUTLOOK_MCP_ALLOWED_DOMAINS="yourcorp.com" is strongly recommended.

ALLOW_SEND = os.environ.get("OUTLOOK_MCP_ALLOW_SEND", "0") == "1"
_allowed_raw = os.environ.get("OUTLOOK_MCP_ALLOWED_DOMAINS", "")
ALLOWED_DOMAINS = {d.strip().lower() for d in _allowed_raw.split(",") if d.strip()}
logger.info("guardrails: allow_send=%s allowed_domains=%s", ALLOW_SEND, sorted(ALLOWED_DOMAINS))

# =============================================================================
# Outlook connection
# =============================================================================


def _connect():
    """Return (Application, MAPI Namespace). Prefer GetActiveObject, fall back to Dispatch."""
    try:
        app = win32com.client.GetActiveObject("Outlook.Application")
        logger.info("connected via GetActiveObject")
    except pythoncom.com_error:
        logger.info("Outlook not running, falling back to Dispatch (may launch Outlook)")
        app = win32com.client.Dispatch("Outlook.Application")
    ns = app.GetNamespace("MAPI")
    return app, ns


def _ns():
    """Wrapper that retries once if Outlook was closed between two calls."""
    try:
        _, ns = _connect()
        # Force a MAPI access to validate the session
        _ = ns.DefaultStore.DisplayName
        return ns
    except Exception as e:
        logger.exception("connect failed")
        raise RuntimeError(f"OUTLOOK_NOT_RUNNING: {e}") from e


def _app_ns():
    """Variant that also returns the Application (for CreateItem)."""
    try:
        app, ns = _connect()
        _ = ns.DefaultStore.DisplayName
        return app, ns
    except Exception as e:
        logger.exception("connect failed")
        raise RuntimeError(f"OUTLOOK_NOT_RUNNING: {e}") from e


# =============================================================================
# Conversion helpers
# =============================================================================


def _to_iso(value) -> str | None:
    """Outlook returns tz-aware datetimes (pywintypes.TimeType). Convert to ISO UTC."""
    if value is None:
        return None
    try:
        dt = datetime.fromisoformat(str(value)) if not hasattr(value, "tzinfo") else value
        if dt.tzinfo is None:
            dt = dt.replace(tzinfo=timezone.utc)
        return dt.astimezone(timezone.utc).isoformat()
    except Exception:
        return str(value)


def _safe(getter, default=None):
    """Read a COM property swallowing com_error (some properties are absent depending on item type)."""
    try:
        return getter()
    except (pywintypes.com_error, AttributeError):
        return default


# MAPI DASL property tags used to resolve X.500 -> SMTP
_PR_SMTP_ADDRESS = "http://schemas.microsoft.com/mapi/proptag/0x39FE001E"
_PR_SENDER_SMTP_ADDRESS = "http://schemas.microsoft.com/mapi/proptag/0x5D01001E"
_PR_SENT_REPRESENTING_SMTP_ADDRESS = "http://schemas.microsoft.com/mapi/proptag/0x5D02001E"


def _resolve_smtp(item, prop_tag: str) -> str | None:
    """Resolve an SMTP address via PropertyAccessor (X.500 -> SMTP)."""
    try:
        return item.PropertyAccessor.GetProperty(prop_tag) or None
    except (pywintypes.com_error, AttributeError):
        return None


def _sender_smtp(item) -> str | None:
    """Try several locations to retrieve the sender's SMTP address from a MailItem."""
    raw = _safe(lambda: item.SenderEmailAddress)
    # If SenderEmailAddress is already in SMTP format, keep it
    if raw and "@" in raw and not raw.startswith("/"):
        return raw
    # Otherwise resolve via PropertyAccessor
    smtp = _resolve_smtp(item, _PR_SENDER_SMTP_ADDRESS)
    if smtp:
        return smtp
    return _resolve_smtp(item, _PR_SENT_REPRESENTING_SMTP_ADDRESS)


def _recipient_smtp(recipient) -> str | None:
    """SMTP of a Recipient (from a MailItem or AppointmentItem)."""
    addr = _safe(lambda: recipient.Address)
    if addr and "@" in addr and not addr.startswith("/"):
        return addr
    return _resolve_smtp(recipient, _PR_SMTP_ADDRESS)


def _summary(item) -> dict:
    """Summary of a MailItem -> serializable dict. Includes sender_smtp if resolved."""
    return {
        "entry_id": _safe(lambda: item.EntryID),
        "received_at": _to_iso(_safe(lambda: item.ReceivedTime)),
        "sent_at": _to_iso(_safe(lambda: item.SentOn)),
        "sender": _safe(lambda: item.SenderName),
        "sender_address": _safe(lambda: item.SenderEmailAddress),
        "sender_smtp": _sender_smtp(item),
        "to": _safe(lambda: item.To),
        "cc": _safe(lambda: item.CC),
        "subject": _safe(lambda: item.Subject) or "",
        "is_read": not bool(_safe(lambda: item.UnRead, False)),
        "has_attachments": bool(_safe(lambda: item.Attachments.Count, 0)),
        "size_bytes": _safe(lambda: item.Size),
        "importance": _safe(lambda: item.Importance),  # 0=Low 1=Normal 2=High
        "folder": _safe(lambda: item.Parent.FolderPath),
    }


def _folder_tree(folder, depth=0, max_depth=4) -> dict:
    """Recursive folder walk, truncated at max_depth."""
    node = {
        "name": _safe(lambda: folder.Name),
        "path": _safe(lambda: folder.FolderPath),
        "item_count": _safe(lambda: folder.Items.Count, 0),
        "unread_count": _safe(lambda: folder.UnReadItemCount, 0),
        "children": [],
    }
    if depth >= max_depth:
        return node
    try:
        for sub in folder.Folders:
            node["children"].append(_folder_tree(sub, depth + 1, max_depth))
    except pywintypes.com_error:
        pass
    return node


def _resolve_folder(ns, path_or_name: str | None):
    """Resolve a folder by full path (\\Store\\Inbox\\Sub) or by standard name (Inbox/Sent/Drafts)."""
    if not path_or_name or path_or_name.lower() == "inbox":
        return ns.GetDefaultFolder(OL_FOLDER_INBOX)
    if path_or_name.lower() == "sent":
        return ns.GetDefaultFolder(OL_FOLDER_SENT)
    if path_or_name.lower() == "drafts":
        return ns.GetDefaultFolder(OL_FOLDER_DRAFTS)
    # Full path "\\store\\path\\sub"
    if path_or_name.startswith("\\\\"):
        parts = [p for p in path_or_name.split("\\") if p]
        store_name = parts[0]
        store = None
        for s in ns.Stores:
            if s.DisplayName == store_name:
                store = s
                break
        if not store:
            raise ValueError(f"FOLDER_NOT_FOUND: store '{store_name}'")
        cur = store.GetRootFolder()
        for p in parts[1:]:
            cur = cur.Folders[p]  # raises com_error if absent
        return cur
    # Otherwise look up by name inside the default store
    root = ns.DefaultStore.GetRootFolder()
    for sub in root.Folders:
        if sub.Name.lower() == path_or_name.lower():
            return sub
    raise ValueError(f"FOLDER_NOT_FOUND: {path_or_name}")


def _get_by_entry_id(ns, entry_id: str):
    try:
        return ns.GetItemFromID(entry_id)
    except pywintypes.com_error as e:
        raise ValueError(f"ENTRY_ID_INVALID: {e}") from e


# =============================================================================
# Restrict / DASL helpers
# =============================================================================
# DASL properties cf. https://learn.microsoft.com/fr-fr/office/vba/api/outlook.olbody
DASL_SUBJECT = "urn:schemas:httpmail:subject"
DASL_FROM_NAME = "urn:schemas:httpmail:fromname"
DASL_FROM_EMAIL = "urn:schemas:httpmail:fromemail"
DASL_RECEIVED = "urn:schemas:httpmail:datereceived"
DASL_TEXT_DESC = "urn:schemas:httpmail:textdescription"
DASL_UNREAD = "urn:schemas:httpmail:read"  # 0 if unread


def _escape_dasl(s: str) -> str:
    return s.replace("'", "''")


# =============================================================================
# MCP server
# =============================================================================

mcp = FastMCP("outlook-com")


@mcp.tool()
def list_folders(max_depth: int = 3) -> str:
    """Tree of Outlook folders, up to max_depth levels.

    Args:
        max_depth: maximum exploration depth (default 3, max 6).
    """
    max_depth = max(1, min(max_depth, 6))
    ns = _ns()
    out = []
    for store in ns.Stores:
        try:
            root = store.GetRootFolder()
            tree = _folder_tree(root, 0, max_depth)
            tree["store"] = store.DisplayName
            out.append(tree)
        except pywintypes.com_error:
            continue
    logger.info("list_folders depth=%d stores=%d", max_depth, len(out))
    return json.dumps(out, ensure_ascii=False, indent=2)


@mcp.tool()
def list_mail(
    folder: str = "inbox",
    limit: int = DEFAULT_LIMIT,
    unread_only: bool = False,
    since_iso: str | None = None,
) -> str:
    """List recent mails from a folder (summary only, no body).

    Args:
        folder: 'inbox' (default), 'sent', 'drafts', a sub-folder name, or a full path '\\\\Store\\\\Folder'.
        limit: max number of mails (default 20, max 1000).
        unread_only: if True, only return unread items.
        since_iso: ISO 8601 date (e.g. '2026-04-20T00:00:00'); only return mails received after.

    Returns a JSON with:
        items, count, folder, query (echo of the params),
        coverage: {oldest_received, newest_received} effective bounds of returned items,
        truncated: True if limit was hit AND more items exist (paginate via since_iso).
    """
    limit = max(1, min(limit, MAX_LIMIT))
    ns = _ns()
    folder_obj = _resolve_folder(ns, folder)
    items = folder_obj.Items
    items.Sort("[ReceivedTime]", True)  # server-side descending sort

    filters = []
    if unread_only:
        filters.append(f"\"{DASL_UNREAD}\" = 0")
    if since_iso:
        try:
            dt = datetime.fromisoformat(since_iso.replace("Z", "+00:00"))
            # Inside @SQL=, you must use DASL URIs + short ISO 8601 format (not the US %m/%d/%Y format).
            # Mixing [ReceivedTime] (Jet syntax) with @SQL= silently breaks the Outlook parser.
            iso = dt.strftime("%Y-%m-%dT%H:%M:%S")
            filters.append(f"\"{DASL_RECEIVED}\" >= '{iso}'")
        except ValueError:
            return json.dumps({"error": f"invalid since_iso: {since_iso}"}, ensure_ascii=False)

    if filters:
        try:
            items = items.Restrict("@SQL=" + " AND ".join(filters))
        except pywintypes.com_error as e:
            return json.dumps({"error": f"Restrict failed: {e}"}, ensure_ascii=False)

    out = []
    item = items.GetFirst()
    count = 0
    while item is not None and count < limit:
        try:
            # Filter out non-MailItem entries (reports, meeting items, …) — keep only MailItem (43)
            if _safe(lambda: item.Class, 0) == 43:  # olMail
                out.append(_summary(item))
                count += 1
        except pywintypes.com_error:
            pass
        item = items.GetNext()

    # Truncation detection: if we hit limit, check whether at least one more MailItem exists
    truncated = False
    if count >= limit and item is not None:
        probe = item
        while probe is not None:
            if _safe(lambda: probe.Class, 0) == 43:
                truncated = True
                break
            probe = items.GetNext()

    # Effective coverage (bounds of returned items)
    coverage = {"oldest_received": None, "newest_received": None}
    if out:
        dates = [m.get("received_at") for m in out if m.get("received_at")]
        if dates:
            coverage = {"oldest_received": min(dates), "newest_received": max(dates)}

    logger.info("list_mail folder=%s limit=%d unread=%s since=%s -> %d truncated=%s", folder, limit, unread_only, since_iso, len(out), truncated)
    return json.dumps({
        "folder": folder_obj.FolderPath,
        "count": len(out),
        "truncated": truncated,
        "coverage": coverage,
        "query": {"folder": folder, "limit": limit, "unread_only": unread_only, "since_iso": since_iso},
        "items": out,
    }, ensure_ascii=False, indent=2)


@mcp.tool()
def read_mail(entry_id: str, max_body_bytes: int = MAX_BODY_BYTES, include_html: bool = False) -> str:
    """Read the full content of a mail by its EntryID.

    Args:
        entry_id: MAPI identifier (returned by list_mail/search_mail).
        max_body_bytes: truncate body to N bytes (default 50000). 0 = no limit.
        include_html: include HTMLBody (can be heavy).
    """
    ns = _ns()
    item = _get_by_entry_id(ns, entry_id)
    body = _safe(lambda: item.Body) or ""
    truncated = False
    if max_body_bytes and len(body.encode("utf-8")) > max_body_bytes:
        body = body.encode("utf-8")[:max_body_bytes].decode("utf-8", errors="ignore")
        truncated = True

    out = _summary(item)
    out["body"] = body
    out["body_truncated"] = truncated
    if include_html:
        html = _safe(lambda: item.HTMLBody) or ""
        if max_body_bytes and len(html.encode("utf-8")) > max_body_bytes:
            html = html.encode("utf-8")[:max_body_bytes].decode("utf-8", errors="ignore")
            out["html_truncated"] = True
        out["html"] = html

    # Recipients with SMTP resolution
    recipients = []
    try:
        for r in item.Recipients:
            recipients.append(
                {
                    "name": _safe(lambda: r.Name),
                    "address": _safe(lambda: r.Address),
                    "smtp": _recipient_smtp(r),
                    "type": _safe(lambda: r.Type),  # MailItem: 1=To 2=Cc 3=Bcc
                }
            )
    except pywintypes.com_error:
        pass
    out["recipients"] = recipients

    # Attachments (metadata only)
    atts = []
    try:
        for i in range(1, item.Attachments.Count + 1):
            a = item.Attachments.Item(i)
            atts.append({"index": i, "filename": _safe(lambda: a.FileName), "size": _safe(lambda: a.Size)})
    except pywintypes.com_error:
        pass
    out["attachments"] = atts

    logger.info("read_mail entry_id=%s body_bytes=%d trunc=%s atts=%d", entry_id[:16], len(body), truncated, len(atts))
    return json.dumps(out, ensure_ascii=False, indent=2)


@mcp.tool()
def search_mail(
    query: str,
    folder: str = "inbox",
    scope: str = "subject_body",
    limit: int = 50,
) -> str:
    """Search mails via Items.Restrict (fast, server-side).

    Args:
        query: search term.
        folder: target folder (default 'inbox', see list_mail).
        scope: 'subject_body' (default), 'subject', 'sender', 'all'.
        limit: max number of results (default 50, max 1000).

    Also returns truncated/coverage (see list_mail).
    """
    limit = max(1, min(limit, MAX_LIMIT))
    q = _escape_dasl(query)
    ns = _ns()
    folder_obj = _resolve_folder(ns, folder)
    items = folder_obj.Items
    items.Sort("[ReceivedTime]", True)

    if scope == "subject":
        clause = f"\"{DASL_SUBJECT}\" LIKE '%{q}%'"
    elif scope == "sender":
        clause = f"(\"{DASL_FROM_NAME}\" LIKE '%{q}%' OR \"{DASL_FROM_EMAIL}\" LIKE '%{q}%')"
    elif scope == "all":
        clause = (
            f"\"{DASL_SUBJECT}\" LIKE '%{q}%' OR "
            f"\"{DASL_TEXT_DESC}\" LIKE '%{q}%' OR "
            f"\"{DASL_FROM_NAME}\" LIKE '%{q}%' OR "
            f"\"{DASL_FROM_EMAIL}\" LIKE '%{q}%'"
        )
    else:  # subject_body (default)
        clause = f"\"{DASL_SUBJECT}\" LIKE '%{q}%' OR \"{DASL_TEXT_DESC}\" LIKE '%{q}%'"

    try:
        restricted = items.Restrict("@SQL=" + clause)
    except pywintypes.com_error as e:
        return json.dumps({"error": f"Restrict failed: {e}", "clause": clause}, ensure_ascii=False)

    out = []
    item = restricted.GetFirst()
    while item is not None and len(out) < limit:
        if _safe(lambda: item.Class, 0) == 43:
            out.append(_summary(item))
        item = restricted.GetNext()

    # Truncation detection
    truncated = False
    if len(out) >= limit and item is not None:
        probe = item
        while probe is not None:
            if _safe(lambda: probe.Class, 0) == 43:
                truncated = True
                break
            probe = restricted.GetNext()

    coverage = {"oldest_received": None, "newest_received": None}
    if out:
        dates = [m.get("received_at") for m in out if m.get("received_at")]
        if dates:
            coverage = {"oldest_received": min(dates), "newest_received": max(dates)}

    logger.info("search_mail q=%r scope=%s folder=%s -> %d truncated=%s", query, scope, folder, len(out), truncated)
    return json.dumps({
        "query": query,
        "scope": scope,
        "folder": folder_obj.FolderPath,
        "count": len(out),
        "truncated": truncated,
        "coverage": coverage,
        "items": out,
    }, ensure_ascii=False, indent=2)


@mcp.tool()
def download_attachment(entry_id: str, index: int, dest_dir: str | None = None) -> str:
    """Download an attachment to local disk.

    Args:
        entry_id: EntryID of the mail (see read_mail).
        index: 1-based attachment index (see read_mail.attachments[].index).
        dest_dir: target directory. Default: %LOCALAPPDATA%\\outlook-com-mcp\\downloads\\.
    """
    ns = _ns()
    item = _get_by_entry_id(ns, entry_id)
    if index < 1 or index > item.Attachments.Count:
        return json.dumps({"error": f"index {index} out of range (1..{item.Attachments.Count})"}, ensure_ascii=False)
    att = item.Attachments.Item(index)
    base = Path(dest_dir) if dest_dir else (LOG_DIR / "downloads")
    base.mkdir(parents=True, exist_ok=True)
    fname = att.FileName or f"attachment_{index}.bin"
    target = base / fname
    # Avoid filename collision
    n = 1
    while target.exists():
        target = base / f"{target.stem}_{n}{target.suffix}"
        n += 1
    att.SaveAsFile(str(target))
    logger.info("download_attachment entry_id=%s idx=%d -> %s", entry_id[:16], index, target)
    return json.dumps({"saved_to": str(target), "size": _safe(lambda: att.Size)}, ensure_ascii=False)


@mcp.tool()
def health_check() -> str:
    """Server diagnostic: Outlook running, COM latency, active guardrails.

    Call this first if something is broken. Returns enough state for support to triage.
    """
    import platform
    import time as _time

    out: dict = {
        "server_pid": os.getpid(),
        "python": sys.version.split()[0],
        "platform": platform.platform(),
        "log_path": str(LOG_PATH),
        "guardrails": {
            "allow_send_direct": ALLOW_SEND,
            "allowed_domains": sorted(ALLOWED_DOMAINS),
        },
    }

    t0 = _time.perf_counter()
    try:
        app, ns = _app_ns()
        out["outlook_running"] = True
        out["connect_ms"] = round((_time.perf_counter() - t0) * 1000, 1)
        out["outlook_version"] = _safe(lambda: app.Version)
        out["default_store"] = _safe(lambda: ns.DefaultStore.DisplayName)
        out["accounts"] = [
            {"display": _safe(lambda: a.DisplayName), "smtp": _safe(lambda: a.SmtpAddress)}
            for a in ns.Accounts
        ]
        inbox = ns.GetDefaultFolder(OL_FOLDER_INBOX)
        out["inbox"] = {
            "count": _safe(lambda: inbox.Items.Count),
            "unread": _safe(lambda: inbox.UnReadItemCount),
        }
        cal = ns.GetDefaultFolder(OL_FOLDER_CALENDAR)
        out["calendar"] = {"count": _safe(lambda: cal.Items.Count)}
        out["status"] = "OK"
    except Exception as e:
        out["outlook_running"] = False
        out["status"] = "ERROR"
        out["error"] = str(e)
        out["hint"] = "Check that Outlook is running. If it is, see the server log."

    return json.dumps(out, ensure_ascii=False, indent=2)


@mcp.tool()
def whoami() -> str:
    """Return the current Outlook identity (useful as a COM connection smoke test)."""
    ns = _ns()
    accounts = []
    try:
        for a in ns.Accounts:
            accounts.append({"display_name": a.DisplayName, "smtp": _safe(lambda: a.SmtpAddress)})
    except pywintypes.com_error:
        pass
    inbox = ns.GetDefaultFolder(OL_FOLDER_INBOX)
    return json.dumps(
        {
            "default_store": ns.DefaultStore.DisplayName,
            "accounts": accounts,
            "inbox_count": inbox.Items.Count,
            "inbox_unread": inbox.UnReadItemCount,
        },
        ensure_ascii=False,
        indent=2,
    )


# =============================================================================
# Write helpers
# =============================================================================


def _split_recipients(s: str | None) -> list[str]:
    if not s:
        return []
    parts = []
    for chunk in s.replace("\n", ",").replace(";", ",").split(","):
        c = chunk.strip()
        if c:
            parts.append(c)
    return parts


def _check_recipients(recipients: list[str]) -> list[str]:
    """Return the list of recipients outside the allowlist. Empty = all OK."""
    rejected = []
    for r in recipients:
        # Extract the email from either 'Name <a@b.c>' or 'a@b.c'
        addr = r
        if "<" in r and ">" in r:
            addr = r.split("<", 1)[1].split(">", 1)[0]
        if "@" not in addr:
            rejected.append(r)
            continue
        domain = addr.rsplit("@", 1)[1].lower().strip()
        if domain not in ALLOWED_DOMAINS:
            rejected.append(r)
    return rejected


def _attach_files(mail, attachments: list[str] | None):
    if not attachments:
        return []
    attached = []
    for path in attachments:
        p = Path(path).expanduser().resolve()
        if not p.is_file():
            raise ValueError(f"ATTACHMENT_NOT_FOUND: {p}")
        mail.Attachments.Add(str(p))
        attached.append(str(p))
    return attached


# =============================================================================
# Write tools
# =============================================================================


@mcp.tool()
def create_draft(
    to: str,
    subject: str,
    body: str,
    cc: str | None = None,
    bcc: str | None = None,
    html: bool = False,
    attachments: list[str] | None = None,
) -> str:
    """Create a draft visible in Outlook's Drafts folder. DOES NOT SEND.

    The draft appears in Outlook's Drafts folder; the user can review it then click
    Send manually, or call send_draft.

    Args:
        to: recipient(s), separated by ',' or ';'.
        subject: mail subject.
        body: body (text or HTML depending on `html`).
        cc, bcc: optional.
        html: True to interpret body as HTML.
        attachments: list of absolute paths to local files.
    """
    app, ns = _app_ns()
    mail = app.CreateItem(OL_MAIL_ITEM)

    to_list = _split_recipients(to)
    cc_list = _split_recipients(cc)
    bcc_list = _split_recipients(bcc)
    rejected = _check_recipients(to_list + cc_list + bcc_list)
    if rejected:
        return json.dumps(
            {
                "error": "RECIPIENTS_NOT_ALLOWED",
                "rejected": rejected,
                "allowed_domains": sorted(ALLOWED_DOMAINS),
                "hint": "Adjust OUTLOOK_MCP_ALLOWED_DOMAINS in your .mcp.json if needed.",
            },
            ensure_ascii=False,
        )

    mail.To = "; ".join(to_list)
    if cc_list:
        mail.CC = "; ".join(cc_list)
    if bcc_list:
        mail.BCC = "; ".join(bcc_list)
    mail.Subject = subject
    if html:
        mail.BodyFormat = OL_FORMAT_HTML
        mail.HTMLBody = body
    else:
        mail.BodyFormat = OL_FORMAT_PLAIN
        mail.Body = body

    try:
        attached = _attach_files(mail, attachments)
    except ValueError as e:
        return json.dumps({"error": str(e)}, ensure_ascii=False)

    mail.Save()  # persists into Drafts; does NOT send
    eid = mail.EntryID
    logger.info("create_draft to=%s subject=%r entry_id=%s atts=%d", to_list, subject, eid[:16], len(attached))
    return json.dumps(
        {
            "ok": True,
            "entry_id": eid,
            "to": to_list,
            "cc": cc_list,
            "bcc": bcc_list,
            "subject": subject,
            "attachments": attached,
            "next_step": "Review in Outlook > Drafts, then call send_draft(entry_id, confirm=True) to send.",
        },
        ensure_ascii=False,
        indent=2,
    )


@mcp.tool()
def send_draft(entry_id: str, confirm: bool = False) -> str:
    """Send an existing draft. Requires confirm=True AND passes the allowlist re-check.

    Args:
        entry_id: EntryID of the draft (returned by create_draft).
        confirm: must be True to actually send.
    """
    if not confirm:
        return json.dumps(
            {"error": "CONFIRM_REQUIRED", "hint": "Call send_draft(entry_id, confirm=True)."},
            ensure_ascii=False,
        )
    ns = _ns()
    item = _get_by_entry_id(ns, entry_id)
    # Re-validate recipients at send time (the user may have edited them in Outlook).
    # Outlook rewrites internal addresses to X.500 on Save -> resolve back to SMTP via PropertyAccessor first.
    all_recip = []
    try:
        for r in item.Recipients:
            smtp = _recipient_smtp(r)
            all_recip.append(smtp or _safe(lambda: r.Address) or _safe(lambda: r.Name))
    except pywintypes.com_error:
        pass
    rejected = _check_recipients([r for r in all_recip if r])
    if rejected and not ALLOW_SEND:
        return json.dumps(
            {
                "error": "RECIPIENTS_NOT_ALLOWED_AT_SEND",
                "rejected": rejected,
                "allowed_domains": sorted(ALLOWED_DOMAINS),
                "hint": "Set OUTLOOK_MCP_ALLOW_SEND=1 to bypass, or adjust ALLOWED_DOMAINS.",
            },
            ensure_ascii=False,
        )
    item.Send()
    logger.warning("send_draft SENT entry_id=%s recipients=%s", entry_id[:16], all_recip)
    return json.dumps({"ok": True, "sent_entry_id": entry_id, "recipients": all_recip}, ensure_ascii=False)


@mcp.tool()
def reply_mail(
    entry_id: str,
    body: str,
    reply_all: bool = False,
    html: bool = False,
    save_only: bool = True,
) -> str:
    """Create a reply to a mail. By default saves as draft (save_only=True).

    To send directly, pass save_only=False; this requires OUTLOOK_MCP_ALLOW_SEND=1.

    Args:
        entry_id: EntryID of the original mail.
        body: body of the reply.
        reply_all: True for Reply All.
        html: True if body is HTML.
        save_only: True (default) saves as draft. False = direct send (under guardrails).
    """
    ns = _ns()
    original = _get_by_entry_id(ns, entry_id)
    reply = original.ReplyAll() if reply_all else original.Reply()
    if html:
        reply.BodyFormat = OL_FORMAT_HTML
        # Prepend body before the quoted reply
        reply.HTMLBody = body + "<br><br>" + (reply.HTMLBody or "")
    else:
        reply.BodyFormat = OL_FORMAT_PLAIN
        reply.Body = body + "\n\n" + (reply.Body or "")

    # Check recipients (X.500 -> SMTP resolution required to match the allowlist)
    recip = []
    try:
        for r in reply.Recipients:
            smtp = _recipient_smtp(r)
            recip.append(smtp or _safe(lambda: r.Address) or _safe(lambda: r.Name))
    except pywintypes.com_error:
        pass
    rejected = _check_recipients([r for r in recip if r])

    if save_only:
        reply.Save()
        logger.info("reply_mail saved draft entry_id=%s reply_all=%s rejected=%s", entry_id[:16], reply_all, rejected)
        return json.dumps(
            {
                "ok": True,
                "draft_entry_id": reply.EntryID,
                "recipients": recip,
                "recipients_outside_allowlist": rejected,
                "next_step": "Call send_draft(draft_entry_id, confirm=True) to send.",
            },
            ensure_ascii=False,
            indent=2,
        )

    if rejected and not ALLOW_SEND:
        reply.Save()
        return json.dumps(
            {
                "error": "RECIPIENTS_NOT_ALLOWED",
                "rejected": rejected,
                "saved_draft_entry_id": reply.EntryID,
                "hint": "Reply saved as a draft. Edit recipients in Outlook, then call send_draft.",
            },
            ensure_ascii=False,
        )
    if not ALLOW_SEND:
        reply.Save()
        return json.dumps(
            {
                "error": "ALLOW_SEND_DISABLED",
                "saved_draft_entry_id": reply.EntryID,
                "hint": "Set OUTLOOK_MCP_ALLOW_SEND=1 in your .mcp.json to allow direct send.",
            },
            ensure_ascii=False,
        )
    reply.Send()
    logger.warning("reply_mail SENT entry_id=%s reply_all=%s recipients=%s", entry_id[:16], reply_all, recip)
    return json.dumps({"ok": True, "sent": True, "recipients": recip}, ensure_ascii=False)


@mcp.tool()
def move_mail(entry_id: str, target_folder: str) -> str:
    """Move a mail to a folder (folder resolution as in list_mail).

    Args:
        entry_id: EntryID of the mail.
        target_folder: short name (Inbox/Sent/Drafts), sub-folder name, or full path '\\\\Store\\\\Folder'.
    """
    ns = _ns()
    item = _get_by_entry_id(ns, entry_id)
    target = _resolve_folder(ns, target_folder)
    moved = item.Move(target)
    logger.info("move_mail entry_id=%s -> %s (new_eid=%s)", entry_id[:16], target.FolderPath, _safe(lambda: moved.EntryID, "")[:16])
    return json.dumps(
        {
            "ok": True,
            "new_entry_id": _safe(lambda: moved.EntryID),
            "target_folder": target.FolderPath,
        },
        ensure_ascii=False,
    )


@mcp.tool()
def mark_read(entry_id: str, read: bool = True) -> str:
    """Mark a mail as read (read=True) or unread (read=False)."""
    ns = _ns()
    item = _get_by_entry_id(ns, entry_id)
    item.UnRead = not read
    item.Save()
    logger.info("mark_read entry_id=%s read=%s", entry_id[:16], read)
    return json.dumps({"ok": True, "entry_id": entry_id, "is_read": read}, ensure_ascii=False)


@mcp.tool()
def flag_mail(entry_id: str, flag: bool = True) -> str:
    """Set or clear the follow-up flag on a mail."""
    ns = _ns()
    item = _get_by_entry_id(ns, entry_id)
    # FlagStatus: 0=NoFlag, 1=Complete, 2=Marked
    item.FlagStatus = 2 if flag else 0
    item.Save()
    logger.info("flag_mail entry_id=%s flag=%s", entry_id[:16], flag)
    return json.dumps({"ok": True, "entry_id": entry_id, "flagged": flag}, ensure_ascii=False)


@mcp.tool()
def guardrails_status() -> str:
    """Return the current state of write guardrails (ALLOW_SEND, ALLOWED_DOMAINS)."""
    return json.dumps(
        {
            "allow_send_direct": ALLOW_SEND,
            "allowed_domains": sorted(ALLOWED_DOMAINS),
            "policy": "create_draft never sends. send_draft requires confirm=True. "
                     "Direct send via reply_mail(save_only=False) requires OUTLOOK_MCP_ALLOW_SEND=1 "
                     "AND all recipients in ALLOWED_DOMAINS.",
        },
        ensure_ascii=False,
        indent=2,
    )


# =============================================================================
# Calendar
# =============================================================================


def _event_summary(item) -> dict:
    return {
        "entry_id": _safe(lambda: item.EntryID),
        "subject": _safe(lambda: item.Subject) or "",
        "start": _to_iso(_safe(lambda: item.Start)),
        "end": _to_iso(_safe(lambda: item.End)),
        "duration_minutes": _safe(lambda: item.Duration),
        "location": _safe(lambda: item.Location) or "",
        "organizer": _safe(lambda: item.Organizer),
        "is_recurring": bool(_safe(lambda: item.IsRecurring, False)),
        "is_all_day": bool(_safe(lambda: item.AllDayEvent, False)),
        "meeting_status": _safe(lambda: item.MeetingStatus),  # 0=non meeting, 1=meeting, 3=cancelled
        "response_status": _safe(lambda: item.ResponseStatus),  # 0..4
        "required": _safe(lambda: item.RequiredAttendees) or "",
        "optional": _safe(lambda: item.OptionalAttendees) or "",
        "online_meeting_url": _safe(lambda: item.GetConversationTopic and None),  # placeholder
    }


def _to_outlook_dt_str(dt: datetime) -> str:
    """Outlook Restrict expects dates in US locale format."""
    return dt.strftime("%m/%d/%Y %I:%M %p")


def _parse_iso(s: str) -> datetime:
    return datetime.fromisoformat(s.replace("Z", "+00:00"))


@mcp.tool()
def list_events(days_ahead: int = 7, days_back: int = 0, limit: int = 50, calendar: str | None = None) -> str:
    """List calendar events in a window [now-days_back, now+days_ahead].

    Includes recurrences. Sorted ascending by start date.

    Args:
        days_ahead: number of days in the future (default 7).
        days_back: number of days in the past (default 0).
        limit: max number of events (default 50, max 1000).
        calendar: specific calendar name (default: primary calendar).
    """
    limit = max(1, min(limit, MAX_LIMIT))
    ns = _ns()
    cal = ns.GetDefaultFolder(OL_FOLDER_CALENDAR) if not calendar else _resolve_folder(ns, calendar)

    now = datetime.now()
    start = now - __import__("datetime").timedelta(days=days_back)
    end = now + __import__("datetime").timedelta(days=days_ahead)

    items = cal.Items
    items.IncludeRecurrences = True
    items.Sort("[Start]")
    clause = f"[Start] >= '{_to_outlook_dt_str(start)}' AND [Start] <= '{_to_outlook_dt_str(end)}'"
    try:
        restricted = items.Restrict(clause)
    except pywintypes.com_error as e:
        return json.dumps({"error": f"Restrict failed: {e}", "clause": clause}, ensure_ascii=False)

    out = []
    item = restricted.GetFirst()
    while item is not None and len(out) < limit:
        out.append(_event_summary(item))
        item = restricted.GetNext()

    logger.info("list_events ahead=%d back=%d -> %d", days_ahead, days_back, len(out))
    return json.dumps(
        {
            "calendar": cal.FolderPath,
            "window": {"start": start.isoformat(timespec="seconds"), "end": end.isoformat(timespec="seconds")},
            "count": len(out),
            "events": out,
        },
        ensure_ascii=False,
        indent=2,
    )


@mcp.tool()
def read_event(entry_id: str, max_body_bytes: int = MAX_BODY_BYTES) -> str:
    """Read the full detail of an event (body + attendees)."""
    ns = _ns()
    item = _get_by_entry_id(ns, entry_id)
    out = _event_summary(item)

    body = _safe(lambda: item.Body) or ""
    truncated = False
    if max_body_bytes and len(body.encode("utf-8")) > max_body_bytes:
        body = body.encode("utf-8")[:max_body_bytes].decode("utf-8", errors="ignore")
        truncated = True
    out["body"] = body
    out["body_truncated"] = truncated

    # Recipients list (with SMTP resolution)
    recipients = []
    try:
        for r in item.Recipients:
            recipients.append(
                {
                    "name": _safe(lambda: r.Name),
                    "address": _safe(lambda: r.Address),
                    "smtp": _recipient_smtp(r),
                    "type": _safe(lambda: r.Type),  # 1=required 2=optional 3=resource
                    "response": _safe(lambda: r.MeetingResponseStatus),
                }
            )
    except pywintypes.com_error:
        pass
    out["recipients"] = recipients
    return json.dumps(out, ensure_ascii=False, indent=2)


@mcp.tool()
def find_freeslots(
    duration_minutes: int = 30,
    days_ahead: int = 5,
    business_hours: str = "09:00-18:00",
    weekdays_only: bool = True,
    slot_step_minutes: int = 30,
    limit: int = 10,
) -> str:
    """Find free slots in the primary calendar.

    Args:
        duration_minutes: desired slot duration (default 30).
        days_ahead: search window in days (default 5).
        business_hours: 'HH:MM-HH:MM' (default '09:00-18:00').
        weekdays_only: True to skip Saturdays/Sundays.
        slot_step_minutes: scan step (default 30).
        limit: max number of slots to return (default 10).
    """
    from datetime import timedelta as _td

    limit = max(1, min(limit, 50))
    duration_minutes = max(5, min(duration_minutes, 480))
    ns = _ns()
    cal = ns.GetDefaultFolder(OL_FOLDER_CALENDAR)

    try:
        bh_start_str, bh_end_str = business_hours.split("-")
        bh_start_h, bh_start_m = map(int, bh_start_str.split(":"))
        bh_end_h, bh_end_m = map(int, bh_end_str.split(":"))
    except Exception:
        return json.dumps({"error": f"invalid business_hours: {business_hours}"}, ensure_ascii=False)

    now = datetime.now()
    end = now + _td(days=days_ahead)

    # NOTE: Restrict + IncludeRecurrences + Sort yields inconsistent results
    # (see https://learn.microsoft.com/office/vba/api/outlook.items.includerecurrences).
    # Iterate directly, sorted, with an early break when Start > end.
    items = cal.Items
    items.Sort("[Start]")
    items.IncludeRecurrences = True

    busy = []
    item = items.GetFirst()
    scanned = 0
    while item is not None:
        scanned += 1
        s = _safe(lambda: item.Start)
        e = _safe(lambda: item.End)
        if s is None or e is None:
            item = items.GetNext()
            continue
        try:
            s_dt = datetime(s.year, s.month, s.day, s.hour, s.minute, s.second)
            e_dt = datetime(e.year, e.month, e.day, e.hour, e.minute, e.second)
        except Exception:
            item = items.GetNext()
            continue
        if s_dt > end:
            break  # ascending sort, we're done
        if e_dt <= now:
            item = items.GetNext()
            continue
        status = _safe(lambda: item.BusyStatus, 2)  # 0=Free 1=Tent 2=Busy 3=OOF 4=Wherever
        if status == 0:  # Free, ignore
            item = items.GetNext()
            continue
        if _safe(lambda: item.AllDayEvent, False):
            item = items.GetNext()
            continue
        busy.append((s_dt, e_dt))
        item = items.GetNext()

    busy.sort()
    logger.debug("find_freeslots scanned=%d busy=%d", scanned, len(busy))

    def overlaps(slot_s: datetime, slot_e: datetime) -> bool:
        for bs, be in busy:
            if slot_s < be and slot_e > bs:
                return True
        return False

    slots = []
    cur = now.replace(second=0, microsecond=0)
    cur += _td(minutes=(slot_step_minutes - cur.minute % slot_step_minutes) % slot_step_minutes)

    while cur < end and len(slots) < limit:
        if weekdays_only and cur.weekday() >= 5:
            cur = (cur + _td(days=1)).replace(hour=bh_start_h, minute=bh_start_m)
            continue
        bh_start = cur.replace(hour=bh_start_h, minute=bh_start_m)
        bh_end = cur.replace(hour=bh_end_h, minute=bh_end_m)
        if cur < bh_start:
            cur = bh_start
            continue
        slot_end = cur + _td(minutes=duration_minutes)
        if slot_end > bh_end:
            cur = (cur + _td(days=1)).replace(hour=bh_start_h, minute=bh_start_m)
            continue
        if not overlaps(cur, slot_end):
            slots.append(
                {
                    "start": cur.isoformat(timespec="minutes"),
                    "end": slot_end.isoformat(timespec="minutes"),
                    "weekday": cur.strftime("%A"),
                }
            )
        cur += _td(minutes=slot_step_minutes)

    logger.info("find_freeslots dur=%d days=%d -> %d slots", duration_minutes, days_ahead, len(slots))
    return json.dumps(
        {
            "duration_minutes": duration_minutes,
            "business_hours": business_hours,
            "weekdays_only": weekdays_only,
            "busy_count": len(busy),
            "slots_found": len(slots),
            "slots": slots,
        },
        ensure_ascii=False,
        indent=2,
    )


# =============================================================================
# Multi-person Free/Busy (Recipient.FreeBusy + Exchange)
# =============================================================================
# Recipient.FreeBusy(Start, MinPerChar, CompleteFormat=True) returns a string
# of characters '0'..'4' representing slots of MinPerChar minutes starting at
# Start (local midnight). 0=Free, 1=Tentative, 2=Busy, 3=OOF, 4=WorkingElsewhere.
# Available with default Exchange free/busy permissions.


def _freebusy_to_intervals(busy_string: str, start: datetime, slot_minutes: int) -> list:
    """Decode a FreeBusy string -> list of merged busy intervals (start, end)."""
    from datetime import timedelta as _td
    intervals: list[tuple[datetime, datetime]] = []
    for i, ch in enumerate(busy_string):
        if ch != "0":  # 0=Free, anything else = unavailable
            s = start + _td(minutes=i * slot_minutes)
            e = s + _td(minutes=slot_minutes)
            if intervals and intervals[-1][1] == s:
                intervals[-1] = (intervals[-1][0], e)
            else:
                intervals.append((s, e))
    return intervals


def _self_busy_from_calendar(ns, start: datetime, days_ahead: int) -> list:
    """Iterate the local calendar to extract busy intervals.
    More reliable than FreeBusy for the current user: published Exchange free/busy
    can be incomplete (missing events, 30-min granularity, stale data).
    """
    from datetime import timedelta as _td
    cal = ns.GetDefaultFolder(OL_FOLDER_CALENDAR)
    items = cal.Items
    items.Sort("[Start]")
    items.IncludeRecurrences = True

    busy: list[tuple[datetime, datetime]] = []
    end_window = start + _td(days=days_ahead + 1)
    item = items.GetFirst()
    while item is not None:
        s = _safe(lambda: item.Start)
        e = _safe(lambda: item.End)
        if s is None or e is None:
            item = items.GetNext()
            continue
        try:
            s_dt = datetime(s.year, s.month, s.day, s.hour, s.minute, s.second)
            e_dt = datetime(e.year, e.month, e.day, e.hour, e.minute, e.second)
        except Exception:
            item = items.GetNext()
            continue
        if s_dt > end_window:
            break
        if e_dt < start:
            item = items.GetNext()
            continue
        # Skip cancelled / free events / all-day
        if _safe(lambda: item.MeetingStatus, 0) == 5:  # olMeetingCanceled
            item = items.GetNext()
            continue
        if _safe(lambda: item.BusyStatus, 2) == 0:  # Free
            item = items.GetNext()
            continue
        if _safe(lambda: item.AllDayEvent, False):
            item = items.GetNext()
            continue
        busy.append((s_dt, e_dt))
        item = items.GetNext()
    return busy


def _get_attendee_busy(ns, smtp: str, start: datetime, slot_minutes: int = 30) -> tuple[list, str | None]:
    """Return (busy intervals, error). error=None if OK."""
    try:
        recipient = ns.CreateRecipient(smtp)
        if not recipient.Resolve():
            return [], "unresolved (not found in the GAL)"
        result = recipient.FreeBusy(start, slot_minutes, True)
        if not result:
            return [], "FreeBusy returned empty (Exchange has no data or access denied)"
        return _freebusy_to_intervals(result, start, slot_minutes), None
    except pywintypes.com_error as e:
        return [], f"COM error: {e}"


@mcp.tool()
def find_freeslots_multi(
    attendees: str,
    duration_minutes: int = 60,
    days_ahead: int = 5,
    business_hours: str = "09:00-18:00",
    weekdays_only: bool = True,
    slot_step_minutes: int = 30,
    limit: int = 10,
    include_self: bool = True,
) -> str:
    """Find common free slots between you and several attendees via Exchange free/busy.

    Uses `Recipient.FreeBusy`, which works with default Exchange permissions
    (no need to be a delegate on each attendee's calendar).

    Args:
        attendees: attendees' emails, separated by ',' or ';'.
        duration_minutes: desired slot duration (default 60, max 480).
        days_ahead: window in days (default 5).
        business_hours: 'HH:MM-HH:MM' (default '09:00-18:00').
        weekdays_only: skip Saturdays/Sundays.
        slot_step_minutes: scan step (default 30).
        limit: max number of slots (default 10, max 50).
        include_self: include the current user's own calendar in the check (default True).
    """
    from datetime import timedelta as _td

    limit = max(1, min(limit, 50))
    duration_minutes = max(5, min(duration_minutes, 480))

    attendee_list = _split_recipients(attendees)
    if not attendee_list:
        return json.dumps({"error": "attendees is empty"}, ensure_ascii=False)

    try:
        bh_start_str, bh_end_str = business_hours.split("-")
        bh_start_h, bh_start_m = map(int, bh_start_str.split(":"))
        bh_end_h, bh_end_m = map(int, bh_end_str.split(":"))
    except Exception:
        return json.dumps({"error": f"invalid business_hours: {business_hours}"}, ensure_ascii=False)

    ns = _ns()

    # Exchange granularity: 30 min per char (15 min on some modern tenants)
    fb_slot = 30
    now = datetime.now()
    fb_start = now.replace(hour=0, minute=0, second=0, microsecond=0)

    # Collect each attendee's busy intervals
    all_busy: list = []
    attendee_status: dict = {}

    if include_self:
        # For the current user: iterate the local calendar (published FreeBusy is often incomplete)
        my_smtp = None
        try:
            my_smtp = ns.Accounts.Item(1).SmtpAddress
        except Exception:
            my_smtp = "self"
        try:
            self_busy = _self_busy_from_calendar(ns, fb_start, days_ahead)
            attendee_status[my_smtp + " (self)"] = {
                "busy_intervals": len(self_busy),
                "source": "local_calendar",
                "error": None,
            }
            all_busy.extend(self_busy)
        except Exception as e:
            attendee_status[my_smtp + " (self)"] = {"busy_intervals": 0, "error": f"local_calendar failed: {e}"}

    for smtp in attendee_list:
        intervals, err = _get_attendee_busy(ns, smtp, fb_start, fb_slot)
        attendee_status[smtp] = {"busy_intervals": len(intervals), "source": "exchange_freebusy", "error": err}
        if not err:
            all_busy.extend(intervals)

    # Merge intervals
    all_busy.sort()
    merged: list = []
    for s, e in all_busy:
        if merged and s <= merged[-1][1]:
            merged[-1] = (merged[-1][0], max(merged[-1][1], e))
        else:
            merged.append((s, e))

    def overlaps(slot_s, slot_e):
        for bs, be in merged:
            if slot_s < be and slot_e > bs:
                return True
        return False

    end_window = now + _td(days=days_ahead)
    slots = []
    cur = now.replace(second=0, microsecond=0)
    cur += _td(minutes=(slot_step_minutes - cur.minute % slot_step_minutes) % slot_step_minutes)

    while cur < end_window and len(slots) < limit:
        if weekdays_only and cur.weekday() >= 5:
            cur = (cur + _td(days=1)).replace(hour=bh_start_h, minute=bh_start_m)
            continue
        bh_start = cur.replace(hour=bh_start_h, minute=bh_start_m)
        bh_end = cur.replace(hour=bh_end_h, minute=bh_end_m)
        if cur < bh_start:
            cur = bh_start
            continue
        slot_end = cur + _td(minutes=duration_minutes)
        if slot_end > bh_end:
            cur = (cur + _td(days=1)).replace(hour=bh_start_h, minute=bh_start_m)
            continue
        if not overlaps(cur, slot_end):
            slots.append(
                {
                    "start": cur.isoformat(timespec="minutes"),
                    "end": slot_end.isoformat(timespec="minutes"),
                    "weekday": cur.strftime("%A"),
                }
            )
        cur += _td(minutes=slot_step_minutes)

    logger.info(
        "find_freeslots_multi dur=%d days=%d attendees=%d -> %d slots",
        duration_minutes, days_ahead, len(attendee_list), len(slots),
    )
    return json.dumps(
        {
            "duration_minutes": duration_minutes,
            "business_hours": business_hours,
            "weekdays_only": weekdays_only,
            "attendees_checked": list(attendee_status.keys()),
            "attendee_status": attendee_status,
            "merged_busy_intervals": len(merged),
            "slots_found": len(slots),
            "slots": slots,
        },
        ensure_ascii=False,
        indent=2,
    )


@mcp.tool()
def create_event_draft(
    subject: str,
    start_iso: str,
    duration_minutes: int = 30,
    body: str | None = None,
    location: str | None = None,
    required_attendees: str | None = None,
    optional_attendees: str | None = None,
) -> str:
    """Create an event/meeting as a DRAFT (saved without sending invites).

    To send the invites, call send_event_invites(entry_id, confirm=True).
    Recipients outside ALLOWED_DOMAINS are rejected at creation time.

    Args:
        subject: event title.
        start_iso: ISO date/time (e.g. '2026-04-30T14:00:00').
        duration_minutes: duration (default 30, max 480).
        body: optional body.
        location: location / Teams link.
        required_attendees: emails separated by ',' or ';'.
        optional_attendees: emails separated by ',' or ';'.
    """
    from datetime import timedelta as _td

    duration_minutes = max(5, min(duration_minutes, 480))
    try:
        start = _parse_iso(start_iso)
    except ValueError:
        return json.dumps({"error": f"invalid start_iso: {start_iso}"}, ensure_ascii=False)

    req_list = _split_recipients(required_attendees)
    opt_list = _split_recipients(optional_attendees)
    rejected = _check_recipients(req_list + opt_list)
    if rejected:
        return json.dumps(
            {
                "error": "RECIPIENTS_NOT_ALLOWED",
                "rejected": rejected,
                "allowed_domains": sorted(ALLOWED_DOMAINS),
            },
            ensure_ascii=False,
        )

    app, ns = _app_ns()
    appt = app.CreateItem(OL_APPOINTMENT_ITEM)
    appt.Subject = subject
    appt.Start = start
    appt.Duration = duration_minutes
    if location:
        appt.Location = location
    if body:
        appt.Body = body

    if req_list or opt_list:
        appt.MeetingStatus = 1  # olMeeting
        for addr in req_list:
            r = appt.Recipients.Add(addr)
            r.Type = OL_RECIP_REQUIRED
        for addr in opt_list:
            r = appt.Recipients.Add(addr)
            r.Type = OL_RECIP_OPTIONAL
        appt.Recipients.ResolveAll()

    appt.Save()
    eid = appt.EntryID
    logger.info(
        "create_event_draft subject=%r start=%s dur=%d req=%s opt=%s eid=%s",
        subject, start_iso, duration_minutes, req_list, opt_list, eid[:16],
    )
    return json.dumps(
        {
            "ok": True,
            "entry_id": eid,
            "subject": subject,
            "start": start.isoformat(timespec="minutes"),
            "duration_minutes": duration_minutes,
            "required": req_list,
            "optional": opt_list,
            "is_meeting": bool(req_list or opt_list),
            "next_step": (
                "Invites NOT sent. Review in Outlook > Calendar, then call "
                "send_event_invites(entry_id, confirm=True) to send."
                if (req_list or opt_list)
                else "Event saved in the calendar (no attendees)."
            ),
        },
        ensure_ascii=False,
        indent=2,
    )


@mcp.tool()
def send_event_invites(entry_id: str, confirm: bool = False) -> str:
    """Send invites for an existing meeting event. Requires confirm=True."""
    if not confirm:
        return json.dumps({"error": "CONFIRM_REQUIRED"}, ensure_ascii=False)
    ns = _ns()
    item = _get_by_entry_id(ns, entry_id)
    if _safe(lambda: item.MeetingStatus, 0) == 0:
        return json.dumps({"error": "NOT_A_MEETING", "hint": "create_event_draft was called without attendees: nothing to invite."}, ensure_ascii=False)
    item.Send()
    logger.warning("send_event_invites SENT entry_id=%s", entry_id[:16])
    return json.dumps({"ok": True, "sent_entry_id": entry_id}, ensure_ascii=False)


@mcp.tool()
def respond_meeting(entry_id: str, response: str = "accepted", send_response: bool = False, comment: str | None = None) -> str:
    """Respond to a meeting invite (from inbox or calendar).

    Args:
        entry_id: EntryID of the MeetingItem (inbox) or AppointmentItem (calendar).
        response: 'accepted' (default), 'tentative', 'declined'.
        send_response: True to notify the organizer, False to respond silently.
        comment: optional message to include in the response.
    """
    mapping = {
        "accepted": OL_RESPONSE_ACCEPTED,
        "tentative": OL_RESPONSE_TENTATIVE,
        "declined": OL_RESPONSE_DECLINED,
    }
    if response not in mapping:
        return json.dumps({"error": f"invalid response: {response}", "valid": list(mapping)}, ensure_ascii=False)
    ns = _ns()
    item = _get_by_entry_id(ns, entry_id)
    target_status = mapping[response]
    # Both MeetingItem (inbox) and AppointmentItem expose a Respond method
    try:
        resp_item = item.Respond(target_status, True)  # True = no UI prompt
        if comment and hasattr(resp_item, "Body"):
            resp_item.Body = comment + "\n\n" + (resp_item.Body or "")
        if send_response:
            resp_item.Send()
            action = "sent"
        else:
            resp_item.Save()
            action = "saved"
    except pywintypes.com_error as e:
        return json.dumps({"error": f"Respond failed: {e}"}, ensure_ascii=False)
    logger.info("respond_meeting entry_id=%s response=%s action=%s", entry_id[:16], response, action)
    return json.dumps({"ok": True, "response": response, "action": action}, ensure_ascii=False)


# =============================================================================
# Outlook rules
# =============================================================================
# API: Namespace.DefaultStore.GetRules() -> Rules collection. Each Rule has
# Conditions / Actions / Exceptions. After modification you must call
# rules.Save() to persist. Rules can be server-side (executed by Exchange even
# when Outlook is closed) or local (Outlook open).

OL_RULE_RECEIVE = 0
OL_RULE_EXECUTE_ALL_MESSAGES = 0  # ApplyRuleNow scope
OL_IMPORTANCE_HIGH = 2

OL_RECIPIENT_TYPE_TO = 1
OL_RECIPIENT_TYPE_CC = 2


def _rule_summary(rule) -> dict:
    """Summary of a Rule -> serializable dict. Decodes Conditions and Actions."""
    conds: list[str] = []
    acts: list[str] = []

    try:
        c = rule.Conditions
        if c.From.Enabled:
            addrs = []
            try:
                for r in c.From.Recipients:
                    addrs.append(_recipient_smtp(r) or _safe(lambda: r.Name))
            except pywintypes.com_error:
                pass
            conds.append(f"From IN [{', '.join(a for a in addrs if a)}]")
        if c.SentTo.Enabled:
            addrs = []
            try:
                for r in c.SentTo.Recipients:
                    addrs.append(_recipient_smtp(r) or _safe(lambda: r.Name))
            except pywintypes.com_error:
                pass
            conds.append(f"SentTo IN [{', '.join(a for a in addrs if a)}]")
        if c.Subject.Enabled:
            conds.append(f"Subject CONTAINS {list(c.Subject.Text)}")
        if c.BodyOrSubject.Enabled:
            conds.append(f"BodyOrSubject CONTAINS {list(c.BodyOrSubject.Text)}")
        if c.HasAttachment.Enabled:
            conds.append("HasAttachment")
        if c.Importance.Enabled:
            conds.append(f"Importance = {c.Importance.Importance}")
    except (pywintypes.com_error, AttributeError):
        pass

    try:
        a = rule.Actions
        if a.MoveToFolder.Enabled:
            try:
                acts.append(f"Move -> {a.MoveToFolder.Folder.FolderPath}")
            except (pywintypes.com_error, AttributeError):
                acts.append("Move -> <unresolved>")
        if a.CopyToFolder.Enabled:
            try:
                acts.append(f"Copy -> {a.CopyToFolder.Folder.FolderPath}")
            except (pywintypes.com_error, AttributeError):
                acts.append("Copy -> <unresolved>")
        if a.MarkAsTask.Enabled:
            acts.append("MarkAsTask")
        if a.Delete.Enabled:
            acts.append("Delete")
        if a.DeletePermanently.Enabled:
            acts.append("DeletePermanently")
        if a.AssignToCategory.Enabled:
            acts.append(f"AssignToCategory {list(a.AssignToCategory.Categories)}")
        if a.Forward.Enabled:
            addrs = []
            try:
                for r in a.Forward.Recipients:
                    addrs.append(_recipient_smtp(r) or _safe(lambda: r.Name))
            except pywintypes.com_error:
                pass
            acts.append(f"Forward -> [{', '.join(a for a in addrs if a)}]")
        if hasattr(a, "Stop") and a.Stop.Enabled:
            acts.append("StopProcessingMoreRules")
    except (pywintypes.com_error, AttributeError):
        pass

    return {
        "name": _safe(lambda: rule.Name),
        "enabled": bool(_safe(lambda: rule.Enabled, False)),
        "execution_order": _safe(lambda: rule.ExecutionOrder),
        "is_local_rule": bool(_safe(lambda: rule.IsLocalRule, False)),
        "rule_type": _safe(lambda: rule.RuleType),  # 0=Receive 1=Send
        "conditions": conds,
        "actions": acts,
    }


@mcp.tool()
def list_rules() -> str:
    """List all Outlook rules from the default store with decoded conditions/actions."""
    ns = _ns()
    rules = ns.DefaultStore.GetRules()
    out = []
    for i in range(1, rules.Count + 1):
        try:
            out.append(_rule_summary(rules.Item(i)))
        except pywintypes.com_error as e:
            out.append({"name": f"<rule {i}>", "error": str(e)})
    logger.info("list_rules count=%d", len(out))
    return json.dumps({"count": len(out), "rules": out}, ensure_ascii=False, indent=2)


@mcp.tool()
def toggle_rule(name: str, enabled: bool) -> str:
    """Enable or disable an existing rule (lookup by name).

    Args:
        name: exact rule name (see list_rules).
        enabled: True to enable, False to disable.
    """
    ns = _ns()
    rules = ns.DefaultStore.GetRules()
    target = None
    for i in range(1, rules.Count + 1):
        r = rules.Item(i)
        if _safe(lambda: r.Name) == name:
            target = r
            break
    if target is None:
        return json.dumps({"error": "RULE_NOT_FOUND", "name": name}, ensure_ascii=False)
    target.Enabled = enabled
    rules.Save(False)
    logger.info("toggle_rule name=%r enabled=%s", name, enabled)
    return json.dumps({"ok": True, "name": name, "enabled": enabled}, ensure_ascii=False)


@mcp.tool()
def create_rule(
    name: str,
    from_addresses: list[str] | None = None,
    subject_contains: list[str] | None = None,
    body_or_subject_contains: list[str] | None = None,
    sent_to: list[str] | None = None,
    has_attachment: bool = False,
    importance_high: bool = False,
    move_to_folder: str | None = None,
    copy_to_folder: str | None = None,
    assign_categories: list[str] | None = None,
    forward_to: list[str] | None = None,
    stop_processing: bool = True,
    enabled: bool = True,
    apply_now: bool = False,
) -> str:
    """Create a receive rule. At least ONE condition + ONE action are required.

    Args:
        name: rule name (unique recommended).
        from_addresses: list of emails (OR-chain). Outside ALLOWED_DOMAINS = rejected.
        subject_contains: list of strings (OR-chain) in the subject.
        body_or_subject_contains: same, in subject OR body.
        sent_to: list of emails (OR-chain) among recipients.
        has_attachment: matches mails with an attachment.
        importance_high: matches mails with High importance.
        move_to_folder: target folder (resolved as in list_mail).
        copy_to_folder: same, but copy instead of move.
        assign_categories: list of Outlook categories.
        forward_to: list of forward addresses (those outside allowlist are rejected).
        stop_processing: stop the next rules (default True).
        enabled: rule active (default True).
        apply_now: apply immediately to existing Inbox mails.
    """
    # Allowlist validation on outgoing addresses (forward) and incoming addresses (filter)
    rejected = _check_recipients(forward_to or []) + _check_recipients(from_addresses or []) + _check_recipients(sent_to or [])
    if rejected:
        return json.dumps(
            {"error": "ADDRESSES_NOT_ALLOWED", "rejected": rejected, "allowed_domains": sorted(ALLOWED_DOMAINS)},
            ensure_ascii=False,
        )

    has_condition = any([from_addresses, subject_contains, body_or_subject_contains, sent_to, has_attachment, importance_high])
    has_action = any([move_to_folder, copy_to_folder, assign_categories, forward_to])
    if not has_condition:
        return json.dumps({"error": "NO_CONDITION", "hint": "At least one condition is required."}, ensure_ascii=False)
    if not has_action:
        return json.dumps({"error": "NO_ACTION", "hint": "At least one action is required."}, ensure_ascii=False)

    ns = _ns()
    rules = ns.DefaultStore.GetRules()

    # Reject if a rule with the same name already exists
    for i in range(1, rules.Count + 1):
        if _safe(lambda: rules.Item(i).Name) == name:
            return json.dumps({"error": "RULE_NAME_EXISTS", "name": name}, ensure_ascii=False)

    rule = rules.Create(name, OL_RULE_RECEIVE)

    # Conditions
    if from_addresses:
        for addr in from_addresses:
            rule.Conditions.From.Recipients.Add(addr)
        rule.Conditions.From.Recipients.ResolveAll()
        rule.Conditions.From.Enabled = True
    if sent_to:
        for addr in sent_to:
            rule.Conditions.SentTo.Recipients.Add(addr)
        rule.Conditions.SentTo.Recipients.ResolveAll()
        rule.Conditions.SentTo.Enabled = True
    if subject_contains:
        rule.Conditions.Subject.Text = list(subject_contains)
        rule.Conditions.Subject.Enabled = True
    if body_or_subject_contains:
        rule.Conditions.BodyOrSubject.Text = list(body_or_subject_contains)
        rule.Conditions.BodyOrSubject.Enabled = True
    if has_attachment:
        rule.Conditions.HasAttachment.Enabled = True
    if importance_high:
        rule.Conditions.Importance.Importance = OL_IMPORTANCE_HIGH
        rule.Conditions.Importance.Enabled = True

    # Actions
    if move_to_folder:
        try:
            target = _resolve_folder(ns, move_to_folder)
        except (ValueError, pywintypes.com_error) as e:
            return json.dumps({"error": "MOVE_FOLDER_NOT_FOUND", "details": str(e)}, ensure_ascii=False)
        rule.Actions.MoveToFolder.Folder = target
        rule.Actions.MoveToFolder.Enabled = True
    if copy_to_folder:
        try:
            target = _resolve_folder(ns, copy_to_folder)
        except (ValueError, pywintypes.com_error) as e:
            return json.dumps({"error": "COPY_FOLDER_NOT_FOUND", "details": str(e)}, ensure_ascii=False)
        rule.Actions.CopyToFolder.Folder = target
        rule.Actions.CopyToFolder.Enabled = True
    if assign_categories:
        rule.Actions.AssignToCategory.Categories = list(assign_categories)
        rule.Actions.AssignToCategory.Enabled = True
    if forward_to:
        for addr in forward_to:
            rule.Actions.Forward.Recipients.Add(addr)
        rule.Actions.Forward.Recipients.ResolveAll()
        rule.Actions.Forward.Enabled = True
    if stop_processing:
        try:
            rule.Actions.Stop.Enabled = True
        except (pywintypes.com_error, AttributeError):
            pass  # Stop not supported in every Outlook version

    rule.Enabled = enabled

    try:
        rules.Save(False)
    except pywintypes.com_error as e:
        return json.dumps({"error": "RULES_SAVE_FAILED", "details": str(e), "hint": "Conditions/actions are probably incompatible with a server-side rule."}, ensure_ascii=False)

    applied = 0
    if apply_now:
        try:
            inbox = ns.GetDefaultFolder(OL_FOLDER_INBOX)
            rule.Execute(False, inbox, False, OL_RULE_EXECUTE_ALL_MESSAGES)
            applied = inbox.Items.Count  # approximation: all inbox mails were scanned
        except pywintypes.com_error as e:
            logger.warning("rule.Execute apply_now failed: %s", e)

    logger.info("create_rule name=%r enabled=%s apply_now=%s", name, enabled, apply_now)
    return json.dumps(
        {
            "ok": True,
            "name": name,
            "enabled": enabled,
            "is_local_rule": bool(_safe(lambda: rule.IsLocalRule, False)),
            "execution_order": _safe(lambda: rule.ExecutionOrder),
            "summary": _rule_summary(rule),
            "apply_now_executed": apply_now,
        },
        ensure_ascii=False,
        indent=2,
    )


@mcp.tool()
def delete_rule(name: str, confirm: bool = False) -> str:
    """Delete an existing rule. Requires confirm=True.

    Args:
        name: exact rule name.
        confirm: must be True to actually delete.
    """
    if not confirm:
        return json.dumps({"error": "CONFIRM_REQUIRED", "hint": "Call delete_rule(name, confirm=True)."}, ensure_ascii=False)
    ns = _ns()
    rules = ns.DefaultStore.GetRules()
    found = False
    for i in range(1, rules.Count + 1):
        if _safe(lambda: rules.Item(i).Name) == name:
            rules.Remove(i)
            found = True
            break
    if not found:
        return json.dumps({"error": "RULE_NOT_FOUND", "name": name}, ensure_ascii=False)
    rules.Save(False)
    logger.warning("delete_rule DELETED name=%r", name)
    return json.dumps({"ok": True, "name": name, "deleted": True}, ensure_ascii=False)


# =============================================================================
# Contacts
# =============================================================================


def _exchange_user_dict(eu) -> dict:
    """Extract the useful fields from an ExchangeUser COM object."""
    return {
        "name": _safe(lambda: eu.Name),
        "smtp_address": _safe(lambda: eu.PrimarySmtpAddress),
        "alias": _safe(lambda: eu.Alias),
        "job_title": _safe(lambda: eu.JobTitle),
        "department": _safe(lambda: eu.Department),
        "office_location": _safe(lambda: eu.OfficeLocation),
        "company_name": _safe(lambda: eu.CompanyName),
        "business_phone": _safe(lambda: eu.BusinessTelephoneNumber),
        "mobile_phone": _safe(lambda: eu.MobileTelephoneNumber),
    }


@mcp.tool()
def search_contacts(query: str, limit: int = 20) -> str:
    """Search the Exchange address book (GAL) by name or email.

    Args:
        query: search term (name, first name, partial email).
        limit: max number of results (default 20, max 1000).
    """
    limit = max(1, min(limit, MAX_LIMIT))
    ns = _ns()

    # Fast path: ADODB/LDAP with server-side ANR (Ambiguous Name Resolution)
    try:
        conn = win32com.client.Dispatch("ADODB.Connection")
        conn.Provider = "ADsDSOObject"
        conn.Open("Active Directory Provider")
        cmd = win32com.client.Dispatch("ADODB.Command")
        cmd.ActiveConnection = conn
        escaped = query.replace("\\", "\\5c").replace("(", "\\28").replace(")", "\\29").replace("*", "\\2a")
        cmd.CommandText = (
            f"<LDAP://>;(&(objectCategory=person)(objectClass=user)(anr={escaped}));"
            "cn,mail,title,department,physicalDeliveryOfficeName,telephoneNumber,mobile,company;"
            "subtree"
        )
        cmd.Properties("Page Size").Value = limit
        cmd.Properties("Size Limit").Value = limit
        rs, _ = cmd.Execute()
        out = []
        while not rs.EOF and len(out) < limit:
            def _field(name):
                try:
                    v = rs.Fields(name).Value
                    return v if v is not None else None
                except Exception:
                    return None
            out.append({
                "name": _field("cn"),
                "smtp_address": _field("mail"),
                "job_title": _field("title"),
                "department": _field("department"),
                "office_location": _field("physicalDeliveryOfficeName"),
                "business_phone": _field("telephoneNumber"),
                "mobile_phone": _field("mobile"),
                "company_name": _field("company"),
            })
            rs.MoveNext()
        rs.Close()
        conn.Close()
        logger.info("search_contacts (LDAP) q=%r -> %d", query, len(out))
        return json.dumps({"query": query, "count": len(out), "contacts": out}, ensure_ascii=False, indent=2)
    except Exception as ldap_err:
        logger.warning("search_contacts LDAP failed (%s), fallback CreateRecipient", ldap_err)

    # Fallback: CreateRecipient (Exchange ANR resolution, 1 result max)
    try:
        recipient = ns.CreateRecipient(query)
        resolved = recipient.Resolve()
        ae = _safe(lambda: recipient.AddressEntry)
        if ae is None:
            return json.dumps({"query": query, "count": 0, "contacts": []}, ensure_ascii=False, indent=2)
        eu = _safe(lambda: ae.GetExchangeUser())
        if eu is None:
            return json.dumps({"query": query, "count": 0, "contacts": [],
                               "hint": "Resolved, but not an Exchange user."}, ensure_ascii=False, indent=2)
        out = [_exchange_user_dict(eu)]
        logger.info("search_contacts (fallback) q=%r -> %d", query, len(out))
        return json.dumps({"query": query, "count": len(out), "contacts": out,
                           "hint": "CreateRecipient fallback (1 result max). LDAP unavailable."}, ensure_ascii=False, indent=2)
    except pywintypes.com_error as e:
        return json.dumps({"error": f"search failed: {e}"}, ensure_ascii=False)


@mcp.tool()
def get_contact_details(smtp_address: str) -> str:
    """Get the full details of an Exchange contact from its SMTP address.

    Args:
        smtp_address: contact email address (e.g. firstname.lastname@example.com).
    """
    ns = _ns()
    try:
        recipient = ns.CreateRecipient(smtp_address)
        if not recipient.Resolve():
            return json.dumps({"error": "CONTACT_NOT_FOUND", "smtp_address": smtp_address,
                               "hint": "Not found in the Exchange GAL."}, ensure_ascii=False)
    except pywintypes.com_error as e:
        return json.dumps({"error": f"Resolve failed: {e}"}, ensure_ascii=False)

    ae = _safe(lambda: recipient.AddressEntry)
    if ae is None:
        return json.dumps({"error": "ADDRESS_ENTRY_NULL", "smtp_address": smtp_address}, ensure_ascii=False)

    eu = _safe(lambda: ae.GetExchangeUser())
    if eu is None:
        return json.dumps({"error": "NOT_EXCHANGE_USER", "smtp_address": smtp_address,
                           "hint": "Contact exists but is not an Exchange user (external contact?)."}, ensure_ascii=False)

    details = _exchange_user_dict(eu)
    # Additional fields available on individual lookup
    details["street_address"] = _safe(lambda: eu.StreetAddress)
    details["city"] = _safe(lambda: eu.City)
    details["state"] = _safe(lambda: eu.StateOrProvince)
    details["postal_code"] = _safe(lambda: eu.PostalCode)

    # Manager
    mgr = _safe(lambda: eu.GetExchangeUserManager())
    if mgr is not None:
        details["manager"] = {
            "name": _safe(lambda: mgr.Name),
            "smtp_address": _safe(lambda: mgr.PrimarySmtpAddress),
        }
    else:
        details["manager"] = None

    logger.info("get_contact_details smtp=%s -> %s", smtp_address, details.get("name"))
    return json.dumps(details, ensure_ascii=False, indent=2)


# =============================================================================
# Main
# =============================================================================

def main() -> None:
    """Entry point for `python -m outlook_com_mcp` and the console script."""
    try:
        mcp.run(transport="stdio")
    except Exception:
        logger.exception("server crashed")
        sys.exit(1)


if __name__ == "__main__":
    main()
