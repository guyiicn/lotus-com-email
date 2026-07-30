"""
Lightweight MCP server for Lotus Notes mail.

Speaks MCP over stdio (JSON-RPC 2.0). Hand-written instead of depending on
the `mcp` Python SDK because that SDK has a Rust extension that won't build
under 32-bit Python (required for Lotus Notes COM). The protocol surface we
need is small: initialize, tools/list, tools/call.

Wire format: one JSON object per line on stdin, one JSON object per line
on stdout. Notifications have no id and get no response.
"""

import json
import os
import sys
import traceback

# ensure src/ is importable when launched as a script by an MCP host
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from notes_client import get_client

SERVER_NAME = "lotus-notes-mail"
SERVER_VERSION = "0.1.0"
PROTOCOL_VERSION = "2024-11-05"

# ── tool schemas ──────────────────────────────────────────────────────

TOOLS = [
    {
        "name": "notes_list_mail",
        "description": "List recent emails from a folder. Returns sender, subject, date, and a universal_id you pass to notes_get_mail to read the body. Defaults to the 20 newest across all folders. Auto-push/newsletter mails (Bitcoin crawler, Google Alerts, etc.) are excluded by default so real correspondence surfaces first; set exclude_auto=false to see everything.",
        "inputSchema": {
            "type": "object",
            "properties": {
                "limit": {"type": "integer", "default": 20, "minimum": 1, "maximum": 100, "description": "How many emails to return."},
                "folder": {"type": "string", "default": "All", "description": "View name substring: 'All', 'Inbox', 'Drafts', 'Sent'. Use notes_list_folders to see all."},
                "newest_first": {"type": "boolean", "default": True},
                "exclude_auto": {"type": "boolean", "default": True, "description": "Hide auto-push/newsletter mails (rules in auto_filters.json). Default true; set false to include everything."},
            },
        },
    },
    {
        "name": "notes_get_mail",
        "description": "Read one email's full body, recipients, and attachment list by its universal_id (from notes_list_mail).",
        "inputSchema": {
            "type": "object",
            "properties": {"universal_id": {"type": "string", "description": "The UniversalID returned by notes_list_mail."}},
            "required": ["universal_id"],
        },
    },
    {
        "name": "notes_search_mail",
        "description": "Find emails by sender name/address, subject keyword, or body keyword. Filters are AND-combined; at least one required. Set include_body=true to read the mail content directly in the results — use this to fetch mail content by person name or subject keyword in a single call (no separate notes_get_mail needed). Sender can be a Chinese name, a Notes address (CN=SomeName/OU=.../O=...), or an email address; matching is case-insensitive.",
        "inputSchema": {
            "type": "object",
            "properties": {
                "from": {"type": "string", "description": "Substring to match against sender (name, Notes address, or email)."},
                "subject": {"type": "string", "description": "Substring to match against subject."},
                "body_keyword": {"type": "string", "description": "Keyword to find in body."},
                "limit": {"type": "integer", "default": 20, "minimum": 1, "maximum": 100},
                "include_body": {"type": "boolean", "default": False, "description": "If true, each result includes its body text and attachment list — use to read content found by name/subject."},
                "body_limit": {"type": "integer", "default": 4000, "minimum": 100, "maximum": 20000, "description": "Max body chars per result when include_body is true."},
                "newest_first": {"type": "boolean", "default": True, "description": "Sort results newest-first by date."},
                "exclude_auto": {"type": "boolean", "default": False, "description": "Drop auto-push/newsletter mails (rules in auto_filters.json). Default false for search, so you can still find your own crawler/alert mails; set true to focus on real correspondence."},
            },
        },
    },
    {
        "name": "notes_download_attachment",
        "description": "Download attachments from a mail to a local directory by its universal_id. Use after notes_get_mail / notes_search_mail to pull the actual file(s). Default downloads ALL attachments; pass a specific filename, substring, or 1-based index to select one. By default collisions are auto-renamed (e.g. 'f.pdf' -> 'f (1).pdf') so re-runs never clobber — set overwrite=true to replace.",
        "inputSchema": {
            "type": "object",
            "properties": {
                "universal_id": {"type": "string", "description": "The UniversalID (or NoteID) of the mail, from notes_list_mail / notes_get_mail."},
                "out_dir": {"type": "string", "description": "Target directory for the downloaded file(s); created if missing."},
                "attachment": {"type": "string", "default": "*", "description": "Which attachment(s) to pull. '*' (default) = all; a filename = exact match; a substring = any name containing it; a number string like '1' = 1-based index into the attachment list."},
                "overwrite": {"type": "boolean", "default": False, "description": "If false (default), name collisions are auto-renamed to '(1)','(2)',... ; if true, existing files are overwritten."},
            },
            "required": ["universal_id", "out_dir"],
        },
    },
    {
        "name": "notes_send_mail",
        "description": "Send a new email. Recipients can be a single address, or comma/semicolon-separated, or a list. Use notes_find_contact first if you only know a name. Supports file attachments.",
        "inputSchema": {
            "type": "object",
            "properties": {
                "to": {"type": "string", "description": "Recipient(s): email, Notes address, or name."},
                "subject": {"type": "string"},
                "body": {"type": "string"},
                "cc": {"type": "string", "description": "Optional CC."},
                "bcc": {"type": "string", "description": "Optional BCC."},
                "attachments": {
                    "description": "Optional file path(s) to attach. Accepts a single path string or a list of path strings.",
                    "oneOf": [
                        {"type": "string"},
                        {"type": "array", "items": {"type": "string"}},
                    ],
                },
            },
            "required": ["to", "subject", "body"],
        },
    },
    {
        "name": "notes_reply_mail",
        "description": "Reply to an email by its universal_id. Set reply_all=True to reply to everyone (original From + CC).",
        "inputSchema": {
            "type": "object",
            "properties": {
                "universal_id": {"type": "string"},
                "body": {"type": "string"},
                "reply_all": {"type": "boolean", "default": False},
            },
            "required": ["universal_id", "body"],
        },
    },
    {
        "name": "notes_find_contact",
        "description": "Look up a person in the Notes address book by name. Returns internet + Notes email addresses you can pass to notes_send_mail.",
        "inputSchema": {
            "type": "object",
            "properties": {"name": {"type": "string", "description": "Person name (or part of it) to search."}},
            "required": ["name"],
        },
    },
    {
        "name": "notes_list_folders",
        "description": "List all folder/view names in the mail database. Use these as the `folder` argument to notes_list_mail.",
        "inputSchema": {"type": "object"},
    },
    {
        "name": "notes_whoami",
        "description": "Health check: return the logged-in Notes user identity. Confirms the COM session is alive.",
        "inputSchema": {"type": "object"},
    },
]


# ── tool dispatch ─────────────────────────────────────────────────────


def _ok_text(text):
    return {"content": [{"type": "text", "text": str(text)}]}


def _err_text(message):
    return {"content": [{"type": "text", "text": str(message)}], "isError": True}


def call_tool(name, args):
    c = get_client()
    if name == "notes_list_mail":
        data = c.list_mail(
            limit=args.get("limit", 20),
            folder=args.get("folder", "All"),
            newest_first=args.get("newest_first", True),
            exclude_auto=args.get("exclude_auto", True),
        )
        return _ok_text(json.dumps(data, ensure_ascii=False, indent=2, default=str))
    if name == "notes_get_mail":
        data = c.get_mail(args["universal_id"])
        return _ok_text(json.dumps(data, ensure_ascii=False, indent=2, default=str))
    if name == "notes_search_mail":
        data = c.search_mail(
            from_filter=args.get("from", ""),
            subject_filter=args.get("subject", ""),
            body_keyword=args.get("body_keyword", ""),
            limit=args.get("limit", 20),
            include_body=args.get("include_body", False),
            body_limit=args.get("body_limit", 4000),
            newest_first=args.get("newest_first", True),
            exclude_auto=args.get("exclude_auto", False),
        )
        return _ok_text(json.dumps(data, ensure_ascii=False, indent=2, default=str))
    if name == "notes_download_attachment":
        data = c.download_attachment(
            universal_id=args["universal_id"],
            out_dir=args["out_dir"],
            attachment=args.get("attachment", "*"),
            overwrite=args.get("overwrite", False),
        )
        return _ok_text(json.dumps(data, ensure_ascii=False, indent=2, default=str))
    if name == "notes_send_mail":
        data = c.send_mail(
            to=args["to"], subject=args["subject"], body=args["body"],
            cc=args.get("cc", ""), bcc=args.get("bcc", ""),
            attachments=args.get("attachments"),
        )
        return _ok_text("Sent.\n" + json.dumps(data, ensure_ascii=False, indent=2))
    if name == "notes_reply_mail":
        data = c.reply_mail(
            universal_id=args["universal_id"], body=args["body"],
            reply_all=args.get("reply_all", False),
        )
        return _ok_text("Reply sent.\n" + json.dumps(data, ensure_ascii=False, indent=2))
    if name == "notes_find_contact":
        data = c.find_contact(args["name"])
        return _ok_text(json.dumps(data, ensure_ascii=False, indent=2))
    if name == "notes_list_folders":
        data = c.list_folders()
        return _ok_text(json.dumps(data, ensure_ascii=False, indent=2))
    if name == "notes_whoami":
        return _ok_text("Notes user: " + c.whoami())
    return _err_text(f"Unknown tool: {name}")


# ── JSON-RPC loop ─────────────────────────────────────────────────────


def _send(obj):
    sys.stdout.write(json.dumps(obj, ensure_ascii=False) + "\n")
    sys.stdout.flush()


def handle(req):
    """Process one JSON-RPC request. Returns response dict or None (notification)."""
    method = req.get("method")
    req_id = req.get("id")
    params = req.get("params", {})

    if method == "initialize":
        return {
            "jsonrpc": "2.0",
            "id": req_id,
            "result": {
                "protocolVersion": PROTOCOL_VERSION,
                "capabilities": {"tools": {}},
                "serverInfo": {"name": SERVER_NAME, "version": SERVER_VERSION},
            },
        }
    if method == "notifications/initialized":
        return None
    if method == "tools/list":
        return {"jsonrpc": "2.0", "id": req_id, "result": {"tools": TOOLS}}
    if method == "tools/call":
        name = params.get("name")
        args = params.get("arguments", {}) or {}
        try:
            result = call_tool(name, args)
            return {"jsonrpc": "2.0", "id": req_id, "result": result}
        except Exception as e:
            tb = traceback.format_exc()
            return {
                "jsonrpc": "2.0",
                "id": req_id,
                "result": _err_text(f"{type(e).__name__}: {e}\n{tb}"),
            }
    # unknown method
    return {
        "jsonrpc": "2.0",
        "id": req_id,
        "error": {"code": -32601, "message": f"Method not found: {method}"},
    }


def main():
    # Force UTF-8 on stdout/stderr — Lotus Notes returns Chinese strings and
    # the Windows console defaults to GBK, which would mojibake the JSON we
    # emit. MCP hosts read bytes, not the console, so this must be set in-process.
    try:
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")
    except Exception:
        pass  # older Python without reconfigure; PYTHONIOENCODING handles it

    # Notes ID lock check: warn early if the Notes client is running,
    # so the caller gets a clear message instead of a cryptic COM error.
    try:
        import subprocess
        out = subprocess.check_output(
            ["tasklist", "/FI", "IMAGENAME eq nlnotes.exe"], stderr=subprocess.DEVNULL
        ).decode("ascii", "ignore")
        if "nlnotes.exe" in out.lower():
            sys.stderr.write(
                "WARNING: IBM Notes client (nlnotes.exe) is running. "
                "The ID file is locked — close Notes before using mail tools.\n"
            )
    except Exception:
        pass  # not worth failing over

    for line in sys.stdin:
        line = line.strip()
        if not line:
            continue
        try:
            req = json.loads(line)
        except json.JSONDecodeError as e:
            _send({"jsonrpc": "2.0", "id": None, "error": {"code": -32700, "message": f"Parse error: {e}"}})
            continue
        resp = handle(req)
        if resp is not None:  # notifications return None → no response
            _send(resp)


if __name__ == "__main__":
    main()
