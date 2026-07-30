"""
Lotus Notes COM client wrapper.

All Lotus Notes automation goes through this module. It owns the single
NotesSession (ID file can only be held by one process at a time, so the
whole server keeps one long-lived session).

Design:
- Lazy init: session created on first use, reused afterwards.
- Config loaded from config.json next to this file (or via env override).
- Every public method returns plain Python data (dicts/lists), never COM
  objects — keeps the MCP tool layer trivial and JSON-serializable.
"""

import json
import os
from contextlib import contextmanager

import win32com.client

# Resolve paths relative to project root (parent of src/)
_PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def _load_config():
    """Load config.json from project root. Env vars override file values."""
    config_path = os.path.join(_PROJECT_ROOT, "config.json")
    cfg = {}
    if os.path.exists(config_path):
        with open(config_path, "r", encoding="utf-8") as f:
            cfg = json.load(f)
    # env overrides (handy for testing without editing the file)
    if os.environ.get("NOTES_PASSWORD"):
        cfg["notes_id_password"] = os.environ["NOTES_PASSWORD"]
    if os.environ.get("NOTES_ID_FILE"):
        cfg["notes_id_file"] = os.environ["NOTES_ID_FILE"]
    return cfg


def _load_auto_filters():
    """Load auto-push/newsletter filter rules from auto_filters.json.

    Returns (from_substrings, subject_substrings) — both lists of lowercase
    substrings. A mail is considered 'auto' if its From OR Subject contains
    any of these. Returns ([], []) if the file is missing so the feature
    degrades to no-op rather than crashing."""
    path = os.path.join(_PROJECT_ROOT, "auto_filters.json")
    from_subs = []
    subj_subs = []
    if os.path.exists(path):
        try:
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)
            from_subs = [str(s).lower() for s in data.get("from_filters", [])]
            subj_subs = [str(s).lower() for s in data.get("subject_filters", [])]
        except Exception:
            pass
    return from_subs, subj_subs


# loaded once at import; the file is meant to be edited then server restarted
_FROM_FILTERS, _SUBJECT_FILTERS = _load_auto_filters()


def _is_auto_mail(from_str, subject_str):
    """True if a mail matches any auto-push filter rule (case-insensitive)."""
    f = (from_str or "").lower()
    s = (subject_str or "").lower()
    for sub in _FROM_FILTERS:
        if sub and sub in f:
            return True
    for sub in _SUBJECT_FILTERS:
        if sub and sub in s:
            return True
    return False


class NotesClient:
    """Thin wrapper over Lotus.NotesSession COM object."""

    def __init__(self):
        self._session = None
        self._mail_db = None
        self._config = _load_config()

    # ── connection lifecycle ──────────────────────────────────────────

    def _ensure_session(self):
        """Create + Initialize the session once, reuse on subsequent calls."""
        if self._session is not None:
            return self._session
        pwd = self._config.get("notes_id_password", "")
        if not pwd:
            raise RuntimeError(
                "notes_id_password missing in config.json — cannot unlock ID file"
            )
        self._session = win32com.client.Dispatch("Lotus.NotesSession")
        self._session.Initialize(pwd)
        return self._session

    def _ensure_mail_db(self):
        """Open the current user's mail database once, reuse afterwards."""
        if self._mail_db is not None:
            return self._mail_db
        s = self._ensure_session()
        server = self._config.get("mail_server") or s.GetEnvironmentString("MailServer", True)
        mailfile = self._config.get("mail_file") or s.GetEnvironmentString("MailFile", True)
        db = s.GetDatabase(server, mailfile, False)
        if not db.IsOpen:
            raise RuntimeError(f"Could not open mail db {server}!!{mailfile}")
        self._mail_db = db
        return db

    def whoami(self):
        """Return the logged-in user identity — handy health check."""
        return str(self._ensure_session().UserName)

    # ── helpers ────────────────────────────────────────────────────────

    def _find_view(self, name_substr):
        """Find a view by substring match on its name (case-sensitive).
        Notes view names are awkward to pass verbatim (parens, localization),
        so we iterate the Views collection and match on a substring."""
        db = self._ensure_mail_db()
        for v in db.Views:
            if name_substr in str(v.Name):
                return v
        return None

    @staticmethod
    def _item(doc, field):
        """Safely read a single-value item from a doc. Returns '' if absent."""
        try:
            vals = doc.GetItemValue(field)
            if vals:
                return str(vals[0])
        except Exception:
            pass
        return ""

    @staticmethod
    def _doc_summary(doc, index=None):
        """Build the standard summary dict used by list/search results."""
        return {
            "index": index,
            "universal_id": str(doc.UniversalID),
            "subject": NotesClient._item(doc, "Subject"),
            "from": NotesClient._item(doc, "From"),
            "posted_date": NotesClient._item(doc, "PostedDate"),
        }

    # ── reading ────────────────────────────────────────────────────────

    def list_mail(self, limit=20, folder="All", newest_first=True, exclude_auto=True):
        """List up to `limit` docs from a folder view. `folder` is matched
        as a substring against view names: 'All', 'Inbox', 'Drafts', ...

        exclude_auto=True (default) drops auto-push/newsletter mails (Bitcoin
        crawler, Google Alerts, bank SMS-style digests, etc.) per the rules in
        auto_filters.json, so real correspondence surfaces to the top. To see
        everything including auto-push, pass exclude_auto=False.

        When filtering, we keep walking past filtered docs until we collect
        `limit` real mails, capped at `limit * 20` scanned docs so a very
        noisy mailbox doesn't trigger a full scan."""
        view = self._find_view(folder)
        if view is None:
            raise RuntimeError(f"No view matching '{folder}' found")
        results = []
        doc = view.GetLastDocument() if newest_first else view.GetFirstDocument()
        next_fn = view.GetPrevDocument if newest_first else view.GetNextDocument
        scanned = 0
        # scan cap: when filtering we may need to read many more docs than we keep
        scan_cap = limit * 20 if exclude_auto else limit
        while doc is not None and len(results) < limit and scanned < scan_cap:
            if exclude_auto:
                frm = self._item(doc, "From")
                subj = self._item(doc, "Subject")
                if _is_auto_mail(frm, subj):
                    doc = next_fn(doc)
                    scanned += 1
                    continue
            results.append(self._doc_summary(doc, len(results) + 1))
            doc = next_fn(doc)
            scanned += 1
        return results

    def _get_doc_by_id(self, doc_id):
        """Resolve a doc by UniversalID OR NoteID.

        The previous code passed the 32-hex UniversalID to GetDocumentByID,
        but that method actually expects a *NoteID* (a small decimal/hex int).
        The mismatch made every notes_get_mail / notes_reply_mail call fail
        with 'No document'. UniversalID must go through GetDocumentByUNID.

        We accept either: a 32-char hex string -> UNID, otherwise treat the
        value as a NoteID and fall back to GetDocumentByID."""
        db = self._ensure_mail_db()
        doc = None
        doc_id = str(doc_id).strip()
        # UniversalID is 32 hex chars; NoteID is short (<=8 hex).
        if len(doc_id) == 32 and all(ch in "0123456789abcdefABCDEF" for ch in doc_id):
            try:
                doc = db.GetDocumentByUNID(doc_id)
            except Exception:
                doc = None
        if doc is None:
            try:
                doc = db.GetDocumentByID(doc_id)
            except Exception:
                doc = None
        return doc

    def get_mail(self, universal_id):
        """Read one mail by its UniversalID (or NoteID): returns subject/from/body/attachments."""
        doc = self._get_doc_by_id(universal_id)
        if doc is None:
            raise RuntimeError(f"No document with id {universal_id}")
        body = ""
        try:
            rt = doc.GetFirstItem("Body")
            if rt is not None:
                body = str(rt.Text)
        except Exception:
            body = self._item(doc, "Body")
        attachments = self._collect_attachment_names(doc)
        return {
            "universal_id": str(doc.UniversalID),
            "note_id": str(doc.NoteID),
            "subject": self._item(doc, "Subject"),
            "from": self._item(doc, "From"),
            "to": self._item(doc, "SendTo"),
            "cc": self._item(doc, "CopyTo"),
            "bcc": self._item(doc, "BlindCopyTo"),
            "posted_date": self._item(doc, "PostedDate"),
            "body": body,
            "attachments": attachments,
        }

    @staticmethod
    def _iter_embedded_objects(rt_item):
        """Yield attachment EmbeddedObjects from a RichText item, robust to
        whether win32com hands us a COM collection (with .Count/.Item) or a
        plain tuple/list. Used by _collect_attachment_names and
        download_attachment. Only yields objects the caller can inspect."""
        try:
            embs = rt_item.EmbeddedObjects
        except Exception:
            return
        if embs is None:
            return
        # COM collection style: has .Count and .Item(i) (1-based)
        if hasattr(embs, "Count") and hasattr(embs, "Item"):
            try:
                count = int(embs.Count)
            except Exception:
                count = 0
            for i in range(1, count + 1):
                try:
                    yield embs.Item(i)
                except Exception:
                    pass
            return
        # tuple / list style: 0-based
        if isinstance(embs, (tuple, list)):
            for obj in embs:
                try:
                    yield obj
                except Exception:
                    pass

    def _collect_attachment_names(self, doc):
        """Collect attachment filenames from a doc via two strategies:
        (1) $File items (item.Type == 1084 = ATTACHMENT), and
        (2) RichText 'Body' EmbeddedObjects (Type == 1454 = EMBED_ATTACHMENT).
        Returns a de-duplicated list of filenames, preserving first-seen order.
        Used by both get_mail (listing) and download_attachment (extraction)."""
        names = []
        seen = set()

        # Strategy 1: $File items — most reliable for filenames
        if doc.HasEmbedded:
            for item in doc.Items:
                try:
                    if item.Type == 1084:
                        vals = item.Values
                        name = str(vals[0]) if vals and vals[0] else ""
                        if name and name not in seen:
                            seen.add(name)
                            names.append(name)
                except Exception:
                    pass

        # Strategy 2: Body RichText EmbeddedObjects — catches inline attachments
        # that sometimes don't surface as $File items.
        try:
            rt = doc.GetFirstItem("Body")
            if rt is not None:
                for obj in self._iter_embedded_objects(rt):
                    try:
                        if obj.Type == self.EMBED_ATTACHMENT:  # 1454
                            nm = str(obj.Source) or str(obj.Name)
                            if nm and nm not in seen:
                                seen.add(nm)
                                names.append(nm)
                    except Exception:
                        pass
        except Exception:
            pass

        return names

    def download_attachment(self, universal_id, out_dir, attachment="*", overwrite=False):
        """Download one or all attachments from a mail to out_dir.

        Args:
            universal_id: doc UNID (or NoteID) to read attachments from.
            out_dir: target directory; created if missing.
            attachment: which attachment to pull.
                "*" (default)  -> all attachments;
                a filename      -> that exact filename;
                a substring     -> any attachment whose name contains it;
                an int / digit-string -> 1-based index into the attachment list.
            overwrite: if False (default), collisions are auto-renamed
                (e.g. "f.pdf" -> "f (1).pdf") so re-downloads never clobber.

        Returns dict:
            out_dir, requested, downloaded: [{name, path, size, renamed_from?}],
            skipped: [{name, reason}], errors: [{name, error}]
        """
        doc = self._get_doc_by_id(universal_id)
        if doc is None:
            raise RuntimeError(f"No document with id {universal_id}")

        all_names = self._collect_attachment_names(doc)

        # Resolve which attachments to pull based on the `attachment` selector.
        sel = str(attachment).strip()
        if sel in ("", "*", "all", "ALL"):
            targets = list(all_names)
        elif sel.lstrip("-").isdigit():
            idx = int(sel)
            if idx < 1 or idx > len(all_names):
                raise RuntimeError(
                    f"attachment index {idx} out of range (1..{len(all_names)})"
                )
            targets = [all_names[idx - 1]]
        else:
            # exact match first, then substring fallback
            targets = [n for n in all_names if n == sel]
            if not targets:
                targets = [n for n in all_names if sel.lower() in n.lower()]
            if not targets:
                raise RuntimeError(
                    f"No attachment matching '{sel}'. Available: {all_names}"
                )

        # Prepare output directory.
        out_dir = os.path.abspath(out_dir)
        os.makedirs(out_dir, exist_ok=True)

        def _safe_path(name):
            """Return an output path that doesn't clobber an existing file
            (unless overwrite=True). Mirrors Windows "(n)" suffix style."""
            base = os.path.join(out_dir, name)
            if overwrite or not os.path.exists(base):
                return base, None
            stem, ext = os.path.splitext(name)
            n = 1
            while True:
                cand = os.path.join(out_dir, f"{stem} ({n}){ext}")
                if not os.path.exists(cand):
                    return cand, name
                n += 1

        downloaded = []
        skipped = []
        errors = []

        for name in targets:
            out_path, renamed_from = _safe_path(name)
            extracted = False

            # Path A: RichText EmbeddedObject direct extraction (preferred —
            # handles inline attachments and preserves the original stream).
            try:
                rt = doc.GetFirstItem("Body")
                if rt is not None:
                    for obj in self._iter_embedded_objects(rt):
                        try:
                            if obj.Type == self.EMBED_ATTACHMENT:
                                src = str(obj.Source) or str(obj.Name)
                                if src == name:
                                    obj.ExtractFile(out_path)
                                    extracted = True
                                    break
                        except Exception:
                            pass
            except Exception as e:
                errors.append({"name": name, "error": f"EmbeddedObject: {e}"})

            # Path B: doc.GetAttachment(name).ExtractFile — catches attachments
            # stored outside the Body RichText (e.g. $File-only items).
            if not extracted:
                try:
                    att = doc.GetAttachment(name)
                    if att is not None:
                        att.ExtractFile(out_path)
                        extracted = True
                except Exception as e:
                    errors.append({"name": name, "error": f"GetAttachment: {e}"})

            if extracted:
                try:
                    size = os.path.getsize(out_path)
                except Exception:
                    size = None
                entry = {"name": os.path.basename(out_path), "path": out_path, "size": size}
                if renamed_from:
                    entry["renamed_from"] = renamed_from
                downloaded.append(entry)
            else:
                # Only keep the "not extracted" skip if no real error was logged
                if not any(e["name"] == name for e in errors):
                    skipped.append({"name": name, "reason": "extraction returned no file"})

        return {
            "universal_id": str(doc.UniversalID),
            "out_dir": out_dir,
            "available": all_names,
            "requested": targets,
            "downloaded": downloaded,
            "skipped": skipped,
            "errors": errors,
        }

    @staticmethod
    def _formula_escape(s):
        """Escape a user string for safe embedding inside a Notes @Formula
        double-quoted string. Backslash and double-quote must be escaped."""
        return str(s).replace("\\", "\\\\").replace('"', '\\"')

    def search_mail(self, from_filter="", subject_filter="", body_keyword="",
                    limit=20, include_body=False, body_limit=4000, newest_first=True,
                    exclude_auto=False):
        """Search mail by sender / subject / body keyword. Filters are AND-combined;
        at least one is required.

        Uses db.Search with a Notes @Formula selection. Unlike db.FTSearch this
        needs NO full-text index, so it never hangs waiting for an index to be
        built (which was the previous 2-minute timeout failure).

        Args:
            from_filter: substring matched against From (sender name or address).
            subject_filter: substring matched against Subject.
            body_keyword: substring matched against Body text.
            limit: max results.
            include_body: if True, each result includes the body text (truncated
                to body_limit chars) — use this when you want to read content by
                person name or subject keyword in a single call.
            body_limit: max chars of body to return per result when include_body.
            newest_first: sort results by PostedDate descending.
            exclude_auto: drop auto-push/newsletter mails per auto_filters.json.
                Defaults to False for search (you usually *want* to search your
                own crawler/alert mails); set True to focus on real correspondence.

        Returns list of result dicts. Each has universal_id, note_id, subject,
        from, to, cc, posted_date; and body + attachments when include_body.
        """
        db = self._ensure_mail_db()
        # Build a Notes selection formula. @IsAvailable guards docs missing a
        # field so @Contains doesn't error on them. Each filter is case-insensitive
        # via @LowerCase.
        terms = []
        if from_filter:
            f = self._formula_escape(from_filter).lower()
            terms.append(f'@Contains(@LowerCase(From); "{f}")')
        if subject_filter:
            s = self._formula_escape(subject_filter).lower()
            terms.append(f'@Contains(@LowerCase(Subject); "{s}")')
        if body_keyword:
            b = self._formula_escape(body_keyword).lower()
            terms.append(f'@Contains(@LowerCase(Body); "{b}")')
        if not terms:
            raise RuntimeError("At least one of from_filter/subject_filter/body_keyword required")
        formula = " & ".join(terms)

        col = db.Search(formula, None, 0)
        results = []
        doc = col.GetFirstDocument()
        while doc is not None:
            results.append(doc)
            doc = col.GetNextDocument(doc)

        # sort by posted_date (string compare is fine for ISO-ish Notes dates)
        def posted(d):
            return self._item(d, "PostedDate")
        results.sort(key=posted, reverse=newest_first)

        out = []
        for d in results:
            if len(out) >= limit:
                break
            if exclude_auto and _is_auto_mail(self._item(d, "From"), self._item(d, "Subject")):
                continue
            row = self._doc_summary(d, len(out) + 1)
            row["note_id"] = str(d.NoteID)
            row["to"] = self._item(d, "SendTo")
            row["cc"] = self._item(d, "CopyTo")
            if include_body:
                body = ""
                try:
                    rt = d.GetFirstItem("Body")
                    if rt is not None:
                        body = str(rt.Text)
                except Exception:
                    body = self._item(d, "Body")
                row["body"] = body[:body_limit]
                row["attachments"] = self._collect_attachment_names(d)
            out.append(row)
        return out

    # ── writing ────────────────────────────────────────────────────────

    def _parse_recipients(self, value):
        """Accept a string or list, return a list of recipient strings."""
        if value is None or value == "":
            return []
        if isinstance(value, list):
            return [str(v) for v in value]
        # Notes accepts comma or semicolon separated
        return [r.strip() for r in str(value).replace(";", ",").split(",") if r.strip()]

    # EMBED_ATTACHMENT constant from the Notes C API / LotusScript
    EMBED_ATTACHMENT = 1454

    def _embed_attachments(self, rt, attachments):
        """Attach files into a RichText body item. `attachments` is a list of
        paths (strings). Non-existent files are skipped with a warning on stderr.
        Returns the list of basenames actually attached."""
        import sys as _sys
        attached = []
        for path in attachments:
            path = str(path).strip().strip('"')
            if not path:
                continue
            if not os.path.exists(path):
                _sys.stderr.write(f"send_mail: attachment not found, skipped: {path}\n")
                continue
            try:
                rt.EmbedObject(self.EMBED_ATTACHMENT, "", path)
                attached.append(os.path.basename(path))
            except Exception as e:
                _sys.stderr.write(f"send_mail: failed to attach {path}: {e}\n")
        return attached

    def send_mail(self, to, subject, body, cc="", bcc="", attachments=None):
        """Compose and send a new memo immediately.

        `attachments` (optional): path string, or list of path strings. Each
        existing file is embedded into the Body rich-text item as an attachment."""
        db = self._ensure_mail_db()
        doc = db.CreateDocument()
        doc.ReplaceItemValue("Form", "Memo")
        doc.ReplaceItemValue("SendTo", self._parse_recipients(to))
        if cc:
            doc.ReplaceItemValue("CopyTo", self._parse_recipients(cc))
        if bcc:
            doc.ReplaceItemValue("BlindCopyTo", self._parse_recipients(bcc))
        doc.ReplaceItemValue("Subject", subject)
        rt = doc.CreateRichTextItem("Body")
        rt.AppendText(body)
        attached = []
        if attachments:
            if isinstance(attachments, str):
                attachments = [attachments]
            attached = self._embed_attachments(rt, attachments)
        doc.Send(False)
        return {
            "sent": True,
            "to": self._parse_recipients(to),
            "subject": subject,
            "attachments": attached,
        }

    def reply_mail(self, universal_id, body, reply_all=False):
        """Reply to a mail. reply_all=True puts original CC into the new To."""
        orig = self._get_doc_by_id(universal_id)
        if orig is None:
            raise RuntimeError(f"No document with id {universal_id}")
        orig_from = self._item(orig, "From")
        orig_cc = orig.GetItemValue("CopyTo") if reply_all else []
        db = self._ensure_mail_db()
        doc = db.CreateDocument()
        doc.ReplaceItemValue("Form", "Memo")
        doc.ReplaceItemValue("SendTo", [orig_from])
        if orig_cc:
            doc.ReplaceItemValue("CopyTo", [str(c) for c in orig_cc if str(c)])
        subj = self._item(orig, "Subject")
        if not subj.lower().startswith("re:"):
            subj = "Re: " + subj
        doc.ReplaceItemValue("Subject", subj)
        rt = doc.CreateRichTextItem("Body")
        rt.AppendText(body)
        doc.Send(False)
        return {
            "sent": True,
            "to": [orig_from],
            "reply_all": reply_all,
            "subject": subj,
        }

    # ── contacts ───────────────────────────────────────────────────────

    def find_contact(self, name):
        """Look up a person by name substring in the local names.nsf address book.
        Returns matching contacts with their internet + Notes addresses."""
        s = self._ensure_session()
        # local address book is names.nsf on the local machine
        nab_path = self._config.get("names_nsf", "names.nsf")
        nab = s.GetDatabase("", nab_path, False)
        if not nab.IsOpen:
            raise RuntimeError(f"Could not open address book {nab_path}")
        view = nab.GetView("($Users)")
        if view is None:
            # try People view as fallback
            for v in nab.Views:
                if "People" in str(v.Name) or "User" in str(v.Name):
                    view = v
                    break
        if view is None:
            raise RuntimeError("No people view found in address book")
        results = []
        doc = view.GetFirstDocument()
        name_lower = name.lower()
        while doc is not None:
            full = self._item(doc, "FullName").lower()
            last = self._item(doc, "LastName").lower()
            first = self._item(doc, "FirstName").lower()
            if name_lower in full or name_lower in last or name_lower in first:
                results.append({
                    "full_name": self._item(doc, "FullName"),
                    "internet_address": self._item(doc, "InternetAddress"),
                    "notes_address": self._item(doc, "ShortName"),
                    "company": self._item(doc, "CompanyName"),
                })
            doc = view.GetNextDocument(doc)
        return results

    def list_folders(self):
        """List all view/folder names in the mail db — helps pick the right
        `folder` argument for list_mail."""
        db = self._ensure_mail_db()
        names = []
        for v in db.Views:
            names.append(str(v.Name))
        return names


# module-level singleton: the whole process keeps one session
_client = None


def get_client():
    global _client
    if _client is None:
        _client = NotesClient()
    return _client
