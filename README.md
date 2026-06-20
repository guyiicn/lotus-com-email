# lotus-com-email

A self-contained **MCP server** that gives an AI agent (ZCode, Claude Desktop,
Cursor, ...) full mail-client capabilities over **Lotus Notes / IBM Notes (HCL
Notes)** via the Lotus COM API (`Lotus.NotesSession`).

Read, search, reply, send (with attachments), and look up contacts — all as
clean MCP tools the agent can call directly. No screen-scraping, no UI
automation: it talks to Notes through its native COM automation interface.

---

## How it works (the principle)

Lotus Notes is an old Java/SWT application with **no modern API**. Generic
"computer use" drivers (that walk the UI via UIA/ accessibility trees) time
out on it because SWT's accessibility provider is unresponsive. The one
reliable programmatic path into Notes is its **COM automation interface**
(`Lotus.NotesSession`), the same API VBScript and Notes macros have used for
decades.

Three hard constraints shape this project:

1. **COM is 32-bit only.** The Notes COM server (`nlsxbe.dll`) is a 32-bit
   in-proc server. A 64-bit process gets `REGDB_E_CLASSNOTREG` when trying to
   `CreateObject("Lotus.NotesSession")`. So the whole server must run as a
   **32-bit process**. This project bundles a 32-bit Python runtime for that.

2. **The ID file is single-occupancy.** Notes' identity file (`.id`) can be
   held by only one process at a time. The IBM Notes client takes that lock
   when it's running. So you must **close the Notes client** before using any
   mail tool here. The server checks and warns on stderr.

3. **The `mcp` PyPI SDK won't build on 32-bit Python.** It has a Rust/maturin
   backend with no 32-bit Windows wheels. So `src/server.py` implements the
   minimal MCP-over-stdio JSON-RPC surface **by hand** — `initialize`,
   `notifications/initialized`, `tools/list`, `tools/call`. It's ~200 lines,
   speaks exactly the protocol MCP hosts expect, and has zero third-party
   runtime deps beyond `pywin32`.

**Data flow:**

```
AI agent (ZCode/Claude/Cursor)
   │  stdio (JSON-RPC, one object per line)
   ▼
server.py  ──hand-written MCP loop──▶  notes_client.py
                                          │  win32com.client.Dispatch
                                          ▼
                                   Lotus.NotesSession (COM, 32-bit)
                                          │
                                          ▼
                                   Notes mail DB (.nsf) + address book
```

Every public method in `notes_client.py` returns plain JSON-serializable data
(dict/list), never COM objects — so the MCP tool layer stays trivial.

---

## Prerequisites

- **Windows** (Lotus Notes COM is Windows-only)
- **IBM / HCL Notes installed** (client or just the runtime) — provides
  `nlsxbe.dll` and the `.id` file
- **Your Notes ID password** (set during Notes setup)

---

## Setup (one-time)

### 1. Get a 32-bit Python runtime into `python/`

The `python/` directory is gitignored (too big, reproducible). Populate it:

```bat
:: download the 32-bit embeddable package (win32, NOT amd64) from python.org
:: e.g. https://www.python.org/ftp/python/3.12.8/python-3.12.8-embed-win32.zip
:: unzip into the project's python\ folder
```

Then enable site-packages and install pip + the one runtime dependency:

```bat
:: 1) edit python\python312._pth — uncomment the "import site" line and add:
::    Lib\site-packages
::    (the _pth should look like:)
::      python312.zip
::      .
::      Lib\site-packages
::
::      import site

:: 2) install pip (get-pip.py from https://bootstrap.pypa.io/get-pip.py)
python\python.exe get-pip.py

:: 3) install pywin32 (provides win32com, the COM bridge)
python\python.exe -m pip install pywin32
```

Verify it's 32-bit:
```bat
python\python.exe -c "import struct; print(struct.calcsize('P')*8, 'bit')"
:: -> 32 bit
```

### 2. Register the Lotus COM server (32-bit, run once as admin)

Notes' COM DLL must be registered with the **32-bit** regsvr32 (the one in
`SysWOW64`, not `System32`):

```bat
C:\Windows\SysWOW64\regsvr32.exe /s "C:\Program Files (x86)\IBM\Notes\nlsxbe.dll"
```

Adjust the path if Notes lives elsewhere on your machine.

### 3. Create `config.json` (your secrets — gitignored)

```bat
copy config.example.json config.json
```

Then edit `config.json` and fill in **your** values:

```json
{
  "notes_id_password": "your-id-password",
  "notes_id_file": "C:\\Users\\YOU\\AppData\\Local\\IBM\\Notes\\Data\\YOUR_NAME.id",
  "mail_server": "CN=YOUR_MAIL_SERVER/O=YOUR_ORG",
  "mail_file": "mail\\YOUR_NAME.nsf",
  "names_nsf": "names.nsf"
}
```

- `notes_id_password` — the password to unlock your `.id` file
- `notes_id_file` — absolute path to your Notes ID file
- `mail_server` / `mail_file` — your mail server DN and mailbox file (find
  them in Notes' location settings, or via `session.GetEnvironmentString`)
- `names_nsf` — the address book DB name (usually `names.nsf`)

`config.json` is in `.gitignore` — your password never gets committed.

### 4. (Optional) Tune auto-push filtering

`auto_filters.json` holds the rules `notes_list_mail`/`notes_search_mail` use
to hide newsletter/crawler noise (so real mail surfaces first). Edit freely —
add sender or subject substrings. Restart the server to pick up changes.

---

## Usage constraint: ID lock

The Notes ID file can be held by **only one process at a time**. Before using
any mail tool, **close the IBM Notes client** (`nlnotes.exe` / `notes2.exe`).
The server detects this on startup and warns on stderr if Notes is running.
This is a Notes limitation, not a bug here.

---

## Wire it into your AI agent

### ZCode / OpenCode

Add to `~/.zcode/cli/config.json` (or workspace config):

```json
{
  "mcpServers": {
    "lotus-notes-mail": {
      "command": "C:\\path\\to\\lotus-com-email\\run.bat"
    }
  }
}
```

Restart ZCode. The tools appear as `mcp__lotus-notes-mail__*`.

### Claude Desktop / Cursor / other MCP hosts

Point the server command at `run.bat` (Windows) in your host's MCP config.
`run.bat` sets UTF-8 env vars (`PYTHONUTF8=1`, `PYTHONIOENCODING=utf-8`) and
launches the bundled 32-bit Python — both are required because Notes returns
non-ASCII strings and the default Windows console codepage (GBK) would
mojibake the JSON emitted over stdio.

---

## Tools

| Tool | What it does | Key params |
|---|---|---|
| `notes_list_mail` | List newest N emails (sender/subject/date/universal_id) | `limit`, `folder`, `exclude_auto` (default true — hides newsletter/crawler noise) |
| `notes_get_mail` | Read one email's body, recipients, attachments | `universal_id` (or note_id) |
| `notes_search_mail` | Filter by from / subject / body keyword | `from`, `subject`, `body_keyword`, `include_body`, `exclude_auto` |
| `notes_send_mail` | Send new mail (to/cc/bcc + file attachments) | `to`, `subject`, `body`, `cc`, `bcc`, `attachments` |
| `notes_reply_mail` | Reply / reply-all | `universal_id`, `body`, `reply_all` |
| `notes_find_contact` | Look up a person → internet + Notes address | `name` |
| `notes_list_folders` | List view names (for the `folder` arg) | — |
| `notes_whoami` | Health check — logged-in user identity | — |

Notable capabilities:
- **`exclude_auto`** — `list_mail` defaults to hiding auto-push mails (Bitcoin
  crawlers, Google Alerts, bank digests) per `auto_filters.json`, so real
  correspondence surfaces first. Set `exclude_auto=false` to see everything.
- **`include_body`** on `search_mail` — read matched mails' content in one
  call (find-by-name-then-read in a single step).
- **`attachments`** on `send_mail` — pass a path or list of paths; embedded
  into the Body rich-text item via `rt.EmbedObject`.
- **Search via `db.Search` (@Formula)**, not `FTSearch` — no full-text index
  required, never hangs. Doc lookup by UNID via `GetDocumentByUNID`.

---

## Layout

```
lotus-com-email/
├── src/
│   ├── notes_client.py   # Lotus COM wrapper (the core)
│   └── server.py         # hand-written MCP-over-stdio server (8 tools)
├── config.json           # your secrets — NOT in git (.gitignore'd)
├── config.example.json   # template, safe to commit
├── auto_filters.json     # auto-push filter rules (editable)
├── run.bat               # launches server with bundled 32-bit Python + UTF-8
├── screenshot.ps1        # helper: capture foreground window (for attaching)
├── shot_now.bat          # helper: one-shot foreground screenshot
├── attachments/          # working dir for attachments
└── python/               # bundled 32-bit Python 3.12 (embeddable) — NOT in git
```

---

## Troubleshooting

| Symptom | Cause / fix |
|---|---|
| `REGDB_E_CLASSNOTREG` / can't create NotesSession | Running 64-bit Python. Must use the 32-bit runtime in `python/`. |
| COM class not registered | Register `nlsxbe.dll` with 32-bit regsvr32 (step 2 above). |
| "ID file is locked by another process" | The IBM Notes client is running. Close it (nlnotes.exe) first. |
| Chinese garbled in agent output | `run.bat` must set `PYTHONUTF8=1` + `PYTHONIOENCODING=utf-8` (it does). |
| `mcp` package install fails on 32-bit | Expected — that's why the server is hand-written, no `mcp` dependency. |
| `notes_get_mail` says "No document" | Old bug: was passing UNID to `GetDocumentByID` (wants NoteID). Fixed — now uses `GetDocumentByUNID`, accepts either. |

---

## License

Private project. Lotus Notes is a trademark of HCL; this project is
independent and not affiliated with or endorsed by HCL or IBM.
