"""Excel 365 live connector (MCP server) over Microsoft Graph.

Reads/writes cells in a co-authored OneDrive/SharePoint workbook while the team keeps editing.
Bulk analysis (map / find / trace / check / changes) runs on a read-only snapshot of the file;
writes go cell-by-cell through the Graph workbook API, are conflict-guarded and logged for undo.

Auth: browser sign-in (opens automatically on first use; device code as fallback) as the user,
Microsoft's "Microsoft Graph Command Line Tools" app; any Microsoft 365 work/school account can sign in
(set EXCEL365_TENANT, e.g. contoso.com, to restrict to one organisation). The refresh token is DPAPI-encrypted
in %USERPROFILE%\\.xlgraph\\refresh.dat. The password never passes through this server.

The MCP stdio protocol is implemented here on the standard library, so openpyxl is the only dependency
(installed by setup.ps1 into %LOCALAPPDATA%\\excel365\\lib).
"""
import atexit
import base64
import ctypes
import ctypes.wintypes as wt
import datetime as dt
import functools
import hashlib
import http.server
import inspect
import json
import os
import re
import secrets
import shutil
import socket
import sys
import threading
import time
import traceback
import typing
import urllib.error
import urllib.parse
import urllib.request
import webbrowser

# Try IPv4 addresses before IPv6. Where IPv6 is broken (it is on some office networks), Python otherwise waits
# Windows' 21 s connect timeout on every IPv6 address before falling back - minutes per call - while browsers,
# which race both, look perfectly fast. IPv6 is still used when there is no IPv4 route.
_getaddrinfo = socket.getaddrinfo
socket.getaddrinfo = lambda *a, **k: sorted(_getaddrinfo(*a, **k), key=lambda ai: ai[0] != socket.AF_INET)

_LIB = os.path.join(os.environ.get("LOCALAPPDATA", ""), "excel365", "lib")
if os.path.isdir(_LIB):
    sys.path.insert(0, _LIB)

import openpyxl
from openpyxl.formula.tokenizer import Token, Tokenizer
from openpyxl.formula.translate import Translator
from openpyxl.utils import get_column_letter, column_index_from_string, range_boundaries

VERSION = "1.4.0"

TENANT = os.environ.get("EXCEL365_TENANT", "").strip() or "organizations"  # "organizations" = any work/school account
CLIENT_ID = "14d82eec-204b-4c2f-b7e8-296a70dab67e"  # Microsoft Graph Command Line Tools
SCOPE = "https://graph.microsoft.com/Files.ReadWrite.All offline_access"
GRAPH = "https://graph.microsoft.com/v1.0"
STORE = os.path.join(os.environ["USERPROFILE"], ".xlgraph")
os.makedirs(STORE, exist_ok=True)
MAX_OUT = 40000
ERRORS = ("#REF!", "#DIV/0!", "#N/A", "#VALUE!", "#NAME?", "#NUM!", "#NULL!", "#SPILL!", "#CALC!")


# ---------------------------------------------------------------- DPAPI (PowerShell SecureString compatible)
class _Blob(ctypes.Structure):
    _fields_ = [("cbData", wt.DWORD), ("pbData", ctypes.POINTER(ctypes.c_char))]


def _dpapi(data: bytes, protect: bool) -> bytes:
    buf = ctypes.create_string_buffer(data, len(data))
    src, out = _Blob(len(data), buf), _Blob()
    fn = ctypes.windll.crypt32.CryptProtectData if protect else ctypes.windll.crypt32.CryptUnprotectData
    if not fn(ctypes.byref(src), None, None, None, None, 0, ctypes.byref(out)):
        raise RuntimeError("DPAPI failed - run xl_login again")
    try:
        return ctypes.string_at(out.pbData, out.cbData)
    finally:
        ctypes.windll.kernel32.LocalFree(out.pbData)


class NotSignedIn(Exception):
    pass


def _load_refresh() -> str:
    path = os.path.join(STORE, "refresh.dat")
    if not os.path.exists(path):
        raise NotSignedIn("Not signed in")
    blob = bytes.fromhex(open(path, encoding="utf-8-sig").read().strip())
    return _dpapi(blob, False).decode("utf-16-le")


def _save_refresh(rt: str):
    with open(os.path.join(STORE, "refresh.dat"), "w") as f:
        f.write(_dpapi(rt.encode("utf-16-le"), True).hex())


# ---------------------------------------------------------------- HTTP / Graph
def _post_form(url, fields):
    req = urllib.request.Request(url, data=urllib.parse.urlencode(fields).encode(), method="POST")
    for attempt in range(4):
        _net_check()
        try:
            with urllib.request.urlopen(req, timeout=TIMEOUT) as r:
                return json.loads(r.read())
        except urllib.error.HTTPError as e:
            return json.loads(e.read() or b"{}")
        except (urllib.error.URLError, ConnectionError, TimeoutError) as e:
            if isinstance(getattr(e, "reason", e), TimeoutError) or "timed out" in str(e):
                _net_stalled("Microsoft sign-in")
            if attempt == 3:
                raise
            time.sleep(1 + attempt)


_tok = {"access": None, "exp": 0.0}
_lock = threading.RLock()


def _token() -> str:
    with _lock:
        if _tok["access"] and time.time() < _tok["exp"]:
            return _tok["access"]
        r = _post_form(f"https://login.microsoftonline.com/{TENANT}/oauth2/v2.0/token", {
            "grant_type": "refresh_token", "client_id": CLIENT_ID, "refresh_token": _load_refresh(), "scope": SCOPE})
        if "access_token" not in r:
            raise NotSignedIn(f"Sign-in expired ({r.get('error')})")
        _save_refresh(r["refresh_token"])
        _tok.update(access=r["access_token"], exp=time.time() + r["expires_in"] - 120)
        return _tok["access"]


class GraphError(Exception):
    pass


TIMEOUT = 30       # seconds per Graph call
NET_PAUSE = 60     # after a call times out, further calls fail at once for this long instead of each hanging


class NetworkDown(GraphError):
    pass


_net = {"down_until": 0.0, "why": ""}


def _net_check():
    if time.time() < _net["down_until"]:
        raise NetworkDown(f"The connection to Microsoft 365 is not responding ({_net['why']}). Nothing was sent. "
                          f"Check this PC's internet/VPN and retry in about {int(_net['down_until'] - time.time()) + 5} s. "
                          "If Excel desktop shows an upload or merge error, it is the same connection problem.")


def _net_stalled(what):
    _net.update(down_until=time.time() + NET_PAUSE, why=f"{what} timed out at {time.strftime('%H:%M:%S')}")
    raise NetworkDown(f"The connection to Microsoft 365 is not responding ({_net['why']}). This request may or may not "
                      "have reached Microsoft: re-read the cells before writing them again. Check this PC's internet/VPN "
                      f"and retry in about {NET_PAUSE} s. If Excel desktop shows an upload or merge error, it is the same "
                      "connection problem.")


def _graph(method, path, body=None, raw=False, session=None, timeout=TIMEOUT):
    # encode spaces etc. in paths like /worksheets('MB5B Costing'); existing %xx escapes are kept
    url = path if path.startswith("http") else GRAPH + urllib.parse.quote(path, safe="/()'=?&$,:!@*+;%-._~[]")
    headers = {"Authorization": f"Bearer {_token()}"}
    data = None
    if body is not None:
        data = json.dumps(body).encode()
        headers["Content-Type"] = "application/json"
    if session:
        headers["workbook-session-id"] = session
    req = urllib.request.Request(url, data=data, headers=headers, method=method)
    for attempt in range(4):
        _net_check()
        try:
            with urllib.request.urlopen(req, timeout=timeout) as r:
                b = r.read()
                _net["down_until"] = 0.0
                return b if raw else (json.loads(b) if b else {})
        except urllib.error.HTTPError as e:
            msg = e.read().decode(errors="replace")
            if e.code in (429, 503, 504) and attempt < 3:
                time.sleep(min(int(e.headers.get("Retry-After", 2 ** attempt)), 10))
                continue
            raise GraphError(f"{e.code} {method} {path}: {msg[:600]}")
        except (urllib.error.URLError, ConnectionError, TimeoutError) as e:
            reason = getattr(e, "reason", e)
            # a stalled call (timeout) means the connection is down or very slow: no retry, and later calls fail
            # at once for NET_PAUSE seconds instead of each hanging. A dropped connection (WinError 10054 on this
            # network) fails fast and usually works on retry: up to 4 tries.
            if isinstance(reason, TimeoutError) or "timed out" in str(reason):
                _net_stalled(f"{method} {path.split('?')[0][-60:]}")
            if attempt == 3:
                raise GraphError(f"network error on {method} {path} after {attempt + 1} tries: {e}")
            time.sleep(0.5 * (attempt + 1))


def _download(url):
    """Fetch a pre-authenticated download URL (no Graph token), with retries."""
    for attempt in range(4):
        _net_check()
        try:
            with urllib.request.urlopen(url, timeout=120) as r:
                return r.read()
        except (urllib.error.URLError, ConnectionError, TimeoutError) as e:
            if isinstance(getattr(e, "reason", e), TimeoutError) or "timed out" in str(e):
                _net_stalled("workbook download")
            if attempt == 3:
                raise
            time.sleep(1 + attempt)


# ---------------------------------------------------------------- workbook state
_state_file = os.path.join(STORE, "workbooks.json")


def _state():
    return json.load(open(_state_file)) if os.path.exists(_state_file) else {"current": None, "books": {}}


def _save_state(s):
    json.dump(s, open(_state_file, "w"), indent=1)


# Which workbook a tool acts on is decided per CHAT, never shared between chats.
# Each Claude Code chat runs its own claude.exe with --resume=<chat id> on its command line, and the app restarts
# that process (and this server) when a chat is resumed after a pause. The chat id survives the restart, so the
# open workbook is remembered in chats.json under it. Where no chat id can be found (e.g. Claude Desktop chat),
# the workbook is kept in memory only - it is never borrowed from another chat.
_current = {"item": None}
_CHATS = os.path.join(STORE, "chats.json")
_chat = {"id": None, "done": threading.Event()}


def _find_chat_id():
    try:
        ps = ("$p = Get-CimInstance Win32_Process -Filter \"ProcessId=%d\"; "
              "for ($i = 0; $i -lt 6 -and $p; $i++) { "
              "if ($p.CommandLine -match '--(resume|session-id)[= ]+\"?([0-9a-fA-F-]{36})') { $Matches[2]; break }; "
              "$p = Get-CimInstance Win32_Process -Filter \"ProcessId=$($p.ParentProcessId)\" }") % os.getpid()
        import subprocess
        r = subprocess.run(["powershell", "-NoProfile", "-NonInteractive", "-Command", ps], capture_output=True,
                           text=True, timeout=20, creationflags=0x08000000)  # CREATE_NO_WINDOW
        m = re.search(r"[0-9a-fA-F-]{36}", r.stdout)
        _chat["id"] = m.group(0).lower() if m else None
    except Exception:
        _chat["id"] = None
    finally:
        _chat["done"].set()


threading.Thread(target=_find_chat_id, daemon=True).start()


def _chat_id():
    _chat["done"].wait(25)
    return _chat["id"]


def _chats():
    try:
        return json.load(open(_CHATS))
    except Exception:
        return {}


def _remember_workbook(item):
    cid = _chat_id()
    if not cid:
        return
    c = {k: v for k, v in _chats().items() if time.time() - v.get("ts", 0) < 30 * 86400}
    c[cid] = {"item": item, "ts": time.time()}
    json.dump(c, open(_CHATS, "w"), indent=1)


def _set_current(item):
    _current["item"] = item
    _remember_workbook(item)


def _book():
    if not _current["item"]:
        cid = _chat_id()
        saved = _chats().get(cid) if cid else None
        if saved and saved["item"] in _state()["books"]:
            _current["item"] = saved["item"]  # same chat, restarted connector: restore its own workbook
        else:
            raise RuntimeError("No workbook open in this chat - call xl_open with a link or file name first "
                               "(or pass workbook=<file name> to this tool)")
    return _state()["books"][_current["item"]]


def _use_workbook(name):
    """Switch this chat to a workbook by name, id or link. Known workbooks switch instantly; others are opened.
    Returns None on success or a message explaining why it could not switch."""
    s = _state()
    n = name.strip().lower()
    if not n.startswith("http"):
        exact = [k for k, b in s["books"].items() if b["name"].lower() in (n, n + ".xlsx") or k.lower() == n]
        part = [k for k, b in s["books"].items() if n in b["name"].lower()]
        hit = exact or part
        if len(hit) == 1:
            _set_current(hit[0])
            return None
    msg = xl_open(name)  # on success its first line starts with the opened workbook's name
    cur = _current["item"]
    if cur and msg.startswith(_state()["books"][cur]["name"]):
        return None
    return msg


_sessions = {}        # item -> workbook session id
_close_due = {}       # item -> (drive, time after which the session is closed to save Claude's edits)
_wb_lock = threading.RLock()
CLOSE_AFTER = 15      # seconds of no further edits before the session is closed (= edits saved to the file)


def _wb(method, sub, body=None):
    """Workbook API call with a persistent session (recreated if it expired)."""
    b = _book()
    base = f"/drives/{b['drive']}/items/{b['item']}/workbook"
    with _wb_lock:
        for attempt in range(2):
            sid = _sessions.get(b["item"])
            if not sid:
                try:
                    sid = _graph("POST", base + "/createSession", {"persistChanges": True}, timeout=25)["id"]
                    _sessions[b["item"]] = sid
                except GraphError:
                    # session could not be opened (busy co-authored file): run this call without one;
                    # sessionless calls save immediately, just a little slower each
                    return _graph(method, base + sub, body)
            try:
                return _graph(method, base + sub, body, session=sid)
            except GraphError as e:
                if attempt == 0 and ("ession" in str(e) or " 404 " in str(e)):
                    _sessions.pop(b["item"], None)
                    continue
                raise


def _close_session(item, drive):
    with _wb_lock:
        _close_due.pop(item, None)
        sid = _sessions.pop(item, None)
        if sid:
            try:
                _graph("POST", f"/drives/{drive}/items/{item}/workbook/closeSession", {}, session=sid, timeout=20)
            except NetworkDown:
                # connection is down: keep the session and try closing (= saving) again once it is back
                _sessions[item] = sid
                _close_due[item] = (drive, _net["down_until"] + 5)
            except Exception:
                pass


def _closer():
    while True:
        time.sleep(2)
        for item, (drive, due) in list(_close_due.items()):
            if time.time() >= due:
                _close_session(item, drive)


threading.Thread(target=_closer, daemon=True).start()


def _commit():
    """After an edit: close the session (which saves the edits to the file) once CLOSE_AFTER seconds pass with no
    further edits. Back-to-back edits and the reads between them reuse one open session instead of reopening the
    workbook each time; an open session left alone can lose its last edits, so it is always closed in the end."""
    b = _book()
    _snaps.get(b["item"], {})["loaded"] = None  # Claude just changed the file: next snapshot use re-checks it
    _close_due[b["item"]] = (b["drive"], time.time() + CLOSE_AFTER)


def _flush_all():
    for item, (drive, _) in list(_close_due.items()):
        _close_session(item, drive)


atexit.register(_flush_all)


def _commits(fn):
    """Decorator for write tools: always save (close the session) before returning."""
    @functools.wraps(fn)
    def run(*a, **k):
        try:
            return fn(*a, **k)
        finally:
            _commit()
    return run


def _sheet_path(sheet):
    return "/worksheets('" + urllib.parse.quote(sheet.replace("'", "''"), safe="") + "')"


def _split(address, sheet=None):
    """'BS Notes'!B5:C9 -> ('BS Notes', 'B5:C9')."""
    if "!" in address:
        sh, rng = address.rsplit("!", 1)
        sh = sh.strip()
        if sh.startswith("'") and sh.endswith("'"):
            sh = sh[1:-1].replace("''", "'")
        return sh, rng.replace("$", "")
    if not sheet:
        raise ValueError(f"Address '{address}' needs a sheet: use Sheet!A1 form")
    return sheet, address.replace("$", "")


def _areas(address):
    """"'MB5B Costing'!A306:AN306,A309:AN309,Summary!B2" -> [(sheet, rng), ...]; an area without a sheet uses the previous one."""
    parts, cur, quoted = [], "", False
    for ch in address:
        if ch == "'":
            quoted = not quoted
        if ch == "," and not quoted:
            parts.append(cur); cur = ""
        else:
            cur += ch
    parts.append(cur)
    out, sheet = [], None
    for p in (p.strip() for p in parts if p.strip()):
        sheet, rng = _split(p, sheet)
        out.append((sheet, rng))
    return out


def _q(sheet):
    return f"'{sheet.replace(chr(39), chr(39) * 2)}'" if re.search(r"[^A-Za-z0-9_]", sheet) else sheet


def _fmt(v):
    if v is None:
        return ""
    if isinstance(v, float):
        return f"{v:.10g}"
    if isinstance(v, (dt.datetime, dt.date)):
        return v.isoformat()[:10]
    return str(v)


# ---------------------------------------------------------------- snapshot index
_snaps = {}  # item -> {"ctag", "f" (index), "loaded", "modified", "by", cached graphs...}; one per workbook


def _snap_dir():
    d = os.path.join(STORE, "snapshots", _book()["item"])
    os.makedirs(d, exist_ok=True)
    return d


def _formula_of(val):
    if isinstance(val, str) and val.startswith("="):
        return val
    if val is not None and hasattr(val, "text"):  # ArrayFormula
        return val.text if str(val.text).startswith("=") else "=" + str(val.text)
    return None


def _load_index(path):
    wf = openpyxl.load_workbook(path, data_only=False)
    wv = openpyxl.load_workbook(path, data_only=True)
    idx = {}
    for ws in wf.worksheets:
        vs = wv[ws.title]
        cells = {}
        for row in ws.iter_rows():
            for c in row:
                if c.value is None:
                    continue
                f = _formula_of(c.value)
                cells[(c.row, c.column)] = (f, vs.cell(c.row, c.column).value if f else c.value)
        idx[ws.title] = {"cells": cells, "state": ws.sheet_state}
    return idx


SNAP_MAX_AGE = 180  # seconds a snapshot is reused before checking the file again (live files change on every autosave)


def _snapshot(force=False):
    """Fresh-enough index of the whole workbook. In order: reuse the in-memory index if under SNAP_MAX_AGE old
    (and no write since); reuse it or the on-disk copy if the file is unchanged (cTag); otherwise download."""
    b = _book()
    snap = _snaps.setdefault(b["item"], {"ctag": None, "f": None, "loaded": None})
    if not force and snap["loaded"] and time.time() - snap["loaded"] < SNAP_MAX_AGE:
        return snap
    meta = _graph("GET", f"/drives/{b['drive']}/items/{b['item']}")  # $select would drop the downloadUrl
    info = {"modified": meta["lastModifiedDateTime"], "by": meta.get("lastModifiedBy", {}).get("user", {}).get("displayName")}
    if not force and snap["f"] is not None and snap["ctag"] == meta["cTag"]:
        snap.update(loaded=time.time(), **info)
        return snap
    d = _snap_dir()
    latest, latest_meta = os.path.join(d, "latest.xlsx"), os.path.join(d, "latest.json")
    on_disk = json.load(open(latest_meta)) if os.path.exists(latest_meta) and os.path.exists(latest) else {}
    if force or on_disk.get("ctag") != meta["cTag"]:
        data = _download(meta["@microsoft.graph.downloadUrl"])
        with open(latest + ".tmp", "wb") as f:
            f.write(data)
        os.replace(latest + ".tmp", latest)
        json.dump({"ctag": meta["cTag"]}, open(latest_meta, "w"))
    if not os.path.exists(os.path.join(d, "baseline.xlsx")):
        shutil.copy(latest, os.path.join(d, "baseline.xlsx"))
    snap.clear()  # drop graphs cached for the previous version of the file
    snap.update(ctag=meta["cTag"], f=_load_index(latest), names=_count_names(latest), loaded=time.time(), **info)
    return snap


def _count_names(path):
    """(workbook-level named ranges, how many point at #REF!) - thousands of stale names bloat a file."""
    try:
        wb = openpyxl.load_workbook(path, read_only=True, keep_links=False)
        names = list(wb.defined_names.values())
        broken = sum(1 for n in names if "#REF" in str(getattr(n, "attr_text", "")))
        wb.close()
        return len(names), broken
    except Exception:
        return 0, 0


def _sheet_cells(sheet):
    idx = _snapshot()["f"]
    if sheet not in idx:
        near = [s for s in idx if s.strip().lower() == sheet.strip().lower()]
        if near:
            sheet = near[0]
        else:
            raise ValueError(f"No sheet '{sheet}'. Sheets: {', '.join(idx)}")
    return sheet, idx[sheet]["cells"]


# ---------------------------------------------------------------- formula reference parsing
_CELL = re.compile(r"^\$?([A-Za-z]{1,3})\$?(\d+)$")
_COL = re.compile(r"^\$?([A-Za-z]{1,3})$")
_ROW = re.compile(r"^\$?(\d+)$")


def _bounds(ref):
    """'A1' / 'A1:B5' / 'A:A' / '3:5' -> (r1, c1, r2, c2) or None (named range etc.)."""
    parts = ref.replace("$", "").split(":")
    if len(parts) == 1:
        m = _CELL.match(parts[0])
        if not m:
            return None
        c, r = column_index_from_string(m[1].upper()), int(m[2])
        return r, c, r, c
    if len(parts) != 2:
        return None
    a, b = parts
    if _CELL.match(a) and _CELL.match(b):
        c1, r1, c2, r2 = range_boundaries(f"{a}:{b}")
        return min(r1, r2), min(c1, c2), max(r1, r2), max(c1, c2)
    if _COL.match(a) and _COL.match(b):
        c1, c2 = column_index_from_string(a.upper()), column_index_from_string(b.upper())
        return 1, min(c1, c2), 1048576, max(c1, c2)
    if _ROW.match(a) and _ROW.match(b):
        return min(int(a), int(b)), 1, max(int(a), int(b)), 16384
    return None


def _refs(formula, own_sheet):
    """[(sheet, ref_text, bounds|None)] for every range operand in a formula."""
    out = []
    try:
        toks = Tokenizer(formula).items
    except Exception:
        return out
    for t in toks:
        if t.type == Token.OPERAND and t.subtype == Token.RANGE:
            v = t.value
            if "!" in v:
                sh, ref = v.rsplit("!", 1)
                sh = sh.strip("'").replace("''", "'")
            else:
                sh, ref = own_sheet, v
            out.append((sh, ref, _bounds(ref)))
    return out


def _r1c1(formula, row, col):
    """Relative signature of a formula so copies down/across group together."""
    try:
        toks = Tokenizer(formula).items
    except Exception:
        return formula
    parts = []
    for t in toks:
        v = t.value
        if t.type == Token.OPERAND and t.subtype == Token.RANGE:
            pre, ref = (v.rsplit("!", 1)[0] + "!", v.rsplit("!", 1)[1]) if "!" in v else ("", v)

            def conv(m):
                cabs, cl, rabs, rw = m.group(1), m.group(2), m.group(3), m.group(4)
                c, r = column_index_from_string(cl.upper()), int(rw)
                rs = f"R{r}" if rabs else ("R" if r == row else f"R[{r - row}]")
                cs = f"C{c}" if cabs else ("C" if c == col else f"C[{c - col}]")
                return rs + cs
            v = pre + re.sub(r"(\$?)([A-Za-z]{1,3})(\$?)(\d+)", conv, ref)
        parts.append(v)
    return "=" + "".join(parts)


def _clip(text):
    return text if len(text) <= MAX_OUT else text[:MAX_OUT] + f"\n... [truncated {len(text) - MAX_OUT} chars - narrow the range]"


# ---------------------------------------------------------------- write log / undo
_LOG = os.path.join(STORE, "writes.jsonl")


_CALLS = os.path.join(STORE, "calls.jsonl")
_PROC = secrets.token_hex(4)  # fallback chat tag when no chat id is found


def _sid():
    """Tag for this chat in the logs: the chat id (stable across restarts) or, failing that, this process."""
    return _chat_id() or _PROC


def _log(entry):
    entry["sid"] = _sid()
    with open(_LOG, "a", encoding="utf-8") as f:
        f.write(json.dumps(entry) + "\n")


def _log_call(tool, ms, ok):
    try:
        with open(_CALLS, "a", encoding="utf-8") as f:
            f.write(json.dumps({"ts": dt.datetime.now().isoformat(timespec="seconds"), "tool": tool, "ms": ms,
                                "ok": ok, "sid": _chat["id"] or _PROC, "item": _current["item"]}) + "\n")
    except Exception:
        pass


def _norm(x):
    """Formula text for comparisons: '=+A1' and '=A1' are the same formula."""
    s = str(x)
    return "=" + s[2:] if s.startswith("=+") else s


def _live(sheet, rng, select="address,values,formulas"):
    return _wb("GET", f"{_sheet_path(sheet)}/range(address='{rng}')?$select={select}")


def _flat(m):
    return [_norm(x) for row in m for x in row]


def _same_grid(a, b):
    """Grids equal? Formulas as text ('=+A1' == '=A1'); constants as numbers within 1e-6 relative (Graph returns full
    precision, e.g. 93716470.3500001, while the agent read 93716470.35)."""
    fa, fb = _flat(a), _flat(b)
    if len(fa) != len(fb):
        return False
    for x, y in zip(fa, fb):
        if x == y:
            continue
        try:
            if abs(float(x) - float(y)) <= 1e-6 * max(1.0, abs(float(x))):
                continue
        except ValueError:
            pass
        return False
    return True


def _shape(rng):
    c1, r1, c2, r2 = range_boundaries(rng)
    return r2 - r1 + 1, c2 - c1 + 1


def _do_write(sheet, rng, formulas, expect, note, cur=None):
    cur = cur or _live(sheet, rng)
    if expect is not None and not _same_grid(cur["formulas"], expect):
        return f"CONFLICT {_q(sheet)}!{rng}: now {cur['formulas']} (expected {expect}) - not written"
    res = _wb("PATCH", f"{_sheet_path(sheet)}/range(address='{rng}')", {"formulas": formulas})
    _log({"ts": dt.datetime.now().isoformat(timespec="seconds"), "item": _book()["item"], "sheet": sheet,
          "range": rng, "old": cur["formulas"], "new": formulas, "note": note})
    vals = res.get("values", [])
    preview = "; ".join(_fmt(v) for v in _flat(vals)[:6]) + (" ..." if len(_flat(vals)) > 6 else "")
    return f"OK {_q(sheet)}!{rng} -> {preview}"


# ================================================================ minimal MCP (stdio JSON-RPC)
_JSON_TYPES = {str: "string", int: "integer", float: "number", bool: "boolean", dict: "object", list: "array"}


def _schema(tp):
    if typing.get_origin(tp) in (typing.Union, getattr(__import__("types"), "UnionType", ())):
        tp = next(a for a in typing.get_args(tp) if a is not type(None))
    origin = typing.get_origin(tp)
    if origin is list:
        args = typing.get_args(tp)
        return {"type": "array", "items": _schema(args[0]) if args else {}}
    return {"type": _JSON_TYPES.get(origin or tp, "string")}


_NO_WORKBOOK_PARAM = {"xl_status", "xl_open"}


class MiniMCP:
    def __init__(self, name, instructions):
        self.name, self.instructions, self.tools = name, instructions, {}

    def tool(self):
        def deco(fn):
            hints = typing.get_type_hints(fn)
            props, required = {}, []
            for p in inspect.signature(fn).parameters.values():
                props[p.name] = _schema(hints.get(p.name, str))
                if p.default is inspect.Parameter.empty:
                    required.append(p.name)
                elif p.default not in ("", False, None, -1):  # the required list already says what is optional
                    props[p.name]["default"] = p.default
            if fn.__name__ not in _NO_WORKBOOK_PARAM:
                props["workbook"] = {"type": "string"}  # explained once in the server instructions
            self.tools[fn.__name__] = (fn, {"name": fn.__name__, "description": inspect.cleandoc(fn.__doc__ or ""),
                                            "inputSchema": {"type": "object", "properties": props, "required": required}})
            return fn
        return deco

    def call(self, name, args):
        fn = self.tools[name][0]
        args = dict(args)
        wb = args.pop("workbook", None)
        t0, ok = time.time(), False
        try:
            if wb:
                problem = _use_workbook(wb)
                if problem:
                    return f"Could not switch to workbook '{wb}': {problem}", True
            out, ok = str(fn(**args)), True
            return out, False
        except NotSignedIn as e:
            return _start_sign_in(str(e)), True
        except Exception as e:
            return f"{type(e).__name__}: {e}", True
        finally:
            _log_call(name, int((time.time() - t0) * 1000), ok)

    def _handle(self, msg):
        method, mid, params = msg.get("method"), msg.get("id"), msg.get("params") or {}
        if mid is None:
            return None  # notification
        if method == "initialize":
            res = {"protocolVersion": params.get("protocolVersion", "2025-06-18"), "capabilities": {"tools": {}},
                   "serverInfo": {"name": self.name, "version": VERSION}, "instructions": self.instructions}
        elif method == "ping":
            res = {}
        elif method == "tools/list":
            res = {"tools": [t[1] for t in self.tools.values()]}
        elif method == "tools/call":
            if params.get("name") not in self.tools:
                return {"jsonrpc": "2.0", "id": mid, "error": {"code": -32602, "message": f"unknown tool {params.get('name')}"}}
            text, err = self.call(params["name"], params.get("arguments") or {})
            res = {"content": [{"type": "text", "text": text}], "isError": err}
        else:
            return {"jsonrpc": "2.0", "id": mid, "error": {"code": -32601, "message": f"method not found: {method}"}}
        return {"jsonrpc": "2.0", "id": mid, "result": res}

    def run(self, transport="stdio"):
        out = sys.stdout.buffer
        for line in sys.stdin.buffer:
            if not line.strip():
                continue
            try:
                reply = self._handle(json.loads(line))
            except Exception:
                reply = {"jsonrpc": "2.0", "id": None, "error": {"code": -32603, "message": traceback.format_exc()[-800:]}}
            if reply is not None:
                out.write(json.dumps(reply).encode("utf-8") + b"\n")
                out.flush()
        _flush_all()  # chat closed: save any edits still waiting in an open session


# ---------------------------------------------------------------- sign-in
_login = {"status": "not started", "url": None}
_AUTHORIZE = f"https://login.microsoftonline.com/{TENANT}/oauth2/v2.0/authorize"
_TOKEN_URL = f"https://login.microsoftonline.com/{TENANT}/oauth2/v2.0/token"


def _got_tokens(t):
    _save_refresh(t["refresh_token"])
    _tok.update(access=t["access_token"], exp=time.time() + t["expires_in"] - 120)
    _login["status"] = "signed in"


def _browser_login():
    """Authorization-code + PKCE flow with a one-shot localhost listener. Returns the sign-in URL."""
    verifier = secrets.token_urlsafe(64)
    challenge = base64.urlsafe_b64encode(hashlib.sha256(verifier.encode()).digest()).decode().rstrip("=")
    state = secrets.token_urlsafe(16)
    got = {}

    class Handler(http.server.BaseHTTPRequestHandler):
        def do_GET(self):
            q = urllib.parse.parse_qs(urllib.parse.urlparse(self.path).query)
            if q.get("state", [None])[0] == state:
                got.update({k: v[0] for k, v in q.items()})
            ok = "code" in got
            self.send_response(200)
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self.end_headers()
            self.wfile.write(("<html><body style='font-family:Segoe UI;margin:3em'><h2>"
                              + ("Signed in to Excel 365 connector. You can close this tab." if ok else "Sign-in did not complete: "
                                 + got.get("error_description", "unknown error"))
                              + "</h2></body></html>").encode())

        def log_message(self, *a):
            pass

    srv = http.server.HTTPServer(("127.0.0.1", 0), Handler)
    redirect = f"http://localhost:{srv.server_address[1]}"
    url = _AUTHORIZE + "?" + urllib.parse.urlencode({
        "client_id": CLIENT_ID, "response_type": "code", "redirect_uri": redirect, "scope": SCOPE,
        "code_challenge": challenge, "code_challenge_method": "S256", "state": state, "prompt": "select_account"})

    def wait():
        srv.timeout = 5
        end = time.time() + 600
        while time.time() < end and not got:
            srv.handle_request()
        srv.server_close()
        if "code" not in got:
            _login["status"] = f"failed: {got.get('error', 'timed out')}"
            return
        t = _post_form(_TOKEN_URL, {"grant_type": "authorization_code", "client_id": CLIENT_ID, "code": got["code"],
                                    "redirect_uri": redirect, "code_verifier": verifier, "scope": SCOPE})
        if "access_token" in t:
            _got_tokens(t)
        else:
            _login["status"] = f"failed: {t.get('error_description', t.get('error'))}"[:300]

    _login.update(status="waiting for user (browser)", url=url, started=time.time())
    threading.Thread(target=wait, daemon=True).start()
    webbrowser.open(url)
    return url


def _start_sign_in(reason):
    """Called when any tool needs a token and there is none: opens the browser sign-in once."""
    if not (_login["status"].startswith("waiting") and time.time() - _login.get("started", 0) < 600):
        _browser_login()
    return (f"{reason}. A Microsoft sign-in page has opened in the user's browser. Ask them to sign in with their "
            f"Microsoft 365 work account, then retry. If no browser opened, they can use this link: "
            f"{_login['url']}")


# ================================================================ MCP server
mcp = MiniMCP(
    name="excel365",
    instructions=(
        "Live Excel 365 (OneDrive/SharePoint) workbooks that a team co-edits. Per chat: xl_open once (remembered across "
        "restarts; workbook= on any tool switches). Unfamiliar sheet: xl_map before bulk xl_read. Edit with xl_write/xl_fill "
        "and expect= (a teammate's newer edit is never overwritten; batches are all-or-nothing; watch= returns recalculated "
        "cells in the same call). xl_check = error origins + circular references (suspect a loop when totals don't reconcile). "
        "Addresses: Sheet!A1; quote names with spaces: 'BS Notes'!C5. Date-like text needs as_text. map/find/trace/check use "
        "a snapshot up to ~2 min behind live. Never download-edit-upload by hand: xl_file replace is the guarded way. "
        "Before inserting/deleting rows ask whether anyone has the file open in desktop Excel (save with AutoSave on, or "
        "close). On 'connection not responding': stop, tell the user to check internet/VPN. On 'sign-in page opened': the "
        "user signs in in their browser, then retry; never ask for a password."),
)


def _do_login(method):
    if method == "browser":
        url = _browser_login()
        return f"Microsoft sign-in opened in the user's browser. If it did not appear, they can open: {url}"
    r = _post_form(f"https://login.microsoftonline.com/{TENANT}/oauth2/v2.0/devicecode", {"client_id": CLIENT_ID, "scope": SCOPE})
    if "device_code" not in r:
        return f"Failed: {r}"

    def poll():
        end = time.time() + r["expires_in"]
        while time.time() < end:
            time.sleep(r["interval"])
            t = _post_form(f"https://login.microsoftonline.com/{TENANT}/oauth2/v2.0/token", {
                "grant_type": "urn:ietf:params:oauth:grant-type:device_code", "client_id": CLIENT_ID, "device_code": r["device_code"]})
            if "access_token" in t:
                _got_tokens(t)
                return
            if t.get("error") not in ("authorization_pending", "slow_down"):
                _login["status"] = f"failed: {t.get('error')}"
                return
        _login["status"] = "failed: code expired"
    _login["status"] = "waiting for user"
    threading.Thread(target=poll, daemon=True).start()
    return (f"Ask the user to open https://login.microsoftonline.com/{TENANT if TENANT != 'organizations' else 'common'}/oauth2/deviceauth "
            f"(or https://microsoft.com/devicelogin), enter code {r['user_code']} and sign in with their work account.")


@mcp.tool()
def xl_status(login: str = "") -> str:
    """Sign-in state, this chat's workbook, version. login='browser' opens Microsoft's sign-in page (normally automatic);
    'code' = device code when no browser can open here."""
    if login:
        return _do_login(login)
    try:
        _token()
        auth = "signed in"
    except Exception as e:
        auth = f"{_login['status']} ({e}) - any tool call opens the sign-in page"
    s = _state()
    try:
        cur_id = _book()["item"]  # restores this chat's own workbook after a connector restart
    except Exception:
        cur_id = None
    cur = s["books"].get(cur_id, {}).get("name") if cur_id else "none - call xl_open first"
    others = [b["name"] for k, b in s["books"].items() if k != cur_id]
    chat = _chat_id()
    return (f"auth: {auth}\nworkbook open in this chat: {cur}\nother known workbooks: {', '.join(others) or '-'}\n"
            f"chat id: {chat or 'not found - the workbook is remembered only until this connector restarts'}  version {VERSION}")


def _search_books(text):
    """Excel files the user can access (own + shared) whose name matches text."""
    q = re.sub(r"\.xls[xmb]?$", "", text, flags=re.I).replace('"', "")
    r = _graph("POST", "/search/query", {"requests": [{"entityTypes": ["driveItem"], "size": 15,
                                                        "query": {"queryString": f'"{q}" AND (filetype:xlsx OR filetype:xlsm)'}}]})
    hits = []
    for c in r.get("value", [{}])[0].get("hitsContainers", []):
        for h in c.get("hits", []):
            res = h["resource"]
            hits.append({"name": res["name"], "drive": res["parentReference"]["driveId"], "item": res["id"],
                         "modified": res.get("lastModifiedDateTime", "")[:16], "url": res.get("webUrl", "")})
    return hits


def _open_item(target):
    """Make target this chat's workbook (and remember it); returns its file metadata."""
    s = _state()
    meta = _graph("GET", f"/drives/{target['drive']}/items/{target['item']}?$select=name,size,lastModifiedDateTime,lastModifiedBy")
    target["name"] = meta["name"]  # files get renamed; keep the current name
    s["books"][target["item"]] = target
    s["current"] = target["item"]
    _save_state(s)
    _set_current(target["item"])
    return meta


@mcp.tool()
def xl_open(link_or_name: str, details: bool = False) -> str:
    """Open a workbook by link or file name (own and shared files; a partial name is enough). Lists sheets.
    details=True adds used range and formula count per sheet (downloads a snapshot)."""
    s = _state()
    target = None
    if link_or_name.startswith("http"):
        enc = "u!" + base64.urlsafe_b64encode(link_or_name.encode()).decode().rstrip("=")
        try:
            it = _graph("GET", f"/shares/{enc}/driveItem?$select=id,name,parentReference")
            target = {"name": it["name"], "drive": it["parentReference"]["driveId"], "item": it["id"], "link": link_or_name}
        except GraphError as e:
            # dead/changed sharing link: fall back to the file name in the URL, if any
            fname = urllib.parse.parse_qs(urllib.parse.urlparse(link_or_name).query).get("file", [None])[0]
            if not fname:
                return (f"That link could not be opened ({str(e)[:120]}). It may have been removed or changed. "
                        "Ask for a fresh link, or call xl_open with the file name instead.")
            link_or_name = fname
    if target is None:
        known = [b for b in s["books"].values() if link_or_name.lower() in b["name"].lower()]
        if len(known) == 1:
            target = known[0]
        else:
            # exact file in the user's own OneDrive root (works for brand-new files search hasn't indexed yet)
            try:
                it = _graph("GET", "/me/drive/root:/" + urllib.parse.quote(link_or_name) + ":?$select=id,name,parentReference")
                if it["name"].lower().endswith((".xlsx", ".xlsm")):
                    target = {"name": it["name"], "drive": it["parentReference"]["driveId"], "item": it["id"]}
            except GraphError:
                pass
        if target is None and len(known) != 1:
            hits = _search_books(link_or_name)
            exact = [h for h in hits if h["name"].lower() == link_or_name.lower()]
            if len(exact) == 1 or len(hits) == 1:
                target = (exact or hits)[0]
            elif not hits:
                return f"No Excel file matching '{link_or_name}' found among files this account can access."
            else:
                return "Several files match - call xl_open again with the exact name:\n" + "\n".join(
                    f"  {h['name']}  (modified {h['modified']})  {h['url']}" for h in hits)
    meta = _open_item(target)
    b = _book()
    sheets = _wb("GET", "/worksheets?$select=name,position,visibility")["value"]
    lines = [f"{b['name']}  ({meta['size'] // 1024} KB, last saved {meta['lastModifiedDateTime']} by "
             f"{meta.get('lastModifiedBy', {}).get('user', {}).get('displayName')})"]
    idx = _snapshot()["f"] if details else {}
    for sh in sheets:
        line = f"{sh['position']:>3} {sh['name']!r:<34} {sh['visibility']}"
        if details:
            cells = idx.get(sh["name"], {}).get("cells", {})
            nf = sum(1 for f, _ in cells.values() if f)
            used = f"A1:{get_column_letter(max(c for _, c in cells))}{max(r for r, _ in cells)}" if cells else "-"
            line = f"{line:<52} used {used:<10} {nf} formulas"
        lines.append(line)
    return "\n".join(lines)


@mcp.tool()
def xl_read(address: str, mode: str = "both") -> str:
    """Live cells, several areas per call: "'BS Notes'!B5:H40,Summary!E20". mode: both | formulas | values |
    text (as displayed) | format (number formats). Blank cells skipped; lines 'A1 | formula | value'."""
    blocks = []
    for sheet, rng in _areas(address):
        c1, r1, c2, r2 = range_boundaries(rng)
        out, step = [], max(1, 8000 // (c2 - c1 + 1))
        for rs in range(r1, r2 + 1, step):
            re_ = min(r2, rs + step - 1)
            sub = f"{get_column_letter(c1)}{rs}:{get_column_letter(c2)}{re_}"
            if mode in ("text", "format"):
                d = _live(sheet, sub, select="address,values,text,numberFormat")
                grid = d["text"] if mode == "text" else d["numberFormat"]
                for i, row in enumerate(grid):
                    for j, t in enumerate(row):
                        if d["values"][i][j] == "" and (mode == "text" or t == "General"):
                            continue
                        out.append(f"{get_column_letter(c1 + j)}{rs + i} | {t}")
                continue
            d = _live(sheet, sub)
            for i, (frow, vrow) in enumerate(zip(d["formulas"], d["values"])):
                for j, (f, v) in enumerate(zip(frow, vrow)):
                    if f == "" and v == "":
                        continue
                    a = f"{get_column_letter(c1 + j)}{rs + i}"
                    is_f = isinstance(f, str) and f.startswith("=")
                    if mode == "formulas":
                        out.append(f"{a} | {f}")
                    elif mode == "values":
                        out.append(f"{a} | {_fmt(v)}")
                    else:
                        out.append(f"{a} | {f} | {_fmt(v)}" if is_f else f"{a} | {_fmt(v)}")
        blocks.append(f"{_q(sheet)}!{rng} (live)\n" + ("\n".join(out) or "(all blank)"))
    return _clip("\n".join(blocks))


_ORDINARY_NUMBERS = {0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 12, 24, 30, 31, 60, 100, 365, 366, 1000, 10 ** 5, 10 ** 6, 10 ** 7, 0.5}


def _typed_numbers(formula):
    """Numbers typed into a formula (e.g. =ROUND(840332/10^5,2)): usually amounts keyed in instead of linked."""
    try:
        toks = Tokenizer(formula).items
    except Exception:
        return []
    out = []
    for t in toks:
        if t.type == Token.OPERAND and t.subtype == Token.NUMBER:
            try:
                if float(t.value) not in _ORDINARY_NUMBERS:
                    out.append(t.value)
            except ValueError:
                pass
    return out


@mcp.tool()
def xl_map(sheet: str, address: str = "", labels: bool = True) -> str:
    """Snapshot structure of a sheet: formula blocks collapsed by pattern, hardcoded numbers, labels, typed-in amounts
    flagged. Use before bulk reads of an unfamiliar sheet. address limits it, e.g. A1:H80."""
    sheet, cells = _sheet_cells(sheet)
    if address:
        c1, r1, c2, r2 = range_boundaries(_split(address, sheet)[1])
        cells = {k: v for k, v in cells.items() if r1 <= k[0] <= r2 and c1 <= k[1] <= c2}
    # runs down each column
    keyed = {}
    for (r, c), (f, v) in cells.items():
        if f:
            key = ("f", _r1c1(f, r, c))
        elif isinstance(v, bool) or not isinstance(v, (int, float)):
            key = ("t", r, c)  # every label stands alone
        else:
            key = ("n",)
        keyed[(r, c)] = key
    runs = []
    for c in sorted({c for _, c in keyed}):
        rows = sorted(r for r, cc in keyed if cc == c)
        start = prev = rows[0]
        for r in rows[1:] + [None]:
            if r is not None and r == prev + 1 and keyed[(r, c)] == keyed[(start, c)]:
                prev = r
                continue
            runs.append((start, prev, c, keyed[(start, c)]))
            if r is not None:
                start = prev = r
    # merge identical runs in adjacent columns into blocks
    by = {}
    for r1_, r2_, c, k in runs:
        by.setdefault((r1_, r2_, k if k[0] != "t" else ("t",)), []).append(c)
    blocks = []
    for (r1_, r2_, k), cols in by.items():
        cols.sort()
        s = p = cols[0]
        for c in cols[1:] + [None]:
            if c is not None and c == p + 1 and k[0] != "t":
                p = c
                continue
            blocks.append((r1_, s, r2_, p, k))
            if c is not None:
                s = p = c
    blocks.sort()
    out, nf, nn, typed = [], 0, 0, 0
    for r1_, c1_, r2_, c2_, k in blocks:
        a = f"{get_column_letter(c1_)}{r1_}"
        rng = a if (r1_, c1_) == (r2_, c2_) else f"{a}:{get_column_letter(c2_)}{r2_}"
        n = (r2_ - r1_ + 1) * (c2_ - c1_ + 1)
        f, v = cells[(r1_, c1_)]
        if k[0] == "f":
            nf += n
            nums = _typed_numbers(f)
            typed += bool(nums)
            out.append(f"{rng:<12} {f}" + (f"   -> {_fmt(v)}" if n == 1 else f"   [x{n} same pattern, first -> {_fmt(v)}]")
                       + (f"   [typed-in numbers: {', '.join(nums[:4])}]" if nums else ""))
        elif k[0] == "n":
            nn += n
            out.append(f"{rng:<12} # {_fmt(v)}" if n == 1 else f"{rng:<12} # hardcoded x{n} (first {_fmt(v)})")
        elif labels:
            out.append(f"{rng:<12} '{str(v)[:60]}'")
    hdr = (f"{_q(sheet)} map  ({nf} formula cells, {nn} hardcoded numbers, {typed} formula blocks with typed-in numbers; "
           f"snapshot {_snapshot().get('modified')})")
    return _clip(hdr + "\n" + "\n".join(out))


def _visible(idx, sh):
    return idx.get(sh, {}).get("state", "visible") == "visible"


def _sheet_filter(idx, sheets, sheets_exclude, visible_only):
    only = {s.strip().lower() for s in sheets.split(",") if s.strip()}
    skip = {s.strip().lower() for s in sheets_exclude.split(",") if s.strip()}
    return [sh for sh in idx if (not only or sh.lower() in only) and sh.lower() not in skip
            and (not visible_only or _visible(idx, sh))]


@mcp.tool()
def xl_find(text: str, regex: bool = False, where: str = "both", sheets: str = "", limit: int = 300,
            sheets_exclude: str = "", visible_only: bool = False, group_by: str = "") -> str:
    """Search formulas/values (snapshot). where: formulas | values | both. sheets / sheets_exclude: comma lists.
    group_by='sheet' -> counts per sheet. Always reports the total match count."""
    idx = _snapshot()["f"]
    pat = re.compile(text if regex else re.escape(text), re.I)
    out, total, per = [], 0, {}
    for sh in _sheet_filter(idx, sheets, sheets_exclude, visible_only):
        for (r, c), (f, v) in sorted(idx[sh]["cells"].items()):
            hay = []
            if f and where in ("formulas", "both"):
                hay.append(f)
            if where in ("values", "both"):
                hay.append(_fmt(v))
            if any(pat.search(h) for h in hay):
                total += 1
                per[sh] = per.get(sh, 0) + 1
                if len(out) < limit and group_by != "sheet":
                    out.append(f"{_q(sh)}!{get_column_letter(c)}{r} | {f or ''} | {_fmt(v)}")
    if group_by == "sheet":
        return "\n".join(f"{n:>6}  {_q(sh)}{'' if _visible(idx, sh) else '  (hidden)'}" for sh, n in
                         sorted(per.items(), key=lambda kv: -kv[1])) + f"\ntotal {total} matches" if total else "no matches"
    tail = f"\n... showing {len(out)} of {total} matches (raise limit, narrow sheets, or use group_by='sheet')" if total > len(out) else ""
    return _clip(("\n".join(out) or "no matches") + tail)


# ---------------------------------------------------------------- formula analysis helpers
def _refs_index(snap):
    """Parsed references of every formula, cached on the snapshot: {(sheet, r, c): [(ref_sheet, bounds), ...]}."""
    if "refs" not in snap:
        snap["refs"] = {(sh, rr, cc): [(rsh, bd) for rsh, _, bd in _refs(f, sh) if bd]
                        for sh, d in snap["f"].items() for (rr, cc), (f, _) in d["cells"].items() if f}
    return snap["refs"]


def _calls(formula):
    """Function calls in a formula: [(NAME, [arg tokens, ...]), ...] (inner calls first)."""
    try:
        toks = Tokenizer(formula).items
    except Exception:
        return []
    stack, done = [], []
    for t in toks:
        if t.type == Token.FUNC and t.subtype == Token.OPEN:
            stack.append((t.value[:-1].upper().replace("_XLFN.", "").replace("@", ""), [[]]))
        elif t.type == Token.FUNC and t.subtype == Token.CLOSE and stack:
            done.append(stack.pop())
            if stack:
                stack[-1][1][-1].append(None)  # nested call result
        elif t.type == Token.SEP and t.subtype == Token.ARG and stack:
            stack[-1][1].append([])
        elif stack and t.type != Token.WSPACE:
            stack[-1][1][-1].append(t)
    return done


def _arg(tokens, own_sheet, idx):
    """('range', sheet, bounds) | ('value', v) | None for an argument made of one operand."""
    toks = [t for t in tokens if t is not None]
    if len(toks) != 1 or None in tokens or toks[0].type != Token.OPERAND:
        return None
    t = toks[0]
    if t.subtype == Token.RANGE:
        sh, ref = (t.value.rsplit("!", 1)[0].strip("'").replace("''", "'"), t.value.rsplit("!", 1)[1]) if "!" in t.value else (own_sheet, t.value)
        bd = _bounds(ref)
        if not bd:
            return None
        if bd[0] == bd[2] and bd[1] == bd[3]:
            return ("value", idx.get(sh, {}).get("cells", {}).get((bd[0], bd[1]), (None, None))[1], sh, bd)
        return ("range", sh, bd)
    if t.subtype == Token.TEXT:
        return ("value", t.value[1:-1].replace('""', '"'))
    if t.subtype == Token.NUMBER:
        return ("value", float(t.value))
    return None


def _eq(a, b):
    if a is None or b is None:
        return False
    try:
        return abs(float(a) - float(b)) < 1e-9
    except (TypeError, ValueError):
        return str(a).strip().lower() == str(b).strip().lower()


def _col_cells(idx, sh, bd):
    """{row: value} for the first column of a range, from the snapshot."""
    return {r: v for (r, c), (_, v) in idx.get(sh, {}).get("cells", {}).items() if c == bd[1] and bd[0] <= r <= bd[2]}


def _resolve_lookups(formula, own_sheet, idx, max_rows=12):
    """Which rows SUMIFS / SUMIF / COUNTIFS / XLOOKUP / VLOOKUP actually pick up, from snapshot values."""
    out = []
    for name, args in _calls(formula):
        a = [_arg(x, own_sheet, idx) for x in args]
        try:
            if name in ("SUMIFS", "COUNTIFS", "AVERAGEIFS", "SUMIF", "COUNTIF"):
                if name == "SUMIF":
                    sum_r = a[2] if len(a) > 2 and a[2] else a[0]
                    pairs = [(a[0], a[1])]
                elif name in ("COUNTIFS", "COUNTIF"):
                    sum_r, pairs = None, [(a[i], a[i + 1]) for i in range(0, len(a) - 1, 2)]
                else:
                    sum_r, pairs = a[0], [(a[i], a[i + 1]) for i in range(1, len(a) - 1, 2)]
                if any(not p or not q or p[0] != "range" or q[0] != "value" or q[1] is None for p, q in pairs):
                    continue
                if any(isinstance(q[1], str) and (q[1][:1] in "<>=" or "*" in q[1] or "?" in q[1]) for _, q in pairs):
                    out.append(f"{name}: criteria use operators/wildcards - not resolved")
                    continue
                rows = None
                for p, q in pairs:
                    match = {r for r, v in _col_cells(idx, p[1], p[2]).items() if _eq(v, q[1])}
                    rows = match if rows is None else rows & {r - p[2][0] + pairs[0][0][2][0] for r in match}
                rows = sorted(rows or [])
                crit = ", ".join(f"{_fmt(q[1])!r}" for _, q in pairs)
                if sum_r and sum_r[0] == "range":
                    sh, bd = sum_r[1], sum_r[2]
                    off = bd[0] - pairs[0][0][2][0]
                    cells = idx.get(sh, {}).get("cells", {})
                    vals = [(r + off, cells.get((r + off, bd[1]), (None, None))[1]) for r in rows]
                    total = sum(v for _, v in vals if isinstance(v, (int, float)))
                    out.append(f"{name}({crit}) -> {len(rows)} row(s), total {_fmt(total)}:")
                    out += [f"   {_q(sh)}!{get_column_letter(bd[1])}{r} = {_fmt(v)}" for r, v in vals[:max_rows]]
                    if len(vals) > max_rows:
                        out.append(f"   ... {len(vals) - max_rows} more rows")
                else:
                    out.append(f"{name}({crit}) -> {len(rows)} matching row(s): {rows[:max_rows]}")
            elif name in ("XLOOKUP", "VLOOKUP") and len(a) >= 3 and a[0] and a[0][0] == "value" and a[1] and a[1][0] == "range":
                key, (_, sh, bd) = a[0][1], a[1]
                hit = next((r for r, v in sorted(_col_cells(idx, sh, bd).items()) if _eq(v, key)), None)
                if hit is None:
                    out.append(f"{name}({_fmt(key)!r}) -> NOT FOUND in {_q(sh)}!{get_column_letter(bd[1])}:{get_column_letter(bd[1])}")
                    continue
                if name == "XLOOKUP" and a[2] and a[2][0] == "range":
                    rsh, rbd = a[2][1], a[2][2]
                    rr, rc = hit - bd[0] + rbd[0], rbd[1]
                elif name == "VLOOKUP" and a[2] and a[2][0] == "value":
                    rsh, rr, rc = sh, hit, bd[1] + int(float(a[2][1])) - 1
                else:
                    continue
                v = idx.get(rsh, {}).get("cells", {}).get((rr, rc), (None, None))
                out.append(f"{name}({_fmt(key)!r}) -> matched row {hit}: {_q(rsh)}!{get_column_letter(rc)}{rr} | {v[0] or ''} | {_fmt(v[1])}")
        except Exception:
            continue
    return out


@mcp.tool()
def xl_trace(address: str, direction: str = "precedents", depth: int = 2) -> str:
    """precedents: what a cell is built from (SUMIFS/SUMIF/COUNTIFS/XLOOKUP/VLOOKUP over columns resolved to the
    matching rows and values); dependents: formulas that use it."""
    snap = _snapshot()
    idx = snap["f"]
    sheet, rng = _split(address)
    sheet = _sheet_cells(sheet)[0]
    c, r = range_boundaries(rng)[:2]
    out = []
    if direction.startswith("pre"):
        seen = set()

        def walk(sh, rr, cc, lvl):
            if (sh, rr, cc) in seen or lvl > depth:
                return
            seen.add((sh, rr, cc))
            f, v = idx.get(sh, {}).get("cells", {}).get((rr, cc), (None, None))
            out.append(f"{'  ' * lvl}{_q(sh)}!{get_column_letter(cc)}{rr} | {f or ''} | {_fmt(v)}")
            if not f:
                return
            big = False
            for rsh, ref, bd in _refs(f, sh):
                if bd and (bd[2] - bd[0] + 1) * (bd[3] - bd[1] + 1) <= 12:
                    for r_ in range(bd[0], bd[2] + 1):
                        for c_ in range(bd[1], bd[3] + 1):
                            walk(rsh, r_, c_, lvl + 1)
                else:
                    big = True
            if big:
                resolved = _resolve_lookups(f, sh, idx)
                out.extend(f"{'  ' * (lvl + 1)}{line}" for line in resolved or
                           [f"{_q(rsh)}!{ref}  (large range)" for rsh, ref, bd in _refs(f, sh)
                            if not bd or (bd[2] - bd[0] + 1) * (bd[3] - bd[1] + 1) > 12])
        walk(sheet, r, c, 0)
    else:
        refs_idx = _refs_index(snap)
        frontier, seen = [(sheet, r, c)], {(sheet, r, c)}
        for lvl in range(1, depth + 1):
            nxt = []
            for (sh, rr, cc), refs in refs_idx.items():
                if (sh, rr, cc) in seen:
                    continue
                f, v = idx[sh]["cells"][(rr, cc)]
                for rsh, bd in refs:
                    if any(rsh == ts and bd[0] <= tr <= bd[2] and bd[1] <= tc <= bd[3] for ts, tr, tc in frontier):
                        seen.add((sh, rr, cc))
                        nxt.append((sh, rr, cc))
                        out.append(f"{'  ' * lvl}{_q(sh)}!{get_column_letter(cc)}{rr} | {f} | {_fmt(v)}")
                        break
                if len(out) > 400:
                    return _clip("\n".join(out) + "\n... stopped at 400 dependents")
            frontier = nxt
            if not frontier:
                break
        out.insert(0, f"dependents of {_q(sheet)}!{rng}:" if out else "no dependents found")
    return _clip("\n".join(out))


def _as_text(grid):
    """Keep literal strings as text (Excel would turn 'March 31, 2026' into a date serial): prefix an apostrophe."""
    return [[f"'{v}" if isinstance(v, str) and v and not v.startswith(("'", "=")) else v for v in row] for row in grid]


def _watch(watch):
    if not watch:
        return ""
    try:
        return "\n--- watch (recalculated) ---\n" + xl_read(watch, mode="values")
    except Exception as e:
        return f"\n--- watch failed: {e}"


def _runs(cells):
    """{"'Sheet'!C5": v, ...} -> edits of contiguous vertical runs per sheet and column."""
    by = {}
    for addr, v in cells.items():
        sheet, ref = _split(addr)
        m = re.fullmatch(r"([A-Za-z]{1,3})(\d+)", ref)
        if not m:
            raise ValueError(f"cells: '{addr}' must be a single cell like Sheet!C5")
        by.setdefault((sheet, m[1].upper()), []).append((int(m[2]), v))
    edits = []
    for (sheet, col), items in by.items():
        items.sort()
        run = [items[0]]
        for it in items[1:] + [None]:
            if it and it[0] == run[-1][0] + 1:
                run.append(it)
                continue
            edits.append({"address": f"{_q(sheet)}!{col}{run[0][0]}:{col}{run[-1][0]}", "formulas": [[v] for _, v in run]})
            if it:
                run = [it]
    return edits


@mcp.tool()
@_commits
def xl_write(edits: list[dict] | None = None, cells: dict | None = None, note: str = "", atomic: bool = True,
             as_text: bool = False, watch: str = "") -> str:
    """Live write. edits: [{address, formulas (2D grid; "" clears) or formula (one cell), expect (grid you last read;
    a changed cell is not overwritten), text (keep strings as text)}]. cells: {"'Sheet'!C5": value, ...} flat map,
    grouped into runs for you. atomic: one conflict -> nothing written. as_text: all strings stay text (dates!).
    watch: areas read back after writing. note: audit reason. Undo with xl_undo."""
    edits = list(edits or []) + (_runs(cells) if cells else [])
    if not edits:
        return "nothing to write: give edits or cells"
    plan, res = [], []
    for e in edits:
        sheet, rng = _split(e["address"])
        f = e.get("formulas")
        if f is None and "formula" in e:
            f = [[e["formula"]]]
        exp = e.get("expect")
        if isinstance(exp, str):
            exp = [[exp]]
        rows, cols = _shape(rng)
        if f is None or len(f) != rows or any(len(x) != cols for x in f):
            res.append(f"SKIPPED {e['address']}: formulas shape {len(f or [])}x{len(f[0]) if f else 0} != range {rows}x{cols} "
                       "(use xl_fill to copy one formula across a range)")
            continue
        if as_text or e.get("text"):
            f = _as_text(f)
        plan.append((e["address"], sheet, rng, f, exp))
    if atomic and res:
        return "\n".join(res) + "\nNOTHING WRITTEN (atomic): fix the edits above and resend the batch."
    cur = {}
    if atomic:
        conflicts = []
        for addr, sheet, rng, f, exp in plan:
            cur[addr] = _live(sheet, rng)
            if exp is not None and not _same_grid(cur[addr]["formulas"], exp):
                conflicts.append(f"CONFLICT {_q(sheet)}!{rng}: now {cur[addr]['formulas']} (expected {exp})")
        if conflicts:
            return "\n".join(conflicts) + f"\nNOTHING WRITTEN (atomic): {len(conflicts)} of {len(plan)} edits conflict. " \
                                          "Re-read those cells and resend, or pass atomic=false to apply the rest."
    for addr, sheet, rng, f, exp in plan:
        try:
            res.append(_do_write(sheet, rng, f, exp, note, cur=cur.get(addr)))
        except NotSignedIn:
            raise
        except Exception as ex:
            res.append(f"FAILED {addr}: {ex}")
    return "\n".join(res) + _watch(watch)


@mcp.tool()
@_commits
def xl_fill(address: str, formula: str, expect_blank: bool = False, note: str = "", watch: str = "") -> str:
    """One formula, written as for the top-left cell, filled across the range with relative refs adjusting.
    expect_blank refuses if any target cell has content. watch: areas read back after writing."""
    sheet, rng = _split(address)
    c1, r1, c2, r2 = range_boundaries(rng)
    origin = f"{get_column_letter(c1)}{r1}"
    grid = [[Translator(formula, origin).translate_formula(f"{get_column_letter(c)}{r}") if formula.startswith("=") else formula
             for c in range(c1, c2 + 1)] for r in range(r1, r2 + 1)]
    exp = None
    if expect_blank:
        exp = [["" for _ in range(c1, c2 + 1)] for _ in range(r1, r2 + 1)]
    return _do_write(sheet, rng, grid, exp, note or f"fill {formula}") + _watch(watch)


@mcp.tool()
@_commits
def xl_undo(count: int = 1, any_chat: bool = False) -> str:
    """Revert this chat's last writes (newest first; any_chat includes other chats). Cells changed since are skipped."""
    if not os.path.exists(_LOG):
        return "nothing to undo"
    lines = open(_LOG, encoding="utf-8").read().splitlines()
    item = _book()["item"]

    def mine(e):
        return e["item"] == item and not e.get("undone") and e.get("old") is not None and (any_chat or e.get("sid") == _sid())
    picked = [i for i in range(len(lines) - 1, -1, -1) if lines[i].strip() and mine(json.loads(lines[i]))][:count]
    res = []
    for i in picked:
        e = json.loads(lines[i])
        cur = _live(e["sheet"], e["range"])["formulas"]
        if not _same_grid(cur, e["new"]):
            res.append(f"SKIP {_q(e['sheet'])}!{e['range']}: changed since Claude wrote it (now {cur})")
            continue
        _wb("PATCH", f"{_sheet_path(e['sheet'])}/range(address='{e['range']}')", {"formulas": e["old"]})
        e["undone"] = dt.datetime.now().isoformat(timespec="seconds")
        lines[i] = json.dumps(e)
        res.append(f"UNDONE {_q(e['sheet'])}!{e['range']} -> restored {e['old']}")
    open(_LOG, "w", encoding="utf-8").write("\n".join(lines) + "\n")
    return "\n".join(res) or "nothing to undo"


@mcp.tool()
def xl_log(what: str = "writes", limit: int = 20, this_chat_only: bool = False, reset_baseline: bool = False) -> str:
    """what: writes = audit trail of edits ([this chat]/[other chat]) | calls = tool-call timings | changes = cells
    anyone changed since the baseline snapshot (reset_baseline makes now the baseline)."""
    if what == "changes":
        return _changes(reset_baseline)
    item = _book()["item"]
    if what == "calls":
        if not os.path.exists(_CALLS):
            return "no calls logged yet"
        rows = [json.loads(x) for x in open(_CALLS, encoding="utf-8").read().splitlines()[-5000:] if x.strip()]
        rows = [r for r in rows if r.get("item") == item and (not this_chat_only or r.get("sid") == _sid())]
        stats = {}
        for r in rows:
            stats.setdefault(r["tool"], []).append(r["ms"])
        summary = "\n".join(f"  {t:<14} calls {len(v):>4}  avg {sum(v) // len(v):>6} ms  max {max(v):>6} ms"
                            for t, v in sorted(stats.items(), key=lambda kv: -sum(kv[1])))
        recent = "\n".join(f"{r['ts']} {r['tool']:<14} {r['ms']:>6} ms {'ok' if r['ok'] else 'ERROR'}"
                           f"{'' if r.get('sid') == _sid() else '  [other chat]'}" for r in rows[-limit:])
        return f"per tool:\n{summary}\nrecent:\n{recent}"
    if not os.path.exists(_LOG):
        return "no writes yet"
    rows = [json.loads(x) for x in open(_LOG, encoding="utf-8").read().splitlines() if x.strip()]
    me = _sid()
    rows = [r for r in rows if r["item"] == item and (not this_chat_only or r.get("sid") == me)][-limit:]
    return "\n".join(f"{r['ts']} {'[this chat]' if r.get('sid') == me else '[other chat]'} {_q(r['sheet'])}!{r['range']}: "
                     f"{r['old']} -> {r['new']}{'  [undone]' if r.get('undone') else ''}{'  # ' + r['note'] if r.get('note') else ''}"
                     for r in rows) or "no writes yet"


def _changes(reset_baseline, include_values=False):
    snap = _snapshot(force=True)
    d = _snap_dir()
    base = _load_index(os.path.join(d, "baseline.xlsx"))
    cur = snap["f"]
    out = []
    for sh in sorted(set(base) | set(cur)):
        if sh not in cur:
            out.append(f"[sheet removed] {sh}"); continue
        if sh not in base:
            out.append(f"[sheet added] {sh}"); continue
        a, b = base[sh]["cells"], cur[sh]["cells"]
        for k in sorted(set(a) | set(b)):
            fa, va = a.get(k, (None, None)); fb, vb = b.get(k, (None, None))
            inp_a = fa if fa else (None if va is None else _fmt(va))
            inp_b = fb if fb else (None if vb is None else _fmt(vb))
            addr = f"{_q(sh)}!{get_column_letter(k[1])}{k[0]}"
            if inp_a != inp_b:
                out.append(f"{addr}: {inp_a!r} -> {inp_b!r}")
            elif include_values and fa and _fmt(va) != _fmt(vb):
                out.append(f"{addr}: value {_fmt(va)} -> {_fmt(vb)}")
    when = time.strftime("%Y-%m-%d %H:%M", time.localtime(os.path.getmtime(os.path.join(d, "baseline.xlsx"))))
    if reset_baseline:
        shutil.copy(os.path.join(d, "latest.xlsx"), os.path.join(d, "baseline.xlsx"))
    hdr = f"changes since baseline {when} (last saved {snap['modified']} by {snap['by']}): {len(out)}"
    return _clip(hdr + "\n" + "\n".join(out))


def _expand(bd, limit=50):
    return [(r, c) for r in range(bd[0], bd[2] + 1) for c in range(bd[1], bd[3] + 1)] \
        if (bd[2] - bd[0] + 1) * (bd[3] - bd[1] + 1) <= limit else None


def _error_origins(idx, refs_idx, sheets_ok):
    """Group error cells by where the error starts. Returns (origins [(cell, reason, n_downstream)], total errors)."""
    err = {(sh, r, c) for sh in idx for (r, c), (_, v) in idx[sh]["cells"].items() if isinstance(v, str) and v in ERRORS}
    by_sheet = {}
    for sh, r, c in err:
        by_sheet.setdefault(sh, []).append((r, c))
    children, origins = {}, []
    for cell in err:
        sh, r, c = cell
        f = idx[sh]["cells"][(r, c)][0]
        if not f:
            origins.append((cell, "typed-in error value"))
            continue
        lit = next((e for e in ERRORS if e in f), None)
        parents = set()
        for rsh, bd in refs_idx.get(cell, []):
            cand = _expand(bd)
            pts = [(rsh, rr, cc) for rr, cc in cand if (rsh, rr, cc) in err] if cand is not None else \
                  [(rsh, rr, cc) for rr, cc in by_sheet.get(rsh, []) if bd[0] <= rr <= bd[2] and bd[1] <= cc <= bd[3]]
            parents.update(p for p in pts if p != cell)
        for p in parents:
            children.setdefault(p, []).append(cell)
        if lit:
            origins.append((cell, f"formula contains {lit} (broken reference)"))
        elif not parents:
            origins.append((cell, "formula errors by itself (lookup not found, divide by zero, bad argument...)"))
    res = []
    for cell, reason in origins:
        if cell[0] not in sheets_ok:
            continue
        seen, stack = {cell}, [cell]
        while stack:
            for ch in children.get(stack.pop(), []):
                if ch not in seen:
                    seen.add(ch); stack.append(ch)
        res.append((cell, reason, len(seen) - 1))
    res.sort(key=lambda x: -x[2])
    return res, len(err)


def _cycles(idx, refs_idx, max_cycles=10):
    """Circular references: strongly connected groups in the formula graph. Whole-column refs become one node per
    distinct range (pointing at the formula cells inside it) so big workbooks stay fast."""
    fcells = {}
    for sh in idx:
        for (r, c), (f, _) in idx[sh]["cells"].items():
            if f:
                fcells.setdefault((sh, c), []).append(r)
    for v in fcells.values():
        v.sort()
    import bisect
    graph = {}
    for cell, refs in refs_idx.items():
        nbrs = graph.setdefault(cell, set())
        for rsh, bd in refs:
            small = _expand(bd)
            if small is not None:
                nbrs.update((rsh, rr, cc) for rr, cc in small if (rsh, rr, cc) in refs_idx)
            else:
                node = ("RANGE", rsh, bd)
                nbrs.add(node)
                if node not in graph:
                    inside = set()
                    for cc in range(bd[1], min(bd[3], 16384) + 1):
                        rows = fcells.get((rsh, cc))
                        if rows:
                            lo, hi = bisect.bisect_left(rows, bd[0]), bisect.bisect_right(rows, bd[2])
                            inside.update((rsh, rr, cc) for rr in rows[lo:hi])
                    graph[node] = inside
    # iterative Tarjan
    index, low, onstack, stack, sccs, counter = {}, {}, set(), [], [], [0]
    for root in graph:
        if root in index:
            continue
        work = [(root, iter(graph.get(root, ())))]
        index[root] = low[root] = counter[0]; counter[0] += 1
        stack.append(root); onstack.add(root)
        while work:
            node, it = work[-1]
            advanced = False
            for nxt in it:
                if nxt not in index:
                    index[nxt] = low[nxt] = counter[0]; counter[0] += 1
                    stack.append(nxt); onstack.add(nxt)
                    work.append((nxt, iter(graph.get(nxt, ()))))
                    advanced = True
                    break
                if nxt in onstack:
                    low[node] = min(low[node], index[nxt])
            if advanced:
                continue
            work.pop()
            if work:
                low[work[-1][0]] = min(low[work[-1][0]], low[node])
            if low[node] == index[node]:
                comp = []
                while True:
                    x = stack.pop(); onstack.discard(x); comp.append(x)
                    if x == node:
                        break
                cells = [x for x in comp if x[0] != "RANGE"]
                if len(comp) > 1 or node in graph.get(node, ()):
                    sccs.append(cells)
    sccs.sort(key=len)
    return sccs[:max_cycles], len(sccs)


@mcp.tool()
def xl_check(sheets: str = "", visible_only: bool = False, limit: int = 40, cycles: bool = True) -> str:
    """Snapshot health: error ORIGINS (cells where an error starts, with how many cells each breaks - fix those),
    circular references, formulas containing #REF!, external links, named ranges. sheets: comma list; limit per list."""
    snap = _snapshot()
    idx = snap["f"]
    refs_idx = _refs_index(snap)
    ok = set(_sheet_filter(idx, sheets, "", visible_only))
    origins, n_err = _error_origins(idx, refs_idx, ok)
    out = [f"ERRORS: {n_err} error cells in the workbook come from {len(origins)} origin cell(s)"
           + (" in the selected sheets" if sheets or visible_only else "") + " - fix these first:"]
    for (sh, r, c), reason, n in origins[:limit]:
        f, v = idx[sh]["cells"][(r, c)]
        out.append(f"  {_q(sh)}!{get_column_letter(c)}{r} = {v}  breaks {n} other cell(s)  | {f or ''}  [{reason}]")
    if len(origins) > limit:
        out.append(f"  ... {len(origins) - limit} more origins (raise limit or filter sheets)")
    if cycles:
        found, total = _cycles(idx, refs_idx)
        out.append(f"CIRCULAR REFERENCES: {total} loop(s)" + (":" if total else " - none"))
        for comp in found:
            names = [f"{_q(sh)}!{get_column_letter(c)}{r}" for sh, r, c in sorted(comp)]
            out.append(f"  loop of {len(comp)} cell(s): {', '.join(names[:10])}" + (" ..." if len(names) > 10 else ""))
    refs, ext = [], []
    for sh in ok:
        for (r, c), (f, _) in sorted(idx[sh]["cells"].items()):
            if not f:
                continue
            a = f"{_q(sh)}!{get_column_letter(c)}{r}"
            if "#REF!" in f:
                refs.append(f"  {a} | {f}")
            if re.search(r"\[\d+\]|\.xls[xmb]?\]", re.sub(r'"[^"]*"', "", f), re.I):
                ext.append(f"  {a} | {f}")
    out += [f"FORMULAS CONTAINING #REF!: {len(refs)}", *refs[:limit], f"EXTERNAL LINKS: {len(ext)}", *ext[:limit]]
    n_names, n_broken = snap.get("names", (0, 0))
    out.append(f"NAMED RANGES: {n_names}" + (f" ({n_broken} broken with #REF!)" if n_broken else "")
               + ("  - thousands of names bloat a file; only xl_file replace can remove them in bulk" if n_names > 500 else ""))
    return _clip("\n".join(out))


_WHOLE_COL = re.compile(r"[A-Za-z]{1,3}(:[A-Za-z]{1,3})?")
_WHOLE_ROW = re.compile(r"\d+(:\d+)?")


@mcp.tool()
@_commits
def xl_shift(address: str, action: str = "insert", shift: str = "", note: str = "") -> str:
    """Insert or delete cells: rows 'BS Notes'!12:14, columns Sheet!B:C, or a cell range. action: insert | delete
    (delete is not undoable: confirm first). shift: Down|Right / Up|Left, automatic for whole rows/columns.
    Saved at once; ask first if anyone has the file open in desktop Excel."""
    sheet, rng = _split(address)
    ins = action.lower().startswith("ins")
    if not shift:
        shift = ("Right" if ins else "Left") if _WHOLE_COL.fullmatch(rng) else ("Down" if ins else "Up")
    _wb("POST", f"{_sheet_path(sheet)}/range(address='{rng}')/{'insert' if ins else 'delete'}", {"shift": shift.capitalize()})
    _log({"ts": dt.datetime.now().isoformat(timespec="seconds"), "item": _book()["item"], "sheet": sheet, "range": rng,
          "old": None, "new": f"{'INSERT' if ins else 'DELETE'} {shift}", "note": note, "undone": "n/a (structural)"})
    _close_session(_book()["item"], _book()["drive"])  # save now: co-authors should see the new layout at once
    return f"{'inserted' if ins else 'deleted'} {_q(sheet)}!{rng} shift {shift}"


_H = {"left": "Left", "center": "Center", "centre": "Center", "right": "Right", "general": "General"}
_V = {"top": "Top", "center": "Center", "centre": "Center", "middle": "Center", "bottom": "Bottom"}
_EDGES = {"all": ["EdgeTop", "EdgeBottom", "EdgeLeft", "EdgeRight", "InsideHorizontal", "InsideVertical"],
          "outline": ["EdgeTop", "EdgeBottom", "EdgeLeft", "EdgeRight"], "top": ["EdgeTop"], "bottom": ["EdgeBottom"],
          "left": ["EdgeLeft"], "right": ["EdgeRight"]}
_EDGES["none"] = _EDGES["all"]


def _format_one(sheet, rng, props):
    p = f"{_sheet_path(sheet)}/range(address='{rng}')"
    done = []
    whole_col, whole_row = _WHOLE_COL.fullmatch(rng), _WHOLE_ROW.fullmatch(rng)
    nf = props.get("number_format")
    if nf:
        if whole_col or whole_row:
            raise ValueError("number_format needs explicit rows and columns (e.g. A2:A700), not whole columns/rows")
        rows, cols = _shape(rng)
        if rows * cols > 250000:
            raise ValueError("number_format area over 250k cells: split it")
        _wb("PATCH", p, {"numberFormat": [[nf] * cols for _ in range(rows)]}); done.append("number_format")
    font = {k2: props[k] for k, k2 in (("bold", "bold"), ("italic", "italic"), ("size", "size"), ("font_color", "color"))
            if props.get(k) is not None}
    if font:
        _wb("PATCH", p + "/format/font", font); done.append("font")
    fill = props.get("fill")
    if fill:
        if str(fill).lower() in ("none", "clear"):
            _wb("POST", p + "/format/fill/clear"); done.append("fill cleared")
        else:
            _wb("PATCH", p + "/format/fill", {"color": fill}); done.append("fill")
    body = {}
    if props.get("wrap") is not None:
        body["wrapText"] = bool(props["wrap"])
    if props.get("align"):
        body["horizontalAlignment"] = _H.get(str(props["align"]).lower(), str(props["align"]).capitalize())
    if props.get("valign"):
        body["verticalAlignment"] = _V.get(str(props["valign"]).lower(), str(props["valign"]).capitalize())
    w = props.get("width")
    if w not in (None, "", "auto"):
        body["columnWidth"] = float(w)
    if body:
        _wb("PATCH", p + "/format", body); done.append("layout")
    if w == "auto":
        _wb("POST", p + "/format/autofitColumns"); done.append("autofit")
    if props.get("hidden") is not None:
        hidden = bool(props["hidden"])
        if whole_col:
            a, b = rng.split(":")[0], rng.split(":")[-1]
            c1, c2, r1, r2 = column_index_from_string(a.upper()), column_index_from_string(b.upper()), 1, 1
        elif whole_row:
            a, b = rng.split(":")[0], rng.split(":")[-1]
            r1, r2, c1, c2 = int(a), int(b), 1, 1
        else:
            c1, r1, c2, r2 = range_boundaries(rng)
        if whole_col or (not whole_row and r1 == r2):
            # columns: a row-1 address answers in 0.3 s; a whole-column address takes ~50 s on Graph
            _wb("PATCH", f"{_sheet_path(sheet)}/range(address='{get_column_letter(c1)}1:{get_column_letter(c2)}1')", {"columnHidden": hidden})
            done.append(("hid" if hidden else "unhid") + " columns")
        else:
            _wb("PATCH", f"{_sheet_path(sheet)}/range(address='A{r1}:A{r2}')", {"rowHidden": hidden})
            done.append(("hid" if hidden else "unhid") + " rows")
    border = props.get("border")
    if border:
        edges = _EDGES.get(str(border).lower())
        if not edges:
            raise ValueError("border must be all | outline | top | bottom | left | right | none")
        style = {"style": "None"} if str(border).lower() == "none" else {"style": "Continuous", "color": props.get("border_color", "#000000")}
        for e in edges:
            _wb("PATCH", f"{p}/format/borders/{e}", style)
        done.append(f"border {border}")
    return done


@mcp.tool()
@_commits
def xl_format(address: str = "", props: dict | None = None, specs: list[dict] | None = None) -> str:
    """Format live cells. address: one or many areas ("Sheet!A1:B2,C4:D9"). props keys: number_format, bold, italic,
    size, font_color, fill ('#RRGGBB' or 'none'), wrap, align (left|center|right), valign (top|center|bottom), width
    (points or 'auto'), hidden (true/false: the area's columns, or rows for a column range), border (all|outline|top|
    bottom|left|right|none). specs: [{address, ...props}] for many areas with different props in one call."""
    jobs = [(address, props or {})] if address else []
    for s in specs or []:
        s = dict(s)
        jobs.append((s.pop("address", ""), s))
    res = []
    for addr, pr in jobs:
        if not addr:
            res.append("SKIPPED spec without address")
            continue
        for sheet, rng in _areas(addr):
            try:
                done = _format_one(sheet, rng, pr)
                res.append(f"formatted {_q(sheet)}!{rng}: {', '.join(done) or 'nothing (no props given)'}")
            except (GraphError, ValueError) as e:
                res.append(f"FAILED {_q(sheet)}!{rng}: {str(e)[:200]}")
    return "\n".join(res) or "nothing to format: give address + props, or specs"


_VIS = {"hide": "Hidden", "unhide": "Visible", "show": "Visible", "veryhide": "VeryHidden"}


@mcp.tool()
@_commits
def xl_sheet(names: str = "", action: str = "list", new_name: str = "", position: int = -1, color: str = "") -> str:
    """Worksheets. action: list | hide | unhide | veryhide | add | delete | rename | move | color. names: one sheet, a
    comma list, or * for all (add: the new name). rename/move take one sheet (new_name / 0-based position);
    color '#RRGGBB'. delete is not undoable: confirm first."""
    act = action.lower().strip() or "list"
    sheets = _wb("GET", "/worksheets?$select=name,visibility,position")["value"]
    if act == "list":
        return "\n".join(f"{s['position']:>3} {s['name']!r:<34} {s['visibility']}" for s in sheets)
    if act == "add":
        r = _wb("POST", "/worksheets/add", {"name": names} if names.strip() else {})
        return f"added sheet {r.get('name')!r} at position {r.get('position')}"
    if names.strip() == "*":
        targets = [s["name"] for s in sheets]
    else:
        wanted = [n.strip() for n in names.split(",") if n.strip()]
        by_lower = {s["name"].strip().lower(): s["name"] for s in sheets}
        missing = [n for n in wanted if n.lower() not in by_lower]
        if missing or not wanted:
            return f"No sheet(s) {missing or '(none given)'}. Sheets: {', '.join(s['name'] for s in sheets)}"
        targets = [by_lower[n.lower()] for n in wanted]
    if act == "delete":
        for t in targets:
            _wb("DELETE", _sheet_path(t))
        _log({"ts": dt.datetime.now().isoformat(timespec="seconds"), "item": _book()["item"], "sheet": ",".join(targets),
              "range": "", "old": None, "new": "DELETE SHEET", "note": ""})
        return "deleted sheet(s): " + ", ".join(repr(t) for t in targets)
    if act in _VIS:
        body = {"visibility": _VIS[act]}
    elif act == "rename":
        body = {"name": new_name}
    elif act == "move":
        body = {"position": position}
    elif act == "color":
        body = {"tabColor": color}
    else:
        return "action must be list | hide | unhide | veryhide | add | delete | rename | move | color"
    if act in ("rename", "move") and len(targets) != 1:
        return f"{act} needs exactly one sheet"
    out = []
    for t in targets:
        try:
            r = _wb("PATCH", _sheet_path(t), body)
            out.append(f"OK {t!r} -> {r.get('name')!r} {r.get('visibility')} pos {r.get('position')}")
        except GraphError as e:
            out.append(f"FAILED {t!r}: {str(e)[:200]}")
    _log({"ts": dt.datetime.now().isoformat(timespec="seconds"), "item": _book()["item"], "sheet": ",".join(targets),
          "range": "", "old": None, "new": f"SHEET {body}", "note": "sheet properties"})
    return "\n".join(out)


def _upload(drive, item, data):
    """Replace a file's content through an upload session, 10 MiB chunks (a multiple of the 320 KiB Graph requires)."""
    up = _graph("POST", f"/drives/{drive}/items/{item}/createUploadSession", {"item": {"@microsoft.graph.conflictBehavior": "replace"}})
    url, total, pos, body = up["uploadUrl"], len(data), 0, b""
    while pos < total:
        chunk = data[pos:pos + 10 * 1024 * 1024]
        req = urllib.request.Request(url, data=chunk, method="PUT", headers={
            "Content-Length": str(len(chunk)), "Content-Range": f"bytes {pos}-{pos + len(chunk) - 1}/{total}"})
        with urllib.request.urlopen(req, timeout=300) as r:
            body = r.read()
        pos += len(chunk)
    return json.loads(body) if body else {}


@mcp.tool()
def xl_file(action: str, name: str = "", path: str = "") -> str:
    """File-level actions. copy: server-side copy of the open workbook as `name`, then opens it | link: org-wide edit
    link | list: your Excel files (matching `name`) | snapshot: local path of an up-to-date read-only copy for
    openpyxl/pandas | save: flush pending edits now | calc: full recalculation | replace: upload the local file at
    `path` over the open workbook, only for changes Graph cannot make (e.g. deleting thousands of named ranges);
    refused if anyone edited the file since the snapshot. Confirm replace with the user first."""
    act = action.lower().strip()
    if act == "list":
        q = name.strip().lower()
        r = _graph("GET", "/me/drive/root/children?$select=name,lastModifiedDateTime&$orderby=lastModifiedDateTime desc&$top=200")
        seen = {}
        for it in r.get("value", []):  # own OneDrive root: immediate, even for a file created seconds ago
            if it["name"].lower().endswith((".xlsx", ".xlsm")) and q in it["name"].lower():
                seen[it["name"]] = it["lastModifiedDateTime"][:16]
        if q:
            for h in _search_books(q):  # shared files and subfolders (the search index lags a few minutes)
                seen.setdefault(h["name"], h["modified"])
        return "\n".join(f"  {n}  (modified {m})" for n, m in list(seen.items())[:60]) or "no Excel files found"
    b = _book()
    base = f"/drives/{b['drive']}/items/{b['item']}"
    if act == "copy":
        if not name.strip():
            return "copy needs name"
        if not name.lower().endswith((".xlsx", ".xlsm")):
            name += ".xlsx"
        _flush_all()
        parent = _graph("GET", base + "?$select=parentReference")["parentReference"]["id"]
        _graph("POST", base + "/copy", {"name": name})  # accepted (202); the copy lands next to the source
        it = None
        for _ in range(40):
            time.sleep(1.5)
            try:
                it = _graph("GET", f"/drives/{b['drive']}/items/{parent}:/{name}:?$select=id,name")
                break
            except GraphError:
                pass
        if not it:
            return f"copy requested, but '{name}' has not appeared yet - xl_open('{name}') in a minute"
        _open_item({"name": it["name"], "drive": b["drive"], "item": it["id"]})
        return f"copied and opened {it['name']!r} - it is now this chat's workbook\n" + xl_file("link")
    if act == "link":
        r = _graph("POST", base + "/createLink", {"type": "edit", "scope": "organization"})
        return f"edit link (organisation): {r.get('link', {}).get('webUrl', '?')}"
    if act == "save":
        _flush_all()
        return "saved: pending edits are in the file"
    if act == "calc":
        _wb("POST", "/application/calculate", {"calculationType": "Full"})
        mode = _wb("GET", "/application").get("calculationMode")
        _commit()
        return f"recalculated (calculation mode: {mode})"
    if act == "snapshot":
        _flush_all()
        snap = _snapshot(force=True)
        return (f"{os.path.join(_snap_dir(), 'latest.xlsx')}\n(last saved {snap['modified']} by {snap['by']}; read-only copy: "
                "analyse it, but make changes through xl_write - or xl_file replace for what Graph cannot do)")
    if act == "replace":
        if not path or not os.path.isfile(path) or not path.lower().endswith((".xlsx", ".xlsm")):
            return "replace needs path to a local .xlsx/.xlsm file"
        latest_meta = os.path.join(_snap_dir(), "latest.json")
        if not os.path.exists(latest_meta):
            return "REFUSED: take xl_file snapshot first, make the offline changes on that copy, then replace"
        _flush_all()
        meta = _graph("GET", base + "?$select=cTag,lastModifiedDateTime,lastModifiedBy")
        if json.load(open(latest_meta)).get("ctag") != meta["cTag"]:
            who = meta.get("lastModifiedBy", {}).get("user", {}).get("displayName")
            return (f"REFUSED: the file changed since your snapshot (last saved {meta['lastModifiedDateTime']} by {who}); "
                    "uploading would discard those edits. Take a fresh snapshot, redo the offline changes on it, then replace.")
        data = open(path, "rb").read()
        try:
            _upload(b["drive"], b["item"], data)
        except (GraphError, urllib.error.HTTPError) as e:
            code = getattr(e, "code", None) or (423 if " 423 " in str(e) else None)
            if code == 423:
                return "LOCKED: someone has the file open in Excel (desktop or browser). Ask them to close it, then retry."
            return f"upload failed: {str(e)[:300]}"
        _sessions.pop(b["item"], None)
        _snaps.pop(b["item"], None)
        _log({"ts": dt.datetime.now().isoformat(timespec="seconds"), "item": b["item"], "sheet": "", "range": "",
              "old": None, "new": f"REPLACE FILE from {os.path.basename(path)} ({len(data) // 1024} KB)", "note": name,
              "undone": "n/a (file)"})
        snap = _snapshot(force=True)
        return f"replaced {b['name']!r} with {os.path.basename(path)} ({len(data) // 1024} KB); saved {snap['modified']}"
    return "action must be copy | link | list | snapshot | save | calc | replace"


@mcp.tool()
def xl_api(method: str, path: str, body: dict | None = None) -> str:
    """Raw Microsoft Graph for anything else. Relative path = the open workbook (/names, /tables, /worksheets('BS')/charts,
    /application/calculate); absolute (/me/..., /drives/...) = plain Graph v1.0. Not available on Graph: cell comments,
    freeze panes, conditional formats."""
    method = method.upper()
    wb_path = path[len("/workbook"):] if path.startswith("/workbook") else path
    is_workbook = path.startswith("/workbook") or not path.startswith(("/me", "/drives", "/sites", "/users", "/shares", "/groups", "/search"))
    try:
        if is_workbook:
            res = _wb(method, wb_path if wb_path.startswith("/") else "/" + wb_path, body)
            if method != "GET":
                _log({"ts": dt.datetime.now().isoformat(timespec="seconds"), "item": _book()["item"], "sheet": "",
                      "range": path, "old": None, "new": f"API {method} {json.dumps(body)[:300]}", "note": "xl_api"})
                _commit()
        else:
            res = _graph(method, path, body)
    except GraphError as e:
        return f"Graph error: {e}"
    return _clip(json.dumps(res, indent=1, default=str))


if __name__ == "__main__":
    mcp.run("stdio")
