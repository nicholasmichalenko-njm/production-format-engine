"""Content-free diagnostics for FormatCheck: errors, quiet failures and a feature activity trail.
Ported from Stationkeeper's diag.py so both tools log the same way and one reader understands both.

What a log line may hold (enforced here, by tools/check_diag.py in the build, and again before export):
    time, app version, process role, level, area, a fixed message written in our source code, the exception TYPE,
    code locations (our file name, line, function), and whitelisted counts / file extensions / short codes.
What it never holds:
    file names, paths, folder names, document text, style or field contents, exception MESSAGES, user names.

Files (all in logs\\ next to the program, never sent anywhere by the app):
    events.jsonl     errors and warnings, one JSON object per line (rotated at 5 MB, 3 files kept)
    activity.jsonl   feature actions and clicks with counts (rotated the same way)
    crash_raw.txt    written by Python's faulthandler only on a hard crash of the window; converted to a
                     content-free event at the next start and then deleted (it holds paths, so it is never exported)
"""
import datetime, faulthandler, json, os, re, sys, threading, traceback

VERSION = "?"
ROLE = "app"               # "app" (window or command line) or "word_worker" (the per-document Word process)
LOG_DIR = None
_lock = threading.Lock()
_seen = {}
_flushed = {}
_where = {}
_run = None                # counters while a check runs
_recent = []
_crash_fh = None
_names = {}
_last_action = [None]
usage_on = True
MAX_BYTES = 5 * 1024 * 1024
KEEP = 3

# Count fields a caller may attach. Values must be ints, or a file extension for "ext", or a short code for "code".
COUNT_KEYS = {"n", "files", "docs", "pages", "pages_after", "paragraphs", "tables", "sections", "seconds", "ms",
              "steps", "ok", "failed", "refused", "accepted", "errors", "warnings", "guard_trips", "op_errors",
              "items", "exit_code", "linked", "drift", "mb", "timeout", "word", "recursive", "blocked", "timed_out",
              "hresult", "scode", "broken_refs", "reports", "kept", "stopped", "dirty_parts"}
_EXT = re.compile(r"^\.[a-z0-9]{1,6}$")
_CODE = re.compile(r"^[a-z][a-z0-9_]{0,40}$")
_AREA = re.compile(r"^[a-z][a-z0-9_]{0,30}$")
_W = re.compile(r"^[a-z][a-z0-9_.]{0,60}$")


def setup(log_dir, version, role="app", crash_capture=True):
    """Call once at start: hooks for uncaught errors in any thread, hard-crash capture, old crash conversion."""
    global LOG_DIR, VERSION, ROLE, _crash_fh
    LOG_DIR, VERSION, ROLE = log_dir, version, role
    try:
        os.makedirs(LOG_DIR, exist_ok=True)
    except OSError:
        LOG_DIR = None
        return
    if crash_capture:
        _convert_old_crash()

    def _hook(t, e, tb):
        error("app", "Uncaught error", exc=(t, e, tb))
        if not getattr(sys, "frozen", False):
            sys.__excepthook__(t, e, tb)
    sys.excepthook = _hook

    def _thook(a):
        error("thread", "Uncaught error in a background task", exc=(a.exc_type, a.exc_value, a.exc_traceback))
        if not getattr(sys, "frozen", False):
            threading.__excepthook__(a)
    threading.excepthook = _thook
    if crash_capture:
        try:
            _crash_fh = open(os.path.join(LOG_DIR, "crash_raw.txt"), "w", encoding="utf-8")
            faulthandler.enable(_crash_fh, all_threads=True)
        except Exception:
            _crash_fh = None


def shutdown():
    try:
        _flush_repeats()
        if _crash_fh is not None:
            faulthandler.disable(); _crash_fh.close()
            p = os.path.join(LOG_DIR, "crash_raw.txt")
            if os.path.isfile(p) and os.path.getsize(p) == 0:
                os.remove(p)
    except Exception:
        pass


# ---------------------------------------------------------------- recording
def error(area, msg, exc=None, **counts):
    _record("error", area, msg, _exc(exc), counts)


def warning(area, msg, exc=None, **counts):
    _record("warning", area, msg, _exc(exc), counts)


def swallowed(area=None):
    """Call inside an `except` block that recovers on purpose. Logs the error type and where, once per place."""
    t, e, tb = sys.exc_info()
    if t is None or issubclass(t, FileNotFoundError):
        return
    if area is None:
        area = _clean_area(sys._getframe(1).f_code.co_name)
    _record("warning", area, "Recovered from an error", (t, e, tb), {})


def activity(action, w=None, state=None, **counts):
    """A feature action: 'run_started', 'doc_done', 'feedback_saved'... with counts only.
    w = internal widget name for clicks (code identifiers only), state = 'disabled' for a greyed-out control."""
    if not _AREA.match(action or ""):
        action = "other"
    now = datetime.datetime.now()
    entry = {"time": now.isoformat(timespec="seconds"), "version": VERSION, "action": action}
    if ROLE != "app":
        entry["role"] = ROLE
    if _last_action[0] is not None:
        entry["ms"] = min(int((now - _last_action[0]).total_seconds() * 1000), 3600000)
    _last_action[0] = now
    if isinstance(w, str) and _W.match(w):
        entry["w"] = w
    if state == "disabled":
        entry["state"] = "disabled"
    entry.update(_counts(counts))
    with _lock:
        _recent.append(entry); del _recent[:-200]
    _write("activity.jsonl", entry)


def recent_activity(n=20):
    with _lock:
        return list(_recent[-n:])


def issues_this_session():
    with _lock:
        return sum(c for k, c in _seen.items() if k[0] in ("error", "warning"))


# ---------------------------------------------------------------- per-run health
def run_started(**counts):
    global _run
    with _lock:
        _run = {"t0": datetime.datetime.now(), "e0": sum(c for k, c in _seen.items() if k[0] == "error"),
                "w0": sum(c for k, c in _seen.items() if k[0] == "warning")}
    activity("run_started", **counts)


def run_finished(stopped=False, **counts):
    global _run
    with _lock:
        s, _run = _run, None
        e1 = sum(c for k, c in _seen.items() if k[0] == "error")
        w1 = sum(c for k, c in _seen.items() if k[0] == "warning")
    if s is None:
        return
    secs = int((datetime.datetime.now() - s["t0"]).total_seconds())
    activity("run_stopped" if stopped else "run_finished", seconds=secs, errors=e1 - s["e0"], warnings=w1 - s["w0"], **counts)
    _flush_repeats()


# ---------------------------------------------------------------- usage: which controls get clicked (tkinter)
def register(widget, name):
    if isinstance(name, str) and _W.match(name):
        _names[id(widget)] = name


def install_click_logger(root):
    """One line per click: the nearest named control, and whether it was greyed out. Never text or anything typed."""
    def on_click(e):
        if not usage_on:
            return
        try:
            w, name, state, hops = e.widget, None, None, 0
            while w is not None and hops < 12:
                try:
                    if state is None and str(w.cget("state")) == "disabled":
                        state = "disabled"
                except Exception:
                    pass
                if id(w) in _names:
                    name = _names[id(w)]; break
                w = getattr(w, "master", None); hops += 1
            if name is None:
                name = "unnamed." + re.sub(r"[^a-z0-9_]", "", type(e.widget).__name__.lower())[:30]
            activity("click", w=name[:60], state=state)
        except Exception:
            pass
    root.bind_all("<ButtonRelease-1>", on_click, add="+")


# ---------------------------------------------------------------- internals
def _exc(exc):
    if exc is True:
        return sys.exc_info()
    return exc


def _now():
    return datetime.datetime.now().isoformat(timespec="seconds")


def _clean_area(name):
    name = re.sub(r"[^a-z0-9_]", "_", (name or "app").lower()).strip("_") or "app"
    return name[:30] if _AREA.match(name[:30]) else "app"


def _counts(counts):
    out = {}
    for k, v in (counts or {}).items():
        if k == "ext":
            if isinstance(v, str) and _EXT.match(v.lower()):
                out["ext"] = v.lower()
        elif k == "code":
            if isinstance(v, str) and _CODE.match(v):
                out["code"] = v
        elif (k in COUNT_KEYS or re.match(r"^ext_[a-z0-9]{1,6}$", k)) and isinstance(v, int) and not isinstance(v, bool):
            out[k] = v
        elif k in COUNT_KEYS and isinstance(v, bool):
            out[k] = int(v)
    return out


def _frames(tb):
    """Code locations only: file NAME (never its folder), line number, function name."""
    out = []
    for fr in traceback.extract_tb(tb)[-10:]:
        name = re.sub(r"[^A-Za-z0-9_.\-]", "_", os.path.basename(fr.filename))[:57]
        if not name.endswith(".py"):
            name += ".py"
        fn = re.sub(r"[^A-Za-z0-9_<>]", "_", fr.name or "?")[:60]
        out.append(f"{name}:{fr.lineno}:{fn}")
    return out


def _record(level, area, msg, exc, counts):
    area = area if isinstance(area, str) and _AREA.match(area) else "app"
    msg = msg if isinstance(msg, str) else "?"
    entry = {"time": _now(), "version": VERSION, "level": level, "area": area, "msg": msg[:160]}
    if ROLE != "app":
        entry["role"] = ROLE
    if exc is not None and exc[0] is not None:
        t, _e, tb = exc
        entry["type"] = re.sub(r"[^A-Za-z0-9_.]", "_", getattr(t, "__name__", "Error"))[:60]
        entry["where"] = _frames(tb)
    entry.update(_counts(counts))
    key = (level, area, entry["msg"], entry.get("type"), (entry.get("where") or [""])[-1])
    with _lock:
        n = _seen.get(key, 0) + 1
        _seen[key] = n
        if n == 1:
            _where[key] = entry.get("where", [])
    if n == 1:
        _write("events.jsonl", entry)


def _flush_repeats():
    with _lock:
        rep = [(k, c) for k, c in _seen.items() if c > 1 and _flushed.get(k) != c]
        for k, c in rep:
            _flushed[k] = c
    for (level, area, msg, typ, where), c in rep:
        e = {"time": _now(), "version": VERSION, "level": level, "area": area, "msg": msg,
             "where": _where.get((level, area, msg, typ, where), [where] if where else []), "repeats": c}
        if typ:
            e["type"] = typ
        if ROLE != "app":
            e["role"] = ROLE
        _write("events.jsonl", e)


def _write(name, entry):
    if not LOG_DIR:
        return
    p = os.path.join(LOG_DIR, name)
    try:
        with _lock:
            if os.path.isfile(p) and os.path.getsize(p) > MAX_BYTES:
                for i in range(KEEP - 1, 0, -1):
                    a, b = f"{p}.{i}", f"{p}.{i + 1}"
                    if os.path.exists(a):
                        os.replace(a, b)
                os.replace(p, p + ".1")
            with open(p, "a", encoding="utf-8") as f:
                f.write(json.dumps(entry, ensure_ascii=False) + "\n")
    except Exception:
        pass


_FH_LINE = re.compile(r'File "(?P<f>[^"]+)", line (?P<l>\d+) in (?P<fn>\S+)')


def _convert_old_crash():
    p = os.path.join(LOG_DIR, "crash_raw.txt")
    try:
        if os.path.isfile(p) and os.path.getsize(p) > 0:
            raw = open(p, encoding="utf-8", errors="replace").read()
            where = []
            for m in _FH_LINE.finditer(raw):
                f = re.sub(r"[^A-Za-z0-9_.\-]", "_", os.path.basename(m["f"].replace(chr(92), "/")))[:57]
                fn = re.sub(r"[^A-Za-z0-9_<>]", "_", m["fn"])[:60]
                where.append(f"{f if f.endswith('.py') else f + '.py'}:{m['l']}:{fn}")
            kind = "Fatal error"
            for k in ("Segmentation fault", "access violation", "Fatal Python error", "Aborted"):
                if k.lower() in raw.lower():
                    kind = k.replace(" ", "_")
            _write("events.jsonl", {"time": _now(), "version": VERSION, "level": "crash", "area": "app",
                                    "msg": "The previous session crashed", "type": kind, "where": where[:12]})
        if os.path.isfile(p):
            os.remove(p)
    except Exception:
        pass
