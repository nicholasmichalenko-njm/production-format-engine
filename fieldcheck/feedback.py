"""FormatCheck feedback: bugs, improvements and ideas about the tool's FEATURES, written as plain text,
and one Export for Claude zip that carries the reports, the content-free logs and the run summaries.
Same rules and same export layout as Stationkeeper:
  - No screenshots, video or attachments. Nothing the tool processes may leave the PC.
  - Reports describe features, never document content.
  - A content guard checks every report and every file of an export against what this PC knows is sensitive
    (file and folder names of every volume FormatCheck has checked, the folders picked, the Windows user and PC
    name). A match blocks the save or the export. The guard runs only on this PC.
  - Machine-written files (logs, run summaries) are also proven against a fixed schema before export.
Files: feedback\\reports\\<time>.json (one per report). The export is one text-only zip saved where the user chooses.
The app never sends anything anywhere.
"""
import datetime, glob, json, os, platform, re, sys, zipfile
import diag

TYPES = ["Bug", "Improvement", "Idea"]
FEATURES = ["Pick a folder", "Run a check", "Engine operations", "Full fix on the copy", "Word check", "Report",
            "Speed", "Window and layout", "Feedback and export", "Other"]
IMPORTANCE = ["Low", "Medium", "High"]
GENERIC = {"users", "desktop", "documents", "downloads", "onedrive", "sharepoint", "c:", "d:", "program files",
           "appdata", "local", "temp", "public", "home", "pictures", "formatcheck", "formatcheck_runs", "files",
           "new folder", "shared documents", "sites", "volume", "volumes", "final", "draft", "docx", "word"}
COMMON = set("""a about after all also an and any are as at be because been before being both but by can could did do
does done each even for from get got had has have how if in into is it its just like make many may more most much must
my no not now of on one only or other our out over same see should so some such than that the their them then there
these they this those to too two up us use used very was way we well were what when where which while who why will with
would you your table tables page pages word style styles heading headings caption captions header footer report run
check folder copy original fix field fields number numbers section sections list lists error errors button window open
opens opened file files document documents format formatting engine guard step steps time slow fast after before click
dialog option options shows show showed said says could should want need better easier please would like""".split())
_DATA = [None]


def data_dir():
    import fieldcheck as fc
    return fc.HERE


def reports_dir():
    d = os.path.join(data_dir(), "feedback", "reports")
    os.makedirs(d, exist_ok=True)
    return d


def runs_dir():
    import fieldcheck as fc
    return fc.RUNS_DIR


def log_dir():
    import fieldcheck as fc
    return fc.LOG_DIR


# ------------------------------------------------------------------ the guard
class Guard:
    """Knows what must not leave the PC. Built fresh each time from this PC's FormatCheck history."""

    def __init__(self, extra_paths=()):
        self.exact, self.phrases = set(), set()
        paths = list(extra_paths)
        for key in glob.glob(os.path.join(runs_dir(), "run_*", "KEY_stays_on_this_PC.txt")):
            try:
                for line in open(key, encoding="utf-8"):
                    if "\t" in line:
                        paths.append(line.rstrip("\n").split("\t", 1)[1])
            except Exception:
                diag.swallowed("feedback")
        for p in paths:
            self._add_path(p)
        for v in (os.environ.get("USERNAME"), os.environ.get("USER"), os.environ.get("COMPUTERNAME"),
                  os.environ.get("USERDOMAIN")):
            if v and len(v) >= 3:
                self.phrases.add(v.lower())
        self.phrases = {p for p in self.phrases if len(p) >= 4 and p not in COMMON}
        self.exact = {e for e in self.exact if len(e) >= 2}

    def _add_path(self, p):
        parts = [x for x in re.split(r"[\\/]+", p or "") if x]
        for i, comp in enumerate(parts):
            if comp.lower() in GENERIC or re.fullmatch(r"[a-zA-Z]:", comp):
                continue
            stem = os.path.splitext(comp)[0] if i == len(parts) - 1 else comp
            if len(stem) >= 4 and stem.lower() not in COMMON:
                self.phrases.add(stem.lower())
            for piece in re.split(r"[_\-\s.()\[\],]+", stem):
                if not piece or piece.isdigit():
                    continue
                if self._acronymish(piece):
                    self.exact.add(piece)
                elif len(piece) >= 4 and piece.lower() not in COMMON and piece.lower() not in GENERIC:
                    self.phrases.add(piece.lower())

    @staticmethod
    def _acronymish(w):
        return len(w) >= 2 and (sum(c.isupper() for c in w) >= 2 or (any(c.isdigit() for c in w) and any(c.isalpha() for c in w)))

    def find(self, text):
        """Matches of sensitive items in text: list of (start, end). The items are never stored or shown outside this PC."""
        hits = []
        if not text:
            return hits
        low = text.lower()
        for ph in self.phrases:
            for m in re.finditer(r"(?<![a-z0-9])" + re.escape(ph) + r"(?![a-z0-9])", low):
                hits.append((m.start(), m.end()))
        for ex in self.exact:
            for m in re.finditer(r"(?<![A-Za-z0-9])" + re.escape(ex) + r"(?![A-Za-z0-9])", text):
                hits.append((m.start(), m.end()))
        return sorted(set(hits))


# ------------------------------------------------------------------ reports
def context():
    import fieldcheck as fc
    return {"version": fc.VERSION, "time": datetime.datetime.now().isoformat(timespec="seconds"),
            "recent_actions": [a.get("action") + (":" + a["w"] if a.get("w") else "") for a in diag.recent_activity(20)],
            "issues_logged_this_session": diag.issues_this_session()}


def save_report(rtype, feature, importance, title, text, guard=None):
    """Returns (ok, message). Blocks the save when the guard finds content from this PC."""
    title, text = (title or "").strip(), (text or "").strip()
    if not title or not text:
        return False, "Add a title and a description."
    g = guard or Guard()
    bad = [k for k, v in (("title", title), ("description", text)) if g.find(v)]
    if bad:
        diag.activity("feedback_blocked")
        return False, ("Not saved: the highlighted words look like a file or folder name from your volumes, or your user "
                       f"or PC name ({' and '.join(bad)}). Describe the feature instead.")
    rep = {"type": rtype if rtype in TYPES else "Idea", "feature": feature if feature in FEATURES else "Other",
           "importance": importance if importance in IMPORTANCE else "Medium", "title": title, "text": text,
           "context": context(), "exported": False}
    name = datetime.datetime.now().strftime("%Y%m%d-%H%M%S-%f") + ".json"
    with open(os.path.join(reports_dir(), name), "w", encoding="utf-8") as f:
        json.dump(rep, f, ensure_ascii=False, indent=1)
    diag.activity("feedback_saved")
    return True, "Saved on this PC. Use Export for Claude when you want to hand reports over."


def list_reports():
    out = []
    for p in sorted(glob.glob(os.path.join(reports_dir(), "*.json"))):
        try:
            out.append((p, json.load(open(p, encoding="utf-8"))))
        except Exception:
            diag.swallowed("feedback")
    return out


# ------------------------------------------------------------------ system and run summaries
def system_info():
    import fieldcheck as fc
    info = {"version": fc.VERSION, "frozen": bool(getattr(sys, "frozen", False)), "windows": platform.platform(terse=True),
            "python": platform.python_version(), "cpus": os.cpu_count()}
    builds = [r.get("word_build") for r in run_summaries() if r.get("word_build")]
    info["word_build"] = builds[-1] if builds else None
    try:
        import ctypes

        class MS(ctypes.Structure):
            _fields_ = [("dwLength", ctypes.c_ulong), ("dwMemoryLoad", ctypes.c_ulong), ("ullTotalPhys", ctypes.c_ulonglong),
                        ("ullAvailPhys", ctypes.c_ulonglong), ("a", ctypes.c_ulonglong), ("b", ctypes.c_ulonglong),
                        ("c", ctypes.c_ulonglong), ("d", ctypes.c_ulonglong), ("e", ctypes.c_ulonglong)]
        m = MS(); m.dwLength = ctypes.sizeof(MS)
        if ctypes.windll.kernel32.GlobalMemoryStatusEx(ctypes.byref(m)):
            info["ram_gb"] = round(m.ullTotalPhys / 2 ** 30, 1)
    except Exception:
        pass
    return info


def _allowed_strings():
    import fieldcheck as fc
    words = {"accepted", "would change", "no change", "refused", "error", "ok", "failed", "dialog", "other"}
    words |= {n for n, *_ in fc.OPS} | set(fc.STRUCT) | set(fc.CATEGORIES) | set(fc.KNOWN_DIALOGS)
    return words


_STR_RX = [re.compile(x) for x in (r"^\d{8}-\d{6}$", r"^Doc-\d{1,5}$", r"^\d+(\.\d+){1,3}$", r"^\.(docx|docm|dotx|dotm)$",
                                    r"^(refused|error): [A-Za-z][A-Za-z0-9 \-]{0,60}$", r"^error [A-Z][A-Za-z0-9_]{0,60}$",
                                    r"^[A-Z][A-Za-z0-9_]{0,60}(Error|Exception|Warning|Exit|Interrupt|Fault)$")]


def _clean(v, depth=0):
    """Run reports are written by FormatCheck with counts and categories only. Every string must be one of
    FormatCheck's own words or a fixed pattern (time stamp, version, file type, error type); anything else becomes '?'.
    Dictionary keys must be identifiers or Word field keywords."""
    if depth > 6:
        return None
    if isinstance(v, bool) or v is None or isinstance(v, (int, float)):
        return v
    if isinstance(v, str):
        if v in _clean.allowed or any(r.match(v) for r in _STR_RX):
            if v.startswith(("refused: ", "error: ")) and v.split(": ", 1)[1] not in _clean.allowed and \
                    not re.match(r"^[A-Z][A-Za-z0-9_]{0,60}$", v.split(": ", 1)[1]):
                return "?"
            return v
        return "?"
    if isinstance(v, list):
        return [_clean(x, depth + 1) for x in v[:200]]
    if isinstance(v, dict):
        return {str(k)[:40]: _clean(x, depth + 1) for k, x in list(v.items())[:200]
                if re.match(r"^[a-z][a-z0-9_]{0,40}$", str(k)) or str(k) in _clean.allowed or str(k) in _clean.fields
                or re.match(r"^[a-z0-9_]+\.xml:(new )?[a-zA-Z]{1,30}$", str(k))}
    return None


_clean.allowed = set()
_clean.fields = set()


def run_summaries(limit=30):
    """One compact, content-free summary per FormatCheck run on this PC (newest last)."""
    import fieldcheck as fc
    _clean.allowed = _allowed_strings()
    _clean.fields = set(fc.WORD_FIELDS) | {"EMPTY", "OTHER"}
    out = []
    for p in sorted(glob.glob(os.path.join(runs_dir(), "run_*", "FormatCheck_report.json")))[-limit:]:
        try:
            r = json.load(open(p, encoding="utf-8"))
        except Exception:
            diag.swallowed("feedback")
            continue
        s = {"run": r.get("run"), "version": r.get("version"), "word_available": r.get("word_available"),
             "stopped": r.get("stopped"), "word_link_option_check": r.get("word_link_option_check"), "documents": []}
        for d in r.get("documents", []):
            wk = d.get("word", {}) or {}
            if wk.get("word_build"):
                s["word_build"] = wk["word_build"]
            dd = {"id": d.get("id"), "type": d.get("type"), "load": d.get("load"), "load_s": d.get("load_s"),
                  "original_untouched": d.get("original_untouched"), "structure": d.get("structure"),
                  "dry": d.get("dry"), "recipe": d.get("recipe"),
                  "word": {k: wk.get(k) for k in ("seconds", "timed_out", "blocked_by_dialog", "word_ended", "word_left_running",
                                                 "worker_failed", "worker_error", "worker_code", "links_option_restored")
                           if k in wk}}
            for side in ("before", "after"):
                if side in wk:
                    dd["word"][side] = wk[side]
            s["documents"].append(dd)
        out.append(_clean(s))
    return out


# ------------------------------------------------------------------ REPORT.md
def _events():
    ev = []
    for p in sorted(glob.glob(os.path.join(log_dir(), "events.jsonl*")), reverse=True):
        try:
            ev += [json.loads(l) for l in open(p, encoding="utf-8") if l.strip()]
        except Exception:
            diag.swallowed("feedback")
    return ev


def build_report_md(reports, runs):
    import fieldcheck as fc
    L = ["# FormatCheck feedback export", "",
         f"Exported {datetime.datetime.now():%Y-%m-%d %H:%M}, FormatCheck {fc.VERSION}.",
         "Feature feedback and content-free diagnostics only: no file names, paths or document text.", "",
         f"## Reports ({len(reports)})", ""]
    if not reports:
        L += ["No new reports.", ""]
    for i, (_p, r) in enumerate(reports, 1):
        c = r.get("context", {})
        L += [f"### {i}. [{r['type']}] {r['title']}", f"- Feature: {r['feature']}", f"- Importance: {r['importance']}",
              f"- Written: {c.get('time', '')} (version {c.get('version', '')})",
              f"- Recent actions: {', '.join(c.get('recent_actions', [])) or 'none'}",
              f"- Issues logged that session: {c.get('issues_logged_this_session', 0)}", "", r["text"].strip(), ""]
    ev = _events()
    L += ["## Errors and warnings", ""]
    if not ev:
        L.append("None logged.")
    else:
        groups = {}
        for e in ev:
            k = (e.get("level"), e.get("role", "app"), e.get("area"), e.get("msg"), e.get("type"), (e.get("where") or [""])[-1])
            g = groups.setdefault(k, {"count": 0, "first": e.get("time"), "last": e.get("time"), "versions": set(), "where": e.get("where") or []})
            g["count"] = max(g["count"] + (0 if e.get("repeats") else 1), e.get("repeats") or 0)
            g["last"] = max(g["last"] or "", e.get("time") or "")
            g["first"] = min(g["first"] or "", e.get("time") or "")
            g["versions"].add(e.get("version"))
        L += ["| Level | Process | Area | Message | Error type | Count | First | Last | Versions | Raised at |",
              "|---|---|---|---|---|---|---|---|---|---|"]
        for (lvl, role, area, msg, typ, _w), g in sorted(groups.items(), key=lambda kv: -kv[1]["count"]):
            L.append(f"| {lvl} | {role} | {area} | {msg} | {typ or ''} | {g['count']} | {g['first']} | {g['last']} | "
                     f"{', '.join(sorted(v for v in g['versions'] if v))} | {(g['where'] or [''])[-1]} |")
        L += ["", "Full code locations (call stack, our files only) for each group are in events.jsonl."]
    L += ["", f"## Runs on this PC ({len(runs)}, newest last)", ""]
    if runs:
        L += ["| Run | Version | Docs | Accepted | Refused (why) | Guard trips | Op errors | Fix stopped (step: why) | "
              "Word: ok / dialog / timeout | Re-save drift | Pages before to after | Linked fields | Originals untouched |",
              "|---|---|---|---|---|---|---|---|---|---|---|---|---|"]
        for r in runs:
            docs = r.get("documents") or []
            acc = sum(1 for d in docs if d.get("load") == "accepted")
            ref = [str(d.get("load", "")).split(": ", 1)[-1] for d in docs if str(d.get("load", "")).startswith("refused")]
            trips = sum(1 for d in docs for v in (d.get("dry") or {}).values() if (v or {}).get("why") == "touch elements it did not declare")
            operr = sum(1 for d in docs for v in (d.get("dry") or {}).values() if (v or {}).get("result") == "error")
            stops = [f"{s.get('op')}: {s.get('why')}" for d in docs for s in (d.get("recipe") or []) if not s.get("ok")]
            wok = sum(1 for d in docs if (d.get("word") or {}).get("after", {}).get("open") == "ok")
            wdlg = sum(1 for d in docs if (d.get("word") or {}).get("blocked_by_dialog"))
            wto = sum(1 for d in docs if (d.get("word") or {}).get("timed_out"))
            drift = sorted({f for d in docs for f in ((d.get("word") or {}).get("after", {}) or {}).get("resave_drift", []) or []})
            pages = [f"{(d.get('word') or {}).get('before', {}).get('pages')}>{(d.get('word') or {}).get('after', {}).get('pages')}"
                     for d in docs if (d.get("word") or {}).get("after", {}).get("pages") is not None]
            linked = sum(((d.get("word") or {}).get("before", {}) or {}).get("linked_fields_locked", 0) or 0 for d in docs)
            kept = sum(1 for d in docs if d.get("original_untouched"))
            L.append(f"| {r.get('run')} | {r.get('version')} | {len(docs)} | {acc} | {len(ref)} ({', '.join(sorted(set(ref)))}) | "
                     f"{trips} | {operr} | {'; '.join(stops[:6]) or 'none'} | {wok} / {wdlg} / {wto} | {', '.join(drift) or 'none'} | "
                     f"{', '.join(pages[:8])} | {linked} | {kept} of {len(docs)} |")
        L += ["", "Per document and per operation detail (counts, categories, seconds) is in runs.json."]
    act = []
    for p in sorted(glob.glob(os.path.join(log_dir(), "activity.jsonl*")), reverse=True):
        try:
            act += [json.loads(l) for l in open(p, encoding="utf-8") if l.strip()]
        except Exception:
            diag.swallowed("feedback")
    if act:
        from collections import Counter
        c = Counter(a.get("action") + (":" + a["w"] if a.get("action") == "click" and a.get("w") else "") for a in act)
        L += ["", "## Activity (counts)", "", "| Action | Count |", "|---|---|"] + [f"| {k} | {v} |" for k, v in c.most_common(40)]
        L += ["", "The time-ordered trail (with ms between actions and disabled-control clicks) is in activity.jsonl."]
    return "\n".join(L) + "\n"


def export_files(reports):
    """{name in zip: text}. Text only."""
    runs = run_summaries()
    files = {"REPORT.md": build_report_md(reports, runs),
             "reports.json": json.dumps([r for _p, r in reports], indent=1, ensure_ascii=False),
             "runs.json": json.dumps(runs, indent=1, ensure_ascii=False),
             "system.json": json.dumps(system_info(), indent=1, ensure_ascii=False)}
    for name in ("events.jsonl", "activity.jsonl"):
        parts = []
        for p in sorted(glob.glob(os.path.join(log_dir(), name + "*")), reverse=True):
            try:
                parts.append(open(p, encoding="utf-8").read())
            except Exception:
                diag.swallowed("feedback")
        files[name] = "".join(parts)
    return files


# ------------------------------------------------------------------ machine-file proof
_RX = {"time": r"^\d{4}-\d\d-\d\dT\d\d:\d\d:\d\d$", "version": r"^[0-9A-Za-z.?\-]{1,20}$",
       "level": r"^(error|warning|crash)$", "area": r"^[a-z][a-z0-9_]{0,30}$", "action": r"^[a-z][a-z0-9_]{0,30}$",
       "type": r"^[A-Za-z_][A-Za-z0-9_.]{0,60}$", "ext": r"^\.[a-z0-9]{1,6}$", "code": r"^[a-z][a-z0-9_]{0,40}$",
       "w": r"^[a-z][a-z0-9_.]{0,60}$", "state": r"^disabled$", "role": r"^(app|word_worker)$"}
_WHERE = re.compile(r"^[A-Za-z0-9_.\-]{1,60}\.py:\d{1,6}:[A-Za-z0-9_<>]{1,60}$")


def check_machine_file(name, text):
    """Logs are content-free by construction; this proves it again before export. Returns a list of problems."""
    problems = []
    for n, line in enumerate(text.splitlines(), 1):
        if not line.strip():
            continue
        try:
            e = json.loads(line)
        except ValueError:
            problems.append(f"{name} line {n}: not valid JSON"); continue
        for k, v in e.items():
            if k in _RX:
                if not isinstance(v, str) or not re.match(_RX[k], v):
                    problems.append(f"{name} line {n}: unexpected {k}")
            elif k == "msg":
                if not isinstance(v, str) or len(v) > 160 or re.search(r"[\\/:]", v):
                    problems.append(f"{name} line {n}: unexpected message")
            elif k == "where":
                if not isinstance(v, list) or not all(isinstance(w, str) and _WHERE.match(w) for w in v):
                    problems.append(f"{name} line {n}: unexpected code location")
            elif k in diag.COUNT_KEYS or re.match(r"^ext_[a-z0-9]{1,6}$", k) or k == "repeats":
                if not isinstance(v, int) or isinstance(v, bool):
                    problems.append(f"{name} line {n}: count {k} is not a number")
            else:
                problems.append(f"{name} line {n}: unexpected field {k}")
    return problems


def prepare_export():
    """Everything the zip would hold, and every reason it must not be saved. Nothing is written."""
    reports = [(p, r) for p, r in list_reports() if not r.get("exported")]
    files = export_files(reports)
    g = Guard()
    blocks = []
    for _p, r in reports:
        if g.find(r.get("title", "")) or g.find(r.get("text", "")):
            blocks.append(f"report '{r.get('title', '')[:40]}'")
    for name in ("events.jsonl", "activity.jsonl"):
        if check_machine_file(name, files[name]):
            blocks.append(name)
    if g.find(files["system.json"]):
        blocks.append("system.json")
    # runs.json is built from fixed words and patterns only (_clean); prove it again
    if "\\" in files["runs.json"] or re.search(r"[A-Za-z]:[\\/]", files["runs.json"]):
        blocks.append("runs.json")
    return reports, files, blocks


def write_zip(path, reports, files):
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        for name, text in files.items():
            z.writestr(name, text)
    for p, r in reports:
        r["exported"] = True
        with open(p, "w", encoding="utf-8") as f:
            json.dump(r, f, ensure_ascii=False, indent=1)
    diag.activity("feedback_exported", reports=len(reports))


def default_zip_name():
    import fieldcheck as fc
    return f"FormatCheck_feedback_{fc.VERSION}_{datetime.datetime.now():%Y-%m-%d_%H%M}.zip"
