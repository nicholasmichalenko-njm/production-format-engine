"""Format Engine field check (read-only).

Run it on real proposal volumes. It never writes to the originals: it reads each file once to copy it into
a local work folder, does everything on the copy, and proves at the end that every original is byte
identical. The report has counts only: no document text, no file names (documents are Doc-1, Doc-2...).
The key that maps Doc-N to a file stays in the local work folder and is not part of the report.

What it measures on each copy:
  1. Load: would the engine accept the file, or refuse it, and why (category only)
  2. Size and structure counts (paragraphs, tables, sections, fields by type, styles, lists, Quick Parts)
  3. Each engine operation, in memory: would it change anything, how much, and did the safety guard hold
     (an operation that tries to touch undeclared content on a real document is an engine bug)
  4. The full recipe through the real write path on the copy, step by step
  5. With Word installed: Word opens the copy and the result, re-saves them, reports pages, broken
     references and any structural drift. Each Word job runs in its own worker under a time limit.

Usage: FormatCheck.exe                      opens the FormatCheck window
       FormatCheck.exe <folder> [--no-word] [--keep] [--timeout 900] [--recursive]   runs without the window

Diagnostics: content-free error and activity logs in the logs folder next to the program (diag.py); feedback reports in
the feedback folder; Export for Claude makes one text-only zip (feedback.py). Nothing is sent anywhere.
"""
import argparse
import collections
import datetime
import hashlib
import html
import json
import os
import shutil
import subprocess
import sys
import tempfile
import time
import traceback
import uuid

HERE = os.path.dirname(os.path.abspath(sys.executable if getattr(sys, "frozen", False) else __file__))
sys.path.insert(0, os.path.dirname(HERE))
sys.path.insert(0, HERE)

from engine import safewrite
from engine.pkg import Package, Refused, W, NS
from engine.fingerprint import fingerprint, logical, diff, fields as field_list
from engine.ops import word_ops as ops
from engine.ops import glossary as G
import diag

VERSION = "0.3.0"
LOG_DIR = os.path.join(HERE, "logs")
RUNS_DIR = os.path.join(HERE, "FormatCheck_runs")
FIELD_KIND = __import__("re").compile(r"^[A-Z][A-Z0-9]{0,19}$")
EXTS = (".docx", ".docm", ".dotx", ".dotm")
STRUCT = ["text", "text_len", "deleted_text", "fields", "bookmarks", "bookmark_context", "tables", "drawings",
          "objects", "content_controls", "content_control_props", "revisions", "note_refs", "comment_refs",
          "section_count", "paragraphs", "alt_content", "hyperlinks", "media", "external_links"]

# LINK, INCLUDETEXT, INCLUDEPICTURE, INCLUDE, DDE, DDEAUTO, IMPORT (Word WdFieldType values)
EXTERNAL_FIELD_TYPES = {56, 68, 67, 36, 45, 46, 55}

OPS = [  # name, function, args, allowed content facets
    ("hf_clean", ops.op_hf_clean, {}, ("fields",)),
    ("styles", ops.op_styles, {"spec": {"Caption": {"bold": True}}}, ()),
    ("heading_numbering", ops.op_heading_numbering, {}, ()),
    ("page_numbers", ops.op_page_numbers, {}, ()),
    ("fix_tables", ops.op_fix_tables, {}, ()),
    ("remove_empty_notes", ops.op_remove_empty_notes, {}, ("text", "text_len", "paragraphs", "tables")),
    ("captions", ops.op_captions, {}, ("fields",)),
    ("xref_audit", ops.op_xref_audit, {}, ()),
]


def sha(path):
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for b in iter(lambda: f.read(1 << 20), b""):
            h.update(b)
    return h.hexdigest()


CATEGORIES = ("not a zip", "damaged", "not well-formed", "DOCTYPE", "zip bomb", "Unsafe part name", "encrypted",
              "Strict Open XML", "altChunk", "bound to data", "marked as final", "not a Word", "digital signature",
              "Restrict Editing", "write-protected", "read-only", "open", "touch elements it did not declare",
              "alter content", "Field changes do not match", "Element order problem", "operation stopped",
              "changed on disk", "differs from the checked result", "Untouched part", "snapshot", "could not copy")
KNOWN_DIALOGS = ("Microsoft Word Security Notice", "Microsoft Word", "Password", "File Conversion", "Show Repairs",
                 "Convert File", "Save As", "Open", "Find and Replace", "Microsoft Visual Basic for Applications")
WORD_FIELDS = ("ADDRESSBLOCK ADVANCE ASK AUTHOR AUTONUM AUTONUMLGL AUTONUMOUT AUTOTEXT AUTOTEXTLIST BARCODE BIBLIOGRAPHY "
               "BIDIOUTLINE CITATION COMMENTS COMPARE CREATEDATE DATABASE DATE DISPLAYBARCODE DOCPROPERTY DOCVARIABLE "
               "EDITTIME EQ FILENAME FILESIZE FILLIN FORMCHECKBOX FORMDROPDOWN FORMTEXT GOTOBUTTON GREETINGLINE HYPERLINK "
               "IF INCLUDEPICTURE INCLUDETEXT INDEX INFO KEYWORDS LASTSAVEDBY LINK LISTNUM MACROBUTTON MERGEBARCODE "
               "MERGEFIELD MERGEREC MERGESEQ NEXT NEXTIF NOTEREF NUMCHARS NUMPAGES NUMWORDS PAGE PAGEREF PRINT PRINTDATE "
               "PRIVATE QUOTE RD REF REVNUM SAVEDATE SECTION SECTIONPAGES SEQ SET SKIPIF STYLEREF SUBJECT SYMBOL TA TC "
               "TEMPLATE TIME TITLE TOA TOC USERADDRESS USERINITIALS USERNAME XE DDE DDEAUTO IMPORT INCLUDE EMBED SHAPE "
               "CONTROL DATA GLOSSARY").split()


def category(msg):
    """A refusal or error as a short category, with no document content."""
    m = str(msg)
    for key in ("not a zip", "damaged", "not well-formed", "DOCTYPE", "zip bomb", "Unsafe part name", "encrypted",
                "Strict Open XML", "altChunk", "bound to data", "marked as final", "not a Word", "digital signature",
                "Restrict Editing", "write-protected", "read-only", "open", "touch elements it did not declare",
                "alter content", "Field changes do not match", "Element order problem", "operation stopped",
                "changed on disk", "differs from the checked result", "Untouched part", "snapshot"):
        if key.lower() in m.lower():
            return key
    return "other"


def guard_detail(msg):
    """Which kinds of units a guard refusal named (part name and element name only)."""
    m = str(msg)
    out = collections.Counter()
    for piece in m.split(";"):
        piece = piece.strip()
        if ": " in piece and " outside the declared" in piece:
            part, rest = piece.split(": ", 1)
            part = part.split("declare: ")[-1]
            what = rest.split(" ")[0] if not rest.startswith("new ") else "new " + rest.split(" ")[1]
            out[f"{os.path.basename(part)}:{what}"] += 1
    return dict(out)


def structure(pkg):
    root = pkg.tree(pkg.main)
    s = {"parts": len(pkg.raw), "mb": round(sum(len(v) for v in pkg.raw.values()) / 1e6, 2),
         "main_xml_mb": round(len(pkg.raw[pkg.main]) / 1e6, 2)}
    s["paragraphs"] = sum(1 for _ in root.iter(f"{{{W}}}p"))
    tbls = [t for t in root.iter(f"{{{W}}}tbl")]
    s["tables"] = len(tbls)
    s["nested_tables"] = sum(1 for t in tbls if any(True for _ in t.iterancestors(f"{{{W}}}tbl")))
    s["figure_tables"] = sum(1 for t in tbls if ops.is_figure_table(t))
    s["merged_cells"] = sum(1 for _ in root.iter(f"{{{W}}}vMerge")) + sum(1 for _ in root.iter(f"{{{W}}}gridSpan"))
    s["sections"] = sum(1 for _ in root.iter(f"{{{W}}}sectPr"))
    kinds = collections.Counter()
    for story in pkg.stories():
        for ins in field_list(pkg.tree(story)):
            k = (ins.split(" ")[0] if ins else "").upper() or "EMPTY"
            kinds[k if k in WORD_FIELDS or k == "EMPTY" else "OTHER"] += 1  # Word's own field keywords only
    s["fields"] = dict(kinds.most_common(15))
    s["bookmarks"] = sum(1 for _ in root.iter(f"{{{W}}}bookmarkStart"))
    s["tracked_changes"] = sum(1 for _ in root.iter(f"{{{W}}}ins")) + sum(1 for _ in root.iter(f"{{{W}}}del"))
    s["content_controls"] = sum(1 for _ in root.iter(f"{{{W}}}sdt"))
    s["drawings"] = sum(1 for _ in root.iter(f"{{{W}}}drawing")) + sum(1 for _ in root.iter(f"{{{W}}}pict"))
    s["text_boxes"] = sum(1 for _ in root.iter(f"{{{W}}}txbxContent"))
    s["footnotes_refs"] = sum(1 for _ in root.iter(f"{{{W}}}footnoteReference"))
    s["headers_footers"] = len([p for p in pkg.stories() if "/header" in p or "/footer" in p])
    sp = pkg.word_part("/styles")
    if sp:
        st = pkg.tree(sp).findall("w:style", NS)
        s["styles_defined"] = len(st)
        s["styles_custom"] = sum(1 for x in st if x.get(f"{{{W}}}customStyle") in ("1", "true"))
        s["styles_used"] = len(ops.styles_in_use(pkg))
    nb = pkg.word_part("/numbering")
    if nb:
        s["list_definitions"] = len(pkg.tree(nb).findall("w:abstractNum", NS))
        s["list_instances"] = len(pkg.tree(nb).findall("w:num", NS))
    s["quick_parts"] = len(G.entries(pkg))
    s["macros"] = pkg.has_macros
    return s


def dry_ops(path):
    """Each operation alone, in memory, on a fresh load: change size, guard result, time. Nothing is written."""
    out = {}
    for name, fn, args, allowed in OPS:
        t = time.perf_counter()
        rec = {}
        try:
            r = safewrite.apply(path, fn, allowed=allowed, op_args=args, dry_run=True)
            rep = r.report
            rec["result"] = "would change" if r.changed else "no change"
            rec["dirty_parts"] = len(getattr(r, "dirty", []) or [])
            for k in ("tables", "figures_skipped", "fixed", "rows_removed", "refs"):
                if k in rep:
                    rec[k] = rep[k]
            if "roles" in rep:
                rec["roles"] = rep["roles"]
            if "rewritten" in rep:
                rec["fields_rewritten"] = len(rep["rewritten"])
            if "missing" in rep:
                rec["broken_refs"] = len(rep["missing"])
            if "kinds" in rep:
                rec["section_kinds"] = dict(collections.Counter(rep["kinds"]))
            if "skipped" in rep:
                rec["skipped"] = dict(collections.Counter(str(x).split(":")[0] for x in rep["skipped"]))
        except Refused as e:
            rec["result"] = "refused"
            rec["why"] = category(e)
            if rec["why"] == "touch elements it did not declare":
                diag.warning("guard", "Guard refused an operation in memory", exc=True)
            elif rec["why"] == "other":
                diag.warning("engine", "Operation refused for an unlisted reason", exc=True)
            if rec["why"] == "touch elements it did not declare":
                rec["guard"] = guard_detail(e)
        except Exception as e:
            rec["result"] = "error"
            rec["why"] = type(e).__name__
            diag.error("engine", "Operation raised an error in memory", exc=True)
        rec["seconds"] = round(time.perf_counter() - t, 2)
        out[name] = rec
    return out


def recipe(path):
    steps = []
    for name, fn, args, allowed in OPS:
        t = time.perf_counter()
        try:
            r = safewrite.apply(path, fn, allowed=allowed, op_args=args)
            steps.append({"op": name, "ok": True, "changed": r.changed, "seconds": round(time.perf_counter() - t, 2)})
        except Refused as e:
            steps.append({"op": name, "ok": False, "why": category(e), "seconds": round(time.perf_counter() - t, 2)})
            if category(e) not in ("open", "read-only", "write-protected"):
                diag.warning("recipe", "Recipe step refused", exc=True)
        except Exception as e:
            steps.append({"op": name, "ok": False, "why": "error " + type(e).__name__, "seconds": round(time.perf_counter() - t, 2)})
            diag.error("recipe", "Recipe step raised an error", exc=True)
    return steps


# ---------------------------------------------------------------- Word worker (its own process)
def word_worker(before, after, out_json):
    """Runs in a separate process. Opens both copies in a new hidden Word, macros off, never the user's Word."""
    import pythoncom
    import win32com.client
    diag.setup(LOG_DIR, VERSION, role="word_worker", crash_capture=False)
    res = {}
    pythoncom.CoInitialize()
    pre = winword_pids()
    app = win32com.client.DispatchEx("Word.Application")
    mine = [p for p in winword_pids() - pre if is_automation_word(p)]
    if len(mine) == 1:
        with open(out_json + ".pid", "w") as fh:
            fh.write(str(mine[0]))
    try:
        res["word_build"] = app.Build
        app.Visible = False
        app.DisplayAlerts = 0
        app.ScreenUpdating = False
        # Word shows a modal security notice for any document with fields that read other files when
        # "Update automatic links at open" is on, and a hidden Word then waits forever. The option is turned
        # off in this automation Word only and set back before Quit (tested: a forced stop never saves it,
        # a normal Quit saves whatever value is set, so it must be set back first).
        res["update_links_at_open_original"] = bool(app.Options.UpdateLinksAtOpen)
        app.Options.UpdateLinksAtOpen = False
        for label, p in (("before", before), ("after", after)):
            if not p or not os.path.exists(p):
                continue
            app.AutomationSecurity = 3
            t = time.perf_counter()
            try:
                doc = app.Documents.Open(os.path.abspath(p), False, False, False, "", "", True, "", "", 0, 0, False, False, 0, True)
            except Exception as e:
                res[label] = {"open": "failed", "why": type(e).__name__, "code": com_code(e)}
                diag.error("word", "Word could not open a copy", exc=True, **(com_code(e) or {}))
                continue
            r = {"open": "ok", "open_s": round(time.perf_counter() - t, 1)}
            resaved = p + ".word.docx"
            doc.SaveAs2(os.path.abspath(resaved), 16)
            r["resave_drift"] = [k.split("|", 1)[1] for k in diff(logical(Package(p)), logical(Package(resaved)))
                                 if k.split("|", 1)[1] in STRUCT]
            t = time.perf_counter()
            linked = 0
            for story in doc.StoryRanges:
                s = story
                while s is not None:
                    # fields that pull from other files are locked first: updating them reaches outside the PC
                    # and, when the source is not reachable, Word replaces the linked table with an error
                    for f in s.Fields:
                        if f.Type in EXTERNAL_FIELD_TYPES and not f.Locked:
                            f.Locked = True
                            linked += 1
                    s.Fields.Update()
                    try:
                        s = s.NextStoryRange
                    except Exception:
                        s = None
            for i in range(1, doc.TablesOfContents.Count + 1):
                doc.TablesOfContents(i).Update()
            for i in range(1, doc.TablesOfFigures.Count + 1):
                doc.TablesOfFigures(i).Update()
            r["update_s"] = round(time.perf_counter() - t, 1)
            r["linked_fields_locked"] = linked
            r["pages"] = doc.ComputeStatistics(2)
            errs = 0
            for f in doc.Fields:
                try:
                    if "Error!" in (f.Result.Text or ""):
                        errs += 1
                except Exception:
                    diag.swallowed("word")
            r["fields_with_error_after_update"] = errs
            doc.Close(0)
            os.remove(resaved)
            res[label] = r
    except Exception as e:
        res["worker_error"] = type(e).__name__
        res["worker_code"] = com_code(e)
        diag.error("word", "Word check stopped with an error", exc=True, **(com_code(e) or {}))
    finally:
        restored = False
        try:
            if "update_links_at_open_original" in res:
                app.Options.UpdateLinksAtOpen = res["update_links_at_open_original"]
                restored = bool(app.Options.UpdateLinksAtOpen) == res["update_links_at_open_original"]
        except Exception:
            diag.error("word", "Could not set the link option back", exc=True)
        res["links_option_restored"] = restored
        try:
            app.NormalTemplate.Saved = True
        except Exception:
            pass
        if restored or "update_links_at_open_original" not in res:
            try:
                app.Quit(0)
            except Exception:
                pass
        # if it could not be set back, the Word is left to the watchdog: a forced stop does not save options
    diag.shutdown()
    json.dump(res, open(out_json, "w"))


def option_check(value, out_path):
    """End of run: a fresh automation Word confirms the user's link setting is the value it had at the start."""
    import pythoncom
    import win32com.client
    pythoncom.CoInitialize()
    app = win32com.client.DispatchEx("Word.Application")
    r = {}
    try:
        now = bool(app.Options.UpdateLinksAtOpen)
        r["was"] = now
        if now != value:
            app.Options.UpdateLinksAtOpen = value
            r["set_back"] = True
        r["ok"] = bool(app.Options.UpdateLinksAtOpen) == value
    except Exception:
        diag.setup(LOG_DIR, VERSION, role="word_worker", crash_capture=False)
        diag.error("word", "Link option check failed", exc=True)
        r["ok"] = None
    finally:
        try:
            app.NormalTemplate.Saved = True
            app.Quit(0)
        except Exception:
            pass
    json.dump(r, open(out_path, "w"))


def com_code(e):
    """HRESULT and Word's own error number only. Word's message is never kept: it can name the document."""
    try:
        h = getattr(e, "hresult", None)
        info = getattr(e, "excepinfo", None) or ()
        sc = info[5] if len(info) > 5 else None
        return {k: v for k, v in (("hresult", h), ("scode", sc)) if isinstance(v, int)}
    except Exception:
        return None


def winword_pids():
    try:
        out = subprocess.run(["tasklist", "/FI", "IMAGENAME eq WINWORD.EXE", "/FO", "CSV", "/NH"],
                             capture_output=True, text=True, creationflags=0x08000000).stdout
        return {int(line.split('","')[1]) for line in out.splitlines() if line.startswith('"WINWORD')}
    except Exception:
        diag.swallowed("word")
        return set()


def is_automation_word(pid):
    """True when the process is a Word started for automation (command line has /Automation or -Embedding),
    never a Word a person opened."""
    try:
        out = subprocess.run(["powershell", "-NoProfile", "-Command",
                              f"(Get-CimInstance Win32_Process -Filter 'ProcessId={int(pid)}').CommandLine"],
                             capture_output=True, text=True, creationflags=0x08000000, timeout=30).stdout
        return "WINWORD" in out.upper() and ("/AUTOMATION" in out.upper() or "-EMBEDDING" in out.upper())
    except Exception:
        diag.swallowed("word")
        return False


def visible_dialogs(pid):
    """Titles of visible dialog windows owned by the process (a blocked hidden Word shows one).
    A title that could hold a file name is reported only as 'dialog'."""
    try:
        import win32gui, win32process
    except Exception:
        return []
    found = []

    def cb(h, _):
        try:
            if win32gui.IsWindowVisible(h) and win32process.GetWindowThreadProcessId(h)[1] == pid \
                    and win32gui.GetClassName(h) in ("NUIDialog", "#32770", "bosa_sdm_msword"):
                t = win32gui.GetWindowText(h) or "dialog"
                found.append(t if t in KNOWN_DIALOGS else "dialog")  # an unknown title could name a document
        except Exception:
            pass
        return True
    try:
        win32gui.EnumWindows(cb, None)
    except Exception:
        diag.swallowed("word")
    return found


def end_word(pid_file):
    """End only the Word this job started, and only if it is still that automation Word."""
    try:
        pid = int(open(pid_file).read().strip())
    except Exception:
        return False
    if pid in winword_pids() and is_automation_word(pid):
        subprocess.run(["taskkill", "/PID", str(pid), "/F"], capture_output=True, creationflags=0x08000000)
        return True
    return False


def run_word(before, after, timeout):
    out_json = after + ".wordcheck.json"
    pid_file = out_json + ".pid"
    exe = [sys.executable] if getattr(sys, "frozen", False) else [sys.executable, os.path.abspath(__file__)]
    t = time.perf_counter()
    proc = subprocess.Popen(exe + ["--word-worker", before, after, out_json], creationflags=0x08000000)
    diag.activity("word_job_started")
    stop = None
    while proc.poll() is None:
        if time.perf_counter() - t > timeout:
            stop = {"timed_out": True}
            break
        pid = None
        try:
            pid = int(open(pid_file).read().strip())
        except Exception:
            pass
        if pid:
            d = visible_dialogs(pid)
            if d:
                stop = {"blocked_by_dialog": d[0]}
                break
        time.sleep(2)
    if stop is not None:
        proc.kill()
        stop["word_ended"] = end_word(pid_file)
        if not stop["word_ended"]:
            stop["word_left_running"] = True  # could not be identified for sure: never end a Word we did not start
        stop["seconds"] = round(time.perf_counter() - t, 1)
        for f in (pid_file, out_json):
            if os.path.exists(f):
                os.remove(f)
        return stop
    if not os.path.exists(out_json):
        if os.path.exists(pid_file):
            os.remove(pid_file)
        return {"worker_failed": True}
    r = json.load(open(out_json))
    os.remove(out_json)
    if "update_links_at_open_original" in r and not r.get("links_option_restored"):
        r["word_ended_unsaved"] = end_word(pid_file)  # a forced stop never saves the changed option
    if os.path.exists(pid_file):
        os.remove(pid_file)
    r["seconds"] = round(time.perf_counter() - t, 1)
    return r


# ---------------------------------------------------------------- report
def write_report(run_dir, data):
    jpath = os.path.join(run_dir, "FormatCheck_report.json")
    json.dump(data, open(jpath, "w"), indent=1)
    rows = []
    for d in data["documents"]:
        dry = d.get("dry", {})
        guard_trips = sum(1 for v in dry.values() if v.get("why") == "touch elements it did not declare")
        errors = sum(1 for v in dry.values() if v.get("result") == "error")
        steps = d.get("recipe", [])
        rec_ok = all(s["ok"] for s in steps) if steps else None
        wk = d.get("word", {})
        after = wk.get("after", {})
        rows.append("<tr>" + "".join(f"<td>{html.escape(str(x))}</td>" for x in [
            d["id"], d.get("load", ""), d.get("structure", {}).get("paragraphs", ""), d.get("structure", {}).get("tables", ""),
            d.get("structure", {}).get("sections", ""), guard_trips, errors,
            "" if rec_ok is None else ("all steps" if rec_ok else "stopped: " + ", ".join(f'{s["op"]} ({s["why"]})' for s in steps if not s["ok"])),
            after.get("open", (wk.get("blocked_by_dialog") and "blocked: " + wk["blocked_by_dialog"]) or (wk.get("timed_out") and "timed out") or ""), ", ".join(after.get("resave_drift", [])) or ("none" if after else ""),
            wk.get("before", {}).get("pages", ""), after.get("pages", ""),
            wk.get("before", {}).get("fields_with_error_after_update", ""), after.get("fields_with_error_after_update", ""),
            wk.get("before", {}).get("linked_fields_locked", ""),
            "yes" if d.get("original_untouched") else "NO"]) + "</tr>")
    oc = data.get("word_link_option_check")
    option_line = ("" if oc is None else
                   "Word setting 'Update automatic links at open': unchanged at the end of the run." if oc.get("ok") and not oc.get("set_back") else
                   "Word setting 'Update automatic links at open' was found changed and was set back." if oc.get("ok") else
                   "Word setting 'Update automatic links at open' could not be confirmed: check it in Word Options, Advanced, General.")
    head = ["Doc", "Load", "Paragraphs", "Tables", "Sections", "Guard trips", "Op errors", "Recipe on copy",
            "Word opens result", "Word re-save drift", "Pages before", "Pages after", "Broken refs before", "Broken refs after",
            "Linked fields (not updated)",
            "Original untouched"]
    page = f"""<!doctype html><html><head><meta charset="utf-8"><title>FormatCheck report</title>
<style>body{{font:14px Segoe UI,system-ui,sans-serif;margin:24px;color:#14202e}}table{{border-collapse:collapse}}
td,th{{border:1px solid #ccd;padding:6px 8px;text-align:left;vertical-align:top}}th{{background:#eef1f5}}</style></head><body>
<h1>FormatCheck report</h1>
<p>Version {VERSION}. Run {html.escape(data['run'])}. {len(data['documents'])} documents. Counts only: no document text and no file names.
The key from Doc-N to file stays on this PC in the work folder.</p>
<p>{html.escape(option_line)}</p>
<p>Originals untouched: <b>{sum(1 for d in data['documents'] if d.get('original_untouched'))} of {len(data['documents'])}</b>.
Word available: {data.get('word_available')}.</p>
<table><tr>{''.join(f'<th>{h}</th>' for h in head)}</tr>{''.join(rows)}</table>
<p>Full detail per document and per operation is in FormatCheck_report.json next to this page.</p></body></html>"""
    hpath = os.path.join(run_dir, "FormatCheck_report.html")
    open(hpath, "w", encoding="utf-8").write(page)
    return jpath, hpath


def find_files(folder, recursive):
    files = []
    for dp, dn, fn in os.walk(folder):
        for f in sorted(fn):
            if f.lower().endswith(EXTS) and not f.startswith(("~$", "~pe_tmp_")):
                files.append(os.path.join(dp, f))
        if not recursive:
            break
    return files


def _exe():
    return [sys.executable] if getattr(sys, "frozen", False) else [sys.executable, os.path.abspath(__file__)]


def run_check(folder, recursive=False, use_word=True, timeout=1200, keep=False, work=None,
              progress=print, should_stop=lambda: False, on_doc=None):
    """Check every Word file in folder on local copies. progress(text) gets status lines with Doc-N only, never names.
    should_stop() is asked between documents. on_doc(i, n) after each document. Returns (data, html_path)."""
    files = find_files(folder, recursive)
    stamp = datetime.datetime.now().strftime("%Y%m%d-%H%M%S")
    # next to the program, not in AppData: packaged apps can redirect AppData so Word would not see the copies
    base = os.path.join(work or RUNS_DIR, f"run_{stamp}")
    os.makedirs(base, exist_ok=True)
    keymap = open(os.path.join(base, "KEY_stays_on_this_PC.txt"), "w", encoding="utf-8")
    links_option = None
    word_ok = use_word and sys.platform == "win32"
    data = {"version": VERSION, "run": stamp, "word_available": word_ok, "documents": []}
    exts = collections.Counter(os.path.splitext(f)[1].lower() for f in files)
    diag.run_started(docs=len(files), word=int(word_ok), recursive=int(recursive), timeout=timeout,
                     **{f"ext_{e.strip('.')}": c for e, c in exts.items()})
    progress(f"FormatCheck {VERSION}: {len(files)} documents.")
    stopped = False
    for i, src in enumerate(files, 1):
        if should_stop():
            stopped = True
            progress("Stopped before the next document.")
            break
        did = f"Doc-{i}"
        t_doc = time.perf_counter()
        ext = os.path.splitext(src)[1].lower()
        keymap.write(f"{did}\t{src}\n"); keymap.flush()
        d = {"id": did, "type": ext}
        progress(f"{did} of {len(files)}: copying")
        try:
            h0 = sha(src)
            work_copy = os.path.join(base, did + ext)
            shutil.copyfile(src, work_copy)  # copyfile: data only, no attributes such as read-only
        except Exception:
            d["load"] = "refused: could not copy"
            diag.error("copy", "Could not read or copy a document", exc=True)
            data["documents"].append(d); write_report(base, data)
            continue
        try:
            t = time.perf_counter()
            pkg = Package(work_copy)
            d["load"] = "accepted"
            d["load_s"] = round(time.perf_counter() - t, 2)
            d["structure"] = structure(pkg)
            del pkg
        except Refused as e:
            d["load"] = "refused: " + category(e)
        except Exception as e:
            d["load"] = "error: " + type(e).__name__
            diag.error("load", "Loading a document raised an error", exc=True)
        if d["load"] == "accepted":
            progress(f"{did}: operations in memory")
            d["dry"] = dry_ops(work_copy)
            before_copy = os.path.join(base, did + "_original" + ext)
            shutil.copyfile(work_copy, before_copy)
            progress(f"{did}: full fix on the copy")
            d["recipe"] = recipe(work_copy)
            if word_ok and ext in (".docx", ".docm"):
                progress(f"{did}: Word check (up to {timeout} s)")
                # a fresh name for every Word job: after a forced stop Word remembers the file name and asks
                # "it caused a serious error, open it anyway?" the next time a file of that name opens
                tag = uuid.uuid4().hex[:8]
                wb, wa = os.path.join(base, f"{did}_{tag}_b{ext}"), os.path.join(base, f"{did}_{tag}_a{ext}")
                shutil.copyfile(before_copy, wb); shutil.copyfile(work_copy, wa)
                try:
                    d["word"] = run_word(wb, wa, timeout)
                except Exception as e:
                    d["word"] = {"worker_failed": True, "why": type(e).__name__}
                    diag.error("word", "Starting the Word check failed", exc=True)
                for f in (wb, wa):
                    try:
                        if os.path.exists(f):
                            os.remove(f)
                    except Exception:
                        diag.swallowed("cleanup")
                wk = d["word"]
                if wk.get("blocked_by_dialog"):
                    diag.warning("word", "Word showed a dialog and the check ended", code=_code(wk["blocked_by_dialog"]))
                if wk.get("timed_out"):
                    diag.warning("word", "Word check ran out of time", seconds=int(wk.get("seconds", 0)))
                if wk.get("worker_failed"):
                    diag.error("word", "Word check gave no result")
                if wk.get("word_left_running"):
                    diag.warning("word", "A Word could not be identified and was left running")
                if "update_links_at_open_original" in wk and links_option is None:
                    links_option = wk["update_links_at_open_original"]
        try:
            d["original_untouched"] = sha(src) == h0
        except Exception:
            d["original_untouched"] = None
            diag.error("copy", "Could not re-read an original", exc=True)
        if d.get("original_untouched") is False:
            diag.error("copy", "An original changed during the run")
        data["documents"].append(d)
        write_report(base, data)
        _doc_activity(d, time.perf_counter() - t_doc)
        if on_doc:
            on_doc(i, len(files))
    keymap.close()
    if links_option is not None:
        out_path = os.path.join(base, "option_check.json")
        try:
            subprocess.run(_exe() + ["--option-check", "1" if links_option else "0", out_path], timeout=120,
                           creationflags=0x08000000)
            data["word_link_option_check"] = json.load(open(out_path))
            os.remove(out_path)
        except Exception as e:
            data["word_link_option_check"] = {"ok": None, "why": type(e).__name__}
            diag.error("word", "Link option check did not run", exc=True)
        if not (data["word_link_option_check"] or {}).get("ok"):
            diag.error("word", "Link option could not be confirmed")
    if not keep:
        for f in os.listdir(base):
            p = os.path.join(base, f)
            try:
                if f.startswith("Doc-") or f == safewrite.SNAP_DIR:
                    shutil.rmtree(p) if os.path.isdir(p) else os.remove(p)
            except Exception:
                diag.swallowed("cleanup")
    data["stopped"] = stopped
    j, h = write_report(base, data)
    docs = data["documents"]
    diag.run_finished(stopped=stopped, docs=len(docs), accepted=sum(1 for d in docs if d.get("load") == "accepted"),
                      refused=sum(1 for d in docs if str(d.get("load", "")).startswith("refused")),
                      kept=sum(1 for d in docs if d.get("original_untouched")))
    progress(f"Done. {len(docs)} documents checked. Report: {os.path.basename(h)}")
    return data, h


def _code(title):
    """A dialog title as a short code for the log ('microsoft_word_security_notice'), or 'dialog'."""
    import re as _re
    c = _re.sub(r"[^a-z0-9]+", "_", (title or "").lower()).strip("_")[:40]
    return c if _re.match(r"^[a-z][a-z0-9_]{0,40}$", c or "") and c not in ("",) else "dialog"


def _doc_activity(d, secs):
    dry = d.get("dry", {})
    steps = d.get("recipe", [])
    wk = d.get("word", {})
    st = d.get("structure", {})
    load = str(d.get("load", ""))
    counts = {"seconds": int(secs), "accepted": int(load == "accepted"), "refused": int(load.startswith("refused")),
              "paragraphs": st.get("paragraphs", 0) if isinstance(st.get("paragraphs"), int) else 0,
              "tables": st.get("tables", 0) if isinstance(st.get("tables"), int) else 0,
              "guard_trips": sum(1 for v in dry.values() if v.get("why") == "touch elements it did not declare"),
              "op_errors": sum(1 for v in dry.values() if v.get("result") == "error"),
              "failed": sum(1 for s_ in steps if not s_.get("ok"))}
    if wk:
        b, a = wk.get("before", {}), wk.get("after", {})
        counts.update({"word": 1, "blocked": int(bool(wk.get("blocked_by_dialog"))), "timed_out": int(bool(wk.get("timed_out"))),
                       "drift": len(set(a.get("resave_drift", []) or [])), "linked": b.get("linked_fields_locked", 0) or 0})
        if isinstance(b.get("pages"), int):
            counts["pages"] = b["pages"]
        if isinstance(a.get("pages"), int):
            counts["pages_after"] = a["pages"]
    if load.startswith("refused"):
        counts["code"] = _code(load.split(": ", 1)[-1])
    diag.activity("doc_done", ext=d.get("type"), **counts)


def main():
    if len(sys.argv) > 1 and sys.argv[1] == "--word-worker":
        word_worker(sys.argv[2], sys.argv[3], sys.argv[4])
        return
    if len(sys.argv) > 1 and sys.argv[1] == "--option-check":
        option_check(sys.argv[2] == "1", sys.argv[3])
        return
    diag.setup(LOG_DIR, VERSION)
    try:
        if len(sys.argv) == 1:
            sys.modules.setdefault("fieldcheck", sys.modules[__name__])  # the window imports this module by name
            import gui
            gui.main()
            return
        ap = argparse.ArgumentParser()
        ap.add_argument("folder")
        ap.add_argument("--no-word", action="store_true")
        ap.add_argument("--keep", action="store_true", help="keep the local copies after the run")
        ap.add_argument("--timeout", type=int, default=1200, help="seconds per Word job")
        ap.add_argument("--recursive", action="store_true")
        ap.add_argument("--work", help="folder for the local copies and the report (default: FormatCheck_runs next to the program)")
        ap.add_argument("--no-open", action="store_true", help="do not open the report at the end")
        a = ap.parse_args()
        if not os.path.isdir(a.folder):
            print("Folder not found."); return
        diag.activity("command_line_run")
        data, h = run_check(a.folder, a.recursive, not a.no_word, a.timeout, a.keep, a.work)
        print("Share FormatCheck_report.json or .html if you want: they hold counts only. Keep KEY_stays_on_this_PC.txt local.")
        if not a.no_open:
            try:
                os.startfile(h)
            except Exception:
                diag.swallowed("report")
    finally:
        diag.shutdown()


if __name__ == "__main__":
    main()
