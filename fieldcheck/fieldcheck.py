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

Usage: FormatCheck.exe [folder] [--no-word] [--keep] [--timeout 900]
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

HERE = os.path.dirname(os.path.abspath(sys.executable if getattr(sys, "frozen", False) else __file__))
sys.path.insert(0, os.path.dirname(HERE))
sys.path.insert(0, HERE)

from engine import safewrite
from engine.pkg import Package, Refused, W, NS
from engine.fingerprint import fingerprint, diff, fields as field_list
from engine.ops import word_ops as ops
from engine.ops import glossary as G

VERSION = "0.1.0"
EXTS = (".docx", ".docm", ".dotx", ".dotm")
STRUCT = ["text", "text_len", "deleted_text", "fields", "bookmarks", "bookmark_context", "tables", "drawings",
          "objects", "content_controls", "content_control_props", "revisions", "note_refs", "comment_refs",
          "section_count", "paragraphs", "alt_content", "hyperlinks", "media", "external_links"]

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
            kinds[(ins.split(" ")[0] if ins else "").upper() or "EMPTY"] += 1
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
                rec["guard"] = guard_detail(e)
        except Exception as e:
            rec["result"] = "error"
            rec["why"] = type(e).__name__
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
        except Exception as e:
            steps.append({"op": name, "ok": False, "why": "error " + type(e).__name__, "seconds": round(time.perf_counter() - t, 2)})
    return steps


# ---------------------------------------------------------------- Word worker (its own process)
def word_worker(before, after, out_json):
    """Runs in a separate process. Opens both copies in a new hidden Word, macros off, never the user's Word."""
    import pythoncom
    import win32com.client
    res = {}
    pythoncom.CoInitialize()
    app = win32com.client.DispatchEx("Word.Application")
    try:
        res["word_build"] = app.Build
        app.Visible = False
        app.DisplayAlerts = 0
        app.ScreenUpdating = False
        for label, p in (("before", before), ("after", after)):
            if not p or not os.path.exists(p):
                continue
            app.AutomationSecurity = 3
            t = time.perf_counter()
            try:
                doc = app.Documents.Open(os.path.abspath(p), False, False, False, "", "", True, "", "", 0, 0, False, False, 0, True)
            except Exception as e:
                res[label] = {"open": "failed", "why": type(e).__name__, "code": com_code(e)}
                continue
            r = {"open": "ok", "open_s": round(time.perf_counter() - t, 1)}
            resaved = p + ".word.docx"
            doc.SaveAs2(os.path.abspath(resaved), 16)
            r["resave_drift"] = [k.split("|", 1)[1] for k in diff(fingerprint(Package(p)), fingerprint(Package(resaved)))
                                 if k.split("|", 1)[1] in STRUCT]
            t = time.perf_counter()
            for story in doc.StoryRanges:
                s = story
                while s is not None:
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
            r["pages"] = doc.ComputeStatistics(2)
            errs = 0
            for f in doc.Fields:
                try:
                    if "Error!" in (f.Result.Text or ""):
                        errs += 1
                except Exception:
                    pass
            r["fields_with_error_after_update"] = errs
            doc.Close(0)
            os.remove(resaved)
            res[label] = r
    except Exception as e:
        res["worker_error"] = type(e).__name__
        res["worker_code"] = com_code(e)
    finally:
        try:
            app.NormalTemplate.Saved = True
        except Exception:
            pass
        try:
            app.Quit(0)
        except Exception:
            pass
    json.dump(res, open(out_json, "w"))


def com_code(e):
    """HRESULT and Word's own error number and message, with any file path removed."""
    try:
        h = getattr(e, "hresult", None)
        info = getattr(e, "excepinfo", None) or ()
        desc = str(info[2]) if len(info) > 2 and info[2] else ""
        desc = " ".join(w for w in desc.split() if "\\" not in w and "/" not in w and ":" not in w)[:160]
        return {"hresult": h, "scode": info[5] if len(info) > 5 else None, "message": desc}
    except Exception:
        return None


def winword_pids():
    try:
        out = subprocess.run(["tasklist", "/FI", "IMAGENAME eq WINWORD.EXE", "/FO", "CSV", "/NH"],
                             capture_output=True, text=True, creationflags=0x08000000).stdout
        return {int(line.split('","')[1]) for line in out.splitlines() if line.startswith('"WINWORD')}
    except Exception:
        return set()


def run_word(before, after, timeout):
    out_json = after + ".wordcheck.json"
    exe = [sys.executable] if getattr(sys, "frozen", False) else [sys.executable, os.path.abspath(__file__)]
    pre = winword_pids()
    t = time.perf_counter()
    proc = subprocess.Popen(exe + ["--word-worker", before, after, out_json], creationflags=0x08000000)
    try:
        proc.wait(timeout=timeout)
    except subprocess.TimeoutExpired:
        proc.kill()
        mine = winword_pids() - pre  # only a Word that started during this job
        for pid in mine:
            subprocess.run(["taskkill", "/PID", str(pid), "/F"], capture_output=True, creationflags=0x08000000)
        return {"timed_out": True, "seconds": round(time.perf_counter() - t, 1), "word_processes_ended": len(mine)}
    if not os.path.exists(out_json):
        return {"worker_failed": True}
    r = json.load(open(out_json))
    os.remove(out_json)
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
            after.get("open", wk.get("timed_out") and "timed out" or ""), ", ".join(after.get("resave_drift", [])) or ("none" if after else ""),
            wk.get("before", {}).get("pages", ""), after.get("pages", ""),
            wk.get("before", {}).get("fields_with_error_after_update", ""), after.get("fields_with_error_after_update", ""),
            "yes" if d.get("original_untouched") else "NO"]) + "</tr>")
    head = ["Doc", "Load", "Paragraphs", "Tables", "Sections", "Guard trips", "Op errors", "Recipe on copy",
            "Word opens result", "Word re-save drift", "Pages before", "Pages after", "Broken refs before", "Broken refs after",
            "Original untouched"]
    page = f"""<!doctype html><html><head><meta charset="utf-8"><title>FormatCheck report</title>
<style>body{{font:14px Segoe UI,system-ui,sans-serif;margin:24px;color:#14202e}}table{{border-collapse:collapse}}
td,th{{border:1px solid #ccd;padding:6px 8px;text-align:left;vertical-align:top}}th{{background:#eef1f5}}</style></head><body>
<h1>FormatCheck report</h1>
<p>Version {VERSION}. Run {html.escape(data['run'])}. {len(data['documents'])} documents. Counts only: no document text and no file names.
The key from Doc-N to file stays on this PC in the work folder.</p>
<p>Originals untouched: <b>{sum(1 for d in data['documents'] if d.get('original_untouched'))} of {len(data['documents'])}</b>.
Word available: {data.get('word_available')}.</p>
<table><tr>{''.join(f'<th>{h}</th>' for h in head)}</tr>{''.join(rows)}</table>
<p>Full detail per document and per operation is in FormatCheck_report.json next to this page.</p></body></html>"""
    hpath = os.path.join(run_dir, "FormatCheck_report.html")
    open(hpath, "w", encoding="utf-8").write(page)
    return jpath, hpath


def pick_folder():
    try:
        import tkinter
        from tkinter import filedialog
        root = tkinter.Tk(); root.withdraw()
        d = filedialog.askdirectory(title="FormatCheck: pick the folder with the volumes to check (read-only)")
        root.destroy()
        return d
    except Exception:
        return None


def main():
    if len(sys.argv) > 1 and sys.argv[1] == "--word-worker":
        word_worker(sys.argv[2], sys.argv[3], sys.argv[4])
        return
    ap = argparse.ArgumentParser()
    ap.add_argument("folder", nargs="?")
    ap.add_argument("--no-word", action="store_true")
    ap.add_argument("--keep", action="store_true", help="keep the local copies after the run")
    ap.add_argument("--timeout", type=int, default=1200, help="seconds per Word job")
    ap.add_argument("--recursive", action="store_true")
    ap.add_argument("--work", help="folder for the local copies and the report (default: FormatCheck_runs next to the program)")
    a = ap.parse_args()
    folder = a.folder or pick_folder()
    if not folder or not os.path.isdir(folder):
        print("No folder selected."); return
    files = []
    for dp, dn, fn in os.walk(folder):
        for f in sorted(fn):
            if f.lower().endswith(EXTS) and not f.startswith(("~$", "~pe_tmp_")):
                files.append(os.path.join(dp, f))
        if not a.recursive:
            break
    stamp = datetime.datetime.now().strftime("%Y%m%d-%H%M%S")
    # next to the program, not in AppData: packaged apps can redirect AppData so Word would not see the copies
    base = os.path.join(a.work or os.path.join(HERE, "FormatCheck_runs"), f"run_{stamp}")
    os.makedirs(base, exist_ok=True)
    keymap = open(os.path.join(base, "KEY_stays_on_this_PC.txt"), "w", encoding="utf-8")
    word_ok = (not a.no_word) and sys.platform == "win32"
    data = {"version": VERSION, "run": stamp, "word_available": word_ok, "documents": []}
    print(f"FormatCheck {VERSION}: {len(files)} documents. Work folder: {base}")
    for i, src in enumerate(files, 1):
        did = f"Doc-{i}"
        ext = os.path.splitext(src)[1].lower()
        keymap.write(f"{did}\t{src}\n"); keymap.flush()
        d = {"id": did, "type": ext}
        h0 = sha(src)
        work = os.path.join(base, did + ext)
        shutil.copyfile(src, work)  # copyfile: data only, no attributes such as read-only
        try:
            t = time.perf_counter()
            pkg = Package(work)
            d["load"] = "accepted"
            d["load_s"] = round(time.perf_counter() - t, 2)
            d["structure"] = structure(pkg)
            del pkg
        except Refused as e:
            d["load"] = "refused: " + category(e)
        except Exception as e:
            d["load"] = "error: " + type(e).__name__
        if d["load"] == "accepted":
            print(f"  {did}: operations in memory")
            d["dry"] = dry_ops(work)
            before_copy = os.path.join(base, did + "_original" + ext)
            shutil.copyfile(work, before_copy)
            print(f"  {did}: full recipe on the copy")
            d["recipe"] = recipe(work)
            if word_ok and ext in (".docx", ".docm"):
                print(f"  {did}: Word check (up to {a.timeout} s)")
                d["word"] = run_word(before_copy, work, a.timeout)
        d["original_untouched"] = sha(src) == h0
        data["documents"].append(d)
        write_report(base, data)
    keymap.close()
    if not a.keep:
        for f in os.listdir(base):
            p = os.path.join(base, f)
            if f.startswith("Doc-") or f == safewrite.SNAP_DIR:
                shutil.rmtree(p) if os.path.isdir(p) else os.remove(p)
    j, h = write_report(base, data)
    print(f"Done. Report: {h}")
    print("Share FormatCheck_report.json or .html if you want: they hold counts only. Keep KEY_stays_on_this_PC.txt local.")
    try:
        os.startfile(h)
    except Exception:
        pass


if __name__ == "__main__":
    try:
        main()
    except Exception:
        traceback.print_exc()
        input("FormatCheck stopped with an error. Press Enter to close.")
