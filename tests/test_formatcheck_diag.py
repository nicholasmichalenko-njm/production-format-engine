"""FormatCheck diagnostics, feedback guard and export: runs anywhere (no Word, no window).
Usage: python tests/test_formatcheck_diag.py"""
import json, os, shutil, subprocess, sys, tempfile, zipfile
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT); sys.path.insert(0, os.path.join(ROOT, "fieldcheck"))
import diag, fieldcheck as fc, feedback as FB

ok = []
def check(name, cond):
    ok.append((name, bool(cond))); print(("PASS " if cond else "FAIL ") + name)

tmp = tempfile.mkdtemp(prefix="fc_diag_")
fc.HERE = tmp; fc.LOG_DIR = os.path.join(tmp, "logs"); fc.RUNS_DIR = os.path.join(tmp, "FormatCheck_runs")
diag.setup(fc.LOG_DIR, fc.VERSION)

# 1. an exception whose message holds a path and a name never reaches the log
try:
    raise ValueError(r"C:\Users\nick\Proposals\Kestrel_RFP_Vol1.docx could not open")
except ValueError:
    diag.error("load", "Loading a document raised an error", exc=True, docs=3, files="Kestrel")
ev = open(os.path.join(fc.LOG_DIR, "events.jsonl")).read()
check("exception message and string counts never logged", "Kestrel" not in ev and "nick" not in ev and "Users" not in ev)
e = json.loads(ev.splitlines()[0])
check("error has type, code location and counts", e["type"] == "ValueError" and e["where"][0].startswith("test_formatcheck_diag.py:")
      and e["docs"] == 3 and "files" not in e)

# 2. a real run without Word on two fixtures and a damaged file
vols = os.path.join(tmp, "Kestrel Program Volumes"); os.makedirs(vols)
shutil.copy(os.path.join(ROOT, "fixtures", "raw_small.docx"), os.path.join(vols, "Kestrel_Vol1_Technical.docx"))
open(os.path.join(vols, "Damaged_Kestrel_Mgmt.docx"), "wb").write(open(os.path.join(ROOT, "fixtures", "raw_small.docx"), "rb").read()[:20000])
lines = []
data, h = fc.run_check(vols, use_word=False, progress=lines.append)
check("run checks both documents", len(data["documents"]) == 2)
check("progress lines never name a file", not any("Kestrel" in l or "Damaged" in l for l in lines))
act = [json.loads(l) for l in open(os.path.join(fc.LOG_DIR, "activity.jsonl"))]
names = [a["action"] for a in act]
check("activity has run_started, doc_done x2, run_finished", names.count("doc_done") == 2 and "run_started" in names and "run_finished" in names)
dd = [a for a in act if a["action"] == "doc_done"]
check("doc_done carries counts and refusal code", any(a.get("refused") == 1 and a.get("code") for a in dd) and any(a.get("accepted") == 1 and a.get("paragraphs", 0) > 0 for a in dd))
check("activity is content-free", "Kestrel" not in open(os.path.join(fc.LOG_DIR, "activity.jsonl")).read())

# 3. machine-file proof
for n in ("events.jsonl", "activity.jsonl"):
    check(f"{n} passes the schema proof", FB.check_machine_file(n, open(os.path.join(fc.LOG_DIR, n)).read()) == [])
bad = '{"time":"2026-10-29T10:00:00","version":"0.3.0","level":"error","area":"load","msg":"C:\\\\Kestrel"}'
check("schema proof rejects a path in a message", FB.check_machine_file("x", bad) != [])
check("schema proof rejects an unknown field", FB.check_machine_file("x", '{"action":"click","name":"Kestrel"}') != [])

# 4. guard
g = FB.Guard()
check("guard knows checked file and folder names", g.find("It failed on Kestrel volume one") and g.find("the KESTREL folder"))
check("guard allows feature text", not g.find("The Word check timed out on a big volume and the report was empty"))
okk, m = FB.save_report("Bug", "Word check", "High", "Kestrel run hung", "Word stopped")
check("save blocked when the title names a volume", not okk)
okk, m = FB.save_report("Improvement", "Report", "Medium", "Show pages per document", "Add a pages column to the window summary.")
check("clean report saved", okk and len(FB.list_reports()) == 1)

# 5. export
reports, files, blocks = FB.prepare_export()
check("export is not blocked", blocks == [])
check("export holds the six files", set(files) == {"REPORT.md", "reports.json", "runs.json", "system.json", "events.jsonl", "activity.jsonl"})
allt = "".join(files.values())
check("export is content-free (no names, no paths, no user)", "Kestrel" not in allt and "Damaged" not in allt and tmp not in allt)
check("REPORT.md has reports, errors, runs and activity", all(x in files["REPORT.md"] for x in ("## Reports (1)", "## Errors and warnings", "## Runs on this PC", "## Activity")))
z = os.path.join(tmp, FB.default_zip_name()); FB.write_zip(z, reports, files)
check("zip written and report marked exported", zipfile.ZipFile(z).namelist() and all(r.get("exported") for _p, r in FB.list_reports()))
# a report typed with a name after the fact blocks the export
p = FB.list_reports()[0][0]; r = json.load(open(p)); r["exported"] = False; r["text"] += " Kestrel"; json.dump(r, open(p, "w"))
check("export blocked when a saved report names a volume", FB.prepare_export()[2] != [])

# 6. build gate
gate = subprocess.run([sys.executable, os.path.join(ROOT, "tools", "check_diag.py")], capture_output=True, text=True)
check("build gate passes on the source", gate.returncode == 0)
diag.shutdown()
shutil.rmtree(tmp, ignore_errors=True)
print(f"{sum(1 for _n, v in ok if v)}/{len(ok)} passed")
sys.exit(0 if all(v for _n, v in ok) else 1)
