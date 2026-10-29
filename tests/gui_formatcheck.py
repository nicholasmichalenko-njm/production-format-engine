"""Drive the real FormatCheck window on Windows: run a check (no Word), save a feedback report (one blocked by the
guard, one clean), open the export review, save the zip, close. Usage: python tests/gui_formatcheck.py <volumes folder>"""
import json, os, shutil, sys, tempfile, time, zipfile
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT); sys.path.insert(0, os.path.join(ROOT, "fieldcheck"))
import diag, fieldcheck as fc, feedback as FB

tmp = tempfile.mkdtemp(prefix="fc_gui_")
fc.HERE = tmp; fc.LOG_DIR = os.path.join(tmp, "logs"); fc.RUNS_DIR = os.path.join(tmp, "FormatCheck_runs")
diag.setup(fc.LOG_DIR, fc.VERSION)
import gui
vols = sys.argv[1]
res = {}
app = gui.App()
app._open = lambda path: res.setdefault("opened", []).append(os.path.basename(path))
app.use_word.set(False)
app.folder.set(vols)
steps = []


def step(fn, delay=300):
    steps.append((fn, delay))


def run_steps(i=0):
    if i >= len(steps):
        return
    fn, delay = steps[i]
    try:
        fn()
    except Exception as e:
        res.setdefault("errors", []).append(f"step {i}: {type(e).__name__}")
    app.after(delay, run_steps, i + 1)


def start():
    app.btn_again.configure(state="normal")
    app.btn_again.invoke()
    res["busy_stop_enabled"] = str(app.btn_stop.cget("state")) == "normal"


def wait_done():
    t0 = time.time()
    while app.worker is not None and app.worker.is_alive() and time.time() - t0 < 300:
        app.update(); time.sleep(0.1)
    for _ in range(20):
        app.update(); time.sleep(0.05)
    res["status"] = app.status.cget("text")
    res["report_button"] = str(app.btn_report.cget("state"))
    res["log_lines"] = app.log.get("1.0", "end").count("\n")
    res["log_has_names"] = any(os.path.splitext(f)[0] in app.log.get("1.0", "end") for f in os.listdir(vols))


def feedback_blocked():
    w = gui.FeedbackWindow(app)
    res["fb_window"] = w.winfo_exists()
    name = os.path.splitext(sorted(os.listdir(vols))[0])[0].split("_")[0]
    w.e_title.insert(0, f"{name} run was slow"); w.box.insert("1.0", "It took long.")
    w.save()
    res["fb_blocked_msg"] = w.msg.cget("text").startswith("Not saved")
    w.e_title.delete(0, "end"); w.box.delete("1.0", "end")
    w.e_title.insert(0, "Show pages per document in the window"); w.box.insert("1.0", "After a run, list pages before and after for each Doc.")
    w.v_type.set("Improvement"); w.v_feat.set("Window and layout")
    w.save()
    res["fb_saved"] = len(FB.list_reports()) == 1
    w.destroy()


def export():
    w = gui.ExportWindow(app)
    res["export_blocks"] = w.blocks
    out = os.path.join(tmp, "fb.zip")
    gui.filedialog.asksaveasfilename = lambda **k: out
    w.save()
    z = zipfile.ZipFile(out)
    res["zip_files"] = sorted(z.namelist())
    allt = "".join(z.read(n).decode("utf-8") for n in z.namelist())
    res["zip_has_names"] = any(os.path.splitext(f)[0] in allt for f in os.listdir(vols)) or vols in allt
    res["zip_report_md_has_run"] = "## Runs on this PC (1" in z.read("REPORT.md").decode()
    w.destroy()


def finish():
    act = [json.loads(l) for l in open(os.path.join(fc.LOG_DIR, "activity.jsonl"))]
    res["clicks"] = sorted({a.get("w") for a in act if a["action"] == "click" and a.get("w")})
    res["actions"] = sorted({a["action"] for a in act})
    res["events"] = open(os.path.join(fc.LOG_DIR, "events.jsonl")).read().count("\n") if os.path.exists(os.path.join(fc.LOG_DIR, "events.jsonl")) else 0
    app.destroy()


step(start, 200); step(wait_done, 100); step(feedback_blocked, 200); step(export, 200); step(finish, 100)
app.after(500, run_steps)
app.mainloop()
diag.shutdown()
res["pass"] = (res.get("busy_stop_enabled") and res.get("report_button") == "normal" and not res.get("log_has_names")
               and res.get("fb_blocked_msg") and res.get("fb_saved") and res.get("export_blocks") == []
               and len(res.get("zip_files", [])) == 6 and not res.get("zip_has_names") and res.get("zip_report_md_has_run")
               and "run_finished" in res.get("actions", []) and not res.get("errors"))
print(json.dumps(res, indent=1))
shutil.rmtree(tmp, ignore_errors=True)
sys.exit(0 if res["pass"] else 1)
