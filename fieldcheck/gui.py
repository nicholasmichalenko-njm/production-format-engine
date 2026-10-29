"""The FormatCheck window. Plain tkinter (part of Python, nothing extra to approve).

Main window: pick a folder, options, progress with Doc-N lines only, stop after the current document, open the
report, Feedback (Ctrl+Shift+F), Export for Claude, open the logs folder. The check runs in a background thread;
the window stays responsive. Every click is logged by internal control name only (diag.install_click_logger).
"""
import os, queue, threading, tkinter as tk
from tkinter import ttk, filedialog, messagebox
import diag
import feedback as FB
import fieldcheck as fc

PAD = 10


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(f"FormatCheck {fc.VERSION}")
        self.geometry("760x600"); self.minsize(640, 500)
        self.report_callback_exception = self._tk_error
        self.q = queue.Queue()
        self.worker = None
        self.stop_flag = threading.Event()
        self.last_report = None
        self.last_run_dir = None
        self.folder = tk.StringVar(value="")
        self.recursive = tk.BooleanVar(value=False)
        self.use_word = tk.BooleanVar(value=True)
        style = ttk.Style(self)
        try:
            style.theme_use("vista")
        except Exception:
            diag.swallowed("window")
        style.configure("Head.TLabel", font=("Segoe UI", 15, "bold"))
        style.configure("Note.TLabel", foreground="#475569")
        self._build()
        diag.install_click_logger(self)
        self.bind_all("<Control-Shift-F>", lambda e: self.open_feedback())
        self.bind_all("<Control-Shift-f>", lambda e: self.open_feedback())
        self.protocol("WM_DELETE_WINDOW", self.on_close)
        self.after(150, self._poll)
        diag.activity("app_started")

    # ------------------------------------------------------------ layout
    def _build(self):
        top = ttk.Frame(self, padding=PAD); top.pack(fill="x")
        ttk.Label(top, text="FormatCheck", style="Head.TLabel").pack(anchor="w")
        ttk.Label(top, style="Note.TLabel", wraplength=720, justify="left",
                  text="Checks Word volumes on local copies. The originals are never changed: FormatCheck proves "
                       "they are byte for byte the same at the end. The report holds counts only.").pack(anchor="w", pady=(2, 8))
        row = ttk.Frame(top); row.pack(fill="x")
        self.btn_pick = ttk.Button(row, text="Check a folder...", command=self.pick_and_run)
        self.btn_pick.pack(side="left")
        self.btn_again = ttk.Button(row, text="Check it again", command=self.run_again, state="disabled")
        self.btn_again.pack(side="left", padx=(8, 0))
        self.btn_stop = ttk.Button(row, text="Stop after this document", command=self.stop, state="disabled")
        self.btn_stop.pack(side="left", padx=(8, 0))
        opts = ttk.Frame(top); opts.pack(fill="x", pady=(8, 0))
        self.chk_sub = ttk.Checkbutton(opts, text="Include subfolders", variable=self.recursive)
        self.chk_sub.pack(side="left")
        self.chk_word = ttk.Checkbutton(opts, text="Check with Word (hidden, macros off)", variable=self.use_word)
        self.chk_word.pack(side="left", padx=(16, 0))

        mid = ttk.Frame(self, padding=(PAD, 0)); mid.pack(fill="both", expand=True)
        self.bar = ttk.Progressbar(mid, mode="determinate"); self.bar.pack(fill="x", pady=(4, 4))
        self.status = ttk.Label(mid, text="Pick a folder with the Word volumes to check.", anchor="w")
        self.status.pack(fill="x")
        frame = ttk.Frame(mid); frame.pack(fill="both", expand=True, pady=(6, 0))
        self.log = tk.Text(frame, height=14, wrap="word", font=("Consolas", 10), relief="solid", borderwidth=1,
                           state="disabled", background="#FBFCFD")
        sb = ttk.Scrollbar(frame, command=self.log.yview); self.log.configure(yscrollcommand=sb.set)
        self.log.pack(side="left", fill="both", expand=True); sb.pack(side="right", fill="y")
        res = ttk.Frame(mid); res.pack(fill="x", pady=(8, 0))
        self.btn_report = ttk.Button(res, text="Open report", command=self.open_report, state="disabled")
        self.btn_report.pack(side="left")
        self.btn_runfolder = ttk.Button(res, text="Open run folder", command=self.open_run_folder, state="disabled")
        self.btn_runfolder.pack(side="left", padx=(8, 0))

        bot = ttk.Frame(self, padding=PAD); bot.pack(fill="x")
        ttk.Separator(bot).pack(fill="x", pady=(0, 8))
        self.btn_feedback = ttk.Button(bot, text="Feedback...  (Ctrl+Shift+F)", command=self.open_feedback)
        self.btn_feedback.pack(side="left")
        self.btn_export = ttk.Button(bot, text="Export for Claude...", command=self.open_export)
        self.btn_export.pack(side="left", padx=(8, 0))
        self.btn_logs = ttk.Button(bot, text="Open logs folder", command=lambda: self._open(fc.LOG_DIR))
        self.btn_logs.pack(side="left", padx=(8, 0))
        self.issues = ttk.Label(bot, text="", style="Note.TLabel"); self.issues.pack(side="right")
        for name in ("btn_pick", "btn_again", "btn_stop", "chk_sub", "chk_word", "btn_report", "btn_runfolder",
                     "btn_feedback", "btn_export", "btn_logs", "log", "bar"):
            diag.register(getattr(self, name), name)

    # ------------------------------------------------------------ run
    def pick_and_run(self):
        d = filedialog.askdirectory(parent=self, title="Pick the folder with the volumes to check (read-only)")
        if not d:
            diag.activity("pick_cancelled")
            return
        self.folder.set(d)
        self.start()

    def run_again(self):
        if self.folder.get() and os.path.isdir(self.folder.get()):
            self.start()

    def start(self):
        folder = self.folder.get()
        files = fc.find_files(folder, self.recursive.get())
        if not files:
            diag.activity("pick_empty")
            messagebox.showinfo("FormatCheck", "No Word files (.docx, .docm, .dotx, .dotm) in that folder"
                                + ("" if self.recursive.get() else ". Tick Include subfolders to look deeper") + ".", parent=self)
            return
        self.stop_flag.clear()
        self._clear_log()
        self.bar.configure(maximum=len(files), value=0)
        self._busy(True)
        self._line(f"{len(files)} Word files found. Copies go to the FormatCheck_runs folder next to the program.")
        use_word = self.use_word.get()
        rec = self.recursive.get()

        def work():
            try:
                data, h = fc.run_check(folder, rec, use_word, progress=lambda t: self.q.put(("line", t)),
                                       should_stop=self.stop_flag.is_set,
                                       on_doc=lambda i, n: self.q.put(("doc", i, n)))
                self.q.put(("done", data, h))
            except Exception:
                diag.error("run", "The check stopped with an error", exc=True)
                self.q.put(("failed",))
        self.worker = threading.Thread(target=work, daemon=True)
        self.worker.start()

    def stop(self):
        self.stop_flag.set()
        self.btn_stop.configure(state="disabled")
        self.status.configure(text="Stopping after the current document...")

    def _poll(self):
        try:
            while True:
                m = self.q.get_nowait()
                if m[0] == "line":
                    self._line(m[1]); self.status.configure(text=m[1])
                elif m[0] == "doc":
                    self.bar.configure(value=m[1])
                elif m[0] == "done":
                    self._finished(m[1], m[2])
                elif m[0] == "failed":
                    self._busy(False)
                    self.status.configure(text="The check stopped with an error. It is in the logs; use Feedback to tell us what you did.")
        except queue.Empty:
            pass
        n = diag.issues_this_session()
        self.issues.configure(text=f"{n} issue(s) logged this session" if n else "")
        self.after(200, self._poll)

    def _finished(self, data, h):
        self._busy(False)
        self.last_report, self.last_run_dir = h, os.path.dirname(h)
        docs = data.get("documents", [])
        kept = sum(1 for d in docs if d.get("original_untouched"))
        acc = sum(1 for d in docs if d.get("load") == "accepted")
        stopped_steps = sum(1 for d in docs for s in d.get("recipe", []) if not s.get("ok"))
        word_issue = sum(1 for d in docs if (d.get("word") or {}).get("blocked_by_dialog") or (d.get("word") or {}).get("timed_out")
                         or (d.get("word") or {}).get("worker_failed"))
        msg = (f"{len(docs)} checked{' (stopped early)' if data.get('stopped') else ''}: {acc} accepted, "
               f"{stopped_steps} fix step(s) stopped, {word_issue} Word check problem(s). Originals untouched: {kept} of {len(docs)}.")
        self._line(msg)
        self.status.configure(text=msg)
        self.btn_report.configure(state="normal"); self.btn_runfolder.configure(state="normal")
        self.open_report()

    def _busy(self, on):
        st = "disabled" if on else "normal"
        for b in (self.btn_pick, self.chk_sub, self.chk_word, self.btn_export):
            b.configure(state=st)
        self.btn_again.configure(state="disabled" if on or not self.folder.get() else "normal")
        self.btn_stop.configure(state="normal" if on else "disabled")

    # ------------------------------------------------------------ helpers
    def _line(self, t):
        self.log.configure(state="normal"); self.log.insert("end", t + "\n"); self.log.see("end"); self.log.configure(state="disabled")

    def _clear_log(self):
        self.log.configure(state="normal"); self.log.delete("1.0", "end"); self.log.configure(state="disabled")

    def _open(self, path):
        try:
            os.makedirs(path, exist_ok=True) if not os.path.splitext(path)[1] else None
            os.startfile(path)
        except Exception:
            diag.swallowed("window")

    def open_report(self):
        if self.last_report and os.path.exists(self.last_report):
            diag.activity("report_opened")
            self._open(self.last_report)

    def open_run_folder(self):
        if self.last_run_dir:
            self._open(self.last_run_dir)

    def on_close(self):
        if self.worker is not None and self.worker.is_alive():
            if not messagebox.askyesno("FormatCheck", "A check is running. Close anyway? The current Word job is ended "
                                       "and the copies stay in the run folder.", parent=self):
                return
            diag.activity("closed_during_run")
        diag.activity("app_closed")
        self.destroy()

    def _tk_error(self, t, e, tb):
        diag.error("window", "Error in the window", exc=(t, e, tb))
        try:
            self.status.configure(text="Something went wrong in the window. It is in the logs.")
        except Exception:
            pass

    # ------------------------------------------------------------ feedback
    def open_feedback(self):
        FeedbackWindow(self)

    def open_export(self):
        ExportWindow(self)


class FeedbackWindow(tk.Toplevel):
    _open = None

    def __new__(cls, app):
        if cls._open is not None and cls._open.winfo_exists():
            cls._open.deiconify(); cls._open.lift(); cls._open.focus_force()
            return cls._open
        return super().__new__(cls)

    def __init__(self, app):
        if FeedbackWindow._open is self:
            return
        super().__init__(app)
        FeedbackWindow._open = self
        diag.activity("feedback_opened")
        self.title("Feedback"); self.geometry("640x600"); self.minsize(560, 500); self.transient(app)
        f = ttk.Frame(self, padding=PAD); f.pack(fill="both", expand=True)
        ttk.Label(f, text="Feedback: bugs, improvements and ideas", style="Head.TLabel").pack(anchor="w")
        ttk.Label(f, style="Note.TLabel", wraplength=600, justify="left",
                  text="Describe the feature, not your documents. Never type file names, folder names or text from a "
                       "volume: FormatCheck checks and will not save it. Your recent actions and any errors logged are "
                       "attached automatically, with no names or content.").pack(anchor="w", pady=(2, 8))
        form = ttk.Frame(f); form.pack(fill="x")
        form.columnconfigure(1, weight=1)
        self.v_type = tk.StringVar(value="Bug"); self.v_feat = tk.StringVar(value=FB.FEATURES[1]); self.v_imp = tk.StringVar(value="Medium")
        ttk.Label(form, text="Type").grid(row=0, column=0, sticky="w", pady=3)
        tf = ttk.Frame(form); tf.grid(row=0, column=1, sticky="w")
        for t in FB.TYPES:
            rb = ttk.Radiobutton(tf, text=t, value=t, variable=self.v_type); rb.pack(side="left", padx=(0, 10))
            diag.register(rb, "fb_type_" + t.lower())
        ttk.Label(form, text="Feature").grid(row=1, column=0, sticky="w", pady=3)
        cb = ttk.Combobox(form, values=FB.FEATURES, textvariable=self.v_feat, state="readonly", width=30)
        cb.grid(row=1, column=1, sticky="w"); diag.register(cb, "fb_feature")
        ttk.Label(form, text="Importance").grid(row=2, column=0, sticky="w", pady=3)
        imf = ttk.Frame(form); imf.grid(row=2, column=1, sticky="w")
        for t in FB.IMPORTANCE:
            rb = ttk.Radiobutton(imf, text=t, value=t, variable=self.v_imp); rb.pack(side="left", padx=(0, 10))
            diag.register(rb, "fb_imp_" + t.lower())
        ttk.Label(form, text="Title").grid(row=3, column=0, sticky="w", pady=3, padx=(0, 10))
        self.e_title = ttk.Entry(form); self.e_title.grid(row=3, column=1, sticky="ew"); diag.register(self.e_title, "fb_title")
        ttk.Label(f, text="What happened, or what you want (steps, expected vs actual, the idea):").pack(anchor="w", pady=(10, 2))
        self.box = tk.Text(f, height=12, wrap="word", font=("Segoe UI", 10), relief="solid", borderwidth=1)
        self.box.pack(fill="both", expand=True); diag.register(self.box, "fb_text")
        self.box.tag_configure("leak", background="#FDE68A", foreground="#7C2D12")
        self.msg = ttk.Label(f, text="", wraplength=600, justify="left"); self.msg.pack(fill="x", pady=(6, 0))
        bar = ttk.Frame(f); bar.pack(fill="x", pady=(8, 0))
        b1 = ttk.Button(bar, text="Save report", command=self.save); b1.pack(side="left"); diag.register(b1, "fb_save")
        b2 = ttk.Button(bar, text="Export for Claude...", command=lambda: ExportWindow(app)); b2.pack(side="left", padx=8)
        diag.register(b2, "fb_export")
        b3 = ttk.Button(bar, text="Close", command=self.destroy); b3.pack(side="right"); diag.register(b3, "fb_close")
        self.count = ttk.Label(f, text=self._count(), style="Note.TLabel"); self.count.pack(anchor="w", pady=(6, 0))
        self.bind("<Escape>", lambda e: self.destroy())
        self.e_title.focus_set()

    def _count(self):
        rs = FB.list_reports()
        return f"{len(rs)} report(s) saved, {sum(1 for _p, r in rs if not r.get('exported'))} not yet exported."

    def save(self):
        g = FB.Guard()
        title, text = self.e_title.get(), self.box.get("1.0", "end")
        self.box.tag_remove("leak", "1.0", "end")
        for s, e in g.find(text):
            self.box.tag_add("leak", f"1.0+{s}c", f"1.0+{e}c")
        ok, m = FB.save_report(self.v_type.get(), self.v_feat.get(), self.v_imp.get(), title, text, guard=g)
        self.msg.configure(text=m, foreground="#047857" if ok else "#B91C1C")
        if ok:
            self.e_title.delete(0, "end"); self.box.delete("1.0", "end")
            self.count.configure(text=self._count())


class ExportWindow(tk.Toplevel):
    def __init__(self, app):
        super().__init__(app)
        self.title("Export for Claude: review"); self.geometry("820x640"); self.transient(app)
        diag.activity("export_opened")
        self.reports, self.files, self.blocks = FB.prepare_export()
        f = ttk.Frame(self, padding=PAD); f.pack(fill="both", expand=True)
        ttk.Label(f, wraplength=780, justify="left", font=("Segoe UI", 11, "bold"),
                  text=f"{len(self.reports)} new report(s), the content-free logs and the run summaries from this PC. "
                       "This is everything in the zip. Read it before saving.").pack(anchor="w")
        frame = ttk.Frame(f); frame.pack(fill="both", expand=True, pady=6)
        view = tk.Text(frame, wrap="word", font=("Consolas", 9), relief="solid", borderwidth=1)
        sb = ttk.Scrollbar(frame, command=view.yview); view.configure(yscrollcommand=sb.set)
        view.pack(side="left", fill="both", expand=True); sb.pack(side="right", fill="y")
        for name, text in self.files.items():
            view.insert("end", f"===== {name} ({len(text):,} characters) =====\n{text[:30000]}\n"
                               + ("[... longer, shown in part ...]\n" if len(text) > 30000 else "") + "\n")
        view.configure(state="disabled")
        self.msg = ttk.Label(f, text="", wraplength=780, justify="left"); self.msg.pack(fill="x")
        bar = ttk.Frame(f); bar.pack(fill="x", pady=(8, 0))
        self.btn = ttk.Button(bar, text="Save zip...", command=self.save); self.btn.pack(side="left")
        diag.register(self.btn, "export_save")
        c = ttk.Button(bar, text="Close", command=self.destroy); c.pack(side="right"); diag.register(c, "export_close")
        if self.blocks:
            self.btn.configure(state="disabled")
            self.msg.configure(foreground="#B91C1C", text="Blocked: something in " + ", ".join(self.blocks) + " matches a file "
                               "or folder name from this PC, or a log line is not in the expected form. Edit or delete the "
                               "report in the feedback folder (next to FormatCheck.exe), then try again.")
            diag.activity("export_blocked", items=len(self.blocks))

    def save(self):
        start = os.path.join(os.path.expanduser("~"), "Desktop")
        path = filedialog.asksaveasfilename(parent=self, title="Save the feedback zip", defaultextension=".zip",
                                            initialfile=FB.default_zip_name(),
                                            initialdir=start if os.path.isdir(start) else None, filetypes=[("Zip", "*.zip")])
        if not path:
            return
        try:
            FB.write_zip(path, self.reports, self.files)
        except Exception:
            diag.error("export", "Saving the feedback zip failed", exc=True)
            self.msg.configure(foreground="#B91C1C", text="Could not save the zip there. Pick another folder.")
            return
        self.msg.configure(foreground="#047857", text="Saved. Hand this zip to Claude.")
        self.btn.configure(state="disabled")


def main():
    app = App()
    app.mainloop()
