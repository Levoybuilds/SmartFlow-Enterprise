import contextlib
import io
import os
import queue
import sys
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from pathlib import Path

ROOT = Path(sys.executable).resolve().parent if getattr(sys, "frozen", False) else Path(__file__).resolve().parent
os.environ["SMARTFLOW_PROJECT_ROOT"] = str(ROOT)

# Static import lets PyInstaller discover the pipeline runner.
import run as pipeline_runner

OUTPUT = ROOT / "output"
STAGES = [
    "File Discovery", "Data Profiling", "Business Validation",
    "Data Quality Report", "Transformation Engine",
    "Data Mart Generation", "KPI Engine", "Executive Report"
]


class QueueWriter(io.TextIOBase):
    def __init__(self, q):
        self.q = q
    def write(self, text):
        if text:
            for line in text.splitlines():
                self.q.put(("log", line))
        return len(text)
    def flush(self):
        pass


class App:
    def __init__(self, root):
        self.root = root
        self.root.title("SmartFlow Enterprise")
        self.root.geometry("900x650")
        self.root.minsize(800, 580)
        self.q = queue.Queue()
        self.running = False
        self.selected_file = tk.StringVar(value="No file selected")

        bg, navy, blue = "#F5F7FB", "#172554", "#2563EB"
        root.configure(bg=bg)
        head = tk.Frame(root, bg=navy, height=105)
        head.pack(fill="x")
        head.pack_propagate(False)
        tk.Label(head, text="SMARTFLOW ENTERPRISE", bg=navy, fg="white",
                 font=("Segoe UI", 24, "bold")).pack(anchor="w", padx=32, pady=(20, 0))
        tk.Label(head, text="Automated Business Analytics Platform",
                 bg=navy, fg="#CBD5E1", font=("Segoe UI", 11)).pack(anchor="w", padx=34)

        main = tk.Frame(root, bg=bg)
        main.pack(fill="both", expand=True, padx=28, pady=22)
        card = tk.Frame(main, bg="white", highlightbackground="#E2E8F0", highlightthickness=1)
        card.pack(fill="x")
        tk.Label(card, text="DATA SOURCE", bg="white", fg="#64748B",
                 font=("Segoe UI", 9, "bold")).pack(anchor="w", padx=20, pady=(16, 5))
        pickrow = tk.Frame(card, bg="white")
        pickrow.pack(fill="x", padx=20, pady=(0, 18))
        tk.Label(pickrow, textvariable=self.selected_file, bg="#F8FAFC", fg="#172033",
                 anchor="w", padx=12, font=("Segoe UI", 10)).pack(side="left", fill="x", expand=True, ipady=8)
        self.browse = tk.Button(pickrow, text="Browse Excel/CSV", command=self.browse_file,
                                bg="#E2E8F0", fg=navy, relief="flat", padx=14, pady=8,
                                font=("Segoe UI", 9, "bold"), cursor="hand2")
        self.browse.pack(side="left", padx=(10, 0))

        row = tk.Frame(main, bg=bg)
        row.pack(fill="x", pady=18)
        self.status = tk.StringVar(value="Select a file to begin")
        tk.Label(row, textvariable=self.status, bg=bg, fg="#172033",
                 font=("Segoe UI", 11, "bold")).pack(side="left")
        self.run = tk.Button(row, text="▶  RUN ANALYSIS", command=self.start,
                             bg=blue, fg="white", relief="flat", bd=0,
                             padx=28, pady=11, font=("Segoe UI", 11, "bold"), cursor="hand2")
        self.run.pack(side="right")

        pc = tk.Frame(main, bg="white", highlightbackground="#E2E8F0", highlightthickness=1)
        pc.pack(fill="x")
        tk.Label(pc, text="PIPELINE PROGRESS", bg="white", fg="#64748B",
                 font=("Segoe UI", 9, "bold")).pack(anchor="w", padx=20, pady=(16, 8))
        self.progress = ttk.Progressbar(pc, maximum=8)
        self.progress.pack(fill="x", padx=20)
        self.plabel = tk.Label(pc, text="0 / 8 stages completed", bg="white",
                               fg="#64748B", font=("Segoe UI", 9))
        self.plabel.pack(anchor="w", padx=20, pady=(6, 16))

        lc = tk.Frame(main, bg="white", highlightbackground="#E2E8F0", highlightthickness=1)
        lc.pack(fill="both", expand=True, pady=(18, 0))
        hr = tk.Frame(lc, bg="white")
        hr.pack(fill="x", padx=20, pady=(14, 6))
        tk.Label(hr, text="ACTIVITY LOG", bg="white", fg="#64748B",
                 font=("Segoe UI", 9, "bold")).pack(side="left")
        self.openbtn = tk.Button(hr, text="Open Output Folder", command=self.open_output,
                                 state="disabled", bg="white", fg=navy, relief="flat", bd=0,
                                 font=("Segoe UI", 9, "bold"), cursor="hand2")
        self.openbtn.pack(side="right")
        self.log = tk.Text(lc, bg="#0F172A", fg="#E2E8F0", relief="flat", bd=0,
                           font=("Consolas", 9), wrap="word", state="disabled")
        self.log.pack(fill="both", expand=True, padx=20, pady=(0, 18))
        self.write("SmartFlow Enterprise is ready.")
        self.write("1. Browse and select an Excel or CSV file.")
        self.write("2. Click RUN ANALYSIS to execute the 8-stage pipeline.")
        self.poll()

    def browse_file(self):
        path = filedialog.askopenfilename(
            title="Select SmartFlow input file",
            filetypes=[("Excel and CSV files", "*.xlsx *.xls *.csv"),
                       ("Excel workbooks", "*.xlsx *.xls"), ("CSV files", "*.csv")]
        )
        if path:
            self.selected_file.set(path)
            self.status.set("Ready to analyze")
            self.write(f"Selected input: {path}")

    def write(self, text):
        self.log.config(state="normal")
        self.log.insert("end", text + "\n")
        self.log.see("end")
        self.log.config(state="disabled")

    def start(self):
        if self.running:
            return
        path = self.selected_file.get()
        if path == "No file selected" or not Path(path).is_file():
            messagebox.showwarning("SmartFlow", "Please browse and select an Excel or CSV file first.")
            return
        if Path(path).name.startswith("~$"):
            messagebox.showwarning("SmartFlow", "Please select the workbook itself, not a temporary lock file.")
            return
        self.running = True
        self.run.config(state="disabled", text="⏳  RUNNING...")
        self.browse.config(state="disabled")
        self.openbtn.config(state="disabled")
        self.progress["value"] = 0
        self.plabel.config(text="0 / 8 stages completed")
        self.status.set("Starting SmartFlow pipeline...")
        self.write("\n" + "=" * 68)
        self.write("SMARTFLOW ANALYSIS STARTED")
        self.write("=" * 68)
        threading.Thread(target=self.execute, args=(path,), daemon=True).start()

    def execute(self, path):
        writer = QueueWriter(self.q)
        try:
            os.environ["PYTHONIOENCODING"] = "utf-8"
            os.environ["PYTHONUTF8"] = "1"
            with contextlib.redirect_stdout(writer), contextlib.redirect_stderr(writer):
                code = pipeline_runner.main(path)
            self.q.put(("done", code))
        except Exception as exc:
            self.q.put(("log", f"ERROR: {type(exc).__name__}: {exc}"))
            self.q.put(("error", str(exc)))

    def poll(self):
        try:
            while True:
                typ, data = self.q.get_nowait()
                if typ == "log":
                    self.write(data)
                    for i, stage in enumerate(STAGES, 1):
                        if f"✓ COMPLETED: {stage}" in data:
                            self.progress["value"] = i
                            self.plabel.config(text=f"{i} / 8 stages completed")
                            self.status.set(f"Completed: {stage}")
                elif typ == "done":
                    self.finish(data)
                else:
                    self.fail(data)
        except queue.Empty:
            pass
        self.root.after(100, self.poll)

    def finish(self, code):
        self.running = False
        self.run.config(state="normal", text="▶  RUN ANALYSIS")
        self.browse.config(state="normal")
        if code == 0:
            self.progress["value"] = 8
            self.plabel.config(text="8 / 8 stages completed")
            self.status.set("Analysis complete")
            self.openbtn.config(state="normal")
            messagebox.showinfo("SmartFlow Complete",
                                "Analysis completed successfully.\\n\\nReports are in the output folder beside SmartFlow.")
        else:
            self.status.set("Pipeline failed")
            messagebox.showerror("SmartFlow Error", "The pipeline failed. Review the Activity Log.")
    def fail(self, error):
        self.running = False
        self.run.config(state="normal", text="▶  RUN ANALYSIS")
        self.browse.config(state="normal")
        self.status.set("Launcher error")
        self.write("Launcher error: " + error)
        messagebox.showerror("SmartFlow Launcher Error", error)
    def open_output(self):
        OUTPUT.mkdir(parents=True, exist_ok=True)
        os.startfile(OUTPUT)


if __name__ == "__main__":
    root = tk.Tk()
    App(root)
    root.mainloop()
