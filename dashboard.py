"""TPI MB Downloader — local dashboard. Save IDs once, one click to download new bills."""
from __future__ import annotations

import json
import os
import queue
import re
import subprocess
import sys
import threading
import tkinter as tk
from datetime import datetime
from pathlib import Path
from tkinter import filedialog, messagebox, simpledialog, ttk

from vault import days_left, master_is_set, protect, session_unlock, set_master, unprotect, verify_master

ROOT = Path(__file__).resolve().parent

def _find_root() -> Path:
    cands = [ROOT, ROOT.parent, Path.cwd(), Path.cwd().parent]
    try:
        cands.extend(ROOT.parent.glob("tpi_mb_downloader*"))
        cands.extend(Path.cwd().glob("tpi_mb_downloader*"))
    except Exception:
        pass
    for p in cands:
        p = Path(p)
        if p.is_file():
            p = p.parent
        if (p / "tpi_mb_downloader.py").exists():
            return p
    return ROOT

ROOT = _find_root()
os.chdir(ROOT)
DATA = ROOT / ".tpidata"
DATA.mkdir(exist_ok=True)
ACCOUNTS_PATH = DATA / "accounts.bin"
SCRIPT = ROOT / "tpi_mb_downloader.py"
UPLOADER = ROOT / "tpi_uploader.py"
DOWNLOADS = ROOT / "downloads"
PY = ROOT / ".venv" / "Scripts" / "python.exe"
if not PY.exists():
    PY = Path(sys.executable)
SETTINGS_PATH = DATA / "settings.json"
_OLD_ACCOUNTS = ROOT / "accounts.json"


def load_settings() -> dict:
    if not SETTINGS_PATH.exists():
        return {}
    try:
        return json.loads(SETTINGS_PATH.read_text(encoding="utf-8"))
    except Exception:
        return {}


def save_settings(data: dict) -> None:
    DATA.mkdir(exist_ok=True)
    SETTINGS_PATH.write_text(json.dumps(data, indent=2), encoding="utf-8")


def load_accounts() -> list[dict]:
    path = ACCOUNTS_PATH
    if not path.exists() and _OLD_ACCOUNTS.exists():
        path = _OLD_ACCOUNTS
    if not path.exists():
        return []
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
        rows = list(data.get("accounts") or [])
    except Exception:
        return []
    migrated = False
    out = []
    for a in rows:
        rec = dict(a)
        plain = ""
        if rec.get("pass_enc"):
            try:
                plain = unprotect(rec["pass_enc"])
            except Exception:
                plain = ""
        elif rec.get("pass"):
            plain = rec["pass"]
            rec["pass_enc"] = protect(plain)
            migrated = True
        rec["pass"] = plain
        rec.pop("pass_enc_plain", None)
        out.append(rec)
    if migrated:
        save_accounts(out)
    return out


def save_accounts(rows: list[dict]) -> None:
    disk = []
    for a in rows:
        item = {
            "name": a.get("name") or "",
            "user": a.get("user") or "",
            "enabled": bool(a.get("enabled", True)),
            "last_run": a.get("last_run") or "",
            "last_session": a.get("last_session") or "",
            "last_update": a.get("last_update") or "",
            "pass_enc": a.get("pass_enc") or "",
        }
        pw = a.get("pass") or ""
        if pw:
            item["pass_enc"] = protect(pw)
            a["pass_enc"] = item["pass_enc"]
        disk.append(item)
    ACCOUNTS_PATH.parent.mkdir(exist_ok=True)
    ACCOUNTS_PATH.write_text(
        json.dumps({"accounts": disk}, indent=2, ensure_ascii=False),
        encoding="utf-8",
    )
    if _OLD_ACCOUNTS.exists():
        try:
            _OLD_ACCOUNTS.unlink()
        except Exception:
            pass
    try:
        import subprocess
        subprocess.run(["attrib", "+h", "+s", str(DATA)], check=False, creationflags=0x08000000)
    except Exception:
        pass


def unlock_or_exit(root: tk.Tk) -> bool:
    if session_unlock():
        return True
    if not master_is_set():
        while True:
            p1 = simpledialog.askstring("Setup", "Create master password (8+ characters):", show="*", parent=root)
            if p1 is None:
                return False
            if len(p1.strip()) < 8:
                messagebox.showerror("Password", "At least 8 characters.", parent=root)
                continue
            p2 = simpledialog.askstring("Setup", "Confirm master password:", show="*", parent=root)
            if p2 is None:
                return False
            if p1 != p2:
                messagebox.showerror("Password", "Passwords do not match.", parent=root)
                continue
            try:
                set_master(p1)
            except Exception as exc:
                messagebox.showerror("Password", str(exc), parent=root)
                continue
            return True
    for _ in range(3):
        pw = simpledialog.askstring(
            "Renew",
            "Enter master password to renew for 15 days:",
            show="*",
            parent=root,
        )
        if pw is None:
            return False
        try:
            verify_master(pw)
            return True
        except Exception:
            messagebox.showerror("Password", "Incorrect password.", parent=root)
    return False


def safe_folder(name: str) -> str:
    keep = "".join(c if c.isalnum() or c in "-_ " else "_" for c in (name or "id"))
    return (keep.strip() or "id")[:40]


class AccountDialog(tk.Toplevel):
    def __init__(self, master, title: str, acc: dict | None = None):
        super().__init__(master)
        self.title(title)
        self.resizable(False, False)
        self.result = None
        acc = acc or {}
        self.transient(master)
        self.grab_set()

        frm = ttk.Frame(self, padding=16)
        frm.grid()
        ttk.Label(frm, text="Name (e.g. TPI Civil 1)").grid(row=0, column=0, sticky="w")
        self.e_name = ttk.Entry(frm, width=36)
        self.e_name.grid(row=1, column=0, pady=(0, 8))
        self.e_name.insert(0, acc.get("name") or "")

        ttk.Label(frm, text="Portal login ID").grid(row=2, column=0, sticky="w")
        self.e_user = ttk.Entry(frm, width=36)
        self.e_user.grid(row=3, column=0, pady=(0, 8))
        self.e_user.insert(0, acc.get("user") or "")

        ttk.Label(frm, text="Portal password (hidden)").grid(row=4, column=0, sticky="w")
        self.e_pass = ttk.Entry(frm, width=36, show="*")
        self.e_pass.grid(row=5, column=0, pady=(0, 8))
        self._old_pass = acc.get("pass") or ""
        self._old_enc = acc.get("pass_enc") or ""
        if self._old_pass or self._old_enc:
            ttk.Label(frm, text="Leave blank to keep the saved password", foreground="#666").grid(
                row=6, column=0, sticky="w"
            )
            row_btns = 7
        else:
            row_btns = 6

        btns = ttk.Frame(frm)
        btns.grid(row=row_btns, column=0, sticky="e")
        ttk.Button(btns, text="Cancel", command=self.destroy).pack(side="right", padx=4)
        ttk.Button(btns, text="Save", command=self._ok).pack(side="right")
        self.e_name.focus_set()
        self.bind("<Return>", lambda e: self._ok())

    def _ok(self):
        name = self.e_name.get().strip()
        user = self.e_user.get().strip()
        pw = self.e_pass.get()
        if not name or not user:
            messagebox.showerror("Required", "Name and Login ID are required.", parent=self)
            return
        if not pw and not self._old_pass and not self._old_enc:
            messagebox.showerror("Required", "Password is required.", parent=self)
            return
        self.result = {
            "name": name,
            "user": user,
            "pass": pw if pw else self._old_pass,
            "pass_enc": "" if pw else self._old_enc,
            "enabled": True,
        }
        self.destroy()


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("TPI MB Downloader")
        self.geometry("1180x740")
        self.minsize(980, 620)
        self.configure(bg="#f4f5f7")
        self.log_q: queue.Queue[str] = queue.Queue()
        self.running = False
        self._apply_style()
        if not unlock_or_exit(self):
            self.destroy()
            return
        self.accounts = load_accounts()

        head = ttk.Frame(self, padding=(16, 12, 16, 6))
        head.pack(fill="x")
        ttk.Label(head, text="TPI Measurement Book", style="Title.TLabel").pack(anchor="w")
        ttk.Label(
            head,
            text="Portal IDs stay encrypted. Download left · Upload centre · Log on the right (drag the divider).",
            style="Hint.TLabel",
        ).pack(anchor="w")

        pathf = ttk.Frame(self, padding=(16, 0, 16, 8))
        pathf.pack(fill="x")
        ttk.Label(pathf, text="Save bills to (Kautilya E-MB folder)").pack(anchor="w")
        pr = ttk.Frame(pathf)
        pr.pack(fill="x")
        self.save_var = tk.StringVar(value=load_settings().get("save_root") or "")
        self.save_entry = ttk.Entry(pr, textvariable=self.save_var)
        self.save_entry.pack(side="left", fill="x", expand=True)
        ttk.Button(pr, text="Browse…", command=self.pick_save).pack(side="left", padx=6)

        body = ttk.Panedwindow(self, orient="horizontal")
        body.pack(fill="both", expand=True, padx=12, pady=(0, 12))
        left_outer = ttk.Frame(body)
        right = ttk.Frame(body, padding=(8, 0, 0, 0))
        body.add(left_outer, weight=3)
        body.add(right, weight=2)

        self.left_canvas = tk.Canvas(left_outer, highlightthickness=0, bg="#f4f5f7")
        left_scroll = ttk.Scrollbar(left_outer, orient="vertical", command=self.left_canvas.yview)
        left = ttk.Frame(self.left_canvas, padding=(0, 0, 8, 0))
        self._left_win = self.left_canvas.create_window((0, 0), window=left, anchor="nw")
        self.left_canvas.configure(yscrollcommand=left_scroll.set)

        def _left_inner_cfg(_e=None):
            self.left_canvas.configure(scrollregion=self.left_canvas.bbox("all"))

        def _left_canvas_cfg(e):
            self.left_canvas.itemconfigure(self._left_win, width=max(e.width, 1))

        left.bind("<Configure>", _left_inner_cfg)
        self.left_canvas.bind("<Configure>", _left_canvas_cfg)

        def _wheel(e):
            self.left_canvas.yview_scroll(int(-1 * (e.delta / 120)), "units")

        def _wheel_linux(e):
            self.left_canvas.yview_scroll(-1 if e.num == 4 else 1, "units")

        def _bind_wheel(_e=None):
            self.left_canvas.bind_all("<MouseWheel>", _wheel)
            self.left_canvas.bind_all("<Button-4>", _wheel_linux)
            self.left_canvas.bind_all("<Button-5>", _wheel_linux)

        def _unbind_wheel(_e=None):
            self.left_canvas.unbind_all("<MouseWheel>")
            self.left_canvas.unbind_all("<Button-4>")
            self.left_canvas.unbind_all("<Button-5>")

        left.bind("<Enter>", _bind_wheel)
        left.bind("<Leave>", _unbind_wheel)
        self.left_canvas.pack(side="left", fill="both", expand=True)
        left_scroll.pack(side="right", fill="y")

        mid = ttk.Frame(left)
        mid.pack(fill="x")
        cols = ("on", "name", "user", "last", "sess", "upd")
        self.tree = ttk.Treeview(mid, columns=cols, show="headings", height=3, selectmode="browse")
        self.tree.heading("on", text="On")
        self.tree.heading("name", text="Name")
        self.tree.heading("user", text="Login ID")
        self.tree.heading("last", text="Last run")
        self.tree.heading("sess", text="Last session")
        self.tree.heading("upd", text="Last update")
        self.tree.column("on", width=40, anchor="center")
        self.tree.column("name", width=120)
        self.tree.column("user", width=160)
        self.tree.column("last", width=110)
        self.tree.column("sess", width=90)
        self.tree.column("upd", width=110)
        self.tree.pack(side="left", fill="x", expand=True)
        sb = ttk.Scrollbar(mid, orient="vertical", command=self.tree.yview)
        sb.pack(side="right", fill="y")
        self.tree.configure(yscrollcommand=sb.set)

        actions = ttk.Frame(left, padding=(0, 6, 0, 4))
        actions.pack(fill="x")
        ttk.Button(actions, text="+ Add ID", command=self.add_acc).pack(side="left")
        ttk.Button(actions, text="Edit", command=self.edit_acc).pack(side="left", padx=4)
        ttk.Button(actions, text="Delete", command=self.del_acc).pack(side="left")
        ttk.Button(actions, text="On/Off", command=self.toggle_acc).pack(side="left", padx=4)
        ttk.Button(actions, text="Open save folder", command=self.open_dl).pack(side="right")

        self.nb = ttk.Notebook(left)
        self.nb.pack(fill="both", expand=True, pady=(6, 0))
        tab_dl = ttk.Frame(self.nb, padding=10)
        tab_up = ttk.Frame(self.nb, padding=10)
        tab_sum = ttk.Frame(self.nb, padding=10)
        self.nb.add(tab_dl, text="  1. Download  ")
        self.nb.add(tab_up, text="  2. Upload  ")
        self.nb.add(tab_sum, text="  3. Summary  ")

        ttk.Label(
            tab_dl,
            text="Tick what to pull. Master Excel tracks every portal bill so the next session only fetches remaining data.",
            style="Hint.TLabel",
        ).pack(anchor="w")
        mf = ttk.Frame(tab_dl)
        mf.pack(fill="x", pady=(6, 4))
        ttk.Label(mf, text="Master Excel (all IDs, continues across sessions)").pack(anchor="w")
        mr = ttk.Frame(mf)
        mr.pack(fill="x")
        default_master = ""
        try:
            sr = load_settings().get("save_root") or ""
            default_master = load_settings().get("master_path") or (str(Path(sr) / "tpi_master.xlsx") if sr else "")
        except Exception:
            default_master = ""
        self.master_var = tk.StringVar(value=default_master)
        ttk.Entry(mr, textvariable=self.master_var).pack(side="left", fill="x", expand=True)
        ttk.Button(mr, text="Load previous…", command=self.pick_master).pack(side="left", padx=6)
        dlf = ttk.LabelFrame(tab_dl, text="  Download this run  ")
        dlf.pack(fill="x", pady=(8, 8))
        self.dl_excel = tk.BooleanVar(value=True)
        self.dl_mb = tk.BooleanVar(value=False)
        self.dl_docs = tk.BooleanVar(value=False)
        self.dl_loa = tk.BooleanVar(value=True)
        self.dl_comments = tk.BooleanVar(value=False)
        ttk.Checkbutton(dlf, text="1. Session Excel only (list, no bill open)", variable=self.dl_excel).grid(row=0, column=0, columnspan=2, sticky="w", padx=8, pady=3)
        ttk.Checkbutton(dlf, text="1b. Open each bill → LOA Details + Grand Total (no PDF)", variable=self.dl_loa).grid(row=1, column=0, columnspan=2, sticky="w", padx=8, pady=3)
        ttk.Checkbutton(dlf, text="1c. Comments (TPI received date from last TPI comment)", variable=self.dl_comments).grid(row=2, column=0, columnspan=2, sticky="w", padx=8, pady=3)
        ttk.Checkbutton(dlf, text="2. Signed MB PDF + Signed MB Excel + Abstract", variable=self.dl_mb).grid(row=3, column=0, columnspan=2, sticky="w", padx=8, pady=3)
        ttk.Checkbutton(dlf, text="3. Supporting documents (Uploaded Documents)", variable=self.dl_docs).grid(row=4, column=0, sticky="w", padx=8, pady=3)
        ttk.Button(dlf, text="Select all", command=self.dl_steps_all).grid(row=0, column=2, padx=8)
        ttk.Button(dlf, text="Clear", command=self.dl_steps_none).grid(row=1, column=2, padx=8)
        runbar = ttk.Frame(tab_dl, padding=(0, 8, 0, 0))
        runbar.pack(fill="x")
        self.btn_all = ttk.Button(runbar, text="  ▶  Download all IDs  ", style="Green.TButton", command=self.run_all)
        self.btn_all.pack(side="left")
        self.btn_one = ttk.Button(runbar, text="  Download selected ID  ", style="Accent.TButton", command=self.run_one)
        self.btn_one.pack(side="left", padx=8)
        self.status = ttk.Label(runbar, text="Ready", font=("Segoe UI", 9, "bold"))
        self.status.pack(side="right")
        progf = ttk.Frame(tab_dl)
        progf.pack(fill="x", pady=(6, 0))
        self.prog_lbl = ttk.Label(progf, text="0 / 0  bills", style="Hint.TLabel")
        self.prog_lbl.pack(anchor="w")
        self.prog = ttk.Progressbar(progf, mode="determinate", maximum=100)
        self.prog.pack(fill="x", pady=(2, 0))
        gf = ttk.Frame(tab_dl)
        gf.pack(fill="x", pady=(6, 0))
        ttk.Label(gf, text="Google Sheet URL (optional)").pack(anchor="w")
        self.gsheet_var = tk.StringVar(value=load_settings().get("gsheet_url") or "")
        ttk.Entry(gf, textvariable=self.gsheet_var).pack(fill="x")

        ttk.Label(
            tab_up,
            text="Scan checks every Excel row has matching PDFs under 7 MB. Tick steps + rows, then Upload.",
            style="Hint.TLabel",
        ).pack(anchor="w", pady=(0, 6))
        stepsf = ttk.LabelFrame(tab_up, text="  Steps this run  ")
        stepsf.pack(fill="x", pady=(0, 6))
        self.step_files = tk.BooleanVar(value=True)
        self.step_fill = tk.BooleanVar(value=True)
        self.step_dsc = tk.BooleanVar(value=True)
        self.step_action = tk.BooleanVar(value=True)
        ttk.Checkbutton(stepsf, text="Upload reports / PDFs", variable=self.step_files).grid(row=0, column=0, sticky="w", padx=8, pady=3)
        ttk.Checkbutton(stepsf, text="Fill Amount + Letter + Remark + Save Draft", variable=self.step_fill).grid(row=0, column=1, sticky="w", padx=8, pady=3)
        ttk.Checkbutton(stepsf, text="Attach DSC (Verified) + Yes, please!", variable=self.step_dsc).grid(row=1, column=0, sticky="w", padx=8, pady=3)
        ttk.Checkbutton(stepsf, text="Action + OTP 000000 + Forward / Return", variable=self.step_action).grid(row=1, column=1, sticky="w", padx=8, pady=3)
        ttk.Button(stepsf, text="Select all steps", command=self.steps_all).grid(row=0, column=2, padx=8, pady=2)
        ttk.Button(stepsf, text="Clear steps", command=self.steps_none).grid(row=1, column=2, padx=8, pady=2)

        uf = ttk.Frame(tab_up)
        uf.pack(fill="x")
        ttk.Label(uf, text="Batch folder (Upload Template.xlsx + PDFs)").pack(anchor="w")
        ur = ttk.Frame(uf)
        ur.pack(fill="x")
        self.batch_var = tk.StringVar(value=load_settings().get("batch_folder") or "")
        ttk.Entry(ur, textvariable=self.batch_var).pack(side="left", fill="x", expand=True)
        ttk.Button(ur, text="Browse…", command=self.pick_batch).pack(side="left", padx=6)
        ttk.Button(ur, text="Reset", command=self.reset_batch).pack(side="left", padx=4)
        self.btn_up_scan = ttk.Button(ur, text="  Scan Excel + files  ", style="Accent.TButton", command=self.run_upload_scan)
        self.btn_up_scan.pack(side="left", padx=4)
        self.up_summary = ttk.Label(tab_up, text="No batch loaded.", style="Hint.TLabel")
        self.up_summary.pack(anchor="w", pady=4)

        ubar = ttk.Frame(tab_up)
        ubar.pack(fill="x", pady=(6, 8))
        ttk.Button(ubar, text="Select all rows", command=self.up_select_all).pack(side="left")
        ttk.Button(ubar, text="Clear rows", command=self.up_select_none).pack(side="left", padx=4)
        self.btn_up = ttk.Button(ubar, text="  ▶  PROCEED — Upload selected  ", style="Green.TButton", command=self.run_upload)
        self.btn_up.pack(side="left", padx=10)
        self.btn_up_all = ttk.Button(ubar, text="  Upload all rows  ", style="Accent.TButton", command=self.run_upload_all)
        self.btn_up_all.pack(side="left")

        ut = ttk.Frame(tab_up)
        ut.pack(fill="both", expand=True)
        ucols = ("sel", "sid", "mb", "letter", "status", "files", "warn")
        self.up_tree = ttk.Treeview(ut, columns=ucols, show="headings", height=7, selectmode="browse")
        self.up_tree.heading("sel", text="Sel")
        self.up_tree.heading("sid", text="Scheme ID")
        self.up_tree.heading("mb", text="MB No.")
        self.up_tree.heading("letter", text="Letter")
        self.up_tree.heading("status", text="Status")
        self.up_tree.heading("files", text="PDFs")
        self.up_tree.heading("warn", text="Check")
        self.up_tree.column("sel", width=36, anchor="center")
        self.up_tree.column("sid", width=88)
        self.up_tree.column("mb", width=100)
        self.up_tree.column("letter", width=56)
        self.up_tree.column("status", width=72)
        self.up_tree.column("files", width=160)
        self.up_tree.column("warn", width=140)
        usb = ttk.Scrollbar(ut, orient="vertical", command=self.up_tree.yview)
        self.up_tree.pack(side="left", fill="both", expand=True)
        usb.pack(side="right", fill="y")
        self.up_tree.configure(yscrollcommand=usb.set)
        self.up_tree.bind("<Button-1>", self._up_click)

        ttk.Label(
            tab_sum,
            text="Pending at TPI (last download, not yet forwarded/returned).\n"
                 "Forward / Return by date — Return amount uses portal Grand Total (TPI amount is 0.00).",
            style="Hint.TLabel",
        ).pack(anchor="w")
        sf = ttk.Frame(tab_sum)
        sf.pack(fill="x", pady=8)
        ttk.Label(sf, text="From (YYYY-MM-DD)").pack(side="left")
        self.sum_from = tk.StringVar(value=datetime.now().strftime("%Y-%m-01"))
        ttk.Entry(sf, textvariable=self.sum_from, width=12).pack(side="left", padx=6)
        ttk.Label(sf, text="To").pack(side="left")
        self.sum_to = tk.StringVar(value=datetime.now().strftime("%Y-%m-%d"))
        ttk.Entry(sf, textvariable=self.sum_to, width=12).pack(side="left", padx=6)
        ttk.Button(sf, text="  Refresh  ", style="Accent.TButton", command=self.refresh_summary).pack(side="left", padx=8)

        ttk.Label(tab_sum, text="Pending at our level", font=("Segoe UI", 10, "bold")).pack(anchor="w", pady=(6, 2))
        pcols = ("cluster", "district", "bills", "grand")
        self.sum_pending = ttk.Treeview(tab_sum, columns=pcols, show="headings", height=5)
        self.sum_pending.heading("cluster", text="Cluster")
        self.sum_pending.heading("district", text="District")
        self.sum_pending.heading("bills", text="Bills")
        self.sum_pending.heading("grand", text="Grand Total")
        self.sum_pending.column("cluster", width=140)
        self.sum_pending.column("district", width=140)
        self.sum_pending.column("bills", width=70, anchor="e")
        self.sum_pending.column("grand", width=120, anchor="e")
        self.sum_pending.pack(fill="x")
        self.sum_pending_tot = ttk.Label(tab_sum, text="Pending total: —", font=("Segoe UI", 9, "bold"))
        self.sum_pending_tot.pack(anchor="e", pady=(2, 8))

        ttk.Label(tab_sum, text="Forwarded / Returned", font=("Segoe UI", 10, "bold")).pack(anchor="w", pady=(4, 2))
        acols = ("date", "cluster", "district", "fwd", "ret", "fwd_amt", "ret_grand", "bills")
        self.sum_act = ttk.Treeview(tab_sum, columns=acols, show="headings", height=6)
        for c, t, w in (
            ("date", "Date", 90),
            ("cluster", "Cluster", 110),
            ("district", "District", 110),
            ("fwd", "Forward", 70),
            ("ret", "Return", 70),
            ("fwd_amt", "Forward amount", 110),
            ("ret_grand", "Return Grand Total", 130),
            ("bills", "Total bills", 80),
        ):
            self.sum_act.heading(c, text=t)
            self.sum_act.column(c, width=w, anchor="e" if c not in ("date", "cluster", "district") else "w")
        self.sum_act.pack(fill="both", expand=True)
        self.sum_act_tot = ttk.Label(tab_sum, text="Period total: —", font=("Segoe UI", 9, "bold"))
        self.sum_act_tot.pack(anchor="e", pady=4)

        loghead = ttk.Frame(right)
        loghead.pack(fill="x")
        ttk.Label(loghead, text="LOG", font=("Segoe UI", 10, "bold")).pack(side="left")
        ttk.Label(loghead, text="CAPTCHA + SIGN IN", style="Hint.TLabel").pack(side="left", padx=8)
        self.status2 = ttk.Label(loghead, text="")
        self.status2.pack(side="right")
        logf = ttk.Frame(right)
        logf.pack(fill="both", expand=True, pady=(4, 0))
        self.log = tk.Text(
            logf,
            wrap="word",
            font=("Consolas", 9),
            bg="#111318",
            fg="#d7dbdf",
            insertbackground="#fff",
            relief="flat",
            padx=8,
            pady=8,
        )
        lsb = ttk.Scrollbar(logf, orient="vertical", command=self.log.yview)
        self.log.pack(side="left", fill="both", expand=True)
        lsb.pack(side="right", fill="y")
        self.log.configure(yscrollcommand=lsb.set)
        self.log.tag_configure("err", foreground="#ff6b6b")
        self.log.tag_configure("ok", foreground="#69db7c")

        self.refresh_tree()
        self.after(200, self.drain_log)
        self.after(400, self.refresh_upload)
        self.after(600, self.refresh_summary)
        if not self.accounts:
            self.after(400, lambda: self.write(
                "Add at least one ID. Portal passwords are encrypted with the master password.\n"
                "If an ID has no password, open Edit and save the portal password again.\n"
            ))

    def _apply_style(self):
        s = ttk.Style(self)
        try:
            s.theme_use("clam")
        except Exception:
            pass
        bg = "#f4f5f7"
        s.configure(".", background=bg, font=("Segoe UI", 9))
        s.configure("TFrame", background=bg)
        s.configure("TLabel", background=bg, font=("Segoe UI", 9))
        s.configure("Title.TLabel", font=("Segoe UI", 18, "bold"), foreground="#14213d", background=bg)
        s.configure("Hint.TLabel", foreground="#5c6370", font=("Segoe UI", 9), background=bg)
        s.configure("TLabelframe", background=bg)
        s.configure("TLabelframe.Label", background=bg, font=("Segoe UI", 9, "bold"))
        s.configure("TButton", padding=(10, 6), font=("Segoe UI", 9))
        s.configure("Accent.TButton", background="#e85d04", foreground="#ffffff", padding=(12, 7), font=("Segoe UI", 9, "bold"))
        s.map("Accent.TButton", background=[("active", "#c2410c"), ("disabled", "#d1d5db")], foreground=[("disabled", "#888")])
        s.configure("Green.TButton", background="#2d6a4f", foreground="#ffffff", padding=(12, 7), font=("Segoe UI", 9, "bold"))
        s.map("Green.TButton", background=[("active", "#1b4332"), ("disabled", "#d1d5db")], foreground=[("disabled", "#888")])
        s.configure("TNotebook", background=bg)
        s.configure("TNotebook.Tab", padding=(14, 7), font=("Segoe UI", 10, "bold"))
        s.configure("Treeview", rowheight=26, font=("Segoe UI", 9), fieldbackground="#fff")
        s.configure("Treeview.Heading", font=("Segoe UI", 9, "bold"))
        s.configure("TCheckbutton", background=bg, font=("Segoe UI", 9))
        s.configure("TPanedwindow", background=bg)

    def refresh_tree(self):
        for i in self.tree.get_children():
            self.tree.delete(i)
        for idx, a in enumerate(self.accounts):
            self.tree.insert(
                "",
                "end",
                iid=str(idx),
                values=(
                    "YES" if a.get("enabled", True) else "no",
                    a.get("name") or "",
                    a.get("user") or "",
                    a.get("last_run") or "-",
                    a.get("last_session") or "-",
                    a.get("last_update") or "-",
                ),
            )

    def selected_index(self) -> int | None:
        sel = self.tree.selection()
        if not sel:
            return None
        return int(sel[0])

    def refresh_summary(self):
        from activity import action_rows, fmt_money, load_events, pending_rows
        dfrom = (self.sum_from.get() or "").strip()
        dto = (self.sum_to.get() or "").strip()
        events = load_events()
        for tree in (self.sum_pending, self.sum_act):
            for i in tree.get_children():
                tree.delete(i)
        pending = pending_rows(events, dfrom, dto)
        pb = 0
        pg = 0.0
        for r in pending:
            pb += r["bills"]
            pg += r["grand"]
            self.sum_pending.insert(
                "",
                "end",
                values=(r["cluster"], r["district"], r["bills"], fmt_money(r["grand"])),
            )
        self.sum_pending_tot.config(text=f"Pending total: {pb} bills   Grand Total {fmt_money(pg)}")
        acts = action_rows(events, dfrom, dto)
        tf = tr = 0
        fa = rg = 0.0
        for r in acts:
            tf += r["forward"]
            tr += r["returned"]
            fa += r["fwd_amt"]
            rg += r["ret_grand"]
            self.sum_act.insert(
                "",
                "end",
                values=(
                    r["date"],
                    r["cluster"],
                    r["district"],
                    r["forward"],
                    r["returned"],
                    fmt_money(r["fwd_amt"]),
                    fmt_money(r["ret_grand"]),
                    r["forward"] + r["returned"],
                ),
            )
        self.sum_act_tot.config(
            text=f"Period: Forward {tf} ({fmt_money(fa)})   Return {tr} (Grand Total {fmt_money(rg)})   Bills {tf + tr}"
        )

    def reset_batch(self):
        self.batch_var.set("")
        s = load_settings()
        s["batch_folder"] = ""
        save_settings(s)
        for i in self.up_tree.get_children():
            self.up_tree.delete(i)
        self.up_summary.config(text="Batch cleared. Browse a folder, then Scan Excel + files.")
        self.write("Batch folder and Excel list reset.\n")

    def pick_batch(self):
        start = self.batch_var.get().strip() or str(Path.home() / "Downloads")
        chosen = filedialog.askdirectory(title="Batch folder (Upload Template + PDFs)", initialdir=start)
        if chosen:
            self.batch_var.set(chosen)
            s = load_settings()
            s["batch_folder"] = chosen
            save_settings(s)
            self.refresh_upload()

    def refresh_days(self):
        self.refresh_upload()

    def refresh_sessions(self):
        self.refresh_upload()

    def refresh_upload(self):
        for i in self.up_tree.get_children():
            self.up_tree.delete(i)
        folder = Path(self.batch_var.get().strip()) if self.batch_var.get().strip() else None
        if not folder or not folder.exists():
            self.up_summary.config(text="Browse to the Batch folder (Upload Template.xlsx + PDFs).")
            return
        try:
            from tpi_uploader import load_batch
            rows = load_batch(folder)
        except Exception as exc:
            self.up_summary.config(text=f"Could not read batch: {exc}")
            return
        n_files = 0
        n_err = 0
        for idx, b in enumerate(rows):
            fill = b.get("fill") or {}
            ok_names = []
            bad = []
            for f in b.get("files") or []:
                fp = Path(f)
                n_files += 1
                try:
                    sz = fp.stat().st_size
                except Exception:
                    sz = 0
                if sz > 7 * 1024 * 1024:
                    bad.append(f"{fp.name} ({sz/1024/1024:.1f} MB)")
                    n_err += 1
                else:
                    ok_names.append(fp.name)
            warn = "OK"
            if not (b.get("files") or []):
                warn = f"ERROR: no PDF for letter {fill.get('letter') or '?'}"
                n_err += 1
            elif bad:
                warn = "ERROR >7MB: " + "; ".join(bad)
            names = ", ".join(ok_names) if ok_names else "-"
            self.up_tree.insert(
                "",
                "end",
                iid=str(idx),
                values=(
                    "☑",
                    b.get("scheme_id") or "",
                    b.get("mb_no") or "",
                    fill.get("letter") or "",
                    fill.get("status") or fill.get("bill_type") or "",
                    names,
                    warn,
                ),
                tags=("bad",) if bad else ("ok",),
            )
        try:
            self.up_tree.tag_configure("bad", foreground="#b00020")
        except Exception:
            pass
        msg = f"{len(rows)} Excel row(s) · {n_files} file(s) matched"
        if n_err:
            msg += f" · ERROR: {n_err} file(s) over 7 MB (will not upload)"
        self.up_summary.config(text=msg)
        if n_err:
            self.write(f"SCAN ERROR: {n_err} file(s) larger than 7 MB — skipped.\n")

    def dl_steps_all(self):
        self.dl_excel.set(True)
        self.dl_loa.set(True)
        self.dl_comments.set(True)
        self.dl_mb.set(True)
        self.dl_docs.set(True)

    def dl_steps_none(self):
        self.dl_excel.set(False)
        self.dl_loa.set(False)
        self.dl_comments.set(False)
        self.dl_mb.set(False)
        self.dl_docs.set(False)

    def steps_all(self):
        self.step_files.set(True)
        self.step_fill.set(True)
        self.step_dsc.set(True)
        self.step_action.set(True)

    def steps_none(self):
        self.step_files.set(False)
        self.step_fill.set(False)
        self.step_dsc.set(False)
        self.step_action.set(False)

    def _up_click(self, event):
        row = self.up_tree.identify_row(event.y)
        col = self.up_tree.identify_column(event.x)
        if row and col == "#1":
            vals = list(self.up_tree.item(row, "values"))
            vals[0] = "☐" if vals[0] == "☑" else "☑"
            self.up_tree.item(row, values=vals)
            return "break"

    def up_select_all(self):
        for iid in self.up_tree.get_children():
            vals = list(self.up_tree.item(iid, "values"))
            vals[0] = "☑"
            self.up_tree.item(iid, values=vals)

    def up_select_none(self):
        for iid in self.up_tree.get_children():
            vals = list(self.up_tree.item(iid, "values"))
            vals[0] = "☐"
            self.up_tree.item(iid, values=vals)

    def _checked_keys(self, all_rows: bool) -> list:
        keys = []
        for iid in self.up_tree.get_children():
            vals = self.up_tree.item(iid, "values")
            if all_rows or (vals and vals[0] == "☑"):
                keys.append(f"{vals[1]}|{vals[2]}")
        return keys

    def run_upload_scan(self):
        self.refresh_upload()
        bad = []
        for iid in self.up_tree.get_children():
            w = str(self.up_tree.item(iid, "values")[6] or "")
            if w and w != "OK":
                vals = self.up_tree.item(iid, "values")
                bad.append(f"{vals[1]} {vals[2]}: {w}")
        if bad:
            messagebox.showerror("Scan errors", "Fix these before upload:\n\n" + "\n".join(bad[:20]))
            self.write("SCAN FAILED: " + "; ".join(bad) + "\n")
        else:
            n = len(self.up_tree.get_children())
            messagebox.showinfo("Scan OK", f"{n} Excel row(s) — matching PDFs found, all under 7 MB.")
            self.write(f"Scan OK: {n} row(s), all PDFs under 7 MB.\n")

    def run_upload(self):
        self._start_upload(all_rows=False)

    def run_upload_all(self):
        self._start_upload(all_rows=True)

    def _start_upload(self, scan_only: bool = False, all_rows: bool = False):
        i = self.selected_index()
        if i is None:
            messagebox.showinfo("Select", "Select a portal ID in the list above.")
            return
        if not UPLOADER.exists():
            messagebox.showerror("Missing", "tpi_uploader.py was not found.")
            return
        batch = self.batch_var.get().strip()
        if not batch or not Path(batch).exists():
            messagebox.showinfo("Batch", "Browse to the batch folder (Upload Template.xlsx + PDFs).")
            return
        if not self.up_tree.get_children():
            self.refresh_upload()
        keys = self._checked_keys(all_rows)
        if not keys:
            messagebox.showinfo("Rows", "Tick at least one row, or use Upload all rows.")
            return
        over = []
        for iid in self.up_tree.get_children():
            vals = self.up_tree.item(iid, "values")
            if vals[6] not in ("OK", "") and (all_rows or vals[0] == "☑"):
                over.append(str(vals[6]))
        if over:
            messagebox.showerror(
                "7 MB limit",
                "One or more files are larger than 7 MB.\nThose files will be skipped.\n\n" + "\n".join(over[:8]),
            )
        self._start([self.accounts[i]], upload=True, scan_only=False, only_keys=keys)

    def add_acc(self):
        d = AccountDialog(self, "New ID")
        self.wait_window(d)
        if d.result:
            d.result["last_run"] = ""
            self.accounts.append(d.result)
            save_accounts(self.accounts)
            self.refresh_tree()

    def edit_acc(self):
        i = self.selected_index()
        if i is None:
            messagebox.showinfo("Select", "Select an ID from the list first.")
            return
        d = AccountDialog(self, "Edit ID", self.accounts[i])
        self.wait_window(d)
        if d.result:
            d.result["last_run"] = self.accounts[i].get("last_run") or ""
            d.result["enabled"] = self.accounts[i].get("enabled", True)
            self.accounts[i] = d.result
            save_accounts(self.accounts)
            self.refresh_tree()

    def del_acc(self):
        i = self.selected_index()
        if i is None:
            return
        if messagebox.askyesno("Delete", f"Remove {self.accounts[i].get('name')}?"):
            self.accounts.pop(i)
            save_accounts(self.accounts)
            self.refresh_tree()

    def toggle_acc(self):
        i = self.selected_index()
        if i is None:
            return
        self.accounts[i]["enabled"] = not self.accounts[i].get("enabled", True)
        save_accounts(self.accounts)
        self.refresh_tree()

    def pick_master(self):
        start = self.master_var.get().strip() or self.save_var.get().strip() or str(Path.home())
        chosen = filedialog.askopenfilename(
            title="Load previous master Excel",
            initialdir=str(Path(start).parent if Path(start).suffix else start),
            filetypes=[("Excel", "*.xlsx"), ("All", "*.*")],
        )
        if chosen:
            self.master_var.set(chosen)
            s = load_settings()
            s["master_path"] = chosen
            save_settings(s)
            self.write(f"Master Excel: {chosen}\n")

    def pick_save(self):
        cur = self.save_var.get().strip() or str(Path.home())
        chosen = filedialog.askdirectory(title="Kautilya E-MB folder chuno", initialdir=cur)
        if chosen:
            self.save_var.set(chosen)
            s = load_settings()
            s["save_root"] = chosen
            save_settings(s)
            self.refresh_days()

    def save_root(self) -> Path | None:
        p = self.save_var.get().strip()
        if not p:
            return None
        return Path(p)

    def open_dl(self):
        p = self.save_root()
        if not p:
            messagebox.showinfo("Path", "Choose the Kautilya E-MB folder with Browse.")
            return
        p.mkdir(parents=True, exist_ok=True)
        os.startfile(str(p))

    def write(self, msg: str):
        tag = ""
        low = msg.lower()
        m = re.search(r"PROGRESS\s+(\d+)/(\d+)", msg)
        if m:
            cur, tot = int(m.group(1)), max(int(m.group(2)), 1)
            try:
                self.prog["maximum"] = tot
                self.prog["value"] = cur
                self.prog_lbl.config(text=f"{cur} (Download in progress) / {tot} (Total bills)")
                self.status.config(text=f"{cur}/{tot}")
            except Exception:
                pass
        if "error" in low or "fail" in low or "scan failed" in low:
            tag = "err"
        elif "ok" in low or "success" in low or "uploaded" in low or "finished" in low:
            tag = "ok"
        self.log.insert("end", msg, tag)
        self.log.see("end")

    def drain_log(self):
        try:
            while True:
                self.write(self.log_q.get_nowait())
        except queue.Empty:
            pass
        self.after(200, self.drain_log)

    def run_all(self):
        rows = [a for a in self.accounts if a.get("enabled", True)]
        if not rows:
            messagebox.showinfo("IDs", "No ID is enabled. Add or enable an account first.")
            return
        self._start(rows)

    def run_one(self):
        i = self.selected_index()
        if i is None:
            messagebox.showinfo("Select", "Select an ID from the list first.")
            return
        self._start([self.accounts[i]])

    def _start(self, rows: list[dict], upload: bool = False, scan_only: bool = False, only_keys: list | None = None):
        if self.running:
            messagebox.showinfo("Busy", "A job is already running.")
            return
        if upload and not UPLOADER.exists():
            messagebox.showerror("Missing", "tpi_uploader.py was not found.")
            return
        if not upload and not SCRIPT.exists():
            messagebox.showerror(
                "Missing",
                f"tpi_mb_downloader.py was not found.\n\nLooked in:\n{ROOT}\n\n"
                "Put dashboard.py, tpi_mb_downloader.py, tpi_uploader.py and activity.py in the SAME folder.",
            )
            return
        root = self.save_root()
        if not root:
            messagebox.showinfo("Path", "Set the Kautilya E-MB folder with Browse first.")
            self.pick_save()
            root = self.save_root()
            if not root:
                return
        s = load_settings()
        s["save_root"] = str(root)
        save_settings(s)
        self.running = True
        self.btn_all.state(["disabled"])
        self.btn_one.state(["disabled"])
        try:
            self.btn_up.state(["disabled"])
            self.btn_up_scan.state(["disabled"])
            self.btn_up_all.state(["disabled"])
        except Exception:
            pass
        self.status.config(text="Running… complete CAPTCHA in the browser")
        threading.Thread(target=self._worker, args=(rows, upload, scan_only, only_keys or []), daemon=True).start()

    def _worker(self, rows: list[dict], upload: bool = False, scan_only: bool = False, only_keys: list | None = None):
        try:
            save_root = self.save_root()
            if upload:
                batch = self.batch_var.get().strip()
                if not batch:
                    self.log_q.put("ERROR: batch folder is not set\n")
                    return
                self.log_q.put(f"Upload batch: {batch}\n")
                for i, acc in enumerate(rows, 1):
                    name = acc.get("name") or acc.get("user")
                    self.log_q.put(f"\n======== UPLOAD [{i}/{len(rows)}] {name} ========\n")
                    env = os.environ.copy()
                    env["PORTAL_USER"] = acc.get("user") or ""
                    env["PORTAL_PASS"] = acc.get("pass") or ""
                    if save_root:
                        env["DOWNLOAD_DIR"] = str(save_root)
                    env["TPI_BATCH"] = batch
                    env["TPI_ONLY"] = ";".join(only_keys or [])
                    env["TPI_CLUSTER"] = name
                    steps = []
                    if self.step_files.get():
                        steps.append("files")
                    if self.step_fill.get():
                        steps.append("fill")
                    if self.step_dsc.get():
                        steps.append("dsc")
                    if self.step_action.get():
                        steps.append("action")
                    env["TPI_STEPS"] = ",".join(steps)
                    env["PYTHONUNBUFFERED"] = "1"
                    cmd = [str(PY), str(UPLOADER), "--headed", "--batch", batch]
                    if scan_only:
                        cmd.append("--scan-only")
                    proc = subprocess.Popen(
                        cmd,
                        cwd=str(ROOT),
                        env=env,
                        stdout=subprocess.PIPE,
                        stderr=subprocess.STDOUT,
                        text=True,
                        bufsize=1,
                    )
                    assert proc.stdout is not None
                    for line in proc.stdout:
                        self.log_q.put(line)
                    rc = proc.wait()
                    self.log_q.put(f"--- upload {name} exit={rc} ---\n")
                self.log_q.put("\nUpload job finished.\n")
                return
            if not save_root:
                self.log_q.put("ERROR: save path is not set\n")
                return
            save_root.mkdir(parents=True, exist_ok=True)
            day = datetime.now().strftime("%d.%m.%Y")
            day_dir = save_root / day
            n = 1
            if day_dir.exists():
                nums = []
                for p in day_dir.iterdir():
                    m = re.search(r"session\s*(\d+)", p.name, re.I) if p.is_dir() else None
                    if m:
                        nums.append(int(m.group(1)))
                if nums:
                    n = max(nums) + 1
            session = f"Session {n}"
            self.log_q.put(f"Save: {save_root} \\ {day} \\ {session} \\ Cluster \\ bill  (all IDs share this session)\n")
            for i, acc in enumerate(rows, 1):
                name = acc.get("name") or acc.get("user")
                self.log_q.put(f"\n======== [{i}/{len(rows)}] {name}  |  CAPTCHA + SIGN IN ========\n")
                env = os.environ.copy()
                env["PORTAL_USER"] = acc.get("user") or ""
                env["PORTAL_PASS"] = acc.get("pass") or ""
                env["DOWNLOAD_DIR"] = str(save_root)
                env["TPI_SESSION"] = session
                env["TPI_CLUSTER"] = name
                mpath = self.master_var.get().strip()
                if not mpath:
                    mpath = str(Path(save_root) / "tpi_master.xlsx")
                    self.master_var.set(mpath)
                env["TPI_MASTER"] = mpath
                env["ACTION_DELAY"] = "0.15"
                env["PYTHONUNBUFFERED"] = "1"
                env["TPI_GSHEET"] = self.gsheet_var.get().strip()
                dls = []
                if self.dl_excel.get():
                    dls.append("excel")
                if getattr(self, "dl_loa", None) and self.dl_loa.get():
                    dls.append("loa")
                if getattr(self, "dl_comments", None) and self.dl_comments.get():
                    dls.append("comments")
                if self.dl_mb.get():
                    dls.append("mb")
                if self.dl_docs.get():
                    dls.append("docs")
                if not dls:
                    self.log_q.put("ERROR: tick at least one Download option\n")
                    return
                env["TPI_DL_STEPS"] = ",".join(dls)
                self.log_q.put(f"Steps: {env['TPI_DL_STEPS']}\n")
                cmd = [str(PY), str(SCRIPT), "--headed", "--with-abstract", "--session", session]
                proc = subprocess.Popen(
                    cmd,
                    cwd=str(ROOT),
                    env=env,
                    stdout=subprocess.PIPE,
                    stderr=subprocess.STDOUT,
                    text=True,
                    bufsize=1,
                )
                assert proc.stdout is not None
                for line in proc.stdout:
                    self.log_q.put(line)
                rc = proc.wait()
                acc["last_run"] = datetime.now().strftime("%Y-%m-%d %H:%M")
                acc["last_session"] = session
                acc["last_update"] = datetime.now().strftime("%Y-%m-%d %H:%M")
                save_accounts(self.accounts)
                self.log_q.put(f"--- {name} finished exit={rc} ---\n")
            self.log_q.put(f"\nComplete. Folder: {save_root}\\{day}\\{session}\n")
        except Exception as exc:
            self.log_q.put(f"\nDASHBOARD ERROR: {exc}\n")
        finally:
            self.running = False
            self.after(0, self._unlock)

    def _unlock(self):
        self.btn_all.state(["!disabled"])
        self.btn_one.state(["!disabled"])
        try:
            self.btn_up.state(["!disabled"])
            self.btn_up_scan.state(["!disabled"])
            self.btn_up_all.state(["!disabled"])
        except Exception:
            pass
        self.status.config(text="Ready")
        self.refresh_tree()
        try:
            self.refresh_summary()
        except Exception:
            pass
        try:
            s = load_settings()
            s["gsheet_url"] = self.gsheet_var.get().strip()
            s["master_path"] = self.master_var.get().strip()
            save_settings(s)
        except Exception:
            pass


if __name__ == "__main__":
    app = App()
    try:
        if app.winfo_exists():
            app.mainloop()
    except tk.TclError:
        pass
