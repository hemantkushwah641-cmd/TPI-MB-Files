"""Usage log: which PC / Windows user / version, and daily activity."""
from __future__ import annotations

import csv
import getpass
import json
import os
import socket
import uuid
from collections import defaultdict
from datetime import datetime
from pathlib import Path

from version import APP_VERSION

ROOT = Path(__file__).resolve().parent
LOCAL = ROOT / ".tpidata" / "usage.jsonl"
CSV_HEADERS = ["DateTime", "Computer", "WindowsUser", "Version", "Action", "Detail"]


def whoami() -> dict:
    return {
        "version": APP_VERSION,
        "computer": os.environ.get("COMPUTERNAME") or socket.gethostname() or "-",
        "win_user": os.environ.get("USERNAME") or getpass.getuser() or "-",
    }


def _share_dir() -> Path | None:
    settings = ROOT / ".tpidata" / "settings.json"
    try:
        if settings.exists():
            data = json.loads(settings.read_text(encoding="utf-8"))
            p = (data.get("usage_share") or "").strip()
            if p:
                path = Path(p)
                path.mkdir(parents=True, exist_ok=True)
                return path
    except Exception:
        pass
    return None


def log_usage(action: str, detail: str = "") -> None:
    rec = {
        **whoami(),
        "ts": datetime.now().isoformat(timespec="seconds"),
        "action": str(action or "")[:80],
        "detail": str(detail or "")[:300],
    }
    try:
        LOCAL.parent.mkdir(parents=True, exist_ok=True)
        with LOCAL.open("a", encoding="utf-8") as f:
            f.write(json.dumps(rec, ensure_ascii=False) + "\n")
    except Exception:
        pass
    share = _share_dir()
    if not share:
        return
    try:
        inbox = share / "inbox"
        inbox.mkdir(parents=True, exist_ok=True)
        name = f"{datetime.now().strftime('%Y%m%d_%H%M%S')}_{rec['computer']}_{uuid.uuid4().hex[:6]}.json"
        (inbox / name).write_text(json.dumps(rec, ensure_ascii=False), encoding="utf-8")
    except Exception:
        pass
    try:
        csv_path = share / "tpi_usage.csv"
        new = not csv_path.exists()
        with csv_path.open("a", encoding="utf-8", newline="") as f:
            w = csv.writer(f)
            if new:
                w.writerow(CSV_HEADERS)
            w.writerow(
                [rec["ts"], rec["computer"], rec["win_user"], rec["version"], rec["action"], rec["detail"]]
            )
    except Exception:
        pass


def _read_jsonl(path: Path) -> list[dict]:
    if not path.exists():
        return []
    out = []
    try:
        text = path.read_text(encoding="utf-8")
    except Exception:
        return []
    for line in text.splitlines():
        line = line.strip()
        if not line:
            continue
        try:
            out.append(json.loads(line))
        except Exception:
            continue
    return out


def load_usage(dfrom: str = "", dto: str = "") -> list[dict]:
    rows = _read_jsonl(LOCAL)
    share = _share_dir()
    if share:
        inbox = share / "inbox"
        if inbox.exists():
            for p in inbox.glob("*.json"):
                try:
                    rows.append(json.loads(p.read_text(encoding="utf-8")))
                except Exception:
                    continue
        csv_path = share / "tpi_usage.csv"
        if csv_path.exists():
            try:
                with csv_path.open(encoding="utf-8", newline="") as f:
                    r = csv.DictReader(f)
                    for rec in r:
                        rows.append(
                            {
                                "ts": rec.get("DateTime") or "",
                                "computer": rec.get("Computer") or "",
                                "win_user": rec.get("WindowsUser") or "",
                                "version": rec.get("Version") or "",
                                "action": rec.get("Action") or "",
                                "detail": rec.get("Detail") or "",
                            }
                        )
            except Exception:
                pass
    seen = set()
    uniq = []
    for r in rows:
        key = (r.get("ts"), r.get("computer"), r.get("win_user"), r.get("action"), r.get("detail"))
        if key in seen:
            continue
        seen.add(key)
        day = (r.get("ts") or "")[:10]
        if dfrom and day < dfrom:
            continue
        if dto and day > dto:
            continue
        uniq.append(r)
    uniq.sort(key=lambda x: x.get("ts") or "")
    return uniq


def daily_summary(events: list[dict]) -> list[dict]:
    groups: dict[tuple, dict] = defaultdict(
        lambda: {"opens": 0, "downloads": 0, "uploads": 0, "other": 0, "last": "", "version": ""}
    )
    for e in events:
        day = (e.get("ts") or "")[:10]
        key = (day, e.get("computer") or "-", e.get("win_user") or "-")
        g = groups[key]
        g["version"] = e.get("version") or g["version"]
        g["last"] = e.get("ts") or g["last"]
        act = (e.get("action") or "").lower()
        if act in ("app_open", "open"):
            g["opens"] += 1
        elif "download" in act:
            g["downloads"] += 1
        elif "upload" in act:
            g["uploads"] += 1
        else:
            g["other"] += 1
    out = []
    for (day, computer, user), g in sorted(groups.items(), reverse=True):
        out.append(
            {
                "date": day,
                "computer": computer,
                "win_user": user,
                "version": g["version"],
                "opens": g["opens"],
                "downloads": g["downloads"],
                "uploads": g["uploads"],
                "other": g["other"],
                "last": (g["last"] or "")[11:19],
            }
        )
    return out


def pc_status(events: list[dict]) -> list[dict]:
    """One row per computer + Windows user."""
    now = datetime.now()
    today = now.strftime("%Y-%m-%d")
    acc: dict[tuple, dict] = {}
    for e in events:
        key = (e.get("computer") or "-", e.get("win_user") or "-")
        g = acc.setdefault(
            key,
            {
                "computer": key[0],
                "win_user": key[1],
                "version": "",
                "last": "",
                "opens": 0,
                "downloads": 0,
                "uploads": 0,
            },
        )
        ts = e.get("ts") or ""
        if ts >= g["last"]:
            g["last"] = ts
            g["version"] = e.get("version") or g["version"]
        if ts[:10] != today:
            continue
        act = (e.get("action") or "").lower()
        if act in ("app_open", "open"):
            g["opens"] += 1
        elif "download" in act:
            g["downloads"] += 1
        elif "upload" in act:
            g["uploads"] += 1
    out = []
    for g in acc.values():
        status = "—"
        try:
            last = datetime.fromisoformat(g["last"])
            mins = (now - last).total_seconds() / 60.0
            if mins <= 15:
                status = "Active"
            elif last.strftime("%Y-%m-%d") == today:
                status = "Today"
            else:
                status = "Idle"
        except Exception:
            status = "—"
        g["status"] = status
        out.append(g)
    out.sort(key=lambda r: r.get("last") or "", reverse=True)
    return out
