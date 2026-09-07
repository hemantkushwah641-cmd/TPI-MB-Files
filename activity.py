"""Local activity log for the Summary tab (pending / forward / return)."""
from __future__ import annotations

import json
from datetime import datetime
from pathlib import Path

ROOT = Path(__file__).resolve().parent
LOG = ROOT / ".tpidata" / "activity.jsonl"


def log_event(**kw) -> None:
    kw["ts"] = kw.get("ts") or datetime.now().isoformat(timespec="seconds")
    LOG.parent.mkdir(parents=True, exist_ok=True)
    with LOG.open("a", encoding="utf-8") as f:
        f.write(json.dumps(kw, ensure_ascii=False) + "\n")


def load_events() -> list[dict]:
    if not LOG.exists():
        return []
    out = []
    for line in LOG.read_text(encoding="utf-8").splitlines():
        line = line.strip()
        if not line:
            continue
        try:
            out.append(json.loads(line))
        except Exception:
            continue
    return out


def parse_money(val) -> float:
    if val is None:
        return 0.0
    s = str(val).replace(",", "").replace("₹", "").replace("Rs", "").strip()
    s = s.replace(" ", "")
    if not s or s == "-":
        return 0.0
    try:
        return float(s)
    except Exception:
        return 0.0


def fmt_money(n: float) -> str:
    try:
        return f"{n:,.2f}"
    except Exception:
        return "0.00"


def in_range(ts: str, dfrom: str, dto: str) -> bool:
    day = (ts or "")[:10]
    if dfrom and day < dfrom:
        return False
    if dto and day > dto:
        return False
    return True


def pending_rows(events: list[dict], dfrom: str, dto: str) -> list[dict]:
    """Last download/list of each bill with no later forward/return."""
    latest: dict[str, dict] = {}
    for e in events:
        key = f"{e.get('scheme_id')}|{e.get('mb_no')}"
        if e.get("kind") == "download":
            latest[key] = dict(e)
            latest[key]["_done"] = False
        elif e.get("kind") == "upload" and e.get("action") in ("forward", "return"):
            if key in latest:
                latest[key]["_done"] = True
            else:
                latest[key] = dict(e)
                latest[key]["_done"] = True
    groups: dict[tuple, dict] = {}
    for rec in latest.values():
        if rec.get("_done"):
            continue
        cl = rec.get("cluster") or "-"
        dist = rec.get("district") or rec.get("cluster") or "-"
        g = groups.setdefault((cl, dist), {"cluster": cl, "district": dist, "bills": 0, "grand": 0.0})
        g["bills"] += 1
        g["grand"] += parse_money(rec.get("grand_total"))
    return sorted(groups.values(), key=lambda x: (x["cluster"], x["district"]))


def action_rows(events: list[dict], dfrom: str, dto: str) -> list[dict]:
    """Forward / return counts and amounts by date + cluster + district."""
    groups: dict[tuple, dict] = {}
    for e in events:
        if e.get("kind") != "upload":
            continue
        act = (e.get("action") or "").lower()
        if act not in ("forward", "return"):
            continue
        if not in_range(e.get("ts") or "", dfrom, dto):
            continue
        day = (e.get("ts") or "")[:10]
        cl = e.get("cluster") or "-"
        dist = e.get("district") or "-"
        g = groups.setdefault(
            (day, cl, dist),
            {
                "date": day,
                "cluster": cl,
                "district": dist,
                "forward": 0,
                "returned": 0,
                "fwd_amt": 0.0,
                "ret_grand": 0.0,
                "grand": 0.0,
            },
        )
        grand = parse_money(e.get("grand_total"))
        amt = parse_money(e.get("amount"))
        g["grand"] += grand
        if act == "return":
            g["returned"] += 1
            g["ret_grand"] += grand  # TPI amount is 0.00 — use Grand Total
        else:
            g["forward"] += 1
            g["fwd_amt"] += amt if amt else grand
    return sorted(groups.values(), key=lambda x: (x["date"], x["cluster"], x["district"]))
