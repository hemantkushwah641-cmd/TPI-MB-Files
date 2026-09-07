"""
UPJN-Kautilya TPI Civil — Measurement Book downloader.
Credentials stay in local .env next to this script.
"""

from __future__ import annotations

import argparse
import json
import os
import re
import sys
from collections import defaultdict
from datetime import datetime
from pathlib import Path

from dotenv import load_dotenv
from openpyxl import Workbook
from openpyxl.styles import Font
from playwright.sync_api import TimeoutError as PwTimeout
from playwright.sync_api import sync_playwright

ROOT = Path(__file__).resolve().parent
os.chdir(ROOT)
load_dotenv(ROOT / ".env")
load_dotenv(ROOT / ".env.txt")
load_dotenv(ROOT / ".env.example")  # fallback if user filled example file

BASE_URL = os.getenv("PORTAL_BASE_URL", "https://upjnr-kautilya.in").rstrip("/")
LIST_PATH = "/measurement/normal/third-party-inspection/record"
USER = os.getenv("PORTAL_USER", "").strip()
PASSWORD = os.getenv("PORTAL_PASS", "").strip()
DOWNLOAD_DIR = Path(os.getenv("DOWNLOAD_DIR") or os.getenv("SAVE_ROOT") or "downloads").resolve()
ACTION_DELAY = float(os.getenv("ACTION_DELAY", "0.4"))
DEBUG_DIR = (ROOT / ".tpidata" / "debug")
INDEX_PATH = DOWNLOAD_DIR / ".tpi_index.json"


def log(msg: str) -> None:
    print(f"[{datetime.now().strftime('%H:%M:%S')}] {msg}", flush=True)


def today_stamp() -> str:
    return datetime.now().strftime("%d.%m.%Y")


def win_name(text: str, max_len: int = 80) -> str:
    """OneDrive-style: spaces/hyphens rakho, Windows-illegal hatao."""
    t = str(text or "").replace("\u00a0", " ")
    t = re.sub(r"[\r\n\t]+", " ", t)
    t = t.replace("/", "-").replace("\\", "-")
    t = re.sub(r'[<>:"|?*]', "", t)
    t = re.sub(r" +", " ", t).strip(" .")
    return (t or "NA")[:max_len]


def district_folder(name: str) -> str:
    t = win_name(name, 60)
    t = re.sub(r"\s+District$", "", t, flags=re.I).strip()
    return t or "Unknown-District"


def next_session_no(root: Path, day: str) -> int:
    """Aaj ke existing Session N dekho, agla number do."""
    day_dir = root / day
    if not day_dir.exists():
        return 1
    nums: list[int] = []
    for p in day_dir.rglob("*"):
        if p.is_dir():
            m = re.search(r"^Session\s+(\d+)$", p.name, re.I)
            if m:
                nums.append(int(m.group(1)))
    return (max(nums) + 1) if nums else 1


def bill_folder_name(row: dict) -> str:
    """20013299_Tikatha - JJM SCHEME_M-S NCC LIMITED_RA-CIVIL-010_213-ED-2021-22-IV"""
    sid = win_name(row.get("scheme_id") or "id", 20)
    scheme = win_name(row.get("scheme") or "scheme", 70)
    contr = win_name(row.get("contractor") or "contractor", 50)
    mb = win_name((row.get("mb_no") or "RA-CIVIL-000").replace("/", "-"), 20)
    loa = win_name(row.get("loa_number") or row.get("list_ref") or sid, 40)
    name = f"{sid}_{scheme}_{contr}_{mb}_{loa}"
    extra = row.get("dup_suffix") or ""
    return (name + extra)[:180]


def bill_key(row: dict) -> str:
    sid = str(row.get("scheme_id") or "").strip()
    mb = str(row.get("mb_no") or "").strip().upper()
    dt = str(row.get("date_iso") or row.get("date") or "").strip()
    return f"{sid}|{mb}|{dt}"


def load_index() -> dict:
    if not INDEX_PATH.exists():
        return {}
    try:
        data = json.loads(INDEX_PATH.read_text(encoding="utf-8"))
        return data if isinstance(data, dict) else {}
    except Exception as exc:
        log(f"Index padhne mein error (naya index): {exc}")
        return {}


def save_index(index: dict) -> None:
    INDEX_PATH.parent.mkdir(parents=True, exist_ok=True)
    INDEX_PATH.write_text(json.dumps(index, indent=2, ensure_ascii=False), encoding="utf-8")


def disk_keys(root: Path) -> set[str]:
    """Pehle se Date/Session/District/schemeId_... folders — skip ke liye."""
    found: set[str] = set()
    if not root.exists():
        return found
    try:
        for p in root.rglob("*"):
            if not p.is_dir():
                continue
            m = re.match(r"^(\d{6,})_", p.name)
            if not m:
                continue
            sid = m.group(1)
            ra = re.search(r"RA[-_/]?CIVIL[-_/]?\d+", p.name, re.I)
            ra_n = (ra.group(0) if ra else "").upper().replace("_", "-").replace("/", "-")
            found.add(f"{sid}|{ra_n}")
            pdfs = list(p.glob("*SignedMB*.pdf")) + list(p.glob("*Signed*.pdf"))
            if pdfs:
                found.add(f"pdf:{sid}|{ra_n}")
    except Exception as exc:
        log(f"disk scan: {exc}")
    log(f"Disk pe pehle se bills: {len(found)}")
    return found


def already_on_disk(row: dict, keys: set[str]) -> bool:
    sid = str(row.get("scheme_id") or "").strip()
    mb = str(row.get("mb_no") or "").upper().replace("/", "-").replace("_", "-")
    if not sid:
        return False
    k = f"{sid}|{mb}"
    if k in keys or f"pdf:{k}" in keys:
        return True
    for existing in keys:
        if existing.startswith(f"{sid}|") and mb and mb in existing:
            return True
    return False


def already_downloaded(row: dict, index: dict) -> dict | None:
    rec = index.get(bill_key(row))
    if not rec:
        return None
    pdf = rec.get("signed_pdf") or ""
    if pdf and Path(pdf).exists():
        return rec
    return None


def safe_name(text: str, max_len: int = 60) -> str:
    text = str(text or "").replace("\u00a0", " ")
    text = re.sub(r"[\r\n\t]+", " ", text)
    text = re.sub(r"[^A-Za-z0-9._-]+", "_", text)
    text = re.sub(r"_+", "_", text).strip("._ ")
    return (text or "NA")[:max_len]


def parse_portal_date(text: str) -> str:
    text = re.sub(r"\s+", " ", (text or "").strip())
    for fmt in ("%d %b, %Y", "%d %B, %Y", "%d %b %Y", "%d-%m-%Y", "%d/%m/%Y", "%Y-%m-%d"):
        try:
            return datetime.strptime(text, fmt).strftime("%Y-%m-%d")
        except ValueError:
            continue
    return safe_name(text) or "unknown-date"


def ensure_dirs() -> None:
    DOWNLOAD_DIR.mkdir(parents=True, exist_ok=True)
    DEBUG_DIR.mkdir(parents=True, exist_ok=True)


def save_debug(page, name: str) -> None:
    path = DEBUG_DIR / f"{datetime.now().strftime('%H%M%S')}_{safe_name(name)}.png"
    try:
        page.screenshot(path=str(path), full_page=True)
        log(f"Debug screenshot: {path}")
    except Exception as exc:
        log(f"Screenshot failed: {exc}")


def bill_type_of(mb_no: str) -> str:
    t = (mb_no or "").upper()
    if re.search(r"\bE\s*&\s*M\b|\bEM\b|ELECTR", t):
        return "E&M"
    return "civil"


def write_tpi_output(rows: list[dict], path: Path) -> None:
    """Session Excel — tpi_output_HHMMSS jaisa."""
    wb = Workbook()
    ws = wb.active
    ws.title = "TPI Data"
    headers = [
        "LoA No.",
        "MB No.",
        "Measurement Date",
        "Contractor Name",
        "Scheme Name",
        "Scheme ID",
        "Expenditure Type",
        "State",
        "ID TYPE",
        "District Name",
        "TPI Date",
        "Bill Type",
        "Cluster",
    ]
    ws.append(headers)
    for cell in ws[1]:
        cell.font = Font(bold=True)
    id_type = (os.getenv("TPI_ID_TYPE") or "TL").strip() or "TL"
    tpi_date = datetime.now().strftime("%Y-%m-%d")
    for r in rows:
        mb = str(r.get("mb_no") or "")
        sid = str(r.get("scheme_id") or "")
        dist = str(r.get("district") or r.get("district_folder") or "").replace(" District", "").strip()
        cluster = str(r.get("cluster") or os.getenv("TPI_CLUSTER") or dist)
        ws.append([
            str(r.get("list_ref") or r.get("loa_number") or ""),
            mb,
            str(r.get("date") or r.get("measurement_date") or ""),
            str(r.get("contractor") or ""),
            str(r.get("scheme") or ""),
            sid,
            str(r.get("type") or "-"),
            str(r.get("state") or "Uttar Pradesh"),
            str(r.get("id_type") or id_type),
            dist or cluster,
            tpi_date,
            bill_type_of(mb),
            cluster,
        ])
    ws.auto_filter.ref = ws.dimensions
    ws.freeze_panes = "A2"
    for col in ws.columns:
        width = min(max(len(str(c.value or "")) for c in col) + 2, 55)
        ws.column_dimensions[col[0].column_letter].width = width
    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)
    log(f"TPI output Excel: {path}")


def merge_session_master(sess_dir: Path) -> Path | None:
    """Combine every tpi_output_*.xlsx in this session into tpi_output_MASTER.xlsx."""
    sess_dir = Path(sess_dir)
    if not sess_dir.exists():
        return None
    files = sorted(
        p for p in sess_dir.glob("tpi_output*.xlsx")
        if p.is_file() and "master" not in p.name.lower()
    )
    if not files:
        return None
    from openpyxl import load_workbook
    headers = None
    rows_out = []
    seen = set()
    for path in files:
        try:
            wb = load_workbook(path, data_only=True, read_only=True)
            ws = wb.active
            data = list(ws.iter_rows(values_only=True))
            wb.close()
        except Exception as exc:
            log(f"  merge skip {path.name}: {exc}")
            continue
        if not data:
            continue
        if headers is None:
            headers = [str(c or "") for c in data[0]]
        for raw in data[1:]:
            key = (str(raw[5] if len(raw) > 5 else ""), str(raw[1] if len(raw) > 1 else ""))
            if key in seen:
                continue
            seen.add(key)
            rows_out.append(list(raw))
    if not headers:
        return None
    dest = sess_dir / "tpi_output_MASTER.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.title = "TPI Data"
    ws.append(headers)
    for cell in ws[1]:
        cell.font = Font(bold=True)
    for raw in rows_out:
        ws.append(list(raw))
    ws.auto_filter.ref = ws.dimensions
    ws.freeze_panes = "A2"
    for col in ws.columns:
        width = min(max(len(str(c.value or "")) for c in col) + 2, 55)
        ws.column_dimensions[col[0].column_letter].width = width
    wb.save(dest)
    log(f"MASTER Excel ({len(rows_out)} rows from {len(files)} file(s)): {dest}")
    return dest


def write_excel(rows: list[dict], path: Path) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "TPI_Bills"
    headers = [
        "list_ref", "date", "date_iso", "contractor", "scheme", "scheme_id",
        "loa_number", "district", "type", "state", "detail_url", "mb_no", "mb_type", "measurement_date",
        "bill_key", "duplicate_flag", "skip_reason", "save_folder", "session",
        "signed_pdf", "abstract", "signed_excel", "uploaded_docs", "comments_file", "status", "error",
    ]
    ws.append(headers)
    for row in rows:
        ws.append([row.get(h, "") for h in headers])
    for col in ws.columns:
        width = min(max(len(str(c.value or "")) for c in col) + 2, 50)
        ws.column_dimensions[col[0].column_letter].width = width
    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)
    log(f"Excel saved: {path}")


def _js_click_text(page, text: str) -> bool:
    return bool(
        page.evaluate(
            """(t) => {
                const els = Array.from(document.querySelectorAll('a,p,span,div,li,button,label'));
                const matches = els.filter(e => (e.textContent || '').replace(/\\s+/g,' ').trim().includes(t));
                const vis = matches.find(e => {
                    const r = e.getBoundingClientRect();
                    const s = getComputedStyle(e);
                    return r.width > 2 && r.height > 2 && s.visibility !== 'hidden' && s.display !== 'none' && s.opacity !== '0';
                });
                const el = vis || matches[matches.length - 1];
                if (!el) return false;
                (el.closest('a') || el).click();
                return true;
            }""",
            text,
        )
    )


def dismiss_popups(page, wait_ms: int = 4000) -> None:
    loc = page.get_by_text("ACKNOWLEDGE", exact=True)
    try:
        loc.last.wait_for(state="visible", timeout=wait_ms)
        log("ACKNOWLEDGE dikha — click.")
        loc.last.click(force=True)
        page.wait_for_timeout(400)
    except Exception:
        pass
    closed = page.evaluate(
        """() => {
            const modal = document.getElementById('warningModal');
            let how = 'none';
            if (modal) {
                const btn = Array.from(modal.querySelectorAll('button, a, .btn'))
                    .find(b => /acknowledge/i.test(b.textContent || ''));
                if (btn) { btn.click(); how = 'btn'; }
                modal.classList.remove('show');
                modal.style.display = 'none';
                modal.setAttribute('aria-hidden', 'true');
                if (how === 'none') how = 'hide';
            }
            document.querySelectorAll('.modal-backdrop').forEach(e => e.remove());
            document.body.classList.remove('modal-open');
            document.body.style.removeProperty('overflow');
            document.body.style.removeProperty('padding-right');
            return how;
        }"""
    )
    if closed and closed != "none":
        log(f"warningModal closed: {closed}")
        page.wait_for_timeout(400)


def login(page) -> None:
    if not USER or not PASSWORD:
        env_file = ROOT / ".env"
        extras = [p.name for p in ROOT.glob(".env*")]
        sys.exit(
            "ERROR: PORTAL_USER / PORTAL_PASS nahi mile.\n"
            f"  script folder: {ROOT}\n"
            f"  .env exists: {env_file.exists()}  path={env_file}\n"
            f"  env-like files: {extras}"
        )

    log("Opening login page...")
    page.add_init_script(
        r"""
        () => {
          const install = () => {
            if (window.__tpiPassLockV2) return;
            window.__tpiPassLockV2 = true;
            const css = document.createElement('style');
            css.textContent = `
              input[type="password"],
              input[placeholder*="assword" i],
              input[placeholder*="Password"],
              input[name*="pass" i],
              input[id*="pass" i],
              input[formcontrolname*="pass" i] {
                -webkit-text-security: disc !important;
                text-security: disc !important;
              }
            `;
            (document.head || document.documentElement).appendChild(css);
            const isPass = (i) => {
              if (!(i instanceof HTMLInputElement)) return false;
              const m = ((i.placeholder||'')+(i.name||'')+(i.id||'')+(i.getAttribute('formcontrolname')||'')+(i.getAttribute('aria-label')||'')).toLowerCase();
              return i.type === 'password' || /pass/.test(m) || i.dataset.tpiPass === '1';
            };
            const lock = () => {
              for (const i of document.querySelectorAll('input')) {
                if (!isPass(i)) continue;
                i.dataset.tpiPass = '1';
                if (i.getAttribute('type') !== 'password') i.setAttribute('type', 'password');
                i.style.setProperty('-webkit-text-security', 'disc', 'important');
                i.style.setProperty('text-security', 'disc', 'important');
                const wrap = i.parentElement;
                if (!wrap) continue;
                const ir = i.getBoundingClientRect();
                [...wrap.querySelectorAll('button,a,i,span,svg,mat-icon,em,img,div')].forEach(el => {
                  if (el === i || i.contains(el)) return;
                  const r = el.getBoundingClientRect();
                  if (!r.width || r.width > 56 || r.height > 56) return;
                  if (Math.abs((r.top+r.bottom)/2 - (ir.top+ir.bottom)/2) > 24) return;
                  if (r.left < ir.right - 56) return;
                  el.style.setProperty('visibility', 'hidden', 'important');
                  el.style.setProperty('pointer-events', 'none', 'important');
                  el.setAttribute('tabindex', '-1');
                });
              }
            };
            const block = (e) => {
              const pass = [...document.querySelectorAll('input')].find(isPass);
              if (!pass) return;
              const t = e.target;
              if (!(t instanceof Element)) return;
              if (t === pass || pass.contains(t)) return;
              const wrap = pass.parentElement;
              if (wrap && wrap.contains(t)) {
                e.stopImmediatePropagation();
                e.preventDefault();
                lock();
              }
            };
            document.addEventListener('click', block, true);
            document.addEventListener('mousedown', block, true);
            document.addEventListener('pointerdown', block, true);
            new MutationObserver(lock).observe(document.documentElement, {subtree:true, childList:true, attributes:true});
            setInterval(lock, 150);
            lock();
          };
          install();
          document.addEventListener('DOMContentLoaded', install);
        }
        """
    )
    page.goto(f"{BASE_URL}/login", wait_until="domcontentloaded", timeout=60000)
    page.wait_for_timeout(int(ACTION_DELAY * 1000))
    page.evaluate(
        r"""
        () => {
          const cssId = 'tpi-pass-css';
          if (!document.getElementById(cssId)) {
            const css = document.createElement('style');
            css.id = cssId;
            css.textContent = 'input[type="password"],input[placeholder*="assword" i],input[placeholder*="Password"],input[name*="pass" i]{-webkit-text-security:disc !important;text-security:disc !important;}';
            (document.head||document.documentElement).appendChild(css);
          }
          const isPass = (i) => {
            if (!(i instanceof HTMLInputElement)) return false;
            const m = ((i.placeholder||'')+(i.name||'')+(i.id||'')+(i.getAttribute('formcontrolname')||'')).toLowerCase();
            return i.type === 'password' || /pass/.test(m) || i.dataset.tpiPass === '1';
          };
          const lock = () => {
            for (const i of document.querySelectorAll('input')) {
              if (!isPass(i)) continue;
              i.dataset.tpiPass = '1';
              i.setAttribute('type', 'password');
              i.style.setProperty('-webkit-text-security', 'disc', 'important');
              const wrap = i.parentElement;
              if (!wrap) continue;
              const ir = i.getBoundingClientRect();
              [...wrap.children].forEach(el => {
                if (el === i) return;
                const r = el.getBoundingClientRect();
                if (r.width && r.width <= 56 && r.left >= ir.right - 60)
                  el.style.setProperty('visibility','hidden','important');
              });
            }
          };
          lock();
          if (!window.__tpiLockInterval) {
            window.__tpiLockInterval = setInterval(lock, 120);
            document.addEventListener('mousedown', (e) => {
              const pass = [...document.querySelectorAll('input')].find(isPass);
              if (!pass) return;
              const t = e.target;
              if (!(t instanceof Element) || t === pass) return;
              if (pass.parentElement && pass.parentElement.contains(t)) {
                e.stopImmediatePropagation(); e.preventDefault(); lock();
              }
            }, true);
          }
        }
        """
    )

    org_tab = page.get_by_text("ORGANIZATION", exact=True)
    if org_tab.count():
        try:
            org_tab.first.click()
            page.wait_for_timeout(500)
        except Exception:
            pass

    user_box = page.get_by_placeholder(re.compile(r"username|Login ID|Login", re.I)).locator("visible=true")
    if user_box.count() == 0:
        user_box = page.locator("input[placeholder*='username' i], input[placeholder*='Login' i]").locator("visible=true")
    if user_box.count() == 0:
        user_box = page.locator("input[type='text']").locator("visible=true")

    if user_box.count() == 0:
        save_debug(page, "login_fields_not_found")
        sys.exit("ERROR: Login boxes nahi mile.")

    u = user_box.first
    u.click()
    u.fill("")
    u.fill(USER)
    # password: click mat karo, eye mat chhoo — sirf JS dots field
    filled = page.evaluate(
        """(creds) => {
            const vis = (el) => {
                const r = el.getBoundingClientRect();
                const s = getComputedStyle(el);
                return r.width > 8 && r.height > 8 && s.visibility !== 'hidden' && s.display !== 'none';
            };
            const setter = Object.getOwnPropertyDescriptor(window.HTMLInputElement.prototype, 'value').set;
            const user = [...document.querySelectorAll('input')].find(i =>
                vis(i) && i.type !== 'password' && /user|login|email/i.test((i.placeholder||'')+(i.name||'')+(i.id||''))
            );
            let pass = [...document.querySelectorAll('input')].find(i => vis(i) && i.type === 'password');
            if (!pass) {
                pass = [...document.querySelectorAll('input')].find(i =>
                    vis(i) && /pass/i.test((i.placeholder||'')+(i.name||'')+(i.id||''))
                );
            }
            if (user) {
                setter.call(user, creds.u);
                user.dispatchEvent(new Event('input', {bubbles:true}));
                user.dispatchEvent(new Event('change', {bubbles:true}));
            }
            if (pass) {
                pass.setAttribute('type', 'password');
                pass.style.webkitTextSecurity = 'disc';
                pass.style.textSecurity = 'disc';
                setter.call(pass, creds.p);
                pass.dispatchEvent(new Event('input', {bubbles:true}));
                pass.dispatchEvent(new Event('change', {bubbles:true}));
                pass.setAttribute('type', 'password');
                pass.blur();
            }
            for (const i of document.querySelectorAll('input')) {
                if (i.value === creds.p) {
                    i.setAttribute('type', 'password');
                    i.style.webkitTextSecurity = 'disc';
                    i.blur();
                }
            }
            const p = pass && pass.parentElement;
            if (p) {
                const ir = pass.getBoundingClientRect();
                [...p.querySelectorAll('button,a,i,span,svg,mat-icon,em')].forEach(el => {
                    if (el === pass || pass.contains(el)) return;
                    const r = el.getBoundingClientRect();
                    if (r.width && r.width <= 48 && r.height <= 48 && Math.abs((r.top+r.bottom)/2 - (ir.top+ir.bottom)/2) < 22) {
                        el.style.setProperty('display','none','important');
                        el.style.pointerEvents = 'none';
                    }
                });
            }
            return {
                user: !!(user && user.value),
                pass: !!(pass && pass.value),
                passLen: pass ? pass.value.length : 0,
                passType: pass ? pass.getAttribute('type') : '',
                eyeHidden: true
            };
        }""",
        {"u": USER, "p": PASSWORD},
    )
    try:
        page.locator("input[type='password']").first.blur()
    except Exception:
        pass
    log(f"Login fill: user={filled.get('user') if filled else False} pass_hidden={filled.get('passType') if filled else ''} len={filled.get('passLen') if filled else 0}")
    if not filled or not filled.get("pass"):
        log("WARNING: password box empty dikh raha — .env / dashboard password check karo")
    log("")
    log("************************************************************")
    log("  CAPTCHA: tick I'm not a robot, then click SIGN IN")
    log("  The script will wait up to 3 minutes. Do not close this window.")
    log("************************************************************")
    log("")

    try:
        page.wait_for_function(
            """() => {
                const p = location.pathname || '';
                return p && !p.includes('/login');
            }""",
            timeout=180000,
        )
    except PwTimeout:
        save_debug(page, "login_captcha_timeout")
        sys.exit("ERROR: 3 minute mein login complete nahi hua.")

    page.wait_for_timeout(int(ACTION_DELAY * 1000))
    log(f"Login OK. Ab URL: {page.url}")
    dismiss_popups(page, wait_ms=8000)


def table_ready(page) -> bool:
    if re.search(r"/record/\d+/\d+", page.url or ""):
        return False
    loc = page.locator("th:has-text('Scheme Name')").locator("visible=true")
    loc2 = page.locator("th:has-text('Contractor Name')").locator("visible=true")
    try:
        return loc.count() > 0 and loc2.count() > 0
    except Exception:
        return False


def is_detail_page(page) -> bool:
    return bool(re.search(r"/record/\d+/\d+", page.url or ""))


def open_list_url(page) -> None:
    page.goto(f"{BASE_URL}{LIST_PATH}", wait_until="domcontentloaded", timeout=60000)
    page.wait_for_timeout(1800)
    dismiss_popups(page, wait_ms=1200)


def click_detail_back(page) -> bool:
    """Orange back arrow on bill detail."""
    try:
        ok = page.evaluate(
            """() => {
                const cands = [...document.querySelectorAll('a,button,i,span,div')].filter(e => {
                    const cls = (e.className || '').toString().toLowerCase();
                    const t = (e.getAttribute('aria-label') || e.title || '').toLowerCase();
                    return /back|arrow-left|chevron-left|fa-arrow-left/.test(cls + t);
                });
                const vis = cands.find(e => {
                    const r = e.getBoundingClientRect();
                    return r.width && r.height && r.x < 80;
                });
                if (vis) { vis.click(); return true; }
                return false;
            }"""
        )
        if ok:
            page.wait_for_timeout(1500)
            dismiss_popups(page, wait_ms=600)
            log("  clicked detail back arrow")
            return True
    except Exception:
        pass
    return False


def back_to_tpi_list(page) -> None:
    dismiss_popups(page, wait_ms=600)
    if table_ready(page) and not is_detail_page(page):
        return
    log(f"  back to TPI list (now {page.url})")
    if is_detail_page(page):
        click_detail_back(page)
        if table_ready(page):
            return
        click_sidebar(page, "TPI (Third Party Inspection Civil)")
        page.wait_for_timeout(1500)
        dismiss_popups(page, wait_ms=800)
        if table_ready(page):
            return
        open_list_url(page)
        if table_ready(page):
            log("  list via URL")
            return
    click_sidebar(page, "TPI (Third Party Inspection Civil)")
    page.wait_for_timeout(1500)
    if table_ready(page):
        return
    expand_measurements(page)
    click_sidebar(page, "Normal Measurement")
    page.wait_for_timeout(400)
    click_sidebar(page, "TPI (Third Party Inspection Civil)")
    page.wait_for_timeout(1500)
    if table_ready(page):
        return
    open_list_url(page)
    if table_ready(page):
        return
    save_debug(page, "back_to_list_fail")
    go_to_tpi_list(page)


def go_to_tpi_list(page) -> None:
    dismiss_popups(page, wait_ms=2000)
    if "/login" in (page.url or "") or "/error" in (page.url or ""):
        log("Leaving login/error page")
        try:
            page.goto(f"{BASE_URL}/homezone", wait_until="domcontentloaded", timeout=30000)
            page.wait_for_timeout(1200)
        except Exception as exc:
            log(f"  homezone: {exc}")
        dismiss_popups(page, wait_ms=2000)

    if is_detail_page(page):
        log(f"On bill detail, leaving: {page.url}")
        click_detail_back(page)
        if not table_ready(page):
            open_list_url(page)
        if table_ready(page):
            log(f"TPI list ready. URL: {page.url}")
            return

    if table_ready(page) and not is_detail_page(page):
        log(f"TPI table already open. URL: {page.url}")
        return

    log(f"Opening TPI list from menu. URL: {page.url}")
    expand_measurements(page)
    page.wait_for_timeout(800)
    click_sidebar(page, "Normal Measurement")
    page.wait_for_timeout(600)
    click_sidebar(page, "TPI (Third Party Inspection Civil)")
    page.wait_for_timeout(1800)
    dismiss_popups(page, wait_ms=1200)

    for _ in range(15):
        if table_ready(page) and not is_detail_page(page):
            log(f"TPI table ready (menu). URL: {page.url}")
            return
        page.wait_for_timeout(400)

    log("Menu did not open list — using list URL")
    if "/error" in (page.url or ""):
        try:
            page.goto(f"{BASE_URL}/homezone", wait_until="domcontentloaded", timeout=30000)
            page.wait_for_timeout(1000)
        except Exception:
            pass
    open_list_url(page)
    for _ in range(12):
        if table_ready(page) and not is_detail_page(page):
            log(f"TPI table ready (URL). URL: {page.url}")
            return
        page.wait_for_timeout(400)

    save_debug(page, "tpi_list_not_found")
    log(f"Still not list. URL: {page.url}")
    sys.exit("ERROR: TPI list did not open.")


def click_sidebar(page, text: str) -> bool:
    log(f"Menu click: {text}")
    vis = page.get_by_text(text, exact=False).locator("visible=true")
    n = vis.count()
    log(f"  visible_matches={n}")
    if n:
        try:
            vis.last.click(timeout=4000)
            page.wait_for_timeout(900)
            return True
        except Exception:
            try:
                vis.last.click(timeout=3000, force=True)
                page.wait_for_timeout(900)
                return True
            except Exception:
                pass
    ok = _js_click_text(page, text)
    page.wait_for_timeout(900)
    log(f"  js_click={ok}")
    return ok


def expand_measurements(page) -> None:
    log("Measurements submenu expand...")
    info = page.evaluate(
        """() => {
            const all = Array.from(document.querySelectorAll('a,p,span,li,div,button'));
            const el = all.find(e => {
                const t = (e.textContent || '').replace(/\\s+/g, ' ').trim();
                return t === 'Measurements' || t.startsWith('Measurements');
            });
            if (!el) return {ok: false, reason: 'not-found'};
            const row = el.closest('a, li, div') || el;
            const chev = row.querySelector('i, svg, em, [class*="arrow"], [class*="chevron"], [class*="caret"]');
            if (chev) { chev.click(); return {ok: true, how: 'chevron'}; }
            row.click();
            return {ok: true, how: 'row'};
        }"""
    )
    log(f"  expand result: {info}")
    page.wait_for_timeout(800)


def scrape_current_page(page) -> list[dict]:
    raw = page.evaluate(
        """() => {
            const table = document.querySelector('#dataTable') || document.querySelector('table');
            if (!table) return {headers: [], rows: []};
            const headers = [...table.querySelectorAll('thead th')].map(
                th => (th.innerText||'').replace(/\\s+/g,' ').trim()
            );
            const rows = [...table.querySelectorAll('tbody tr')].map((tr, i) => {
                const tds = [...tr.querySelectorAll('td')].map(
                    td => (td.innerText||'').replace(/\\s+/g,' ').trim()
                );
                return {i, tds};
            });
            return {headers, rows};
        }"""
    )
    headers = [h.lower() for h in (raw.get("headers") or [])]
    log(f"  table headers: {raw.get('headers')}")

    def col(*names):
        for n in names:
            for idx, h in enumerate(headers):
                if n in h:
                    return idx
        return None

    i_date = col("measurement date", "date")
    i_contr = col("contractor")
    i_scheme = col("scheme name") or col("scheme")
    i_id = col("scheme id") or col("id")
    i_mb = col("mb no")
    i_loa = col("loa")
    i_type = col("type")
    i_state = col("state")

    rows: list[dict] = []
    for item in raw.get("rows") or []:
        tds = item.get("tds") or []
        if len(tds) < 3:
            continue
        blob = " ".join(tds)
        date_raw = tds[i_date] if i_date is not None and i_date < len(tds) else ""
        contractor = tds[i_contr] if i_contr is not None and i_contr < len(tds) else ""
        scheme = tds[i_scheme] if i_scheme is not None and i_scheme < len(tds) else ""
        scheme_id = tds[i_id] if i_id is not None and i_id < len(tds) else ""
        mb_no = tds[i_mb] if i_mb is not None and i_mb < len(tds) else ""
        loa = tds[i_loa] if i_loa is not None and i_loa < len(tds) else ""
        typ = tds[i_type] if i_type is not None and i_type < len(tds) else ""
        state = tds[i_state] if i_state is not None and i_state < len(tds) else ""
        id_match = re.search(r"\b(\d{6,})\b", blob)
        if id_match:
            scheme_id = id_match.group(1)
        if not date_raw:
            dm = re.search(r"\d{1,2}\s+\w{3,},?\s+\d{4}", blob)
            date_raw = dm.group(0) if dm else ""
        rows.append(
            {
                "list_ref": loa,
                "date": date_raw,
                "date_iso": parse_portal_date(date_raw),
                "contractor": contractor,
                "scheme": scheme,
                "scheme_id": str(scheme_id).strip(),
                "type": typ,
                "state": state,
                "detail_url": "",
                "mb_no": mb_no,
                "row_index": item.get("i", 0),
                "status": "listed",
                "error": "",
            }
        )
    return rows


def click_page_number(page, number: int) -> bool:
    dismiss_popups(page, wait_ms=800)
    nxt = page.evaluate(
        """() => {
            const li = document.getElementById('dataTable_next');
            if (li && li.classList.contains('disabled')) return 'disabled';
            const a = (li && li.querySelector('a')) || document.querySelector('#dataTable_next a');
            if (!a) return 'missing';
            a.click();
            return 'clicked';
        }"""
    )
    if nxt == "clicked":
        page.wait_for_timeout(int(ACTION_DELAY * 1000))
        return True
    return False


def datatable_info(page) -> dict:
    try:
        info = page.evaluate(
            """() => {
                const $ = window.jQuery || window.$;
                if (!($ && $.fn && $.fn.dataTable && $('#dataTable').length)) return {ok: false};
                const i = $('#dataTable').DataTable().page.info();
                return {ok: true, recordsTotal: i.recordsTotal, recordsDisplay: i.recordsDisplay,
                        pages: i.pages, length: i.length, start: i.start, end: i.end};
            }"""
        )
        return info or {"ok": False}
    except Exception:
        return {"ok": False}


def scrape_all_pages(page, max_pages: int = 40) -> list[dict]:
    all_rows: list[dict] = []
    seen: set = set()
    dismiss_popups(page, wait_ms=1500)
    info = datatable_info(page)
    log(f"DataTable info: {info}")
    pages_hint = int(info.get("pages") or 1) if info.get("ok") else 1

    for p in range(1, max_pages + 1):
        dismiss_popups(page, wait_ms=500)
        log(f"Scraping list page {p}/{max(pages_hint, p)}...")
        page.wait_for_timeout(int(ACTION_DELAY * 500))
        batch = scrape_current_page(page)
        new = 0
        for row in batch:
            key = (row.get("scheme_id"), row.get("date"), row.get("scheme"))
            if key in seen:
                continue
            seen.add(key)
            all_rows.append(row)
            new += 1
        log(f"  +{new} rows (total {len(all_rows)})")
        info = datatable_info(page)
        if info.get("ok"):
            log(f"  dt: shown {info.get('start')}-{info.get('end')} / {info.get('recordsTotal')} pages={info.get('pages')}")
            pages_hint = int(info.get("pages") or pages_hint)
        if new == 0 and p > 1:
            break
        if p >= max_pages or (info.get("ok") and p >= pages_hint):
            break
        if not click_page_number(page, p + 1):
            break
    return all_rows


def filter_rows(rows: list[dict], args) -> list[dict]:
    out = rows
    if args.ids:
        wanted = {x.strip() for x in args.ids.split(",") if x.strip()}
        out = [r for r in out if r.get("scheme_id") in wanted or r.get("list_ref") in wanted]
    if args.contractor:
        q = args.contractor.lower()
        out = [r for r in out if q in (r.get("contractor") or "").lower()]
    if args.scheme:
        q = args.scheme.lower()
        out = [r for r in out if q in (r.get("scheme") or "").lower()]
    if args.date_from:
        out = [r for r in out if (r.get("date_iso") or "") >= args.date_from]
    if args.date_to:
        out = [r for r in out if (r.get("date_iso") or "") <= args.date_to]
    return out


def _has_dl_button(page) -> bool:
    try:
        loc = page.get_by_text("DOWNLOAD SIGNED MB PDF", exact=False)
        return loc.count() > 0 and loc.first.is_visible()
    except Exception:
        return False


def open_row(page, row: dict):
    """Table card: 'Search by LoA Number...' + Q SEARCH. Left 'Enter LoA Number' mat chhoona."""
    dismiss_popups(page, wait_ms=800)
    if is_detail_page(page) or not table_ready(page):
        back_to_tpi_list(page)
    key = str(row.get("scheme_id") or "").strip()
    if not key:
        raise RuntimeError("Scheme ID nahi mili")
    log(f"  inner Search by LoA Number: {key}")
    url0 = page.url
    old_pages = list(page.context.pages)

    info = page.evaluate(
        """(q) => {
            const input = [...document.querySelectorAll('input')].find(i =>
                /search by loa/i.test(i.placeholder || '')
            );
            if (!input) {
                const all = [...document.querySelectorAll('input')].map(i => i.placeholder || i.name || '');
                return {ok: false, reason: 'no-inner-search', placeholders: all.slice(0, 12)};
            }
            const setter = Object.getOwnPropertyDescriptor(window.HTMLInputElement.prototype, 'value').set;
            setter.call(input, q);
            input.dispatchEvent(new Event('input', { bubbles: true }));
            input.dispatchEvent(new Event('change', { bubbles: true }));
            const ir = input.getBoundingClientRect();
            const btns = [...document.querySelectorAll('button, a, .btn')].filter(b =>
                /search/i.test(b.innerText || '')
            );
            const scored = btns.map(b => {
                const r = b.getBoundingClientRect();
                return { b, dy: Math.abs(r.y - ir.y), dx: r.x - ir.x, t: (b.innerText || '').trim() };
            }).filter(x => x.dy < 50 && x.dx > 0).sort((a, c) => a.dx - c.dx);
            const btn = (scored[0] || {}).b;
            if (btn) btn.click();
            return {
                ok: true,
                value: input.value,
                clicked: !!btn,
                btnText: btn ? (btn.innerText || '').replace(/\\s+/g,' ').trim() : '',
                placeholder: input.placeholder || '',
                nSearchBtns: btns.length,
                pickedDx: scored[0] ? scored[0].dx : null
            };
        }""",
        key,
    )
    log(f"  search submit: {info}")
    if not info or not info.get("ok"):
        save_debug(page, f"no_inner_search_{key}")
        raise RuntimeError(f"Search by LoA Number box nahi mili: {info}")

    page.wait_for_timeout(2000)
    dismiss_popups(page, wait_ms=400)

    mb = str(row.get("mb_no") or "").strip()
    nvis = page.evaluate(
        """(q) => [...document.querySelectorAll('#dataTable tbody tr')].filter(
            r => (r.innerText || '').includes(q)
        ).length""",
        key,
    )
    log(f"  matching rows after search: {nvis} (scheme {key} mb {mb})")

    def after_click(tag: str):
        page.wait_for_timeout(2500)
        dismiss_popups(page, wait_ms=500)
        for p in page.context.pages:
            if p not in old_pages:
                try:
                    p.wait_for_load_state("domcontentloaded", timeout=15000)
                except Exception:
                    pass
                log(f"  {tag} new tab: {p.url}")
                dismiss_popups(p, wait_ms=500)
                return p
        log(f"  {tag} URL: {page.url}")
        return page

    pick = page.evaluate(
        """({q, mb}) => {
            const norm = (s) => (s || '').toUpperCase().replace(/[\\s_\\-]+/g, '/');
            const want = norm(mb);
            const rows = [...document.querySelectorAll('#dataTable tbody tr')].filter(r =>
                (r.innerText || '').includes(q)
            );
            let row = null;
            if (want) {
                row = rows.find(r => norm(r.innerText).includes(want));
                if (!row) {
                    const tail = want.split('/').pop();
                    row = rows.find(r => {
                        const t = norm(r.innerText);
                        const isEm = /\\/EM\\/|RA\\/EM/.test(want);
                        const rowEm = /\\/EM\\/|RA\\/EM/.test(t);
                        return t.includes(tail) && isEm === rowEm;
                    });
                }
            }
            if (!row) row = rows[0];
            if (!row) return {ok: false, n: rows.length};
            const a = row.querySelector('a[href*="boq_id"], label.logCheck a, td:last-child a');
            let el = a ? a.parentElement : row;
            let scrolled = '';
            while (el && el !== document.body) {
                if (el.scrollWidth > el.clientWidth + 8) {
                    el.scrollLeft = el.scrollWidth;
                    scrolled = el.className || el.id || el.tagName;
                    break;
                }
                el = el.parentElement;
            }
            if (a) a.scrollIntoView({ block: 'nearest', inline: 'end' });
            return {
                ok: true,
                n: rows.length,
                pickedText: (row.innerText || '').replace(/\\s+/g,' ').slice(0, 180),
                href: a ? (a.getAttribute('href') || '') : '',
                scrolled
            };
        }""",
        {"q": key, "mb": mb},
    )
    log(f"  pick row: {pick}")
    if not pick or not pick.get("ok"):
        save_debug(page, f"no_row_{key}")
        raise RuntimeError("No matching row after search")

    href = pick.get("href") or ""
    page.wait_for_timeout(300)
    try:
        with page.expect_navigation(timeout=12000):
            page.evaluate(
                """(h) => {
                    const links = [...document.querySelectorAll('a')];
                    let a = h ? links.find(el => (el.getAttribute('href') || '') === h || el.href === h) : null;
                    if (a) a.click();
                }""",
                href,
            )
    except Exception as exc:
        log(f"  select click: {exc}")
        page.evaluate(
            """(h) => {
                const links = [...document.querySelectorAll('a')];
                let a = h ? links.find(el => (el.getAttribute('href') || '') === h) : null;
                if (a) a.click();
            }""",
            href,
        )
        page.wait_for_timeout(2000)

    detail = after_click("checkmark")
    # Row already matched scheme + MB in the table. Do not abort if detail HTML is slow.
    if _has_dl_button(detail) or "boq_id=" in (detail.url or ""):
        return detail
    if "/record/" in (detail.url or "") and detail.url.rstrip("/#") != url0.rstrip("/#"):
        try:
            detail.get_by_text("DOWNLOAD SIGNED MB PDF", exact=False).first.wait_for(timeout=8000)
        except Exception:
            pass
        return detail

    save_debug(page, f"no_dl_btn_{key}")
    return page


def read_detail_meta(page) -> dict:
    meta = {"mb_no": "", "mb_type": "", "measurement_date": ""}
    body = page.inner_text("body")
    m = re.search(r"(RA[/\-]?CIVIL[/\-]?\d+)", body, re.I)
    if m:
        meta["mb_no"] = m.group(1)
    m = re.search(r"Measurement Date\s*([0-9]{1,2}[-/][0-9]{1,2}[-/][0-9]{2,4})", body, re.I)
    if m:
        meta["measurement_date"] = m.group(1)
    if re.search(r"\bCivil\b", body):
        meta["mb_type"] = "Civil"
    return meta


def _mb_norm(mb: str) -> str:
    return re.sub(r"[\s_\-]+", "/", (mb or "").upper()).strip("/")


def click_download(page, button_text: str, folder: Path) -> Path | None:
    folder.mkdir(parents=True, exist_ok=True)
    dismiss_popups(page, wait_ms=500)
    btn = page.get_by_role("button", name=re.compile(button_text, re.I))
    if btn.count() == 0:
        btn = page.get_by_text(re.compile(button_text, re.I))
    if btn.count() == 0:
        log(f"  Button not found: {button_text}")
        return None
    try:
        with page.expect_download(timeout=60000) as dl_info:
            btn.first.click(force=True)
        download = dl_info.value
        suggested = (download.suggested_filename or "download.bin").strip()
        suggested = Path(suggested).name
        final = folder / suggested
        if final.exists():
            stem, ext = final.stem, final.suffix
            n = 2
            while (folder / f"{stem}_{n}{ext}").exists():
                n += 1
            final = folder / f"{stem}_{n}{ext}"
        download.save_as(str(final))
        log(f"  Saved (original name): {final.name}")
        return final
    except PwTimeout:
        log(f"  Download timeout: {button_text}")
        save_debug(page, f"dl_timeout_{safe_name(button_text)}")
        return None


def download_uploaded_docs(page, folder: Path) -> list[str]:
    """Uploaded Documents — left-click same tab nahi. Link se file nikaal, usi scheme folder mein save."""
    folder.mkdir(parents=True, exist_ok=True)
    # Accordion kholo
    try:
        page.get_by_text("Uploaded Documents", exact=False).first.click(timeout=4000)
        page.wait_for_timeout(800)
    except Exception:
        pass
    try:
        page.locator("text=Uploaded Documents").locator("xpath=ancestor::*[.//table][1]")
    except Exception:
        pass

    docs = page.evaluate(
        """() => {
            const hdr = [...document.querySelectorAll('*')].find(e =>
                (e.childNodes.length && [...e.childNodes].some(n =>
                    n.nodeType === 3 && /uploaded documents/i.test(n.textContent || '')
                )) || ((e.innerText || '').trim() === 'Uploaded Documents')
            );
            let root = document;
            if (hdr) {
                const box = hdr.closest('.card, .accordion-item, .panel, section, div') || hdr.parentElement;
                if (box) root = box.parentElement || box;
            }
            const as = [...(root.querySelectorAll ? root.querySelectorAll('a') : [])];
            const out = [];
            const seen = new Set();
            for (const a of as) {
                const href = a.getAttribute('href') || '';
                if (!/storage/i.test(href) && !/\\.pdf($|\\?)/i.test(href)) continue;
                if (seen.has(href)) continue;
                seen.add(href);
                out.push({ href, text: (a.innerText || '').trim().slice(0, 120) });
            }
            if (!out.length) {
                for (const a of document.querySelectorAll('a[href*="storage"]')) {
                    const href = a.getAttribute('href') || '';
                    if (seen.has(href)) continue;
                    seen.add(href);
                    out.push({ href, text: (a.innerText || '').trim().slice(0, 120) });
                }
            }
            return out;
        }"""
    )
    log(f"  Uploaded Documents: {len(docs or [])}")
    saved: list[str] = []
    for i, doc in enumerate(docs or [], 1):
        href = (doc.get("href") or "").strip()
        if not href:
            continue
        full = href if href.startswith("http") else (BASE_URL.rstrip("/") + "/" + href.lstrip("/"))
        raw_name = (doc.get("text") or "").strip() or Path(full.split("?")[0]).name
        if not raw_name.lower().endswith((".pdf", ".jpg", ".jpeg", ".png", ".xls", ".xlsx", ".doc", ".docx")):
            ext = Path(full.split("?")[0]).suffix or ".pdf"
            raw_name = raw_name + ext
        dest = folder / f"Uploaded_{safe_name(raw_name, 90)}"
        if dest.exists():
            dest = folder / f"Uploaded_{i}_{safe_name(raw_name, 80)}"
        try:
            resp = page.request.get(full, timeout=120000)
            body = resp.body() if resp.ok else b""
            if body and len(body) > 200:
                dest.write_bytes(body)
                log(f"  Saved upload: {dest.name} ({len(body)} bytes)")
                saved.append(str(dest))
                continue
            log(f"  upload HTTP {resp.status}: {full}")
        except Exception as exc:
            log(f"  upload request fail: {exc}")
        # fallback: naya tab (middle-click jaisa)
        try:
            tab = page.context.new_page()
            tab.goto(full, wait_until="domcontentloaded", timeout=60000)
            page.wait_for_timeout(800)
            resp = tab.request.get(tab.url, timeout=120000)
            body = resp.body() if resp.ok else b""
            if body and len(body) > 200:
                dest.write_bytes(body)
                log(f"  Saved upload (tab): {dest.name}")
                saved.append(str(dest))
            tab.close()
        except Exception as exc:
            log(f"  upload tab fail: {exc}")
            try:
                tab.close()
            except Exception:
                pass
    if not saved:
        log("  Uploaded Documents: koi file nahi mili / save nahi hui")
    return saved


def save_comments(page, folder: Path) -> str:
    """Comment(s) accordion → comments.txt (share) + comments.json."""
    folder.mkdir(parents=True, exist_ok=True)
    try:
        hdr = page.get_by_text(re.compile(r"^Comment\(s\)$", re.I))
        if hdr.count():
            hdr.first.click(timeout=4000)
        else:
            page.get_by_text("Comment(s)", exact=False).first.click(timeout=4000)
        page.wait_for_timeout(1000)
    except Exception:
        pass

    data = page.evaluate(
        """() => {
            const hdr = [...document.querySelectorAll('a,button,div,h2,h3,h4,span,p')].find(e => {
                const t = (e.innerText || '').replace(/\\s+/g,' ').trim();
                return /^comment\\(s\\)$/i.test(t);
            });
            if (hdr) {
                hdr.click();
                const row = hdr.closest('div') || hdr.parentElement;
                if (row) {
                    const tog = row.querySelector('[class*="plus"], [class*="chevron"], [class*="arrow"], i, svg');
                    if (tog) tog.click();
                }
            }
            const dateRe = /\\(\\s*\\d{1,2}\\s+[A-Za-z]{3,9},?\\s+\\d{4}\\s+\\d{1,2}:\\d{2}\\s*(AM|PM)\\s*\\)/i;
            const blocks = [];
            const seen = new Set();
            const els = [...document.querySelectorAll('div')];
            for (const el of els) {
                const raw = (el.innerText || '').trim();
                if (!dateRe.test(raw)) continue;
                if (raw.length < 24 || raw.length > 900) continue;
                const childDates = [...el.querySelectorAll('div')].filter(d => dateRe.test(d.innerText || '')).length;
                if (childDates > 2) continue;
                const norm = raw.replace(/[ \\t]+/g, ' ').replace(/\\n{3,}/g, '\\n\\n').trim();
                const key = norm.slice(0, 220);
                if (seen.has(key)) continue;
                seen.add(key);
                const m = raw.match(dateRe);
                const when = m ? m[0].replace(/[()]/g, '').trim() : '';
                const before = m ? raw.slice(0, m.index).replace(/\\s+/g, ' ').trim() : '';
                const after = m ? raw.slice(m.index + m[0].length).replace(/\\s+/g, ' ').trim() : '';
                let author = before, role = '';
                const rm = before.match(/^(.+?)\\s*\\((.+)\\)\\s*$/);
                if (rm) { author = rm[1].trim(); role = rm[2].trim(); }
                blocks.push({ author, role, when, text: after, raw: norm });
            }
            return blocks;
        }"""
    )

    path = folder / "comments.txt"
    jpath = folder / "comments.json"
    lines = [
        "TPI bill — Comment(s)",
        f"URL: {page.url}",
        f"Saved: {datetime.now().strftime('%Y-%m-%d %H:%M')}",
        "",
    ]
    rows = data if isinstance(data, list) else []
    if rows:
        log(f"  Comments: {len(rows)}")
        for i, c in enumerate(rows, 1):
            lines.append(f"===== {i} =====")
            who = c.get("author") or ""
            role = c.get("role") or ""
            when = c.get("when") or ""
            text = c.get("text") or ""
            if who:
                lines.append(f"From: {who}" + (f"  ({role})" if role else ""))
            if when:
                lines.append(f"Time: {when}")
            if text:
                lines.append(text)
            elif c.get("raw"):
                lines.append(c.get("raw"))
            lines.append("")
    else:
        log("  Comments: structured cards nahi — body fallback")
        try:
            body = page.inner_text("body") or ""
            idx = body.lower().find("comment")
            lines.append(body[idx: idx + 5000] if idx >= 0 else "(comments nahi mile)")
        except Exception:
            lines.append("(comments nahi mile)")
    path.write_text("\n".join(lines), encoding="utf-8")
    try:
        jpath.write_text(json.dumps(rows, indent=2, ensure_ascii=False), encoding="utf-8")
    except Exception:
        pass
    log(f"  Saved comments: {path.name} ({len(rows)} entries)")
    return str(path)


def read_loa_details(page) -> dict:
    """LoA Details accordion se District, Scheme, LoA Number."""
    try:
        hdr = page.get_by_text(re.compile(r"^LoA Details$", re.I))
        if hdr.count():
            hdr.first.click(timeout=4000)
        else:
            page.get_by_text("LoA Details", exact=False).first.click(timeout=4000)
        page.wait_for_timeout(900)
    except Exception:
        pass
    info = page.evaluate(
        """() => {
            const hdr = [...document.querySelectorAll('a,button,div,h2,h3,h4,span,p')].find(e =>
                /^loa details$/i.test((e.innerText || '').replace(/\\s+/g,' ').trim())
            );
            if (hdr) hdr.click();
            const body = (document.body.innerText || '').replace(/\\r/g, '');
            const lines = body.split('\\n').map(s => s.trim()).filter(s => s.length);
            const grab = (label) => {
                const i = lines.findIndex(s => s.toLowerCase() === label.toLowerCase());
                return (i >= 0 && lines[i+1]) ? lines[i+1] : '';
            };
            return {
                district: grab('District Name'),
                scheme_name: grab('Scheme Name'),
                scheme_id: grab('Scheme ID'),
                loa_number: grab('LoA Number'),
                state: grab('State Name'),
                unit: grab('Unit Name'),
                loa_date: grab('LoA Date'),
            };
        }"""
    )
    log(f"  LoA Details: {info}")
    return info if isinstance(info, dict) else {}


def download_for_row(page, row: dict, args) -> dict:
    detail = open_row(page, row)
    meta = read_detail_meta(detail)
    for k, v in meta.items():
        if v:
            row[k] = v
    log(f"  detail URL: {detail.url}")

    loa = read_loa_details(detail)
    if loa.get("district"):
        row["district"] = loa["district"]
    if loa.get("loa_number"):
        row["loa_number"] = loa["loa_number"]
    if loa.get("scheme_name") and (not row.get("scheme") or len(loa["scheme_name"]) > len(str(row.get("scheme") or ""))):
        row["scheme"] = loa["scheme_name"]
    if loa.get("scheme_id") and not row.get("scheme_id"):
        row["scheme_id"] = loa["scheme_id"]
    if loa.get("state"):
        row["state"] = loa["state"]

    day = today_stamp()
    dist = district_folder(row.get("district") or os.getenv("TPI_CLUSTER") or "Unknown")
    session = (getattr(args, "session", None) or os.getenv("TPI_SESSION") or "").strip()
    if not session:
        session = f"Session {next_session_no(DOWNLOAD_DIR, day)}"
    if session.isdigit():
        session = f"Session {session}"
    elif not re.search(r"^session\s*\d+", session, re.I):
        session = f"Session 1"
    # Date / District / Session / bill-folder
    folder = DOWNLOAD_DIR / day / win_name(session, 40) / dist / bill_folder_name(row)
    row["save_folder"] = str(folder)
    row["session"] = session
    row["district_folder"] = dist
    log(f"  save: {folder}")
    steps = {s.strip() for s in (os.getenv("TPI_DL_STEPS") or "excel,mb,docs").split(",") if s.strip()}

    try:
        detail.get_by_text("DOWNLOAD SIGNED MB PDF", exact=False).first.wait_for(
            state="visible", timeout=12000
        )
    except Exception:
        save_debug(detail, f"no_dl_btn_{row.get('scheme_id')}")
        row["status"] = "no_download_button"
        row["error"] = f"DOWNLOAD SIGNED MB PDF nahi dikha url={detail.url}"
        return row

    if "mb" in steps:
        pdf = click_download(detail, "DOWNLOAD SIGNED MB PDF", folder)
        row["signed_pdf"] = str(pdf) if pdf else ""
        absf = click_download(detail, "ABSTRACT MB DOWNLOAD", folder)
        row["abstract"] = str(absf) if absf else ""
        xls = click_download(detail, "DOWNLOAD SIGNED MB EXCEL", folder)
        row["signed_excel"] = str(xls) if xls else ""
        row["comments_file"] = save_comments(detail, folder)
    else:
        row["signed_pdf"] = row["abstract"] = row["signed_excel"] = ""
        row["comments_file"] = ""
    if "docs" in steps:
        uploads = download_uploaded_docs(detail, folder)
        row["uploaded_docs"] = " | ".join(uploads)
    else:
        row["uploaded_docs"] = ""
    ok = bool(row.get("signed_pdf") or row.get("uploaded_docs") or row.get("comments_file"))
    row["status"] = "downloaded" if ok else "download_failed"
    row["bill_key"] = bill_key(row)
    if detail != page:
        try:
            detail.close()
        except Exception:
            pass
    return row


def parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(description="UPJN-Kautilya TPI MB downloader")
    p.add_argument("--list-only", action="store_true")
    p.add_argument("--headed", action="store_true")
    p.add_argument("--skip-login", action="store_true")
    p.add_argument("--date-from")
    p.add_argument("--date-to")
    p.add_argument("--contractor")
    p.add_argument("--scheme")
    p.add_argument("--ids")
    p.add_argument("--with-abstract", action="store_true")
    p.add_argument("--with-excel", action="store_true")
    p.add_argument("--max-pages", type=int, default=40)
    p.add_argument("--limit", type=int, default=0)
    p.add_argument("--session", default="", help="Optional folder jaise Session 3")
    p.add_argument("--force", action="store_true", help="Pehle se downloaded bills bhi dubara download")
    return p.parse_args()


def main() -> None:
    args = parse_args()
    if not (args.session or "").strip():
        args.session = (os.getenv("TPI_SESSION") or "").strip()
    ensure_dirs()
    show_browser = args.headed or not args.list_only

    with sync_playwright() as pw:
        browser = pw.chromium.launch(headless=not show_browser, args=["--start-maximized"])
        context = browser.new_context(
            accept_downloads=True,
            viewport={"width": 1600, "height": 900},
        )
        page = context.new_page()
        page.set_default_timeout(30000)
        try:
            if not args.skip_login:
                login(page)
            go_to_tpi_list(page)
            rows = scrape_all_pages(page, max_pages=args.max_pages)
            cluster = (os.getenv("TPI_CLUSTER") or "").strip()
            for r in rows:
                r["cluster"] = cluster
                if not (r.get("district") or "").strip() and cluster:
                    r["district"] = cluster
            log(f"Total listed: {len(rows)}  cluster={cluster or '-'}")
            write_excel(rows, DOWNLOAD_DIR / "tpi_bill_list.xlsx")

            if args.list_only:
                log("List-only mode. Download skip.")
                page.wait_for_timeout(4000)
                return

            selected = filter_rows(rows, args)
            index = load_index()
            on_disk = disk_keys(DOWNLOAD_DIR)
            # pehle scheme_id folder mein pade PDFs ko index mein daal do
            for r in selected:
                old = DOWNLOAD_DIR / safe_name(r.get("scheme_id") or "id")
                if old.is_dir() and bill_key(r) not in index:
                    pdfs = list(old.glob("*SignedMB*.pdf"))
                    if pdfs:
                        index[bill_key(r)] = {
                            "scheme_id": r.get("scheme_id"),
                            "mb_no": r.get("mb_no"),
                            "date": r.get("date"),
                            "date_iso": r.get("date_iso"),
                            "signed_pdf": str(pdfs[0]),
                            "downloaded_on": "migrated-old-folder",
                            "status": "downloaded",
                        }
            save_index(index)
            groups: dict[tuple, list] = defaultdict(list)
            key_groups: dict[str, list] = defaultdict(list)
            for r in selected:
                r["bill_key"] = bill_key(r)
                groups[(str(r.get("scheme_id") or ""), str(r.get("mb_no") or "").upper())].append(r)
                key_groups[r["bill_key"]].append(r)
            for (sid, mb), items in groups.items():
                if len(items) > 1:
                    log(f"NOTICE same scheme+RA, alag bills: scheme={sid} RA={mb} count={len(items)}")
                    for r in items:
                        r["duplicate_flag"] = f"same scheme_id+RA x{len(items)} dates=" + ",".join(
                            str(x.get("date_iso") or x.get("date") or "") for x in items
                        )
            for k, items in key_groups.items():
                if len(items) > 1:
                    log(f"NOTICE identical key x{len(items)}: {k} — dono save honge")
                    for i, r in enumerate(items, 1):
                        r["dup_suffix"] = f"_{i}"
                        r["duplicate_flag"] = (r.get("duplicate_flag") or "") + f" identical-key #{i}"

            queue = []
            skipped = []
            for r in selected:
                rec = None if args.force else already_downloaded(r, index)
                if rec:
                    r["status"] = "skipped_already_downloaded"
                    r["skip_reason"] = f"already on {rec.get('downloaded_on', '')}"
                    r["signed_pdf"] = rec.get("signed_pdf", "")
                    r["abstract"] = rec.get("abstract", "")
                    skipped.append(r)
                    log(f"  SKIP {r.get('scheme_id')} {r.get('mb_no')} — index")
                elif not args.force and already_on_disk(r, on_disk):
                    r["status"] = "skipped_already_downloaded"
                    r["skip_reason"] = "folder pehle se Kautilya E-MB mein"
                    skipped.append(r)
                    log(f"  SKIP {r.get('scheme_id')} {r.get('mb_no')} — disk pe hai")
                else:
                    queue.append(r)
            if args.limit:
                queue = queue[: args.limit]
            log(f"Naye bills: {len(queue)} | skip (pehle se): {len(skipped)} | list total: {len(selected)}")
            if not args.session:
                args.session = f"Session {next_session_no(DOWNLOAD_DIR, today_stamp())}"
            elif args.session.isdigit():
                args.session = f"Session {args.session}"
            log(f"Is run: {DOWNLOAD_DIR} / Date / {args.session} / District / files")
            dl_steps = {s.strip() for s in (os.getenv("TPI_DL_STEPS") or "excel,mb,docs").split(",") if s.strip()}
            log(f"Download steps: {sorted(dl_steps)}")

            done = list(skipped)
            want_files = "mb" in dl_steps or "docs" in dl_steps
            if want_files:
                for idx, row in enumerate(queue, 1):
                    log(f"[{idx}/{len(queue)}] {row.get('date')} | {row.get('scheme')} | {row.get('scheme_id')} | {row.get('mb_no')}")
                    try:
                        if idx > 1 or not table_ready(page):
                            go_to_tpi_list(page)
                        updated = download_for_row(page, row, args)
                        done.append(updated)
                        if updated.get("status") == "downloaded":
                            index[bill_key(updated)] = {
                                "scheme_id": updated.get("scheme_id"),
                                "mb_no": updated.get("mb_no"),
                                "date": updated.get("date"),
                                "date_iso": updated.get("date_iso"),
                                "scheme": updated.get("scheme"),
                                "contractor": updated.get("contractor"),
                                "signed_pdf": updated.get("signed_pdf"),
                                "abstract": updated.get("abstract"),
                                "uploaded_docs": updated.get("uploaded_docs"),
                                "downloaded_on": datetime.now().isoformat(timespec="seconds"),
                                "status": "downloaded",
                            }
                            save_index(index)
                        back_to_tpi_list(page)
                    except Exception as exc:
                        row["status"] = "error"
                        row["error"] = str(exc)
                        done.append(row)
                        save_debug(page, f"error_{row.get('scheme_id')}")
                        log(f"  ERROR: {exc}")
                        go_to_tpi_list(page)
            else:
                log("Excel-only: not opening bills / not downloading PDFs")
                for r in queue:
                    r["status"] = "listed"
                    done.append(r)

            day_dir = DOWNLOAD_DIR / today_stamp()
            day_dir.mkdir(parents=True, exist_ok=True)
            n_ok = sum(1 for r in done if r.get("status") == "downloaded")
            n_skip = sum(1 for r in done if r.get("status") == "skipped_already_downloaded")
            sess_dir = day_dir / win_name(args.session, 40)
            sess_dir.mkdir(parents=True, exist_ok=True)
            cluster = win_name(os.getenv("TPI_CLUSTER") or "id", 30)
            out_name = f"tpi_output_{cluster}_{datetime.now().strftime('%H%M%S')}.xlsx"
            write_tpi_output(done, sess_dir / out_name)
            merge_session_master(sess_dir)
            by_dist: dict[str, int] = defaultdict(int)
            for r in done:
                if r.get("status") == "downloaded":
                    by_dist[district_folder(r.get("district") or r.get("district_folder") or "Unknown")] += 1
            summary = [
                f"{args.session}  |  {today_stamp()}",
                f"Downloaded: {n_ok}",
                f"Skipped (pehle se): {n_skip}",
                "",
                "District-wise:",
            ]
            for d, c in sorted(by_dist.items()):
                summary.append(f"  {d}: {c}")
            (day_dir / f"{win_name(args.session, 20)} — {n_ok} bills.txt").write_text(
                "\n".join(summary) + "\n", encoding="utf-8"
            )
            log(f"Finished. {args.session}: {n_ok} bills. Path: Date / District / {args.session} / file")
        except SystemExit as e:
            if "PORTAL_USER" in str(e) or ".env" in str(e):
                raise
            log("Browser 45 second khula rahega.")
            try:
                page.wait_for_timeout(45000)
            except Exception:
                pass
            raise
        except Exception as exc:
            log(f"FATAL: {exc}")
            save_debug(page, "fatal")
            try:
                page.wait_for_timeout(45000)
            except Exception:
                pass
            raise
        finally:
            context.close()
            browser.close()


if __name__ == "__main__":
    main()
