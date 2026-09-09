"""
TPI Measurement Book — Upload.

Puts files from each bill's TO_UPLOAD folder onto the open bill
on https://upjnr-kautilya.in (file inputs on the detail page).

Folder layout (same as download):
  Kautilya E-MB / DD.MM.YYYY / Session N / District / {scheme}_{...} / TO_UPLOAD / *

Usage:
  python tpi_uploader.py --headed --session "Session 1" --day 07.09.2026
  python tpi_uploader.py --headed --scan-only
"""
from __future__ import annotations

import argparse
import os
import re
import sys
from datetime import datetime
from pathlib import Path

from tpi_mb_downloader import (
    ACTION_DELAY,
    BASE_URL,
    DOWNLOAD_DIR,
    USER,
    PASSWORD,
    back_to_tpi_list,
    bill_folder_name,
    dismiss_popups,
    go_to_tpi_list,
    log,
    login,
    open_row,
    save_debug,
    scrape_grand_total,
    today_stamp,
    win_name,
)

from playwright.sync_api import sync_playwright


def parse_bill_folder(name: str) -> dict:
    m = re.match(r"^(\d{6,})_", name)
    sid = m.group(1) if m else ""
    ra = re.search(r"RA[-_/]?(CIVIL|EM)[-_/]?(\d+)", name, re.I)
    mb = ""
    if ra:
        kind = ra.group(1).upper()
        num = ra.group(2)
        mb = f"RA/{kind}/{num}" if kind == "CIVIL" else f"RA/EM/{num}"
    return {"scheme_id": sid, "mb_no": mb, "folder_name": name}


UPLOAD_DIR_NAMES = ("For Upload", "TO_UPLOAD")
MAX_BYTES = 7 * 1024 * 1024


def _norm_bill_type(text: str) -> str:
    t = (text or "").strip().lower()
    if t.startswith("ret") or t in ("r", "return bill"):
        return "return"
    return "forward"


def fill_from_excel(bill_dir: Path) -> dict:
    """Amount Approved, TPI Letter Number, Remark from xlsx in bill or For Upload."""
    out = {"amount": "", "letter": "", "remark": "", "bill_type": "forward"}
    paths = []
    for name in UPLOAD_DIR_NAMES:
        d = bill_dir / name
        if d.exists():
            paths.extend(d.glob("*.xlsx"))
            paths.extend(d.glob("*.xls"))
    paths.extend(bill_dir.glob("*.xlsx"))
    paths = [p for p in paths if p.is_file() and not p.name.startswith("~$")
             and "signed" not in p.name.lower() and "tpi_output" not in p.name.lower()]
    if not paths:
        return out
    from openpyxl import load_workbook
    for path in paths:
        try:
            wb = load_workbook(path, data_only=True, read_only=True)
            ws = wb.active
            rows = list(ws.iter_rows(max_row=30, max_col=12, values_only=True))
            wb.close()
        except Exception:
            continue
        if not rows:
            continue
        headers = [str(c or "").strip().lower() for c in rows[0]]

        def col(*keys):
            for i, h in enumerate(headers):
                if any(k in h for k in keys):
                    return i
            return None

        ia = col("amount approved", "amount")
        il = col("letter")
        ir = col("remark", "comment")
        it = col("bill type", "action")
        # vertical key-value sheet
        if ia is None and il is None and it is None:
            for row in rows:
                k = str(row[0] or "").strip().lower()
                v = row[1] if len(row) > 1 else ""
                if "amount" in k:
                    out["amount"] = str(v or "").strip()
                elif "letter" in k:
                    out["letter"] = str(v or "").strip()
                elif "remark" in k or "comment" in k:
                    out["remark"] = str(v or "").strip()
                elif "bill type" in k or k in ("type", "action", "forward", "return"):
                    out["bill_type"] = str(v or "").strip()
            if out["amount"] or out["letter"] or out["remark"] or out["bill_type"]:
                out["bill_type"] = _norm_bill_type(out["bill_type"])
                if out["bill_type"] == "return":
                    out["amount"] = "0.00"
                log(f"  fill excel: {path.name} type={out['bill_type']}")
                return out
            continue
        data = rows[1] if len(rows) > 1 else None
        if not data:
            continue
        if ia is not None:
            out["amount"] = str(data[ia] or "").strip()
        if il is not None:
            out["letter"] = str(data[il] or "").strip()
        if ir is not None:
            out["remark"] = str(data[ir] or "").strip()
        if it is not None:
            out["bill_type"] = str(data[it] or "").strip()
        out["bill_type"] = _norm_bill_type(out["bill_type"])
        if out["bill_type"] == "return":
            out["amount"] = "0.00"
        log(f"  fill excel: {path.name} type={out['bill_type']}")
        return out
    return out


def list_upload_files(bill_dir: Path) -> list[Path]:
    files: list[Path] = []
    for name in UPLOAD_DIR_NAMES:
        d = bill_dir / name
        if not d.exists():
            continue
        for p in sorted(d.iterdir()):
            if not p.is_file() or p.name.startswith("~$") or p.name.startswith("."):
                continue
            if p.suffix.lower() in {".xlsx", ".xls"}:
                continue
            files.append(p)
    return files


def scan_session(root: Path, day: str, session: str) -> list[dict]:
    sess = root / day / session
    if not sess.exists():
        log(f"Session folder not found: {sess}")
        return []
    bills = []
    for dist in sorted(p for p in sess.iterdir() if p.is_dir()):
        if dist.name.lower().startswith("session"):
            continue
        for bill in sorted(p for p in dist.iterdir() if p.is_dir()):
            meta = parse_bill_folder(bill.name)
            files = list_upload_files(bill)
            fill = fill_from_excel(bill)
            bills.append({
                **meta,
                "district": dist.name,
                "path": str(bill),
                "files": [str(f) for f in files],
                "n_files": len(files),
                "fill": fill,
            })
    return bills


def list_days(root: Path) -> list[str]:
    if not root.exists():
        return []
    days = [p.name for p in root.iterdir() if p.is_dir() and re.match(r"\d{2}\.\d{2}\.\d{4}$", p.name)]
    return sorted(days, reverse=True)


def list_sessions(root: Path, day: str) -> list[str]:
    d = root / day
    if not d.exists():
        return []
    sess = [p.name for p in d.iterdir() if p.is_dir() and re.search(r"^session\s+\d+", p.name, re.I)]
    return sorted(sess, key=lambda s: int(re.search(r"\d+", s).group(0)))


def upload_files_on_detail(page, files: list[str]) -> list[str]:
    """Choose File → Description → UPLOAD, one file at a time (max 7 MB)."""
    saved: list[str] = []
    try:
        page.get_by_text("Upload file if any", exact=False).first.scroll_into_view_if_needed(timeout=8000)
    except Exception:
        page.evaluate("window.scrollTo(0, document.body.scrollHeight)")
    page.wait_for_timeout(500)

    file_input = page.locator("input[type='file']")
    if file_input.count() == 0:
        save_debug(page, "no_file_input")
        log("  No Choose File control on this page")
        return saved

    for path in files:
        p = Path(path)
        if not p.exists():
            continue
        if p.stat().st_size > MAX_BYTES:
            log(f"  ERROR: file over 7 MB, skipped: {p.name} ({p.stat().st_size/1024/1024:.1f} MB)")
            continue
        log(f"  uploading {p.name}")
        try:
            file_input.first.set_input_files(str(p))
            page.wait_for_timeout(400)
            desc = page.get_by_placeholder(re.compile("file description", re.I))
            if desc.count() == 0:
                desc = page.locator("input[placeholder*='Description' i], textarea[placeholder*='Description' i]")
            if desc.count():
                desc.first.fill(p.stem[:80])
            btn = page.get_by_role("button", name=re.compile(r"^\s*upload\s*$", re.I))
            if btn.count() == 0:
                btn = page.get_by_text(re.compile(r"^\s*upload\s*$", re.I))
            if btn.count():
                btn.first.click(timeout=8000)
            else:
                page.evaluate(
                    """() => {
                        const b = [...document.querySelectorAll('button,a,.btn')].find(e =>
                            /^\\s*upload\\s*$/i.test((e.innerText||'').replace(/\\s+/g,' ').trim())
                        );
                        if (b) b.click();
                    }"""
                )
            page.wait_for_timeout(800)
            click_modal_ok(page, wait_ms=4000)
            dismiss_popups(page, wait_ms=400)
            saved.append(str(p))
            log(f"  uploaded {p.name}")
        except Exception as exc:
            log(f"  upload fail {p.name}: {exc}")
            save_debug(page, f"upload_fail_{p.stem[:30]}")
    return saved


def _type_into(page, loc, value: str) -> bool:
    """Angular fields ignore paste — type like a keyboard."""
    value = str(value)
    loc.scroll_into_view_if_needed(timeout=8000)
    loc.click(timeout=8000)
    page.wait_for_timeout(200)
    loc.press("Control+A")
    page.wait_for_timeout(80)
    loc.press("Backspace")
    page.wait_for_timeout(80)
    page.keyboard.type(value, delay=70)
    page.wait_for_timeout(150)
    loc.press("Tab")
    page.wait_for_timeout(250)
    try:
        got = loc.input_value()
    except Exception:
        got = ""
    if str(got).strip() != value.strip():
        loc.click()
        loc.press("Control+A")
        loc.press("Backspace")
        page.keyboard.type(value, delay=90)
        loc.press("Tab")
        page.wait_for_timeout(200)
        try:
            got = loc.input_value()
        except Exception:
            got = ""
    log(f"    typed '{value}' now='{got}'")
    return str(got).strip() == value.strip() or (got.replace(",", "") == value.replace(",", ""))


def fill_tpi_fields(page, fill: dict) -> None:
    amount = (fill or {}).get("amount") or ""
    letter = (fill or {}).get("letter") or ""
    remark = (fill or {}).get("remark") or ""
    if not (amount or letter or remark):
        log("  no Amount / Letter / Remark in excel")
        return

    try:
        page.get_by_text("Amount Approved by TPI", exact=False).first.scroll_into_view_if_needed(timeout=8000)
    except Exception:
        pass
    page.wait_for_timeout(400)

    if amount:
        loc = page.locator(
            "xpath=//*[contains(normalize-space(.),'Amount Approved by TPI')][1]/following::input[1]"
        )
        if loc.count() == 0:
            loc = page.get_by_label(re.compile("Amount Approved by TPI", re.I))
        ok = False
        if loc.count():
            ok = _type_into(page, loc.first, amount)
        log(f"  Amount Approved by TPI = {amount} ({ok})")

    if letter:
        loc = page.get_by_placeholder(re.compile("TPI Letter Number", re.I))
        if loc.count() == 0:
            loc = page.locator(
                "xpath=//*[contains(normalize-space(.),'TPI Letter Number')][1]/following::input[1]"
            )
        ok = False
        if loc.count():
            ok = _type_into(page, loc.first, letter)
        log(f"  TPI Letter Number = {letter} ({ok})")

    if remark:
        ta = page.locator("textarea")
        ok = False
        try:
            if ta.count():
                box = ta.last
                box.scroll_into_view_if_needed(timeout=5000)
                box.click()
                box.fill("")
                page.keyboard.type(str(remark), delay=25)
                ok = True
        except Exception as exc:
            log(f"  remark type: {exc}")
        log(f"  Remark filled ({ok})")
    page.wait_for_timeout(500)


def click_modal_ok(page, wait_ms: int = 8000) -> str:
    """OK on Success / OTP sent / Are you sure want to return|forward."""
    page.wait_for_timeout(400)
    for _ in range(max(1, wait_ms // 250)):
        info = page.evaluate(
            """() => {
                const body = (document.body && document.body.innerText) || '';
                const vis = [...document.querySelectorAll('button, .btn, a')].filter(b => {
                    const r = b.getBoundingClientRect();
                    const t = (b.innerText || '').trim();
                    return r.width && r.height && /^OK$/i.test(t);
                });
                const hit = vis.find(b => {
                    const box = b.closest('.modal, .swal2-popup, [class*="modal"], [class*="alert"], [role="dialog"]') || b.parentElement;
                    const t = (box && box.innerText) || body;
                    return /success|data save successfully|please enter|tpi number|otp has been sent|are you sure|want to return|want to forward|press ok/i.test(t);
                }) || vis[0];
                if (!hit) return {ok: false};
                const msg = (hit.closest('.modal, .swal2-popup, [class*="modal"], [role="dialog"]') || hit.parentElement).innerText || '';
                hit.click();
                return {ok: true, msg: msg.replace(/\\s+/g, ' ').trim().slice(0, 160)};
            }"""
        )
        if info and info.get("ok"):
            log(f"  popup OK: {info.get('msg')}")
            page.wait_for_timeout(500)
            return str(info.get("msg") or "ok")
        page.wait_for_timeout(250)
    return ""


def click_save_as_draft(page) -> bool:
    btn = page.get_by_role("button", name=re.compile("save as draft", re.I))
    if btn.count() == 0:
        btn = page.get_by_text(re.compile("save as draft", re.I))
    try:
        if btn.count():
            btn.first.click(timeout=8000)
        else:
            page.evaluate(
                """() => {
                    const b = [...document.querySelectorAll('button,a,.btn')].find(e =>
                        /save as draft/i.test(e.innerText || '')
                    );
                    if (b) b.click();
                }"""
            )
        page.wait_for_timeout(800)
        msg = click_modal_ok(page, wait_ms=10000)
        if re.search(r"Please Enter TPI Number|Enter TPI (Letter )?Number", msg or "", re.I):
            log("  alert: Please Enter TPI Number — letter field empty")
            return False
        if re.search(r"save successfully|Success", msg or "", re.I):
            log("  SAVE AS DRAFT — Data save successfully")
            return True
        # fallback: maybe modal already gone
        try:
            body = page.inner_text("body")[:4000]
        except Exception:
            body = ""
        if re.search(r"Please Enter TPI Number", body, re.I):
            click_modal_ok(page, wait_ms=2000)
            return False
        log("  clicked SAVE AS DRAFT")
        return True
    except Exception as exc:
        log(f"  SAVE AS DRAFT: {exc}")
        save_debug(page, "save_draft_fail")
        return False


def click_attach_signature(page) -> bool:
    """Forward bills only. DSC dongle/PIN is completed by the user if the OS dialog appears."""
    btn = page.get_by_role("button", name=re.compile("attach digital signature", re.I))
    if btn.count() == 0:
        btn = page.get_by_text(re.compile("attach digital signature", re.I))
    try:
        if btn.count():
            btn.first.click(timeout=8000)
        else:
            clicked = page.evaluate(
                """() => {
                    const b = [...document.querySelectorAll('button,a,.btn')].find(e =>
                        /attach digital signature/i.test(e.innerText || '')
                    );
                    if (b) { b.click(); return true; }
                    return false;
                }"""
            )
            if not clicked:
                log("  ATTACH DIGITAL SIGNATURE button not found")
                return False
        log("  clicked ATTACH DIGITAL SIGNATURE")
        # Alert: Are you sure you want to Attach DSC?  →  Yes, please!
        page.wait_for_timeout(600)
        yes = page.get_by_role("button", name=re.compile(r"yes,?\s*please", re.I))
        if yes.count() == 0:
            yes = page.get_by_text(re.compile(r"yes,?\s*please", re.I))
        if yes.count():
            yes.first.click(timeout=8000)
            log("  clicked Yes, please! on DSC confirm")
        else:
            page.evaluate(
                """() => {
                    const b = [...document.querySelectorAll('button,a,.btn')].find(e =>
                        /yes,?\s*please/i.test(e.innerText || '')
                    );
                    if (b) b.click();
                }"""
            )
            log("  DSC confirm Yes, please! (js)")
        page.wait_for_timeout(3000)
        click_modal_ok(page, wait_ms=4000)
        return True
    except Exception as exc:
        log(f"  DSC: {exc}")
        save_debug(page, "dsc_fail")
        return False


def select_action(page, btype: str) -> bool:
    """Portal shows Civil or E&M options itself. We only pick Forward vs Return."""
    want = "return" if btype == "return" else "forward"
    log(f"  Action: pick {want} (Civil / E&M list is automatic)")
    try:
        for i in range(page.locator("select").count()):
            sel = page.locator("select").nth(i)
            opts = [o.strip() for o in sel.locator("option").all_text_contents() if (o or "").strip()]
            hit = next(
                (
                    o
                    for o in opts
                    if (
                        want == "return"
                        and re.search(r"return", o, re.I)
                        and not re.search(r"select one", o, re.I)
                    )
                    or (want == "forward" and re.search(r"forward", o, re.I))
                ),
                None,
            )
            if hit:
                sel.select_option(label=hit)
                page.wait_for_timeout(500)
                log(f"  Action native select → {hit}")
                return True
    except Exception:
        pass
    opened = page.evaluate(
        """() => {
            const norm = t => (t || '').replace(/\\s+/g, ' ').trim();
            const cands = [...document.querySelectorAll(
                'div, span, input, select, [role="combobox"], mat-select, ng-select, .ng-select'
            )].filter(e => {
                const r = e.getBoundingClientRect();
                const t = norm(e.innerText || e.value || e.placeholder || '');
                return r.width > 120 && r.height > 20 && r.height < 90 && /select one/i.test(t);
            });
            cands.sort((a, b) => b.getBoundingClientRect().y - a.getBoundingClientRect().y);
            if (!cands.length) return {ok: false};
            cands[0].click();
            return {ok: true, text: norm(cands[0].innerText || '')};
        }"""
    )
    log(f"  Action dropdown open: {opened}")
    page.wait_for_timeout(450)
    picked = page.evaluate(
        """(want) => {
            const norm = t => (t || '').replace(/\\s+/g, ' ').trim();
            const nodes = [...document.querySelectorAll(
                'li, option, mat-option, .ng-option, [role="option"], div, span, a'
            )];
            const items = nodes.filter(e => {
                const t = norm(e.innerText);
                const r = e.getBoundingClientRect();
                if (!r.width || !r.height || r.height > 60) return false;
                if (t.length < 8 || t.length > 90) return false;
                if (/select one|save as draft|generate otp|attach digital|upload/i.test(t)) return false;
                if (want === 'return') return /return\\s+to/i.test(t);
                return /forward\\s+to/i.test(t);
            });
            items.sort((a, b) => norm(a.innerText).length - norm(b.innerText).length);
            if (!items.length) {
                return {ok: false, visible: nodes.filter(e => {
                    const r = e.getBoundingClientRect();
                    return r.width && r.height && r.height < 50;
                }).slice(0, 12).map(e => norm(e.innerText)).filter(Boolean)};
            }
            const t = norm(items[0].innerText);
            items[0].click();
            return {ok: true, picked: t};
        }""",
        want,
    )
    log(f"  Action option: {picked}")
    if not picked or not picked.get("ok"):
        save_debug(page, "action_select_fail")
        return False
    page.wait_for_timeout(700)
    return True


def _wait_button(page, pattern: str, ms: int = 8000) -> bool:
    deadline = datetime.now().timestamp() + ms / 1000
    rx = re.compile(pattern, re.I)
    while datetime.now().timestamp() < deadline:
        try:
            if page.get_by_role("button", name=rx).count():
                return True
        except Exception:
            pass
        found = page.evaluate(
            """(pat) => {
                const re = new RegExp(pat, 'i');
                return [...document.querySelectorAll('button, a.btn, .btn')].some(b => {
                    const r = b.getBoundingClientRect();
                    return r.width && r.height && re.test((b.innerText || '').trim());
                });
            }""",
            pattern,
        )
        if found:
            return True
        page.wait_for_timeout(200)
    return False


def generate_otp_and_submit(page, btype: str) -> bool:
    if not _wait_button(page, r"generate otp", 10000):
        log("  GENERATE OTP button not visible yet")
        save_debug(page, "no_generate_otp")
        return False
    try:
        btn = page.get_by_role("button", name=re.compile("generate otp", re.I))
        if btn.count():
            btn.first.click(timeout=8000)
        else:
            page.evaluate(
                """() => {
                    const b = [...document.querySelectorAll('button,a,.btn')].find(e =>
                        /generate otp/i.test(e.innerText || '')
                    );
                    if (b) b.click();
                }"""
            )
        log("  clicked GENERATE OTP")
    except Exception as exc:
        log(f"  GENERATE OTP: {exc}")
        return False
    page.wait_for_timeout(600)
    click_modal_ok(page, wait_ms=12000)
    page.wait_for_timeout(400)
    otp = page.get_by_placeholder(re.compile("otp", re.I))
    if otp.count() == 0:
        otp = page.locator("xpath=//*[contains(normalize-space(.),'Please Enter OTP')][1]/following::input[1]")
    if otp.count() == 0:
        otp = page.locator("input[name*='otp' i], input[id*='otp' i], input[placeholder*='OTP' i]")
    if otp.count():
        _type_into(page, otp.first, "000000")
        log("  OTP typed 000000")
    else:
        log("  OTP box not found")
        save_debug(page, "otp_missing")
        return False
    page.wait_for_timeout(500)
    btn_pat = r"return to" if btype == "return" else r"forward to"
    if not _wait_button(page, btn_pat, 8000):
        log(f"  final {btype} button not found")
        save_debug(page, "final_action_missing")
        return False
    clicked = page.evaluate(
        """(want) => {
            const btns = [...document.querySelectorAll('button, a.btn, .btn')].filter(b => {
                const t = (b.innerText || '').replace(/\\s+/g, ' ').trim();
                const r = b.getBoundingClientRect();
                if (!r.width || !r.height) return false;
                if (/save as draft|generate otp|attach digital|upload|select one/i.test(t)) return false;
                if (want === 'return') return /return\\s+to/i.test(t);
                return /forward\\s+to/i.test(t);
            });
            if (!btns.length) return {ok: false};
            const t = (btns[0].innerText || '').replace(/\\s+/g, ' ').trim();
            btns[0].click();
            return {ok: true, text: t};
        }""",
        "return" if btype == "return" else "forward",
    )
    log(f"  clicked final action: {clicked}")
    if not clicked or not clicked.get("ok"):
        save_debug(page, "final_action_fail")
        return False
    page.wait_for_timeout(600)
    click_modal_ok(page, wait_ms=8000)
    try:
        page.wait_for_url(re.compile(r"/success/"), timeout=20000)
        log(f"  success page: {page.url}")
    except Exception:
        log(f"  waiting success page — now {page.url}")
        click_modal_ok(page, wait_ms=3000)
        try:
            page.wait_for_url(re.compile(r"/success/"), timeout=10000)
            log(f"  success page: {page.url}")
        except Exception:
            save_debug(page, "no_success_page")
            return False
    page.wait_for_timeout(400)
    return True


def load_batch(folder: Path) -> list[dict]:
    """Batch folder: Upload Template.xlsx + PDFs named '{Letter} TPI Report...'."""
    folder = Path(folder)
    if not folder.exists():
        log(f"Batch folder not found: {folder}")
        return []
    xlsx = None
    for p in folder.glob("*.xlsx"):
        if p.name.startswith("~$"):
            continue
        xlsx = p
        if "upload" in p.name.lower() or "template" in p.name.lower():
            break
    if not xlsx:
        log("No Excel (Upload Template) in batch folder")
        return []
    from openpyxl import load_workbook
    wb = load_workbook(xlsx, data_only=True, read_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(max_row=500, max_col=16, values_only=True))
    wb.close()
    if not rows:
        return []
    headers = [re.sub(r"\s+", " ", str(c or "")).strip().lower() for c in rows[0]]

    def col(*keys):
        for i, h in enumerate(headers):
            if any(k in h for k in keys):
                return i
        return None

    isid = col("scheme")
    imb = col("mb")
    iamt = col("amount")
    ilet = col("letter")
    irem = col("remark")
    istat = col("status", "bill type", "action")
    log(f"Batch excel: {xlsx.name} headers={headers}")
    bills = []
    for raw in rows[1:]:
        if not raw or all(v is None or str(v).strip() == "" for v in raw):
            continue
        def cell(i):
            if i is None or i >= len(raw):
                return ""
            return str(raw[i] or "").strip()
        sid = cell(isid)
        mb = cell(imb)
        letter = cell(ilet)
        status = cell(istat)
        btype = _norm_bill_type(status)
        if status.lower().startswith("verif") or status.lower() == "forward":
            btype = "forward"
        amount = "0.00" if btype == "return" else cell(iamt)
        files = []
        if letter:
            pat = re.compile(rf"^{re.escape(letter)}(?:[\s_\-]|$)", re.I)
            for p in sorted(folder.iterdir()):
                if not p.is_file() or p.name.startswith("~$"):
                    continue
                if p.suffix.lower() in {".xlsx", ".xls"}:
                    continue
                if pat.match(p.name):
                    files.append(str(p))
        bills.append({
            "scheme_id": sid,
            "mb_no": mb.replace("-", "/"),
            "district": "",
            "path": str(folder),
            "files": files,
            "n_files": len(files),
            "fill": {
                "amount": amount,
                "letter": letter,
                "remark": cell(irem),
                "bill_type": btype,
                "status": status,
            },
        })
    return bills


def parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(description="TPI MB uploader")
    p.add_argument("--headed", action="store_true")
    p.add_argument("--scan-only", action="store_true")
    p.add_argument("--batch", default="", help="Folder with Upload Template.xlsx + TPI Report PDFs")
    p.add_argument("--day", default="", help="DD.MM.YYYY (legacy)")
    p.add_argument("--session", default="")
    p.add_argument("--skip-login", action="store_true")
    return p.parse_args()


def main() -> None:
    args = parse_args()
    batch = (args.batch or os.getenv("TPI_BATCH") or "").strip()
    if batch:
        log(f"Batch folder: {batch}")
        bills = load_batch(Path(batch))
    else:
        root = DOWNLOAD_DIR
        day = args.day.strip() or today_stamp()
        session = args.session.strip()
        if not session:
            sess = list_sessions(root, day)
            session = sess[-1] if sess else "Session 1"
        log(f"Upload root: {root}")
        log(f"Day: {day}  Session: {session}")
        bills = scan_session(root, day, session)
    pending = [
        b for b in bills
        if b["n_files"] or (b.get("fill") or {}).get("letter") or (b.get("fill") or {}).get("amount")
    ]
    only = (os.getenv("TPI_ONLY") or "").strip()
    if only:
        keys = {k.strip() for k in only.split(";") if k.strip()}
        pending = [b for b in pending if f"{b.get('scheme_id')}|{b.get('mb_no')}" in keys]
        log(f"Filtered to selected rows: {len(pending)}")
    log(f"Rows: {len(bills)} | ready: {len(pending)}")
    for b in bills:
        fill = b.get("fill") or {}
        log(f"  {b.get('scheme_id')} | {b.get('mb_no')} | letter={fill.get('letter')} | {fill.get('bill_type')} | files={b['n_files']}")
        for f in b["files"]:
            log(f"      {Path(f).name}")

    if args.scan_only:
        log("Scan only. No portal upload.")
        return
    if not pending:
        log("Nothing to upload. Create For Upload\\ with files and/or an Excel (Amount, Letter, Remark).")
        return
    steps = {s.strip() for s in (os.getenv("TPI_STEPS") or "files,fill,dsc,action").split(",") if s.strip()}
    log(f"Steps: {sorted(steps)}")
    if not USER or not PASSWORD:
        sys.exit("ERROR: PORTAL_USER / PORTAL_PASS missing")

    with sync_playwright() as pw:
        browser = pw.chromium.launch(headless=not args.headed)
        context = browser.new_context(accept_downloads=True)
        page = context.new_page()
        try:
            if not args.skip_login:
                login(page)
            go_to_tpi_list(page)
            for i, b in enumerate(pending, 1):
                log(f"[{i}/{len(pending)}] upload {b['scheme_id']} {b['mb_no']}")
                row = {"scheme_id": b["scheme_id"], "mb_no": b["mb_no"]}
                try:
                    go_to_tpi_list(page)
                    detail = open_row(page, row)
                    fill = b.get("fill") or {}
                    btype = _norm_bill_type(fill.get("bill_type") or fill.get("status"))
                    grand = scrape_grand_total(detail)
                    if grand:
                        log(f"  Grand Total: {grand}")
                    if btype == "return":
                        fill["amount"] = "0.00"
                    log(f"  bill type: {btype}  steps={sorted(steps)}")
                    if "files" in steps:
                        uploaded = upload_files_on_detail(detail, b["files"])
                        log(f"  files uploaded: {len(uploaded)}")
                    saved = True
                    if "fill" in steps:
                        fill_tpi_fields(detail, fill)
                        saved = click_save_as_draft(detail)
                        if not saved:
                            log("  retry fill letter + save")
                            fill_tpi_fields(detail, fill)
                            saved = click_save_as_draft(detail)
                    if "dsc" in steps and btype == "forward":
                        click_attach_signature(detail)
                    elif "dsc" in steps:
                        log("  Return bill — skip DSC")
                    if "action" in steps:
                        if select_action(detail, btype):
                            ok_act = generate_otp_and_submit(detail, btype)
                        else:
                            ok_act = False
                            log("  skip OTP — action not selected")
                        try:
                            batch = Path(b.get("path") or os.getenv("TPI_BATCH") or ".")
                            tag = "Return" if btype == "return" else "Forward"
                            st = "OK" if ok_act else "FAIL"
                            fn = f"{b.get('scheme_id')}_{win_name(b.get('mb_no') or '', 20)}_{tag}_{st}_{datetime.now().strftime('%H%M%S')}.png"
                            dest = batch / fn
                            detail.screenshot(path=str(dest), full_page=True)
                            log(f"  screenshot saved: {dest.name}")
                        except Exception as exc:
                            log(f"  screenshot: {exc}")
                    try:
                        from activity import log_event
                        log_event(
                            kind="upload",
                            action="return" if btype == "return" else "forward",
                            cluster=os.getenv("TPI_CLUSTER") or "",
                            district=b.get("district") or "",
                            scheme_id=b.get("scheme_id") or "",
                            mb_no=b.get("mb_no") or "",
                            amount=(fill or {}).get("amount") or "",
                            grand_total=grand,
                            letter=(fill or {}).get("letter") or "",
                        )
                    except Exception as exc:
                        log(f"  activity: {exc}")
                    if detail != page:
                        try:
                            detail.close()
                        except Exception:
                            pass
                    back_to_tpi_list(page)
                except Exception as exc:
                    log(f"  ERROR: {exc}")
                    save_debug(page, f"upload_err_{b['scheme_id']}")
                    try:
                        go_to_tpi_list(page)
                    except Exception:
                        pass
        finally:
            context.close()
            browser.close()
    log("Upload finished.")


if __name__ == "__main__":
    main()
