"""Master password vault. Portal passwords stay encrypted. Session lasts 15 days."""
from __future__ import annotations

import base64
import hashlib
import json
import os
from datetime import datetime, timedelta
from pathlib import Path

SESSION_DAYS = 15
DATA = Path(__file__).resolve().parent / ".tpidata"
VAULT_PATH = DATA / "vault.bin"

_key: bytes | None = None


class VaultError(Exception):
    pass


def _fernet(key: bytes):
    from cryptography.fernet import Fernet

    return Fernet(base64.urlsafe_b64encode(key))


def _derive(password: str, salt: bytes) -> bytes:
    return hashlib.pbkdf2_hmac("sha256", password.encode("utf-8"), salt, 200_000, dklen=32)


def _dpapi_protect(raw: bytes) -> bytes | None:
    if os.name != "nt":
        return None
    try:
        import ctypes
        import ctypes.wintypes as w

        class DATA_BLOB(ctypes.Structure):
            _fields_ = [("cbData", w.DWORD), ("pbData", ctypes.POINTER(ctypes.c_char))]

        crypt32 = ctypes.windll.crypt32
        kernel32 = ctypes.windll.kernel32
        blob_in = DATA_BLOB(len(raw), ctypes.create_string_buffer(raw, len(raw)))
        blob_out = DATA_BLOB()
        if not crypt32.CryptProtectData(
            ctypes.byref(blob_in), None, None, None, None, 0, ctypes.byref(blob_out)
        ):
            return None
        try:
            out = ctypes.string_at(blob_out.pbData, blob_out.cbData)
        finally:
            kernel32.LocalFree(blob_out.pbData)
        return out
    except Exception:
        return None


def _dpapi_unprotect(raw: bytes) -> bytes | None:
    if os.name != "nt":
        return None
    try:
        import ctypes
        import ctypes.wintypes as w

        class DATA_BLOB(ctypes.Structure):
            _fields_ = [("cbData", w.DWORD), ("pbData", ctypes.POINTER(ctypes.c_char))]

        crypt32 = ctypes.windll.crypt32
        kernel32 = ctypes.windll.kernel32
        blob_in = DATA_BLOB(len(raw), ctypes.create_string_buffer(raw, len(raw)))
        blob_out = DATA_BLOB()
        if not crypt32.CryptUnprotectData(
            ctypes.byref(blob_in), None, None, None, None, 0, ctypes.byref(blob_out)
        ):
            return None
        try:
            out = ctypes.string_at(blob_out.pbData, blob_out.cbData)
        finally:
            kernel32.LocalFree(blob_out.pbData)
        return out
    except Exception:
        return None


def _load() -> dict:
    if not VAULT_PATH.exists():
        return {}
    try:
        return json.loads(VAULT_PATH.read_text(encoding="utf-8"))
    except Exception:
        return {}


def _save(data: dict) -> None:
    DATA.mkdir(parents=True, exist_ok=True)
    VAULT_PATH.write_text(json.dumps(data), encoding="utf-8")
    try:
        import subprocess

        subprocess.run(
            ["attrib", "+h", "+s", str(DATA)],
            check=False,
            creationflags=0x08000000 if os.name == "nt" else 0,
        )
    except Exception:
        pass


def master_is_set() -> bool:
    d = _load()
    return bool(d.get("salt") and d.get("hash"))


def days_left() -> int:
    d = _load()
    until = d.get("session_until") or ""
    if not until:
        return 0
    try:
        end = datetime.fromisoformat(until)
    except Exception:
        return 0
    n = (end - datetime.now()).days
    return max(0, n)


def _set_session(key: bytes) -> None:
    global _key
    _key = key
    d = _load()
    d["session_until"] = (datetime.now() + timedelta(days=SESSION_DAYS)).isoformat(timespec="seconds")
    wrapped = _dpapi_protect(key)
    if wrapped:
        d["session_key"] = base64.b64encode(wrapped).decode("ascii")
    _save(d)


def session_unlock() -> bool:
    """True if 15-day session is still valid (no password prompt)."""
    global _key
    d = _load()
    if not d.get("salt"):
        return False
    until = d.get("session_until") or ""
    try:
        if not until or datetime.fromisoformat(until) < datetime.now():
            return False
    except Exception:
        return False
    blob = d.get("session_key") or ""
    if not blob:
        return False
    try:
        key = _dpapi_unprotect(base64.b64decode(blob))
    except Exception:
        key = None
    if not key or len(key) != 32:
        return False
    _key = key
    return True


def set_master(password: str) -> None:
    password = (password or "").strip()
    if len(password) < 8:
        raise VaultError("At least 8 characters.")
    salt = os.urandom(16)
    key = _derive(password, salt)
    digest = hashlib.sha256(key).hexdigest()
    prev = _load()
    _save({
        "salt": base64.b64encode(salt).decode("ascii"),
        "hash": digest,
        "sealed": prev.get("sealed") or {},
    })
    _set_session(key)


def verify_master(password: str) -> None:
    d = _load()
    if not d.get("salt"):
        raise VaultError("Master password is not set.")
    salt = base64.b64decode(d["salt"])
    key = _derive(password or "", salt)
    digest = hashlib.sha256(key).hexdigest()
    if not hashlib.compare_digest(digest, d.get("hash") or ""):
        raise VaultError("Incorrect password.")
    _set_session(key)


def protect(plain: str) -> str:
    if _key is None:
        raise VaultError("Vault is locked.")
    return _fernet(_key).encrypt((plain or "").encode("utf-8")).decode("ascii")


def unprotect(token: str) -> str:
    if not token:
        return ""
    if _key is None:
        raise VaultError("Vault is locked.")
    return _fernet(_key).decrypt(token.encode("ascii")).decode("utf-8")


def seal_setting(name: str, value: str) -> None:
    d = _load()
    sealed = dict(d.get("sealed") or {})
    sealed[name] = protect(value or "")
    d["sealed"] = sealed
    _save(d)


def unseal_setting(name: str) -> str:
    d = _load()
    tok = (d.get("sealed") or {}).get(name) or ""
    if not tok:
        return ""
    try:
        return unprotect(tok)
    except Exception:
        return ""
