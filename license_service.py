import getpass
import json
import socket
import sys
import time
import urllib.error
import urllib.request
from pathlib import Path

LICENSE_VERIFY_URL = "https://script.google.com/macros/s/AKfycby987WffVGZXKoPt8HzEJttF_IBJg_GRNZh3xl6xSrM1QHGREkTM18m0lYLvrVMQdAZ/exec"
LICENSE_FILE_NAME = ".license_key.json"
APP_NAME = "ppo_automation"
APP_VERSION = "8"
LICENSE_CACHE_TTL_SECONDS = 3 * 24 * 60 * 60


def get_runtime_dir():
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def _license_file_path():
    return get_runtime_dir() / LICENSE_FILE_NAME


def _read_license_payload():
    path = _license_file_path()
    if not path.is_file():
        return {}
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
        return data if isinstance(data, dict) else {}
    except Exception:
        return {}


def _write_license_payload(payload):
    path = _license_file_path()
    path.write_text(json.dumps(payload, ensure_ascii=False), encoding="utf-8")


def read_saved_license_key():
    data = _read_license_payload()
    return str(data.get("license_key", "")).strip()


def save_license_key(key):
    payload = _read_license_payload()
    payload["license_key"] = str(key).strip()
    _write_license_payload(payload)


def mark_license_verified(key):
    payload = _read_license_payload()
    payload["license_key"] = str(key).strip()
    payload["last_verified_key"] = str(key).strip()
    payload["last_verified_at"] = int(time.time())
    _write_license_payload(payload)


def has_recent_verified_license(key, max_age_seconds=LICENSE_CACHE_TTL_SECONDS):
    payload = _read_license_payload()
    saved_key = str(payload.get("last_verified_key", "")).strip()
    verified_at = payload.get("last_verified_at")
    if saved_key != str(key or "").strip():
        return False
    if not isinstance(verified_at, int):
        return False
    return (time.time() - verified_at) <= max(0, int(max_age_seconds))


def verify_license_online(license_key, timeout=12, retries=3):
    payload = {
        "action": "verify",
        "license_key": str(license_key or "").strip(),
        "app_name": APP_NAME,
        "app_version": APP_VERSION,
    }
    data = json.dumps(payload).encode("utf-8")
    req = urllib.request.Request(
        LICENSE_VERIFY_URL,
        data=data,
        headers={"Content-Type": "application/json"},
        method="POST",
    )
    last_reason = "verify_error"
    attempts = max(1, int(retries or 1))

    for attempt in range(attempts):
        try:
            with urllib.request.urlopen(req, timeout=timeout) as resp:
                raw = resp.read().decode("utf-8", errors="ignore")
            obj = json.loads(raw)
            allowed = bool(obj.get("allowed", False))
            reason = str(obj.get("reason", "")).strip() or "unknown"
            if allowed:
                mark_license_verified(license_key)
            return allowed, reason
        except urllib.error.URLError:
            last_reason = "network_error"
        except Exception:
            last_reason = "verify_error"

        if attempt + 1 < attempts:
            time.sleep(1.0 + attempt)

    if last_reason in {"network_error", "verify_error"} and has_recent_verified_license(license_key):
        return True, "cached_ok"

    return False, last_reason


def log_usage_online(license_key, submitted_requests=0, timeout=10):
    payload = {
        "action": "log_usage",
        "license_key": str(license_key or "").strip(),
        "app_name": APP_NAME,
        "app_version": APP_VERSION,
        "pc_name": socket.gethostname(),
        "username": getpass.getuser(),
        "submitted_requests": int(submitted_requests or 0),
    }
    data = json.dumps(payload).encode("utf-8")
    req = urllib.request.Request(
        LICENSE_VERIFY_URL,
        data=data,
        headers={"Content-Type": "application/json"},
        method="POST",
    )
    try:
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            raw = resp.read().decode("utf-8", errors="ignore")
        obj = json.loads(raw)
        return bool(obj.get("ok", False) or obj.get("logged", False) or obj.get("allowed", False))
    except Exception:
        return False
