import getpass
import json
import socket
import sys
import urllib.error
import urllib.request
from pathlib import Path

LICENSE_VERIFY_URL = "https://script.google.com/macros/s/AKfycby987WffVGZXKoPt8HzEJttF_IBJg_GRNZh3xl6xSrM1QHGREkTM18m0lYLvrVMQdAZ/exec"
LICENSE_FILE_NAME = ".license_key.json"
APP_NAME = "ppo_automation"
APP_VERSION = "8"


def get_runtime_dir():
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def read_saved_license_key():
    path = get_runtime_dir() / LICENSE_FILE_NAME
    if not path.is_file():
        return ""
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
        return str(data.get("license_key", "")).strip()
    except Exception:
        return ""


def save_license_key(key):
    path = get_runtime_dir() / LICENSE_FILE_NAME
    payload = {"license_key": str(key).strip()}
    path.write_text(json.dumps(payload, ensure_ascii=False), encoding="utf-8")


def verify_license_online(license_key, timeout=12):
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
    try:
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            raw = resp.read().decode("utf-8", errors="ignore")
        obj = json.loads(raw)
        allowed = bool(obj.get("allowed", False))
        reason = str(obj.get("reason", "")).strip() or "unknown"
        return allowed, reason
    except urllib.error.URLError:
        return False, "network_error"
    except Exception:
        return False, "verify_error"


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
