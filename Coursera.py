import requests
import time
import random
import string
import os
import re
import glob
import traceback
from datetime import datetime
import threading
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import TimeoutException
import concurrent.futures

# ================= Config =================
# RoxyBrowser local API config
API_HOST = "http://127.0.0.1:50000"
API_TOKEN = "4ba21591e96dad03142b86e7ef106598"  
WORKSPACE_ID = os.getenv("ROXY_WORKSPACE_ID", "").strip()
HEADERS = {"token": API_TOKEN}
_WORKSPACE_ID_CACHE = ""
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
ACCOUNT_FILE = os.path.join(SCRIPT_DIR, "account.xml")
LINK_FILE = os.path.join(SCRIPT_DIR, "link.xml")
WORKSPACE_FILE = os.path.join(SCRIPT_DIR, "workspace_id.txt")
PROFILE_FILE = os.path.join(SCRIPT_DIR, "profile_ids.txt")
DEBUG_DIR = os.path.join(SCRIPT_DIR, "debug_artifacts")
LINK_FILE_LOCK = threading.Lock()
RUN_LINKS = set()
COURSE_ENROLL_URL = "https://www.coursera.org/professional-certificates/google-ai?action=enroll"
ACTION_DELAY_SECONDS = 0.0
_ACTION_DELAY_CACHE_KEY = None
_ACTION_DELAY_RANGE_CACHE = (0.0, 0.0)
_HUMAN_SCROLL_RANGE_CACHE_KEY = None
_HUMAN_SCROLL_RANGE_CACHE = (0, 0)
_HUMAN_SCROLL_PROB_CACHE_KEY = None
_HUMAN_SCROLL_PROB_CACHE = 0.0
# ==========================================

def _parse_non_negative_float(raw, default=0.0):
    try:
        value = float((raw or "").strip())
        if value < 0:
            raise ValueError
        return value
    except Exception:
        return default

def get_action_delay_range_seconds():
    global _ACTION_DELAY_CACHE_KEY, _ACTION_DELAY_RANGE_CACHE

    raw = os.getenv("ACTION_DELAY_SECONDS", "0.35,0.9").strip()
    cache_key = raw
    if cache_key == _ACTION_DELAY_CACHE_KEY:
        return _ACTION_DELAY_RANGE_CACHE

    min_delay = 0.0
    max_delay = 0.0
    try:
        if "," in raw:
            left, right = [x.strip() for x in raw.split(",", 1)]
            min_delay = _parse_non_negative_float(left, 0.0)
            max_delay = _parse_non_negative_float(right, min_delay)
            if max_delay < min_delay:
                min_delay, max_delay = max_delay, min_delay
        else:
            fixed = _parse_non_negative_float(raw, 0.0)
            min_delay = fixed
            max_delay = fixed
    except Exception:
        min_delay = 0.0
        max_delay = 0.0

    _ACTION_DELAY_CACHE_KEY = cache_key
    _ACTION_DELAY_RANGE_CACHE = (min_delay, max_delay)
    return _ACTION_DELAY_RANGE_CACHE

def get_startup_stagger_seconds(task_index):
    base = _parse_non_negative_float(os.getenv("STARTUP_STAGGER_SECONDS", "0"), 0.0)
    jitter = _parse_non_negative_float(os.getenv("STARTUP_STAGGER_JITTER_SECONDS", "0"), 0.0)
    if base <= 0 and jitter <= 0:
        return 0.0
    return task_index * base + (random.uniform(0.0, jitter) if jitter > 0 else 0.0)

def demo_pause(profile_id="", note=""):
    min_delay, max_delay = get_action_delay_range_seconds()
    if min_delay <= 0 and max_delay <= 0:
        return
    delay = min_delay if max_delay <= min_delay else random.uniform(min_delay, max_delay)
    if delay <= 0:
        return
    prefix = f"[{profile_id}] " if profile_id else ""
    if note:
        print(f"{prefix}Demo pause {delay:.1f}s: {note}")
    time.sleep(delay)

def _parse_int_range(raw, default_min, default_max):
    min_val = default_min
    max_val = default_max
    try:
        if "," in (raw or ""):
            left, right = [x.strip() for x in raw.split(",", 1)]
            min_val = int(round(_parse_non_negative_float(left, float(default_min))))
            max_val = int(round(_parse_non_negative_float(right, float(default_max))))
        else:
            fixed = int(round(_parse_non_negative_float(raw, float(default_min))))
            min_val = fixed
            max_val = fixed
    except Exception:
        min_val = default_min
        max_val = default_max
    min_val = max(0, min_val)
    max_val = max(0, max_val)
    if max_val < min_val:
        min_val, max_val = max_val, min_val
    return min_val, max_val

def get_human_scroll_range_pixels():
    global _HUMAN_SCROLL_RANGE_CACHE_KEY, _HUMAN_SCROLL_RANGE_CACHE

    raw = os.getenv("HUMAN_SCROLL_PIXELS", "70,220").strip()
    if raw == _HUMAN_SCROLL_RANGE_CACHE_KEY:
        return _HUMAN_SCROLL_RANGE_CACHE
    _HUMAN_SCROLL_RANGE_CACHE_KEY = raw
    _HUMAN_SCROLL_RANGE_CACHE = _parse_int_range(raw, 70, 220)
    return _HUMAN_SCROLL_RANGE_CACHE

def get_human_scroll_probability():
    global _HUMAN_SCROLL_PROB_CACHE_KEY, _HUMAN_SCROLL_PROB_CACHE

    raw = os.getenv("HUMAN_SCROLL_PROBABILITY", "0.55").strip()
    if raw == _HUMAN_SCROLL_PROB_CACHE_KEY:
        return _HUMAN_SCROLL_PROB_CACHE
    _HUMAN_SCROLL_PROB_CACHE_KEY = raw
    try:
        value = float(raw)
    except Exception:
        value = 0.55
    _HUMAN_SCROLL_PROB_CACHE = min(1.0, max(0.0, value))
    return _HUMAN_SCROLL_PROB_CACHE

def _maybe_random_scroll(driver, profile_id="", note=""):
    min_px, max_px = get_human_scroll_range_pixels()
    if min_px <= 0 and max_px <= 0:
        return False
    if random.random() > get_human_scroll_probability():
        return False

    distance = min_px if max_px <= min_px else random.randint(min_px, max_px)
    if distance <= 0:
        return False
    direction = -1 if random.random() < 0.45 else 1
    delta = direction * distance

    try:
        driver.execute_script("window.scrollBy({top: arguments[0], left: 0, behavior: 'smooth'});", delta)
    except Exception:
        try:
            driver.execute_script("window.scrollBy(0, arguments[0]);", delta)
        except Exception:
            return False

    prefix = f"[{profile_id}] " if profile_id else ""
    if note:
        print(f"{prefix}Random scroll {delta}px: {note}")
    time.sleep(random.uniform(0.08, 0.24))
    return True

def _pause_between_actions(driver=None, profile_id="", note="", allow_scroll=False):
    if allow_scroll and driver is not None:
        _maybe_random_scroll(driver, profile_id=profile_id, note=note)
    demo_pause(profile_id=profile_id, note=note)

def _api_get(path, params=None):
    resp = requests.get(f"{API_HOST}{path}", params=params, headers=HEADERS, timeout=10)
    try:
        return resp.json()
    except Exception:
        return {"code": resp.status_code, "msg": resp.text[:200], "data": None}

def _api_post(path, data=None):
    resp = requests.post(f"{API_HOST}{path}", json=data, headers=HEADERS, timeout=10)
    try:
        return resp.json()
    except Exception:
        return {"code": resp.status_code, "msg": resp.text[:200], "data": None}

def _extract_workspace_id_from_logs():
    """Fallback: read workspaceId from local RoxyBrowser logs."""
    appdata = os.getenv("APPDATA", "")
    log_dir = os.path.join(appdata, "RoxyBrowser", "logs")
    if not os.path.isdir(log_dir):
        return ""
    pattern = re.compile(r"workspaceId[\s:'\"]+(\d+)")
    log_files = sorted(glob.glob(os.path.join(log_dir, "*.log")), reverse=True)
    for log_file in log_files:
        try:
            with open(log_file, "r", encoding="utf-8", errors="ignore") as f:
                content = f.read()
            matches = pattern.findall(content)
            if matches:
                return matches[-1]
        except Exception:
            continue
    return ""

def save_workspace_id(workspace_id, output_file=WORKSPACE_FILE):
    with open(output_file, "w", encoding="utf-8") as f:
        f.write(str(workspace_id).strip() + "\n")

def get_workspace_id():
    """Get workspaceId: env -> API -> local logs."""
    global _WORKSPACE_ID_CACHE
    if _WORKSPACE_ID_CACHE:
        return _WORKSPACE_ID_CACHE
    if WORKSPACE_ID:
        _WORKSPACE_ID_CACHE = WORKSPACE_ID
        return _WORKSPACE_ID_CACHE
    try:
        rsp = _api_get("/browser/workspace", {"page_index": 1, "page_size": 50})
        if rsp.get("code") == 0:
            rows = rsp.get("data", {}).get("rows", [])
            if rows and rows[0].get("id") is not None:
                _WORKSPACE_ID_CACHE = str(rows[0]["id"])
                return _WORKSPACE_ID_CACHE
        else:
            print(f"[WorkspaceAPI] code={rsp.get('code')}, msg={rsp.get('msg')}")
    except Exception as e:
        print(f"[WorkspaceAPI] request failed: {e}")
    workspace_id = _extract_workspace_id_from_logs()
    if workspace_id:
        print(f"[WorkspaceAPI] fallback workspaceId from logs: {workspace_id}")
    _WORKSPACE_ID_CACHE = workspace_id
    return _WORKSPACE_ID_CACHE

def _collect_profile_ids(node, profile_ids):
    """Recursively collect profile IDs from JSON payload."""
    if isinstance(node, dict):
        profile_id = node.get("profileId") or node.get("profile_id")
        if profile_id:
            profile_ids.add(str(profile_id))
        for key in ("id", "_id"):
            value = node.get(key)
            if value and isinstance(value, (str, int)):
                profile_ids.add(str(value))
        for value in node.values():
            _collect_profile_ids(value, profile_ids)
    elif isinstance(node, list):
        for item in node:
            _collect_profile_ids(item, profile_ids)

def get_all_profile_ids():
    """Fetch all profile IDs (dirId) using workspaceId."""
    workspace_id = get_workspace_id()
    if not workspace_id:
        print("[ProfileAPI] workspaceId not found.")
        return []
    try:
        rsp = _api_get(
            "/browser/list_v3",
            {"workspaceId": workspace_id, "page_index": 1, "page_size": 200},
        )
        if rsp.get("code") == 0:
            rows = rsp.get("data", {}).get("rows", [])
            profile_ids = sorted(str(row["dirId"]) for row in rows if row.get("dirId"))
            if profile_ids:
                return profile_ids
        print(f"[ProfileAPI] /browser/list_v3 -> code={rsp.get('code')}, msg={rsp.get('msg')}")
    except Exception as e:
        print(f"[ProfileAPI] /browser/list_v3 failed: {e}")

    # Backward-compatible fallback.
    try:
        rsp = _api_get("/browser/list", {"workspaceId": workspace_id, "page_index": 1, "page_size": 200})
        payload = rsp.get("data", rsp)
        profile_ids = set()
        _collect_profile_ids(payload, profile_ids)
        return sorted(profile_ids)
    except Exception:
        return []

def save_profile_ids(profile_ids, output_file=PROFILE_FILE):
    """Write profile IDs to a local text file for quick verification."""
    with open(output_file, "w", encoding="utf-8") as f:
        for profile_id in profile_ids:
            f.write(profile_id + "\n")

def show_startup_config_dialog():
    """
    Show startup settings dialog and write values to environment variables.
    Set SHOW_CONFIG_UI=0 to skip this dialog.
    """
    if os.getenv("SHOW_CONFIG_UI", "1").strip().lower() in {"0", "false", "no"}:
        return

    try:
        import tkinter as tk
        from tkinter import messagebox
    except Exception as e:
        print(f"[ConfigUI] tkinter unavailable, skip dialog: {e}")
        return

    fields = [
        ("LOOP_COUNT", "循环轮次", os.getenv("LOOP_COUNT", "1")),
        ("MAX_WORKERS", "最大并发数（留空=全部）", os.getenv("MAX_WORKERS", "")),
        ("STARTUP_STAGGER_SECONDS", "启动错峰秒数", os.getenv("STARTUP_STAGGER_SECONDS", "0")),
        (
            "STARTUP_STAGGER_JITTER_SECONDS",
            "启动抖动秒数",
            os.getenv("STARTUP_STAGGER_JITTER_SECONDS", "0"),
        ),
        ("PROFILE_LIMIT", "Profile 数量上限（留空=全部）", os.getenv("PROFILE_LIMIT", "")),
        ("PROFILE_IDS", "指定 Profile IDs（逗号分隔，留空=全部）", os.getenv("PROFILE_IDS", "")),
        (
            "ACTION_DELAY_SECONDS",
            "动作延时秒数（固定值或最小,最大）",
            os.getenv("ACTION_DELAY_SECONDS", "0.35,0.9"),
        ),
        (
            "HUMAN_SCROLL_PIXELS",
            "随机滚动像素（最小,最大）",
            os.getenv("HUMAN_SCROLL_PIXELS", "70,220"),
        ),
        (
            "HUMAN_SCROLL_PROBABILITY",
            "随机滚动概率（0-1）",
            os.getenv("HUMAN_SCROLL_PROBABILITY", "0.55"),
        ),
        (
            "STEP3_EXTRA_WAIT_SECONDS",
            "第3步额外等待秒数",
            os.getenv("STEP3_EXTRA_WAIT_SECONDS", "1.2"),
        ),
        (
            "STEP4_RETRY_WAIT_SECONDS",
            "第4步重试等待秒数",
            os.getenv("STEP4_RETRY_WAIT_SECONDS", "1.6"),
        ),
        (
            "RANDOMIZE_FINGERPRINT_ON_START",
            "启动时随机指纹（0/1）",
            os.getenv("RANDOMIZE_FINGERPRINT_ON_START", "1"),
        ),
        (
            "CLEAR_COOKIES_ON_START",
            "启动时清理 Cookie（0/1）",
            os.getenv("CLEAR_COOKIES_ON_START", "1"),
        ),
        ("SIGNUP_RETRY_SECONDS", "注册自动重试秒数", os.getenv("SIGNUP_RETRY_SECONDS", "45")),
        (
            "MANUAL_VERIFY_WAIT_SECONDS",
            "人工验证等待秒数",
            os.getenv("MANUAL_VERIFY_WAIT_SECONDS", "180"),
        ),
    ]
    int_fields = {
        "LOOP_COUNT",
        "SIGNUP_RETRY_SECONDS",
        "MANUAL_VERIFY_WAIT_SECONDS",
    }
    all_tokens = {"all", "auto", "max", "*", "0"}

    root = tk.Tk()
    root.title("Coursera 自动化启动参数")
    root.resizable(False, False)

    vars_map = {}
    for row, (key, label, default) in enumerate(fields):
        tk.Label(root, text=label, anchor="w", width=30).grid(row=row, column=0, padx=8, pady=4, sticky="w")
        var = tk.StringVar(value=default)
        vars_map[key] = var
        tk.Entry(root, textvariable=var, width=45).grid(row=row, column=1, padx=8, pady=4, sticky="w")

    tk.Label(
        root,
        text="每轮会按并发配置运行所有选中的 Profile。",
        anchor="w",
    ).grid(row=len(fields), column=0, columnspan=2, padx=8, pady=(4, 8), sticky="w")

    cancelled = {"value": True}

    def _apply_env():
        for key, _, _ in fields:
            value = vars_map[key].get().strip()
            if key in int_fields and value:
                try:
                    parsed = int(value)
                    if parsed <= 0:
                        raise ValueError
                except Exception:
                    messagebox.showerror("参数错误", f"{key} 必须为正整数。")
                    return
            if key == "MAX_WORKERS" and value:
                lower = value.lower()
                if lower not in all_tokens:
                    try:
                        parsed = int(value)
                        if parsed <= 0:
                            raise ValueError
                    except Exception:
                        messagebox.showerror(
                            "参数错误",
                            "MAX_WORKERS 必须是正整数，或 all/auto/max/*/0。",
                        )
                        return
            if key == "PROFILE_LIMIT" and value:
                lower = value.lower()
                if lower not in all_tokens:
                    try:
                        parsed = int(value)
                        if parsed <= 0:
                            raise ValueError
                    except Exception:
                        messagebox.showerror(
                            "参数错误",
                            "PROFILE_LIMIT 必须是正整数，或 all/auto/max/*/0。",
                        )
                        return
            if key == "ACTION_DELAY_SECONDS" and value:
                try:
                    if "," in value:
                        left, right = [x.strip() for x in value.split(",", 1)]
                        min_delay = float(left)
                        max_delay = float(right)
                        if min_delay < 0 or max_delay < 0:
                            raise ValueError
                    else:
                        parsed = float(value)
                        if parsed < 0:
                            raise ValueError
                except Exception:
                    messagebox.showerror(
                        "参数错误",
                        "ACTION_DELAY_SECONDS 必须 >= 0，或使用 0.2,0.8 这种格式。",
                    )
                    return
            if key == "HUMAN_SCROLL_PIXELS" and value:
                try:
                    if "," in value:
                        left, right = [x.strip() for x in value.split(",", 1)]
                        min_px = float(left)
                        max_px = float(right)
                        if min_px < 0 or max_px < 0:
                            raise ValueError
                    else:
                        parsed = float(value)
                        if parsed < 0:
                            raise ValueError
                except Exception:
                    messagebox.showerror(
                        "参数错误",
                        "HUMAN_SCROLL_PIXELS 必须 >= 0，或使用 70,220 这种格式。",
                    )
                    return
            if key == "HUMAN_SCROLL_PROBABILITY" and value:
                try:
                    parsed = float(value)
                    if parsed < 0 or parsed > 1:
                        raise ValueError
                except Exception:
                    messagebox.showerror(
                        "参数错误",
                        "HUMAN_SCROLL_PROBABILITY 必须在 0 到 1 之间。",
                    )
                    return
            if key in {"STARTUP_STAGGER_SECONDS", "STARTUP_STAGGER_JITTER_SECONDS"} and value:
                try:
                    parsed = float(value)
                    if parsed < 0:
                        raise ValueError
                except Exception:
                    messagebox.showerror("参数错误", f"{key} 必须 >= 0。")
                    return
            if key == "STEP3_EXTRA_WAIT_SECONDS" and value:
                try:
                    parsed = float(value)
                    if parsed < 0:
                        raise ValueError
                except Exception:
                    messagebox.showerror("参数错误", "STEP3_EXTRA_WAIT_SECONDS 必须 >= 0。")
                    return
            if key == "STEP4_RETRY_WAIT_SECONDS" and value:
                try:
                    parsed = float(value)
                    if parsed < 0:
                        raise ValueError
                except Exception:
                    messagebox.showerror("参数错误", "STEP4_RETRY_WAIT_SECONDS 必须 >= 0。")
                    return

        for key, _, _ in fields:
            value = vars_map[key].get().strip()
            if key in {"MAX_WORKERS", "PROFILE_LIMIT"} and value.lower() in all_tokens:
                value = ""
            if value:
                os.environ[key] = value
            else:
                os.environ.pop(key, None)
        cancelled["value"] = False
        root.destroy()

    def _cancel():
        cancelled["value"] = True
        root.destroy()

    btn_row = len(fields) + 1
    tk.Button(root, text="开始", width=14, command=_apply_env).grid(row=btn_row, column=0, padx=8, pady=8, sticky="w")
    tk.Button(root, text="取消", width=14, command=_cancel).grid(row=btn_row, column=1, padx=8, pady=8, sticky="e")
    root.protocol("WM_DELETE_WINDOW", _cancel)
    root.mainloop()

    if cancelled["value"]:
        print("[配置面板] 用户取消启动。")
        raise SystemExit(0)

    print(
        "[配置面板] 已应用参数: "
        f"LOOP_COUNT={os.getenv('LOOP_COUNT', '')}, "
        f"MAX_WORKERS={os.getenv('MAX_WORKERS', '')}, "
        f"STARTUP_STAGGER_SECONDS={os.getenv('STARTUP_STAGGER_SECONDS', '')}, "
        f"STARTUP_STAGGER_JITTER_SECONDS={os.getenv('STARTUP_STAGGER_JITTER_SECONDS', '')}, "
        f"ACTION_DELAY_SECONDS={os.getenv('ACTION_DELAY_SECONDS', '')}, "
        f"HUMAN_SCROLL_PIXELS={os.getenv('HUMAN_SCROLL_PIXELS', '')}, "
        f"HUMAN_SCROLL_PROBABILITY={os.getenv('HUMAN_SCROLL_PROBABILITY', '')}, "
        f"STEP3_EXTRA_WAIT_SECONDS={os.getenv('STEP3_EXTRA_WAIT_SECONDS', '')}, "
        f"STEP4_RETRY_WAIT_SECONDS={os.getenv('STEP4_RETRY_WAIT_SECONDS', '')}, "
        f"RANDOMIZE_FINGERPRINT_ON_START={os.getenv('RANDOMIZE_FINGERPRINT_ON_START', '')}, "
        f"PROFILE_LIMIT={os.getenv('PROFILE_LIMIT', '')}, "
        f"PROFILE_IDS={'已设置' if os.getenv('PROFILE_IDS', '').strip() else '全部'}"
    )

def get_random_account_info():
    """Generate random account data: Gmail, full name, password, zipcode."""
    # List of realistic US zipcodes from different regions
    us_zips = [
        "10001", "90210", "60601", "77001", "85001", "19101", "33101", "30301", "98101", "02101",
        "75201", "20001", "48201", "55401", "63101", "80201", "84101", "94101", "97201", "21201"
    ]
    
    prefix_len = random.randint(8, 12)
    prefix = ''.join(random.choices(string.ascii_lowercase + string.digits, k=prefix_len))
    email = f"{prefix}@gmail.com"
    password = f"{prefix}pw"
    
    first_name = ''.join(random.choices(string.ascii_lowercase, k=5)).capitalize()
    last_name = ''.join(random.choices(string.ascii_lowercase, k=6)).capitalize()
    full_name = f"{first_name} {last_name}"
    
    zipcode = random.choice(us_zips)
    return email, full_name, password, zipcode

def build_password_from_email(email):
    """Build password as '<gmail_prefix>pw'."""
    raw = (email or "").strip()
    if "@" in raw:
        prefix = raw.split("@", 1)[0].strip()
    else:
        prefix = raw
    if not prefix:
        prefix = ''.join(random.choices(string.ascii_lowercase + string.digits, k=10))
    return f"{prefix}pw"

def get_visible_input_value(driver, xpaths):
    element = _find_visible_element(driver, xpaths)
    if element is None:
        return ""
    try:
        return (element.get_attribute("value") or "").strip()
    except Exception:
        return ""

def get_card_from_xml():
    """Pick one card entry from account.xml split by '---'."""
    file_path = ACCOUNT_FILE
    if not os.path.exists(file_path):
        print("account.xml not found, using default test card")
        return "4242424242424242", "12/25", "123"
    
    with open(file_path, 'r', encoding='utf-8') as f:
        lines = [line.strip() for line in f.readlines() if '---' in line]
        
    if not lines:
        return "4242424242424242", "12/25", "123"
        
    selected_line = random.choice(lines)
    parts = [p.strip() for p in selected_line.split('---')]
    if len(parts) < 3:
        return "4242424242424242", "12/25", "123"
    return parts[0], parts[1], parts[2]

def save_link_to_xml(link, output_file=LINK_FILE):
    """Append extracted link to link.xml in a thread-safe way."""
    if not link:
        print("Skip saving empty link")
        return False

    os.makedirs(os.path.dirname(output_file), exist_ok=True)
    with LINK_FILE_LOCK:
        # Ensure links are different within the same run.
        if link in RUN_LINKS:
            print(f"Duplicate link in current run, skip: {link}")
            return False

        existing = set()
        if os.path.exists(output_file):
            with open(output_file, "r", encoding="utf-8") as f:
                existing = {line.strip() for line in f if line.strip()}

        RUN_LINKS.add(link)
        if link not in existing:
            with open(output_file, "a", encoding="utf-8") as f:
                f.write(link + "\n")
            print(f"Saved link to {output_file}: {link}")
        else:
            print(f"Link already exists in {output_file}: {link}")
        return True

def _extract_offer_link_from_text(text):
    if not text:
        return ""
    # Match exact target format: https://one.google.com/offer/<CODE>
    match = re.search(r"https://one\.google\.com/offer/[A-Za-z0-9_-]+", text)
    return match.group(0) if match else ""

def _resolve_offer_via_http(url):
    """Try to resolve intermediary links to final one.google offer link."""
    if not url or not url.startswith("http"):
        return ""
    try:
        rsp = requests.get(url, timeout=20, allow_redirects=True)
        # Final redirected URL
        offer = _extract_offer_link_from_text(rsp.url or "")
        if offer:
            return offer
        # Redirect chain headers
        for hist in rsp.history or []:
            offer = _extract_offer_link_from_text(hist.headers.get("Location", ""))
            if offer:
                return offer
        # HTML body
        offer = _extract_offer_link_from_text(rsp.text or "")
        if offer:
            return offer
    except Exception:
        return ""
    return ""

def extract_final_link(driver, wait):
    """Extract final link, prioritizing one.google offer format."""
    handles = list(driver.window_handles or [])
    if not handles:
        handles = [driver.current_window_handle]

    generic_candidates = []

    # Scan each tab/window, prioritize newest tab first.
    for handle in reversed(handles):
        try:
            driver.switch_to.window(handle)
        except Exception:
            continue

        # Priority A: current URL if already one.google offer.
        current = (driver.current_url or "").strip()
        offer = _extract_offer_link_from_text(current)
        if offer:
            return offer
        if current.startswith("https://"):
            generic_candidates.append(current)

        # Priority B: anchor href that is one.google offer.
        try:
            anchors = driver.find_elements(By.XPATH, "//a[@href]")
        except Exception:
            anchors = []
        for anchor in anchors:
            href = (anchor.get_attribute("href") or "").strip()
            offer = _extract_offer_link_from_text(href)
            if offer:
                return offer
            if href.startswith("https://"):
                generic_candidates.append(href)

        # Priority C: page source regex.
        page = driver.page_source or ""
        offer = _extract_offer_link_from_text(page)
        if offer:
            return offer
        match = re.search(r"https://[^\s\"'<>]+", page)
        if match:
            generic_candidates.append(match.group(0))

    # Priority D: resolve generic candidates via HTTP redirects/body.
    seen = set()
    for candidate in generic_candidates:
        if not candidate or candidate in seen:
            continue
        seen.add(candidate)
        offer = _resolve_offer_via_http(candidate)
        if offer:
            return offer

    # Fallback: return first generic https link if offer format not found.
    for candidate in generic_candidates:
        if candidate.startswith("https://"):
            return candidate

    raise RuntimeError("No https link found on final page")

def _safe_filename(text):
    safe = re.sub(r"[^A-Za-z0-9_.-]+", "_", str(text)).strip("_")
    return safe or "unknown"

def save_debug_artifacts(driver, profile_id, step_name, error):
    """Save screenshot/html/meta when a step fails."""
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    profile_dir = os.path.join(DEBUG_DIR, _safe_filename(profile_id))
    os.makedirs(profile_dir, exist_ok=True)
    base = f"{ts}_{_safe_filename(step_name)}"

    screenshot_path = os.path.join(profile_dir, f"{base}.png")
    html_path = os.path.join(profile_dir, f"{base}.html")
    meta_path = os.path.join(profile_dir, f"{base}.txt")

    current_url = ""
    title = ""
    try:
        current_url = driver.current_url or ""
        title = driver.title or ""
    except Exception:
        pass

    try:
        driver.save_screenshot(screenshot_path)
    except Exception as e:
        screenshot_path = f"save_failed: {e}"

    try:
        page_source = driver.page_source or ""
        with open(html_path, "w", encoding="utf-8") as f:
            f.write(page_source)
    except Exception as e:
        html_path = f"save_failed: {e}"

    try:
        with open(meta_path, "w", encoding="utf-8") as f:
            f.write(f"profile_id: {profile_id}\n")
            f.write(f"step: {step_name}\n")
            f.write(f"time: {ts}\n")
            f.write(f"url: {current_url}\n")
            f.write(f"title: {title}\n")
            f.write(f"error: {error}\n")
            f.write("traceback:\n")
            f.write(traceback.format_exc())
    except Exception as e:
        meta_path = f"save_failed: {e}"

    print(f"[{profile_id}] debug screenshot: {screenshot_path}")
    print(f"[{profile_id}] debug html: {html_path}")
    print(f"[{profile_id}] debug meta: {meta_path}")

def _take_debug_screenshot(driver, profile_id, step_name):
    """Backward-compatible debug capture helper used by workflow steps."""
    try:
        save_debug_artifacts(driver, profile_id, step_name, f"manual debug capture: {step_name}")
    except Exception as e:
        print(f"[{profile_id}] debug capture failed at {step_name}: {e}")

def js_click(driver, element):
    """Force click by JavaScript to bypass overlay issues."""
    driver.execute_script("arguments[0].click();", element)

def _find_visible_element(driver, xpaths):
    for xp in xpaths:
        elements = driver.find_elements(By.XPATH, xp)
        for element in elements:
            try:
                if element.is_displayed():
                    return element
            except Exception:
                continue
    return None

def _wait_for_visible_element(driver, xpaths, timeout=20, raise_error=True):
    end = time.time() + timeout
    while time.time() < end:
        element = _find_visible_element(driver, xpaths)
        if element is not None:
            return element
        time.sleep(0.3)
    if raise_error:
        raise TimeoutException(f"No visible element found for xpaths: {xpaths}")
    return None

def _wait_for_clickable_element(driver, xpaths, timeout=20, raise_error=True):
    end = time.time() + timeout
    while time.time() < end:
        remaining = max(0.2, min(1.2, end - time.time()))
        for xp in xpaths:
            try:
                element = WebDriverWait(driver, remaining).until(
                    EC.element_to_be_clickable((By.XPATH, xp))
                )
                if element is not None:
                    return element
            except Exception:
                continue
        time.sleep(0.15)
    if raise_error:
        raise TimeoutException(f"No clickable element found for xpaths: {xpaths}")
    return None

def _click_first(driver, xpaths, timeout=20):
    element = _wait_for_clickable_element(driver, xpaths, timeout=timeout, raise_error=False)
    if element is None:
        element = _wait_for_visible_element(driver, xpaths, timeout=timeout)
    try:
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", element)
    except Exception:
        pass
    _pause_between_actions(driver=driver, note="click element", allow_scroll=False)
    js_click(driver, element)
    _pause_between_actions(driver=driver, note="after click", allow_scroll=True)
    return element

def _fill_input(driver, xpaths, value, timeout=20):
    element = _wait_for_clickable_element(driver, xpaths, timeout=timeout, raise_error=False)
    if element is None:
        element = _wait_for_visible_element(driver, xpaths, timeout=timeout)
    try:
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", element)
    except Exception:
        pass
    try:
        readonly = (element.get_attribute("readonly") or "").lower()
        disabled = (element.get_attribute("disabled") or "").lower()
        if readonly or disabled:
            return False
    except Exception:
        pass
    try:
        element.click()
    except Exception:
        pass
    try:
        element.clear()
    except Exception:
        pass
    # More robust clear for inputs that ignore clear().
    try:
        element.send_keys(Keys.CONTROL, "a")
        element.send_keys(Keys.BACKSPACE)
    except Exception:
        pass
    _pause_between_actions(driver=driver, note="fill input", allow_scroll=False)
    element.send_keys(value)
    _pause_between_actions(driver=driver, note="after input", allow_scroll=True)
    return True

def _human_click(driver, element):
    """Prefer human-like click sequence before JS fallbacks."""
    try:
        ActionChains(driver).move_to_element(element).pause(random.uniform(0.12, 0.35)).click().perform()
        return True
    except Exception:
        pass
    try:
        element.click()
        return True
    except Exception:
        return False

def _aggressive_click(driver, element):
    """Try multiple click methods for stubborn buttons."""
    try:
        element.click()
        return True
    except Exception:
        pass

    try:
        ActionChains(driver).move_to_element(element).pause(0.05).click().perform()
        return True
    except Exception:
        pass

    try:
        js_click(driver, element)
        return True
    except Exception:
        pass

    try:
        driver.execute_script(
            "arguments[0].dispatchEvent(new MouseEvent('mousedown', {bubbles:true}));"
            "arguments[0].dispatchEvent(new MouseEvent('mouseup', {bubbles:true}));"
            "arguments[0].dispatchEvent(new MouseEvent('click', {bubbles:true}));",
            element,
        )
        return True
    except Exception:
        pass

    return False

def _submit_signup_form(driver):
    """
    Submit signup form via form APIs as a fallback when click is unreliable.
    Returns True if submit signal was sent.
    """
    try:
        submitted = driver.execute_script(
            """
            const btn =
              document.querySelector("button[data-e2e='signup-form-submit-button']") ||
              document.querySelector("button[data-catchpoint='signup-form-submit-button']") ||
              document.querySelector("button[type='submit']");
            if (!btn) return false;
            const form = btn.closest("form");
            if (form) {
              if (typeof form.requestSubmit === "function") {
                form.requestSubmit(btn);
                return true;
              }
              const evt = new Event("submit", { bubbles: true, cancelable: true });
              form.dispatchEvent(evt);
              try { form.submit(); } catch (e) {}
              return true;
            }
            btn.click();
            return true;
            """
        )
        return bool(submitted)
    except Exception:
        return False

def _signup_join_button_xpaths():
    """Exact selectors for signup modal submit button; avoid top-nav Join button."""
    return [
        "//form[@name='signup']//button[@data-e2e='signup-form-submit-button']",
        "//form[@name='signup']//button[@data-catchpoint='signup-form-submit-button']",
        "//div[@role='dialog']//button[@data-e2e='signup-form-submit-button']",
        "//div[@role='dialog']//button[@data-catchpoint='signup-form-submit-button']",
        "//button[@data-e2e='signup-form-submit-button']",
        "//button[@data-catchpoint='signup-form-submit-button']",
        "//form[@name='signup']//button[@type='submit' and contains(normalize-space(.), 'Join for Free')]",
        "//div[@role='dialog']//button[@type='submit' and contains(normalize-space(.), 'Join for Free')]",
    ]

def _get_signup_submit_button(driver):
    """
    Resolve the signup submit button from signup form/dialog scope first.
    Returns WebElement or None.
    """
    try:
        element = driver.execute_script(
            """
            const root =
              document.querySelector("form[name='signup']") ||
              document.querySelector("div[role='dialog']");
            const selectors = [
              "button[data-e2e='signup-form-submit-button']",
              "button[data-catchpoint='signup-form-submit-button']",
              "button[type='submit']",
            ];
            const pools = [];
            if (root) pools.push(...selectors.map(s => [root, s]));
            pools.push(...selectors.map(s => [document, s]));
            for (const [scope, sel] of pools) {
              const list = scope.querySelectorAll(sel);
              for (const btn of list) {
                const txt = (btn.innerText || btn.textContent || "").trim().toLowerCase();
                if (txt.includes("join for free") || btn.dataset?.e2e === "signup-form-submit-button" || btn.dataset?.catchpoint === "signup-form-submit-button") {
                  return btn;
                }
              }
            }
            return null;
            """
        )
        return element
    except Exception:
        return None

def _is_element_enabled(element):
    try:
        disabled_attr = (element.get_attribute("disabled") or "").lower()
        aria_disabled = (element.get_attribute("aria-disabled") or "").lower()
        return not disabled_attr and aria_disabled != "true"
    except Exception:
        return False

def _find_module2_redeem_item(driver):
    """
    Locate Redeem target under Module 2/Week 2 area.
    Preference:
    1) clickable item whose text starts with 'Redeem'
    2) module-local third clickable item if its text starts with 'Redeem'
    """
    try:
        element = driver.execute_script(
            """
            const norm = (s) => (s || "").replace(/\\s+/g, " ").trim().toLowerCase();
            const isVisible = (el) => {
              if (!el) return false;
              const r = el.getBoundingClientRect();
              return r.width > 0 && r.height > 0;
            };
            const uniq = (arr) => {
              const out = [];
              for (const el of arr) if (el && !out.includes(el)) out.push(el);
              return out;
            };
            const pickClickable = (el) => {
              if (!el) return null;
              if (el.matches && el.matches("a,button,[role='button']")) return el;
              return el.closest ? el.closest("a,button,[role='button']") : null;
            };

            // Exact target from course outline: ungraded LTI redeem item.
            const directRedeem = document.querySelector(
              "a[href*='/ungradedLti/'][aria-label*='Redeem your Google AI Pro trial']"
            ) || document.querySelector(
              "a[href*='/ungradedLti/'][href*='google-ai-fundamentals'][aria-label*='Redeem']"
            ) || document.querySelector(
              "a[href*='/ungradedLti/'][aria-label*='Ungraded App Item'][aria-label*='Redeem']"
            );
            if (directRedeem && isVisible(directRedeem)) return directRedeem;

            const headers = Array.from(
              document.querySelectorAll("button,summary,[role='button'],a,div,h2,h3,h4,span")
            ).filter((el) => {
              const t = norm(el.innerText || el.textContent);
              return isVisible(el) && (t.includes("module 2") || t.includes("week 2") || t.includes("practice using ai"));
            });
            if (!headers.length) return null;

            for (const header of headers) {
              const containers = uniq([
                header.closest("section"),
                header.closest("article"),
                header.closest("li"),
                header.closest("div"),
                header.parentElement,
              ]);
              for (const container of containers) {
                if (!container) continue;
                const clickables = Array.from(
                  container.querySelectorAll("a,button,[role='button']")
                ).filter((el) => {
                  if (!isVisible(el)) return false;
                  const t = norm(el.innerText || el.textContent);
                  if (!t) return false;
                  if (t.includes("module 2") || t.includes("week 2")) return false;
                  return true;
                });
                if (!clickables.length) continue;

                const redeem = clickables.find((el) => norm(el.innerText || el.textContent).startsWith("redeem"));
                if (redeem) return redeem;

                const redeemByLabel = clickables.find((el) => {
                  const label = norm(el.getAttribute ? el.getAttribute("aria-label") : "");
                  const href = norm(el.getAttribute ? el.getAttribute("href") : "");
                  return label.includes("redeem your google ai pro trial") || (href.includes("/ungradedlti/") && label.includes("redeem"));
                });
                if (redeemByLabel) return redeemByLabel;

                if (clickables.length >= 3) {
                  const third = clickables[2];
                  const thirdText = norm(third.innerText || third.textContent);
                  if (thirdText.startsWith("redeem")) return third;
                }
              }
            }
            return null;
            """
        )
        return element
    except Exception:
        return None

def _dismiss_common_popups(driver, allow_modal_close=False):
    """
    Best-effort close of blocking popups.
    By default, avoid closing auth/signup modal used in business flow.
    """
    popup_xpaths = [
        "//button[@id='onetrust-accept-btn-handler']",
        "//button[contains(., 'Accept All')]",
        "//button[contains(., 'Accept')]",
    ]
    if allow_modal_close:
        popup_xpaths = [
            "//button[@aria-label='Close modal']",
            "//button[contains(@aria-label, 'Close')]",
            "//button[contains(@class, 'close')]",
            "//button[contains(@data-e2e, 'close')]",
            "//*[@role='dialog']//button[.='×' or .='✕' or .='✖']",
        ] + popup_xpaths
    for xp in popup_xpaths:
        try:
            elements = driver.find_elements(By.XPATH, xp)
            for element in elements:
                try:
                    if element.is_displayed():
                        js_click(driver, element)
                        time.sleep(0.2)
                        break
                except Exception:
                    continue
        except Exception:
            continue

def _is_signup_page(driver):
    try:
        url = (driver.current_url or "").lower()
    except Exception:
        url = ""
    if "authmode=signup" in url:
        return True
    signup_markers = [
        "//form[@name='signup']",
        "//div[@role='dialog']//button[@data-e2e='signup-form-submit-button']",
        "//div[@role='dialog']//button[@data-catchpoint='signup-form-submit-button']",
        "//div[@role='dialog']//input[@name='password' and @type='password']",
    ]
    return _find_visible_element(driver, signup_markers) is not None

def _is_post_signup_ready(driver):
    """Whether signup finished and flow moved to terms/trial/billing pages."""
    post_signup_markers = [
        "//button[contains(., 'I accept')]",
        "//button[contains(., 'Start Free Trial')]",
        "//button[contains(., 'Start free trial')]",
        "//button[contains(., 'Subscribe')]",
        "//select[contains(@name, 'country') or contains(@id, 'country')]",
        "//input[contains(@name, 'zip') or contains(@id, 'postal')]",
    ]
    if _find_visible_element(driver, post_signup_markers) is not None:
        return True
    try:
        url = (driver.current_url or "").lower()
    except Exception:
        url = ""
    return any(k in url for k in ("checkout", "payment", "subscribe", "billing"))

def _wait_after_signup(driver, profile_id, auto_retry_seconds, manual_wait_seconds):
    """Conservative retry after signup submit, then optional manual verification window."""
    join_btn_xpaths = _signup_join_button_xpaths()
    try:
        auto_click_attempts = int(os.getenv("SIGNUP_AUTO_CLICK_ATTEMPTS", "1"))
    except Exception:
        auto_click_attempts = 1
    auto_click_attempts = max(0, min(auto_click_attempts, 3))

    auto_deadline = time.time() + max(0, auto_retry_seconds)
    auto_click_done = 0
    next_click_at = time.time()
    while time.time() < auto_deadline:
        _dismiss_common_popups(driver)
        if _is_post_signup_ready(driver) or not _is_signup_page(driver):
            return True

        if auto_click_done < auto_click_attempts and time.time() >= next_click_at:
            try:
                join_btn = _get_signup_submit_button(driver)
                if join_btn is None:
                    join_btn = _wait_for_clickable_element(driver, join_btn_xpaths, timeout=2.5, raise_error=False)
                if join_btn is not None:
                    if _is_element_enabled(join_btn):
                        _pause_between_actions(
                            driver=driver,
                            profile_id=profile_id,
                            note=f"auto retry Join click {auto_click_done + 1}",
                            allow_scroll=True,
                        )
                        if _human_click(driver, join_btn):
                            auto_click_done += 1
                            next_click_at = time.time() + random.uniform(5.0, 8.5)
            except Exception:
                pass
        time.sleep(1.1)

    manual_wait_seconds = max(0, manual_wait_seconds)
    if manual_wait_seconds == 0:
        return False

    print(
        f"[{profile_id}] Signup verification detected. "
        f"Please complete captcha/verification manually within {manual_wait_seconds}s..."
    )
    manual_deadline = time.time() + manual_wait_seconds
    while time.time() < manual_deadline:
        _dismiss_common_popups(driver)
        if _is_post_signup_ready(driver) or not _is_signup_page(driver):
            return True
        time.sleep(2)
    return False

def _fill_stripe_fields(driver, card_num, exp_date, cvc, timeout=30):
    """Fill Stripe card fields across possible multi-iframe layouts."""
    # Build robust field variants
    targets = [
        ("cardnumber", card_num, [
            "//input[@name='cardnumber']",
            "//input[@aria-label='Card number' or @aria-label='Card Number']",
            "//input[contains(@placeholder, 'Card number') or contains(@placeholder, 'Card Number')]",
            "//input[@autocomplete='cc-number']",
        ]),
        ("exp-date", exp_date, [
            "//input[@name='exp-date']",
            "//input[@aria-label='Expiration date' or @aria-label='Expiration Date']",
            "//input[contains(@placeholder, 'MM / YY')]",
            "//input[@autocomplete='cc-exp']",
        ]),
        ("cvc", cvc, [
            "//input[@name='cvc']",
            "//input[@aria-label='CVC' or @aria-label='Security code']",
            "//input[contains(@placeholder, 'CVC')]",
            "//input[@autocomplete='cc-csc']",
        ])
    ]
    
    for field_id, value, xpaths in targets:
        found = False
        deadline = time.time() + timeout
        while time.time() < deadline and not found:
            # Try current DOM first.
            try:
                element = _find_visible_element(driver, xpaths)
                if element:
                    element.send_keys(value)
                    found = True
                    break
            except Exception:
                pass

            if found:
                break

            # Try every iframe.
            try:
                frames = driver.find_elements(By.TAG_NAME, "iframe")
            except Exception:
                frames = []
            
            for frame in frames:
                try:
                    # Skip invisible small frames that aren't payment related
                    width = int(frame.get_attribute("width") or 0)
                    if width > 0 and width < 10: continue
                    
                    driver.switch_to.frame(frame)
                    element = _find_visible_element(driver, xpaths)
                    if element:
                        element.send_keys(value)
                        found = True
                        break
                except Exception:
                    pass
                finally:
                    driver.switch_to.default_content()
                if found:
                    break
            
            if not found:
                time.sleep(1)
                
        if not found:
            raise TimeoutException(f"Stripe field not found: {field_id}")

def _is_logged_out_page(driver):
    # First check if we see clear logged-in indicators.
    if _is_logged_in_page(driver):
        return False

    logout_markers = [
        "//a[contains(., 'Log in')]",
        "//button[contains(., 'Log in')]",
        "//a[contains(., 'Log In')]",
        "//button[contains(., 'Log In')]",
        "//a[contains(., 'Join for Free')]",
        "//button[contains(., 'Join for Free')]",
        "//button[contains(., 'Enroll for free')]",
        "//a[contains(., 'Enroll for free')]",
    ]
    if _find_visible_element(driver, logout_markers) is not None:
        return True
    try:
        url = (driver.current_url or "").lower()
    except Exception:
        url = ""
    return "authmode=login" in url or "authmode=signup" in url

def _is_logged_in_page(driver):
    logged_in_xpaths = [
        # Account/profile triggers in top nav
        "//button[contains(@aria-label, 'Account')]",
        "//button[contains(@aria-label, 'account')]",
        "//button[contains(@aria-label, 'Profile')]",
        "//button[contains(@aria-label, 'profile')]",
        "//button[contains(@aria-label, 'User dropdown')]",
        "//*[@data-e2e='user-dropdown-trigger']",
        # Avatar button in navigation
        "//button[contains(@class, 'cds-button') and .//*[contains(@class, 'Avatar')]]",
    ]
    return _find_visible_element(driver, logged_in_xpaths) is not None

def _wait_until(condition_fn, timeout=3.0, interval=0.25):
    deadline = time.time() + max(0.0, float(timeout))
    while time.time() < deadline:
        try:
            if condition_fn():
                return True
        except Exception:
            pass
        time.sleep(interval)
    try:
        return bool(condition_fn())
    except Exception:
        return False

def _reset_to_single_window(driver):
    """Keep one browser window/tab and close all others."""
    handles = list(driver.window_handles or [])
    if not handles:
        return
    primary = handles[0]
    try:
        driver.switch_to.window(primary)
    except Exception:
        pass

def _clear_all_tabs_and_open_target(driver, target_url):
    """
    Clear existing tabs by opening a fresh tab, closing old tabs, then opening target URL.
    """
    handles = list(driver.window_handles or [])
    if not handles:
        return _open_target_page_fast(driver, target_url, timeout=5.0)

    fresh_handle = None
    try:
        driver.switch_to.window(handles[0])
        driver.execute_script("window.open('about:blank', '_blank');")
        new_handles = list(driver.window_handles or [])
        for handle in new_handles:
            if handle not in handles:
                fresh_handle = handle
                break
        if fresh_handle is None and new_handles:
            fresh_handle = new_handles[-1]
    except Exception:
        fresh_handle = handles[0]

    for handle in handles:
        try:
            driver.switch_to.window(handle)
            driver.close()
        except Exception:
            continue

    try:
        remain = list(driver.window_handles or [])
        if not remain:
            return False
        if fresh_handle in remain:
            driver.switch_to.window(fresh_handle)
        else:
            driver.switch_to.window(remain[0])
    except Exception:
        return False

    return _open_target_page_fast(driver, target_url, timeout=5.0)

def _open_target_page_fast(driver, target_url, timeout=4.0):
    try:
        driver.get(target_url)
    except Exception:
        return False
    return _wait_until(
        lambda: "coursera.org" in (driver.current_url or "").lower(),
        timeout=timeout,
        interval=0.2,
    )

def logout_and_go_home(driver, profile_id):
    """Logout current Coursera account and return to target page without login state."""
    target_url = COURSE_ENROLL_URL
    logout_urls = (
        "https://www.coursera.org/?authMode=logout",
        "https://www.coursera.org/logout",
    )
    logout_xpaths = [
        "//a[contains(., 'Log Out') or contains(., 'Log out') or contains(., 'Sign out')]",
        "//button[contains(., 'Log Out') or contains(., 'Log out') or contains(., 'Sign out')]",
    ]
    # User dropdown button — must be the avatar button, NOT the notification bell
    user_dropdown_xpaths = [
        "//button[contains(@aria-label, 'dropdown')]",
        "//button[contains(@aria-label, 'User dropdown')]",
        "//button[contains(@aria-label, 'Account')]",
        "//button[contains(@aria-label, 'account')]",
        "//button[contains(@aria-label, 'Profile')]",
        "//button[contains(@aria-label, 'profile')]",
        "//*[@data-e2e='user-dropdown-trigger']",
    ]

    print(f"[{profile_id}] Logout: clearing login state")

    # Strategy 1: direct logout URL first (fast path).
    for url in logout_urls:
        try:
            driver.get(url)
        except Exception:
            continue
        if _wait_until(lambda: _is_logged_out_page(driver), timeout=3.0, interval=0.25):
            print(f"[{profile_id}] Logout via URL success: {url}")
            break

    # Strategy 2: UI dropdown logout if still logged in.
    if _is_logged_in_page(driver):
        _open_target_page_fast(driver, target_url, timeout=3.0)
        for attempt in range(2):
            _dismiss_common_popups(driver)
            if _is_logged_out_page(driver):
                break
            try:
                account_btn = _find_visible_element(driver, user_dropdown_xpaths)
                if account_btn is not None:
                    js_click(driver, account_btn)
                    _wait_until(lambda: _find_visible_element(driver, logout_xpaths) is not None, timeout=1.2, interval=0.2)
            except Exception:
                pass
            try:
                logout_btn = _find_visible_element(driver, logout_xpaths)
                if logout_btn is not None:
                    js_click(driver, logout_btn)
                    if _wait_until(lambda: _is_logged_out_page(driver), timeout=3.0, interval=0.25):
                        print(f"[{profile_id}] Logout via UI dropdown success (attempt {attempt + 1})")
                        break
            except Exception:
                pass

    # Strategy 3: cookie clear fallback.
    if _is_logged_in_page(driver):
        try:
            driver.delete_all_cookies()
            driver.get(logout_urls[0])
            if _wait_until(lambda: _is_logged_out_page(driver), timeout=2.5, interval=0.25):
                print(f"[{profile_id}] Logout via cookie clear success")
        except Exception:
            pass

    # Final: ensure target page is loaded for next step.
    for _ in range(2):
        _open_target_page_fast(driver, target_url, timeout=3.0)
        _dismiss_common_popups(driver, allow_modal_close=True)
        if _wait_until(lambda: _is_logged_out_page(driver), timeout=2.0, interval=0.25):
            print(f"[{profile_id}] Logout confirmed, back to {target_url}")
            return
        if _is_logged_in_page(driver):
            try:
                driver.get(logout_urls[0])
                _wait_until(lambda: _is_logged_out_page(driver), timeout=2.0, interval=0.25)
            except Exception:
                pass

    print(f"[{profile_id}] Logout state not fully confirmed on {target_url}")

def normalize_entry_page(driver, profile_id):
    """
    On entry:
    1) clear all existing tabs and open target URL
    2) default go to Step 2 after cookie/storage clear
    3) only when still logged in, logout then go to Step 2
    """
    print(f"[{profile_id}] Entry: reset tabs, open target page, and normalize login state")
    demo_pause(profile_id, "open entry page")
    _clear_all_tabs_and_open_target(driver, COURSE_ENROLL_URL)
    _dismiss_common_popups(driver, allow_modal_close=True)

    if _is_logged_in_page(driver):
        print(f"[{profile_id}] Entry: logged-in account detected, logging out first")
        logout_and_go_home(driver, profile_id)
        _clear_all_tabs_and_open_target(driver, COURSE_ENROLL_URL)
        _dismiss_common_popups(driver, allow_modal_close=True)
        print(f"[{profile_id}] Entry ready: logout flow done, go to Step 2")
        return

    print(f"[{profile_id}] Entry ready: proceed directly to Step 2")

# ----------------- Browser Control API -----------------
def start_roxy_browser(profile_id):
    """Start RoxyBrowser environment."""
    workspace_id = get_workspace_id()
    if not workspace_id:
        print(f"[{profile_id}] missing workspaceId")
        return None, None
    try:
        response = _api_post("/browser/open", {"workspaceId": workspace_id, "dirId": profile_id})
        if response.get("code") == 0 or response.get("success"):
            data = response.get("data", {})
            debug_addr = data.get("http") or data.get("debug_port")
            webdriver_path = data.get("driver") or data.get("webdriver")
            return debug_addr, webdriver_path
        print(f"[{profile_id}] open failed: code={response.get('code')}, msg={response.get('msg')}")
    except Exception as e:
        print(f"[{profile_id}] Browser start error: {e}")
    return None, None

def close_roxy_browser(profile_id):
    """Close RoxyBrowser environment."""
    try:
        workspace_id = get_workspace_id()
        payload = {"dirId": profile_id}
        if workspace_id:
            payload["workspaceId"] = workspace_id
        _api_post("/browser/close", payload)
    except Exception:
        pass

def randomize_fingerprints_for_profiles(profile_ids):
    """
    Randomize fingerprint for each profile via RoxyBrowser API.
    API: POST /browser/random_env
    """
    workspace_id = get_workspace_id()
    if not workspace_id:
        print("[Fingerprint] skip random_env: workspaceId missing")
        return False
    if not profile_ids:
        print("[Fingerprint] skip random_env: no profiles")
        return False

    success = 0
    failed = 0
    for profile_id in profile_ids:
        try:
            rsp = _api_post("/browser/random_env", {"workspaceId": workspace_id, "dirId": profile_id})
            if rsp.get("code") == 0 or rsp.get("success"):
                success += 1
            else:
                failed += 1
                print(
                    f"[Fingerprint] {profile_id} random_env failed: "
                    f"code={rsp.get('code')}, msg={rsp.get('msg')}"
                )
        except Exception as e:
            failed += 1
            print(f"[Fingerprint] {profile_id} random_env error: {e}")

    print(f"[Fingerprint] random_env done: success={success}, failed={failed}, total={len(profile_ids)}")
    return failed == 0

def randomize_fingerprint_for_profile(profile_id):
    """Randomize fingerprint for one profile right before opening its window."""
    workspace_id = get_workspace_id()
    if not workspace_id:
        print(f"[{profile_id}] Fingerprint skip: workspaceId missing")
        return False
    try:
        rsp = _api_post("/browser/random_env", {"workspaceId": workspace_id, "dirId": profile_id})
        if rsp.get("code") == 0 or rsp.get("success"):
            print(f"[{profile_id}] Fingerprint random_env success")
            return True
        print(
            f"[{profile_id}] Fingerprint random_env failed: "
            f"code={rsp.get('code')}, msg={rsp.get('msg')}"
        )
    except Exception as e:
        print(f"[{profile_id}] Fingerprint random_env error: {e}")
    return False

def clear_cookie_storage_in_session(driver, profile_id):
    """Clear cookie/storage in the same webdriver session after window opens."""
    try:
        _open_target_page_fast(driver, "https://www.coursera.org/", timeout=4.0)
    except Exception:
        pass
    try:
        driver.delete_all_cookies()
    except Exception:
        pass
    try:
        driver.execute_script("window.localStorage.clear(); window.sessionStorage.clear();")
    except Exception:
        pass
    try:
        driver.execute_cdp_cmd("Network.enable", {})
        driver.execute_cdp_cmd("Network.clearBrowserCookies", {})
        driver.execute_cdp_cmd("Network.clearBrowserCache", {})
    except Exception:
        pass
    print(f"[{profile_id}] In-session cookie/storage clear done")

def _env_flag(name, default=False):
    raw = os.getenv(name, "1" if default else "0").strip().lower()
    return raw in {"1", "true", "yes", "on"}

def _parse_positive_int(raw, default):
    try:
        value = int(str(raw).strip())
        if value <= 0:
            raise ValueError
        return value
    except Exception:
        return default

def _ensure_desktop_viewport(driver, profile_id="", reason=""):
    """
    Force desktop-like viewport to reduce responsive layout drift.
    Env:
      WINDOW_WIDTH (default 1600)
      WINDOW_HEIGHT (default 980)
      MIN_VIEWPORT_WIDTH (default 1180)
    """
    target_w = _parse_positive_int(os.getenv("WINDOW_WIDTH", "1600"), 1600)
    target_h = _parse_positive_int(os.getenv("WINDOW_HEIGHT", "980"), 980)
    min_inner_w = _parse_positive_int(os.getenv("MIN_VIEWPORT_WIDTH", "1180"), 1180)
    reason_tag = f" ({reason})" if reason else ""

    try:
        driver.set_window_size(target_w, target_h)
        time.sleep(0.2)
    except Exception:
        try:
            driver.maximize_window()
            time.sleep(0.2)
        except Exception:
            pass

    metrics = {}
    try:
        metrics = driver.execute_script(
            "return {innerW: window.innerWidth || 0, innerH: window.innerHeight || 0,"
            "outerW: window.outerWidth || 0, outerH: window.outerHeight || 0};"
        ) or {}
    except Exception:
        metrics = {}

    inner_w = int(metrics.get("innerW") or 0)
    if inner_w > 0 and inner_w < min_inner_w:
        try:
            grow_w = max(target_w, min_inner_w + 240)
            driver.set_window_size(grow_w, target_h)
            time.sleep(0.2)
            metrics = driver.execute_script(
                "return {innerW: window.innerWidth || 0, innerH: window.innerHeight || 0,"
                "outerW: window.outerWidth || 0, outerH: window.outerHeight || 0};"
            ) or metrics
        except Exception:
            pass

    try:
        print(
            f"[{profile_id}] Viewport{reason_tag}: "
            f"outer={int(metrics.get('outerW') or 0)}x{int(metrics.get('outerH') or 0)}, "
            f"inner={int(metrics.get('innerW') or 0)}x{int(metrics.get('innerH') or 0)}, "
            f"target={target_w}x{target_h}, min_inner_w={min_inner_w}"
        )
    except Exception:
        pass

# ----------------- Core Workflow -----------------
def run_coursera_workflow(driver, profile_id):
    wait = WebDriverWait(driver, 20)
    email, full_name, password, zipcode = get_random_account_info()
    card_num, exp_date, cvc = get_card_from_xml()
    step_name = "init"
    signup_retry_seconds = int(os.getenv("SIGNUP_RETRY_SECONDS", "45"))
    manual_verify_wait_seconds = int(os.getenv("MANUAL_VERIFY_WAIT_SECONDS", "180"))
    terms_wait_seconds = int(os.getenv("TERMS_WAIT_SECONDS", "60"))
    manual_join_try_seconds = int(os.getenv("MANUAL_JOIN_TRY_SECONDS", "0"))
    try:
        step3_extra_wait_seconds = max(0.0, float(os.getenv("STEP3_EXTRA_WAIT_SECONDS", "1.2")))
    except Exception:
        step3_extra_wait_seconds = 1.2
    try:
        step4_retry_wait_seconds = max(0.0, float(os.getenv("STEP4_RETRY_WAIT_SECONDS", "1.6")))
    except Exception:
        step4_retry_wait_seconds = 1.6
    try:
        step4_submit_attempts = int(os.getenv("STEP4_SUBMIT_ATTEMPTS", "2"))
    except Exception:
        step4_submit_attempts = 2
    step4_submit_attempts = max(1, min(step4_submit_attempts, 4))
    try:
        step4_post_submit_wait_seconds = max(1.0, float(os.getenv("STEP4_POST_SUBMIT_WAIT_SECONDS", "6.5")))
    except Exception:
        step4_post_submit_wait_seconds = 6.5
    step4_enable_form_submit = _env_flag("STEP4_ENABLE_FORM_SUBMIT", default=False)
    step4_aggressive_click_fallback = _env_flag("STEP4_AGGRESSIVE_CLICK_FALLBACK", default=False)

    try:
        skip_payment = False
        enroll_btn_xpaths = [
            "//button[contains(normalize-space(.), 'Enroll for free')]",
            "//a[contains(normalize-space(.), 'Enroll for free')]",
            "//button[contains(normalize-space(.), 'Enroll now')]",
            "//a[contains(normalize-space(.), 'Enroll now')]",
            "//button[contains(normalize-space(.), 'Enroll')]",
            "//a[contains(normalize-space(.), 'Enroll')]",
        ]
        email_input_xpaths = [
            "//input[@type='email' and not(@readonly)]",
            "//input[@name='email' and not(@readonly)]",
        ]
        email_any_xpaths = [
            "//input[@type='email']",
            "//input[@name='email']",
        ]
        continue_btn_xpaths = [
            "//button[contains(., 'Continue')]",
        ]
        name_input_xpaths = [
            "//input[@placeholder='Enter your full name' or @name='name']",
            "//input[@aria-label='Full Name']",
        ]
        password_input_xpaths = [
            "//input[@type='password' and (@name='password' or @autocomplete='new-password' or @aria-label='Password')]",
            "//input[@type='password']",
        ]
        join_btn_xpaths = _signup_join_button_xpaths()
        accept_btn_xpaths = [
            "//button[contains(., 'I accept')]",
            "//button[contains(., 'I Accept')]",
            "//button[contains(., 'Accept')]",
        ]

        step_name = "step_2_enroll"
        print(f"[{profile_id}] Step 2: Click Enroll for free")
        demo_pause(profile_id, "prepare Step 2")
        _dismiss_common_popups(driver)
        try:
            _click_first(driver, enroll_btn_xpaths, timeout=10)
        except Exception:
            # Some sessions open directly in signup/auth mode without this button.
            pass

        step_name = "step_3_email_continue"
        print(f"[{profile_id}] Step 3: Input email")
        print(f"[{profile_id}] Generated email: {email}")
        print(f"[{profile_id}] Step 3 extra wait: {step3_extra_wait_seconds:.1f}s")
        demo_pause(profile_id, "prepare Step 3")
        _dismiss_common_popups(driver)
        # Step 3 must transition from email page to full-name/password page.
        transitioned_to_step4 = False
        for attempt in range(1, 4):
            # If already on full-name/password page, no need to continue.
            if _find_visible_element(driver, name_input_xpaths + password_input_xpaths) is not None:
                transitioned_to_step4 = True
                break

            try:
                _fill_input(driver, email_input_xpaths, email, timeout=8)
            except Exception:
                pass
            if step3_extra_wait_seconds > 0:
                time.sleep(step3_extra_wait_seconds)

            entered_email = get_visible_input_value(driver, email_any_xpaths)
            if entered_email and entered_email.lower() != email.lower():
                print(
                    f"[{profile_id}] Step 3 warning: entered email mismatch "
                    f"(try {attempt}): {entered_email}"
                )
                try:
                    _fill_input(driver, email_input_xpaths, email, timeout=5)
                    entered_email = get_visible_input_value(driver, email_any_xpaths)
                except Exception:
                    pass
                if step3_extra_wait_seconds > 0:
                    time.sleep(step3_extra_wait_seconds)

            try:
                _click_first(driver, continue_btn_xpaths, timeout=8)
            except Exception:
                pass
            if step3_extra_wait_seconds > 0:
                time.sleep(step3_extra_wait_seconds)

            end_wait = time.time() + max(6.0, 6.0 + step3_extra_wait_seconds * 2.0)
            while time.time() < end_wait:
                if _find_visible_element(driver, name_input_xpaths + password_input_xpaths) is not None:
                    transitioned_to_step4 = True
                    break
                time.sleep(0.3)
            if transitioned_to_step4:
                break
            if step3_extra_wait_seconds > 0:
                time.sleep(step3_extra_wait_seconds)

        if not transitioned_to_step4 and _find_visible_element(driver, name_input_xpaths + password_input_xpaths) is None:
            raise RuntimeError("Step 3 did not reach full-name/password page after Continue.")

        effective_email = get_visible_input_value(driver, email_any_xpaths) or email
        password = build_password_from_email(effective_email)
        print(f"[{profile_id}] Step 4 password rule email: {effective_email}")
        print(f"[{profile_id}] Step 4 password generated: {password}")
        print(f"[{profile_id}] Step 4 retry wait: {step4_retry_wait_seconds:.1f}s")
        print(
            f"[{profile_id}] Step 4 strategy: attempts={step4_submit_attempts}, "
            f"post_submit_wait={step4_post_submit_wait_seconds:.1f}s, "
            f"form_submit={'on' if step4_enable_form_submit else 'off'}, "
            f"aggressive_click={'on' if step4_aggressive_click_fallback else 'off'}"
        )

        step_name = "step_4_name_password"
        print(f"[{profile_id}] Step 4: Fill name and password")
        demo_pause(profile_id, "prepare Step 4")
        _dismiss_common_popups(driver)
        _pause_between_actions(
            driver=driver,
            profile_id=profile_id,
            note="Step 4 pre-fill stabilization",
            allow_scroll=True,
        )
        signup_form_present = _find_visible_element(driver, name_input_xpaths + password_input_xpaths) is not None
        if signup_form_present:
            try:
                _wait_for_clickable_element(driver, name_input_xpaths, timeout=8, raise_error=False)
                _wait_for_clickable_element(driver, password_input_xpaths, timeout=8, raise_error=False)
            except Exception:
                pass

            try:
                _fill_input(driver, name_input_xpaths, full_name, timeout=8)
            except Exception:
                pass
            _wait_until(
                lambda: bool(get_visible_input_value(driver, name_input_xpaths)),
                timeout=4.0,
                interval=0.2,
            )
            _pause_between_actions(
                driver=driver,
                profile_id=profile_id,
                note="Step 4 after name input",
                allow_scroll=True,
            )
            try:
                _fill_input(driver, password_input_xpaths, password, timeout=8)
            except Exception:
                pass
            _wait_until(
                lambda: bool(get_visible_input_value(driver, password_input_xpaths)),
                timeout=4.0,
                interval=0.2,
            )
            _pause_between_actions(
                driver=driver,
                profile_id=profile_id,
                note="Step 4 after password input",
                allow_scroll=True,
            )

            # Optional: pause at Step 4 so user can manually click "Join for Free".
            if manual_join_try_seconds > 0:
                print(
                    f"[{profile_id}] Step 4 manual window: "
                    f"you can click Join for Free in {manual_join_try_seconds}s"
                )
                manual_deadline = time.time() + manual_join_try_seconds
                while time.time() < manual_deadline:
                    if _find_visible_element(driver, accept_btn_xpaths) is not None:
                        break
                    if not _is_signup_page(driver):
                        break
                    _pause_between_actions(
                        driver=driver,
                        profile_id=profile_id,
                        note="Step 4 manual-join wait tick",
                        allow_scroll=True,
                    )
                    time.sleep(random.uniform(0.5, 1.1))
                if _find_visible_element(driver, accept_btn_xpaths) is not None or not _is_signup_page(driver):
                    print(f"[{profile_id}] Step 4 manual click appears successful, continue workflow")
                else:
                    print(f"[{profile_id}] Step 4 manual window ended, fallback to auto Join click")

            # Step 4 must end with submitting "Join for Free".
            print(f"[{profile_id}] Step 4: Click Join for Free")
            join_clicked = False
            join_selectors = list(join_btn_xpaths)
            if _find_visible_element(driver, accept_btn_xpaths) is not None or not _is_signup_page(driver):
                join_clicked = True
            else:
                for attempt in range(1, step4_submit_attempts + 1):
                    click_sent = False
                    _pause_between_actions(
                        driver=driver,
                        profile_id=profile_id,
                        note=f"Step 4 attempt {attempt} pre-submit",
                        allow_scroll=True,
                    )

                    # Priority 1: real user-like click on Join button.
                    if not click_sent:
                        try:
                            join_btn = _get_signup_submit_button(driver)
                            if join_btn is None:
                                join_btn = _wait_for_clickable_element(
                                    driver,
                                    join_selectors,
                                    timeout=6,
                                    raise_error=False,
                                )
                            if join_btn is None:
                                join_btn = _wait_for_visible_element(driver, join_selectors, timeout=6)
                            try:
                                driver.execute_script("arguments[0].scrollIntoView({block:'center'});", join_btn)
                            except Exception:
                                pass
                            try:
                                if not _is_element_enabled(join_btn):
                                    print(f"[{profile_id}] Step 4: Join button disabled (attempt {attempt})")
                                else:
                                    _pause_between_actions(
                                        driver=driver,
                                        profile_id=profile_id,
                                        note=f"click Join for Free attempt {attempt}",
                                        allow_scroll=False,
                                    )
                                    click_sent = _human_click(driver, join_btn)
                                    if not click_sent and step4_aggressive_click_fallback:
                                        click_sent = _aggressive_click(driver, join_btn)
                            except Exception:
                                pass
                        except Exception:
                            pass

                    # Priority 2: optional form submit fallback.
                    if not click_sent and step4_enable_form_submit:
                        _wait_for_clickable_element(driver, join_selectors, timeout=5, raise_error=False)
                        click_sent = _submit_signup_form(driver)
                        if click_sent:
                            print(f"[{profile_id}] Step 4: form submit used (attempt {attempt})")

                    if click_sent:
                        join_clicked = True
                    else:
                        _pause_between_actions(
                            driver=driver,
                            profile_id=profile_id,
                            note=f"click Join for Free attempt {attempt}",
                            allow_scroll=True,
                        )
                        print(f"[{profile_id}] Step 4: no click/submit signal sent (attempt {attempt})")

                    if _wait_until(
                        lambda: _find_visible_element(driver, accept_btn_xpaths) is not None or not _is_signup_page(driver),
                        timeout=max(step4_post_submit_wait_seconds, step4_retry_wait_seconds * 1.8),
                        interval=0.25,
                    ):
                        join_clicked = True
                        break
                    if step4_retry_wait_seconds > 0 and attempt < step4_submit_attempts:
                        _pause_between_actions(
                            driver=driver,
                            profile_id=profile_id,
                            note=f"Step 4 retry wait attempt {attempt}",
                            allow_scroll=True,
                        )
                        wait_s = random.uniform(
                            max(0.2, step4_retry_wait_seconds * 0.8),
                            max(0.4, step4_retry_wait_seconds * 1.4),
                        )
                        time.sleep(wait_s)

                if not join_clicked:
                    print(f"[{profile_id}] Step 4: Join button click fallback via Enter key")

            if not join_clicked:
                try:
                    password_el = _wait_for_visible_element(driver, password_input_xpaths, timeout=5)
                    password_el.send_keys(Keys.ENTER)
                    join_clicked = True
                except Exception:
                    pass

            if not join_clicked and step4_enable_form_submit and _submit_signup_form(driver):
                join_clicked = True

            if not join_clicked:
                raise RuntimeError("Step 4 failed: cannot click Join for Free")

            if not _wait_after_signup(
                driver,
                profile_id,
                auto_retry_seconds=signup_retry_seconds,
                manual_wait_seconds=manual_verify_wait_seconds,
            ):
                raise RuntimeError(
                    "Signup did not advance after Join for Free; verification/captcha still blocking."
                )

            # After Join for Free, wait until Terms dialog appears (Step 5 screen).
            print(f"[{profile_id}] Step 4: Waiting for terms dialog (I accept)")
            try:
                _wait_for_visible_element(driver, accept_btn_xpaths, timeout=terms_wait_seconds)
            except Exception:
                raise RuntimeError(
                    f"After Join for Free, terms dialog did not appear within {terms_wait_seconds}s."
                )

        step_name = "step_5_accept_terms"
        print(f"[{profile_id}] Step 5: Accept terms")
        demo_pause(profile_id, "prepare Step 5")
        _dismiss_common_popups(driver)
        accept_btn = _find_visible_element(driver, accept_btn_xpaths)
        if accept_btn is not None:
            js_click(driver, accept_btn)

        step_name = "step_6_start_trial"
        print(f"[{profile_id}] Step 6: Confirm free trial start")
        demo_pause(profile_id, "prepare Step 6")
        _dismiss_common_popups(driver)
        go_course_xpaths = [
            "//button[contains(., 'Go to course')]",
            "//a[contains(., 'Go to course')]",
        ]
        # The user requested that the payment flow is MUST.
        # We will not skip even if already enrolled is detected.
        already_enrolled = _find_visible_element(
            driver,
            go_course_xpaths + ["//*[contains(., 'Already enrolled')]"],
        ) is not None
        if already_enrolled:
            print(f"[{profile_id}] Already enrolled detected, but forced to continue payment flow")
        
        skip_payment = False # FORCE PAYMENT FLOW

        start_trial_xpaths = [
            "//button[contains(., 'Start Free Trial')]",
            "//button[contains(., 'Start free trial')]",
            "//button[contains(., 'Free Trial')]",
            "//a[contains(., 'Start Free Trial')]",
            "//a[contains(., 'Free Trial')]",
        ]
        if not skip_payment:
            try:
                _click_first(driver, start_trial_xpaths, timeout=20)
            except Exception:
                if _is_signup_page(driver):
                    if not _wait_after_signup(
                        driver,
                        profile_id,
                        auto_retry_seconds=max(10, signup_retry_seconds // 2),
                        manual_wait_seconds=manual_verify_wait_seconds,
                    ):
                        raise RuntimeError("Still on signup page; trial flow not reachable due verification.")
                    _click_first(driver, start_trial_xpaths, timeout=20)
                else:
                    # If we can't find 'Start Free Trial', maybe we are already on checkout or another step.
                    print(f"[{profile_id}] Step 6 warning: Could not find 'Start Free Trial' button, continuing...")

        if not skip_payment:
            step_name = "step_7_country_zip"
            print(f"[{profile_id}] Step 7: Fill billing country and zipcode")
            demo_pause(profile_id, "prepare Step 7")
            _dismiss_common_popups(driver)
            time.sleep(2)

            # --- Country selection: handle React-Select searchable dropdown ---
            country_filled = False

            # Strategy A: standard <select> element
            try:
                sel_el = driver.find_element(By.XPATH, "//select[contains(@name, 'country') or contains(@id, 'country')]")
                try:
                    Select(sel_el).select_by_visible_text("USA")
                    country_filled = True
                except Exception:
                    try:
                        Select(sel_el).select_by_visible_text("United States")
                        country_filled = True
                    except: pass
            except Exception:
                pass

            # Strategy B: React-Select combobox (data-e2e="country" or id="cc-country")
            if not country_filled:
                react_country_xpaths = [
                    "//input[@data-e2e='country']",
                    "//input[@id='cc-country']",
                    "//div[contains(@class,'Select')]//input[@role='combobox' or @aria-autocomplete='list']",
                ]
                try:
                    # Wait for input to be present
                    country_input = None
                    deadline = time.time() + 15
                    while time.time() < deadline:
                        country_input = _find_visible_element(driver, react_country_xpaths)
                        if country_input: break
                        
                        # Try clicking placeholder or label to activate it
                        activators = _find_visible_element(driver, [
                            "//*[contains(text(), 'Select your country')]",
                            "//*[contains(@class, 'Select-placeholder')]",
                            "//label[@for='cc-country']"
                        ])
                        if activators: js_click(driver, activators)
                        time.sleep(1)

                    if country_input:
                        # Clear and type USA
                        js_click(driver, country_input)
                        time.sleep(0.5)
                        # More aggressive clearing
                        country_input.send_keys(Keys.CONTROL, 'a')
                        country_input.send_keys(Keys.BACKSPACE)
                        time.sleep(0.5)
                        
                        print(f"[{profile_id}] Step 7: Typing USA and hitting ENTER")
                        country_input.send_keys("USA")
                        time.sleep(1.5)
                        country_input.send_keys(Keys.ENTER)
                        time.sleep(2)
                        
                        # Check if USA/United States is now the value
                        val = country_input.get_attribute("value") or ""
                        text_in_parent = ""
                        try:
                            text_in_parent = country_input.find_element(By.XPATH, "./..").text
                        except: pass
                        
                        if "United States" in val or "USA" in val or "United States" in text_in_parent or "USA" in text_in_parent:
                            print(f"[{profile_id}] Step 7: Country set to USA confirmed")
                            country_filled = True
                        else:
                            # Search for the option in the list as a backup
                            us_option_xpaths = [
                                "//*[contains(@class, 'Select-option') and (contains(., 'United States') or contains(., 'USA'))]",
                                "//*[@role='option' and (contains(., 'United States') or contains(., 'USA'))]",
                                "//*[contains(@id, 'react-select') and (contains(., 'United States') or contains(., 'USA'))]",
                                "//*[text()='United States (USA)' or text()='United States' or text()='USA']",
                            ]
                            us_opt = _find_visible_element(driver, us_option_xpaths)
                            if us_opt:
                                try:
                                    us_opt.click()
                                except:
                                    js_click(driver, us_opt)
                                country_filled = True
                                time.sleep(2)
                            else:
                                # Final backup: Enter again
                                country_input.send_keys(Keys.ENTER)
                                country_filled = True
                                time.sleep(2)
                except Exception as e:
                    print(f"[{profile_id}] Step 7 country React-Select failed: {e}")

            # Strategy C: generic text-based dropdown
            if not country_filled:
                try:
                    country_box = wait.until(EC.element_to_be_clickable(
                        (By.XPATH, "//*[contains(text(), 'Select your country') or contains(text(), 'COUNTRY')]/..")
                    ))
                    js_click(driver, country_box)
                    time.sleep(0.5)
                    us_option = wait.until(
                        EC.element_to_be_clickable((By.XPATH, "//*[text()='United States' or text()='USA' or text()='United States (USA)']"))
                    )
                    js_click(driver, us_option)
                    country_filled = True
                    time.sleep(2)
                except Exception as e:
                    print(f"[{profile_id}] Step 7 country generic dropdown failed: {e}")

            if not country_filled:
                raise RuntimeError("Step 7 failed: could not select country")
            
            # --- ZIP/postal code: wait aggressively ---
            zip_xpaths = [
                "//input[@id='cc-zipcode']",
                "//input[@data-e2e='zipcode']",
                "//input[@data-e2e='postalCode']",
                "//input[@id='cc-postalCode']",
                "//input[contains(@name, 'zip') or contains(@name, 'postal')]",
                "//input[contains(@id, 'zip') or contains(@id, 'postal')]",
                "//input[contains(@placeholder, 'ZIP') or contains(@placeholder, 'Zip') or contains(@placeholder, 'postal')]",
            ]
            print(f"[{profile_id}] Step 7: Waiting for ZIP field...")
            zip_input = _wait_for_visible_element(driver, zip_xpaths, timeout=20, raise_error=False) 
            if zip_input:
                _fill_input(driver, zip_xpaths, zipcode, timeout=10)
                print(f"[{profile_id}] Step 7: ZIP code {zipcode} filled")
            else:
                print(f"[{profile_id}] Step 7 warning: ZIP/postal code input not found after 20s")

            step_name = "step_8_card_and_submit"
            print(f"[{profile_id}] Step 8: Fill card details and submit")
            demo_pause(profile_id, "prepare Step 8")
            _dismiss_common_popups(driver)
            _fill_stripe_fields(driver, card_num, exp_date, cvc, timeout=35)
            driver.switch_to.default_content()
            
            # --- Finish Payment with aggressive click ---
            submit_btn_xpaths = [
                "//button[contains(., 'Start Free Trial')]",
                "//button[contains(., 'Submit')]",
                "//button[contains(., 'Subscribe')]",
                "//button[contains(., 'Start trial')]",
            ]
            btn = _wait_for_visible_element(driver, submit_btn_xpaths, timeout=10, raise_error=False)
            if btn:
                print(f"[{profile_id}] Step 8: Clicking final payment button")
                _aggressive_click(driver, btn)
                # Wait longer for payment processing and transition to 'My commitment' or 'Home'
                print(f"[{profile_id}] Step 8: Waiting for payment processing (up to 20s)...")
                time.sleep(15) 
            else:
                print(f"[{profile_id}] Step 8 warning: Payment button not found, maybe already processed?")
                # Short wait just in case it's transitioning
                time.sleep(5)
        else:
            print(f"[{profile_id}] Skip Step 7/8 payment flow for enrolled account")

        step_name = "step_9_start_course"
        print(f"[{profile_id}] Step 9: Confirm commitment and start course")
        demo_pause(profile_id, "prepare Step 9")
        _dismiss_common_popups(driver)
        
        # Try to find the commitment checkbox (e.g., "I commit to completing this course")
        try:
            commit_xpaths = [
                "//*[contains(text(), 'I commit to completing this course')]",
                "//label[contains(., 'I commit')]//input",
                "//label[contains(., 'I commit')]",
                "//input[@type='checkbox']",
            ]
            commit_checkbox = _wait_for_visible_element(driver, commit_xpaths, timeout=12, raise_error=False)
            if commit_checkbox:
                print(f"[{profile_id}] Step 9: Clicking commitment checkbox")
                _aggressive_click(driver, commit_checkbox)
                time.sleep(2)
        except Exception as e:
            print(f"[{profile_id}] Step 9 note: commitment checkbox error (ignored): {e}")

        start_course_xpaths = [
            "//button[contains(., 'Start the course')]",
            "//a[contains(., 'Start the course')]",
            "//button[contains(., 'Start Course')]",
            "//a[contains(., 'Start Course')]",
            "//button[contains(text(), 'Start the course')]",
        ]
        start_course_btn = _wait_for_visible_element(driver, start_course_xpaths, timeout=12, raise_error=False)
        if start_course_btn:
            print(f"[{profile_id}] Step 9: Clicking 'Start the course'")
            _aggressive_click(driver, start_course_btn)
            time.sleep(8)
        else:
            print(f"[{profile_id}] Step 9 note: start-course button not found, checking if already on course home")
            _take_debug_screenshot(driver, profile_id, "step_9_skipped")

        step_name = "step_10_continue"
        print(f"[{profile_id}] Step 10: Confirm popup and continue")
        demo_pause(profile_id, "prepare Step 10")
        _dismiss_common_popups(driver)
        continue_xpaths = [
            "//button[contains(., 'Continue')]",
            "//a[contains(., 'Continue')]",
            "//button[contains(text(), 'Continue')]",
        ]
        continue_btn = _wait_for_visible_element(driver, continue_xpaths, timeout=10, raise_error=False)
        if continue_btn:
            print(f"[{profile_id}] Step 10: Clicking 'Continue'")
            _aggressive_click(driver, continue_btn)
            time.sleep(5)
        else:
            print(f"[{profile_id}] Step 10 note: continue popup/button not found, moving on")

        step_name = "step_11_module2_redeem"
        print(f"[{profile_id}] Step 11: Open Module 2 target item")
        demo_pause(profile_id, "prepare Step 11")
        _ensure_desktop_viewport(driver, profile_id, reason="Step 11")
        _dismiss_common_popups(driver)
        
        module2_xpaths = [
            "//button[contains(normalize-space(.), 'Module 2')]",
            "//button[contains(normalize-space(.), 'Week 2')]",
            "//button[contains(normalize-space(.), 'Practice using AI')]",
            "//*[@role='button' and contains(normalize-space(.), 'Module 2')]",
            "//*[@role='button' and contains(normalize-space(.), 'Week 2')]",
            "//*[@role='button' and contains(normalize-space(.), 'Practice using AI')]",
            "//summary[contains(normalize-space(.), 'Module 2')]",
            "//summary[contains(normalize-space(.), 'Week 2')]",
            "//summary[contains(normalize-space(.), 'Practice using AI')]",
            "//a[contains(normalize-space(.), 'Module 2')]",
            "//a[contains(normalize-space(.), 'Week 2')]",
            "//a[contains(normalize-space(.), 'Practice using AI')]",
            "//*[contains(normalize-space(.), 'Module 2') and (@aria-expanded or @aria-controls)]",
            "//*[contains(normalize-space(.), 'Practice using AI') and (@aria-expanded or @aria-controls)]",
        ]
        redeem_xpaths = [
            "//a[contains(@aria-label, 'Redeem your Google AI Pro trial') and contains(@href, '/ungradedLti/')]",
            "//a[contains(@aria-label, 'Ungraded App Item') and contains(@aria-label, 'Redeem') and contains(@href, '/ungradedLti/')]",
            "//a[contains(@href, '/learn/google-ai-fundamentals/ungradedLti/') and contains(@aria-label, 'Redeem')]",
            "//a[contains(@href, '/ungradedLti/') and starts-with(normalize-space(.), 'Redeem')]",
            "//a[starts-with(normalize-space(.), 'Redeem')]",
            "//button[starts-with(normalize-space(.), 'Redeem')]",
            "//*[@role='button' and starts-with(normalize-space(.), 'Redeem')]",
            "//*[self::a or self::button][contains(normalize-space(.), 'Redeem your Google AI Pro trial')]",
            "//*[self::a or self::button][contains(normalize-space(.), 'Redeem')]",
            "//*[@role='button' and contains(normalize-space(.), 'Redeem')]",
        ]
        module2_third_item_xpaths = [
            "((//button[contains(normalize-space(.), 'Module 2') or contains(normalize-space(.), 'Week 2')]"
            "/ancestor::*[self::section or self::li or self::article or self::div][1]"
            "//*[self::a or self::button][normalize-space(.) != ''])[3])[1]",
            "((//*[@role='button' and (contains(normalize-space(.), 'Module 2') or contains(normalize-space(.), 'Week 2'))]"
            "/ancestor::*[self::section or self::li or self::article or self::div][1]"
            "//*[self::a or self::button][normalize-space(.) != ''])[3])[1]",
        ]
        redeem_item = None
        for attempt in range(1, 4):
            module2 = _wait_for_clickable_element(driver, module2_xpaths, timeout=6, raise_error=False)
            if module2:
                print(f"[{profile_id}] Step 11: Opening Module 2 (attempt {attempt})")
                try:
                    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", module2)
                except Exception:
                    pass
                clicked = _human_click(driver, module2)
                if not clicked:
                    clicked = _aggressive_click(driver, module2)
                if clicked:
                    time.sleep(random.uniform(1.4, 2.4))

            # Prefer module-local Redeem candidate or module-local 3rd clickable fallback.
            redeem_item = _find_module2_redeem_item(driver)
            if redeem_item is None:
                redeem_item = _wait_for_clickable_element(driver, redeem_xpaths, timeout=3.5, raise_error=False)
            if redeem_item is not None:
                break

        if redeem_item is None:
            third_item = _wait_for_clickable_element(driver, module2_third_item_xpaths, timeout=3, raise_error=False)
            if third_item is not None:
                try:
                    third_text = (third_item.text or "").strip()
                except Exception:
                    third_text = ""
                if third_text.lower().startswith("redeem"):
                    print(f"[{profile_id}] Step 11: Using Module 2 third item: {third_text}")
                    redeem_item = third_item

        if redeem_item:
            try:
                redeem_text = (redeem_item.text or "").strip()
            except Exception:
                redeem_text = ""
            print(f"[{profile_id}] Step 11: Clicking Redeem item: {redeem_text}")
            try:
                driver.execute_script("arguments[0].scrollIntoView({block:'center'});", redeem_item)
            except Exception:
                pass
            clicked = _human_click(driver, redeem_item)
            if not clicked:
                clicked = _aggressive_click(driver, redeem_item)
            time.sleep(5)
        else:
            print(f"[{profile_id}] Step 11 note: redeem entry not found, continue to launch/final link")
            _take_debug_screenshot(driver, profile_id, "step_11_redeem_not_found")

        step_name = "step_12_agree_launch"
        print(f"[{profile_id}] Step 12: Accept honor code and launch app")
        demo_pause(profile_id, "prepare Step 12")
        _dismiss_common_popups(driver)
        try:
            agree_checkbox = wait.until(
                EC.presence_of_element_located(
                    (By.XPATH, "//label[contains(., 'I agree')]//input[@type='checkbox'] | //input[@type='checkbox' and (contains(@name,'agree') or contains(@id,'agree'))]")
                )
            )
            js_click(driver, agree_checkbox)
        except Exception:
            try:
                # fallback to first available checkbox on this step
                honor_checkbox = _wait_for_visible_element(driver, ["//input[@type='checkbox']"], timeout=5)
                js_click(driver, honor_checkbox)
            except Exception:
                print(f"[{profile_id}] Step 12 note: no agreement checkbox found")
        launch_xpaths = [
            "//button[contains(., 'Launch')]",
            "//button[contains(., 'Open App')]",
            "//a[contains(., 'Launch')]",
            "//button[contains(., 'Go to course')]",
            "//button[contains(., 'Start the course')]",
            "//a[contains(., 'Go to course')]",
            "//button[contains(., 'Continue')]",
        ]
        launch_btn = _wait_for_visible_element(driver, launch_xpaths, timeout=15, raise_error=False)
        if launch_btn:
            print(f"[{profile_id}] Step 12: Launching app...")
            js_click(driver, launch_btn)
            time.sleep(8)
        else:
            print(f"[{profile_id}] Step 12 failure: Launch button not found")
            _take_debug_screenshot(driver, profile_id, "step_12_failed")
            raise Exception("Step 12: Launch button not found")
        final_link = extract_final_link(driver, wait)
        if not save_link_to_xml(final_link):
            raise RuntimeError(f"Duplicate/invalid link for this run: {final_link}")

        print(f"[{profile_id}] Link saved. Skip remaining steps and close browser.")
        print(f"[{profile_id}] Workflow completed successfully")
        return

    except Exception as e:
        print(f"[{profile_id}] Workflow blocked with error: {e}")
        save_debug_artifacts(driver, profile_id, step_name, e)

def run_automation(profile_id):
    if _env_flag("RANDOMIZE_FINGERPRINT_ON_START", default=True):
        randomize_fingerprint_for_profile(profile_id)
    else:
        print(f"[{profile_id}] Skip fingerprint randomization (RANDOMIZE_FINGERPRINT_ON_START=0)")
    debug_port, webdriver_path = start_roxy_browser(profile_id)
    if not debug_port:
        return

    chrome_options = Options()
    chrome_options.add_experimental_option("debuggerAddress", debug_port)
    service = Service(executable_path=webdriver_path) if webdriver_path else Service()
    
    driver = None
    try:
        driver = webdriver.Chrome(service=service, options=chrome_options)
        _ensure_desktop_viewport(driver, profile_id, reason="startup")
        if _env_flag("CLEAR_COOKIES_ON_START", default=True):
            clear_cookie_storage_in_session(driver, profile_id)
        else:
            print(f"[{profile_id}] Skip cookie clear (CLEAR_COOKIES_ON_START=0)")
        normalize_entry_page(driver, profile_id)
        run_coursera_workflow(driver, profile_id)
        
    except Exception as e:
        print(f"[{profile_id}] WebDriver connection error: {e}")
        if driver:
            save_debug_artifacts(driver, profile_id, "webdriver_connect", e)
    finally:
        if driver:
            try:
                driver.quit() 
            except:
                pass
        close_roxy_browser(profile_id)

def run_automation_with_stagger(profile_id, task_index):
    delay = get_startup_stagger_seconds(task_index)
    if delay > 0:
        print(f"[{profile_id}] Startup stagger sleep {delay:.2f}s")
        time.sleep(delay)
    run_automation(profile_id)

def _parse_positive_int_or_default(raw_value, default_value, env_name, allow_all=False):
    value = (raw_value or "").strip()
    if not value:
        return default_value
    if allow_all and value.lower() in {"all", "auto", "max", "*", "0"}:
        return default_value
    try:
        parsed = int(value)
    except ValueError:
        print(f"Invalid {env_name}='{value}', fallback to {default_value}")
        return default_value
    if parsed <= 0:
        print(f"Invalid {env_name}='{value}', fallback to {default_value}")
        return default_value
    return parsed

def main():
    show_startup_config_dialog()

    # Auto-fetch all available RoxyBrowser profile IDs for batch run
    workspace_id = get_workspace_id()
    if not workspace_id:
        print("No workspaceId found from env/API/logs.")
        return
    save_workspace_id(workspace_id)

    profile_ids = get_all_profile_ids()
    if not profile_ids:
        print("No profile IDs found from RoxyBrowser API.")
        return

    profile_ids_env = os.getenv("PROFILE_IDS", "").strip()
    if profile_ids_env:
        selected = [x.strip() for x in profile_ids_env.split(",") if x.strip()]
        profile_ids = [x for x in profile_ids if x in selected]
        print(f"PROFILE_IDS filter enabled, remaining profiles={len(profile_ids)}")

    profile_limit_env = os.getenv("PROFILE_LIMIT", "").strip()
    limit = _parse_positive_int_or_default(
        profile_limit_env,
        len(profile_ids),
        "PROFILE_LIMIT",
        allow_all=True,
    )
    if limit < len(profile_ids):
        profile_ids = profile_ids[:limit]
        print(f"PROFILE_LIMIT enabled, remaining profiles={len(profile_ids)}")

    if not profile_ids:
        print("No profiles left after filters.")
        return

    save_profile_ids(profile_ids)
    # Default: run with full parallelism according to profile count in RoxyBrowser.
    default_workers = len(profile_ids)
    workers_from_env = os.getenv("MAX_WORKERS", "").strip()
    max_workers = _parse_positive_int_or_default(
        workers_from_env,
        default_workers,
        "MAX_WORKERS",
        allow_all=True,
    )
    max_workers = max(1, min(max_workers, len(profile_ids)))

    default_loops = 1
    loops_from_env = os.getenv("LOOP_COUNT", str(default_loops)).strip()
    try:
        loop_count = int(loops_from_env)
    except ValueError:
        print(f"Invalid LOOP_COUNT='{loops_from_env}', fallback to {default_loops}")
        loop_count = default_loops
    loop_count = max(1, loop_count)

    print(
        f"Start Coursera automation: max_workers={max_workers}, "
        f"loop_count={loop_count}, profiles={len(profile_ids)}, "
        f"startup_stagger={os.getenv('STARTUP_STAGGER_SECONDS', '0')}, "
        f"startup_jitter={os.getenv('STARTUP_STAGGER_JITTER_SECONDS', '0')}, "
        f"action_delay={os.getenv('ACTION_DELAY_SECONDS', '0.35,0.9')}, "
        f"human_scroll_pixels={os.getenv('HUMAN_SCROLL_PIXELS', '70,220')}, "
        f"human_scroll_prob={os.getenv('HUMAN_SCROLL_PROBABILITY', '0.55')}, "
        f"window_size={os.getenv('WINDOW_WIDTH', '1600')}x{os.getenv('WINDOW_HEIGHT', '980')}, "
        f"min_viewport_width={os.getenv('MIN_VIEWPORT_WIDTH', '1180')}, "
        f"step3_extra_wait={os.getenv('STEP3_EXTRA_WAIT_SECONDS', '1.2')}, "
        f"step4_retry_wait={os.getenv('STEP4_RETRY_WAIT_SECONDS', '1.6')}, "
        f"randomize_fingerprint_on_start={os.getenv('RANDOMIZE_FINGERPRINT_ON_START', '1')}, "
        f"clear_cookies_on_start={os.getenv('CLEAR_COOKIES_ON_START', '1')}"
    )
    for loop_index in range(loop_count):
        current_loop = loop_index + 1
        print(
            f"=== Loop {current_loop}/{loop_count} start: "
            f"run all {len(profile_ids)} profiles in parallel ==="
        )
        with LINK_FILE_LOCK:
            RUN_LINKS.clear()
        with concurrent.futures.ThreadPoolExecutor(max_workers=max_workers) as executor:
            future_to_profile = {
                executor.submit(run_automation_with_stagger, profile_id, task_index): profile_id
                for task_index, profile_id in enumerate(profile_ids)
            }
            for future in concurrent.futures.as_completed(future_to_profile):
                profile_id = future_to_profile[future]
                try:
                    future.result()
                except Exception as e:
                    print(f"[{profile_id}] Unexpected loop error: {e}")
        print(f"=== Loop {current_loop}/{loop_count} completed ===")

if __name__ == "__main__":
    main()
