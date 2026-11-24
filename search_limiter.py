import hashlib
import json
import os
import tempfile
import time

import streamlit as st

STORE = os.environ.get("SEARCH_LIMIT_STORE", "search_usage.json")
LIMIT = int(os.environ.get("SEARCH_LIMIT", "8"))
WINDOW_HOURS = float(os.environ.get("SEARCH_LIMIT_HOURS", "12"))
EXEMPT_KEY = "w_limiter_bypass"


def _window():
    return WINDOW_HOURS * 3600.0


def client_key():
    """A per-visitor key. The IP is what Streamlit Cloud can see; when it is
    unavailable we fall back to a per-session id, which resets on reload."""
    try:
        ip = st.context.ip_address
    except Exception:
        ip = None
    if ip:
        return "ip:" + hashlib.sha256(str(ip).encode("utf-8")).hexdigest()[:20]
    try:
        sid = st.session_state.get("_limiter_sid")
        if not sid:
            sid = hashlib.sha256(
                f"{time.time()}{id(st.session_state)}".encode("utf-8")
            ).hexdigest()[:20]
            st.session_state["_limiter_sid"] = sid
        return "sid:" + sid
    except Exception:
        return "anon"


def _load():
    try:
        with open(STORE, encoding="utf-8") as fh:
            data = json.load(fh)
        return data if isinstance(data, dict) else {}
    except (OSError, ValueError):
        return {}


def _save(data):
    try:
        folder = os.path.dirname(os.path.abspath(STORE)) or "."
        fd, tmp = tempfile.mkstemp(dir=folder, suffix=".tmp")
        with os.fdopen(fd, "w", encoding="utf-8") as fh:
            json.dump(data, fh)
        os.replace(tmp, STORE)
    except OSError:
        pass


def _prune(data, now):
    cutoff = now - _window()
    out = {}
    for key, stamps in data.items():
        kept = [t for t in stamps if isinstance(t, (int, float)) and t > cutoff]
        if kept:
            out[key] = kept
    return out


def exempt():
    try:
        return bool(st.session_state.get(EXEMPT_KEY))
    except Exception:
        return False


def status(key=None):
    """(used, remaining, seconds_until_one_frees_up) without recording a use."""
    if exempt():
        return 0, LIMIT, 0.0
    now = time.time()
    key = key or client_key()
    data = _prune(_load(), now)
    stamps = sorted(data.get(key, []))
    used = len(stamps)
    remaining = max(0, LIMIT - used)
    wait = 0.0
    if remaining == 0 and stamps:
        wait = max(0.0, stamps[0] + _window() - now)
    return used, remaining, wait


def consume(key=None):
    """Record one search. Returns (allowed, remaining_after, wait_seconds)."""
    if exempt():
        return True, LIMIT, 0.0
    now = time.time()
    key = key or client_key()
    data = _prune(_load(), now)
    stamps = sorted(data.get(key, []))
    if len(stamps) >= LIMIT:
        wait = max(0.0, stamps[0] + _window() - now)
        return False, 0, wait
    stamps.append(now)
    data[key] = stamps
    _save(data)
    return True, max(0, LIMIT - len(stamps)), 0.0


def human_wait(seconds):
    seconds = int(max(0, seconds))
    hours, rest = divmod(seconds, 3600)
    minutes = rest // 60
    if hours and minutes:
        return f"{hours}h {minutes}m"
    if hours:
        return f"{hours}h"
    if minutes:
        return f"{minutes}m"
    return "under a minute"


def render_quota(container=None):
    target = container or st
    used, remaining, wait = status()
    if exempt():
        return
    if remaining > 0:
        target.caption(f"{remaining} of {LIMIT} searches left "
                       f"(resets {int(WINDOW_HOURS)}h after each one).")
    else:
        target.caption(f"Search limit reached. One frees up in "
                       f"{human_wait(wait)}.")


def block_message(wait, container=None):
    target = container or st
    target.warning(
        f"You have used all {LIMIT} searches for now. The next one becomes "
        f"available in {human_wait(wait)}. Results already on screen stay "
        "available, and you can keep opening appraisals and endpoints.")
