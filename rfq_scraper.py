"""
RFQ Scraper - fetches RFQ status from internal company system
Credentials are read from config.txt in the parent (dist) directory
"""
import requests
import re
import time
import os
import sys
from html.parser import HTMLParser

# ── Config path (works when imported from dashboard.py) ──────────────────────
def _get_config():
    """Read config.txt from the dist folder."""
    # Could be run from dist/ OR _Developer_Tools/
    base = os.path.dirname(os.path.abspath(__file__))
    for candidate in [base, os.path.dirname(base)]:
        p = os.path.join(candidate, "config.txt")
        if os.path.exists(p):
            cfg = {}
            with open(p, encoding="utf-8") as f:
                for line in f:
                    line = line.strip()
                    if "=" in line and not line.startswith("#"):
                        k, v = line.split("=", 1)
                        cfg[k.strip()] = v.strip()
            return cfg
    return {}

_config = _get_config()

RFQ_BASE_URL   = _config.get("RFQ_URL",      "http://192.168.68.33")
RFQ_EMAIL      = _config.get("RFQ_EMAIL",    "kcchee@genxai.com.my")
RFQ_PASSWORD   = _config.get("RFQ_PASSWORD", "12345678")

LOGIN_URL = f"{RFQ_BASE_URL}/genxai/auth/login"
RFQ_URL   = f"{RFQ_BASE_URL}/genxai/rfqs"

# ── Cache (so we don't hammer the server every second) ──────────────────────
_cache_data = None
_cache_time = 0
CACHE_TTL   = 300  # 5 minutes (Balanced frequency)


def _strip_tags(text):
    return re.sub(r"<[^>]+>", "", text or "").strip()


def _login_and_scrape():
    """Login and return list of RFQ dicts."""
    session = requests.Session()
    session.headers.update({"User-Agent": "Mozilla/5.0"})

    # Get CSRF token
    try:
        resp = session.get(LOGIN_URL, timeout=10)
    except Exception as e:
        return None, f"Cannot reach RFQ site: {e}"

    csrf_match = re.search(
        r'<meta name=["\']csrf-token["\'] content=["\']([^"\']+)["\']',
        resp.text
    )
    csrf_token = csrf_match.group(1) if csrf_match else ""

    # Login
    login_resp = session.post(
        LOGIN_URL,
        data={"email": RFQ_EMAIL, "password": RFQ_PASSWORD, "_token": csrf_token},
        allow_redirects=True,
        timeout=10
    )

    if "auth/login" in login_resp.url:   # still on login page = failed
        return None, "RFQ login failed — check credentials in config.txt"

    # Scrape first 2 pages of RFQs to show more history
    all_results = []
    for p_idx in range(1, 3):
        url = RFQ_URL
        if p_idx > 1: url += f"?page={p_idx}"
        
        try:
            resp = session.get(url, timeout=10)
            if resp.status_code == 200:
                all_results.extend(_parse_rfq_table(resp.text))
        except: break # stop if page fails
        
    return all_results, None


def _parse_rfq_table(html):
    """Parse the RFQ HTML table and return list of dicts."""
    # Extract the <table> block
    table_match = re.search(r"<table[^>]*>(.*?)</table>", html, re.DOTALL | re.IGNORECASE)
    if not table_match:
        return []

    table_html = table_match.group(1)

    # Headers from <th>
    headers_raw = re.findall(r"<th[^>]*>(.*?)</th>", table_html, re.DOTALL | re.IGNORECASE)
    headers = [_strip_tags(h) for h in headers_raw]

    # Rows from <tr>
    rows_raw = re.findall(r"<tr[^>]*>(.*?)</tr>", table_html, re.DOTALL | re.IGNORECASE)

    results = []
    for row_html in rows_raw[1:]:   # skip header row
        cells_raw = re.findall(r"<td[^>]*>(.*?)</td>", row_html, re.DOTALL | re.IGNORECASE)
        cells = [_strip_tags(c) for c in cells_raw]
        if not cells:
            continue
        row_dict = {}
        for i, h in enumerate(headers):
            row_dict[h] = cells[i] if i < len(cells) else ""
        results.append(row_dict)

    return results


def get_rfq_data():
    """
    Returns (list_of_rfq_dicts, error_message_or_None).
    Uses a 5-minute cache.
    """
    global _cache_data, _cache_time
    now = time.time()
    if _cache_data is not None and (now - _cache_time) < CACHE_TTL:
        return _cache_data, None

    data, err = _login_and_scrape()
    if data is not None:
        _cache_data = data
        _cache_time = now
    return data, err


def clear_cache():
    """Manual override to force a fresh scrape next time get_rfq_data is called."""
    global _cache_data, _cache_time
    _cache_data = None
    _cache_time = 0


# ── Quick self-test ──────────────────────────────────────────────────────────
if __name__ == "__main__":
    import json
    data, err = get_rfq_data()
    if err:
        print(f"ERROR: {err}")
    else:
        print(f"Fetched {len(data)} RFQ rows")
        if data:
            print("Columns:", list(data[0].keys()))
            print("First row:", json.dumps(data[0], ensure_ascii=False, indent=2))
