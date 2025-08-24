#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import sys, os, time, json, tempfile, shutil, re, unicodedata
import pandas as pd
from datetime import datetime, date

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager

from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import (
    StaleElementReferenceException,
    TimeoutException,
    SessionNotCreatedException,
)
import re

# -------- Tunables (you can tweak if needed) --------
PANEL_READY_TIMEOUT    = 30     # seconds to wait for day panel to appear
EVENT_SETTLE_TIMEOUT   = 25     # seconds to wait for list to “settle”
EVENT_SEARCH_TIMEOUT   = 45     # total seconds to search/scroll for your tile
SCROLL_STEP_FRACTION   = 0.60   # each calendar scroll moves ~60% of visible height
SCROLL_PAUSE           = 0.30   # pause between calendar scrolls
AFTER_DATE_CLICK_PAUSE = 1.5    # time for panel to re-render after date click

# NEW: Attendance table search limits so it never hangs
SHORT_FIND_TIMEOUT        = 2   # (s) small waits for elements during search
PER_STUDENT_MAX_SECONDS   = 5  # (s) hard cap per student search
TABLE_SCROLL_TRIES        = 6   # how many small scroll steps in the table
TABLE_SCROLL_PAUSE        = 0.25

# =========================================================
# SLCM URLs
# =========================================================
HOME_URL  = "https://maheslcmtech.lightning.force.com/lightning/page/home"
BASE_URL  = "https://maheslcmtech.lightning.force.com"
LOGIN_URL = "https://maheslcm.manipal.edu/login"

# =========================================================
# CLI parsing (date, workbook path, absentees, subject details)
#   python maa.py <date> <workbook_path> <absentees> <subject_details>
#   <absentees>        -> comma-separated IDs (e.g., "2301,2302")
#   <subject_details>  -> "CourseName|CourseCode|Semester|ClassSection[|Session]"
#                         Session is optional and ignored if ClassSection is like "B-1".
# =========================================================
def parse_arguments():
    if len(sys.argv) < 5:
        print("❌ Usage: python maa.py <date> <workbook_path> <absentees> <subject_details>")
        print("   Example: python maa.py '8/1/2025' '/path/to/workbook.xlsx' '2301,2302' 'OS|CSE 3123|V|B-1'")
        sys.exit(1)
    selected_date_str   = sys.argv[1]
    workbook_path       = sys.argv[2]
    absentees_str       = sys.argv[3]
    subject_details_str = sys.argv[4]
    return selected_date_str, workbook_path, absentees_str, subject_details_str

# =========================================================
# Date parsing (prefer MM/DD/YYYY for slash dates)
# =========================================================
def excel_serial_to_date(n: float) -> date | None:
    try:
        n = float(n)
    except Exception:
        return None

    def from_base(base_year):
        base = datetime(1899, 12, 30) if base_year == 1900 else datetime(1904, 1, 1)
        return (base + pd.to_timedelta(int(n), unit="D")).date()

    candidates = []
    for base in (1900, 1904):
        try:
            d = from_base(base)
            if 1990 <= d.year <= 2100:
                candidates.append(d)
        except Exception:
            pass
    if not candidates:
        try:
            return from_base(1900)
        except Exception:
            return None
    today = date.today()
    candidates.sort(key=lambda d: abs((d - today).days))
    return candidates[0]

def parse_date_any(s) -> date | None:
    """Treat X/Y/Z as MM/DD/YYYY (so 8/1/2025 = Aug 1, 2025)."""
    if s is None:
        return None
    s = unicodedata.normalize("NFC", str(s)).strip()
    if not s:
        return None

    # Excel serial?
    try:
        as_float = float(s)
        d = excel_serial_to_date(as_float)
        if d:
            print(f"📅 Parsed Excel serial {s} -> {d}")
            return d
    except Exception:
        pass

    m = re.fullmatch(r"\s*(\d{1,2})/(\d{1,2})/(\d{2,4})\s*", s)
    if m:
        m1, d1, y1 = int(m.group(1)), int(m.group(2)), int(m.group(3))
        if y1 < 100:
            y1 += 2000 if y1 < 50 else 1900
        try:
            parsed = date(y1, m1, d1)  # MM/DD/YYYY
            print(f"📅 Parsed '{s}' as MM/DD/YYYY -> {parsed}")
            return parsed
        except ValueError:
            try:
                parsed = date(y1, d1, m1)
                print(f"📅 Parsed '{s}' as DD/MM/YYYY fallback -> {parsed}")
                return parsed
            except ValueError:
                pass

    fmts = [
        "%Y-%m-%d",
        "%d-%m-%Y", "%d-%m-%y",
        "%d-%b-%Y", "%d-%b-%y",
        "%d %b %Y", "%d %B %Y",
        "%A, %d %B %Y at %I:%M:%S %p",
        "%A, %d %B %Y",
    ]
    for f in fmts:
        try:
            parsed = datetime.strptime(s, f).date()
            print(f"📅 Parsed '{s}' using '{f}' -> {parsed}")
            return parsed
        except Exception:
            continue

    try:
        parsed = pd.to_datetime(s, dayfirst=False).date()
        print(f"📅 Parsed '{s}' via pandas (dayfirst=False) -> {parsed}")
        return parsed
    except Exception:
        print(f"❌ Could not parse date: {s}")
        return None

# =========================================================
# Selenium helpers
# =========================================================
def js_click(driver, el):
    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", el)
    driver.execute_script("arguments[0].click();", el)

def ready(driver):
    try:
        return driver.execute_script("return document.readyState") == "complete"
    except Exception:
        return False

def close_blank_tabs(driver):
    handles = driver.window_handles[:]
    for h in handles:
        driver.switch_to.window(h)
        url = driver.current_url
        if url.startswith(("about:blank","chrome://newtab","chrome://")):
            try: driver.close()
            except Exception: pass
    if driver.window_handles:
        driver.switch_to.window(driver.window_handles[-1])

def hard_nav(driver, url, attempts=4):
    for _ in range(attempts):
        try:
            driver.get(url); time.sleep(0.5)
            if ready(driver) and driver.current_url.startswith("http"): return True
        except Exception: pass
        try:
            driver.execute_script("window.location.href = arguments[0];", url); time.sleep(0.5)
            if ready(driver) and driver.current_url.startswith("http"): return True
        except Exception: pass
        try:
            driver.switch_to.new_window('tab'); driver.get(url); time.sleep(0.7)
            if ready(driver) and driver.current_url.startswith("http"):
                close_blank_tabs(driver); return True
        except Exception: pass
        time.sleep(0.3)
    close_blank_tabs(driver); return False

def click_calendar_date_fast(driver, day_number: str):
    js = """
    const wrap = document.querySelector('#calendarSidebar');
    if (!wrap) return false;
    const nodes = wrap.querySelectorAll('table.datepicker .slds-day, .slds-day');
    for (const n of nodes) {
        const txt = (n.textContent || '').trim();
        const disabled = n.getAttribute('aria-disabled') === 'true' || (n.className || '').includes('disabled');
        if (!disabled && txt === arguments[0]) {
            n.scrollIntoView({block:'center'});
            n.click();
            return true;
        }
    }
    return false;
    """
    ok = driver.execute_script(js, day_number)
    if not ok:
        raise RuntimeError(f"❌ Could not click mini calendar date {day_number}")
    print(f"✅ Clicked calendar date: {day_number}")

def _norm(s: str) -> str:
    return " ".join((s or "").split())



def _has_word(haystack: str, needle: str) -> bool:
    # word boundary search on A-Z0-9 underscores; tweak if needed
    return re.search(rf'\b{re.escape(needle)}\b', haystack) is not None



def matches_event_text(txt: str, code: str, sem: str, sec: str, sess: str | None) -> bool:
    """
    Section 'B' must NOT match 'B-1'/'B-2'.
    Section 'B-1' must match exactly that token.
    Session is ignored if section already has '-'.
    """
    def norm(s): return " ".join((s or "").split()).upper()
    T = norm(txt)

    ok = True
    if code:
        ok = ok and (code.upper() in T)
    if sem:
        ok = ok and (f"SEMESTER {sem.upper()}" in T)

    if sec:
        secU = sec.strip().upper()

        if "-" in secU:
            # exact token match for things like B-1 (not part of a longer token)
            pat = rf'(?<![A-Z0-9]){re.escape(secU)}(?![A-Z0-9])'
            ok = ok and re.search(pat, T) is not None
        else:
            # letter section like B — must NOT match B-1/B-2
            # 1) explicit "SEC B" or "SECTION B"
            pat1 = rf'\bSEC(?:TION)?\s*[:\-]?\s*{re.escape(secU)}(?!\s*-\s*\d+)\b'
            # 2) parenthesized "(B)"
            pat2 = rf'\(\s*{re.escape(secU)}\s*\)'
            # 3) standalone token …B… not followed by "-<digits>"
            pat3 = rf'(?<![A-Z0-9]){re.escape(secU)}(?!\s*-\s*\d+)(?![A-Z0-9])'

            ok = ok and (
                re.search(pat1, T) is not None or
                re.search(pat2, T) is not None or
                re.search(pat3, T) is not None
            )

    # Only use session when section is broad (no dash)
    if ok and sess and sess.strip() and (not sec or "-" not in sec.strip()):
        s = sess.strip().upper()
        ok = ok and (f"SESSION {s}" in T)

    return ok


# ---------- Day panel helpers ----------
def day_header_strings(d: date):
    parts = [
        d.strftime("%A, %B %-d") if sys.platform != "win32" else d.strftime("%A, %B %#d"),
        d.strftime("%A, %B %d").lstrip("0").replace(", 0", ", "),
    ]
    seen, out = set(), []
    for s in parts:
        s = _norm(s)
        if s not in seen:
            seen.add(s); out.append(s)
    return out

def find_day_panel_for_date(driver, selected_date):
    headers = driver.find_elements(By.CSS_SELECTOR, "h2.slds-assistive-text")
    wanted = [h.lower() for h in day_header_strings(selected_date)]
    for h in headers:
        try:
            txt = _norm(h.text).lower()
            if txt in wanted:
                panel = h.find_element(By.XPATH, "following-sibling::div[contains(@class,'calendarDay')][1]")
                return panel
        except Exception:
            continue
    return None

def wait_for_day_panel_ready(driver, selected_date, timeout=PANEL_READY_TIMEOUT):
    t0 = time.time()
    while time.time() - t0 < timeout:
        panel = find_day_panel_for_date(driver, selected_date)
        if panel is not None:
            try:
                panel.find_element(By.CSS_SELECTOR, "div.eventList ul.eventListContainer")
                return panel
            except Exception:
                pass
        time.sleep(0.25)
    return None

def wait_for_events_to_settle(driver, panel, timeout=EVENT_SETTLE_TIMEOUT):
    t0 = time.time()
    stable_since = None
    last = (-1, -1)
    while time.time() - t0 < timeout:
        try:
            metrics = driver.execute_script("""
                const panel = arguments[0];
                const list = panel.querySelector("div.eventList");
                const cont = list ? list.querySelector("ul.eventListContainer") : null;
                const h = list ? list.scrollHeight : 0;
                const c = cont ? cont.children.length : 0;
                return [h, c];
            """, panel)
        except Exception:
            metrics = (0, 0)
        if metrics == last:
            if stable_since is None:
                stable_since = time.time()
            elif time.time() - stable_since >= 0.8:
                return True
        else:
            last = metrics
            stable_since = None
        time.sleep(0.25)
    return False

def aria_date_matches_selected(aria_desc: str, selected_date: date) -> bool:
    if not aria_desc:
        return False
    txt = aria_desc.replace("–", "-")
    m = re.search(r"([A-Za-z]+)\s+(\d{1,2})\s+([A-Za-z]+),\s*(\d{4})", txt)
    if not m:
        return False
    try:
        day, mon, yr = int(m.group(2)), m.group(3), int(m.group(4))
        parsed = datetime.strptime(f"{day} {mon} {yr}", "%d %B %Y").date()
        return parsed == selected_date
    except Exception:
        return False

def scroll_day_panel_gradual(driver, panel, max_seconds, code, sem, sec, sess, selected_date):
    start = time.time()
    seen_bottom = False

    def collect_candidates():
        try:
            links = driver.execute_script("""
                const p = arguments[0];
                return Array.from(p.querySelectorAll("a.subject-link, a[data-id='subject-link'], a"))
                    .filter(a => (a.innerText||a.textContent||"").trim().length > 0);
            """, panel)
        except Exception:
            links = panel.find_elements(By.CSS_SELECTOR, "a")
        out = []
        for a in links:
            try:
                title = (a.get_attribute("innerText") or a.text or "").strip()
                if not title or not matches_event_text(title, code, sem, sec, sess):
                    continue
                aria = a.get_attribute("aria-description") or ""
                if aria_date_matches_selected(aria, selected_date):
                    out.append(a)
            except Exception:
                pass
        return out

    cand = collect_candidates()
    if cand:
        return cand[0]

    while time.time() - start < max_seconds:
        try:
            list_container = panel.find_element(By.CSS_SELECTOR, "div.eventList")
        except Exception:
            list_container = panel

        try:
            curTop, curH, cliH = driver.execute_script("""
                const el = arguments[0];
                return [el.scrollTop, el.scrollHeight, el.clientHeight];
            """, list_container)
        except Exception:
            curTop, curH, cliH = 0, 0, 0

        step = max(40, int(cliH * SCROLL_STEP_FRACTION)) if cliH else 250
        newTop = curTop + step
        if curH and newTop >= (curH - cliH - 2):
            newTop = curH
            seen_bottom = True

        try:
            driver.execute_script("arguments[0].scrollTop = arguments[1];", list_container, newTop)
        except Exception:
            pass

        time.sleep(SCROLL_PAUSE)

        cand = collect_candidates()
        if cand:
            return cand[0]

        if seen_bottom:
            time.sleep(0.4)
            cand = collect_candidates()
            if cand:
                return cand[0]
            break

    return None

# =========================================================
# Main
# =========================================================
def main():
    print("🚀 SLCM Attendance Automation Started")
    print("====================================================")

    selected_date_str, workbook_path, absentees_str, subject_details_str = parse_arguments()

    selected_date = parse_date_any(selected_date_str)
    if not selected_date:
        print(f"❌ Could not parse date: {selected_date_str}")
        sys.exit(1)

    absentees = [s.strip() for s in (absentees_str or "").split(",") if s.strip()]

    parts = subject_details_str.split("|")
    if len(parts) < 4:
        print(f"❌ Invalid subject details: {subject_details_str}")
        sys.exit(1)
    course_name   = parts[0].strip()
    course_code   = parts[1].strip()
    semester      = parts[2].strip()
    class_section = parts[3].strip()
    session_no    = parts[4].strip() if len(parts) > 4 else ""  # optional

    print(f"📅 Selected Date : {selected_date}")
    print(f"📂 Workbook      : {workbook_path}")
    print(f"🧑‍🎓 Absentees   : {', '.join(absentees) if absentees else 'None'}")
    print("\n📘 Course Details")
    print(f"   Course Name   : {course_name or '(blank)'}")
    print(f"   Course Code   : {course_code or '(blank)'}")
    print(f"   Semester      : {semester or '(blank)'}")
    print(f"   Class Section : {class_section or '(blank)'}  (supports 'B' or 'B-1')")

    missing = []
    if not course_code:   missing.append("Course Code")
    if not semester:      missing.append("Semester")
    if not class_section: missing.append("Class Section")
    if missing:
        print("❌ Missing required subject details:")
        for m in missing: print(f"   - {m}")
        sys.exit(1)

    # ---- Selenium profile ----
    from pathlib import Path
    def pick_profile_dir():
        home = Path.home()
        d1 = home / ".slcm_automation_profile"
        try:
            d1.mkdir(parents=True, exist_ok=True)
            return str(d1)
        except Exception:
            pass
        d2 = Path(tempfile.gettempdir()) / f"slcm_automation_profile_{os.getuid() if hasattr(os, 'getuid') else 'user'}"
        d2.mkdir(parents=True, exist_ok=True)
        return str(d2)

    PROFILE_DIR = pick_profile_dir()
    print(f"👤 Using Chrome profile dir: {PROFILE_DIR}")

    for name in os.listdir(PROFILE_DIR):
        if name.startswith("Singleton"):
            try: os.remove(os.path.join(PROFILE_DIR, name))
            except Exception: pass

    TEMP_PROFILE_DIR = None
    def build_options(user_data_dir):
        opts = webdriver.ChromeOptions()
        opts.add_argument(f"--user-data-dir={user_data_dir}")
        opts.add_argument("--no-first-run")
        opts.add_argument("--no-default-browser-check")
        # opts.add_argument("--headless=new")
        return opts

    def start_driver_with_fallback():
        nonlocal TEMP_PROFILE_DIR
        try:
            options = build_options(PROFILE_DIR)
            service = Service(ChromeDriverManager().install())
            return webdriver.Chrome(service=service, options=options)
        except SessionNotCreatedException:
            print("⚠️ Profile is locked. Using fresh temp profile…")
            TEMP_PROFILE_DIR = tempfile.mkdtemp(prefix="slcm_profile_")
            options = build_options(TEMP_PROFILE_DIR)
            service = Service(ChromeDriverManager().install())
            return webdriver.Chrome(service=service, options=options)

    driver = start_driver_with_fallback()

    try:
        # Go to Home (user completes SSO if needed)
        if not hard_nav(driver, HOME_URL):
            hard_nav(driver, BASE_URL)
            hard_nav(driver, HOME_URL)

        cur = driver.current_url.lower()
        print("🌐 Current URL:", cur)
        if ("login.microsoftonline.com" in cur) or ("saml" in cur) or ("manipal.edu" in cur and "/login" in cur):
            print("🔐 SSO/login detected. Complete in Chrome, then continue here.")
            input("Press Enter AFTER Salesforce Home is visible… ")
            hard_nav(driver, HOME_URL)

        WebDriverWait(driver, 60).until(EC.presence_of_element_located((By.XPATH, "//a[@title='Calendar']")))
        print("✅ Logged in & on Lightning Home")

        # Open Calendar
        cal_tab = WebDriverWait(driver, 40).until(
            EC.element_to_be_clickable((By.XPATH, "//a[@title='Calendar']"))
        )
        js_click(driver, cal_tab)
        print("✅ Opened Calendar")

        # Wait for mini calendar & click selected date
        WebDriverWait(driver, 12).until(EC.presence_of_element_located((By.ID, "calendarSidebar")))
        time.sleep(0.25)
        day_number = str(selected_date.day).lstrip("0")
        click_calendar_date_fast(driver, day_number)
        time.sleep(AFTER_DATE_CLICK_PAUSE)  # give time to repaint day panel

        # === robust wait for correct day panel, then wait for events to settle
        panel = wait_for_day_panel_ready(driver, selected_date, timeout=PANEL_READY_TIMEOUT)
        if not panel:
            raise RuntimeError("❌ Could not locate the correct day panel for the selected date.")
        if not wait_for_events_to_settle(driver, panel, timeout=EVENT_SETTLE_TIMEOUT):
            print("⚠️ Event list still changing; proceeding with search anyway.")

        # === gradual scroll + polling for the exact tile (up to EVENT_SEARCH_TIMEOUT)
        target = scroll_day_panel_gradual(
            driver, panel, max_seconds=EVENT_SEARCH_TIMEOUT,
            code=course_code, sem=semester, sec=class_section,
            sess=(session_no if "-" not in class_section.strip() else None),
            selected_date=selected_date
        )
        if not target:
            raise RuntimeError("❌ Could not locate the event tile for the selected date (after scrolling & waiting).")

        # Click target
        try:
            driver.execute_script("arguments[0].scrollIntoView({block:'center'});", target)
            time.sleep(0.15)
            target.click()
            print("✅ Opened the event tile")
        except Exception as e:
            raise RuntimeError(f"❌ Failed to click the event tile: {e}")

        # Popover "More Details" (if shown)
        try:
            more_details = WebDriverWait(driver, 8).until(
                EC.element_to_be_clickable((By.XPATH, "//a[normalize-space()='More Details']"))
            )
            js_click(driver, more_details)
            print("✅ Clicked 'More Details'")
        except Exception:
            pass

        # Attendance tab
        time.sleep(0.6)
        def click_attendance_tab_fast(driver):
            js = """
            let el = document.querySelector("a[data-label='Attendance']");
            if (!el) {
                const span = Array.from(document.querySelectorAll('span.title'))
                    .find(s => (s.textContent || '').trim() === 'Attendance');
                if (span) el = span.closest('a, button, [role="tab"]') || span;
            }
            if (el) { el.scrollIntoView({block:'center'}); el.click(); return true; }
            return false;
            """
            if driver.execute_script(js):
                print("✅ Opened Attendance tab (fast)")
                return True
            try:
                att_tab = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.XPATH, "//a[@data-label='Attendance'] | //span[@class='title' and normalize-space()='Attendance']"))
                )
                js_click(driver, att_tab)
                print("✅ Opened Attendance tab (fallback)")
                return True
            except Exception:
                return False

        if not click_attendance_tab_fast(driver):
            print("⚠️ Could not switch to Attendance tab automatically. Please click it, then press Enter here…")
            input()

        # =============================
        # Process absentees (UNTICK boxes) — never hang per student
        # =============================
        print(f"🔎 Processing attendance for {len(absentees)} absentees…")
        unticked_ids, not_found = [], []

        # helper: get the scrollable table container (if any)
        def get_table_container():
            try:
                return WebDriverWait(driver, 4).until(
                    EC.presence_of_element_located((By.XPATH, "//div[contains(@class,'slds-table') or contains(@class,'slds-scrollable')]//table/ancestor::div[contains(@class,'scroll') or contains(@class,'slds-scrollable')][1]"))
                )
            except Exception:
                # fallback to any scrollable table region
                try:
                    return driver.find_element(By.XPATH, "//div[contains(@class,'slds-scrollable')]")
                except Exception:
                    return None

        def try_find_cell_for_id(student_id):
            """
            Fast, bounded attempts to locate the student's id cell.
            Returns WebElement (cell) or None.
            """
            # 1) Direct XPath: exact formatted text
            xps = [
                f"//lightning-base-formatted-text[normalize-space()='{student_id}']",
                f"//td[normalize-space()='{student_id}']",
                f"//*[contains(@class,'formatted-text') and normalize-space()='{student_id}']",
            ]
            for xp in xps:
                try:
                    return WebDriverWait(driver, SHORT_FIND_TIMEOUT).until(
                        EC.presence_of_element_located((By.XPATH, xp))
                    )
                except TimeoutException:
                    continue
                except StaleElementReferenceException:
                    continue

            # 2) JS scan as fallback (quicker than long waits)
            try:
                cell = driver.execute_script("""
                    const sid = arguments[0].trim();
                    const exact = (t) => (t||'').trim() === sid;
                    const nodes = Array.from(document.querySelectorAll(
                       "lightning-base-formatted-text, td, span"
                    ));
                    for (const n of nodes) {
                      const txt = (n.innerText || n.textContent || "").trim();
                      if (txt === sid) return n;
                    }
                    return null;
                """, student_id)
                if cell:
                    return cell
            except Exception:
                pass

            return None

        def find_checkbox_from_cell(cell):
            try:
                row = cell.find_element(By.XPATH, "./ancestor::tr")
                return row.find_element(By.XPATH, ".//input[@type='checkbox']")
            except Exception:
                return None

        def scroll_table_small_steps():
            cont = get_table_container()
            if not cont:
                return
            try:
                for _ in range(TABLE_SCROLL_TRIES):
                    driver.execute_script(
                        "arguments[0].scrollTop = Math.min(arguments[0].scrollTop + Math.max(80, arguments[0].clientHeight*0.35), arguments[0].scrollHeight);",
                        cont
                    )
                    time.sleep(TABLE_SCROLL_PAUSE)
            except Exception:
                pass

        def process_one_absentee(student_id):
            """
            Returns:
              True  => unticked
              False => already unticked
              None  => not found (within time budget)
            """
            t0 = time.time()
            attempts = 0
            while time.time() - t0 < PER_STUDENT_MAX_SECONDS:
                attempts += 1
                try:
                    cell = try_find_cell_for_id(student_id)
                    if not cell:
                        # Try to reveal more rows by small table scrolls
                        scroll_table_small_steps()
                        continue

                    # found a cell, get checkbox
                    checkbox = find_checkbox_from_cell(cell)
                    if not checkbox:
                        scroll_table_small_steps()
                        continue

                    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", checkbox)
                    if checkbox.is_selected():
                        js_click(driver, checkbox)
                        return True
                    else:
                        return False

                except (StaleElementReferenceException, TimeoutException):
                    continue
                except Exception:
                    # Avoid hard stops on weird row types; keep trying within time budget
                    continue

            # out of time for this student
            return None

        for ab in absentees:
            result = process_one_absentee(ab)
            if result is True:
                print(f"✔️ Unticked: {ab}")
                unticked_ids.append(ab)
            elif result is False:
                print(f"ℹ️ Already unticked: {ab}")
            else:
                print(f"❌ Not found on page (skipped safely): {ab}")
                not_found.append(ab)

        print("\n📊 Attendance Summary")
        print("=" * 40)
        print(f"✔️ Successfully unticked: {len(unticked_ids)}")
        print(f"❌ Not found: {len(not_found)}")
        if not_found:
            for nf in not_found:
                print(f"   - {nf}")

        # Submit
        try:
            submit_btn = WebDriverWait(driver, 20).until(
                EC.element_to_be_clickable((By.XPATH, "//button[contains(., 'Submit Attendance')]"))
            )
            js_click(driver, submit_btn)
            print("✅ Clicked Submit Attendance")

            modal = WebDriverWait(driver, 22).until(
                EC.presence_of_element_located((By.XPATH, "//div[contains(@class,'modal-container') or contains(@class,'uiModal') or contains(@class,'slds-modal')]"))
            )
            WebDriverWait(driver, 12).until(EC.visibility_of(modal))
            print("✅ Confirmation modal visible")

            xps = [
                ".//button[normalize-space()='Confirm Submission']",
                ".//button[.//span[normalize-space()='Confirm Submission']]",
                ".//button[contains(.,'Confirm Submission')]",
                ".//footer//*[self::button or self::*[contains(@class,'slds-button')]][contains(.,'Confirm') and contains(@class,'slds-button_brand')]",
                ".//button[contains(.,'Confirm') and contains(@class,'slds-button_brand')]",
            ]
            clicked = False
            for xp in xps:
                try:
                    btn = WebDriverWait(modal, 8).until(EC.element_to_be_clickable((By.XPATH, xp)))
                    js_click(driver, btn)
                    print("✅ Confirmed submission")
                    clicked = True
                    break
                except Exception:
                    continue
            if not clicked:
                btn = driver.execute_script("""
                    const modal = document.querySelector('.modal-container, .uiModal, .slds-modal');
                    if (!modal) return null;
                    const btns = Array.from(modal.querySelectorAll('button, .slds-button'));
                    const norm = t => (t || '').trim().toLowerCase();
                    return btns.find(b => {
                      const txt = norm(b.innerText || b.textContent);
                      return txt === 'confirm submission' || txt === 'confirm' || txt.includes('confirm submission');
                    }) || null;
                """)
                if btn:
                    driver.execute_script("arguments[0].click();", btn)
                    print("✅ Confirmed via JS fallback")
                else:
                    try:
                        modal.send_keys(Keys.ENTER)
                        print("↩️ Sent ENTER to modal (fallback)")
                    except Exception:
                        print("⚠️ Please click Confirm manually.")

        except Exception as e:
            print(f"⚠️ Could not submit attendance: {e}")

        print("\n🎉 SLCM Attendance automation completed!")
        time.sleep(1.3)

    except Exception as e:
        print(f"❌ Error during automation: {e}")
        import traceback; traceback.print_exc()

    finally:
        try: driver.quit()
        except Exception: pass
        if 'TEMP_PROFILE_DIR' in locals() and TEMP_PROFILE_DIR:
            try: shutil.rmtree(TEMP_PROFILE_DIR, ignore_errors=True)
            except Exception: pass

    print("\n====================================================")
    print("👨‍💻 Developed by: Anirudhan Adukkathayar C, SCE, MIT")
    print("====================================================")

if __name__ == "__main__":
    main()
