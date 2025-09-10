#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
maa.py — SLCM Attendance automation (complete file)
Includes:
 - quick-precheck to avoid unnecessary month prev/next navigation
 - collect_candidates() computing top relative to calendar wrapper
 - click_fragment() which records window.__slcm_last_clicked_date
 - robust panel_opened_ok() detection
 - multi-step fallback when calendar navigation fails
"""

import sys
import os
import time
import json
import tempfile
import shutil
import re
import unicodedata
import calendar
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

# -------- Tunables (you can tweak if needed) --------
PANEL_READY_TIMEOUT    = 30
EVENT_SETTLE_TIMEOUT   = 25
EVENT_SEARCH_TIMEOUT   = 45
SCROLL_STEP_FRACTION   = 0.60
SCROLL_PAUSE           = 0.30
AFTER_DATE_CLICK_PAUSE = 1.0  # reduced for speed

SHORT_FIND_TIMEOUT        = 2
PER_STUDENT_MAX_SECONDS   = 5
TABLE_SCROLL_TRIES        = 6
TABLE_SCROLL_PAUSE        = 0.20

# Debugging for calendar nav
DEBUG = False  # set False to reduce calendar debugging output

# =========================================================
# SLCM URLs
# =========================================================
HOME_URL  = "https://maheslcmtech.lightning.force.com/lightning/page/home"
BASE_URL  = "https://maheslcmtech.lightning.force.com"
LOGIN_URL = "https://maheslcm.manipal.edu/login"

# =========================================================
# CLI parsing (date, workbook path, absentees, subject details)
# =========================================================
def parse_arguments():
    if len(sys.argv) < 5:
        print("❌ Usage: python maa.py <date> <workbook_path> <absentees> <subject_details>")
        sys.exit(1)
    selected_date_str   = sys.argv[1]
    workbook_path       = sys.argv[2]
    absentees_str       = sys.argv[3]
    subject_details_str = sys.argv[4]
    return selected_date_str, workbook_path, absentees_str, subject_details_str

# =========================================================
# Date parsing helpers
# =========================================================
def excel_serial_to_date(n):
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

def parse_date_any(s):
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
            if DEBUG: print(f"📅 Parsed Excel serial {s} -> {d}")
            return d
    except Exception:
        pass

    m = re.fullmatch(r"\s*(\d{1,2})/(\d{1,2})/(\d{2,4})\s*", s)
    if m:
        m1, d1, y1 = int(m.group(1)), int(m.group(2)), int(m.group(3))
        if y1 < 100:
            y1 += 2000 if y1 < 50 else 1900
        try:
            parsed = date(y1, m1, d1)
            if DEBUG: print(f"📅 Parsed '{s}' as MM/DD/YYYY -> {parsed}")
            return parsed
        except ValueError:
            try:
                parsed = date(y1, d1, m1)
                if DEBUG: print(f"📅 Parsed '{s}' as DD/MM/YYYY fallback -> {parsed}")
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
            if DEBUG: print(f"📅 Parsed '{s}' using '{f}' -> {parsed}")
            return parsed
        except Exception:
            continue

    try:
        parsed = pd.to_datetime(s, dayfirst=False).date()
        if DEBUG: print(f"📅 Parsed '{s}' via pandas (dayfirst=False) -> {parsed}")
        return parsed
    except Exception:
        if DEBUG: print(f"❌ Could not parse date: {s}")
        return None

# =========================================================
# Subject details parsing
# =========================================================
def parse_subject_details(details):
    raw = unicodedata.normalize("NFC", str(details or "")).strip()
    if not raw:
        return None, "empty subject details"
    if "::" in raw:
        parts = raw.split("::")
    else:
        parts = raw.replace("^|", "|").split("|")
    parts += [""] * (5 - len(parts))
    course_name, course_code, semester, class_section, session_no = [p.strip() for p in parts[:5]]

    missing = []
    if not course_code:   missing.append("Course Code")
    if not semester:      missing.append("Semester")
    if not class_section: missing.append("Class Section")
    if missing:
        return None, f"missing required fields: {', '.join(missing)} (raw={raw!r})"

    return (course_name, course_code, semester, class_section, session_no), None

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

# =========================================================
# Mini-calendar helpers
# =========================================================
def _sidebar_month_label(driver):
    js = """
    const wrap = document.querySelector('#calendarSidebar') || document.querySelector('.calendarSidebar') || document.getElementById('calendarSidebar');
    if (!wrap) return null;
    const all = wrap.querySelectorAll('*');
    const texts = [];
    for (const el of all) {
      const txt = (el.innerText || el.textContent || '').trim();
      if (txt) texts.push(txt);
    }
    return texts.join(" || ");
    """
    try:
        raw = driver.execute_script(js)
    except Exception:
        return None
    if not raw:
        return None

    m = re.search(r"\b(January|February|March|April|May|June|July|August|September|October|November|December|"
                  r"Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Sept|Oct|Nov|Dec)\b(?:[^0-9]{0,3}(\d{4}))?", raw, re.I)
    if m:
        mon = m.group(1)
        yr = m.group(2)
        return f"{mon} {yr}" if yr else mon
    return None

def _parse_month_label_to_date(lbl):
    if not lbl:
        return None, None
    lbl = lbl.strip()

    fmts = ("%B %Y", "%b %Y")
    for f in fmts:
        try:
            dt = datetime.strptime(lbl, f)
            return dt.year, dt.month
        except Exception:
            pass

    try:
        dt = datetime.strptime(lbl, "%B")
    except Exception:
        try:
            dt = datetime.strptime(lbl, "%b")
        except Exception:
            return None, None
    return date.today().year, dt.month

def _click_sidebar_prev_next_once(driver, which):
    js = """
    const which = arguments[0];
    const wrap = document.querySelector('#calendarSidebar') || document.querySelector('.calendarSidebar') || document.getElementById('calendarSidebar');
    if (!wrap) return false;
    const candidates = [
      "button[title='Previous Month']",
      "button[title='Next Month']",
      "button[aria-label='Previous Month']",
      "button[aria-label='Next Month']",
      ".slds-datepicker__nav .slds-button_icon",
      ".uiDatePicker .ui-datepicker-prev",
      ".uiDatePicker .ui-datepicker-next"
    ];
    for (const sel of candidates) {
      try {
        const el = wrap.querySelector(sel);
        if (!el) continue;
        if (sel === ".slds-datepicker__nav .slds-button_icon") {
          const btns = wrap.querySelectorAll(".slds-datepicker__nav .slds-button_icon");
          if (btns && btns.length >= 2) {
            if (which === 'prev') { btns[0].click(); return true; }
            else { btns[1].click(); return true; }
          }
        } else {
          if (sel.includes('Previous') && which==='prev') { el.click(); return true; }
          if (sel.includes('Next') && which==='next') { el.click(); return true; }
          el.click(); return true;
        }
      } catch(e) {}
    }
    const arrows = Array.from(wrap.querySelectorAll('button, a'))
                 .filter(n => (n.innerText||n.textContent||'').includes('◀') || (n.innerText||n.textContent||'').includes('▶') || n.className.indexOf('prev')>=0 || n.className.indexOf('next')>=0);
    if (arrows.length) {
      if (which === 'prev') { arrows[0].click(); return true; }
      else { arrows[arrows.length-1].click(); return true; }
    }
    return false;
    """
    try:
        return bool(driver.execute_script(js, which))
    except Exception:
        return False

# ---------- Enhanced click function with proper date filtering ----------
def click_calendar_date_fast(driver, day_number, target_date=None):
    def get_shown_month_index():
        lbl = _sidebar_month_label(driver)
        if lbl:
            y,m = _parse_month_label_to_date(lbl)
            if y and m:
                return y*12 + m
        td = date.today()
        return td.year*12 + td.month

    def collect_candidates():
        js = """
        const day = arguments[0];
        const targetDate = arguments[1]; // YYYY-MM-DD format
        const wrap = document.querySelector('#calendarSidebar') || document.querySelector('.calendarSidebar') || document.getElementById('calendarSidebar');
        if (!wrap) return [];
        const wrapRect = wrap.getBoundingClientRect();
        const nodes = Array.from(wrap.querySelectorAll('*'));
        const out = [];
        for (const n of nodes) {
          try {
            const txt = (n.innerText || n.textContent || '').trim();
            if (txt !== day) continue;
            
            const dataDate = (n.getAttribute('data-date') || '').trim();
            const cls = (n.className || '').toLowerCase();
            const aria = (n.getAttribute('aria-label') || '').trim();
            const title = (n.getAttribute('title') || '').trim();
            const rect = n.getBoundingClientRect();
            
            // Calculate top relative to the calendar wrapper so 'top row' is always smaller
            let relTop = 0;
            if (wrapRect && rect && typeof wrapRect.top === 'number' && typeof rect.top === 'number') {
              relTop = Math.max(0, rect.top - wrapRect.top);
            } else {
              relTop = rect.top || 0;
            }
            
            // Check if this is a disabled/grayed out date
            const isDisabled = cls.includes('disabled') || 
                              cls.includes('slds-disabled') ||
                              cls.includes('slds-disabled-text') ||
                              cls.includes('prevmonth') || 
                              cls.includes('nextmonth') ||
                              cls.includes('outside') ||
                              cls.includes('adjacent') ||
                              cls.includes('other-month') ||
                              n.getAttribute('aria-disabled') === 'true';
            
            // Exact data-date match gets highest priority
            let priority = 0;
            if (targetDate && dataDate === targetDate) {
              priority = 100;
            } else if (targetDate && dataDate && dataDate.includes(targetDate.split('-')[0] + '-' + targetDate.split('-')[1])) {
              priority = 90;
            } else if (!isDisabled) {
              priority = 50;
            } else {
              priority = 10; // Low priority for disabled dates
            }
            
            out.push({
              top: relTop,
              html: n.outerHTML || '',
              cls: cls,
              aria: aria,
              title: title,
              dataDate: dataDate,
              isDisabled: isDisabled,
              priority: priority
            });
          } catch(e) {}
        }
        return out;
        """
        try:
            target_date_str = target_date.strftime("%Y-%m-%d") if target_date else None
            return driver.execute_script(js, str(day_number), target_date_str)
        except Exception:
            return []

    def click_fragment(html_frag):
        js_click = """
        const frag = arguments[0];
        const wrap = document.querySelector('#calendarSidebar') || document.querySelector('.calendarSidebar') || document.getElementById('calendarSidebar');
        if (!wrap) return false;
        const nodes = Array.from(wrap.querySelectorAll('*'));
        for (const n of nodes) {
          try {
            if ((n.outerHTML || '').indexOf(frag) !== -1) {
              let dataDate = n.getAttribute('data-date') || '';
              if (!dataDate) {
                const aria = (n.getAttribute('aria-label')||n.getAttribute('title')||'').trim();
                const m = aria.match(/([A-Za-z]+)\\s+(\\d{1,2}),?\\s*(\\d{4})/);
                if (m) {
                  const mm = new Date(Date.parse(m[1] + ' 1, ' + m[3])).getMonth() + 1;
                  const dd = String(m[2]).padStart(2,'0');
                  const mo = String(mm).padStart(2,'0');
                  dataDate = `${m[3]}-${mo}-${dd}`;
                }
              }
              try { window.__slcm_last_clicked_date = dataDate || ''; } catch(e){}
              n.scrollIntoView({block:'center'});
              try { n.click(); } catch(e) { n.dispatchEvent(new MouseEvent('click',{bubbles:true})); }
              return true;
            }
          } catch(e){}
        }
        return false;
        """
        try:
            return bool(driver.execute_script(js_click, html_frag))
        except Exception:
            return False

    def click_main_grid(target_date):
        js = """
        const day = arguments[0], month = arguments[1], year = arguments[2];
        const patt1 = month + " " + day + ", " + year;
        const patt2 = month + " " + day + " " + year;
        const nodes = Array.from(document.querySelectorAll('*'));
        for (const n of nodes) {
          try {
            const a = (n.getAttribute('aria-label')||'') + '||' + (n.getAttribute('title')||'') + '||' + (n.getAttribute('data-date')||'') + '||' + (n.innerText||n.textContent||'');
            if (a.indexOf(patt1) !== -1 || a.indexOf(patt2) !== -1 || a.indexOf(month + ' ' + day) !== -1) {
              n.scrollIntoView({block:'center'});
              try { n.click(); } catch(e) { n.dispatchEvent(new MouseEvent('click',{bubbles:true})); }
              return true;
            }
          } catch(e){}
        }
        for (const n of nodes) {
          try { 
            const txt = (n.innerText||n.textContent||'').trim();
            const cls = (n.className || '').toLowerCase();
            const isDisabled = cls.includes('disabled') || cls.includes('slds-disabled') || cls.includes('prevmonth') || cls.includes('nextmonth');
            if (txt === String(day) && !isDisabled) { 
              n.scrollIntoView({block:'center'}); 
              try{ n.click(); } catch(e){ n.dispatchEvent(new MouseEvent('click',{bubbles:true})); } 
              return true; 
            } 
          } catch(e){}
        }
        return false;
        """
        try:
            return bool(driver.execute_script(js, target_date.day, target_date.strftime("%B"), target_date.year))
        except Exception:
            return False

    def panel_opened_ok():
        time.sleep(0.5)

        def _has_event_list(elem):
            try:
                if hasattr(elem, "find_element"):
                    try:
                        elem.find_element(By.CSS_SELECTOR, "div.eventList, div.calendarDay, ul.eventListContainer")
                        return True
                    except Exception:
                        return False
                else:
                    return bool(driver.execute_script("""
                        const n = arguments[0];
                        try {
                          if (!n) return false;
                          if (n.querySelector && (n.querySelector('div.eventList') || n.querySelector('div.calendarDay') || n.querySelector('ul.eventListContainer'))) return true;
                        } catch(e){}
                        return false;
                    """, elem))
            except Exception:
                return False

        last_clicked = None
        try:
            last_clicked = driver.execute_script("return (window.__slcm_last_clicked_date || null);")
        except Exception:
            last_clicked = None

        if DEBUG:
            print("panel_opened_ok: last_clicked_date =", repr(last_clicked))

        if last_clicked:
            try:
                panel_candidate = None
                try:
                    panel_candidate = driver.execute_script("""
                        const target = arguments[0];
                        let el = document.querySelector("[data-date='" + target + "']");
                        if (el) {
                          const p = el.closest('div.calendarDay') || el.closest('section') || el.closest('div');
                          return p || el;
                        }
                        const all = Array.from(document.querySelectorAll('div, section, article, h2, a, span'));
                        for (const n of all) {
                          try {
                            const attrs = ((n.getAttribute && (n.getAttribute('aria-description') || n.getAttribute('aria-label') || n.getAttribute('title'))) || '') + ' ' + (n.innerText || n.textContent || '');
                            if (attrs.indexOf(target) !== -1) return n;
                          } catch(e){}
                        }
                        return null;
                    """, last_clicked)
                except Exception:
                    panel_candidate = None

                if DEBUG:
                    try:
                        desc = None
                        if panel_candidate:
                            desc = driver.execute_script("return (arguments[0].getAttribute && (arguments[0].getAttribute('aria-description') || arguments[0].getAttribute('aria-label') || arguments[0].getAttribute('title'))) || (arguments[0].innerText||arguments[0].textContent||'');", panel_candidate)
                        print("panel_opened_ok: panel_candidate (by data-date) present?:", bool(panel_candidate), " sample_text:", repr(desc))
                    except Exception:
                        print("panel_opened_ok: panel_candidate present (could not fetch sample text)")

                if panel_candidate and _has_event_list(panel_candidate):
                    if DEBUG: print("panel_opened_ok: matched panel by data-date and it contains event list -> OK")
                    return True

                if panel_candidate:
                    try:
                        nearby = driver.execute_script("""
                            const n = arguments[0];
                            if (!n) return null;
                            const p = n.closest ? (n.closest('div.calendarDay') || n.closest('section') || n.closest('div')) : null;
                            return p || null;
                        """, panel_candidate)
                        if DEBUG: print("panel_opened_ok: nearby candidate found:", bool(nearby))
                        if nearby and _has_event_list(nearby):
                            if DEBUG: print("panel_opened_ok: nearby contains event list -> OK")
                            return True
                    except Exception:
                        pass

                try:
                    found_via_aria = driver.execute_script("""
                        const target = arguments[0];
                        const nodes = Array.from(document.querySelectorAll('[aria-description], [aria-label], [title]'));
                        for (const n of nodes) {
                          try {
                            const txt = ((n.getAttribute && (n.getAttribute('aria-description') || n.getAttribute('aria-label') || n.getAttribute('title'))) || '') + ' ' + (n.innerText || n.textContent || '');
                            if (txt.indexOf(target) !== -1) return n;
                          } catch(e){}
                        }
                        return null;
                    """, last_clicked)
                    if found_via_aria and _has_event_list(found_via_aria):
                        if DEBUG: print("panel_opened_ok: matched via aria/title and has event list -> OK")
                        return True
                except Exception:
                    pass

            except Exception:
                if DEBUG: print("panel_opened_ok: exception during data-date panel search (falling back)", sys.exc_info()[0])

        try:
            panel = find_day_panel_for_date(driver, target_date)
            if panel:
                if DEBUG: print("panel_opened_ok: found panel using header-based detection")
                try:
                    panel.find_element(By.CSS_SELECTOR, "div.eventList, div.calendarDay, ul.eventListContainer")
                    if DEBUG: print("panel_opened_ok: header-panel contains event list -> OK")
                    return True
                except Exception:
                    if DEBUG: print("panel_opened_ok: header-panel exists but no event list found; accepting panel existence")
                    return True
        except Exception:
            if DEBUG: print("panel_opened_ok: header-based detection raised exception", sys.exc_info()[0])

        try:
            expected = [h.lower() for h in day_header_strings(target_date)]
            headers = driver.find_elements(By.CSS_SELECTOR, "h2, h3, h1")
            for h in headers:
                try:
                    txt = _norm(h.text).lower()
                    for exp in expected:
                        if exp in txt:
                            if DEBUG: print("panel_opened_ok: matched header text (relaxed) ->", txt)
                            try:
                                cand = h.find_element(By.XPATH, "following-sibling::div[contains(@class,'calendarDay') or contains(@class,'eventList')][1]")
                                if cand:
                                    if DEBUG: print("panel_opened_ok: header-following container present -> OK")
                                    return True
                            except Exception:
                                return True
                except Exception:
                    continue
        except Exception:
            if DEBUG: print("panel_opened_ok: relaxed header scan error", sys.exc_info()[0])

        if DEBUG: print("panel_opened_ok: no matching panel detected for", target_date)
        return False

    # If no target_date, attempt a simple direct click with disabled filtering
    if target_date is None:
        js_direct = """
        const wrap = document.querySelector('#calendarSidebar') || document.querySelector('.calendarSidebar') || document.getElementById('calendarSidebar');
        if (!wrap) return false;
        const nodes = wrap.querySelectorAll('table.datepicker .slds-day, .slds-day, table.datepicker td, table td, td');
        for (const n of nodes) {
            const txt = (n.textContent || '').trim();
            const cls = (n.className || '').toLowerCase();
            const disabled = n.getAttribute && (
                n.getAttribute('aria-disabled') === 'true' ||
                cls.includes('disabled') || 
                cls.includes('slds-disabled') ||
                cls.includes('prevmonth') || 
                cls.includes('nextmonth') ||
                cls.includes('outside') ||
                cls.includes('adjacent')
            );
            if (!disabled && txt === arguments[0]) { 
                n.scrollIntoView({block:'center'}); 
                try{ n.click(); } catch(e){} 
                return true; 
            }
        }
        return false;
        """
        try:
            ok = driver.execute_script(js_direct, str(day_number))
        except Exception:
            ok = False
        if not ok:
            raise RuntimeError(f"❌ Could not click mini calendar date {day_number}")
        return

        # ---------- SAFE QUICK PRE-CHECK (avoid ambiguous sidebar clicks; jump to robust) ----------
    try:
        pre_cands = collect_candidates()
    except Exception:
        pre_cands = []

    # Helper: candidate unambiguous if it has exact dataDate OR month/year in aria/title
    def unambiguous_for_target(c):
        dd = (c.get('dataDate') or "").strip()
        if dd and target_date and dd == target_date.strftime("%Y-%m-%d"):
            return True
        txt = ((c.get('aria') or "") + " " + (c.get('title') or "")).strip()
        if txt:
            mon = target_date.strftime("%B") if target_date else ""
            yr = str(target_date.year) if target_date else ""
            if mon and mon.lower() in txt.lower():
                return True
            if yr and yr in txt:
                return True
        return False

    # 1) Exact data-date in sidebar candidates (highest confidence)
    if pre_cands:
        exact_now = [c for c in pre_cands if c.get('dataDate') and target_date and c.get('dataDate') == target_date.strftime("%Y-%m-%d")]
        if exact_now:
            pick = sorted(exact_now, key=lambda c: (c.get('top', 0)))[0]
            if DEBUG: print("quick-precheck: clicking exact data-date candidate (no nav needed)")
            if click_fragment(pick.get('html')) and panel_opened_ok():
                if DEBUG: print("✅ Quick precheck exact click succeeded")
                return

    # 2) Try global data-date anywhere on page (very reliable)
    try:
        if target_date:
            target_dd = target_date.strftime("%Y-%m-%d")
            if DEBUG: print("quick-precheck: looking for global data-date element", target_dd)
            clicked_global = driver.execute_script("""
                const t = arguments[0];
                const el = document.querySelector("[data-date='" + t + "']");
                if (!el) return false;
                try { el.scrollIntoView({block:'center'}); } catch(e){}
                try { el.click(); } catch(e){ el.dispatchEvent(new MouseEvent('click',{bubbles:true})); }
                return true;
            """, target_dd)
            if clicked_global:
                time.sleep(0.12)
                if panel_opened_ok():
                    if DEBUG: print("✅ Quick precheck global data-date click succeeded")
                    return
                # else continue to robust path if it didn't open correct panel
    except Exception:
        pass

    # 3) If sidebar has only ambiguous candidates (no dataDate and no aria/title month/year),
    #    DO NOT click them — jump straight to robust navigation which is reliable.
    if pre_cands:
        non_disabled_now = [c for c in pre_cands if not c.get('isDisabled', False)]
        if non_disabled_now:
            unamb = [c for c in non_disabled_now if unambiguous_for_target(c)]
            ambiguous = [c for c in non_disabled_now if not unambiguous_for_target(c)]
            if DEBUG and ambiguous:
                sample = []
                for c in ambiguous[:6]:
                    sample.append({
                        'cls': c.get('cls'),
                        'dataDate': c.get('dataDate'),
                        'aria': (c.get('aria') or '')[:60],
                        'top': c.get('top')
                    })
                print("quick-precheck: skipped ambiguous sidebar candidates:", sample)

            # If we have unambiguous candidates, try them (safe).
            if unamb:
                pick = sorted(unamb, key=lambda c: (c.get('top', 0)))[0]
                if DEBUG: print("quick-precheck: clicking unambiguous same-month candidate, cls=", pick.get('cls'))
                if click_fragment(pick.get('html')) and panel_opened_ok():
                    if DEBUG: print("✅ Quick precheck unambiguous click succeeded")
                    return
            else:
                # No safe candidate available — use robust navigation immediately to avoid wasted clicks
                if DEBUG: print("quick-precheck: no safe sidebar candidate found — invoking robust navigator")
                try:
                    click_calendar_date_robust(driver, target_date)
                    return
                except Exception as e:
                    if DEBUG: print("quick-precheck: robust navigator raised:", repr(e))
                    # fall through to the normal nav logic below (it will try again)
    # If we reach here, no safe quick precheck succeeded — proceed to month navigation as before.

    # navigate to target month (bounded)
    MAX_NAV = 24
    target_month_index = target_date.year*12 + target_date.month
    shown_month_index = get_shown_month_index()
    nav_count = 0
    while shown_month_index != target_month_index and nav_count < MAX_NAV:
        direction = 'next' if target_month_index > shown_month_index else 'prev'
        if DEBUG: print("nav step: shown", shown_month_index, " target", target_month_index, " ->", direction)
        if not _click_sidebar_prev_next_once(driver, direction):
            if DEBUG: print("nav control not found for", direction)
            break
        time.sleep(0.10)
        nav_count += 1
        shown_month_index = get_shown_month_index()

    # collect candidates with enhanced filtering
    candidates = collect_candidates()
    if not candidates:
        raise RuntimeError(f"❌ No mini-calendar candidates found for day {day_number}")

    candidates_sorted = sorted(candidates, key=lambda c: (-c.get('priority', 0), c.get('top', 0)))

    click_order = []
    exact_matches = [c for c in candidates_sorted if c.get('priority', 0) >= 90]
    if exact_matches:
        click_order.extend(exact_matches)
    non_disabled = [c for c in candidates_sorted if not c.get('isDisabled', False) and c not in click_order]
    if non_disabled:
        click_order.extend(non_disabled)
    remaining = [c for c in candidates_sorted if c not in click_order]
    click_order.extend(remaining)

    MAX_TRIES = max(6, len(click_order)*2)
    tried = 0

    for idx, candidate in enumerate(click_order):
        if tried >= MAX_TRIES:
            break
        tried += 1
        if DEBUG:
            lbl = _sidebar_month_label(driver)
            print(f"[click attempt] trying day {day_number} (label={lbl}) candidate #{idx+1} priority={candidate.get('priority')} disabled={candidate.get('isDisabled')} cls={candidate.get('cls')}")
        clicked = click_fragment(candidate.get('html'))
        if not clicked:
            if DEBUG: print("candidate click failed, continuing")
            time.sleep(0.10)
            continue
        if panel_opened_ok():
            if DEBUG: print(f"✅ Clicked calendar date: {day_number} (candidate #{idx+1})")
            return
        else:
            if DEBUG: print("⚠️ Clicked but wrong panel — will try next candidate")
            time.sleep(0.10)
            continue

    # fallback to main-grid with enhanced filtering
    if DEBUG: print("Attempting fallback click on main calendar grid for", target_date)
    if click_main_grid(target_date):
        if panel_opened_ok():
            if DEBUG: print("✅ Main-grid fallback click opened correct day panel")
            return
        if DEBUG: print("Fallback main-grid click did not open correct panel")

    # final nudge and reattempt
    if DEBUG: print("Final attempt: nudging month then re-collecting candidates")
    _click_sidebar_prev_next_once(driver, 'next'); time.sleep(0.10)
    _click_sidebar_prev_next_once(driver, 'prev'); time.sleep(0.10)
    candidates = collect_candidates()
    candidates_sorted = sorted(candidates, key=lambda c: (-c.get('priority', 0), c.get('top', 0)))
    for c in candidates_sorted:
        if not c.get('isDisabled', False) and click_fragment(c.get('html')) and panel_opened_ok():
            if DEBUG: print("✅ Clicked after nudge")
            return

    raise RuntimeError(f"❌ Could not click mini calendar date {day_number} (last known label={repr(_sidebar_month_label(driver))}) after retries")

# =========================================================
# Alternative robust approach for problematic calendars
# =========================================================
def click_calendar_date_robust(driver, target_date):
    def get_shown_month_index():
        lbl = _sidebar_month_label(driver)
        if lbl:
            y,m = _parse_month_label_to_date(lbl)
            if y and m:
                return y*12 + m
        td = date.today()
        return td.year*12 + td.month

    current_month_index = get_shown_month_index()
    target_month_index = target_date.year * 12 + target_date.month

    max_nav_attempts = 24
    nav_count = 0
    while current_month_index != target_month_index and nav_count < max_nav_attempts:
        direction = 'next' if target_month_index > current_month_index else 'prev'
        if DEBUG: print(f"Navigating {direction} from month index {current_month_index} to {target_month_index}")
        if not _click_sidebar_prev_next_once(driver, direction):
            if DEBUG: print(f"Navigation control not found for {direction}")
            break
        time.sleep(0.12)
        nav_count += 1
        current_month_index = get_shown_month_index()

    if current_month_index != target_month_index and DEBUG:
        print(f"Warning: Could not navigate to target month. Current: {current_month_index}, Target: {target_month_index}")

    js_click_current_month_date = """
    const day = arguments[0];
    const wrap = document.querySelector('#calendarSidebar') || document.querySelector('.calendarSidebar');
    if (!wrap) return false;
    const cells = wrap.querySelectorAll('td, .slds-day, button, a');
    for (const cell of cells) {
        const text = (cell.innerText || cell.textContent || '').trim();
        const classes = (cell.className || '').toLowerCase();
        if (text === day && 
            !classes.includes('disabled') && 
            !classes.includes('slds-disabled') &&
            !classes.includes('slds-disabled-text') &&
            !classes.includes('prevmonth') &&
            !classes.includes('nextmonth') &&
            !classes.includes('outside') &&
            !classes.includes('adjacent') &&
            !classes.includes('other-month') &&
            cell.getAttribute('aria-disabled') !== 'true') {
            cell.scrollIntoView({block: 'center'});
            try { cell.click(); } catch(e) { cell.dispatchEvent(new MouseEvent('click', {bubbles: true})); }
            return true;
        }
    }
    return false;
    """

    day_str = str(target_date.day)
    if driver.execute_script(js_click_current_month_date, day_str):
        if DEBUG: print(f"✅ Successfully clicked date {target_date} using robust method")
        return True

    raise RuntimeError(f"❌ Could not click date {target_date} after navigating to correct month")

# =========================================================
# Remaining helpers (events scanning, attendance manipulation)
# =========================================================
def _norm(s):
    return " ".join((s or "").split())

def _has_word(haystack, needle):
    return re.search(rf'\b{re.escape(needle)}\b', haystack) is not None

def matches_event_text(txt, code, sem, sec, sess):
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
            pat = rf'(?<![A-Z0-9]){re.escape(secU)}(?![A-Z0-9])'
            ok = ok and re.search(pat, T) is not None
        else:
            pat1 = rf'\bSEC(?:TION)?\s*[:\-]?\s*{re.escape(secU)}(?!\s*-\s*\d+)\b'
            pat2 = rf'\(\s*{re.escape(secU)}\s*\)'
            pat3 = rf'(?<![A-Z0-9]){re.escape(secU)}(?!\s*-\s*\d+)(?![A-Z0-9])'
            ok = ok and (
                re.search(pat1, T) is not None or
                re.search(pat2, T) is not None or
                re.search(pat3, T) is not None
            )

    if ok and sess and sess.strip() and (not sec or "-" not in sec.strip()):
        s = sess.strip().upper()
        ok = ok and (f"SESSION {s}" in T)

    return ok

def day_header_strings(d):
    parts = []
    try:
        if sys.platform != "win32":
            parts.append(d.strftime("%A, %B %-d"))
        else:
            parts.append(d.strftime("%A, %B %#d"))
    except Exception:
        parts.append(d.strftime("%A, %B %d").lstrip("0").replace(", 0", ", "))
    parts.append(d.strftime("%A, %B %d").lstrip("0").replace(", 0", ", "))
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

def aria_date_matches_selected(aria_desc, selected_date):
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

    parsed, err = parse_subject_details(subject_details_str)
    if err:
        print(f"❌ Invalid subject details: {err}")
        print(f"   Received: {subject_details_str!r}")
        sys.exit(1)
    course_name, course_code, semester, class_section, session_no = parsed

    print(f"📅 Selected Date : {selected_date}")
    print(f"📂 Workbook      : {workbook_path}")
    print(f"🧑‍🎓 Absentees   : {', '.join(absentees) if absentees else 'None'}")
    print("\n📘 Course Details")
    print(f"   Course Name   : {course_name or '(blank)'}")
    print(f"   Course Code   : {course_code or '(blank)'}")
    print(f"   Semester      : {semester or '(blank)'}")
    print(f"   Class Section : {class_section or '(blank)'}")
    print(f"   Session No    : {session_no or '(none)'}")

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
        opts.add_argument("--log-level=3")
        opts.add_argument("--disable-logging")
        opts.add_argument("--disable-dev-shm-usage")
        opts.add_argument("--no-sandbox")
        opts.add_argument("--disable-background-timer-throttling")
        opts.add_argument("--disable-backgrounding-occluded-windows")
        opts.add_argument("--disable-renderer-backgrounding")
        opts.add_argument("--disable-features=TranslateUI,VizDisplayCompositor")
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
        # Navigate & login
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

        # Wait for mini calendar & click selected date (robust with fallbacks)
        WebDriverWait(driver, 12).until(EC.presence_of_element_located((By.ID, "calendarSidebar")))
        time.sleep(0.25)
        day_number = str(selected_date.day).lstrip("0")

        def try_wait_panel(timeout=PANEL_READY_TIMEOUT):
            p = wait_for_day_panel_ready(driver, selected_date, timeout=timeout)
            if p:
                if DEBUG: print("wait_for_day_panel_ready: panel found")
            else:
                if DEBUG: print("wait_for_day_panel_ready: panel NOT found")
            return p

        # Primary attempt: fast click then wait
        try:
            if DEBUG: print("Attempting click_calendar_date_fast(...)")
            click_calendar_date_fast(driver, day_number, target_date=selected_date)
        except Exception as e_fast:
            if DEBUG:
                print("click_calendar_date_fast raised:", repr(e_fast))

        time.sleep(AFTER_DATE_CLICK_PAUSE)
        panel = try_wait_panel(timeout=PANEL_READY_TIMEOUT)

        # Fallback 1: use the robust navigation+click method and wait again
        if not panel:
            try:
                if DEBUG: print("Fallback: click_calendar_date_robust(...)")
                ok = click_calendar_date_robust(driver, selected_date)
                if ok:
                    time.sleep(AFTER_DATE_CLICK_PAUSE)
                panel = try_wait_panel(timeout=PANEL_READY_TIMEOUT)
            except Exception as e_rob:
                if DEBUG: print("click_calendar_date_robust raised:", repr(e_rob))
                panel = None

        # Fallback 2: Try to click any element with exact data-date anywhere on the page
        if not panel:
            try:
                target_dd = selected_date.strftime("%Y-%m-%d")
                if DEBUG: print("Fallback: clicking any element with data-date=", target_dd)
                clicked = driver.execute_script("""
                    const target = arguments[0];
                    const el = document.querySelector("[data-date='" + target + "']");
                    if (!el) {
                      const all = Array.from(document.querySelectorAll('[data-date], [data-dt], [data-day]'));
                      for (const n of all) {
                        const v = (n.getAttribute('data-date')||n.getAttribute('data-dt')||n.getAttribute('data-day')||'').trim();
                        if (v === target) { n.scrollIntoView({block:'center'}); try{ n.click(); }catch(e){ n.dispatchEvent(new MouseEvent('click',{bubbles:true})); } return true; }
                      }
                      return false;
                    }
                    el.scrollIntoView({block:'center'});
                    try { el.click(); } catch(e) { el.dispatchEvent(new MouseEvent('click',{bubbles:true})); }
                    return true;
                """, target_dd)
                if DEBUG: print("Fallback data-date click returned:", bool(clicked))
                time.sleep(AFTER_DATE_CLICK_PAUSE)
                panel = try_wait_panel(timeout=PANEL_READY_TIMEOUT)
            except Exception as e:
                if DEBUG: print("data-date fallback raised:", repr(e))
                panel = None

        # Fallback 3: try a sidebar JS click for visible, non-disabled cell matching the day text
        if not panel:
            try:
                if DEBUG: print("Fallback: clicking visible non-disabled cell inside calendarSidebar with day text")
                clicked = driver.execute_script("""
                    const day = arguments[0].trim();
                    const wrap = document.querySelector('#calendarSidebar') || document.querySelector('.calendarSidebar');
                    if (!wrap) return false;
                    const cells = Array.from(wrap.querySelectorAll('td, button, a, div, span'));
                    for (const c of cells) {
                        try {
                            const txt = (c.innerText || c.textContent || '').trim();
                            const cls = (c.className||'').toLowerCase();
                            const disabled = cls.includes('disabled') || cls.includes('prevmonth') || cls.includes('nextmonth') || (c.getAttribute && c.getAttribute('aria-disabled') === 'true');
                            if (!disabled && txt === day) {
                                c.scrollIntoView({block:'center'});
                                try { c.click(); } catch(e) { c.dispatchEvent(new MouseEvent('click',{bubbles:true})); }
                                return true;
                            }
                        } catch(e){}
                    }
                    return false;
                """, day_number)
                if DEBUG: print("Sidebar visible-day click returned:", bool(clicked))
                time.sleep(AFTER_DATE_CLICK_PAUSE)
                panel = try_wait_panel(timeout=PANEL_READY_TIMEOUT)
            except Exception as e:
                if DEBUG: print("sidebar fallback raised:", repr(e))
                panel = None

        # Final diagnostic dump and error if still not found
        if not panel:
            try:
                data_dates = driver.execute_script("""
                    const wrap = document.querySelector('#calendarSidebar') || document.querySelector('.calendarSidebar');
                    if (!wrap) return [];
                    const nodes = Array.from(wrap.querySelectorAll('[data-date]'));
                    const out = [];
                    for (const n of nodes) {
                      try { out.push(n.getAttribute('data-date')); } catch(e) {}
                    }
                    return out;
                """)
            except Exception:
                data_dates = None
            try:
                last_clicked_marker = driver.execute_script("return (window.__slcm_last_clicked_date || null);")
            except Exception:
                last_clicked_marker = None

            if DEBUG:
                print("❗ Diagnostic info: calendarSidebar [data-date] attrs:", data_dates)
                print("❗ Diagnostic info: window.__slcm_last_clicked_date =", repr(last_clicked_marker))
                print("❗ Diagnostic info: visible mini-calendar label:", repr(_sidebar_month_label(driver)))

            raise RuntimeError(
                "❌ Could not locate the correct day panel for the selected date after multiple fallbacks.\n"
                f"Diagnostic: data-dates-in-sidebar={data_dates}, last_clicked={last_clicked_marker}, last_label={_sidebar_month_label(driver)}"
            )

        if not wait_for_events_to_settle(driver, panel, timeout=EVENT_SETTLE_TIMEOUT):
            print("⚠️ Event list still changing; proceeding with search anyway.")

        # find event tile and open it
        target = scroll_day_panel_gradual(
            driver, panel, max_seconds=EVENT_SEARCH_TIMEOUT,
            code=course_code, sem=semester, sec=class_section,
            sess=(session_no if "-" not in class_section.strip() else None),
            selected_date=selected_date
        )
        if not target:
            raise RuntimeError("❌ Could not locate the event tile for the selected date (after scrolling & waiting).")

        try:
            driver.execute_script("arguments[0].scrollIntoView({block:'center'});", target)
            time.sleep(0.12)
            target.click()
            print("✅ Opened the event tile")
        except Exception as e:
            raise RuntimeError(f"❌ Failed to click the event tile: {e}")

        # "More Details" if present
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
        # Process absentees (UNTICK boxes)
        # =============================
        print(f"🔎 Processing attendance for {len(absentees)} absentees…")
        unticked_ids, not_found = [], []

        def get_table_container():
            try:
                return WebDriverWait(driver, 4).until(
                    EC.presence_of_element_located((By.XPATH, "//div[contains(@class,'slds-table') or contains(@class,'slds-scrollable')]//table/ancestor::div[contains(@class,'scroll') or contains(@class,'slds-scrollable')][1]"))
                )
            except Exception:
                try:
                    return driver.find_element(By.XPATH, "//div[contains(@class,'slds-scrollable')]")
                except Exception:
                    return None

        def try_find_cell_for_id(student_id):
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

            try:
                cell = driver.execute_script("""
                    const sid = arguments[0].trim();
                    const nodes = Array.from(document.querySelectorAll("lightning-base-formatted-text, td, span"));
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
            t0 = time.time()
            while time.time() - t0 < PER_STUDENT_MAX_SECONDS:
                try:
                    cell = try_find_cell_for_id(student_id)
                    if not cell:
                        scroll_table_small_steps()
                        continue

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
                    continue
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
