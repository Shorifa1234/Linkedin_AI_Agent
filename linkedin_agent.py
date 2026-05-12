# -*- coding: utf-8 -*-
# linkedin_agent.py  —  Unified Orchestrator
#
# PIPELINE:
#   STEP 1  → LinkedIn search → collect company URLs
#   STEP 2  → Company About + People page scrape
#   STEP 3  → Google / Bing search to find individual LinkedIn profile URLs
#   STEP 4  → Deduplicate + final output
#
# FEATURES:
#   ✅ Single script — no manual chaining
#   ✅ JSON checkpoint — resume from any step after crash/restart
#   ✅ Auto-save after each company / each row
#   ✅ Director-level filtering (CEO, VP, Director, Manager, etc.)
#   ✅ Google CAPTCHA → auto-switch to Bing fallback
#   ✅ Multiple CSS selector fallbacks for LinkedIn UI changes
#   ✅ Chrome crash auto-restart

import sys
import time
import re
import os
import json
import random
from urllib.parse import quote_plus, urlparse, parse_qs, unquote, urlsplit, urlunsplit, parse_qsl, urlencode

# Fix Windows console encoding so print() does not crash on special chars
if sys.stdout.encoding and sys.stdout.encoding.lower() != "utf-8":
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
if sys.stderr.encoding and sys.stderr.encoding.lower() != "utf-8":
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")

import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (
    TimeoutException, NoSuchWindowException, WebDriverException
)
from webdriver_manager.chrome import ChromeDriverManager


# ============================================================
#  CONFIG  — edit these before running
# ============================================================

LINKEDIN_SEARCH_URL = (
    "https://www.linkedin.com/search/results/companies/"
    "?companyHqGeo=%5B%22105080838%22%5D"
    "&keywords=Pet%20Spay%20%26%20Neuter%20Programs"
    "&origin=GLOBAL_SEARCH_HEADER"
)
TOTAL_SEARCH_PAGES = 5           # how many LinkedIn search result pages to scrape

# Director-level keywords — contacts NOT matching are kept but flagged
DIRECTOR_KEYWORDS = [
    "ceo", "coo", "cfo", "cto", "cmo", "founder", "co-founder",
    "president", "owner", "partner", "principal",
    "director", "vp", "vice president",
    "manager", "head of", "chief",
    "executive", "superintendent", "administrator",
]

PEOPLE_SCROLL_ROUNDS  = 8        # scroll rounds on /people page
MAX_PAGES_PER_COMPANY = 200      # pagination cap for employee listing

WAIT_SEC        = 20
POLITE_DELAY    = 1.3
LOOP_DELAY      = (9, 16)        # between Google/Bing search rows
PAGE_WAIT_RANGE = (5, 9)
SAVE_EVERY      = 10             # auto-save every N profile-search rows

# Credentials file — first line: email, second line: password
CREDENTIALS_FILE = "linkedin.txt"

# File paths
STATE_FILE       = "agent_state.json"
COMPANIES_FILE   = "agent_companies.xlsx"
PEOPLE_FILE      = "agent_people.xlsx"
PROFILES_FILE    = "agent_profiles.xlsx"
VERIFIED_FILE    = "agent_verified.xlsx"
FINAL_FILE       = "agent_final.xlsx"

# Google CAPTCHA → switch to Bing after this many consecutive CAPTCHA hits
CAPTCHA_SWITCH_THRESHOLD = 3


# ============================================================
#  CREDENTIALS
# ============================================================

def load_credentials() -> tuple[str, str]:
    """Read email + password from first two lines of linkedin.txt."""
    if not os.path.exists(CREDENTIALS_FILE):
        return "", ""
    lines = []
    with open(CREDENTIALS_FILE, "r", encoding="utf-8", errors="ignore") as f:
        for line in f:
            stripped = line.strip()
            if stripped:
                lines.append(stripped)
            if len(lines) == 2:
                break
    if len(lines) >= 2:
        return lines[0], lines[1]
    return "", ""


def linkedin_auto_login():
    """Attempt automatic LinkedIn login using credentials from linkedin.txt."""
    email, password = load_credentials()
    if not email or not password:
        print("[LOGIN] No credentials found in linkedin.txt — manual login required.")
        input("Login manually in Chrome, then press Enter...")
        return

    print(f"[LOGIN] Auto-login as {email} ...")
    _driver.get("https://www.linkedin.com/login")
    time.sleep(2)

    try:
        email_field = _wait.until(EC.presence_of_element_located((By.ID, "username")))
        email_field.clear()
        email_field.send_keys(email)
        time.sleep(0.5)

        pass_field = _driver.find_element(By.ID, "password")
        pass_field.clear()
        pass_field.send_keys(password)
        time.sleep(0.5)

        _driver.find_element(By.CSS_SELECTOR, "button[type='submit']").click()
        time.sleep(4)

        # Check if login succeeded
        cur = _driver.current_url.lower()
        if "feed" in cur or "mynetwork" in cur or "jobs" in cur:
            print("[LOGIN] Auto-login successful.")
        elif "checkpoint" in cur or "challenge" in cur or "verify" in cur:
            print("[LOGIN] LinkedIn verification/CAPTCHA required.")
            input("Complete verification in Chrome, then press Enter...")
        else:
            print(f"[LOGIN] Unexpected page after login: {_driver.current_url}")
            print("[LOGIN] If not logged in, do it manually then press Enter...")
            input()
    except Exception as e:
        print(f"[LOGIN] Auto-login failed ({e}). Login manually then press Enter...")
        input()


# ============================================================
#  STATE MANAGEMENT
# ============================================================

DEFAULT_STATE = {
    "step": 1,          # 1=collect companies, 2=scrape people, 3=find profiles, 4=dedup
    "google_captcha_streak": 0,
    "use_bing": False,
}

def load_state() -> dict:
    if os.path.exists(STATE_FILE):
        with open(STATE_FILE, "r", encoding="utf-8") as f:
            s = json.load(f)
        print(f"[AGENT] Resuming from state: step={s.get('step')}, bing={s.get('use_bing')}")
        return s
    return dict(DEFAULT_STATE)

def save_state(state: dict):
    with open(STATE_FILE, "w", encoding="utf-8") as f:
        json.dump(state, f, indent=2)


# ============================================================
#  DRIVER
# ============================================================

_driver: webdriver.Chrome | None = None
_wait:   WebDriverWait | None    = None

def ensure_driver():
    global _driver, _wait
    try:
        if _driver is not None:
            _ = _driver.current_url
            return
    except Exception:
        pass

    options = Options()
    options.add_argument("--start-maximized")
    options.add_argument("--disable-blink-features=AutomationControlled")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--disable-gpu")
    options.add_argument("--window-size=1920,1080")
    options.add_argument("--remote-debugging-port=9222")
    options.add_argument(
        "user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
    )

    _driver = webdriver.Chrome(
        service=Service(ChromeDriverManager().install()),
        options=options
    )
    _driver.set_page_load_timeout(60)
    _wait = WebDriverWait(_driver, WAIT_SEC)
    print("[DRIVER] Chrome started.")

def quit_driver():
    global _driver, _wait
    try:
        if _driver:
            _driver.quit()
    except Exception:
        pass
    _driver = None
    _wait = None


# ============================================================
#  HELPERS
# ============================================================

def human_sleep(a, b):
    time.sleep(random.uniform(a, b))

def text_clean(s) -> str:
    return re.sub(r"\s+", " ", (str(s) if s else "")).strip()

def norm(s) -> str:
    return text_clean(s).lower()

def clean_profile_url(url: str) -> str:
    if not url:
        return ""
    url = url.strip()
    parts = urlsplit(url)
    q = dict(parse_qsl(parts.query))
    for k in list(q.keys()):
        if k.lower() in ["miniprofileurn", "lipi", "trk", "trkinfo", "originalreferer"]:
            q.pop(k, None)
    clean = urlunsplit((parts.scheme, parts.netloc, parts.path, urlencode(q, doseq=True), ""))
    return clean.rstrip("/")

def is_director_level(designation: str) -> bool:
    d = norm(designation)
    return any(kw in d for kw in DIRECTOR_KEYWORDS)

def is_linkedin_member(name: str) -> bool:
    return norm(name) in ["linkedin member", "member", "linkedin", ""]

def parse_address_parts(addr: str) -> dict:
    out = {"Street Address": "", "Street Address 02": "", "City": "", "State": "", "Zip Code": "", "Country": ""}
    if not addr:
        return out
    s = text_clean(addr)
    m = re.search(r"^(.*?),\s*(.*?),\s*([^,]+),\s*([A-Za-z\s]+)\s+(\d{5}(?:-\d{4})?)\s*,\s*(.*)$", s)
    if m:
        out["Street Address"]    = m.group(1).strip()
        out["Street Address 02"] = m.group(2).strip()
        out["City"]              = m.group(3).strip()
        out["State"]             = m.group(4).strip()
        out["Zip Code"]          = m.group(5).strip()
        out["Country"]           = m.group(6).strip()
        return out
    parts = [p.strip() for p in s.split(",") if p.strip()]
    if parts:     out["Street Address"]    = parts[0]
    if len(parts) >= 2: out["Street Address 02"] = parts[1]
    if len(parts) >= 3: out["City"]              = parts[-3]
    if len(parts) >= 2: out["Country"]           = parts[-1]
    mz = re.search(r"\b(\d{5}(?:-\d{4})?)\b", s)
    if mz:        out["Zip Code"] = mz.group(1)
    return out


# ============================================================
#  STEP 1 — COLLECT COMPANY URLs FROM LINKEDIN SEARCH
# ============================================================

# Multiple selector fallbacks because LinkedIn changes class names
_COMPANY_LINK_SELECTORS = [
    "a[href*='linkedin.com/company']",
    "a.app-aware-link[href*='/company/']",
    ".search-results__list a[href*='/company/']",
    "li.reusable-search__result-container a[href*='/company/']",
]

def step1_collect_companies(state: dict):
    print("\n========== STEP 1: Collect Company URLs ==========")
    if os.path.exists(COMPANIES_FILE):
        df = pd.read_excel(COMPANIES_FILE)
        if len(df) > 0:
            print(f"[STEP1] Already have {len(df)} companies in {COMPANIES_FILE}. Skipping.")
            return

    ensure_driver()
    linkedin_auto_login()

    company_urls: list[str] = []

    for page in range(1, TOTAL_SEARCH_PAGES + 1):
        print(f"  Scraping search page {page}...")
        _driver.get(LINKEDIN_SEARCH_URL + f"&page={page}")
        time.sleep(3)

        found = []
        for sel in _COMPANY_LINK_SELECTORS:
            try:
                elements = _driver.find_elements(By.CSS_SELECTOR, sel)
                if elements:
                    found = elements
                    break
            except Exception:
                continue

        for elem in found:
            url = elem.get_attribute("href") or ""
            url = url.split("?")[0].rstrip("/")
            if url and url not in company_urls:
                company_urls.append(url)

        print(f"  → {len(found)} companies found on page {page}")
        time.sleep(2)

    df = pd.DataFrame(company_urls, columns=["LinkedIn URL"])
    df.drop_duplicates(subset="LinkedIn URL", inplace=True)
    df.to_excel(COMPANIES_FILE, index=False)
    print(f"[STEP1] Saved {len(df)} company URLs → {COMPANIES_FILE}")

    state["step"] = 2
    save_state(state)


# ============================================================
#  STEP 2 — SCRAPE COMPANY ABOUT + PEOPLE
# ============================================================

def _get_about_field(label: str) -> str:
    xpaths = [
        f"//dl//dt[.//h3[normalize-space()='{label}']]/following-sibling::dd[1]",
        f"//dt[.//h3[normalize-space()='{label}']]/following-sibling::dd[1]",
        f"//h3[normalize-space()='{label}']/ancestor::dt/following-sibling::dd[1]",
        f"//span[normalize-space()='{label}']/following-sibling::span[1]",
    ]
    for xp in xpaths:
        try:
            el = _driver.find_element(By.XPATH, xp)
            t = text_clean(el.text)
            if t:
                return t
        except Exception:
            pass
    return ""

def _get_company_name() -> str:
    selectors = [
        "h1.org-top-card-summary__title",
        "h1.org-top-card-summary__title span",
        "h1",
    ]
    for sel in selectors:
        try:
            el = _driver.find_element(By.CSS_SELECTOR, sel)
            t = text_clean(el.text)
            if t:
                return t
        except Exception:
            pass
    return ""

def _get_company_address() -> str:
    css_candidates = [
        "p.t-14.t-black--light.t-normal.break-words",
        "p.break-words",
        ".org-top-card-summary-info-list__info-item",
    ]
    for sel in css_candidates:
        for p in _driver.find_elements(By.CSS_SELECTOR, sel):
            t = text_clean(p.text)
            if t and "," in t and re.search(r"\d", t):
                return t
    return ""

def _scroll_people_page():
    for _ in range(PEOPLE_SCROLL_ROUNDS):
        _driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        time.sleep(1.1)
        for btn_sel in [
            "button.scaffold-finite-scroll__load-button",
            "button[aria-label*='Show more']",
        ]:
            try:
                btn = _driver.find_element(By.CSS_SELECTOR, btn_sel)
                if btn.is_displayed() and btn.is_enabled():
                    btn.click()
                    time.sleep(1.2)
            except Exception:
                pass

def _scrape_people_cards() -> list[dict]:
    people = []
    card_selectors = [
        "li.org-people-profile-card__profile-card-spacing",
        "li[class*='profile-card']",
        "[role='listitem']",
    ]
    cards = []
    for sel in card_selectors:
        cards = _driver.find_elements(By.CSS_SELECTOR, sel)
        if cards:
            break

    for c in cards:
        try:
            profile_url = ""
            for a_sel in ["a[href*='/in/']", "a[href*='miniProfileUrn=']"]:
                els = c.find_elements(By.CSS_SELECTOR, a_sel)
                if els:
                    profile_url = clean_profile_url(els[0].get_attribute("href") or "")
                    break

            name = ""
            for name_sel in [
                ".artdeco-entity-lockup__title .lt-line-clamp--single-line",
                ".artdeco-entity-lockup__title",
                "span.lt-line-clamp--single-line",
            ]:
                try:
                    name = text_clean(c.find_element(By.CSS_SELECTOR, name_sel).text)
                    if name:
                        break
                except Exception:
                    pass
            if not name:
                name = "LinkedIn Member"

            designation = ""
            for des_sel in [
                ".artdeco-entity-lockup__subtitle .lt-line-clamp--multi-line",
                "div.lt-line-clamp--multi-line",
                ".artdeco-entity-lockup__subtitle",
            ]:
                try:
                    designation = text_clean(c.find_element(By.CSS_SELECTOR, des_sel).text)
                    if designation:
                        break
                except Exception:
                    pass

            people.append({
                "Name": name,
                "Designation": designation,
                "Director Level": "Yes" if is_director_level(designation) else "No",
                "Contact Linkedin URL": profile_url,
                "Personal Email": "",
            })
        except Exception:
            pass

    # de-dupe
    seen = {}
    for p in people:
        key = p["Contact Linkedin URL"] or (p["Name"] + "|" + p["Designation"])
        if key not in seen:
            seen[key] = p
    return list(seen.values())

def step2_scrape_people(state: dict):
    print("\n========== STEP 2: Scrape Company About + People ==========")
    df_companies = pd.read_excel(COMPANIES_FILE)

    if os.path.exists(PEOPLE_FILE):
        df_out = pd.read_excel(PEOPLE_FILE)
        done_urls = set(df_out["LinkedIn URL"].dropna().unique())
        rows = df_out.to_dict("records")
        print(f"[STEP2] Resuming — {len(done_urls)} companies already done.")
    else:
        df_out = None
        done_urls = set()
        rows = []

    ensure_driver()

    for _, r in df_companies.iterrows():
        company_base = str(r.get("LinkedIn URL", "")).strip().rstrip("/")
        if not company_base or company_base.lower() == "nan":
            continue
        if company_base in done_urls:
            continue

        about_url  = company_base + "/about/"
        people_url = company_base + "/people/"

        company_name = website = phone = industry = num_emp = ""
        address_parts = parse_address_parts("")

        # -- About --
        try:
            _driver.get(about_url)
            _wait.until(EC.presence_of_element_located((By.TAG_NAME, "body")))
            time.sleep(2)
            company_name  = _get_company_name()
            addr_full     = _get_company_address()
            address_parts = parse_address_parts(addr_full)
            website  = _get_about_field("Website")
            phone    = _get_about_field("Phone")
            industry = _get_about_field("Industry")
        except Exception:
            pass

        time.sleep(POLITE_DELAY)

        # -- People --
        people = []
        try:
            _driver.get(people_url)
            _wait.until(EC.presence_of_element_located((By.TAG_NAME, "body")))
            time.sleep(2)

            # employee count
            try:
                h2 = _driver.find_element(
                    By.XPATH,
                    "//h2[contains(translate(.,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'associated members')]"
                )
                m = re.search(r"(\d[\d,]*)", h2.text)
                num_emp = int(m.group(1).replace(",", "")) if m else ""
            except Exception:
                num_emp = ""

            _scroll_people_page()
            people = _scrape_people_cards()
        except Exception:
            people = []

        base_row = {
            "LinkedIn URL": company_base,
            "Company Name": company_name,
            "Website": website,
            "Company Phone": phone,
            "Industry": industry,
            "Number of Employees": num_emp,
            **address_parts,
        }

        if not people:
            rows.append({**base_row, "Name": "", "Designation": "", "Director Level": "", "Contact Linkedin URL": "", "Personal Email": ""})
        else:
            for p in people:
                rows.append({**base_row, **p})

        done_urls.add(company_base)
        pd.DataFrame(rows).to_excel(PEOPLE_FILE, index=False)
        print(f"[STEP2] ✅ {company_name or company_base} | contacts: {len(people)} | total rows: {len(rows)}")
        time.sleep(POLITE_DELAY)

    state["step"] = 3
    save_state(state)
    print(f"[STEP2] Done. Saved → {PEOPLE_FILE}")


# ============================================================
#  STEP 3 — FIND INDIVIDUAL LINKEDIN PROFILE URLs
#            Google first, Bing fallback
# ============================================================

def _google_captcha_detected() -> bool:
    try:
        src = norm(_driver.page_source)
        cur = (_driver.current_url or "").lower()
        return (
            "google.com/sorry" in cur
            or "unusual traffic" in src
            or "automated queries" in src
            or "our systems have detected unusual traffic" in src
        )
    except Exception:
        return True

def _extract_linkedin_from_redirect(href: str) -> str:
    try:
        qs  = parse_qs(urlparse(href).query)
        real = qs.get("url", [""])[0]
        return unquote(real)
    except Exception:
        return ""

def _extract_name_from_title(title: str) -> str:
    if not title:
        return ""
    title = re.sub(r"\.\.\.$", "", title).strip()
    if " - " in title:
        return title.split(" - ", 1)[0].strip()
    return ""

# Google result selectors (multiple fallbacks)
_GOOGLE_RESULT_SELECTORS = ["div.b8lM7", "div.g", "div[data-hveid]"]

def _parse_google_results() -> tuple[str, str]:
    for sel in _GOOGLE_RESULT_SELECTORS:
        try:
            _wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, sel)))
            blocks = _driver.find_elements(By.CSS_SELECTOR, sel)
            for block in blocks:
                try:
                    a = None
                    try:
                        a = block.find_element(By.CSS_SELECTOR, 'a[href*="linkedin.com/in/"]')
                    except Exception:
                        a = block.find_element(By.CSS_SELECTOR, 'a[href]')
                    if not a:
                        continue
                    href  = a.get_attribute("href") or ""
                    final = href
                    if "google.com/url" in href.lower():
                        real = _extract_linkedin_from_redirect(href)
                        if real:
                            final = real
                    if "linkedin.com/in/" not in final.lower():
                        try:
                            a2 = block.find_element(By.CSS_SELECTOR, 'a[href*="linkedin.com/in/"]')
                            final = a2.get_attribute("href") or ""
                        except Exception:
                            continue
                    if "linkedin.com/in/" in final.lower():
                        title = ""
                        try:
                            title = block.find_element(By.CSS_SELECTOR, "h3").text.strip()
                        except Exception:
                            pass
                        return final, title
                except Exception:
                    continue
        except Exception:
            continue
    return "", ""

def _google_search(query: str, state: dict) -> tuple[str, str]:
    ensure_driver()
    url = "https://www.google.com/search?q=" + quote_plus(query)
    try:
        _driver.get(url)
    except (NoSuchWindowException, WebDriverException):
        ensure_driver()
        _driver.get(url)
    human_sleep(*PAGE_WAIT_RANGE)

    if _google_captcha_detected():
        state["google_captcha_streak"] += 1
        save_state(state)
        if state["google_captcha_streak"] >= CAPTCHA_SWITCH_THRESHOLD:
            print(f"⚠️ Google CAPTCHA streak={state['google_captcha_streak']} — switching to Bing.")
            state["use_bing"] = True
            save_state(state)
            return "", ""
        print("⚠️ Google CAPTCHA detected. Solve it in Chrome, then press Enter...")
        input()
        # wait for results
        for _ in range(60):
            if not _google_captcha_detected():
                state["google_captcha_streak"] = 0
                save_state(state)
                return _parse_google_results()
            time.sleep(2)
        return "", ""

    state["google_captcha_streak"] = 0
    save_state(state)
    return _parse_google_results()

def _bing_search(query: str) -> tuple[str, str]:
    ensure_driver()
    url = "https://www.bing.com/search?q=" + quote_plus(query)
    try:
        _driver.get(url)
    except (NoSuchWindowException, WebDriverException):
        ensure_driver()
        _driver.get(url)
    human_sleep(*PAGE_WAIT_RANGE)

    # Bing result containers
    for sel in ["li.b_algo", "div.b_algo"]:
        try:
            blocks = _driver.find_elements(By.CSS_SELECTOR, sel)
            for block in blocks:
                try:
                    a = block.find_element(By.CSS_SELECTOR, 'a[href*="linkedin.com/in/"]')
                    href = a.get_attribute("href") or ""
                    if "linkedin.com/in/" in href.lower():
                        title = ""
                        try:
                            title = block.find_element(By.CSS_SELECTOR, "h2").text.strip()
                        except Exception:
                            pass
                        return href, title
                except Exception:
                    continue
        except Exception:
            continue
    return "", ""

def _find_linkedin_url(company: str, designation: str, contact_name: str, state: dict) -> tuple[str, str, str]:
    if is_linkedin_member(contact_name):
        query = f'{company} {designation} LinkedIn'
    else:
        query = f'{contact_name} {designation} {company} LinkedIn'

    if state.get("use_bing"):
        url, title = _bing_search(query)
        return url, title, "Bing"
    else:
        url, title = _google_search(query, state)
        if not url and state.get("use_bing"):
            url, title = _bing_search(query)
            return url, title, "Bing"
        return url, title, "Google"

def step3_find_profiles(state: dict):
    print("\n========== STEP 3: Find Individual LinkedIn Profiles ==========")
    df = pd.read_excel(PROFILES_FILE) if os.path.exists(PROFILES_FILE) else pd.read_excel(PEOPLE_FILE)

    if "Contact LinkedIn URL" not in df.columns:
        df["Contact LinkedIn URL"] = ""
    if "Matched Title" not in df.columns:
        df["Matched Title"] = ""
    if "Search Source" not in df.columns:
        df["Search Source"] = ""

    ensure_driver()

    try:
        for i, row in df.iterrows():
            existing = str(row.get("Contact LinkedIn URL", "")).strip()
            if existing and existing.lower() not in ("nan", ""):
                continue

            contact_name = str(row.get("Name", "")).strip()
            designation  = str(row.get("Designation", "")).strip()
            company      = str(row.get("Company Name", "")).strip()

            print(f"\n🔍 [{i+1}] {contact_name} | {designation} | {company}")

            try:
                url, title, source = _find_linkedin_url(company, designation, contact_name, state)
            except (NoSuchWindowException, WebDriverException):
                ensure_driver()
                url, title, source = _find_linkedin_url(company, designation, contact_name, state)

            if url:
                df.at[i, "Contact LinkedIn URL"] = url
                df.at[i, "Matched Title"]        = title
                df.at[i, "Search Source"]        = source
                extracted = _extract_name_from_title(title)
                if extracted and is_linkedin_member(contact_name):
                    df.at[i, "Name"] = extracted
                    print(f"   🧠 Name updated → {extracted}")
                print(f"   ✅ {source}: {url}")
            else:
                df.at[i, "Search Source"] = source
                print(f"   ❌ Not found ({source})")

            if (i + 1) % SAVE_EVERY == 0:
                df.to_excel(PROFILES_FILE, index=False)
                print(f"💾 Auto-saved at row {i+1}")

            human_sleep(*LOOP_DELAY)

    finally:
        df.to_excel(PROFILES_FILE, index=False)

    state["step"] = 4
    save_state(state)
    print(f"[STEP3] Done. Saved → {PROFILES_FILE}")


# ============================================================
#  STEP 4 — DEDUPLICATE + FINAL OUTPUT
# ============================================================

def step4_dedup(state: dict):
    print("\n========== STEP 4: Deduplicate + Final Output ==========")
    df = pd.read_excel(PROFILES_FILE)

    before = len(df)
    df.drop_duplicates(subset=["Contact LinkedIn URL"], keep="first", inplace=True)
    df.drop_duplicates(subset=["Name", "Company Name", "Designation"], keep="first", inplace=True)
    after = len(df)

    df.sort_values(
        by=["Director Level", "Company Name"],
        ascending=[False, True],
        inplace=True,
        ignore_index=True,
    )

    df.to_excel(FINAL_FILE, index=False)
    print(f"[STEP4] Removed {before - after} duplicates. Final rows: {after}")
    print(f"[STEP4] ✅ Saved → {FINAL_FILE}")

    state["step"] = 5
    save_state(state)


# ============================================================
#  STEP 5 — VERIFY PROFILES: visit each LinkedIn URL, check Company Name
# ============================================================

def _get_profile_name(driver) -> tuple[str, str]:
    """Return (first_name, last_name) from profile h1."""
    selectors = [
        "h1.inline.t-24",
        "h1.text-heading-xlarge",
        "h1",
    ]
    for sel in selectors:
        try:
            h1 = driver.find_element(By.CSS_SELECTOR, sel)
            full = text_clean(h1.text)
            if full:
                parts = full.split(" ", 1)
                first = parts[0]
                last  = parts[1] if len(parts) > 1 else ""
                return first, last
        except Exception:
            pass
    return "", ""

def _get_profile_title(driver) -> str:
    """Return current job title from profile."""
    selectors = [
        "div.text-body-medium.break-words",
        ".pv-text-details__left-panel div.text-body-medium",
    ]
    for sel in selectors:
        try:
            el = driver.find_element(By.CSS_SELECTOR, sel)
            t = text_clean(el.text)
            if t:
                return t
        except Exception:
            pass
    return ""

def _get_current_company_from_experience(driver) -> str:
    """
    Find the CURRENT company from experience section.
    Looks for 'Present' date indicator.
    """
    try:
        # Experience entries with 'Present'
        spans = driver.find_elements(By.CSS_SELECTOR, "span.pvs-entity__caption-wrapper")
        present_indices = []
        for i, sp in enumerate(spans):
            if "present" in sp.text.lower():
                present_indices.append(i)

        if not present_indices:
            return ""

        # For each 'Present' entry, find the company name nearby
        exp_section = driver.find_element(
            By.XPATH,
            "//section[.//span[normalize-space()='Experience']]"
        )
        # Get all bold spans (company/title names) inside experience
        bold_spans = exp_section.find_elements(
            By.CSS_SELECTOR,
            "span[aria-hidden='true']"
        )
        company_candidates = []
        for sp in bold_spans:
            t = text_clean(sp.text)
            if t and len(t) > 1:
                company_candidates.append(t)

        # Return first non-empty candidate after title (heuristic: alternating title/company)
        if len(company_candidates) >= 2:
            return company_candidates[1]
        if company_candidates:
            return company_candidates[0]
    except Exception:
        pass
    return ""

def _company_match(expected: str, found: str) -> bool:
    """Check if expected company name appears in found (case-insensitive, partial)."""
    if not expected or not found:
        return False
    e = norm(expected)
    f = norm(found)
    # exact or partial match
    return e in f or f in e or any(
        word in f for word in e.split() if len(word) > 3
    )

def step5_verify_profiles(state: dict):
    print("\n========== STEP 5: Verify LinkedIn Profiles vs Company Name ==========")

    df = pd.read_excel(VERIFIED_FILE) if os.path.exists(VERIFIED_FILE) else pd.read_excel(FINAL_FILE)

    if "First Name" not in df.columns:
        df["First Name"] = ""
    if "Last Name" not in df.columns:
        df["Last Name"] = ""
    if "Profile Title" not in df.columns:
        df["Profile Title"] = ""
    if "Company Match" not in df.columns:
        df["Company Match"] = ""

    ensure_driver()
    # Make sure we are logged in
    cur_url = _driver.current_url.lower()
    if "linkedin.com" not in cur_url:
        linkedin_auto_login()

    saved = 0

    for i, row in df.iterrows():
        already = str(row.get("Company Match", "")).strip()
        if already in ("Yes", "No"):
            continue

        profile_url = str(row.get("Contact LinkedIn URL", "")).strip()
        if not profile_url or profile_url.lower() == "nan" or "linkedin.com/in/" not in profile_url.lower():
            df.at[i, "Company Match"] = "No URL"
            continue

        company_name = str(row.get("Company Name", "")).strip()

        print(f"[{i+1}] Visiting: {profile_url}")
        try:
            _driver.get(profile_url)
            time.sleep(3)

            # scroll a bit to load experience
            _driver.execute_script("window.scrollTo(0, 800);")
            time.sleep(1.5)

            first, last = _get_profile_name(_driver)
            title        = _get_profile_title(_driver)
            found_co     = _get_current_company_from_experience(_driver)

            match = _company_match(company_name, found_co) or _company_match(company_name, title)

            df.at[i, "First Name"]    = first
            df.at[i, "Last Name"]     = last
            df.at[i, "Profile Title"] = title
            df.at[i, "Company Match"] = "Yes" if match else "No"

            print(f"   Name: {first} {last} | Title: {title} | Co: {found_co} | Match: {'Yes' if match else 'No'}")

        except Exception as e:
            print(f"   ERROR: {e}")
            df.at[i, "Company Match"] = "Error"

        saved += 1
        if saved % SAVE_EVERY == 0:
            df.to_excel(VERIFIED_FILE, index=False)
            print(f"   [SAVE] Auto-saved at row {i+1}")

        human_sleep(3, 6)

    df.to_excel(VERIFIED_FILE, index=False)
    print(f"[STEP5] Done. Saved → {VERIFIED_FILE}")

    state["step"] = 6
    save_state(state)


# ============================================================
#  MAIN ORCHESTRATOR
# ============================================================

def main():
    state = load_state()

    print("=" * 60)
    print("  LinkedIn Agent — Unified Orchestrator")
    print(f"  Starting from Step {state['step']}")
    print("=" * 60)

    try:
        if state["step"] <= 1:
            step1_collect_companies(state)

        if state["step"] <= 2:
            step2_scrape_people(state)

        if state["step"] <= 3:
            step3_find_profiles(state)

        if state["step"] <= 4:
            step4_dedup(state)

        if state["step"] <= 5:
            step5_verify_profiles(state)

        print("\nPIPELINE COMPLETE!")
        print(f"   Verified output  -> {VERIFIED_FILE}")
        print(f"   Unverified final -> {FINAL_FILE}")

        # Reset state for next full run
        if os.path.exists(STATE_FILE):
            os.remove(STATE_FILE)

    finally:
        quit_driver()


if __name__ == "__main__":
    main()
