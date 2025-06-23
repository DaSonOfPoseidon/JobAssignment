import sys
import os
import re
import time
from datetime import datetime, timedelta, date
from collections import defaultdict
from tkinter import messagebox
from dotenv import load_dotenv, set_key
import pandas as pd
from rapidfuzz import fuzz
import threading
import webbrowser
import tkinter as tk
from tkcalendar import DateEntry
from email_reader import connect_imap, find_matching_msg_nums, fetch_body, extract_relevant_section
from tkinter import simpledialog
from tkinter import ttk
from tkinterdnd2 import DND_FILES, TkinterDnD
from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeout
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "../embedded_python/lib")))


# FEATURE ADDONS
# "Remove Old Aassignments" checkbox, removes old assignments and only keeps the new contractors.
#
#
#



HERE         = os.path.abspath(os.path.dirname(__file__))
PROJECT_ROOT = os.path.abspath(os.path.join(HERE, os.pardir))
OUTPUT_DIR   = os.path.join(PROJECT_ROOT, "Outputs")
COOKIE_PATH  = os.path.join(PROJECT_ROOT, "cookies.pkl")
ENV_PATH     = os.path.join(PROJECT_ROOT, ".env")
LOG_FOLDER   = os.path.join(PROJECT_ROOT, "logs")

# ensure folders exist
os.makedirs(OUTPUT_DIR, exist_ok=True)
os.makedirs(LOG_FOLDER, exist_ok=True)

# === CONFIGURATION ===
SHOW_ALL_OUTPUT_IN_CONSOLE = True
load_dotenv(dotenv_path=ENV_PATH)
cookie_lock = threading.Lock()

COLUMN_DATE = 0
COLUMN_TIME = 1
COLUMN_NAME = 2
COLUMN_TYPE = 3
COLUMN_WO = 4
COLUMN_ADDRESS = 5
COLUMN_DROPDOWN = 7

BASE_URL = "http://inside.sockettelecom.com/workorders/view.php?nCount="

CONTRACTOR_LABELS = {
    "(none)": None,
    "SubT": "Subterraneus Installs",
    "TGS": "TGS Fiber",
    "Tex-Star": "Tex-Star Communications",
    "Pifer": "Pifer Quality Communications",
    "Advanced": "Advanced Electric",
    "All Clear": "All Clear",
    "North Sky": "North Sky",
}

CONTRACTOR_NAME_CORRECTIONS = {
    "Subterraneus Installs": {
        "brandon": "Brandon Turner",
        "jeff": "Jeffrey Givens",
        "chris": "Chris Kunkle",
        "cliff": "Clifford Kunkle",
        "simmie": "Simmie Dunn",   
        "andrew": "Andrew Orton",
        "dooley": "Dooley Heflin",
        "george": "George Stone",
        "john": "John Orton",
    },
    "TGS Fiber": {
        "jacob": "Jacob Jones",
        "clinton": "William Woods",
        "nick": "Nick Prichett",
        "kyle": "Kyle Thatcher",
        "blake": "Blake Wellman",
        "adam": "Adam Ward",
    },
    "Tex-Star Communications": {
        "robby": "Robby Cowart",
        "david": "David Villarreal",
        "ryan": "Ryan Sharp",
        "robbie": "Robby Cowart",
        "tommu": "Tommy Estrada",
        "frank": "Francisco Morales",
        "marcus": "Demarcus Blackmon"
    },
    "Pifer Quality Communications": {
        "caleb": "Caleb Pifer",
        "blake": "Blake Pifer",
        "cody": "Cody Wolfe"
    },
    "All Clear": {
        "brandon": "Brandon Thompson",
        "jacob": "Jacob Hein"
    },
    "default": {    
        "will": "William Woods",
        "brandon": "Brandon Turner"
    }
}

log_lines = []

class PlaywrightDriver:
    def __init__(self, headless=True, state_path=None, playwright=None, browser=None):
        # Use state_path for session persistence if provided
        from pathlib import Path
        self.state_path = state_path or os.path.join(PROJECT_ROOT, "state.json")
        if playwright and browser:
            self._pw = playwright
            self.browser = browser
        else:
            self._pw = sync_playwright().start()
            self.browser = self._pw.chromium.launch(headless=headless)
        # Load context (cookies/session)
        if Path(self.state_path).exists():
            self.context = self.browser.new_context(storage_state=self.state_path)
        else:
            self.context = self.browser.new_context()
        self.page = self.context.new_page()
        self.page.on("dialog", lambda dlg: dlg.dismiss())
        self.page.route("**/*.{png,svg}", lambda route: route.abort())

    def goto(self, url, timeout=5000, wait_until="load", retries=3):
        for attempt in range(retries):
            try:
                # Try navigation
                return self.page.goto(url, timeout=timeout, wait_until=wait_until)
            except Exception as e:
                # If navigation was aborted, log and retry
                if "ERR_ABORTED" in str(e) or "Navigation failed" in str(e):
                    log(f"🟡 [Attempt {attempt+1}] Navigation aborted for {url}: {e}")
                    # Try to dismiss any possible alerts
                    try:
                        self.page.keyboard.press("Escape")
                    except Exception:
                        pass
                    self.page.wait_for_timeout(300)
                    continue  # Retry
                elif isinstance(e, PlaywrightTimeout):
                    log(f"🟡 [Attempt {attempt+1}] Timeout navigating to {url}, retrying: {e}")
                    self.page.wait_for_timeout(500)
                    continue
                else:
                    log(f"❌ Navigation error: {e} for {url}")
                    # Optionally: self.page.reload()
                    break
        log(f"❌ Failed to load {url} after {retries} attempts")
        return None
        
    def save_state(self, path=None):
        self.context.storage_state(path=path or self.state_path)
    def __getattr__(self, name):
        # So you can use driver.page APIs directly
        return getattr(self.page, name)
    def close(self):
        self.context.close()
        try:
            self.browser.close()
            self._pw.stop()
        except Exception:
            pass

class Assigner:
    def __init__(self, num_threads=8):
        self.num_threads = num_threads

    def parse_jobs(self, lines):
        # dispatch to your existing per-contractor parsers
        jobs = []
        for label, parser in CONTRACTOR_FORMAT_PARSERS.items():
            # detect label on the fly or pass it in
            parsed = parser(lines)
            jobs.extend(parsed)
        return jobs

    def first_jobs(self, jobs):
        df = pd.DataFrame(jobs)
        assign_jobs(df)
        summary = build_first_jobs_summary(df)
        # flatten summary into a list of lines
        return [line for lines in summary.values() for line in lines]

def log(message):
    log_lines.append(message)
    if SHOW_ALL_OUTPUT_IN_CONSOLE :
        print(message)

def gui_log(message):
    try:
        if HEADLESS_MODE.get() and 'log_output_text' in globals():
            timestamp = datetime.now().strftime("[%H:%M:%S] ")
            new_message = f"{timestamp}{message}"
            log_output_text.config(state="normal")
            log_output_text.insert(tk.END, f"{new_message}\n")
            log_output_text.see(tk.END)
            log_output_text.config(state="disabled")
            log(new_message)
    except Exception as e:
        log(f"⚠️ Failed to update GUI log: {e}")

def ask_date_range(dates):
    top = tk.Toplevel()
    top.title("Select Date Range")
    top.geometry("300x150")
    
    tk.Label(top, text="Start Date:").pack(pady=(10, 0))

    try:
        mindate = min(dates)
        maxdate = max(dates)
    except (ValueError, TypeError):
        mindate = None
        maxdate = None

    start = DateEntry(top, mindate=mindate, maxdate=maxdate)
    start.pack()

    tk.Label(top, text="End Date:").pack(pady=(10, 0))
    end = DateEntry(top, mindate=mindate, maxdate=maxdate)
    end.pack()

    selected = {}

    def submit():
        selected['start'] = start.get_date()
        selected['end'] = end.get_date()
        top.destroy()

    tk.Button(top, text="OK", command=submit).pack(pady=10)
    top.grab_set()
    top.wait_window()

    return selected.get('start'), selected.get('end')

def parse_flexible_time(t):
    t = str(t).strip().lower().replace('.', '')

    # Force PM for common afternoon times like 1:00, 2:00, 3:00, etc.
    force_pm_hours = {1, 2, 3, 4, 5}
    match = re.match(r"^(\d{1,2})(:\d{2})?$", t)
    if match:
        hour = int(match.group(1))
        minute = int(match.group(2)[1:]) if match.group(2) else 0

        if hour in force_pm_hours:
            hour += 12  # convert to 13:00, 14:00, etc.

        return datetime.strptime(f"{hour}:{minute:02d}", "%H:%M")

    # Append am/pm suffix based on heuristics if it's in a known format
    if re.match(r"^\d{1,2}:\d{2}$", t):
        hour = int(t.split(":")[0])
        if hour in force_pm_hours:
            t += " pm"
        else:
            t += " am"

    formats = ("%I:%M %p", "%I %p", "%I%p", "%H:%M", "%H:%M:%S", "%I:%M%p")

    for fmt in formats:
        try:
            return datetime.strptime(t, fmt)
        except:
            continue

    return pd.NaT

def force_pm_if_needed(dt):
    if pd.isna(dt):
        return dt
    if dt.hour in {1, 2, 3, 4, 5}:
        return dt.replace(hour=dt.hour + 12)
    return dt

def flexible_date_parser(date_str):
    try:
        return pd.to_datetime(str(date_str), errors='coerce')
    except:
        return None

def format_time_str(t):
    try:
        if isinstance(t, datetime):
            dt_obj = t
        else:
            dt_obj = datetime.strptime(t.strip(), "%H:%M")
        return dt_obj.strftime("%-I%p").lower() if os.name != 'nt' else dt_obj.strftime("%#I%p").lower()
    except:
        return str(t).strip()

def build_first_jobs_summary(df, name_column="Dropdown"):
    first_jobs = defaultdict(list)

    df['TimeParsed'] = df['Time'].apply(lambda x: parse_flexible_time(str(x)))
    df['TimeParsed'] = df['TimeParsed'].apply(force_pm_if_needed)
    failed_times = df[df['TimeParsed'].isna()]
    if not failed_times.empty:
        log("\n⚠️ Could not parse the following time values:")
        log(failed_times[['Time']].to_string(index=False))

    df = df.dropna(subset=['TimeParsed', 'Name', 'Type', 'Address', 'WO'])

    for date, group in df.groupby(df['Date'].dt.date):
        group = group.sort_values('TimeParsed')
        seen = set()
        for _, row in group.iterrows():
            raw_tech = row.get(name_column, '').strip().lower()
            if raw_tech not in seen:
                seen.add(raw_tech)
                corrected = get_corrected_name(raw_tech, SELECTED_CONTRACTOR.get())
                parts = corrected.strip().split()
                if len(parts) >= 2:
                    tech_display = f"{parts[0].capitalize()} {parts[1][0].upper()}"
                else:
                    tech_display = corrected.capitalize()

                formatted_time = format_time_str(row['TimeParsed'])
                name_display = row['Name'].strip().title()
                type_display = row['Type'].strip().title()
                addr_display = row['Address'].strip().title()
                wo_display = f"WO {str(row['WO']).strip()}"

                line = f"{tech_display} - {formatted_time} - {name_display} - {type_display} - {addr_display} - {wo_display}"
                first_jobs[date].append(line)

    return first_jobs

def get_corrected_name(name_input, contractor_full):
    def normalize_name_key(name_input):
        name_input = name_input.strip().lower()
        parts = name_input.split()
        if len(parts) == 1:
            return parts[0]
        elif len(parts) >= 2:
            return f"{parts[0]} {parts[1][0]}"
        return name_input
    
    key = normalize_name_key(name_input)
    corrections = CONTRACTOR_NAME_CORRECTIONS.get(contractor_full)
    if corrections and key in corrections:
        return corrections[key]
    key_parts = key.split()
    if corrections and key_parts and key_parts[0] in corrections:
        return corrections[key_parts[0]]
    # Fuzzy match fallback
    if corrections and key_parts:
        best_match = None
        highest_score = 0
        for known_key in corrections:
            score = fuzz.ratio(key_parts[0], known_key)
            if score > 90 and score > highest_score:
                best_match = known_key
                highest_score = score
        if best_match:
            return corrections[best_match]
    fallback = CONTRACTOR_NAME_CORRECTIONS.get("default", {})
    if key_parts:
        return fallback.get(key_parts[0], name_input)
    return name_input

def match_dropdown_option(options, tech_name_input, contractor_full):
    """
    Args:
        options: list of option texts from Playwright's all_inner_texts()
        tech_name_input: the technician name you want to assign (e.g., "Brandon T")
        contractor_full: full contractor company name (for name corrections)
    Returns:
        The option text to select (exact match or best fuzzy match), or None if no match.
    """
    # Step 1: Correct the tech name for company-specific nicknames/variants
    corrected_name = get_corrected_name(tech_name_input, contractor_full)
    key_parts = corrected_name.lower().split()
    first_name = key_parts[0] if key_parts else ""
    last_initial = key_parts[1][0] if len(key_parts) > 1 else ""

    # Step 2: Full exact match (case-insensitive)
    for opt in options:
        if opt.strip().lower() == corrected_name.lower():
            return opt

    # Step 3: First name + last initial match
    for opt in options:
        full_parts = opt.lower().strip().split()
        if len(full_parts) >= 2 and full_parts[0] == first_name and full_parts[1][0] == last_initial:
            return opt

    # Step 4: Loose first name match, return if only one match
    potential_matches = [opt for opt in options if opt.lower().startswith(first_name)]
    if len(potential_matches) == 1:
        return potential_matches[0]

    # Step 5: Fuzzy match as fallback
    highest_score = 0
    best_opt = None
    for opt in options:
        score = fuzz.ratio(opt.lower(), corrected_name.lower())
        if score > highest_score:
            highest_score = score
            best_opt = opt
    if highest_score > 90:
        return best_opt

    return None  # If all else fails

def check_env_or_prompt_login(log):
    username = os.getenv("UNITY_USER")
    password = os.getenv("PASSWORD")

    if username and password:
        log("🔐 Loaded stored credentials.")
        return username, password

    while True:
        username, password = prompt_for_credentials()
        if not username or not password:
            messagebox.showerror("Login Cancelled", "Login is required to continue.")
            return None, None

        # Do NOT try logging in here (add ~15 sec delay in most use cases) — just trust them until used (nothing will break)
        save_env_credentials(username, password)
        log("✅ Credentials captured and saved to .env.")
        return username, password

def prompt_for_credentials():
    login_window = tk.Tk()
    login_window.withdraw()

    USERNAME = simpledialog.askstring("Login", "Enter your USERNAME:", parent=login_window)
    PASSWORD = simpledialog.askstring("Login", "Enter your PASSWORD:", parent=login_window, show="*")

    login_window.destroy()
    return USERNAME, PASSWORD

def save_env_credentials(USERNAME, PASSWORD):
    dotenv_path = ".env"
    if not os.path.exists(dotenv_path):
        with open(dotenv_path, "w") as f:
            f.write("")
    set_key(dotenv_path, "UNITY_USER", USERNAME)
    set_key(dotenv_path, "PASSWORD", PASSWORD)

def jobs_to_html(first_jobs, company): #FIRST JOBS FORMAT
    now = datetime.now()
    timestamp = now.strftime("%I:%M%p, %m/%d").lower()
    html_header = f"""
    <div style="font-family: Apatos, Arial, sans-serif; font-size:12pt;">
    <p>Please call into 163 for tech assistance - Please call 633 for dispatcher</p>
    <p>
    This is a reminder. Please make sure we are returning the HST's to the office once we are done using them.<br>
    Please also make sure you are updating WO notes and completing the WO after completing the dispatch/install.
    </p>
    <p>As of {timestamp} the jobs are set as follows.....</p><br><br>
    <h3 style="margin-bottom:0">{company} Contractors -</h3><br>
    <ul style="margin-top:6px;">
    """

    lines = []
    for day, jobs in first_jobs.items():
        for job in jobs:
            # <li> for bullet or <div> for block
            lines.append(f"<div>{job}</div><br>")
    html_footer = """
    </ul>
    <br>
    <br>
    <h3 style="margin-bottom:0">Internals -</h3><br>
    <div>Brandon - 8am - </div><br>
    <div>Billy - 8am - </div><br>
    <div>Cole - 8am - </div><br>
    <div>Gladston - 8am - </div><br>
    <div>Matt - 8am - </div><br>
    <div>Tylor - 8am - </div>
    </div>
    """
    
    return html_header + "\n".join(lines) + html_footer

def verify_work_order_page(driver, wo_number, url, max_attempts=3):
    for attempt in range(1, max_attempts + 1):
        try:
            # If the URL isn't correct, navigate there
            if not driver.page.url.strip().endswith(wo_number):
                driver.goto(url)
            # Wait for the Work Order # field to appear (no need for poll_frequency with Playwright)
            selector = "//td[contains(text(), 'Work Order #:')]/following-sibling::td"
            elem = driver.page.wait_for_selector(f"xpath={selector}", timeout=10_000)
            # Check the displayed WO number
            text = elem.inner_text().strip()
            if text == wo_number:
                return True
        except Exception:
            log(f"🟡 Attempt {attempt}: Unable to verify WO #{wo_number}")
        time.sleep(5)
    return False

def assign_contractor(driver, wo_number, desired_contractor_full): #COMPANY ASSIGNMENT
    try:
        page = driver.page
        # 🧠 Trigger the assignment UI via JavaScript
        page.evaluate(f"assignContractor('{wo_number}');")
        page.wait_for_selector("#ContractorID", timeout=10_000)
        page.wait_for_timeout(500)  # let modal settle

        # ✅ Get current contractor assignment from the page
        contractor_texts = [
            elem.inner_text().strip()
            for elem in page.locator("b").all()
        ]
        current_contractor = None
        for text in contractor_texts:
            if " - (Primary" in text:
                current_contractor = text.split(" - ")[0].strip()
                break

        if current_contractor == desired_contractor_full:
            log(f"✅ Contractor '{current_contractor}' already assigned to WO #{wo_number}")
            return

        log(f"🧹 Reassigning from '{current_contractor}' → '{desired_contractor_full}'")

        # 🧽 Remove currently assigned contractors
        remove_links = page.locator("a").filter(has_text="Remove")
        for i in range(remove_links.count()):
            try:
                link = remove_links.nth(i)
                link.scroll_into_view_if_needed()
                link.click()
                # Confirm alert (Playwright auto-dismisses dialog, but for safety:)
                page.wait_for_timeout(500)
            except Exception as e:
                log(f"❌ Could not remove contractor: {e}")

        # 🏷️ Assign the new contractor
        contractor_dropdown = page.locator("#ContractorID")
        contractor_dropdown.select_option(label=desired_contractor_full)
        role_dropdown = page.locator("#ContractorType")
        role_dropdown.select_option(label="Primary")

        # Hide any modal overlays if needed (FileList)
        try:
            file_list_elem = page.locator("#FileList")
            if file_list_elem.is_visible():
                page.evaluate("el => el.style.display = 'none'", file_list_elem)
        except Exception:
            pass  # It's okay if FileList isn't present

        assign_button = page.locator("input[type='button'][value='Assign']")
        assign_button.click()

        log(f"🏷️ Assigned contractor '{desired_contractor_full}' to WO #{wo_number}")

    except Exception as e:
        log(f"❌ Contractor assignment process failed for WO #{wo_number}: {e}")

def reformat_contractor_text(text):

    def ask_date_with_default(default_date):
        root = tk.Tk()
        root.withdraw()

        popup = tk.Toplevel()
        popup.title("Select Job Date")
        popup.geometry("250x120")

        tk.Label(popup, text="Job Date:").pack(pady=(10, 0))
        cal = DateEntry(popup, width=12, background='darkblue', foreground='white', borderwidth=2)
        cal.set_date(default_date)
        cal.pack(pady=5)

        selected = {}
        def submit():
            selected["date"] = cal.get_date()
            popup.destroy()

        tk.Button(popup, text="OK", command=submit).pack()
        popup.grab_set()
        popup.wait_window()
        root.destroy()
        return selected.get("date")

    lines = [line.strip() for line in text.strip().splitlines() if line.strip()]
    selected_contractor_label = SELECTED_CONTRACTOR.get()
    selected_contractor_name = CONTRACTOR_LABELS.get(selected_contractor_label)

    if not selected_contractor_name:
        gui_log("❌ Please select a contractor from the dropdown before parsing.")
        return []

    parser_fn = CONTRACTOR_FORMAT_PARSERS.get(selected_contractor_name)
    if not parser_fn:
        gui_log(f"❌ No parser available for contractor: {selected_contractor_name}")
        return []

    return parser_fn(lines)

def is_headless():
    try:
        return HEADLESS_MODE.get()
    except:
        return False

def save_and_open_html(html_str, filename="FirstJobsSummary.html"):
    # Write HTML to file
    with open(filename, "w", encoding="utf-8") as f:
        f.write(html_str)
    # Open in the default browser (works everywhere)
    file_url = 'file://' + os.path.realpath(filename)
    webbrowser.open(file_url)

def handle_login(driver):
    # 1) Try to restore state
    driver.goto("http://inside.sockettelecom.com/")
    if "login.php" not in driver.page.url:
        log("✅ Session restored with stored state.")
        clear_first_time_overlays(driver.page)
        return
    # 2) Otherwise, fall back to manual login
    user, pw = check_env_or_prompt_login(log)
    driver.goto("http://inside.sockettelecom.com/system/login.php")
    driver.page.fill("input[name='username']", user)
    driver.page.fill("input[name='password']", pw)
    driver.page.click("#login")
    driver.page.wait_for_selector("iframe#MainView", timeout=10_000)
    clear_first_time_overlays(driver.page)
    # 3) Persist for next runs
    driver.save_state()
    log("✅ Logged in via credentials.")

def clear_first_time_overlays(page):
    selectors = [
        'xpath=//input[@id="valueForm1" and @type="button"]',
        'xpath=//input[@value="Close This" and @type="button"]',
        'xpath=//form[starts-with(@id,"valueForm")]//input[@type="button"]',
        'xpath=//form[@id="f"]//input[@type="button"]'
    ]
    for sel in selectors:
        while True:
            try:
                btn = page.wait_for_selector(sel, timeout=500)
                btn.click()
                page.wait_for_timeout(200)
            except PlaywrightTimeout:
                break

def normalize_subt_multiline_format(lines):
    normalized = []
    buffer = []

    for raw_line in lines:
        line = raw_line.strip()
        if not line:
            continue

        # Auto-join if a full job row is pasted already
        if '\t' in line and line.count('\t') >= 5 and re.match(r"\d{1,2}-[A-Za-z]{3}", line):
            normalized.append(line)
            continue

        buffer.append(line)

        # If we have 3 lines (Date, Time, Job Row), normalize
        if len(buffer) == 3:
            date = buffer[0]
            time = buffer[1]
            details = buffer[2]

            parts = [p.strip() for p in details.split('\t') if p.strip()]
            if len(parts) < 4:
                print(f"[SKIP] Malformed job: {buffer}")
                buffer = []
                continue

            name = parts[0]
            job_type = parts[1]
            wo = parts[2]
            address = parts[3]
            tech = parts[-1] if len(parts) >= 5 else ""

            normalized.append(f"{date}\t{time}\t{name}\t{job_type}\t{wo}\t{address}\t{tech}")
            buffer.clear()

        elif len(buffer) > 3:
            print(f"[SKIP] Malformed job block (too many lines): {buffer}")
            buffer.clear()

    return normalized

def parse_subt_line(fields):
    # Assume fields is a list of columns from splitting the line
    wo_regex = re.compile(r'^WO\s*\d{6}$', re.I)
    type_idx, wo_idx = None, None

    # Only check the likely columns
    for idx in [3, 4]:
        if wo_regex.match(fields[idx].strip()):
            wo_idx = idx
        else:
            type_idx = idx

    # Fallback if both are not found
    if wo_idx is None or type_idx is None:
        return None  # or handle as a parse error

    job_type = fields[type_idx].strip()
    wo = fields[wo_idx].strip().replace("WO", "").strip()
    return job_type, wo

def assign_jobs(df, contractor_label=None):
    gui_log(f"Processing {len(df)} work orders...")
    driver = PlaywrightDriver(headless=is_headless())
    handle_login(driver)

    for index, row in df.iterrows():
        raw_wo = row['WO']
        raw_name = str(row.get('Dropdown') or row.get('Tech', '')).strip()
        raw_name = re.sub(r"^[\-\s]+", "", raw_name)

        name_parts = raw_name.split()
        if len(name_parts) >= 2:
            dropdown_value = f"{name_parts[0].capitalize()} {name_parts[1][0].upper()}"
        else:
            dropdown_value = raw_name.capitalize()

        try:
            wo_number = str(int(raw_wo))
        except (ValueError, TypeError):
            gui_log(f"❌ Invalid WO number '{raw_wo}' on line {index + 2}.")
            continue

        url = BASE_URL + wo_number
        log(f"\n🔗 Opening WO #{wo_number} — {url}")
        driver.goto(url)

        # Check for unscheduled error (optional; can add custom error logic if you want)
        try:
            if "no scheduled dates found" in driver.page.content().lower():
                gui_log(f"❌ WO {wo_number} is not scheduled — skipping.")
                continue
        except Exception:
            pass

        if not verify_work_order_page(driver, wo_number, url):
            gui_log(f"❌ Failed to verify WO #{wo_number}. Skipping.")
            continue

        desired_contractor_label = SELECTED_CONTRACTOR.get()
        desired_contractor_full = CONTRACTOR_LABELS.get(desired_contractor_label)
        if desired_contractor_full:
            assign_contractor(driver, wo_number, desired_contractor_full)

        try:
            page = driver.page  # Your Playwright page object
            page.wait_for_selector("#AssignEmpID", timeout=10000)
            select_elem = page.locator("#AssignEmpID")

            # Get all option texts for robust matching
            options = select_elem.locator("option").all_inner_texts()
            matched_option = match_dropdown_option(options, dropdown_value, desired_contractor_full)
            if not matched_option:
                gui_log(f"❌ No dropdown match for '{dropdown_value}' — skipping WO #{wo_number}")
                continue

            # Check if already assigned
            page.wait_for_selector("#AssignmentsList", timeout=5000)
            assigned_names = page.locator("#AssignmentsList").inner_text()
            if matched_option.lower() in assigned_names.lower():
                log(f"🟡 WO #{wo_number}: '{matched_option}' already assigned.")
                continue

            # Select tech and click Add
            select_elem.select_option(label=matched_option)
            add_button = page.locator("button.button.Socket")
            add_button.click()
            gui_log(f"WO {wo_number} - Assigned to {matched_option} ({desired_contractor_full})")
        except Exception as e:
            gui_log(f"❌ Error assigning WO #{wo_number}: {e}")
        
        # 2. Wait for the page to "settle" after contractor assignment
        try:
            driver.page.wait_for_selector("xpath=//td[contains(text(), 'Work Order #:')]", timeout=10_000)
            driver.page.wait_for_timeout(200)  # slight extra buffer
        except Exception:
            log("⚠️ Could not detect Work Order # after assignment. Retrying may be needed.")

    driver.close()  # Clean up browser at the end
    gui_log(f"\n✅ Done processing work orders.")

# ===== Contractor Parsers =====

def parse_subterraneus_format(lines):
    jobs = []
    current_date = None
    current_time = None

    lines = normalize_subt_multiline_format(lines)
    print(f"[DEBUG] Total normalized lines to parse: {len(lines)}")
    print(f"[DEBUG] Example normalized line: {lines[0] if lines else 'None'}")

    for i, line in enumerate(lines):
        line = line.strip()
        if not line:
            continue

        # Handle date-only lines (e.g. "5-May" or "5/5")
        if re.fullmatch(r"\d{1,2}-[A-Za-z]{3}", line) or re.fullmatch(r"\d{1,2}/\d{1,2}", line):
            current_date = line
            continue

        # Handle time-only lines (e.g. "8:00 AM" or "15:30")
        if re.fullmatch(r"\d{1,2}:\d{2}\s?(?:AM|PM)?", line, re.IGNORECASE):
            current_time = line
            continue

        # Try to split by tabs first (new format)
        parts = line.split('\t')
        print(f"[DEBUG] Split parts: {parts}")
        if len(parts) < 6:
            # Fallback: try splitting by 2+ spaces
            parts = re.split(r'\s{2,}', line)

        if len(parts) < 6:
            print(f"[SKIP] Line malformed or incomplete: {line}")
            continue

        try:
            parts = line.strip().split('\t')
            if len(parts) < 6:
                print(f"[SKIP] Malformed normalized line: {line}")
                continue

            temp = parse_subt_line(parts)

            date_str = parts[0].strip()
            time_str = parts[1].strip()
            name = parts[2].strip()
            job_type = temp[0].strip()
            wo_field = temp[1].strip()
            address = parts[5].strip()

            non_empty_tail = [p.strip() for p in parts[6:] if p.strip()]
            tech = non_empty_tail[-1] if non_empty_tail else ""

            if not tech:
                print(f"[SKIP] Missing tech for: {line}")
                continue

            wo_match = re.search(r"\b(\d{6})\b", wo_field)
            if not wo_match:
                print(f"[SKIP] Invalid WO: {line}")
                continue

            jobs.append({
                "Date": date_str,
                "Time": time_str,
                "Name": name,
                "Type": job_type,
                "WO": wo_match.group(1),
                "Address": address,
                "Tech": tech
            })
        except Exception as e:
            print(f"[ERROR] Failed to parse line: {line} — {e}")

    print(f"[✅ Parsed] {len(jobs)} SubT job(s)")
    return jobs

def parse_tgs_format(lines):
    jobs = []

    current_tech = None
    current_date = None
    last_time = None

    job_line_pattern = re.compile(
        r"(\d{1,2}:\d{2})\s*-\s*(.*?)\s*-\s*[\d\-]+?\s*-\s*(.*?)\s*-\s*(.*?)\s*-\s*WO\s*?(\d+)",
        re.IGNORECASE
    )

    for raw_line in lines:
        line = raw_line.strip()
        if not line:
            continue

        # Skip known filler
        if line.lower() == "clinton" or "legacy drop" in line.lower() or "block" in line.lower():
            continue

        # Pure contractor header (e.g. "DONNELL")
        if re.fullmatch(r"(?:[A-Z]{3,}|[A-Z]+(?:\s+[A-Z]+)+)", line):
            current_tech = line
            continue

        # Pure date-only lines (e.g. "5/5/2025")
        if re.fullmatch(r"\d{1,2}/\d{1,2}/\d{4}", line):
            current_date = line
            continue

        # Pure time-only lines (e.g. "8 AM", "10:00PM")
        if re.fullmatch(r"\d{1,2}(:\d{2})?\s?(?:AM|PM)", line, re.IGNORECASE):
            last_time = line
            continue

        # Now parse the real job lines
        match = job_line_pattern.match(line)
        if not match:
            print(f"[SKIP] Unmatched or malformed TGS line: {line}")
            continue

        time_str, name, job_type, address, wo_number = match.groups()
        formatted_time = format_time_str(time_str.strip())

        jobs.append({
            "Date": current_date,
            "Time": formatted_time,
            "Name": name.strip(),
            "Type": job_type.strip(),
            "WO": wo_number.strip(),
            "Address": address.strip(),
            "Tech": current_tech
        })

    print(f"[✅ Parsed] {len(jobs)} TGS job(s)")
    return jobs

def parse_pifer_format(lines):
    clean_lines = []
    for raw in lines:
        if raw is None:
            continue
        clean_lines.append(str(raw))
    lines = clean_lines
    # 1) Merge only indented continuation lines (not tech / date / time)
    merged = []
    for raw in lines:
        if raw.startswith((" ", "\t")) and not raw.strip().startswith("-"):
            # continuation of the previous line
            merged[-1] = merged[-1].rstrip() + " " + raw.strip()
        else:
            merged.append(raw)

    # 2) Now parse as before
    jobs = []
    current_tech = current_date = current_time = None
    job_line_pattern = re.compile(
        r"-\s*(?P<name>.*?)\s*-\s*[\d\-]+\s*-\s*(?P<type>.*?)\s*-\s*(?P<address>.*?)\s*-\s*WO\s*(?P<wo>\d+)",
        re.IGNORECASE
    )

    for raw in merged:
        line = raw.strip()
        if not line:
            continue

        #  Technician header: two+ words, letters only (no digits), not starting with '-'
        if not line.startswith("-") and not re.search(r"\d", line) and len(line.split()) >= 2:
            current_tech = line
            continue

        # Detect a date (e.g. 6-17-25 or 06-17-2025)
        if re.match(r"^\d{1,2}-\d{1,2}-\d{2,4}$", line):
            current_date = line
            continue

        # Detect a time (e.g. 8AM, 10PM)
        if re.match(r"^\d{1,2}(?:AM|PM)$", line, re.IGNORECASE):
            current_time = line.upper()
            continue

        # Finally, a full job line
        match = job_line_pattern.search(line)
        if match:
            jobs.append({
                "Date":    current_date,
                "Time":    current_time,
                "Tech":    current_tech,
                "Name":    match.group("name").strip(),
                "Type":    match.group("type").strip(),
                "Address": match.group("address").strip(),
                "WO":      match.group("wo").strip(),
            })

    print(f"[✅ Parsed] {len(jobs)} Pifer job(s)")
    return jobs

def parse_texstar_format(lines):
    jobs = []
    current_date = None

    # Original pattern
    job_line_pattern = re.compile(
        r"(\d{1,2}:\d{2})\s*-\s*(.*?)\s*-\s*[\d\-]+?\s*-\s*(.*?)\s*-\s*(.*?)\s*-\s*WO\s*(\d+)\s*[—-]{1,2}\s*(\w+)",
        re.IGNORECASE
    )

    # New alternate pattern
    job_line_alt_pattern = re.compile(
        r"(\d{1,2}:\d{2})\s*-\s*(.*?)\s*-\s*\d{4}-\d{4}-\d{4}\s*-\s*(.*?)\s*-\s*(.*?)\s*-\s*WO\s*(\d+)\s+(\w+)",
        re.IGNORECASE
    )

    for line in lines:
        line = line.strip()
        if not line:
            continue

        # Capture date
        if re.match(r"\d{1,2}[-/]\d{1,2}[-/]\d{2,4}", line):
            current_date = line
            continue

        match = job_line_pattern.match(line) or job_line_alt_pattern.match(line)
        if not match:
            continue

        time, name, job_type, address, wo_number, tech = match.groups()
        jobs.append({
            "Date": current_date,
            "Time": format_time_str(time),
            "Name": name.strip(),
            "Type": job_type.strip(),
            "WO": wo_number.strip(),
            "Address": address.strip(),
            "Tech": tech.strip()
        })

    print(f"[✅ Parsed] {len(jobs)} Tex-Star job(s)")
    return jobs

def parse_all_clear_format(lines):
    jobs = []
    current_date = None
    current_time = None

    dash_format = re.compile(
        r"(\d{1,2}:\d{2})?\s*-\s*(.*?)\s*-\s*[\d\-]+?\s*-\s*(.*?)\s*-\s*(.*?)\s*-\s*WO\s*(\d+)\s*-\s*(.+)",
        re.IGNORECASE
    )

    underscore_format = re.compile(
        r"(.*?)\s*-\s*[\d\-]+?\s*_+\s*(.*?)\s*_+\s*(.*?)\s*_+\s*WO\s*(\d+)\s*-\s*(.+)",
        re.IGNORECASE
    )

    for raw_line in lines:
        line = raw_line.strip()
        if not line:
            continue

        # Check for date (assumed to be outside your snippet—skip)
        if re.match(r"\d{1,2}[-/]\d{1,2}[-/]\d{2,4}", line):
            current_date = line
            continue

        # Update current time if it's a time marker
        if re.fullmatch(r"\d{1,2}(:\d{2})?\s?(AM|PM)?", line, re.IGNORECASE):
            current_time = format_time_str(line)
            continue

        # Dash-delimited format
        dash_match = dash_format.match(line)
        if dash_match:
            time, name, job_type, address, wo_number, tech = dash_match.groups()
            time = format_time_str(time) if time else current_time
            jobs.append({
                "Date": current_date,
                "Time": time,
                "Name": name.strip(),
                "Type": job_type.strip(),
                "WO": wo_number.strip(),
                "Address": address.strip(),
                "Tech": tech.strip()
            })
            continue

        # Underscore-delimited format
        underscore_match = underscore_format.match(line)
        if underscore_match:
            name, job_type, address, wo_number, tech = underscore_match.groups()
            jobs.append({
                "Date": current_date,
                "Time": current_time,
                "Name": name.strip(),
                "Type": job_type.strip(),
                "WO": wo_number.strip(),
                "Address": address.strip(),
                "Tech": tech.strip()
            })

    print(f"[✅ Parsed] {len(jobs)} All Clear job(s)")
    return jobs

def parse_north_sky_format(lines):
    jobs = []
    current_date = None
    current_time = None

    dash_format = re.compile(
        r"(\d{1,2}:\d{2})?\s*-\s*(.*?)\s*-\s*[\d\-]+?\s*-\s*(.*?)\s*-\s*(.*?)\s*-\s*WO\s*(\d+)\s*-\s*(.+)",
        re.IGNORECASE
    )

    underscore_format = re.compile(
        r"(.*?)\s*-\s*[\d\-]+?\s*_+\s*(.*?)\s*_+\s*(.*?)\s*_+\s*WO\s*(\d+)\s*-\s*(.+)",
        re.IGNORECASE
    )

    for raw_line in lines:
        line = raw_line.strip()
        if not line:
            continue

        # Check for date (assumed to be outside your snippet—skip)
        if re.match(r"\d{1,2}[-/]\d{1,2}[-/]\d{2,4}", line):
            current_date = line
            continue

        # Update current time if it's a time marker
        if re.fullmatch(r"\d{1,2}(:\d{2})?\s?(AM|PM)?", line, re.IGNORECASE):
            current_time = format_time_str(line)
            continue

        # Dash-delimited format
        dash_match = dash_format.match(line)
        if dash_match:
            time, name, job_type, address, wo_number, tech = dash_match.groups()
            time = format_time_str(time) if time else current_time
            jobs.append({
                "Date": current_date,
                "Time": time,
                "Name": name.strip(),
                "Type": job_type.strip(),
                "WO": wo_number.strip(),
                "Address": address.strip(),
                "Tech": tech.strip()
            })
            continue

        # Underscore-delimited format
        underscore_match = underscore_format.match(line)
        if underscore_match:
            name, job_type, address, wo_number, tech = underscore_match.groups()
            jobs.append({
                "Date": current_date,
                "Time": current_time,
                "Name": name.strip(),
                "Type": job_type.strip(),
                "WO": wo_number.strip(),
                "Address": address.strip(),
                "Tech": tech.strip()
            })

    print(f"[✅ Parsed] {len(jobs)} North Sky job(s)")
    return jobs

CONTRACTOR_FORMAT_PARSERS = {
    "Subterraneus Installs": parse_subterraneus_format,
    "TGS Fiber": parse_tgs_format,
    "Tex-Star Communications": parse_texstar_format,
    "All Clear": parse_all_clear_format,
    "Pifer Quality Communications": parse_pifer_format,
    "North Sky": parse_north_sky_format,
}

def create_gui():
    app = TkinterDnD.Tk()
    app.title("Drop Excel File or Paste Schedule Text")
    app.geometry("600x800")

    label = tk.Label(app, text="Drag and drop your Excel file here", width=60, height=5, bg="lightgray")
    label.pack(padx=10, pady=10)

    textbox_frame = tk.Frame(app)
    textbox_frame.pack(fill="x", padx=10, pady=(0, 5))  # fixed vertical size

    textbox = tk.Text(textbox_frame, height=10, wrap="word")
    textbox.pack(fill="x", expand=False)

    
    global SELECTED_CONTRACTOR
    SELECTED_CONTRACTOR = tk.StringVar(value="(none)")
    
    dropdown_frame = tk.Frame(app)
    dropdown_frame.pack()
    tk.Label(dropdown_frame, text="Contractor Company:").pack(side="left", padx=5)
    dropdown_menu = ttk.Combobox(dropdown_frame, textvariable=SELECTED_CONTRACTOR, state="readonly")
    dropdown_menu["values"] = list(CONTRACTOR_LABELS.keys())
    dropdown_menu.current(0)
    dropdown_menu.pack(side="left")
    
    global HEADLESS_MODE
    HEADLESS_MODE = tk.BooleanVar(value=True)

    headless_frame = tk.Frame(app)
    headless_frame.pack(pady=5)
    tk.Checkbutton(headless_frame, text="Run Headless", variable=HEADLESS_MODE).pack()

    global WANT_FIRST_JOBS
    WANT_FIRST_JOBS = tk.BooleanVar(value=True)
    tk.Checkbutton(app, text="Show First Jobs Summary", variable=WANT_FIRST_JOBS).pack()

    # === HEADLESS MODE OUTPUT LOG ===
    log_output_frame = tk.Frame(app)
    log_output_frame.pack(fill="both", expand=True, padx=10, pady=(0, 10))
    global log_output_text
    log_output_text = tk.Text(log_output_frame, height=10, wrap="word", state="disabled", bg="#1e1e1e", fg="lime", font=("Consolas", 9))
    log_output_text.pack(padx=5, pady=5, fill="both", expand=True)
    
    def update_log_visibility(*args):
        if HEADLESS_MODE.get():
            log_output_frame.pack(fill="both", expand=True, padx=10, pady=(5, 10))
        else:
            log_output_frame.pack_forget()

    HEADLESS_MODE.trace_add("write", lambda *args: update_log_visibility())
    update_log_visibility()



    def drop(event):
        file_path = event.data.strip('{}')
        def threaded_process():
            try:
                df_test = pd.read_excel(file_path)
                assign_jobs(df_test)
            except Exception as e:
                gui_log(f"❌ Could not process file: {e}")
        threading.Thread(target=threaded_process, daemon=True).start()

    def parse_text():
        raw_text = textbox.get("1.0", tk.END).strip()
        if not raw_text:
            gui_log("No text to process.")
            return
        try:
            temp = reformat_contractor_text(raw_text)

            if not temp or not isinstance(temp, list):
                gui_log("❌ Parsed data is empty or malformed.")
                return

            # Failsafe: ensure all jobs have 'Date'
            for job in temp:
                if not isinstance(job, dict):
                    continue
                if 'Date' not in job or not job['Date']:
                    job['Date'] = (datetime.today() + timedelta(days=1)).strftime("%Y-%m-%d")

            df = pd.DataFrame(temp)

            if 'Date' not in df.columns:
                df['Date'] = (datetime.today() + timedelta(days=1)).strftime("%Y-%m-%d")

            # Parse dates, coerce errors to NaT
            df['Date'] = df['Date'].apply(flexible_date_parser)

            # Fill any blanks (NaT) with tomorrow's date
            tomorrow = datetime.today().date() + timedelta(days=1)
            df['Date'] = df['Date'].fillna(pd.Timestamp(tomorrow))
            df['Time'] = df['Time'].astype(str)

            unique_dates = sorted(df['Date'].dropna().dt.date.unique())
            start_date, end_date = ask_date_range(unique_dates)
            if not start_date or not end_date:
                gui_log("❌ No date range selected.")
                return

            filtered_df = df[(df['Date'].dt.date >= start_date) & (df['Date'].dt.date <= end_date)]

            if filtered_df.empty:
                gui_log("No matching jobs found.")
                return
            
            df['Dropdown'] = df['Tech']  # so it gets matched properly

            log(f"📦 Runtime headless check (before assign): {HEADLESS_MODE.get()}")
            def threaded_assign():
                assign_jobs(filtered_df)
                gui_log("✅ All jobs have been assigned.")

            threading.Thread(target=threaded_assign, daemon=True).start()

            # === Save log
            now = datetime.now()
            filename = f"Output{now.strftime('%m%d%H%M')}.txt"
            log_path = os.path.join(LOG_FOLDER, filename)
            with open(log_path, "w", encoding="utf-8") as f:
                f.write("\n".join(log_lines))

            gui_log(f"✅ Done processing pasted text.")

            first_jobs = build_first_jobs_summary(filtered_df, name_column="Tech")

            if WANT_FIRST_JOBS.get():
                app.after(0, lambda: save_and_open_html(jobs_to_html(first_jobs, SELECTED_CONTRACTOR.get())))

        except Exception as e:
            gui_log(f"❌ Error processing pasted text: {e}")

    btn = tk.Button(app, text="Parse & Assign from Pasted Text", command=parse_text)
    btn.pack(pady=5)

    label.drop_target_register(DND_FILES)
    label.dnd_bind('<<Drop>>', drop)

    app.mainloop()

if __name__ == "__main__":
    # This all works it just takes a second to run and is not useful currently
    #
    #
    # target_env = os.getenv("EMAIL_TARGET_DATE")  # e.g. “2025-05-27”
    # if target_env:
    #     try:
    #         target_date = datetime.fromisoformat(target_env).date()
    #     except ValueError:
    #         print("Invalid EMAIL_TARGET_DATE; should be YYYY-MM-DD")
    # else:
    #     target_date = date.today() + timedelta(days=1)
    # imap = None
    # try:
    #     imap = connect_imap()
    #     msg_nums = find_matching_msg_nums(imap, target_date)
    #     print(f"[Email] Found {len(msg_nums)} message(s) matching {target_date}:\n")
    #     for num in msg_nums:
    #         body = extract_relevant_section(fetch_body(imap, num))
    #         print(f"--- Email #{num.decode()} ---\n{body}\n")
    # except Exception as e:
    #     print(f"[Email] Error: {e}")
    # finally:
    #     if imap:
    #         try: imap.logout()
    #         except: pass
    try:
        create_gui()
    except Exception as e:
        print(f"\n❌ Fatal error: {e}")
        input("\nPress Enter to close...")
