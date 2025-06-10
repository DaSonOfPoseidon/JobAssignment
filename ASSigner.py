import sys
import os
import re
import time
import platform
import shutil
import tempfile
import pickle
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
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select, WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "../embedded_python/lib")))

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
        "frank": "Francisco Morales"
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
        assign_jobs_from_dataframe(df)
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
        log(failed_times[['Time']])

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

def normalize_name_key(name_input):
    name_input = name_input.strip().lower()
    parts = name_input.split()
    if len(parts) == 1:
        return parts[0]
    elif len(parts) >= 2:
        return f"{parts[0]} {parts[1][0]}"  # e.g., "jeff g"
    return name_input

def match_dropdown_option(select: Select, raw_name: str, contractor_full: str) -> str:
    corrected_name = get_corrected_name(raw_name.strip(), contractor_full)
    key_parts = corrected_name.lower().split()
    first_name = key_parts[0] if key_parts else ""
    last_initial = key_parts[1][0] if len(key_parts) > 1 else ""

    # Full name exact match
    for option in select.options:
        if option.text.strip().lower() == corrected_name.lower():
            return option.text

    # First name + last initial match
    for option in select.options:
        full_parts = option.text.lower().strip().split()
        if len(full_parts) >= 2 and full_parts[0].startswith(first_name) and full_parts[1][0] == last_initial:
            return option.text

    # Loose first name match
    potential_matches = [opt.text for opt in select.options if opt.text.lower().startswith(first_name)]
    if len(potential_matches) == 1:
        return potential_matches[0]

    return None  # fallback — caller must handle

def get_corrected_name(name_input, contractor_full):
    if not name_input:
        return name_input

    key = normalize_name_key(name_input)
    
    # First, try contractor-specific corrections
    corrections = CONTRACTOR_NAME_CORRECTIONS.get(contractor_full)
    if corrections and key in corrections:
        return corrections[key]
    
    # Fallback to first word if "brandon t" isn't found but "brandon" exists
    key_parts = key.split()
    if corrections and key_parts and key_parts[0] in corrections:
        return corrections[key_parts[0]]

    # Try fuzzy match on contractor corrections
    if corrections:
        best_match = None
        highest_score = 0
        for known_key in corrections:
            score = fuzz.ratio(key_parts[0], known_key)
            if score > 90 and score > highest_score:
                best_match = known_key
                highest_score = score
        if best_match:
            return corrections[best_match]

    # Fallback to default corrections
    fallback = CONTRACTOR_NAME_CORRECTIONS.get("default", {})
    return fallback.get(key_parts[0], name_input)

def login_failed(driver):
    try:
        return (
            "login.php" in driver.current_url
            or "Username" in driver.page_source
            or "Invalid username or password" in driver.page_source
        )
    except Exception:
        return True  # if we can't read the page, assume failure

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
    <div>Cole - 8am - </div><br>
    <div>Matt - 8am - </div><br>
    <div>Tylor - 8am - </div>
    </div>
    """
    
    return html_header + "\n".join(lines) + html_footer

def show_first_jobs(first_jobs):
    from tkinter import Toplevel, Scrollbar, Text, RIGHT, Y, END, Button

    popup = Toplevel()
    popup.title("First Jobs Summary")
    popup.geometry("700x600")

    text = Text(popup, wrap="word", font=("Apatos", 11))
    scrollbar = Scrollbar(popup, command=text.yview)
    text.configure(yscrollcommand=scrollbar.set)

    text.pack(side="left", fill="both", expand=True)
    scrollbar.pack(side=RIGHT, fill=Y)

    # --- Custom Header Block ---
    company = SELECTED_CONTRACTOR.get()
    now = datetime.now()
    timestamp = now.strftime("%I:%M%p, %m/%d").lower()
    main_header = (
        f"Please call into 163 for tech assistance - Please call 633 for dispatcher\n\n"
        f"This is a reminder. Please make sure we are returning the HST's to the office once we are done using them. Please also make sure you are updating WO notes and completing the WO after completing the dispatch/install.\n\n"
        f"As of {timestamp} the jobs are set as follows.....\n\n"
        f"{company} Contractors - \n\n"
    )
    text.insert(END, main_header)

    # --- Main Jobs Block ---
    # Only show jobs for the selected contractor!
    for day, jobs in first_jobs.items():
        for j in jobs:
            text.insert(END, f"{j}\n\n")
        text.insert(END, "\n")

    # --- Copy to clipboard button ---
    popup.clipboard_clear()
    popup.clipboard_append(text.get("1.0", END).strip())

    def copy_to_clipboard():
        popup.clipboard_clear()
        popup.clipboard_append(text.get("1.0", END).strip())

    copy_button = Button(popup, text="Copy to Clipboard", command=copy_to_clipboard)
    copy_button.pack(pady=5)


def verify_work_order_page(driver, wo_number, url, max_attempts=3):
    for attempt in range(1, max_attempts + 1):
        try:
            if not driver.current_url.strip().endswith(wo_number):
                driver.get(url)
            displayed_wo_elem = WebDriverWait(driver, 10, poll_frequency=0.5).until(
                EC.presence_of_element_located((By.XPATH, "//td[contains(text(), 'Work Order #:')]/following-sibling::td"))
            )
            if displayed_wo_elem.text.strip() == wo_number:
                return True
        except Exception:
            log(f"🟡 Attempt {attempt}: Unable to verify WO #{wo_number}")
        time.sleep(5)
    return False

def process_workorders(file_path):
    gui_log(f"\nProcessing file: {file_path}")
    df_raw = pd.read_excel(file_path)

    df = pd.DataFrame()
    df['Date'] = df_raw.iloc[:, COLUMN_DATE].apply(flexible_date_parser)
    df['Time'] = df_raw.iloc[:, COLUMN_TIME].astype(str)
    df['Name'] = df_raw.iloc[:, COLUMN_NAME]
    df['Type'] = df_raw.iloc[:, COLUMN_TYPE]
    df['WO'] = df_raw.iloc[:, COLUMN_WO]
    df['Address'] = df_raw.iloc[:, COLUMN_ADDRESS]
    df['Dropdown'] = df_raw.iloc[:, COLUMN_DROPDOWN]

    df = df.dropna(subset=['Date', 'WO', 'Dropdown'])

    unique_dates = sorted(df['Date'].dropna().dt.date.unique())
    start_date, end_date = ask_date_range(unique_dates)
    if not start_date or not end_date:
        gui_log("❌ No date range selected.")
        return

    filtered_df = df[(df['Date'].dt.date >= start_date) & (df['Date'].dt.date <= end_date)]

    if filtered_df.empty:
        gui_log("No matching jobs found for that date range.")
        input("\nPress Enter to close...")
        return

    gui_log(f"\nProcessing {len(filtered_df)} work orders...")
    
    driver = create_driver(is_headless())
    handle_login(driver)

    for index, row in filtered_df.iterrows():
        raw_wo = row['WO']
        raw_name = str(row['Dropdown']).strip()
        name_parts = raw_name.split()
        if len(name_parts) >= 2:
            dropdown_value = f"{name_parts[0].capitalize()} {name_parts[1][0].upper()}"
        else:
            dropdown_value = raw_name.capitalize()

        try:
            wo_number = str(int(raw_wo))
        except (ValueError, TypeError):
            excel_row = index + 2
            gui_log(f"❌ Invalid WO number '{raw_wo}' on spreadsheet line {excel_row}.")
            continue

        url = BASE_URL + wo_number
        log(f"\n🔗 Opening WO #{wo_number} — {url}")
        driver.get(url)

        if not verify_work_order_page(driver, wo_number, url):
            gui_log(f"❌ Failed to verify WO #{wo_number}. Skipping.")
            continue

        desired_contractor_label = SELECTED_CONTRACTOR.get()
        desired_contractor_full = CONTRACTOR_LABELS.get(desired_contractor_label)

        if desired_contractor_full:
            assign_contractor(driver, wo_number, desired_contractor_full)

        try:
            dropdown = WebDriverWait(driver, 60).until(
                EC.presence_of_element_located((By.ID, "AssignEmpID"))
            )
            select = Select(dropdown)
            matched_option = match_dropdown_option(select, dropdown_value, desired_contractor_full)

            if not matched_option:
                gui_log(f"❌ No dropdown match for '{dropdown_value}' — skipping WO #{wo_number}", level="warning")
                continue

            # Check and remove other assignments first
            assignments_div = WebDriverWait(driver, 5).until(
                EC.presence_of_element_located((By.ID, "AssignmentsList"))
            )
            assigned_rows = assignments_div.find_elements(By.XPATH, ".//tr")

            already_assigned = False

            for row in assigned_rows:
                try:
                    row_text = row.text.strip().lower()
                    if matched_option.lower() in row_text:
                        already_assigned = True
                    else:
                        # Remove incorrect assignment
                        remove_link = row.find_element(By.XPATH, ".//a[contains(@onclick, 'removeWorkOrderAssignment')]")
                        driver.execute_script("arguments[0].scrollIntoView(true);", remove_link)
                        driver.execute_script("arguments[0].click();", remove_link)

                        # Confirm alert
                        try:
                            alert = WebDriverWait(driver, 2).until(EC.alert_is_present())
                            alert_text = alert.text
                            log(f"⚠️ Alert: {alert_text}")
                            alert.accept()
                            log("✅ Removed incorrect tech assignment.")
                        except:
                            pass
                except Exception as e:
                    log(f"⚠️ Skipping row removal due to error: {e}")

            if already_assigned:
                log(f"🟡 WO #{wo_number}: '{matched_option}' already assigned.")
                continue


            select.select_by_visible_text(matched_option)
            add_button = WebDriverWait(driver, 60).until(
                EC.element_to_be_clickable((By.XPATH, "//button[contains(@class, 'Socket')]"))
            )
            add_button.click()

        except Exception as e:
            gui_log(f"❌ Error on WO #{wo_number}: {e}")

    now = datetime.now()
    filename = f"Output{now.strftime('%m%d%H%M')}.txt"
    log_path = os.path.join(LOG_FOLDER, filename)
    with open(log_path, "w", encoding="utf-8") as f:
        f.write("\n".join(log_lines))

    gui_log(f"\n✅ Done processing work orders.")
    gui_log(f"🗂️ Output saved to: {log_path}")

    first_jobs = build_first_jobs_summary(filtered_df, name_column="Dropdown")
    company = SELECTED_CONTRACTOR.get()
    html_str = jobs_to_html(first_jobs, company)
    save_and_open_html(html_str)

def create_driver(headless: bool = True) -> webdriver.Chrome:
    # 1) Locate the browser binary
    chrome_bin = (
        os.environ.get("CHROME_BIN")
        or shutil.which("google-chrome")
        or shutil.which("chrome")
        or shutil.which("chromium-browser")
        or shutil.which("chromium")
    )
    if not chrome_bin and platform.system() == "Windows":
        # Probe standard Windows install paths
        for p in (
            r"C:\Program Files\Google\Chrome\Application\chrome.exe",
            r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
        ):
            if os.path.exists(p):
                chrome_bin = p
                break

    if not chrome_bin or not os.path.exists(chrome_bin):
        raise RuntimeError(
            "Chrome/Chromium binary not found—install it or set CHROME_BIN"
        )

    # 2) Build ChromeOptions
    opts = webdriver.ChromeOptions()
    opts.binary_location = chrome_bin
    if headless:
        opts.add_argument("--headless=new")
        opts.add_argument("--disable-gpu")
        opts.add_argument("--no-sandbox")
        opts.add_argument("--disable-usb-keyboard-detect")
        opts.add_argument("--disable-hid-detection")
    opts.add_argument("--log-level=3")
    opts.add_argument("--disable-dev-shm-usage")
    opts.set_capability("unhandledPromptBehavior", "dismiss")
    opts.page_load_strategy = "eager"

    # 3) Pick driver based on platform
    system = platform.system()
    arch = platform.machine().lower()
    if system == "Linux" and ("arm" in arch or "aarch64" in arch):
        # Raspberry Pi / ARM Linux → use distro’s chromedriver
        driver_path = "/usr/bin/chromedriver"
        if not os.path.exists(driver_path):
            raise RuntimeError("ARM chromedriver not found; please `apt install chromium-driver`")
        service = Service(driver_path)
    else:
        # x86 Linux or Windows → download via webdriver-manager
        service = Service(ChromeDriverManager().install())

    # 4) Launch
    return webdriver.Chrome(service=service, options=opts)

def assign_jobs_from_dataframe(df):
    gui_log(f"Processing {len(df)} work orders from pasted text...")
    driver = create_driver(is_headless())
    handle_login(driver)

    for index, row in df.iterrows():
        raw_wo = row['WO']
        raw_name = str(row.get('Dropdown') or row.get('Tech', '')).strip()
        raw_name = re.sub(r"^[\-\s]+", "", raw_name)  # Remove leading dashes/spaces

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
        driver.get(url)

        # Check if WO is unscheduled
        try:
            error_elem = driver.find_element(By.CLASS_NAME, "errors")
            if "no scheduled dates found" in error_elem.text.lower():
                gui_log(f"❌ WO {wo_number} is not scheduled — skipping.")
                continue
        except:
            pass  # no error block found, continue

        if not verify_work_order_page(driver, wo_number, url):
            gui_log(f"❌ Failed to verify WO #{wo_number}. Skipping.")
            continue

        desired_contractor_label = SELECTED_CONTRACTOR.get()
        desired_contractor_full = CONTRACTOR_LABELS.get(desired_contractor_label)
        if desired_contractor_full:
            assign_contractor(driver, wo_number, desired_contractor_full)

        try:
            dropdown = WebDriverWait(driver, 60).until(
                EC.presence_of_element_located((By.ID, "AssignEmpID"))
            )
            select = Select(dropdown)
            matched_option = match_dropdown_option(select, dropdown_value.strip(), desired_contractor_full)
            if not matched_option:
                gui_log(f"❌ No dropdown match for '{dropdown_value}' — skipping WO #{wo_number}")
                continue

            assignments_div = WebDriverWait(driver, 5).until(
                EC.presence_of_element_located((By.ID, "AssignmentsList"))
            )
            assigned_names = assignments_div.text.lower()

            if matched_option.lower() in assigned_names:
                log(f"🟡 WO #{wo_number}: '{matched_option}' already assigned.")
                continue

            select.select_by_visible_text(matched_option)
            add_button = WebDriverWait(driver, 60).until(
                EC.element_to_be_clickable((By.XPATH, "//button[contains(@class, 'Socket')]"))
            )
            add_button.click()
            gui_log(f"WO {wo_number} - Assigned to {matched_option} with {desired_contractor_full}")


        except Exception as e:
            gui_log(f"❌ Error assigning WO #{wo_number}: {e}")

def assign_contractor(driver, wo_number, desired_contractor_full):
    try:
        # 🧠 Trigger the assignment UI manually using JavaScript
        try:
            driver.execute_script(f"assignContractor('{wo_number}');")
            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, "ContractorID"))
            )
            time.sleep(0.5)  # let modal settle
        except Exception as e:
            gui_log(f"Could not trigger assignContractor JS on WO #{wo_number}: {e}")
            return

        # ✅ Get current contractor assignment from the page
        current_contractor = get_contractor_assignments(driver)
        if current_contractor == desired_contractor_full:
            log(f"✅ Contractor '{current_contractor}' already assigned to WO #{wo_number}")
            return  # No change needed

        log(f"🧹 Reassigning from '{current_contractor}' → '{desired_contractor_full}'")

        # 🧽 Remove the currently assigned contractor (if any)
        try:
            remove_links = driver.find_elements(By.XPATH, "//a[contains(@onclick, 'removeContractor')]")
            for link in remove_links:
                driver.execute_script("arguments[0].scrollIntoView(true);", link)
                driver.execute_script("arguments[0].click();", link)
                time.sleep(1)
        except Exception as e:
            gui_log(f"❌ Could not remove existing contractor on WO #{wo_number}: {e}")

        # 🏷️ Assign the new contractor
        try:
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "ContractorID")))
            contractor_dropdown = Select(driver.find_element(By.ID, "ContractorID"))
            contractor_dropdown.select_by_visible_text(desired_contractor_full)

            role_dropdown = Select(driver.find_element(By.ID, "ContractorType"))
            role_dropdown.select_by_visible_text("Primary")

            # Handle potential overlay (e.g. FileList blocking Assign button)
            try:
                file_list_elem = driver.find_element(By.ID, "FileList")
                driver.execute_script("arguments[0].style.display = 'none';", file_list_elem)
            except:
                pass  # it's okay if FileList isn't present

            assign_button = driver.find_element(By.XPATH, "//input[@type='button' and @value='Assign']")
            driver.execute_script("arguments[0].click();", assign_button)

            log(f"🏷️ Assigned contractor '{desired_contractor_full}' to WO #{wo_number}")
        except Exception as e:
            log(f"❌ Failed to assign contractor on WO #{wo_number}: {e}")

    except Exception as e:
        gui_log(f"❌ Contractor assignment process failed for WO #{wo_number}: {e}")

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

def get_contractor_assignments(driver):
    try:
        # Wait for the contractor section to load
        WebDriverWait(driver, 5).until(
            lambda d: "Primary" in d.find_element(By.CLASS_NAME, "contractorsection").text
        )

        # Find ALL <b> tags and look for the one that contains " - (Primary"
        for elem in driver.find_elements(By.TAG_NAME, "b"):
            text = elem.text.strip()
            if " - (Primary" in text:
                contractor = text.split(" - ")[0].strip()
                #log(f"🔍 Detected contractor: {contractor}")
                return contractor

        log("⚠️ No 'Primary' contractor detected.")
        return "Unknown"
        
    except Exception as e:
        gui_log(f"❌ Could not find contractor name: {e}")
        return "Unknown"

def is_headless():
    try:
        return HEADLESS_MODE.get()
    except:
        return False

def save_cookies(driver, filename="cookies.pkl"):
    with cookie_lock:
        with open(COOKIE_PATH, "wb") as f:
            pickle.dump(driver.get_cookies(), f)

def load_cookies(driver, filename="cookies.pkl"):
    if not os.path.exists(COOKIE_PATH): return False
    try:
        with cookie_lock:
            with open(COOKIE_PATH, "rb") as f:
                cookies = pickle.load(f)

        driver.get("http://inside.sockettelecom.com/")
        for cookie in cookies:
            driver.add_cookie(cookie)
        driver.refresh()
        clear_first_time_overlays(driver)
        return True
    except Exception:
        if os.path.exists(COOKIE_PATH): os.remove(COOKIE_PATH)
        return False

def save_and_open_html(html_str, filename="FirstJobsSummary.html"):
    # Write HTML to file
    with open(filename, "w", encoding="utf-8") as f:
        f.write(html_str)
    # Open in the default browser (works everywhere)
    file_url = 'file://' + os.path.realpath(filename)
    webbrowser.open(file_url)

def handle_login(driver):
    driver.get("http://inside.sockettelecom.com/")

    if load_cookies(driver):
        if not login_failed(driver):
            log("✅ Session restored via cookies.")
            clear_first_time_overlays(driver)
            return 

    # Cookies failed or expired, now try credentials
    USERNAME, PASSWORD = check_env_or_prompt_login(log)

    while True:
        perform_login(driver, USERNAME, PASSWORD)
        time.sleep(2)
        if not login_failed(driver):
            save_cookies(driver)
            return
        else:
            gui_log("❌ Login failed. Please re-enter your credentials.")
            USERNAME, PASSWORD = prompt_for_credentials()
            save_env_credentials(USERNAME, PASSWORD)

def perform_login(driver, USERNAME, PASSWORD):
    driver.get("http://inside.sockettelecom.com/system/login.php")
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.NAME, "username")))
    driver.find_element(By.NAME, "username").send_keys(USERNAME)
    driver.find_element(By.NAME, "password").send_keys(PASSWORD)
    driver.find_element(By.ID, "login").click()
    clear_first_time_overlays(driver)
    gui_log("Logged in successfully")

def clear_first_time_overlays(driver):
    # Dismiss alert if present
    try:
        WebDriverWait(driver, 0.5).until(EC.alert_is_present())
        driver.switch_to.alert.dismiss()
    except:
        pass

    # Known popup buttons
    buttons = [
        "//form[@id='valueForm']//input[@type='button']",
        "//form[@id='f']//input[@type='button']"
    ]
    for xpath in buttons:
        try:
            WebDriverWait(driver, 0.5).until(EC.element_to_be_clickable((By.XPATH, xpath))).click()
        except:
            pass

    # Iframe switch loop
    for _ in range(3):
        try:
            WebDriverWait(driver, 0.5).until(EC.frame_to_be_available_and_switch_to_it((By.ID, "MainView")))
            return
        except:
            time.sleep(0.25)
    log("❌ Could not switch to MainView iframe.")

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
        if re.fullmatch(r"[A-Z]{3,}", line):
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
    jobs = []

    current_tech = None
    current_date = None
    current_time = None

    job_line_pattern = re.compile(
        r"-\s*(.*?)\s*-\s*[\d\-]+?\s*-\s*(.*?)\s*-\s*(.*?)\s*-\s*WO\s*?(\d+)",
        re.IGNORECASE
    )

    for i, raw_line in enumerate(lines):
        line = raw_line.strip()
        if not line:
            continue

        # Tech name (capitalized full names)
        if re.fullmatch(r"[A-Z][a-z]+ [A-Z][a-z]+", line):
            current_tech = line
            continue

        # Date (e.g. 5-5-25 or 4/28/25)
        if re.match(r"\d{1,2}[-/]\d{1,2}[-/]\d{2,4}", line):
            current_date = line
            continue

        # Time block (e.g. 8AM, 10AM)
        if re.fullmatch(r"\d{1,2}(:\d{2})?\s?(AM|PM)?", line, re.IGNORECASE):
            current_time = format_time_str(line)
            continue

        # Full job line with embedded time (optional)
        if re.match(r"\d{1,2}:\d{2}", line):
            # Pull time from the front of the line
            embedded_time = re.match(r"(\d{1,2}:\d{2})", line).group(1)
            current_time = format_time_str(embedded_time)
            # Remove time prefix before parsing
            line = re.sub(r"^\d{1,2}:\d{2}\s*-\s*", "- ", line)

        match = job_line_pattern.search(line)
        if match:
            name, job_type, address, wo_number = match.groups()
            jobs.append({
                "Date": current_date,
                "Time": current_time,
                "Name": name.strip(),
                "Type": job_type.strip(),
                "WO": wo_number.strip(),
                "Address": address.strip(),
                "Tech": current_tech
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

CONTRACTOR_FORMAT_PARSERS = {
    "Subterraneus Installs": parse_subterraneus_format,
    "TGS Fiber": parse_tgs_format,
    "Tex-Star Communications": parse_texstar_format,
    "All Clear": parse_all_clear_format,
    "Pifer Quality Communications": parse_pifer_format
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
                process_workorders(file_path)
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
                gui_log("❌ 'Date' column missing after parsing. Aborting.")
                return

            df['Date'] = df['Date'].apply(flexible_date_parser)
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
                assign_jobs_from_dataframe(filtered_df)
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
