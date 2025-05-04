import sys
import os
import re
import time
import tempfile
import pickle
from datetime import datetime, timedelta
from collections import defaultdict
from tkinter import messagebox
from dotenv import load_dotenv, set_key
import pandas as pd
from rapidfuzz import fuzz
from threading import Thread
import threading
import tkinter as tk
from tkcalendar import DateEntry
from tkinter import simpledialog
from tkinter import ttk
from tkinterdnd2 import DND_FILES, TkinterDnD
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select, WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "../embedded_python/lib")))

# === CONFIGURATION ===
SHOW_ALL_OUTPUT_IN_CONSOLE = True
CHROMEDRIVER_PATH = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "chromedriver.exe"))
LOG_FOLDER = os.path.join(os.path.dirname(__file__), "..", "logs")
os.makedirs(LOG_FOLDER, exist_ok=True)
load_dotenv(dotenv_path=".env")
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

def ask_single_date(available_dates=None, title="Select Job Date"):
    top = tk.Toplevel()
    top.title(title)
    top.geometry("250x120")

    tk.Label(top, text="Job Date:").pack(pady=(10, 0))

    try:
        available_dates = sorted(d for d in available_dates if d)
        mindate = min(available_dates)
        maxdate = max(available_dates)
        default_date = available_dates[0]
    except (ValueError, TypeError):
        mindate = None
        maxdate = None
        default_date = datetime.today().date()

    cal = DateEntry(top, width=12, background='darkblue', foreground='white', borderwidth=2, mindate=mindate, maxdate=maxdate)
    cal.set_date(default_date)
    cal.pack(pady=5)

    selected = {}

    def submit():
        selected["date"] = cal.get_date()
        top.destroy()

    tk.Button(top, text="OK", command=submit).pack()
    top.grab_set()
    top.wait_window()

    return selected.get("date")

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

def show_first_jobs(first_jobs):
    from tkinter import Toplevel, Scrollbar, Text, RIGHT, Y, END, Button

    popup = Toplevel()
    popup.title("First Jobs Summary")
    popup.geometry("600x400")

    text = Text(popup, wrap="word")
    scrollbar = Scrollbar(popup, command=text.yview)
    text.configure(yscrollcommand=scrollbar.set)

    text.pack(side="left", fill="both", expand=True)
    scrollbar.pack(side=RIGHT, fill=Y)

    all_text = ""

    for day, jobs in first_jobs.items():
        all_text += f"{day.strftime('%A %m/%d/%Y')}\n"
        for j in jobs:
            all_text += f"{j}\n"
        all_text += "\n"

    text.insert(END, all_text)
    popup.clipboard_clear()
    popup.clipboard_append(all_text.strip())

    def copy_to_clipboard():
        popup.clipboard_clear()
        popup.clipboard_append(all_text.strip())

    copy_button = Button(popup, text="Copy to Clipboard", command=copy_to_clipboard)
    copy_button.pack(pady=5)

def process_jobs_from_list(job_list):
    df = pd.DataFrame(job_list)

    df['Date'] = df['Date'].apply(flexible_date_parser)
    df['Time'] = df['Time'].astype(str)
    
    tomorrow = datetime.today().date() + timedelta(days=1)
    df['Date'] = df['Date'].fillna(pd.Timestamp(tomorrow))      

    df = df.dropna(subset=['Date', 'WO'])

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

    # Save to temporary Excel file and reuse the existing function
    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
        temp_path = tmp.name
        filtered_df.to_excel(temp_path, index=False)
        log(f"\n📄 Temporary Excel created: {temp_path}")

    process_workorders(temp_path)

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
    options = webdriver.ChromeOptions()
    if is_headless():
        options.add_argument("--no-sandbox")
        options.add_argument("--disable-dev-shm-usage")
        options.add_argument("--disable-gpu")
        options.add_argument("--disable-extensions")
        options.add_argument("--headless=new")
        options.add_argument("--window-size=1920,1080")
    else:
        options.add_argument("--start-maximized")
    
    driver = webdriver.Chrome(service=Service(CHROMEDRIVER_PATH), options=options)
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
    show_first_jobs(first_jobs)

def parse_undated_tabular_with_wo(lines):
    jobs = []
    for line in lines:
        parts = re.split(r'\t+|\s{2,}', line.strip())
        if len(parts) < 6:
            continue

        time = parts[0]
        name = parts[1]
        job_type = parts[2]
        wo_field = parts[3]
        address = parts[4]
        tech = parts[5]

        wo_match = re.search(r"\bWO\s*(\d{6})\b", wo_field)
        if not wo_match:
            continue

        jobs.append({
            "Date": "",  # no date in this format
            "Time": time.strip(),
            "Name": name.strip(),
            "Type": job_type.strip(),
            "WO": wo_match.group(1),
            "Address": address.strip(),
            "Tech": tech.strip()
        })
    return jobs

def assign_jobs_from_dataframe(df):
    gui_log(f"Processing {len(df)} work orders from pasted text...")

    options = webdriver.ChromeOptions()
    if is_headless():
        options.add_argument("--headless=new")  # or just "--headless" if issues
        options.add_argument("--window-size=1920,1080")
        options.add_argument("--no-sandbox")
        options.add_argument("--disable-dev-shm-usage")
        options.add_argument("--disable-gpu")
        options.add_argument("--disable-extensions")

    else:
        options.add_argument("--start-maximized")

    driver = webdriver.Chrome(service=Service(CHROMEDRIVER_PATH), options=options)
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
            matched_option = None

            # Get corrected name with contractor context
            corrected_name = get_corrected_name(dropdown_value.strip(), desired_contractor_full)
            key_parts = corrected_name.strip().lower().split()

            first_name = key_parts[0] if key_parts else ""
            last_initial = key_parts[1][0] if len(key_parts) > 1 else ""

            # Try full corrected name match
            for option in select.options:
                if option.text.lower().strip() == corrected_name.lower():
                    matched_option = option.text
                    break

            # Try first + last initial match
            if not matched_option:
                for option in select.options:
                    full_parts = option.text.lower().strip().split()
                    if len(full_parts) >= 2 and full_parts[0].startswith(first_name) and full_parts[1][0] == last_initial:
                        matched_option = option.text
                        break

            # Try loose match by first name
            if not matched_option:
                matches = [opt.text for opt in select.options if opt.text.lower().startswith(first_name)]
                if len(matches) == 1:
                    matched_option = matches[0]

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

def parse_dated_tabular_with_city(lines):
    jobs = []
    for i, line in enumerate(lines, start=1):
        # split by tabs or 2+ spaces
        parts = re.split(r'\t+|\s{2,}', line.strip())
        parts = [p.strip() for p in parts]

        if len(parts) < 8:
            print(f"[❌ MALFORMED LINE {i}] Too few parts ({len(parts)}): {line}")
            continue

        date = parts[0]
        time = parts[1]
        name = parts[2]
        job_type = parts[3]
        wo_field = parts[4]
        address = parts[5]
        city = parts[6]
        tech = parts[7]

        wo_match = re.search(r"\bWO\s*(\d{6})\b", wo_field)
        if not wo_match:
            print(f"[❌ MALFORMED LINE {i}] Could not extract WO from: '{wo_field}' — full line: {line}")
            continue

        wo = wo_match.group(1)
        jobs.append({
            "Date": date.strip(),
            "Time": time.strip(),
            "Name": name.strip(),
            "Type": job_type.strip(),
            "WO": wo.strip(),
            "Address": f"{address.strip()}, {city.strip()}",
            "Tech": tech.strip()
        })

    print(f"[✅ Parsed] {len(jobs)} valid job(s) from dated tabular format.")
    return jobs

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

def parse_dmy_tabular_variant(lines):
    jobs = []
    print("[DEBUG] Triggering parse_dmy_tabular_variant()")
    for i, line in enumerate(lines):
        parts = re.split(r'\t+|\s{2,}', line.strip())
        original_parts = parts.copy()

        if len(parts) == 7 and re.match(r"\d{1,2}:\d{2} ?[APap][Mm] .+", parts[1]):
            # Split the hybrid time-name field
            match = re.match(r"(\d{1,2}:\d{2} ?[APap][Mm])\s+(.+)", parts[1])
            if match:
                time_val = match.group(1)
                name_val = match.group(2)
                parts = [parts[0], time_val, name_val] + parts[2:]

        if len(parts) != 8:
            print(f"[WARN] Skipping line with insufficient parts: {original_parts}")
            continue

        date, time, name, job_type, wo_field, address, city, tech = parts

        wo_match = re.search(r"\bWO\s*(\d{6})\b", wo_field)
        if not wo_match:
            print(f"[WARN] Skipping line with invalid WO: {wo_field}")
            continue

        jobs.append({
            "Date": date.strip(),
            "Time": time.strip(),
            "Name": name.strip(),
            "Type": job_type.strip(),
            "WO": wo_match.group(1),
            "Address": f"{address.strip()}, {city.strip()}",
            "Tech": tech.strip()
        })

    return jobs

def looks_like_grouped_tech_date_format(lines):
    date_pattern = re.compile(r"\d{1,2}[-/]\d{1,2}([-/]\d{2,4})?")
    for i in range(len(lines) - 1):
        name_line = lines[i].strip()
        date_line = lines[i + 1].strip()
        if name_line.isalpha() and date_pattern.match(date_line):
            return True
    return False

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
    jobs = []

    current_tech = None
    current_date = None
    current_time = None

    print(f"[DEBUG] Raw line count: {len(lines)}")
    for i, l in enumerate(lines[:5], 1):
        print(f"[Line {i}] {repr(l)}")

    # === Format: Dated tabular with city + tech (tabs or clean spacing)
    if all("WO" in l and len(re.split(r'\t+|\s{2,}', l)) >= 8 for l in lines[:5]):
        print("[DEBUG] Triggering parse_dated_tabular_with_city()")
        return parse_dated_tabular_with_city(lines)

    def detect_undated_tabular_with_wo(lines):
        valid_lines = 0
        for line in lines[:5]:
            parts = re.split(r'\t+|\s{2,}', line.strip())
            if len(parts) >= 6 and any("WO" in p for p in parts):
                valid_lines += 1
        return valid_lines >= 4
    
    # Match d-MMM formats like '2-May'
    dmy_pattern = re.compile(r"^\d{1,2}-[A-Za-z]{3}$")
    if all(dmy_pattern.match(re.split(r'\t+|\s{2,}', l)[0]) for l in lines[:3]):
        print("[DEBUG] Triggering parse_dmy_tabular_variant()")
        return parse_dmy_tabular_variant(lines)


    is_grouped_inline_with_tech = all("WO" in l and l.count(" - ") >= 6 for l in lines[:5])
    date_pattern = re.compile(r"\d{1,2}/\d{1,2}/\d{4}")
    date_alt_pattern = re.compile(r"\d{1,2}-[A-Za-z]{3}")  # 21-Apr format
    time_block_pattern = re.compile(r"^\d{1,2}( AM| PM)?$", re.IGNORECASE)
    job_line_pattern = re.compile(r"(\d{1,2}:\d{2})\s*-\s*(.*?)\s*-\s*(\d{4}-\d{4}-\d{4})\s*-\s*(.*?)\s*-\s*(.*?)\s*-\s*WO\s*(\d+)", re.IGNORECASE)
    is_subt_tabular = all(re.match(r"\d{1,2}/\d{1,2}/\d{4}", l) for l in lines[:3]) and all("\t" in l for l in lines[:3])
    is_dated_tabular_with_city = all("WO" in l and l.count("  ") >= 3 for l in lines[:5])

    def format_tech_name(name):
        return name.capitalize() if name.isupper() else name

    date_found = False
    # Detect format early to avoid fallback hijacking
    if looks_like_grouped_tech_date_format(lines):
        print("[DEBUG] Triggering parse_grouped_tech_date_schedule()")
        return parse_grouped_tech_date_schedule(lines, selected_date=None)

    if detect_undated_tabular_with_wo(lines):
        print("[DEBUG] Triggering parse_undated_tabular_with_wo()")
        return parse_undated_tabular_with_wo(lines)

    for line in lines:
        if not line:
            continue

        # Handle SubT format
        parts = re.split(r"\t+| {2,}", line)
        if len(parts) >= 8:
            raw_date = parts[0].strip()
            raw_time = parts[1].strip()
            name = parts[2].strip()
            job_type = parts[3].strip()
            wo_field = parts[4].strip()
            address = parts[5].strip()
            city = parts[6].strip()
            tech = format_tech_name(parts[7].strip())

            wo = re.search(r"\bWO\s*(\d{6})\b", wo_field)

            try:
                if "-" in raw_date and len(raw_date.split("-")) == 2:
                    parsed_date = datetime.strptime(raw_date + f"-{datetime.today().year}", "%d-%b-%Y").date()
                else:
                    parsed_date = datetime.strptime(raw_date, "%m/%d/%Y").date()
                date_found = True
            except:
                parsed_date = None

            jobs.append({
                "Tech": tech,
                "Date": parsed_date,
                "Time": raw_time,
                "Name": name,
                "Account": "",
                "Type": job_type,
                "Address": f"{address}, {city}",
                "WO": wo
            })
            continue

        # Tech name detection (all caps or capitalized)
        if re.match(r"^[A-Z][A-Za-z\s]+$", line):
            current_tech = format_tech_name(line.strip())
            continue

        # Date detection
        if date_pattern.match(line):
            try:
                current_date = datetime.strptime(line.strip(), "%m/%d/%Y").date()
                date_found = True
                continue
            except:
                pass
        elif date_alt_pattern.match(line):
            try:
                current_date = datetime.strptime(line.strip() + f"-{datetime.today().year}", "%d-%b-%Y").date()
                date_found = True
                continue
            except:
                pass
        elif is_subt_tabular:
            jobs = parse_subt_tabular_variant(lines)


        # Time block header (ignored)
        if time_block_pattern.match(line):
            continue
        
        if is_grouped_inline_with_tech:
            jobs = []
            pattern = re.compile(r"^(.*?) - (.*?) - (.*?) - (.*?) - (.*?) - WO (\d{6}) - (.+)$", re.IGNORECASE)
            for line in lines:
                match = pattern.match(line)
                if match:
                    time_str, name, acc, job_type, address, wo, tech = match.groups()
                    jobs.append({
                        "Time": time_str.strip(),
                        "Name": name.strip(),
                        "Account": acc.strip(),
                        "Type": job_type.strip(),
                        "Address": address.strip(),
                        "WO": wo.strip(),
                        "Tech": tech.strip(),
                        "Date": current_date or "",  # <- optional fallback if needed
                    })
            return jobs
        
        if is_dated_tabular_with_city:
            return parse_dated_tabular_with_city(lines)

        match = job_line_pattern.match(line)
        if match:
            time_str, name, acc, job_type, address, wo = match.groups()

            # Try parsing time
            try:
                dt_obj = datetime.strptime(time_str.strip(), "%H:%M")
                if dt_obj.hour < 8:
                    dt_obj = dt_obj.replace(hour=dt_obj.hour + 12)  # assume PM
                formatted_time = dt_obj.strftime("%H:%M")
            except:
                formatted_time = time_str.strip()

            jobs.append({
                "Tech": current_tech or "",
                "Date": current_date or "",
                "Time": formatted_time,
                "Name": name.strip(),
                "Account": acc.strip(),
                "Type": job_type.strip(),
                "Address": address.strip(),
                "WO": wo.strip()
            })

    if not date_found:
        # Prompt user for date if not found
        tomorrow = datetime.today().date() + timedelta(days=1)
        fallback_date = ask_date_with_default(tomorrow)
        for job in jobs:
            job["Date"] = fallback_date
    
    if all('\t' in line for line in lines[:3]) and len(re.split(r'\t+', lines[0])) == 7:
        return parse_subt_tabular_variant(lines)
    
    if any(re.match(r"^\d{1,2}[-/]\d{1,2}[-/]\d{2,4}$", line) for line in lines) and any("WO" in line and re.search(r"—|\-", line) for line in lines):
        print("[DEBUG] Triggering parse_time_prefixed_format_with_date_header()")
        return parse_time_prefixed_format_with_date_header(lines)

    return jobs

def parse_time_prefixed_format_with_date_header(lines):
    jobs = []
    current_date = None

    for line in lines:
        line = line.strip()
        if not line:
            continue

        # If line looks like a standalone date (e.g. "5-2-25")
        if re.match(r"\d{1,2}[-/]\d{1,2}[-/]\d{2,4}", line):
            current_date = line.replace('/', '-')
            continue

        # Match job lines like: "10:00 - Name - Account - Type - Address - WO ###### — Tech"
        job_match = re.match(
            r"(?P<time>\d{1,2}:\d{2})\s*-\s*(?P<name>.+?)\s*-\s*(?P<account>\d{4}-\d{4}-\d{4})\s*-\s*(?P<type>.+?)\s*-\s*(?P<address>.+?)\s*-\s*WO\s*(?P<wo>\d+)\s*[-—]\s*(?P<tech>.+)",
            line
        )

        if job_match and current_date:
            groups = job_match.groupdict()
            jobs.append({
                "Date": current_date,
                "Time": groups["time"].strip(),
                "Name": groups["name"].strip(),
                "Account": groups["account"].strip(),
                "Type": groups["type"].strip(),
                "Address": groups["address"].strip(),
                "WO": groups["wo"].strip(),
                "Tech": groups["tech"].strip()
            })

    return jobs

def parse_subt_tabular_variant(lines):
    jobs = []
    for line in lines:
        print(f"[DEBUG] Raw line: {repr(line)}")
        parts = re.split(r'\s{2,}|\t+', line.strip())
        print(f"[DEBUG] Split parts: {parts}")
        if len(parts) < 7:
            continue  # Skip malformed lines

        date = parts[0]
        time = parts[1]
        name = parts[2]
        job_type = parts[3]
        wo = parts[4]
        address = parts[5]
        tech = parts[6]

        jobs.append({
            "Date": date.strip(),
            "Time": time.strip(),
            "Name": name.strip(),
            "Type": job_type.strip(),
            "WO": wo.strip(),
            "Address": address.strip(),
            "Tech": tech.strip(),
        })

    return jobs

def parse_grouped_tech_date_schedule(lines, selected_date=None):
    jobs = []

    current_tech = None
    current_date = None
    current_time = None

    time_block_pattern = re.compile(r"^\d{1,2}(:\d{2})?\s*(AM|PM)?$", re.IGNORECASE)
    lines = [line for line in lines if not time_block_pattern.match(line.strip())]  

    date_pattern = re.compile(r"\d{1,2}[-/]\d{1,2}([-/]\d{2,4})?")
    time_pattern = re.compile(r"^\d{1,2}(?::\d{2})?\s?(AM|PM)?$", re.IGNORECASE)
    job_pattern = re.compile(
        r"(?P<time>\d{1,2}:\d{2})\s*-\s*(?P<name>.+?) - (?P<account>\d{4}-\d{4}-\d{4}) - (?P<type>.+?) - (?P<address>.+?) - WO (?P<wo>\d+)",
        re.IGNORECASE
    )


    for i, line in enumerate(lines):
        line = line.strip()
        if not line:
            continue

        # Detect date (after name line)
        if date_pattern.match(line):
            for fmt in ("%m/%d/%Y", "%m/%d/%y", "%m/%d"):
                try:
                    parsed = datetime.strptime(line.replace('-', '/'), fmt)
                    if fmt == "%m/%d":  # No year — assume current
                        parsed = parsed.replace(year=datetime.today().year)
                    break
                except:
                    parsed = None
            if parsed:
                current_date = parsed.strftime("%-m/%-d/%y") if os.name != 'nt' else parsed.strftime("%#m/%#d/%y")
            continue

        # Detect new tech name (assumes followed by a date)
        if i + 1 < len(lines) and date_pattern.match(lines[i + 1].strip()):
            current_tech = line.strip()
            continue

        # Detect time header
        if time_pattern.match(line):
        # Normalize times like "8AM" → "8:00 AM"
            time_match = re.match(r"^(\d{1,2})(AM|PM)$", line.strip(), re.IGNORECASE)
            if time_match:
                hour, meridiem = time_match.groups()
                current_time = f"{hour}:00 {meridiem.upper()}"
            else:
                current_time = line.strip()

            continue

        # Match job line
        match = job_pattern.match(line)
        if match and current_tech and current_date:
            data = match.groupdict()
            jobs.append({
                "Tech": current_tech,
                "Date": current_date,
                "Time": data["time"].strip(),
                "Name": data["name"].strip(),
                "Account": data["account"].strip(),
                "Type": data["type"].strip(),
                "Address": data["address"].strip(),
                "WO": data["wo"].strip()
            })

    return jobs

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

def assign_contractor_company(driver, wo_number, contractor_name, contractor_id):
    try:
        # === Get currently assigned contractor from page ===
        WebDriverWait(driver, 5).until(
            EC.presence_of_element_located((By.CLASS_NAME, "contractorsection"))
        )
        section_elem = driver.find_element(By.CLASS_NAME, "contractorsection")
        section_text = section_elem.text.strip()

        lines = section_text.splitlines()
        assigned_contractor = None
        for line in lines:
            if " - (Primary" in line:
                assigned_contractor = line.split(" - ")[0].strip()
                break

        # fallback if 'Primary' not found
        if not assigned_contractor:
            for line in lines:
                if "assigned to this work order" not in line and "Contractors" not in line:
                    assigned_contractor = line.split(" - ")[0].strip()
                    break

        # === Compare with selected contractor ===
        if assigned_contractor and assigned_contractor.lower() == contractor_name.lower():
            log(f"✅ Contractor already assigned on WO #{wo_number}: {assigned_contractor}")
            return

        # === Attempt to remove incorrect contractor ===
        log(f"🧹 Removing incorrect contractor '{assigned_contractor}' on WO #{wo_number}")
        try:
            remove_link = section_elem.find_element(By.LINK_TEXT, "Remove")
            driver.execute_script("arguments[0].click();", remove_link)
            time.sleep(1.5)
        except Exception as e:
            log(f"❌ Could not remove existing contractor on WO #{wo_number}: {e}")

        # === Assign correct contractor ===
        contractor_select = Select(driver.find_element(By.ID, "ContractorID"))
        contractor_select.select_by_value(str(contractor_id))
        driver.execute_script("assignContractor('{}');".format(wo_number))
        time.sleep(1.5)
        log(f"✅ Assigned contractor '{contractor_name}' to WO #{wo_number}")

    except Exception as e:
        log(f"❌ Failed to assign contractor on WO #{wo_number}: {e}")

def is_headless():
    try:
        return HEADLESS_MODE.get()
    except:
        return False

def save_cookies(driver, filename="cookies.pkl"):
    with cookie_lock:
        with open(filename, "wb") as f:
            pickle.dump(driver.get_cookies(), f)

def load_cookies(driver, filename="cookies.pkl"):
    if not os.path.exists(filename): return False
    try:
        with cookie_lock:
            with open(filename, "rb") as f:
                cookies = pickle.load(f)

        driver.get("http://inside.sockettelecom.com/")
        for cookie in cookies:
            driver.add_cookie(cookie)
        driver.refresh()
        clear_first_time_overlays(driver)
        return True
    except Exception:
        if os.path.exists(filename): os.remove(filename)
        return False

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

def create_gui():
    app = TkinterDnD.Tk()
    app.title("Drop Excel File or Paste Schedule Text")
    app.geometry("600x400")

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
        Thread(target=threaded_process, daemon=True).start()

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
            
            log(f"📦 Runtime headless check (before assign): {HEADLESS_MODE.get()}")
            def threaded_assign():
                assign_jobs_from_dataframe(filtered_df)
                gui_log("✅ All jobs have been assigned.")

            Thread(target=threaded_assign, daemon=True).start()

            # === Save log
            now = datetime.now()
            filename = f"Output{now.strftime('%m%d%H%M')}.txt"
            log_path = os.path.join(LOG_FOLDER, filename)
            with open(log_path, "w", encoding="utf-8") as f:
                f.write("\n".join(log_lines))

            gui_log(f"✅ Done processing pasted text.")

            first_jobs = build_first_jobs_summary(filtered_df, name_column="Tech")

            if WANT_FIRST_JOBS.get():
                app.after(0, lambda: show_first_jobs(first_jobs))

        except Exception as e:
            gui_log(f"❌ Error processing pasted text: {e}")

    btn = tk.Button(app, text="Parse & Assign from Pasted Text", command=parse_text)
    btn.pack(pady=5)

    label.drop_target_register(DND_FILES)
    label.dnd_bind('<<Drop>>', drop)

    app.mainloop()

if __name__ == "__main__":
    try:
        create_gui()
    except Exception as e:
        print(f"\n❌ Fatal error: {e}")
        input("\nPress Enter to close...")
