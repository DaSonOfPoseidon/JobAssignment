import sys
import os
import re
import time
import tempfile
from datetime import datetime, timedelta
from collections import defaultdict

import pandas as pd
import tkinter as tk
from tkinter import ttk
from tkinterdnd2 import DND_FILES, TkinterDnD
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select, WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service

# === CONFIGURATION ===
SHOW_ALL_OUTPUT_IN_CONSOLE = False
CHROMEDRIVER_PATH = os.path.join(os.path.dirname(__file__), "chromedriver.exe")
LOG_FOLDER = os.path.join(os.path.dirname(__file__), "logs")
os.makedirs(LOG_FOLDER, exist_ok=True)

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
    "advanced": "Advanced Electric",
}

NAME_CORRECTIONS = {
    "jeff t": "Jeffery Thornton",
    "cliff": "Clifford Kunkle",
    "christopher k": "Chris Kunkle",
    "simmie": "Simmie Dunn",
    "will": "William Woods",
    "nick": "Nick Prichett",
    "kyle": "Kyle Thatcher",
    "blake": "Blake Wellman",
    "jacob": "Jacob Jones",
    "adam": "Adam Ward"
}

log_lines = []

def log(message):
    global log_lines
    log_lines.append(message)
    if SHOW_ALL_OUTPUT_IN_CONSOLE and (message.startswith("🟡") or message.startswith("🟢")):
        print(message)
    elif message.startswith("❌") or message.startswith("✅"):
        print(message)

def flexible_date_parser(date_str):
    try:
        return pd.to_datetime(str(date_str), errors='coerce')
    except:
        return None

def format_time_str(t):
    try:
        return datetime.strptime(t.strip(), "%I:%M %p").strftime("%-I%p").lower()
    except:
        return t.strip()

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

    df = df.dropna(subset=['Date', 'WO'])

    unique_dates = sorted(df['Date'].dropna().dt.date.unique())
    print("\nAvailable Dates:")
    for d in unique_dates:
        print(f" - {d}")

    start_input = input("\nEnter start date (YYYY-MM-DD): ")
    end_input = input("Enter end date (YYYY-MM-DD): ")
    try:
        start_date = datetime.strptime(start_input, "%Y-%m-%d").date()
        end_date = datetime.strptime(end_input, "%Y-%m-%d").date()
    except ValueError:
        print("Invalid date format. Use YYYY-MM-DD.")
        input("\nPress Enter to close...")
        return

    filtered_df = df[(df['Date'].dt.date >= start_date) & (df['Date'].dt.date <= end_date)]

    if filtered_df.empty:
        print("No matching jobs found for that date range.")
        input("\nPress Enter to close...")
        return

    # Save to temporary Excel file and reuse the existing function
    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
        temp_path = tmp.name
        filtered_df.to_excel(temp_path, index=False)
        print(f"\n📄 Temporary Excel created: {temp_path}")

    process_workorders(temp_path)

def process_workorders(file_path):
    print(f"\nProcessing file: {file_path}")
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

    unique_dates = sorted(df['Date'].dt.date.unique())
    print("\nAvailable Dates:")
    for d in unique_dates:
        print(f" - {d}")

    start_input = input("\nEnter start date (YYYY-MM-DD): ")
    end_input = input("Enter end date (YYYY-MM-DD): ")
    try:
        start_date = datetime.strptime(start_input, "%Y-%m-%d").date()
        end_date = datetime.strptime(end_input, "%Y-%m-%d").date()
    except ValueError:
        print("Invalid date format. Use YYYY-MM-DD.")
        input("\nPress Enter to close...")
        return

    filtered_df = df[(df['Date'].dt.date >= start_date) & (df['Date'].dt.date <= end_date)]

    if filtered_df.empty:
        print("No matching jobs found for that date range.")
        input("\nPress Enter to close...")
        return

    print(f"\nProcessing {len(filtered_df)} work orders...")
    options = webdriver.ChromeOptions()
    options.add_argument("--start-maximized")
    driver = webdriver.Chrome(service=Service(CHROMEDRIVER_PATH), options=options)

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
            print(f"❌ Invalid WO number '{raw_wo}' on spreadsheet line {excel_row}.")
            continue

        url = BASE_URL + wo_number
        log(f"\n🔗 Opening WO #{wo_number} — {url}")
        driver.get(url)

        max_attempts = 3
        matched_wo = False

        for attempt in range(1, max_attempts + 1):
            try:
                if not driver.current_url.strip().endswith(wo_number):
                    driver.get(url)

                displayed_wo_elem = WebDriverWait(driver, 10, poll_frequency=0.5).until(
                    EC.presence_of_element_located((By.XPATH, "//td[contains(text(), 'Work Order #:')]/following-sibling::td"))
                )
                displayed_wo = displayed_wo_elem.text.strip()

                if displayed_wo == wo_number:
                    matched_wo = True
                    break
                else:
                    log(f"🟡 Attempt {attempt}: Page shows WO #{displayed_wo}, expected #{wo_number}. Retrying...")
            except Exception:
                log(f"🟡 Attempt {attempt}: Unable to find WO number on page. Retrying...")

            time.sleep(5)

        if not matched_wo:
            print(f"❌ Failed to verify correct WO after 3 attempts — skipping WO #{wo_number}")
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

            parts = dropdown_value.strip().lower().split()
            if not parts:
                print(f"❌ No valid name format for WO #{wo_number} — skipping.")
                continue

            first_name = parts[0]
            last_initial = parts[1][0] if len(parts) > 1 else ""
            matched_option = None

            if dropdown_value.strip().lower() in NAME_CORRECTIONS:
                corrected_name = NAME_CORRECTIONS[dropdown_value.strip().lower()]
                for option in select.options:
                    if option.text.strip().lower() == corrected_name.lower():
                        matched_option = option.text
                        break

            if not matched_option:
                for option in select.options:
                    full_text = option.text.lower().strip()
                    full_parts = full_text.split()
                    if len(full_parts) >= 2:
                        full_first = full_parts[0]
                        full_last_initial = full_parts[1][0]
                        if full_first.startswith(first_name) and full_last_initial == last_initial:
                            matched_option = option.text
                            break

            if not matched_option:
                potential_matches = [opt.text for opt in select.options if opt.text.lower().startswith(first_name.lower())]
                if len(potential_matches) == 1:
                    matched_option = potential_matches[0]
                elif len(potential_matches) > 1:
                    print(f"❌ Ambiguous first name '{first_name}' — found multiple matches: {', '.join(potential_matches)}")
                    continue

            if not matched_option:
                print(f"❌ No dropdown match for '{dropdown_value}' — skipping WO #{wo_number}")
                continue

            assignments_div = WebDriverWait(driver, 5).until(
                EC.presence_of_element_located((By.ID, "AssignmentsList"))
            )
            assigned_names = assignments_div.text.lower()

            if matched_option.lower() in assigned_names:
                log(f"🟡 WO #{wo_number}: '{matched_option}' is already assigned — skipping.")
                continue
            elif assigned_names:
                log(f"🟢 Tech assigned: '{matched_option}'")
            else:
                log(f"🟢 Tech assigned: '{matched_option}'")

            select.select_by_visible_text(matched_option)
            add_button = WebDriverWait(driver, 60).until(
                EC.element_to_be_clickable((By.XPATH, "//button[contains(@class, 'Socket')]"))
            )
            add_button.click()

        except Exception as e:
            print(f"❌ Error on WO #{wo_number}: {e}")

    now = datetime.now()
    filename = f"Output{now.strftime('%m%d%H%M')}.txt"
    log_path = os.path.join(LOG_FOLDER, filename)
    with open(log_path, "w", encoding="utf-8") as f:
        f.write("\n".join(log_lines))

    print(f"\n✅ Done processing work orders.")
    print(f"🗂️ Output saved to: {log_path}")

    # === FIRST JOB SUMMARY
    first_jobs = defaultdict(list)
    if not filtered_df.empty:
        grouped = filtered_df.copy()

        def parse_flexible_time(t):
            t = t.strip().lower().replace('.', '')
            formats = ("%I:%M %p", "%I %p", "%I%p", "%H:%M", "%H:%M:%S")
            for fmt in formats:
                try:
                    return datetime.strptime(t, fmt)
                except:
                    continue
            return pd.NaT

        grouped['TimeParsed'] = grouped['Time'].apply(lambda x: parse_flexible_time(str(x)))
        failed_times = grouped[grouped['TimeParsed'].isna()]
        if not failed_times.empty:
            print("\n⚠️ Could not parse the following time values:")
            print(failed_times[['Time']])

        grouped = grouped.dropna(subset=['TimeParsed', 'Dropdown', 'Name', 'Type', 'Address', 'WO'])

        for date, group in grouped.groupby(grouped['Date'].dt.date):
            group = group.sort_values('TimeParsed')
            seen = set()
            for _, row in group.iterrows():
                tech = row['Dropdown']
                if tech not in seen:
                    seen.add(tech)
                    formatted_time = row['TimeParsed'].strftime("%I%p").lstrip("0").lower()
                    candidate_line = f"{tech} - {formatted_time} - {row['Name']} - {row['Type']} - {row['Address']} - WO {row['WO']}"
                    first_jobs[date].append(candidate_line)

    want_first = input("\nOutput First Jobs? (y/n): ").strip().lower()
    if want_first == 'y':
        show_first_jobs(first_jobs)

def assign_jobs_from_dataframe(df):
    print(f"\nProcessing {len(df)} work orders from pasted text...")

    options = webdriver.ChromeOptions()
    options.add_argument("--start-maximized")
    driver = webdriver.Chrome(service=Service(CHROMEDRIVER_PATH), options=options)

    for index, row in df.iterrows():
        raw_wo = row['WO']
        raw_name = str(row.get('Dropdown', row.get('Tech', ''))).strip()
        name_parts = raw_name.split()
        if len(name_parts) >= 2:
            dropdown_value = f"{name_parts[0].capitalize()} {name_parts[1][0].upper()}"
        else:
            dropdown_value = raw_name.capitalize()

        try:
            wo_number = str(int(raw_wo))
        except (ValueError, TypeError):
            print(f"❌ Invalid WO number '{raw_wo}' on line {index + 2}.")
            continue

        url = BASE_URL + wo_number
        log(f"\n🔗 Opening WO #{wo_number} — {url}")
        driver.get(url)

        matched_wo = False
        for attempt in range(1, 4):
            try:
                if not driver.current_url.strip().endswith(wo_number):
                    driver.get(url)

                displayed_wo_elem = WebDriverWait(driver, 10, poll_frequency=0.5).until(
                    EC.presence_of_element_located((By.XPATH, "//td[contains(text(), 'Work Order #:')]/following-sibling::td"))
                )
                displayed_wo = displayed_wo_elem.text.strip()
                if displayed_wo == wo_number:
                    matched_wo = True
                    break
            except Exception:
                log(f"🟡 Attempt {attempt}: Unable to find WO number. Retrying...")
                time.sleep(5)

        if not matched_wo:
            print(f"❌ Failed to verify WO #{wo_number}. Skipping.")
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

            parts = dropdown_value.lower().split()
            first_name = parts[0] if parts else ""
            last_initial = parts[1][0] if len(parts) > 1 else ""

            if dropdown_value.lower() in NAME_CORRECTIONS:
                corrected_name = NAME_CORRECTIONS[dropdown_value.lower()]
                for option in select.options:
                    if option.text.lower() == corrected_name.lower():
                        matched_option = option.text
                        break

            if not matched_option:
                for option in select.options:
                    full_parts = option.text.lower().strip().split()
                    if len(full_parts) >= 2 and full_parts[0].startswith(first_name) and full_parts[1][0] == last_initial:
                        matched_option = option.text
                        break

            if not matched_option:
                matches = [opt.text for opt in select.options if opt.text.lower().startswith(first_name)]
                if len(matches) == 1:
                    matched_option = matches[0]

            if not matched_option:
                print(f"❌ No dropdown match for '{dropdown_value}' — skipping WO #{wo_number}")
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
            log(f"🟢 Assigned tech: '{matched_option}' to WO #{wo_number}")

        except Exception as e:
            print(f"❌ Error assigning WO #{wo_number}: {e}")

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
            print(f"❌ Could not trigger assignContractor JS on WO #{wo_number}: {e}")
            return

        # ✅ Get current contractor assignment from the page
        current_contractor = get_contractor_assignments(driver)
        if current_contractor == desired_contractor_full:
            print(f"✅ Contractor '{current_contractor}' already assigned to WO #{wo_number}")
            return  # No change needed

        print(f"🧹 Reassigning from '{current_contractor}' → '{desired_contractor_full}'")

        # 🧽 Remove the currently assigned contractor (if any)
        try:
            remove_links = driver.find_elements(By.XPATH, "//a[contains(@onclick, 'removeContractor')]")
            for link in remove_links:
                driver.execute_script("arguments[0].scrollIntoView(true);", link)
                driver.execute_script("arguments[0].click();", link)
                time.sleep(1)
        except Exception as e:
            print(f"❌ Could not remove existing contractor on WO #{wo_number}: {e}")

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

            print(f"🏷️ Assigned contractor '{desired_contractor_full}' to WO #{wo_number}")
        except Exception as e:
            print(f"❌ Failed to assign contractor on WO #{wo_number}: {e}")

    except Exception as e:
        print(f"❌ Contractor assignment process failed for WO #{wo_number}: {e}")

def reformat_contractor_text(text):
    lines = [line.strip() for line in text.strip().splitlines() if line.strip()]
    jobs = []

    # Try format detection
    is_tabular = all(any(h in line.lower() for h in ["customer", "street", "wo", "date"]) for line in lines[:2])
    is_vertical = any(re.match(r"\d{4}-\d{4}-\d{4}", line) for line in lines[:10]) and any("wo" in line.lower() for line in lines)
    is_grouped = any(re.match(r"\d{1,2}/\d{1,2}/\d{4}", line) for line in lines) and any(re.match(r"\d{1,2} (AM|PM)", line) for line in lines)

    if is_tabular:
        headers = [h.strip().lower() for h in lines[0].split("\t")]
        for row in lines[1:]:
            values = row.split("\t")
            row_dict = dict(zip(headers, values))
            jobs.append({
                "Tech": row_dict.get("tech", "").strip(),
                "Date": row_dict.get("date", "").strip(),
                "Time": row_dict.get("time frame", "").strip(),
                "Name": row_dict.get("customer name", "").strip(),
                "Account": row_dict.get("account number", "").strip(),
                "Type": row_dict.get("job type", "").strip(),
                "Address": f"{row_dict.get('street', '').strip()} {row_dict.get('city', '').strip()} {row_dict.get('zip', '').strip()}",
                "WO": row_dict.get("work order number", "").strip()
            })

    elif is_vertical:
        current_tech = None
        current_date = None
        current_time = None
        buffer = []

        for line in lines:
            if re.match(r"[A-Z\s]+", line) and "AM" not in line and "PM" not in line:
                current_tech = line.strip()
            elif re.match(r"\d{1,2}/\d{1,2}/\d{4}", line):
                current_date = datetime.strptime(line, "%m/%d/%Y").strftime("%Y-%m-%d")
            elif re.match(r"\d{1,2} (AM|PM)", line):
                current_time = line
            else:
                buffer.append(line)

            if len(buffer) == 7:  # Full job block
                name = buffer[2].strip()
                acc = buffer[3].strip()
                job_type = buffer[1].strip()
                address = f"{buffer[4].strip()}, {buffer[5].strip()} {buffer[6].strip()}"
                wo_match = re.search(r"WO\s*(\d+)", buffer[6])
                wo = wo_match.group(1) if wo_match else ""
                jobs.append({
                    "Tech": current_tech,
                    "Date": current_date,
                    "Time": current_time,
                    "Name": name,
                    "Account": acc,
                    "Type": job_type,
                    "Address": address,
                    "WO": wo
                })
                buffer = []

    elif is_grouped:
        current_tech = None
        current_date = None
        current_time = None
        tech_name_buffer = []
        last_line_was_date = False  # track date transition

        job_pattern = re.compile(r"^(.*?) - (\d{4}-\d{4}-\d{4}) _ (.*?) _ (.*?) _ WO (\d+)", re.IGNORECASE)
        alt_job_pattern = re.compile(r"^(.*?) - WO (\d+)", re.IGNORECASE)

        for line in lines:
            # All-caps name line (possible tech)
            if re.match(r"^[A-Z\s\-]+$", line) and not line.startswith("WO"):
                # if we just saw a date, this is the *first* tech name
                if last_line_was_date:
                    current_tech = line.strip()
                    last_line_was_date = False
                else:
                    tech_name_buffer.append(line.strip())

            elif re.match(r"\d{1,2}/\d{1,2}/\d{4}", line):
                current_date = datetime.strptime(line, "%m/%d/%Y").strftime("%Y-%m-%d")
                if tech_name_buffer:
                    current_tech = tech_name_buffer[0]  # Use first buffered name
                    tech_name_buffer = []
                last_line_was_date = True

            elif re.match(r"\d{1,2} (AM|PM)", line):
                current_time = line
                last_line_was_date = False

            elif "WO" in line:
                last_line_was_date = False  # make sure we reset this

                match = job_pattern.match(line)
                if match:
                    name, acc, job_type, address, wo = match.groups()
                    jobs.append({
                        "Tech": current_tech,
                        "Date": current_date,
                        "Time": current_time,
                        "Name": name.strip(),
                        "Account": acc.strip(),
                        "Type": job_type.strip(),
                        "Address": address.strip(),
                        "WO": wo.strip()
                    })
                else:
                    alt_match = alt_job_pattern.match(line)
                    if alt_match:
                        desc, wo = alt_match.groups()
                        jobs.append({
                            "Tech": current_tech,
                            "Date": current_date,
                            "Time": current_time,
                            "Name": desc.strip(),
                            "Account": "",
                            "Type": "Other",
                            "Address": "",
                            "WO": wo.strip()
                        })


    else:
        # Fallback legacy format with mixed separators
        for line in lines:
            # Normalize to use " - " between parts
            line = re.sub(r'\s*[_-]\s*', ' - ', line)
            parts = [p.strip() for p in line.split(" - ")]

            # Look for WO at the end and extract
            wo_match = re.search(r"WO\s*(\d+)", line, re.IGNORECASE)
            wo = wo_match.group(1) if wo_match else ""

            if len(parts) >= 4 and wo:
                name = parts[0]
                acc = parts[1]
                job_type = parts[2]
                address = parts[3]
                jobs.append({
                    "Tech": "",
                    "Date": "",
                    "Time": "",
                    "Name": name,
                    "Account": acc,
                    "Type": job_type,
                    "Address": address,
                    "WO": wo
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
                #print(f"🔍 Detected contractor: {contractor}")
                return contractor

        print("⚠️ No 'Primary' contractor detected.")
        return "Unknown"
        
    except Exception as e:
        print(f"❌ Could not find contractor name: {e}")
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
            #print(f"✅ Contractor already assigned on WO #{wo_number}: {assigned_contractor}")
            return

        # === Attempt to remove incorrect contractor ===
        print(f"🧹 Removing incorrect contractor '{assigned_contractor}' on WO #{wo_number}")
        try:
            remove_link = section_elem.find_element(By.LINK_TEXT, "Remove")
            driver.execute_script("arguments[0].click();", remove_link)
            time.sleep(1.5)
        except Exception as e:
            print(f"❌ Could not remove existing contractor on WO #{wo_number}: {e}")

        # === Assign correct contractor ===
        contractor_select = Select(driver.find_element(By.ID, "ContractorID"))
        contractor_select.select_by_value(str(contractor_id))
        driver.execute_script("assignContractor('{}');".format(wo_number))
        time.sleep(1.5)
        print(f"✅ Assigned contractor '{contractor_name}' to WO #{wo_number}")

    except Exception as e:
        print(f"❌ Failed to assign contractor on WO #{wo_number}: {e}")

def create_gui():
    app = TkinterDnD.Tk()
    app.title("Drop Excel File or Paste Schedule Text")
    app.geometry("600x400")

    label = tk.Label(app, text="Drag and drop your Excel file here", width=60, height=5, bg="lightgray")
    label.pack(padx=10, pady=10)

    textbox = tk.Text(app, height=10, wrap="word")
    textbox.pack(padx=10, pady=(0, 5), fill="both", expand=True)
    
    global SELECTED_CONTRACTOR
    SELECTED_CONTRACTOR = tk.StringVar(value="(none)")
    
    dropdown_frame = tk.Frame(app)
    dropdown_frame.pack()
    tk.Label(dropdown_frame, text="Contractor Company:").pack(side="left", padx=5)
    dropdown_menu = ttk.Combobox(dropdown_frame, textvariable=SELECTED_CONTRACTOR, state="readonly")
    dropdown_menu["values"] = list(CONTRACTOR_LABELS.keys())
    dropdown_menu.current(0)
    dropdown_menu.pack(side="left")

    def drop(event):
        file_path = event.data.strip('{}')
        try:
            df_test = pd.read_excel(file_path)
            process_workorders(file_path)
        except Exception as e:
            print(f"❌ Could not process file: {e}")

    def parse_text():
        raw_text = textbox.get("1.0", tk.END).strip()
        if not raw_text:
            print("No text to process.")
            return
        try:
            temp = reformat_contractor_text(raw_text)

            df = pd.DataFrame(temp)
            df['Date'] = df['Date'].apply(flexible_date_parser)
            df['Time'] = df['Time'].astype(str)

            unique_dates = sorted(df['Date'].dropna().dt.date.unique())
            print("\nAvailable Dates:")
            for d in unique_dates:
                print(f" - {d}")

            start_input = input("\nEnter start date (YYYY-MM-DD): ")
            end_input = input("Enter end date (YYYY-MM-DD): ")
            try:
                start_date = datetime.strptime(start_input, "%Y-%m-%d").date()
                end_date = datetime.strptime(end_input, "%Y-%m-%d").date()
            except ValueError:
                print("Invalid date format. Use YYYY-MM-DD.")
                return

            filtered_df = df[(df['Date'].dt.date >= start_date) & (df['Date'].dt.date <= end_date)]

            if filtered_df.empty:
                print("No matching jobs found.")
                return

            assign_jobs_from_dataframe(filtered_df)

            # === Save log
            now = datetime.now()
            filename = f"Output{now.strftime('%m%d%H%M')}.txt"
            log_path = os.path.join(LOG_FOLDER, filename)
            with open(log_path, "w", encoding="utf-8") as f:
                f.write("\n".join(log_lines))

            print(f"\n✅ Done processing pasted text.")
            print(f"🗂️ Output saved to: {log_path}")

            # === FIRST JOB SUMMARY
            first_jobs = defaultdict(list)

            def parse_flexible_time(t):
                t = t.strip().lower().replace('.', '')
                formats = ("%I:%M %p", "%I %p", "%I%p", "%H:%M", "%H:%M:%S")
                for fmt in formats:
                    try:
                        return datetime.strptime(t, fmt)
                    except:
                        continue
                return pd.NaT

            filtered_df['TimeParsed'] = filtered_df['Time'].apply(lambda x: parse_flexible_time(str(x)))
            failed_times = filtered_df[filtered_df['TimeParsed'].isna()]
            if not failed_times.empty:
                print("\n⚠️ Could not parse the following time values:")
                print(failed_times[['Time']])

            filtered_df = filtered_df.dropna(subset=['TimeParsed', 'Name', 'Type', 'Address', 'WO'])

            for date, group in filtered_df.groupby(filtered_df['Date'].dt.date):
                group = group.sort_values('TimeParsed')
                seen = set()
                for _, row in group.iterrows():
                    tech = row.get('Dropdown') or row.get('Tech', '')
                    if tech not in seen:
                        seen.add(tech)
                        formatted_time = row['TimeParsed'].strftime("%I%p").lstrip("0").lower()
                        line = f"{tech} - {formatted_time} - {row['Name']} - {row['Type']} - {row['Address']} - WO {row['WO']}"
                        first_jobs[date].append(line)

            want_first = input("\nOutput First Jobs? (y/n): ").strip().lower()
            if want_first == 'y':
                show_first_jobs(first_jobs)

        except Exception as e:
            print(f"❌ Error processing pasted text: {e}")

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
