import sys
import os
import re
import time
import tempfile
from datetime import datetime
from collections import defaultdict

import pandas as pd
import tkinter as tk
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

    first_jobs = defaultdict(list)
    if not filtered_df.empty:
        grouped = filtered_df.copy()

        # More forgiving time parsing
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

        # Debug: show times that failed to parse
        failed_times = grouped[grouped['TimeParsed'].isna()]
        if not failed_times.empty:
            print("\n⚠️ Could not parse the following time values:")
            print(failed_times[['Time']])

        # Drop rows that don't have all required fields
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

def reformat_contractor_text(raw_text):
    lines = [line.strip() for line in raw_text.splitlines() if line.strip()]
    data = []

    def is_date(line):
        try:
            datetime.strptime(line, "%m/%d/%Y")
            return True
        except:
            return False

    def is_short_date(line):
        try:
            datetime.strptime(line, "%d-%b")
            return True
        except:
            return False

    def detect_format(lines):
        if any(re.match(r"^\d{1,2}-[A-Za-z]{3}$", line) for line in lines):
            return "tabular"
        wo_lines = sum(1 for line in lines if re.match(r'^WO\s*\d+', line, re.IGNORECASE))
        if wo_lines >= len(lines) // 10 and any(re.match(r'^\d{1,2}:\d{2}\s?(AM|PM)$', line, re.IGNORECASE) for line in lines[:10]):
            return "tgs"
        if any(is_date(line) for line in lines[:10]) and any(re.match(r'\d{1,2}(am|pm)', line.lower()) for line in lines):
            return "grouped"
        return "legacy"

    fmt = detect_format(lines)

    if fmt == "tabular":
        print("📄 Detected TGS tabular format.")
        i = 0
        while i + 10 < len(lines):
            date_str = lines[i]
            try:
                job_date = datetime.strptime(date_str, "%d-%b").replace(year=datetime.today().year).date()
            except:
                i += 1
                continue

            time_raw = lines[i + 1]
            type_ = lines[i + 2]
            name = lines[i + 3]
            # skip account at i + 4
            street = lines[i + 5]
            city = lines[i + 6]
            zip_ = lines[i + 7]
            # skip state at i + 8
            wo_match = re.search(r'\d+', lines[i + 9])
            tech = lines[i + 10].title()

            if not wo_match:
                i += 11
                continue

            address = f"{street}, {city} {zip_}"
            data.append({
                "Date": job_date,
                "Time": time_raw,
                "Name": name,
                "Type": type_,
                "WO": wo_match.group(),
                "Address": address,
                "Notes": "",
                "Assignment": tech
            })
            i += 11

    elif fmt == "tgs":
        print("📄 Detected TGS (vertical 9-line) format.")
        while True:
            date_input = input("Enter the date for these jobs (YYYY-MM-DD, 'today', or 'tomorrow'): ").strip().lower()
            if date_input == "today":
                job_date = datetime.today().date()
                break
            elif date_input == "tomorrow":
                job_date = datetime.today().date() + timedelta(days=1)
                break
            else:
                try:
                    job_date = datetime.strptime(date_input, "%Y-%m-%d").date()
                    break
                except ValueError:
                    print("❌ Invalid date format. Use YYYY-MM-DD, 'today', or 'tomorrow'.")

        i = 0
        while i + 8 < len(lines):
            block = lines[i:i+9]
            time_raw = block[0]
            type_ = block[1]
            name = block[2]
            address1 = block[4]
            city = block[5]
            zip_ = block[6]
            wo_match = re.match(r'^WO\s*(\d+)', block[7], re.IGNORECASE)
            tech = block[8].title()
            if tech.lower() == "cancelled" or not wo_match:
                i += 9
                continue
            wo = wo_match.group(1)
            address = f"{address1}, {city} {zip_}"
            data.append({
                "Date": job_date, "Time": time_raw, "Name": name, "Type": type_, "WO": wo,
                "Address": address, "Notes": "", "Assignment": tech
            })
            i += 9

    elif fmt == "grouped":
        print("📄 Detected grouped format (Tech → Date → Time → Jobs)")
        i = 0
        current_tech = None
        current_date = None
        current_time = None
        while i < len(lines):
            line = lines[i]
            if re.match(r"^[A-Za-z]+\s*[A-Za-z]*$", line) and not is_date(line):
                current_tech = line.title()
                i += 1
                continue
            if is_date(line):
                current_date = datetime.strptime(line, "%m/%d/%Y").date()
                i += 1
                continue
            if re.match(r"^\d{1,2}(?::\d{2})?\s?(AM|PM)$", line, re.IGNORECASE):
                current_time = line
                i += 1
                continue
            if "WO" in line.upper():
                wo_match = re.search(r'WO\s*(\d+)', line, re.IGNORECASE)
                if not (current_tech and current_date and current_time and wo_match):
                    i += 1
                    continue
                parts = re.split(r'[-_]', line)
                parts = [p.strip() for p in parts if p.strip()]
                if len(parts) < 4:
                    i += 1
                    continue
                name = parts[0]
                type_ = ""
                address = ""
                for p in parts:
                    if any(k in p.lower() for k in ["fiber", "connectorized", "install", "service"]):
                        type_ = p
                    elif "mo" in p.lower() and any(char.isdigit() for char in p):
                        address = p
                data.append({
                    "Date": current_date, "Time": current_time, "Name": name,
                    "Type": type_, "WO": wo_match.group(1), "Address": address,
                    "Notes": "", "Assignment": current_tech
                })
            i += 1

    else:
        print("📄 Detected legacy freeform contractor format.")
                # === ORIGINAL FREEFORM FORMAT ===
        contractor = None
        date = None
        time_slot = None

        contractor_pattern = re.compile(r'^[A-Z]{2,}(?:\s+[A-Z]{2,})?$')
        date_pattern = re.compile(r'\d{1,2}/\d{1,2}/\d{4}')
        time_pattern = re.compile(r'^\d{1,2}\s?(AM|PM)$', re.IGNORECASE)
        wo_pattern = re.compile(r'WO\s*#?\s*(\d+)', re.IGNORECASE)

        skip_contractor_labels = {
            "SPLICING", "SPLICE", "DROP CREW", "BURY", "ENGINEERING",
            "STL", "PRE-BURY", "UNASSIGNED", "CANCELLED"
        }

        i = 0
        while i < len(lines):
            line = lines[i]

            if any(keyword in line.upper() for keyword in ["UNASSIGNED", "RESCHEDULE", "CANCELLED", "PRE-BURY"]):
                contractor = None
                i += 1
                continue

            if contractor_pattern.match(line):
                if i + 1 < len(lines):
                    next_line = lines[i + 1]
                    if contractor_pattern.match(next_line):
                        if next_line.upper() in skip_contractor_labels:
                            contractor = None
                            i += 2
                            continue
                        contractor = line.title()
                        i += 2
                        continue
                contractor = line.title()
                i += 1
                continue

            if date_match := date_pattern.search(line):
                try:
                    date = datetime.strptime(date_match.group(), "%m/%d/%Y").date()
                except:
                    pass
                i += 1
                continue

            if time_pattern.match(line):
                time_slot = line.upper().replace(" ", "")
                if not time_slot.endswith("M"):
                    time_slot += "M"
                i += 1
                continue

            if not (contractor and date and time_slot):
                i += 1
                continue

            if "splice" in line.lower() or "splicing" in line.lower():
                i += 1
                continue

            if "WO" in line.upper():
                wo_match = wo_pattern.search(line)
                if not wo_match:
                    i += 1
                    continue
                wo_number = wo_match.group(1)
                line = line[:wo_match.end()].strip()
                parts = re.split(r'[-_]', line)
                parts = [p.strip() for p in parts if p.strip()]
                if len(parts) < 4:
                    i += 1
                    continue

                customer = parts[0]
                type_ = ""
                address = ""

                for p in parts:
                    if any(k in p.lower() for k in ["fiber", "connectorized", "install", "service"]):
                        type_ = p
                    elif "MO" in p.upper() and any(char.isdigit() for char in p) and "WO" not in p.upper():
                        address = p

                if not (customer and type_ and address and wo_number):
                    i += 1
                    continue

                data.append({
                    "Date": date,
                    "Time": time_slot,
                    "Name": customer,
                    "Type": type_,
                    "WO": wo_number,
                    "Address": address,
                    "Notes": "",
                    "Assignment": contractor
                })

            i += 1

    if not data:
        raise ValueError("No valid WO entries found in pasted text.")

    df = pd.DataFrame(data)
    df = df[["Date", "Time", "Name", "Type", "WO", "Address", "Notes", "Assignment"]]
    temp_path = os.path.join(tempfile.gettempdir(), "formatted_assignments.xlsx")
    df.to_excel(temp_path, index=False)
    return temp_path


def create_gui():
    app = TkinterDnD.Tk()
    app.title("Drop Excel File or Paste Schedule Text")
    app.geometry("600x400")

    label = tk.Label(app, text="Drag and drop your Excel file here", width=60, height=5, bg="lightgray")
    label.pack(padx=10, pady=10)

    textbox = tk.Text(app, height=10, wrap="word")
    textbox.pack(padx=10, pady=(0, 5), fill="both", expand=True)

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
            process_workorders(temp)
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
