# ASSigner - Automated Job Assignment Tool

Automates the parsing, processing, and assignment of contractor jobs from emails and internal work orders. Supports email fetching, multi-format parsing per contractor, and web automation for assignment.

---

## Overview

This project automates the entire workflow of retrieving job assignments from email, parsing multiple contractor job formats, and assigning those jobs automatically within an internal system. It includes:

- Fetching and filtering emails to extract relevant job data by date.
- Supporting multi-contractor schedule formats for parsing.
- Automating job assignment using browser automation.
- Providing summary reports for quick review.
- Supporting concurrency for efficient processing.
- Managing browser sessions and dependencies automatically.
- Securely handling credentials via environment variables.

---

## Features

- Command-line interface for email fetching and job summary output.
- Drag-and-drop GUI supporting Excel files and raw text input.
- Date range selectors and contractor company dropdown for filtering and assignment.
- Choice of headless or headed browser modes for automation.
- Intelligent fuzzy matching to correct technician name assignments.
- Robust error handling with detailed log files.
- Saving outputs and logs to organized directories.
- Ability to update existing job assignments in the system.
- Environment setup and credential management via a `.env` file.

---

## Typical Workflow

1. **Setup**  
   - Configure `.env` with IMAP and internal system credentials.

2. **Run CLI or GUI**  
   - CLI: `python CLIrunner.py --date YYYY-MM-DD` to fetch and summarize emails.  
   - GUI: `python ASSigner.py` to interactively parse and assign jobs.

3. **Process Jobs**  
   - GUI accepts pasted text or Excel files, parses jobs per contractor format.  
   - Assignments are performed automatically in the internal job system via Playwright.

4. **Review Outputs**  
   - Summary reports in HTML are generated and opened automatically.  
   - Logs available in `logs/` folder for diagnostics.

---

## Configuration & Environment

- **Credentials** stored in `.env` (located in `Misc/.env`):  
  - `EMAIL_USER`, `EMAIL_PASS`, `IMAP_SERVER`, `IMAP_PORT` for email. (Optional)  
  - `UNITY_USER`, `PASSWORD` for internal system login.  
- **Contractor mappings and name corrections** editable in code (`ASSigner.py`).  
- **Logging and outputs** saved under `logs/` and `Outputs/`.  
- Playwright browsers auto-installed in `browsers/` folder on first run.

---

## Requirements

- Python 3.8+  
- Dependencies listed in `requirements.txt`, including:  
  - `playwright`  
  - `python-dotenv`  
  - `pandas`, `openpyxl`  
  - `tkcalendar`, `tkinterdnd2`  
  - `RapidFuzz`  
  - `python-dateutil`  
  - `tqdm`

---

## Limitations & Notes

- Playwright downloads ~100 MB of Chromium on first run if missing.  
- Login credentials stored in plain `.env` file; secure accordingly.  
- Some parsing rules are tailored to specific contractor formats and may need updates for changes.  
- Update mode removes existing assignments before adding new ones.  
- GUI log output only visible in headless mode.  
- Tested primarily on internal network and specific internal URLs.

---

## Development & Contribution

- Modular Python code with threading and Playwright sync API.  
- Contributions and bug reports welcome via repository issues.  
- Use `--version` flag to check the version from CLI.  
- Run GUI by executing `ASSigner.py`.

---

## Contact

For questions or support, please submit an Issue on GitHub.
