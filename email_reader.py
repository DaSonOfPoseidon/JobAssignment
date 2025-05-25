# email_reader.py

import os
import re
import imaplib
import email
from email.header import decode_header
from email import policy
from datetime import date, timedelta, datetime
from dotenv import load_dotenv

load_dotenv()  # expects .env to have EMAIL_USER, EMAIL_PASS, IMAP_SERVER, IMAP_PORT

EMAIL_USER  = os.getenv("EMAIL_USER")
EMAIL_PASS  = os.getenv("EMAIL_PASS")
IMAP_SERVER = os.getenv("IMAP_SERVER", "imap.gmail.com")
IMAP_PORT   = int(os.getenv("IMAP_PORT", 993))


def connect_imap():
    imap = imaplib.IMAP4_SSL(IMAP_SERVER, IMAP_PORT)
    imap.login(EMAIL_USER, EMAIL_PASS)
    return imap


def decode_mime_words(s: str) -> str:
    parts = []
    for text, enc in decode_header(s):
        if isinstance(text, bytes):
            parts.append(text.decode(enc or "utf-8", errors="ignore"))
        else:
            parts.append(text)
    return "".join(parts)


def fetch_body(imap, msg_num: bytes) -> str:
    status, data = imap.fetch(msg_num, "(RFC822)")
    if status != "OK":
        return ""
    msg = email.message_from_bytes(data[0][1], policy=policy.default)
    if msg.is_multipart():
        for part in msg.walk():
            if part.get_content_type() == "text/plain" and not part.get_content_disposition():
                return part.get_content()
    return msg.get_content()


def extract_relevant_section(body: str) -> str:
    """
    Extract the chunk starting at the first "From:" line
    up to (but not including) the line that starts with "On "
    and mentions "Corey".
    """
    lines = body.splitlines()
    start_idx = None
    for i, line in enumerate(lines):
        if line.startswith("From:"):
            start_idx = i
            break
    if start_idx is None:
        return ""  # no From: header found

    extracted = []
    for line in lines[start_idx:]:
        if line.startswith("On ") and "wrote:" in line:
            break
        extracted.append(line)
    return "\n".join(extracted).strip()


def subject_matches_date(subject: str, target: date) -> bool:
    m = target.month
    d = target.day

    # full mm/dd-mm/dd (same-month only)
    for rng in re.finditer(r"\b(\d{1,2})/(\d{1,2})\s*-\s*(\d{1,2})/(\d{1,2})\b", subject):
        sm, sd, em, ed = map(int, rng.groups())
        if sm == em == m and sd <= d <= ed:
            return True

    # day-only range
    for rng in re.finditer(r"\b(\d{1,2})\s*-\s*(\d{1,2})\b", subject):
        sd, ed = map(int, rng.groups())
        if sd <= d <= ed:
            return True

    # standalone day
    if re.search(rf"\b{d}\b", subject):
        return True

    return False


def find_matching_msg_nums(imap, target: date) -> list[bytes]:
    imap.select("INBOX")
    status, data = imap.search(None, "ALL")
    if status != "OK":
        raise RuntimeError("IMAP search failed")

    hits = []
    for num in data[0].split():
        status, hdr = imap.fetch(num, '(BODY.PEEK[HEADER.FIELDS (SUBJECT)])')
        if status != "OK":
            continue
        hdr_bytes = hdr[0][1]
        msg = email.message_from_bytes(hdr_bytes, policy=policy.default)
        subj = decode_mime_words(msg.get("Subject", ""))
        if subject_matches_date(subj, target):
            hits.append(num)
    return hits


if __name__ == "__main__":
    # for testing override, set EMAIL_TARGET_DATE in .env as YYYY-MM-DD
    target_env = os.getenv("EMAIL_TARGET_DATE")
    if target_env:
        try:
            tomorrow = datetime.fromisoformat(target_env).date()
        except ValueError:
            print("Invalid EMAIL_TARGET_DATE; should be YYYY-MM-DD")
            exit(1)
    else:
        tomorrow = date.today() + timedelta(days=1)

    imap = None
    try:
        imap = connect_imap()
        msg_nums = find_matching_msg_nums(imap, tomorrow)
        print(f"Found {len(msg_nums)} message(s) matching {tomorrow}:\n")
        for num in msg_nums:
            raw = fetch_body(imap, num)
            relevant = extract_relevant_section(raw)
            print(f"--- Message #{num.decode()} ---\n{relevant}\n")
    except Exception as e:
        print("Error while fetching mail:", e)
    finally:
        if imap:
            try: imap.logout()
            except: pass
