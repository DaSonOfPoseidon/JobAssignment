#!/usr/bin/env python3
"""
CLI entrypoint for automating the email-to-job-assignment workflow.
"""
import argparse
import os
import re
from datetime import datetime
from dotenv import load_dotenv

# Reuse IMAP/email utilities
from email_reader import (
    connect_imap,
    find_matching_msg_nums,
    fetch_body,
    extract_relevant_section
)
# Reuse your parsing and summary logic
from ASSigner import Assigner

# Simple company detection via sender/cc (same regex map as before)
COMPANY_REGEX = {
    'SubT': re.compile(r'iteike@subterraneus\.net|billing@subterraneus\.net|danielle@subterraneus\.net', re.I),
    'Tex-Star': re.compile(r'shagins@tex-starcommunications\.com|douggale01@gmail\.com', re.I),
    'All-Clear': re.compile(r'carriem@allclearmo\.com|shawn@allclearmo\.com', re.I),
    'Pifer': re.compile(r'calebpifer8@gmail\.com', re.I),
    'TGS': re.compile(r'emily\.moody@takbroadband\.com|cmoore@takbroadband\.com', re.I),
    'Advanced': re.compile(r'bdmurphy@advlights\.com|kmurphy@advlights\.com', re.I),
}

def detect_company(sender: str, cc_list: list[str]) -> str:
    combined = sender + ' ' + ' '.join(cc_list)
    for name, rx in COMPANY_REGEX.items():
        if rx.search(combined):
            return name
    return 'Unknown'

def parse_date(s: str):
    from dateutil.parser import parse
    default = datetime(datetime.now().year, 1, 1)
    return parse(s, default=default).date()

def main():
    parser = argparse.ArgumentParser(
        description="Fetch emails for a single date and print each company's First Jobs"
    )
    parser.add_argument('-d', '--date', required=True,
                        help='Target date (e.g. 5/28 or 2025-05-28)')
    parser.add_argument('-t', '--threads', type=int, default=8,
                        help='Worker threads count (default: 8)')
    args = parser.parse_args()

    # Load .env for IMAP_SERVER, EMAIL_USER, EMAIL_PASS, etc.
    load_dotenv()

    # Parse the target date
    target = parse_date(args.date)

    # Connect to IMAP and find messages for that date
    imap = connect_imap()
    try:
        msg_nums = find_matching_msg_nums(imap, target)
    except Exception as e:
        print("Error searching mail:", e)
        return

    # Initialize the Assigner
    assigner = Assigner(num_threads=args.threads)

    # Process each email
    results: dict[str, list[str]] = {}
    for num in msg_nums:
        raw = fetch_body(imap, num)
        section = extract_relevant_section(raw)

        # Extract sender and CC from the 'section' text itself
        from_match = re.search(r'^From:\s*.*?<([^>]+)>', section, re.M)
        cc_match   = re.search(r'^Cc:\s*(.+)', section, re.M)
        sender     = from_match.group(1).strip() if from_match else ''
        cc_list    = [addr.strip() for addr in cc_match.group(1).split(',')] if cc_match else []

        company = detect_company(sender, cc_list)

        # Parse jobs and pull first jobs
        jobs   = assigner.parse_jobs(section)
        firsts = assigner.first_jobs(jobs)
        results.setdefault(company, []).extend(firsts)

    # Logout IMAP
    imap.logout()

    # Print out the summaries
    for company, lines in results.items():
        print(f"=== {company} First Jobs ===")
        for l in lines:
            print(l)

if __name__ == '__main__':
    main()
