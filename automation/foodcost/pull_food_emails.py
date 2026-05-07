"""
KFC Food Cost — Email Puller
=============================

Connects to Outlook and downloads 3 daily food cost reports:
  1. Μεικτό κέρδος (product profitability)
  2. Πωλήσεις KFC ALC (à la carte hourly)
  3. Πωλήσεις KFC Combos (combos hourly)

Saves attachments to _work/ overwriting any existing file.
Designed to run before build_pipeline.py — typical sequence:

    1. python pull_food_emails.py
    2. python build_pipeline.py

Or both via run_daily_food.bat.
"""

import argparse
import logging
import os
import sys
from datetime import datetime, timedelta
from pathlib import Path

# =====================================================================
# CONFIG — Customize these subject patterns based on actual emails
# =====================================================================
REPO_ROOT = Path(r"C:\Users\IT\Documents\GitHub\kfc-sales-dashboard")
WORK_DIR = REPO_ROOT / "_work"

# Each entry defines an email to look for and where to save its attachment.
# Subject patterns are case-INSENSITIVE substring matches.
# Verified subject lines from actual emails (sender: Reports@foodplus.gr):
#   1. "Μεικτό κέρδος_KFC_new(Git)"          → 24 KB  Μεικτό κέρδος_KFC_new(Git).xlsx
#   2. "Πωλήσεις__KFC_ALC_Stores_Hours"      → 451 KB Πωλήσεις_ KFC_ALC_Stores_Hours(Git).xlsx
#   3. "Πωλήσεις_ KFC_Combos_Stores_Hours(Git)" → 1 MB Πωλήσεις_ KFC_Combos_Stores_Hours(Git).xlsx
#
# subject_contains uses case-insensitive substring match.
# We also restrict by sender to avoid false positives from forwarded emails.

EXPECTED_SENDER = "reports@foodplus.gr"

EMAILS_TO_PULL = [
    {
        "name": "Μεικτό κέρδος",
        "subject_contains": "μεικτό κέρδος",   # case-insensitive
        "save_as": "Μεικτό_κέρδος_KFC_new_Git_.xlsx",
        "required": True,
    },
    {
        "name": "ALC Hourly",
        "subject_contains": "kfc_alc_stores_hours",
        "save_as": "Πωλήσεις__KFC_ALC_Stores_Hours.xlsx",
        "required": True,
    },
    {
        "name": "Combos Hourly",
        "subject_contains": "kfc_combos_stores_hours",
        "save_as": "Πωλήσεις__KFC_Combos_Stores_Hours.xlsx",
        "required": True,
    },
]

# How far back to search for emails (in case daily run is missed)
MAX_AGE_DAYS = 3

# Maximum emails to scan per pattern
MAX_ITEMS_TO_SCAN = 300


# =====================================================================
# LOGGING
# =====================================================================
log = logging.getLogger("food_email_puller")
logging.basicConfig(
    level=logging.INFO,
    format='[%(asctime)s] [%(levelname)s] %(message)s',
    datefmt='%H:%M:%S',
)


# =====================================================================
# OUTLOOK FETCH
# =====================================================================
def connect_to_outlook():
    """Open Outlook MAPI connection."""
    try:
        import win32com.client
    except ImportError:
        log.error("pywin32 not installed. Run: pip install pywin32")
        sys.exit(1)
    
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    inbox = outlook.GetDefaultFolder(6)  # 6 = Inbox
    return inbox


def find_and_save_email(inbox, config, work_dir):
    """
    Find latest email matching subject pattern, save its first xlsx attachment.
    Returns saved path or None if not found.
    """
    pattern = config["subject_contains"].lower()
    save_as = config["save_as"]
    cutoff = datetime.now() - timedelta(days=MAX_AGE_DAYS)
    
    log.info(f"  Searching for: subject contains '{config['subject_contains']}'")
    
    # Get items sorted newest-first
    items = inbox.Items
    items.Sort("[ReceivedTime]", True)
    
    found_email = None
    checked = 0
    for item in items:
        checked += 1
        if checked > MAX_ITEMS_TO_SCAN:
            break
        try:
            received = item.ReceivedTime
            received_naive = received.replace(tzinfo=None) if hasattr(received, 'replace') else received
            
            if received_naive < cutoff:
                break  # everything older = beyond cutoff
            
            subject = (item.Subject or "").lower()
            sender = (item.SenderEmailAddress or "").lower()
            
            # Match: sender + subject pattern + has attachment
            sender_ok = (EXPECTED_SENDER in sender) if EXPECTED_SENDER else True
            subject_ok = pattern in subject
            
            if sender_ok and subject_ok and item.Attachments.Count > 0:
                found_email = item
                log.info(f"  ✓ Found: '{item.Subject}' from {sender} ({received_naive})")
                break
        except Exception as e:
            log.warning(f"  Error reading item: {e}")
            continue
    
    if not found_email:
        return None
    
    # Save first xlsx attachment
    work_dir.mkdir(parents=True, exist_ok=True)
    target_path = work_dir / save_as
    
    for att in found_email.Attachments:
        att_name = att.FileName.lower()
        if att_name.endswith(('.xlsx', '.xlsm', '.xls')):
            att.SaveAsFile(str(target_path))
            size_kb = target_path.stat().st_size / 1024
            log.info(f"  ✓ Saved: {att.FileName} → {save_as} ({size_kb:.0f} KB)")
            return target_path
    
    log.warning(f"  Email found but no xlsx attachment")
    return None


# =====================================================================
# MAIN
# =====================================================================
def main():
    parser = argparse.ArgumentParser(description='Pull food cost emails from Outlook')
    parser.add_argument('--work-dir', default=str(WORK_DIR),
                        help='Output folder for attachments (default: _work/)')
    parser.add_argument('--strict', action='store_true',
                        help='Exit with error if any required email is missing')
    args = parser.parse_args()
    
    work_dir = Path(args.work_dir)
    
    print("=" * 60)
    print("KFC Food Cost — Email Puller")
    print(f"Started: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"Work dir: {work_dir}")
    print("=" * 60)
    
    log.info("STEP 1: Connecting to Outlook...")
    inbox = connect_to_outlook()
    log.info(f"  ✓ Connected to inbox")
    
    log.info(f"\nSTEP 2: Pulling {len(EMAILS_TO_PULL)} emails...")
    saved_count = 0
    missing = []
    
    for idx, config in enumerate(EMAILS_TO_PULL, 1):
        log.info(f"\n[{idx}/{len(EMAILS_TO_PULL)}] {config['name']}:")
        result = find_and_save_email(inbox, config, work_dir)
        if result:
            saved_count += 1
        else:
            log.warning(f"  ✗ Not found in last {MAX_AGE_DAYS} days")
            if config.get("required"):
                missing.append(config["name"])
    
    print()
    print("=" * 60)
    if missing and args.strict:
        log.error(f"Missing {len(missing)} required emails: {', '.join(missing)}")
        log.error("Exiting with error (strict mode)")
        sys.exit(1)
    elif missing:
        log.warning(f"Saved {saved_count}/{len(EMAILS_TO_PULL)} (missing: {', '.join(missing)})")
        log.warning("Build will use existing files in _work/ for missing items")
    else:
        log.info(f"✓ All {saved_count} emails saved successfully")
    print(f"Finished: {datetime.now().strftime('%H:%M:%S')}")
    print("=" * 60)
    
    return 0 if not (missing and args.strict) else 1


if __name__ == '__main__':
    sys.exit(main())
