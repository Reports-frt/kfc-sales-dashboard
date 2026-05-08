"""
KFC Food Cost — Email Puller (v2: strict mode + notifications)
==============================================================

In strict mode (default): only accepts TODAY's emails. Sends notification
to user via Outlook if any required email is missing, then exits with
error code so the calling .bat can stop the pipeline.

Usage:
    python pull_food_emails.py             # strict mode (default)
    python pull_food_emails.py --lenient   # accept up to MAX_AGE_DAYS_LENIENT old
    python pull_food_emails.py --no-notify # skip notification email
"""

import argparse
import logging
import sys
from datetime import datetime, timedelta
from pathlib import Path

# =====================================================================
# CONFIG
# =====================================================================
REPO_ROOT = Path(r"C:\Users\IT\Documents\GitHub\kfc-sales-dashboard")
WORK_DIR = REPO_ROOT / "_work"

NOTIFY_TO = "report@frt-ike.gr"
EXPECTED_SENDER = "reports@foodplus.gr"

EMAILS_TO_PULL = [
    {
        "name": "Μεικτό κέρδος",
        "subject_contains": "μεικτό κέρδος",
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

MAX_AGE_DAYS_LENIENT = 3
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
# OUTLOOK
# =====================================================================
def connect_to_outlook():
    """Open Outlook MAPI connection. Returns (Application, Inbox)."""
    try:
        import win32com.client
    except ImportError:
        log.error("pywin32 not installed. Run: pip install pywin32")
        sys.exit(1)
    
    outlook_app = win32com.client.Dispatch("Outlook.Application")
    namespace = outlook_app.GetNamespace("MAPI")
    inbox = namespace.GetDefaultFolder(6)  # 6 = Inbox
    return outlook_app, inbox


def find_and_save_email(inbox, config, work_dir, cutoff):
    """Find latest email matching pattern, save xlsx attachment. Returns result dict."""
    pattern = config["subject_contains"].lower()
    save_as = config["save_as"]
    
    log.info(f"  Searching for: subject contains '{config['subject_contains']}'")
    log.info(f"  Cutoff: emails after {cutoff.strftime('%Y-%m-%d %H:%M')}")
    
    items = inbox.Items
    items.Sort("[ReceivedTime]", True)
    
    found_email = None
    found_received = None
    checked = 0
    
    for item in items:
        checked += 1
        if checked > MAX_ITEMS_TO_SCAN:
            break
        try:
            received = item.ReceivedTime
            received_naive = received.replace(tzinfo=None) if hasattr(received, 'tzinfo') and received.tzinfo else received
            
            if received_naive < cutoff:
                break
            
            subject = (item.Subject or "").lower()
            sender = (item.SenderEmailAddress or "").lower()
            
            sender_ok = (EXPECTED_SENDER in sender) if EXPECTED_SENDER else True
            subject_ok = pattern in subject
            
            if sender_ok and subject_ok and item.Attachments.Count > 0:
                found_email = item
                found_received = received_naive
                log.info(f"  ✓ Found: '{item.Subject}' from {sender} ({received_naive})")
                break
        except Exception as e:
            log.warning(f"  Error reading item: {e}")
            continue
    
    if not found_email:
        return {
            'status': 'missing',
            'name': config['name'],
            'message': f"Δεν βρέθηκε email μετά τις {cutoff.strftime('%Y-%m-%d %H:%M')}",
        }
    
    work_dir.mkdir(parents=True, exist_ok=True)
    target_path = work_dir / save_as
    
    for att in found_email.Attachments:
        att_name = att.FileName.lower()
        if att_name.endswith(('.xlsx', '.xlsm', '.xls')):
            att.SaveAsFile(str(target_path))
            size_kb = target_path.stat().st_size / 1024
            log.info(f"  ✓ Saved: {att.FileName} → {save_as} ({size_kb:.0f} KB)")
            return {
                'status': 'ok',
                'name': config['name'],
                'received': found_received,
                'size_kb': size_kb,
                'path': target_path,
            }
    
    return {
        'status': 'no_attachment',
        'name': config['name'],
        'message': "Email βρέθηκε αλλά δεν είχε xlsx attachment",
    }


def send_notification_email(outlook_app, missing, saved, mode):
    """Send notification email when files are missing."""
    try:
        mail = outlook_app.CreateItem(0)
        mail.To = NOTIFY_TO
        
        subj_prefix = "[STRICT MODE FAIL]" if mode == 'strict' else "[LENIENT WARN]"
        mail.Subject = f"{subj_prefix} KFC Food Pipeline — Missing emails {datetime.now().strftime('%Y-%m-%d %H:%M')}"
        
        body_lines = [
            "KFC Food Pipeline notification",
            "=" * 50,
            "",
            f"Time:    {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}",
            f"Mode:    {mode}",
            f"Status:  {len(missing)}/{len(saved) + len(missing)} emails missing",
            "",
            "Missing emails:",
        ]
        for m in missing:
            body_lines.append(f"  X  {m['name']}: {m['message']}")
        
        if saved:
            body_lines.extend(["", "Successfully pulled:"])
            for s in saved:
                body_lines.append(f"  OK  {s['name']} ({s['size_kb']:.0f} KB, received {s['received']})")
        
        body_lines.extend([
            "",
            "Action items:",
            "  1. Check Outlook Inbox + Junk for emails from Reports@foodplus.gr",
            "  2. If emails exist but were not pulled, run manually:",
            f"     cd {REPO_ROOT}\\automation\\foodcost",
            "     .\\run_daily_food.bat",
            "  3. If emails never arrived, check Targit / foodplus.gr",
            "",
            "Pipeline behavior:",
            "  - STRICT mode: build/push SKIPPED. Dashboard data is stale.",
            "  - LENIENT mode: build runs with old _work/ files. Dashboard may show stale numbers.",
            "",
            "-- KFC Food Pipeline (auto-notification)",
        ])
        
        mail.Body = "\n".join(body_lines)
        mail.Send()
        log.info(f"  ✓ Notification email sent to {NOTIFY_TO}")
        return True
    except Exception as e:
        log.error(f"  ✗ Failed to send notification email: {e}")
        return False


# =====================================================================
# MAIN
# =====================================================================
def main():
    parser = argparse.ArgumentParser(description='Pull food cost emails from Outlook')
    parser.add_argument('--work-dir', default=str(WORK_DIR))
    parser.add_argument('--lenient', action='store_true',
                        help=f'Accept emails up to {MAX_AGE_DAYS_LENIENT} days old (default: only today)')
    parser.add_argument('--no-notify', action='store_true',
                        help='Skip sending notification email on failure')
    args = parser.parse_args()
    
    work_dir = Path(args.work_dir)
    mode = 'lenient' if args.lenient else 'strict'
    
    if args.lenient:
        cutoff = datetime.now() - timedelta(days=MAX_AGE_DAYS_LENIENT)
    else:
        cutoff = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
    
    print("=" * 60)
    print(f"KFC Food Cost — Email Puller ({mode} mode)")
    print(f"Started: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print("=" * 60)
    
    log.info("STEP 1: Connecting to Outlook...")
    outlook_app, inbox = connect_to_outlook()
    log.info(f"  ✓ Connected to inbox")
    
    log.info(f"\nSTEP 2: Pulling {len(EMAILS_TO_PULL)} emails...")
    saved = []
    missing = []
    
    for idx, config in enumerate(EMAILS_TO_PULL, 1):
        log.info(f"\n[{idx}/{len(EMAILS_TO_PULL)}] {config['name']}:")
        result = find_and_save_email(inbox, config, work_dir, cutoff)
        
        if result['status'] == 'ok':
            saved.append(result)
        else:
            log.warning(f"  ✗ {result['message']}")
            if config.get("required"):
                missing.append(result)
    
    print()
    print("=" * 60)
    
    if not missing:
        log.info(f"✓ All {len(saved)} emails saved successfully")
        print(f"Finished: {datetime.now().strftime('%H:%M:%S')}")
        print("=" * 60)
        return 0
    
    log.error(f"✗ Missing {len(missing)} required emails: {', '.join(m['name'] for m in missing)}")
    
    if not args.no_notify:
        log.info("Sending notification email...")
        send_notification_email(outlook_app, missing, saved, mode)
    
    if mode == 'strict':
        log.error("STRICT MODE: Exiting with error code 1 (build/push will be skipped)")
        print("=" * 60)
        return 1
    else:
        log.warning("LENIENT MODE: Continuing — build will use existing files")
        print("=" * 60)
        return 0


if __name__ == '__main__':
    sys.exit(main())
