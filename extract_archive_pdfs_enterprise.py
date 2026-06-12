
# extract_archive_pdfs_enterprise.py
# Fully optimized Outlook Desktop PDF extractor for large Online Archives

import os, re, csv, time
from datetime import datetime
import pythoncom
import win32com.client

YEAR = int(input("Enter Year (Example: 2025): "))
ROOT_OUTPUT = fr"D:\Archive_PDFs\{YEAR}"

LOG_FLUSH_SIZE = 5000
EMAIL_FLUSH_SIZE = 5000

os.makedirs(ROOT_OUTPUT, exist_ok=True)

LOG_FILE = os.path.join(ROOT_OUTPUT, "extraction_log.csv")
PROCESSED_EMAILS_FILE = os.path.join(ROOT_OUTPUT, "processed_emails.csv")
ERROR_LOG = os.path.join(ROOT_OUTPUT, "errors.log")
CHECKPOINT_FILE = os.path.join(ROOT_OUTPUT, "checkpoint.txt")

def clean_filename(name):
    return re.sub(r'[<>:"/\\|?*]', "_", str(name))

processed_files = set()
processed_emails = set()

if os.path.exists(LOG_FILE):
    with open(LOG_FILE, "r", encoding="utf-8") as f:
        r = csv.reader(f)
        next(r, None)
        for row in r:
            if len(row) >= 4:
                processed_files.add(row[3])

if os.path.exists(PROCESSED_EMAILS_FILE):
    with open(PROCESSED_EMAILS_FILE, "r", encoding="utf-8") as f:
        for line in f:
            if line.strip():
                processed_emails.add(line.strip())

if not os.path.exists(LOG_FILE):
    with open(LOG_FILE, "w", newline="", encoding="utf-8") as f:
        csv.writer(f).writerow(
            ["ReceivedDate","Subject","AttachmentName","SavedPath"]
        )

pythoncom.CoInitialize()
outlook = win32com.client.gencache.EnsureDispatch("Outlook.Application")
namespace = outlook.GetNamespace("MAPI")

archive_store = None
for store in namespace.Folders:
    if "Online Archive" in store.Name:
        archive_store = store
        break

if archive_store is None:
    raise Exception("Online Archive mailbox not found")

archive_inbox = archive_store.Folders["Inbox"]

folders = []

def collect_folders(folder):
    folders.append(folder)
    for sub in folder.Folders:
        collect_folders(sub)

print("Discovering folders...")
collect_folders(archive_inbox)

email_count = 0
start_date = f"01/01/{YEAR} 12:00 AM"
end_date = f"12/31/{YEAR} 11:59 PM"

print("Counting emails...")

for folder in folders:
    try:
        items = folder.Items
        filtered = items.Restrict(
            f"[ReceivedTime] >= '{start_date}' AND [ReceivedTime] <= '{end_date}'"
        )
        email_count += filtered.Count
    except:
        pass

remaining = email_count - len(processed_emails)

print(f"Total Emails      : {email_count:,}")
print(f"Processed Earlier : {len(processed_emails):,}")
print(f"Remaining         : {remaining:,}")
print(f"Folders           : {len(folders):,}")

if input("Start extraction? (yes/no): ").lower() != "yes":
    raise SystemExit

pdf_count = 0
skipped_count = 0
current_email = len(processed_emails)

log_buffer = []
processed_email_buffer = []
created_folders = set()

start_time = time.time()

def flush_buffers():
    global log_buffer, processed_email_buffer

    if log_buffer:
        with open(LOG_FILE, "a", newline="", encoding="utf-8") as f:
            csv.writer(f).writerows(log_buffer)
        log_buffer.clear()

    if processed_email_buffer:
        with open(PROCESSED_EMAILS_FILE, "a", encoding="utf-8") as f:
            f.write("\n".join(processed_email_buffer) + "\n")
        processed_email_buffer.clear()

for folder_no, folder in enumerate(folders, start=1):

    try:
        items = folder.Items

        filtered = items.Restrict(
            f"[ReceivedTime] >= '{start_date}' AND [ReceivedTime] <= '{end_date}'"
        )

        try:
            filtered.Sort("[ReceivedTime]", True)
        except:
            pass

        for mail in filtered:

            try:
                entry_id = mail.EntryID
            except:
                continue

            current_email += 1

            if entry_id in processed_emails:
                continue

            if current_email % 50 == 0:

                elapsed = max(time.time() - start_time, 1)

                pct = (current_email / email_count) * 100
                rate = current_email / elapsed

                rem = email_count - current_email
                eta = int(rem / rate) if rate else 0

                print(
                    f"\rFolder {folder_no}/{len(folders)} | "
                    f"{current_email:,}/{email_count:,} | "
                    f"{pct:.2f}% | "
                    f"{rate:.1f} emails/sec | "
                    f"PDFs {pdf_count:,} | "
                    f"Skipped {skipped_count:,} | "
                    f"ETA {eta//3600}h {(eta%3600)//60}m",
                    end="",
                    flush=True
                )

            try:

                if mail.Attachments.Count == 0:
                    processed_email_buffer.append(entry_id)
                    processed_emails.add(entry_id)
                    continue

                received = mail.ReceivedTime

                month_dir = os.path.join(
                    ROOT_OUTPUT,
                    received.strftime("%b")
                )

                if month_dir not in created_folders:
                    os.makedirs(month_dir, exist_ok=True)
                    created_folders.add(month_dir)

                subject = getattr(mail, "Subject", "")

                for i in range(1, mail.Attachments.Count + 1):

                    att = mail.Attachments.Item(i)

                    filename = att.FileName

                    if ".pdf" not in filename.lower():
                        continue

                    filepath = os.path.join(
                        month_dir,
                        clean_filename(filename)
                    )

                    if filepath in processed_files:
                        skipped_count += 1
                        continue

                    if os.path.exists(filepath):
                        skipped_count += 1
                        processed_files.add(filepath)
                        continue

                    att.SaveAsFile(filepath)

                    pdf_count += 1

                    processed_files.add(filepath)

                    log_buffer.append([
                        received.strftime("%Y-%m-%d %H:%M:%S"),
                        subject,
                        filename,
                        filepath
                    ])

                processed_email_buffer.append(entry_id)
                processed_emails.add(entry_id)

                if len(log_buffer) >= LOG_FLUSH_SIZE or \
                   len(processed_email_buffer) >= EMAIL_FLUSH_SIZE:

                    flush_buffers()

                    with open(CHECKPOINT_FILE, "w", encoding="utf-8") as f:
                        f.write(str(current_email))

            except Exception as ex:

                with open(ERROR_LOG, "a", encoding="utf-8") as ef:
                    ef.write(
                        f"{datetime.now()} | MAIL | {ex}\n"
                    )

    except Exception as ex:

        with open(ERROR_LOG, "a", encoding="utf-8") as ef:
            ef.write(
                f"{datetime.now()} | FOLDER | {folder.Name} | {ex}\n"
            )

flush_buffers()

duration = datetime.now() - datetime.fromtimestamp(start_time)

print("\n\n================================")
print("EXTRACTION COMPLETED")
print("================================")
print(f"Emails Total      : {email_count:,}")
print(f"Emails Processed  : {current_email:,}")
print(f"PDFs Extracted    : {pdf_count:,}")
print(f"PDFs Skipped      : {skipped_count:,}")
print(f"Duration          : {duration}")
print(f"Output            : {ROOT_OUTPUT}")
