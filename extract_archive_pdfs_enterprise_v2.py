
# extract_archive_pdfs_enterprise_v2.py
# Features:
# - Online Archive Inbox + all subfolders
# - Resume using EntryID
# - Progress %, emails/sec, ETA
# - Buffered logging
# - Extract only selected file types (PDF / ZIP / BOTH)
# - Month-wise storage
# - Logs extracted email subjects
# - Logs unprocessed emails with reasons

import os
import re
import csv
import time
from datetime import datetime
import pythoncom
import win32com.client

YEAR = int(input("Enter Year (Example: 2025): "))

print("\nWhat do you want to extract?")
print("1 - PDF")
print("2 - ZIP")
print("3 - PDF + ZIP")

choice = input("Enter choice (1/2/3): ").strip()

if choice == "1":
    VALID_EXTENSIONS = {".pdf"}
elif choice == "2":
    VALID_EXTENSIONS = {".zip"}
else:
    VALID_EXTENSIONS = {".pdf", ".zip"}

ROOT_OUTPUT = fr"D:\Archive_Files\{YEAR}"

LOG_FLUSH_SIZE = 5000
EMAIL_FLUSH_SIZE = 5000

os.makedirs(ROOT_OUTPUT, exist_ok=True)

EXTRACTION_LOG = os.path.join(ROOT_OUTPUT, "extraction_log.csv")
PROCESSED_EMAILS_FILE = os.path.join(ROOT_OUTPUT, "processed_emails.csv")
SUBJECT_LOG = os.path.join(ROOT_OUTPUT, "email_subjects.csv")
UNPROCESSED_LOG = os.path.join(ROOT_OUTPUT, "unprocessed_emails.csv")
ERROR_LOG = os.path.join(ROOT_OUTPUT, "errors.log")
CHECKPOINT_FILE = os.path.join(ROOT_OUTPUT, "checkpoint.txt")


def clean_filename(name):
    return re.sub(r'[<>:"/\\|?*]', "_", str(name))


def ensure_csv(path, headers):
    if not os.path.exists(path):
        with open(path, "w", newline="", encoding="utf-8") as f:
            csv.writer(f).writerow(headers)


ensure_csv(
    EXTRACTION_LOG,
    ["ReceivedDate", "Subject", "AttachmentName", "SavedPath", "EntryID"]
)

ensure_csv(
    SUBJECT_LOG,
    ["ReceivedDate", "Subject", "EntryID", "Folder"]
)

ensure_csv(
    UNPROCESSED_LOG,
    ["ReceivedDate", "Subject", "EntryID", "Reason"]
)

processed_emails = set()

if os.path.exists(PROCESSED_EMAILS_FILE):
    with open(PROCESSED_EMAILS_FILE, "r", encoding="utf-8") as f:
        for line in f:
            eid = line.strip()
            if eid:
                processed_emails.add(eid)

pythoncom.CoInitialize()

outlook = win32com.client.gencache.EnsureDispatch(
    "Outlook.Application"
)

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

start_date = f"01/01/{YEAR} 12:00 AM"
end_date = f"12/31/{YEAR} 11:59 PM"

email_count = 0

print("Counting emails...")

for folder in folders:

    try:
        filtered = folder.Items.Restrict(
            f"[ReceivedTime] >= '{start_date}' AND "
            f"[ReceivedTime] <= '{end_date}'"
        )
        email_count += filtered.Count
    except:
        pass

print(f"Total Emails      : {email_count:,}")
print(f"Already Processed : {len(processed_emails):,}")
print(f"Remaining         : {max(email_count - len(processed_emails), 0):,}")
print(f"Folders           : {len(folders):,}")

if input("Start extraction? (yes/no): ").lower() != "yes":
    raise SystemExit

extracted_count = 0
skipped_count = 0
current_email = len(processed_emails)

log_buffer = []
subject_buffer = []
unprocessed_buffer = []
processed_email_buffer = []

created_folders = set()

start_time = time.time()


def flush_buffers():

    global log_buffer
    global subject_buffer
    global unprocessed_buffer
    global processed_email_buffer

    if log_buffer:
        with open(EXTRACTION_LOG, "a", newline="", encoding="utf-8") as f:
            csv.writer(f).writerows(log_buffer)
        log_buffer.clear()

    if subject_buffer:
        with open(SUBJECT_LOG, "a", newline="", encoding="utf-8") as f:
            csv.writer(f).writerows(subject_buffer)
        subject_buffer.clear()

    if unprocessed_buffer:
        with open(UNPROCESSED_LOG, "a", newline="", encoding="utf-8") as f:
            csv.writer(f).writerows(unprocessed_buffer)
        unprocessed_buffer.clear()

    if processed_email_buffer:
        with open(PROCESSED_EMAILS_FILE, "a", encoding="utf-8") as f:
            f.write("\n".join(processed_email_buffer) + "\n")
        processed_email_buffer.clear()


for folder_no, folder in enumerate(folders, start=1):

    try:

        items = folder.Items

        filtered = items.Restrict(
            f"[ReceivedTime] >= '{start_date}' AND "
            f"[ReceivedTime] <= '{end_date}'"
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

                remaining = email_count - current_email
                eta = int(remaining / rate) if rate else 0

                print(
                    f"\rFolder {folder_no}/{len(folders)} | "
                    f"{current_email:,}/{email_count:,} | "
                    f"{pct:.2f}% | "
                    f"{rate:.1f} emails/sec | "
                    f"Files {extracted_count:,} | "
                    f"Skipped {skipped_count:,} | "
                    f"ETA {eta//3600}h {(eta%3600)//60}m",
                    end="",
                    flush=True
                )

            try:

                received = mail.ReceivedTime
                subject = getattr(mail, "Subject", "No Subject")

                subject_buffer.append([
                    received.strftime("%Y-%m-%d %H:%M:%S"),
                    subject,
                    entry_id,
                    folder.Name
                ])

                if mail.Attachments.Count == 0:

                    unprocessed_buffer.append([
                        received.strftime("%Y-%m-%d %H:%M:%S"),
                        subject,
                        entry_id,
                        "No attachments"
                    ])

                    processed_email_buffer.append(entry_id)
                    processed_emails.add(entry_id)
                    continue

                month_dir = os.path.join(
                    ROOT_OUTPUT,
                    received.strftime("%b")
                )

                if month_dir not in created_folders:
                    os.makedirs(month_dir, exist_ok=True)
                    created_folders.add(month_dir)

                found_downloadable = False

                for i in range(1, mail.Attachments.Count + 1):

                    att = mail.Attachments.Item(i)
                    filename = att.FileName

                    extension = os.path.splitext(
                        filename
                    )[1].lower()

                    if extension not in VALID_EXTENSIONS:
                        continue

                    found_downloadable = True

                    filepath = os.path.join(
                        month_dir,
                        clean_filename(filename)
                    )

                    if os.path.exists(filepath):
                        base, ext = os.path.splitext(filepath)
                        counter = 1

                        while os.path.exists(filepath):
                            filepath = (
                                f"{base}_{counter}{ext}"
                            )
                            counter += 1

                    att.SaveAsFile(filepath)

                    extracted_count += 1

                    log_buffer.append([
                        received.strftime("%Y-%m-%d %H:%M:%S"),
                        subject,
                        filename,
                        filepath,
                        entry_id
                    ])

                if not found_downloadable:

                    reason = (
                        f"No matching attachment "
                        f"({', '.join(sorted(VALID_EXTENSIONS))})"
                    )

                    unprocessed_buffer.append([
                        received.strftime("%Y-%m-%d %H:%M:%S"),
                        subject,
                        entry_id,
                        reason
                    ])

                    skipped_count += 1

                processed_email_buffer.append(entry_id)
                processed_emails.add(entry_id)

                if (
                    len(log_buffer) >= LOG_FLUSH_SIZE
                    or len(processed_email_buffer) >= EMAIL_FLUSH_SIZE
                ):
                    flush_buffers()

                    with open(
                        CHECKPOINT_FILE,
                        "w",
                        encoding="utf-8"
                    ) as f:
                        f.write(str(current_email))

            except Exception as ex:

                with open(ERROR_LOG, "a", encoding="utf-8") as ef:
                    ef.write(
                        f"{datetime.now()} | "
                        f"{entry_id} | "
                        f"{str(ex)}\n"
                    )

    except Exception as ex:

        with open(ERROR_LOG, "a", encoding="utf-8") as ef:
            ef.write(
                f"{datetime.now()} | "
                f"Folder {folder.Name} | "
                f"{str(ex)}\n"
            )

flush_buffers()

duration = (
    datetime.now()
    - datetime.fromtimestamp(start_time)
)

print("\n")
print("================================")
print("EXTRACTION COMPLETED")
print("================================")
print(f"Emails Total      : {email_count:,}")
print(f"Emails Processed  : {current_email:,}")
print(f"Files Extracted   : {extracted_count:,}")
print(f"Skipped Emails    : {skipped_count:,}")
print(f"Duration          : {duration}")
print(f"Extraction Log    : {EXTRACTION_LOG}")
print(f"Subject Log       : {SUBJECT_LOG}")
print(f"Unprocessed Log   : {UNPROCESSED_LOG}")
print(f"Output Folder     : {ROOT_OUTPUT}")
