import os
import re
import csv
import pythoncom
import win32com.client
from datetime import datetime

# ===========================================
# CONFIGURATION
# ===========================================

OUTPUT_ROOT = r"D:\PDF_Extraction"

PROCESSED_EMAILS_FILE = os.path.join(
    OUTPUT_ROOT,
    "processed_emails.csv"
)

EXTRACTION_LOG = os.path.join(
    OUTPUT_ROOT,
    "extracted_pdfs.csv"
)

os.makedirs(
    OUTPUT_ROOT,
    exist_ok=True
)

# ===========================================
# HELPERS
# ===========================================

def clean_filename(name):
    return re.sub(
        r'[<>:"/\\|?*]',
        "_",
        str(name)
    )

# ===========================================
# LOAD PROCESSED EMAILS
# ===========================================

processed_emails = set()

if os.path.exists(PROCESSED_EMAILS_FILE):

    print("Loading processed emails...")

    with open(
        PROCESSED_EMAILS_FILE,
        "r",
        encoding="utf-8"
    ) as f:

        for line in f:

            entry_id = line.strip()

            if entry_id:
                processed_emails.add(
                    entry_id
                )

print(
    f"Processed Emails Loaded: "
    f"{len(processed_emails):,}"
)

# ===========================================
# CREATE EXTRACTION LOG
# ===========================================

if not os.path.exists(EXTRACTION_LOG):

    with open(
        EXTRACTION_LOG,
        "w",
        newline="",
        encoding="utf-8"
    ) as f:

        writer = csv.writer(f)

        writer.writerow([
            "ReceivedDate",
            "Subject",
            "AttachmentName",
            "SavedPath",
            "EntryID"
        ])

# ===========================================
# CONNECT OUTLOOK
# ===========================================

pythoncom.CoInitialize()

outlook = (
    win32com.client.gencache.EnsureDispatch(
        "Outlook.Application"
    )
)

namespace = outlook.GetNamespace("MAPI")

archive_store = None

for store in namespace.Folders:

    if "Online Archive" in store.Name:

        archive_store = store
        break

if archive_store is None:
    raise Exception(
        "Online Archive mailbox not found"
    )

archive_inbox = archive_store.Folders[
    "Inbox"
]

# ===========================================
# COLLECT FOLDERS
# ===========================================

folders = []

def collect_folders(folder):

    folders.append(folder)

    for sub in folder.Folders:
        collect_folders(sub)

collect_folders(
    archive_inbox
)

print(
    f"Folders Found: "
    f"{len(folders):,}"
)

# ===========================================
# PROCESS UNPROCESSED EMAILS
# ===========================================

pdf_count = 0
email_count = 0
log_buffer = []

for folder_no, folder in enumerate(
    folders,
    start=1
):

    try:

        items = folder.Items

        try:
            items.Sort(
                "[ReceivedTime]",
                True
            )
        except:
            pass

        for mail in items:

            try:

                if mail.Class != 43:
                    continue

                entry_id = mail.EntryID

                # Skip processed emails
                if (
                    entry_id
                    in processed_emails
                ):
                    continue

                email_count += 1

                # Skip emails with no attachments
                if (
                    mail.Attachments.Count == 0
                ):
                    continue

                received = mail.ReceivedTime

                month_folder = (
                    received.strftime("%b")
                )

                save_dir = os.path.join(
                    OUTPUT_ROOT,
                    month_folder
                )

                os.makedirs(
                    save_dir,
                    exist_ok=True
                )

                subject = getattr(
                    mail,
                    "Subject",
                    "No Subject"
                )

                found_pdf = False

                for i in range(
                    1,
                    mail.Attachments.Count + 1
                ):

                    att = (
                        mail.Attachments.Item(i)
                    )

                    filename = att.FileName

                    if not filename.lower().endswith(
                        ".pdf"
                    ):
                        continue

                    found_pdf = True

                    filepath = os.path.join(
                        save_dir,
                        clean_filename(
                            filename
                        )
                    )

                    # Handle duplicates
                    if os.path.exists(
                        filepath
                    ):

                        base, ext = (
                            os.path.splitext(
                                filepath
                            )
                        )

                        counter = 1

                        while os.path.exists(
                            filepath
                        ):

                            filepath = (
                                f"{base}_{counter}{ext}"
                            )

                            counter += 1

                    att.SaveAsFile(
                        filepath
                    )

                    pdf_count += 1

                    log_buffer.append([
                        received.strftime(
                            "%Y-%m-%d %H:%M:%S"
                        ),
                        subject,
                        filename,
                        filepath,
                        entry_id
                    ])

                # Mark processed only if PDF found
                if found_pdf:

                    with open(
                        PROCESSED_EMAILS_FILE,
                        "a",
                        encoding="utf-8"
                    ) as f:

                        f.write(
                            entry_id + "\n"
                        )

                    processed_emails.add(
                        entry_id
                    )

                # Flush logs
                if (
                    len(log_buffer)
                    >= 500
                ):

                    with open(
                        EXTRACTION_LOG,
                        "a",
                        newline="",
                        encoding="utf-8"
                    ) as f:

                        csv.writer(
                            f
                        ).writerows(
                            log_buffer
                        )

                    log_buffer.clear()

                # Progress
                if email_count % 100 == 0:

                    print(
                        f"\rFolder "
                        f"{folder_no}/{len(folders)} | "
                        f"Emails Checked: "
                        f"{email_count:,} | "
                        f"PDFs Saved: "
                        f"{pdf_count:,}",
                        end="",
                        flush=True
                    )

            except Exception:
                continue

    except Exception:
        continue

# Final log flush
if log_buffer:

    with open(
        EXTRACTION_LOG,
        "a",
        newline="",
        encoding="utf-8"
    ) as f:

        csv.writer(
            f
        ).writerows(
            log_buffer
        )

print("\n")
print("================================")
print("COMPLETED")
print("================================")
print(
    f"Emails Checked : "
    f"{email_count:,}"
)
print(
    f"PDFs Extracted : "
    f"{pdf_count:,}"
)
print(
    f"Output Folder  : "
    f"{OUTPUT_ROOT}"
)
print(
    f"Extraction Log : "
    f"{EXTRACTION_LOG}"
)
