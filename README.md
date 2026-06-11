import win32com.client
import os
import re
from collections import defaultdict

YEAR = int(input("Enter Year (Example: 2025): "))

ROOT_OUTPUT = fr"D:\Archive_PDFs\{YEAR}"

def clean_filename(name):
    return re.sub(r'[<>:"/\\|?*]', '_', name)

outlook = win32com.client.Dispatch("Outlook.Application")
namespace = outlook.GetNamespace("MAPI")

archive_store = None

for store in namespace.Folders:
    if "Online Archive" in store.Name:
        archive_store = store
        break

if archive_store is None:
    raise Exception("Online Archive mailbox not found")

archive_inbox = archive_store.Folders["Inbox"]

email_count = 0
pdf_count = 0
folder_stats = defaultdict(int)

print("Counting emails...")

def count_emails(folder):

    global email_count

    try:
        items = folder.Items

        start_date = f"01/01/{YEAR} 12:00 AM"
        end_date = f"12/31/{YEAR} 11:59 PM"

        filtered = items.Restrict(
            f"[ReceivedTime] >= '{start_date}' AND [ReceivedTime] <= '{end_date}'"
        )

        email_count += filtered.Count

    except Exception:
        pass

    for subfolder in folder.Folders:
        count_emails(subfolder)

count_emails(archive_inbox)

print(f"\nEmails found for {YEAR}: {email_count:,}")

if input("Start PDF extraction? (yes/no): ").lower() != "yes":
    exit()

print("\nStarting extraction...\n")

def extract_pdfs(folder):

    global pdf_count

    try:

        items = folder.Items

        start_date = f"01/01/{YEAR} 12:00 AM"
        end_date = f"12/31/{YEAR} 11:59 PM"

        filtered = items.Restrict(
            f"[ReceivedTime] >= '{start_date}' AND [ReceivedTime] <= '{end_date}'"
        )

        for mail in filtered:

            try:

                received = mail.ReceivedTime
                month_folder = received.strftime("%b")

                save_dir = os.path.join(
                    ROOT_OUTPUT,
                    month_folder
                )

                os.makedirs(save_dir, exist_ok=True)

                for i in range(1, mail.Attachments.Count + 1):

                    att = mail.Attachments.Item(i)

                    if att.FileName.lower().endswith(".pdf"):

                        filename = clean_filename(att.FileName)

                        filepath = os.path.join(save_dir, filename)

                        if os.path.exists(filepath):

                            base, ext = os.path.splitext(filepath)

                            n = 1
                            while os.path.exists(filepath):
                                filepath = f"{base}_{n}{ext}"
                                n += 1

                        att.SaveAsFile(filepath)

                        pdf_count += 1
                        folder_stats[month_folder] += 1

                        if pdf_count % 100 == 0:
                            print(f"{pdf_count:,} PDFs extracted...")

            except Exception:
                continue

    except Exception as e:
        print(f"Folder error: {folder.Name} : {e}")

    for subfolder in folder.Folders:
        extract_pdfs(subfolder)

extract_pdfs(archive_inbox)

print("\n==========================")
print("EXTRACTION COMPLETED")
print("==========================")
print(f"Total PDFs : {pdf_count:,}")

for month in sorted(folder_stats.keys()):
    print(f"{month}: {folder_stats[month]:,}")

print(f"\nSaved to: {ROOT_OUTPUT}")
