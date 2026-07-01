import win32com.client  
import time  
import os  
import csv
from datetime import datetime  
import pytz  
import logging  
import pandas as pd  
import warnings  
import smtplib  
from email.mime.multipart import MIMEMultipart  
from email.mime.text import MIMEText  
from email.mime.application import MIMEApplication
 
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")  
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')  
  
class EmailProcessor:  
    def __init__(self):  
        # Log script startup time and store it to ignore older emails
        self.startup_time = datetime.now(pytz.UTC)
        logging.info(f"Script started at: {self.startup_time} (UTC)")
        
        # Track processed EntryIDs to prevent reprocessing during the same run
        self.processed_entry_ids = set()
        
        # Initialize CSV Log file
        self.csv_log_file = "processed_emails_log.csv"
        self._initialize_csv_log()

    def _initialize_csv_log(self):
        """Creates the CSV file and header if it doesn't exist."""
        if not os.path.exists(self.csv_log_file):
            with open(self.csv_log_file, mode='w', newline='', encoding='utf-8') as file:
                writer = csv.writer(file)
                writer.writerow(["Timestamp", "EntryID", "SenderEmail", "Subject", "ReportGenerated"])

    def log_to_csv(self, entry_id, sender, subject, report_generated):
        """Logs processed email details to the CSV file."""
        with open(self.csv_log_file, mode='a', newline='', encoding='utf-8') as file:
            writer = csv.writer(file)
            writer.writerow([datetime.now().strftime("%Y-%m-%d %H:%M:%S"), entry_id, sender, subject, report_generated])
  
    def filter_and_save_report(self, file_path, sheet_name, header_row, app_column, report_name, application_code):  
        """Filters the Excel sheet based on application code and saves as an HTML report."""
        try:
            df = pd.read_excel(file_path, sheet_name=sheet_name, header=header_row)  
            df.columns = df.columns.str.strip()  
    
            if app_column in df.columns and 'Status' in df.columns:  
                filtered_df = df[df[app_column].str.contains(application_code, case=False, na=False)]  
    
                if 'Unique' in df.columns and 'Status' in df.columns:  
                    if not filtered_df.empty:  
                        result = filtered_df[['Unique','Application Name','Severity','Past due','Owner','Status','Application Code' ]]  
    
                        status_options = result['Status'].dropna().unique().tolist()  
                        dropdown_html = '<select id="statusDropdown" onchange="filterTable()">'  
                        dropdown_html += '<option value="All">All</option>'  
                        for status in status_options:  
                            dropdown_html += f'<option value="{status}">{status}</option>'  
                        dropdown_html += '</select>'  
    
                        javascript = '''  
                        <script>  
                            function filterTable() {  
                                var dropdown = document.getElementById("statusDropdown");  
                                var filter = dropdown.value;  
                                var table = document.getElementById("resultTable");  
                                var rows = table.getElementsByTagName("tr");  
    
                                for (var i = 1; i < rows.length; i++) {  
                                    var statusCell = rows[i].getElementsByTagName("td")[5];  
                                    if (statusCell) {  
                                        var status = statusCell.textContent || statusCell.innerText;  
                                        if (filter === "All" || status === filter) {  
                                            rows[i].style.display = "";  
                                        } else {  
                                            rows[i].style.display = "none";  
                                        }  
                                    }  
                                }  
                            }  
                        </script>  
                        '''  
    
                        table_html = result.to_html(index=False, table_id="resultTable")  
                        current_datetime = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")  
                        
                        # Note: Ensure the C:\Python\Application\CCB\ directory exists
                        output_dir = r'C:\Python\Application\CCB'
                        os.makedirs(output_dir, exist_ok=True)
                        
                        output_html_path = os.path.join(output_dir, f'{application_code}_{current_datetime}_{report_name}.html')  
                        
                        full_html = f'''  
                        <html>  
                        <head><title>{report_name} - Data for {application_code}</title></head>  
                        <body>  
                            <h2>{report_name} - Data for {application_code}</h2>  
                            <label for="statusDropdown">Status: </label>  
                            {dropdown_html}  
                            {javascript}  
                            {table_html}  
                        </body>  
                        </html>  
                        '''  
                        with open(output_html_path, 'w', encoding='utf-8') as f:  
                            f.write(full_html)  
    
                        logging.info(f"{report_name} saved successfully: {output_html_path}")  
                        return output_html_path 
                    else:  
                        logging.info(f"No matching records found for application code '{application_code}' in {report_name}.")  
                else:  
                    logging.warning(f"Required columns not found in {report_name}.")  
            else:  
                logging.warning(f"'{app_column}' or 'Status' missing in {report_name}.")  
        except Exception as e:
            logging.error(f"Error generating report for {application_code}: {e}")
        
        return None
  
    def send_email_with_attachment(self, to_email, subject, body, attachment_path=None):  
        """Sends an email via SMTP and attaches the generated file."""
        from_email = "Infosec_VAPT@crisil.com" 
        smtp_server = "mail.ad.crisil.com" 
        smtp_port = 25 

        try: 
            msg = MIMEMultipart() 
            msg['From'] = from_email 
            msg['To'] = to_email 
            msg['Subject'] = subject 
            msg.attach(MIMEText(body, 'plain')) 

            # Attach the file if it was successfully generated
            if attachment_path and os.path.exists(attachment_path):
                with open(attachment_path, "rb") as f:
                    part = MIMEApplication(f.read(), Name=os.path.basename(attachment_path))
                
                # Add header to make it a downloadable attachment
                part['Content-Disposition'] = f'attachment; filename="{os.path.basename(attachment_path)}"'
                msg.attach(part)
            else:
                logging.warning("No attachment found to send.")

            server = smtplib.SMTP(smtp_server, smtp_port) 
            server.sendmail(from_email, [to_email], msg.as_string()) 
            server.quit() 
            logging.info(f"Successfully replied to {to_email} with report.") 
        except Exception as e: 
            logging.error(f"Error sending email to {to_email}: {e}") 

    def get_sender_smtp_address(self, message):
        """Helper to extract standard SMTP address, avoiding Exchange X.500 address formats."""
        try:
            if message.SenderEmailType == "EX":
                return message.Sender.GetExchangeUser().PrimarySmtpAddress
            return message.SenderEmailAddress
        except Exception:
            return message.SenderEmailAddress
 
    def read_outlook_inbox(self):  
        try:  
            outlook = win32com.client.Dispatch("Outlook.Application")  
            namespace = outlook.GetNamespace("MAPI")  
            inbox = namespace.GetDefaultFolder(6)  # 6 = Inbox
  
            # Restrict to Unread emails only
            messages = inbox.Items.Restrict("[UnRead] = True")  
            messages.Sort('ReceivedTime', True)  
  
            for message in messages:  
                entry_id = message.EntryID

                # Skip if we already processed this exact email in this session
                if entry_id in self.processed_entry_ids:
                    continue
                
                # Ensure date is timezone aware in UTC for accurate comparison
                message_dt = datetime.fromtimestamp(message.ReceivedTime.timestamp(), tz=pytz.UTC)  

                # Ignore emails received before the script started
                if message_dt < self.startup_time:
                    continue

                application_code = message.Subject 
                sender_email = self.get_sender_smtp_address(message)

                if application_code:
                    logging.info(f"Processing new request for: {application_code} from {sender_email}")  
                    
                    reports = [  
                        {
                            "file_path": r'C:\Users\c-abhimanyue\OneDrive - crisil.com\All Trackers - CRISIL\Application VAPT\Application Vulnerability Tracker.xlsx', 
                            "sheet_name": "Advance Search", 
                            "header_row": 3, 
                            "app_column": 'Application Code', 
                            "report_name": "Application VAPT"
                        }
                    ]  
  
                    attachments = []
                    # Run the report filter and save
                    for report in reports:  
                        attachment_path = self.filter_and_save_report(**report, application_code=application_code)  
                        if attachment_path:
                            attachments.append(attachment_path)

                    # Revert to the same recipient with the attachment
                    for attachment_path in attachments:
                        self.send_email_with_attachment( 
                            to_email=sender_email, 
                            subject=f"RE: {application_code} - Requested Report", 
                            body=f"Hello,\n\nPlease find the requested VAPT report for {application_code} attached.\n\nRegards,\nInfosec Team",
                            attachment_path=attachment_path
                        ) 
                    
                    # Log to CSV
                    self.log_to_csv(entry_id, sender_email, application_code, bool(attachments))

                # Mark the email as Read and add to processed set
                message.UnRead = False
                self.processed_entry_ids.add(entry_id)

        except Exception as e:  
            logging.error(f"An error occurred while reading Outlook: {e}")  
  
if __name__ == "__main__":  
    processor = EmailProcessor()  
    logging.info("Starting email monitoring loop. Press Ctrl+C to stop.")
    
    try:
        while True:  
            processor.read_outlook_inbox()  
            time.sleep(5)  # Check every 5 seconds 
    except KeyboardInterrupt:
        logging.info("Script manually terminated.")
