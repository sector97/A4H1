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
        self.startup_time = datetime.now(pytz.UTC)
        logging.info(f"Script started at: {self.startup_time} (UTC)")
        
        self.processed_entry_ids = set()
        self.csv_log_file = "processed_emails_log.csv"
        self._initialize_csv_log()

    def _initialize_csv_log(self):
        if not os.path.exists(self.csv_log_file):
            with open(self.csv_log_file, mode='w', newline='', encoding='utf-8') as file:
                writer = csv.writer(file)
                writer.writerow(["Timestamp", "EntryID", "SenderEmail", "Subject", "ReportGenerated"])

    def log_to_csv(self, entry_id, sender, subject, report_generated):
        with open(self.csv_log_file, mode='a', newline='', encoding='utf-8') as file:
            writer = csv.writer(file)
            writer.writerow([datetime.now().strftime("%Y-%m-%d %H:%M:%S"), entry_id, sender, subject, report_generated])
  
    def filter_and_save_report(self, file_path, sheet_name, header_row, app_column, report_name, application_code, columns_to_extract):  
        """Filters the Excel sheet dynamically and saves as an HTML report."""
        try:
            # Handle multiple sheets by concatenating them
            if isinstance(sheet_name, list):
                df_list = []
                for sheet in sheet_name:
                    temp_df = pd.read_excel(file_path, sheet_name=sheet, header=header_row)
                    temp_df.columns = temp_df.columns.str.strip()
                    df_list.append(temp_df)
                df = pd.concat(df_list, ignore_index=True)
            else:
                df = pd.read_excel(file_path, sheet_name=sheet_name, header=header_row)  
                df.columns = df.columns.str.strip()  
    
            if app_column in df.columns:  
                filtered_df = df[df[app_column].str.contains(application_code, case=False, na=False)]  
    
                if not filtered_df.empty:  
                    # Ensure we only try to extract columns that actually exist in the file
                    valid_columns = [col for col in columns_to_extract if col in filtered_df.columns]
                    
                    if not valid_columns:
                        logging.warning(f"None of the specified columns were found in {report_name}.")
                        return None

                    result = filtered_df[valid_columns]  

                    # Dynamically determine Status column index for JavaScript
                    status_col_index = -1
                    dropdown_html = ""
                    javascript = ""
                    
                    if 'Status' in result.columns:
                        status_col_index = result.columns.get_loc('Status')
                        status_options = result['Status'].dropna().unique().tolist()  
                        dropdown_html = '<select id="statusDropdown" onchange="filterTable()">'  
                        dropdown_html += '<option value="All">All</option>'  
                        for status in status_options:  
                            dropdown_html += f'<option value="{status}">{status}</option>'  
                        dropdown_html += '</select>'  
    
                        javascript = f'''  
                        <script>  
                            function filterTable() {{  
                                var dropdown = document.getElementById("statusDropdown");  
                                var filter = dropdown.value;  
                                var table = document.getElementById("resultTable");  
                                var rows = table.getElementsByTagName("tr");  
    
                                for (var i = 1; i < rows.length; i++) {{  
                                    // Dynamically injected index for Status column
                                    var statusCell = rows[i].getElementsByTagName("td")[{status_col_index}];  
                                    if (statusCell) {{  
                                        var status = statusCell.textContent || statusCell.innerText;  
                                        if (filter === "All" || status === filter) {{  
                                            rows[i].style.display = "";  
                                        }} else {{  
                                            rows[i].style.display = "none";  
                                        }}  
                                    }}  
                                }}  
                            }}  
                        </script>  
                        '''  
    
                    table_html = result.to_html(index=False, table_id="resultTable")  
                    current_datetime = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")  
                    output_dir = r'C:\Python\Application\CCB'
                    os.makedirs(output_dir, exist_ok=True)
                    
                    output_html_path = os.path.join(output_dir, f'{application_code}_{current_datetime}_{report_name}.html')  
                    
                    full_html = f'''  
                    <html>  
                    <head><title>{report_name} - Data for {application_code}</title></head>  
                    <body>  
                        <h2>{report_name} - Data for {application_code}</h2>  
                        {f'<label for="statusDropdown">Status: </label>' if dropdown_html else ''}
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
                logging.warning(f"'{app_column}' column missing in {report_name}.")  
        except Exception as e:
            logging.error(f"Error generating report for {application_code} in {report_name}: {e}")
        
        return None
  
    def send_email_with_attachments(self, to_email, subject, body, attachment_paths):  
        """Sends an email via SMTP and attaches all generated files."""
        from_email = "Infosec_VAPT@crisil.com" 
        smtp_server = "mail.ad.crisil.com" 
        smtp_port = 25 

        try: 
            msg = MIMEMultipart() 
            msg['From'] = from_email 
            msg['To'] = to_email 
            msg['Subject'] = subject 
            msg.attach(MIMEText(body, 'plain')) 

            # Attach all successfully generated files
            for path in attachment_paths:
                if path and os.path.exists(path):
                    with open(path, "rb") as f:
                        part = MIMEApplication(f.read(), Name=os.path.basename(path))
                    part['Content-Disposition'] = f'attachment; filename="{os.path.basename(path)}"'
                    msg.attach(part)

            server = smtplib.SMTP(smtp_server, smtp_port) 
            server.sendmail(from_email, [to_email], msg.as_string()) 
            server.quit() 
            logging.info(f"Successfully replied to {to_email} with reports.") 
        except Exception as e: 
            logging.error(f"Error sending email to {to_email}: {e}") 

    def get_sender_smtp_address(self, message):
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
            inbox = namespace.GetDefaultFolder(6)  
  
            messages = inbox.Items.Restrict("[UnRead] = True")  
            messages.Sort('ReceivedTime', True)  
  
            for message in messages:  
                entry_id = message.EntryID

                if entry_id in self.processed_entry_ids:
                    continue
                
                message_dt = datetime.fromtimestamp(message.ReceivedTime.timestamp(), tz=pytz.UTC)  

                if message_dt < self.startup_time:
                    continue

                application_code = message.Subject 
                sender_email = self.get_sender_smtp_address(message)

                if application_code:
                    logging.info(f"Processing new request for: {application_code} from {sender_email}")  
                    
                    # VAPT, SAST, and SCA Configurations
                    reports = [  
                        {
                            "file_path": r'C:\Users\c-abhimanyue\OneDrive - crisil.com\All Trackers - CRISIL\Application VAPT\Application Vulnerability Tracker.xlsx', 
                            "sheet_name": "Advance Search", 
                            "header_row": 3, 
                            "app_column": 'Application Code', 
                            "report_name": "Application VAPT",
                            "columns_to_extract": [
                                'Unique', 'Application Name', 'Severity', 'Past due', 
                                'Owner', 'Status', 'Application Code'
                            ]
                        },
                        {
                            "file_path": r'C:\Users\c-abhimanyue\OneDrive - crisil.com\All Trackers - CRISIL\ SAST\SAST Updated Tracker.xlsx', 
                            "sheet_name": "Tracker", 
                            "header_row": 0,  
                            "app_column": 'Application Code', 
                            "report_name": "SAST",
                            "columns_to_extract": [
                                'Application Code', 'Application Name', 'Severity', 
                                'Vulnerability Count', 'Date of Review / Report', 
                                'Revalidation vulnerability Count', 'Closed', 
                                'Target Closure Date', 'Past due', 'Ageing', 
                                'Status', 'Application SPOC'
                            ]
                        },
                        {
                            "file_path": r'C:\Users\c-abhimanyue\OneDrive - crisil.com\All Trackers - CRISIL\ SCA\Automation\Dashboard_v3.xlsx', 
                            "sheet_name": ["SLA Tracker 2025", "SLA Tracker 2026"],  # Process both sheets 
                            "header_row": 0,  
                            "app_column": 'Application Code', 
                            "report_name": "SCA",
                            "columns_to_extract": [
                                'Application Code', 'Application Name', 'severity', 
                                'Latest Start Date', 'Count', 'SLA closure date', 
                                'Past Due', 'SLA Email', 'Status', 'Closed', 
                                'Application Owner'
                            ]
                        }
                    ]  
  
                    generated_attachments = []
                    
                    # Run the report filter and save for all configs
                    for report in reports:  
                        attachment_path = self.filter_and_save_report(**report, application_code=application_code)  
                        if attachment_path:
                            generated_attachments.append(attachment_path)

                    # Only send an email if at least one report was generated
                    if generated_attachments:
                        self.send_email_with_attachments( 
                            to_email=sender_email, 
                            subject=f"RE: {application_code} - Requested Reports", 
                            body=f"Hello,\n\nPlease find the requested security reports (VAPT/SAST/SCA) for {application_code} attached.\n\nRegards,\nInfosec Team",
                            attachment_paths=generated_attachments
                        ) 
                    
                    self.log_to_csv(entry_id, sender_email, application_code, bool(generated_attachments))

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
            time.sleep(5) 
    except KeyboardInterrupt:
        logging.info("Script manually terminated.")
