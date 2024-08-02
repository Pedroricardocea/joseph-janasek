import xml.etree.ElementTree as ET
import csv
import data_formatter_reach as data_formatting
import pandas as pd
import gmail_auto 
import reach_automation_FOC as reach
import time
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders


username = "reportdownloads@revolutionptwl.com"
password = "reports1!"
    
reach.ReachAutomation(username, password)


time.sleep(60)

username = "reportdownloads@revolutionptwl.com"
password = "uprt prnd hixl msoc"  # App-specific password
imap_url = "imap.gmail.com"
save_path = "/Users/pedrocastro/Desktop/Revolution/FOC/Leads Report/Data/raw_data"

gmail_auto.gmail(username, password, imap_url, save_path)


# Load and parse the XML file
tree = ET.parse("/Users/pedrocastro/Desktop/Revolution/FOC/Leads Report/Data/raw_data/leads.xls")
root = tree.getroot()

# Extract the column names from the first row
namespace = {'ss': 'urn:schemas-microsoft-com:office:spreadsheet'}
worksheet = root.find('ss:Worksheet', namespace)
table = worksheet.find('ss:Table', namespace)
header_row = table.find('ss:Row', namespace)

headers = [cell.find('ss:Data', namespace).text for cell in header_row.findall('ss:Cell', namespace)]

# Prepare to write to CSV
with open('/Users/pedrocastro/Desktop/Revolution/FOC/Leads Report/Data/raw_data/leads.csv', 'w', newline='') as csvfile:
    csvwriter = csv.writer(csvfile)
    csvwriter.writerow(headers)  # Write header row

    # Iterate over remaining rows
    for row in table.findall('ss:Row', namespace)[1:]:
        cells = [cell.find('ss:Data', namespace).text for cell in row.findall('ss:Cell', namespace)]
        csvwriter.writerow(cells)  # Write data rows

print("XML to CSV conversion completed successfully.")



# Define file paths
scheduler_path = "/Users/pedrocastro/Desktop/Revolution/FOC/Leads Report/Data/raw_data/leads.csv"
output_path = "/Users/pedrocastro/Desktop/Revolution/FOC/Leads Report/Data/formatted_data/formatted_leads.csv"

# Load and clean data
data_formatting.load_and_clean_data(scheduler_path, output_path)

#turn output_path to a pandas dataframe and to excel file
df = pd.read_csv(output_path)
df.to_excel("/Users/pedrocastro/Desktop/Revolution/FOC/Leads Report/Data/formatted_data/formatted_leads.xlsx", index=False)


# Define the sender and recipient email addresses
sender_email = "reportdownloads@revolutionptwl.com"
recipient_email = "reportdownloads@revolutionptwl.com"

# Create a multipart message object
message = MIMEMultipart()
message["From"] = sender_email
message["To"] = recipient_email
message["Subject"] = "Leads Report"

# Add the file as an attachment
attachment_path = "/Users/pedrocastro/Desktop/Revolution/FOC/Leads Report/Data/formatted_data/formatted_leads.xlsx"
attachment_name = "formatted_leads.xlsx"

with open(attachment_path, "rb") as attachment:
    part = MIMEBase("application", "octet-stream")
    part.set_payload(attachment.read())

encoders.encode_base64(part)
part.add_header("Content-Disposition", f"attachment; filename= {attachment_name}")

message.attach(part)

# Connect to the SMTP server and send the email
smtp_server = "smtp.gmail.com"
smtp_port = 587
smtp_username = "reportdownloads@revolutionptwl.com"
smtp_password = "uprt prnd hixl msoc"

with smtplib.SMTP(smtp_server, smtp_port) as server:
    server.starttls()
    server.login(smtp_username, smtp_password)
    server.send_message(message)

print("Email sent successfully.")
