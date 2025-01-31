import pandas as pd
import os
import logging
import json
import base64
from tqdm import tqdm
import re
from mailjet_rest import Client

# Load email configuration from a config file
def load_config():
    with open('mailjetconfig.json', 'r') as f:
        return json.load(f)

config = load_config()

# Mailjet credentials
mailjet_api_key = config["mailjet_api_key"]
mailjet_api_secret = config["mailjet_api_secret"]
mailjet = Client(auth=(mailjet_api_key, mailjet_api_secret), version='v3.1')

# Email subject and body template with certificate number and father's name
subject_template = "Congratulations! Your Certificate for {course_name} is Ready"
body_template = """
<html>
  <body>
    <p>Dear {full_name},</p>
    
    <p>We are pleased to share your certificate (Certificate No. {certificate_no}) and scorecard for the "{course_name}" course, 
    conducted from {start_date} to {end_date}. Please find the attached documents for your reference.</p>
    
    <p><strong>Here are your details:</strong></p>
    <ul>
      <li><strong>Certificate Number:</strong> {certificate_no}</li>
      <li><strong>Full Name:</strong> {full_name}</li>
      <li><strong>Course Name:</strong> {course_name}</li>
    </ul>

    <p><strong>Instructions:</strong></p>
    <p>We encourage you to share your achievement with your professional network by uploading your certificate on LinkedIn. Here's how:</p>
    <ol>
      <li>Log in to your <a href="https://www.linkedin.com" target="_blank">LinkedIn account</a>.</li>
      <li>Go to your profile and click on "Add Profile Section" &gt; "Accomplishments" &gt; "Certifications".</li>
      <li>Enter the following details:</li>
      <ul>
        <li><strong>Certification Name:</strong> {course_name}</li>
        <li><strong>Issuing Organization:</strong> NIELIT Ropar</li>
        <li><strong>Credential ID:</strong> {certificate_no}</li>
      </ul>
      <li>Upload your certificate PDF (attached to this email).</li>
    </ol>

    <p>Don't forget to tag us in your LinkedIn post:</p>
    <ul>
      <li><a href="https://www.linkedin.com/school/nielitindia/" target="_blank">NIELIT India</a>, <a href="https://in.linkedin.com/in/imsarwansingh" target="_blank"> Dr. Sarwan Singh</a></li>
    </ul>
    
    <p>We look forward to seeing your post and celebrating your success!</p>
    
    <p>Best regards,<br/>
    NIELIT Chandigarh</p>
  </body>
</html>
"""

# File paths
excel_file = config["excel_file"]
attachments_folder = config["attachments_folder"]

# Email validation regex
email_regex = r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'

# Set up logging to track email sends
logging.basicConfig(filename="email_sending.log", level=logging.INFO, 
                    format="%(asctime)s - %(levelname)s - %(message)s")

# Function to validate email format
def is_valid_email(email):
    return re.match(email_regex, email)

# Read Excel and process emails
data = pd.read_excel(excel_file)

# Validate required columns
required_columns = {'full_name', 'father_name', 'email', 'course_name', 'start_date', 'end_date', 'issue_date', 'roll_no', 'cert_no'}
if not required_columns.issubset(data.columns):
    missing_cols = required_columns - set(data.columns)
    raise ValueError(f"Missing required columns in Excel file: {missing_cols}")

total_emails = len(data)
emails_sent = 0

# Initialize progress bar
with tqdm(total=total_emails, desc="Sending Emails", unit="email") as pbar:
    for _, row in data.iterrows():
        full_name = row['full_name']
        father_name = row['father_name']
        to_email = row['email']
        course_name = row['course_name']
        start_date = row['start_date']
        end_date = row['end_date']
        issue_date = row['issue_date']
        roll_no = row['roll_no']
        certificate_no = row['cert_no']

        # Validate email address
        if not is_valid_email(to_email):
            logging.error(f"Invalid email address for {full_name} (Roll No: {roll_no}). Email: {to_email}")
            continue

        # File paths
        folder_name = f"{roll_no}_{full_name}"
        certificate_file = os.path.join(attachments_folder, folder_name, f"{folder_name}_certificate.pdf")
        scorecard_file = os.path.join(attachments_folder, folder_name, f"{folder_name}_scorecard.pdf")
        logo_file = 'logo.png'  # Updated path for logo file

        try:
            with open(certificate_file, 'rb') as cert_file, open(scorecard_file, 'rb') as scorecard_file, open(logo_file, 'rb') as logo_file_handle:
                certificate_data = base64.b64encode(cert_file.read()).decode()
                scorecard_data = base64.b64encode(scorecard_file.read()).decode()
                logo_data = base64.b64encode(logo_file_handle.read()).decode()

                # Prepare Mailjet email payload
                data = {
                    'Messages': [
                        {
                            "From": {
                                "Email": config["from_email"],
                                "Name": "NIELIT Chandigarh"
                            },
                            "To": [
                                {
                                    "Email": to_email,
                                    "Name": full_name
                                }
                            ],
                            "Subject": subject_template.format(course_name=course_name),
                            "HTMLPart": body_template.format(
                                full_name=full_name,
                                certificate_no=certificate_no,
                                course_name=course_name,
                                start_date=start_date,
                                end_date=end_date
                            ),
                            "Attachments": [
                                {
                                    "ContentType": "application/pdf",
                                    "Filename": f"{folder_name}_certificate.pdf",
                                    "Base64Content": certificate_data
                                },
                                {
                                    "ContentType": "application/pdf",
                                    "Filename": f"{folder_name}_scorecard.pdf",
                                    "Base64Content": scorecard_data
                                }
                            ],
                            "InlineAttachments": [
                                {
                                    "ContentType": "image/png",
                                    "Filename": "logo.png",
                                    "Base64Content": logo_data,
                                    "ContentID": "logo"
                                }
                            ]
                        }
                    ]
                }

                # Send email using Mailjet API
                result = mailjet.send.create(data=data)

                if result.status_code == 200:
                    logging.info(f"Successfully sent email to {to_email} (Roll No: {roll_no})")
                    emails_sent += 1
                else:
                    logging.error(f"Failed to send email to {to_email} (Roll No: {roll_no}). Error: {result.text}")
        except FileNotFoundError as e:
            logging.warning(f"Files not found for {full_name} (Roll No: {roll_no}). Skipping.")
        except Exception as e:
            logging.error(f"An error occurred while processing {full_name} (Roll No: {roll_no}). Error: {str(e)}")

        pbar.update(1)

# Log email summary
summary_message = f"Total emails: {total_emails}\nSent: {emails_sent}\nFailed: {total_emails - emails_sent}"
logging.info(summary_message)

# Send summary email
summary_data = {
    'Messages': [
        {
            "From": {
                "Email": config["from_email"],
                "Name": "NIELIT Chandigarh"
            },
            "To": [
                {
                    "Email": config["from_email"],
                    "Name": "Administrator"
                }
            ],
            "Subject": "Email Sending Summary",
            "TextPart": summary_message
        }
    ]
}
mailjet.send.create(data=summary_data)
