import pandas as pd
import os
import logging
import json
import base64
from tqdm import tqdm
import re
from mailjet_rest import Client

# Load email configuration from a JSON object
config = {
    "from_email": "xxxxxxxxxxx@gmail.com",               #replace with your mail id registered on mailjet
    "excel_file": "TEST.xlsx",                           #path of the excel sheet
    "attachments_folder": "certificates",                #path of the attachments_folder where certificates are present
    "mailjet_api_key": "90972xxxxxxxxxxxxxxxxxe256d1c427",#replace with your mailjet_api_key" get it for free from mailjet
    "mailjet_api_secret": "49f0cxxxxxxxxxxxxxxxxxxxxxxx"  #replace with your mailjet_api_secret" get it for free from mailjet
}

# Initialize Mailjet Client
mailjet = Client(auth=(config["mailjet_api_key"], config["mailjet_api_secret"]), version='v3.1')

# Email subject and body template
subject_template = "Congratulations! Your Certificate is Ready"
body_template = """
<html>
  <body>
    <p>Dear {full_name},</p>
    
    <p>We are pleased to share your certificate for the Whatever Course name is. Please find the attached document.</p>

    <p><strong>Instructions:</strong></p>
    <p>We encourage you to share your achievement with your professional network by uploading your certificate on LinkedIn. Here's how:</p>
    <ol>
      <li>Log in to your <a href="https://www.linkedin.com" target="_blank">LinkedIn account</a>.</li>
      <li>Go to your profile and click on "Add Profile Section" &gt; "Accomplishments" &gt; "Certifications".</li>
      <li>Enter the following details:</li>
      <ul>
        <li><strong>Certification Name:</strong> Whatever Course name is</li>
        <li><strong>Issuing Organization:</strong> Company Name</li>
        <li><strong>Credential ID:</strong> Your Certificate No.</li>
      </ul>
      <li>Upload your certificate PDF (attached to this email).</li>
    </ol>

    <p>Don't forget to tag us in your LinkedIn post:</p>
    <ul>
      <li><a href="https://www.linkedin.com/school/nielitindia/" target="_blank">NIELIT India</a>, <a href="https://in.linkedin.com/in/lovnishverma" target="_blank"> Lovnish Verma</a></li>
    </ul>
    
    <p>We look forward to seeing your post and celebrating your success!</p>
    
    <p>Best regards,<br/>
    NIELIT Chandigarh</p>
  </body>
</html>
"""

# Email validation regex
email_regex = r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'

# Set up logging
logging.basicConfig(filename="email_sending.log", level=logging.INFO, 
                    format="%(asctime)s - %(levelname)s - %(message)s")

# Function to validate email format
def is_valid_email(email):
    return re.match(email_regex, email)

# Read Excel file
data = pd.read_excel(config["excel_file"])

# Required columns check
required_columns = {'full_name', 'email', 'cert_no'}
if not required_columns.issubset(data.columns):
    missing_cols = required_columns - set(data.columns)
    raise ValueError(f"Missing required columns in Excel file: {missing_cols}")

total_emails = len(data)
emails_sent = 0

# Initialize progress bar
with tqdm(total=total_emails, desc="Sending Emails", unit="email") as pbar:
    for _, row in data.iterrows():
        full_name = row['full_name']
        to_email = row['email']
        cert_no = row['cert_no']

        # Validate email
        if not is_valid_email(to_email):
            logging.error(f"Invalid email address for {full_name}. Email: {to_email}")
            continue

        # File path for certificate
        certificate_file = os.path.join(config["attachments_folder"], f"{cert_no}.pdf")

        try:
            with open(certificate_file, 'rb') as cert_file:
                certificate_data = base64.b64encode(cert_file.read()).decode()

                # Prepare Mailjet email payload
                data = {
                    'Messages': [
                        {
                            "From": {
                                "Email": config["from_email"],
                                "Name": "Enter your company name"  #replace it 
                            },
                            "To": [
                                {
                                    "Email": to_email,
                                    "Name": full_name
                                }
                            ],
                            "Subject": subject_template,
                            "HTMLPart": body_template.format(full_name=full_name, cert_no=cert_no),
                            "Attachments": [
                                {
                                    "ContentType": "application/pdf",
                                    "Filename": f"{cert_no}.pdf",
                                    "Base64Content": certificate_data
                                }
                            ]
                        }
                    ]
                }

                # Send email
                result = mailjet.send.create(data=data)

                if result.status_code == 200:
                    logging.info(f"Email sent to {to_email}")
                    emails_sent += 1
                else:
                    logging.error(f"Failed to send email to {to_email}. Error: {result.text}")
        except FileNotFoundError:
            logging.warning(f"Certificate file not found for {full_name} ({cert_no}). Skipping.")
        except Exception as e:
            logging.error(f"Error processing {full_name} ({cert_no}): {str(e)}")

        pbar.update(1)

# Summary log
summary_message = f"Total emails: {total_emails}\nSent: {emails_sent}\nFailed: {total_emails - emails_sent}"
logging.info(summary_message)

# Send summary email
summary_data = {
    'Messages': [
        {
            "From": {
                "Email": config["from_email"],
                "Name": "compant name here" #replace it 
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
