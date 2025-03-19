# Bulk Email Sender Using Mailjet Version 1.2

This Python script automates the process of sending bulk emails with attachments (certificates) using **Mailjet's API**. The script reads recipient details from an Excel file, validates emails, and sends personalized emails with attachments.

## Features
- Reads recipient details (name, email, certificate number) from an Excel file.
- Sends personalized emails using Mailjet API.
- Attaches PDF certificates to each email.
- Logs success and failure messages.
- Provides a summary email report at the end.

## Requirements
- Python 3.x
- Mailjet account (API Key and Secret Key)
- Required Python Libraries:
  - `pandas`
  - `os`
  - `logging`
  - `json`
  - `base64`
  - `tqdm`
  - `re`
  - `mailjet_rest`

## Installation

1. **Clone the Repository** (if applicable):
   ```sh
   git clone https://github.com/lovnishverma/mailjet.git
   cd mailjet
   ```

2. **Install Required Dependencies:**
   ```sh
   pip install pandas tqdm mailjet-rest
   ```

3. **Update Configuration in the Script:**
   - Edit the `config` dictionary inside the script to add:
     - Your Mailjet API Key and Secret
     - Sender Email (registered with Mailjet)
     - Path to the Excel file
     - Folder containing certificates

## Usage

1. **Prepare the Excel File:**
   - The file should have the following columns:
     - `full_name`: Recipient's full name
     - `email`: Recipient's email address
     - `cert_no`: Certificate number (used to locate the PDF file)
   
2. **Prepare the Certificate Folder:**
   - Store all PDF certificates inside a folder (e.g., `certificates/`).
   - Ensure each certificate is named as `{cert_no}.pdf` (matching `cert_no` in the Excel file).

3. **Run the Script:**
   ```sh
   python mailjet.py
   ```

## Email Template
The script sends emails with the following content:
- Personalized greeting (`Dear {full_name},`)
- Instructions on sharing the certificate on LinkedIn
- Attached PDF certificate
- Signature from `your company`

## Logs and Summary
- The script logs each email's status in `email_sending.log`.
- After execution, a **summary email** is sent to the administrator with total sent/failed emails.

## Troubleshooting
### Error: `FileNotFoundError: Certificate file not found`
- Ensure the `cert_no` matches the file name in the `certificates` folder.

### Error: `Invalid email address`
- The script uses a regex pattern to validate emails before sending. Ensure the email format is correct.

### Error: `Mailjet authentication failed`
- Double-check your Mailjet API Key and Secret Key in the script.

## License
This project is licensed under the MIT License.

