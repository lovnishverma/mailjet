## Important Notice
[Check branch 2 for improved version ](https://github.com/lovnishverma/mailjet/tree/mailjetver1.2)


# Bulk Email Sender Using Mailjet

This Python script automates the process of sending bulk emails with attachments (certificates) using **Mailjet's API**. With Mailjet's free tier, you can send up to **200 emails per day (6,000 emails per month) instantly** without any cost. The script reads recipient details from an Excel file, validates emails, and sends personalized emails with attachments.

## Features
- **Send Bulk Emails for Free**: Mailjet allows sending up to **200 emails/day, 6,000 emails/month** on the free plan.
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


## License
This project is licensed under the MIT License.

