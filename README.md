**Certificate Email Sender - Mailjet GUI version 1.4** app:

---

```markdown
# Certificate Email Sender - Mailjet GUI 🎓📧

A user-friendly desktop application built using **Python & Tkinter** for sending personalized certificate emails in bulk using the **Mailjet API**. It supports both Excel and CSV data formats, attaches PDF certificates, and provides real-time logging and summary.

---

## 🚀 Features

- ✅ Simple and interactive GUI
- 📂 Supports Excel (`.xlsx`) and CSV (`.csv`) data files
- 📎 Sends certificates (PDF attachments) using Mailjet
- 📈 Progress bar and real-time logs
- 💾 Auto-saves API keys and configuration
- 📬 Summary email sent to admin after completion
- 🔒 Hide API secrets with password-like input
- 🔍 Email validation to avoid sending to invalid addresses

---

## 📁 Project Structure

```

certificate\_email\_sender/
│
├── main.py                  # Main application code (MailSenderApp)
├── config.json              # Auto-generated configuration file
├── README.md                # This documentation
└── requirements.txt         # Python dependencies

````

---

## 🛠️ Installation

1. **Clone the repository**
   ```bash
   git clone https://github.com/yourusername/certificate-email-sender.git
   cd certificate-email-sender
````

2. **Create a virtual environment (optional but recommended)**

   ```bash
   python -m venv venv
   source venv/bin/activate  # On Windows: venv\Scripts\activate
   ```

3. **Install required libraries**

   ```bash
   pip install -r requirements.txt
   ```

---

## 📌 Requirements

* Python 3.7+
* A **Mailjet** account (for API Key & Secret)
* Excel or CSV file with the following **columns**:

  * `full_name`
  * `email`
  * `cert_no`
* PDF certificate files named as `<cert_no>.pdf`

---

## 📦 Example Usage

1. **Launch the app**

   ```bash
   python main.py
   ```

2. **Fill in the form**:

   * From Email
   * Mailjet API Key & Secret
   * Choose file type: Excel or CSV
   * Browse and select the data file
   * Browse and select the folder with PDF certificates
   * Optionally disable attachments

3. **Click** `Start Sending Emails`

4. Monitor progress and logs in the GUI.

---

## 📤 Summary Email

After sending all emails, a summary report is emailed to the configured "From Email" address, showing:

* Total records
* Sent count
* Failed count

---

## 📋 Configuration File

All settings are auto-saved in `config.json` on close or email send, so you don’t need to enter them every time.

---

## 🧪 Sample Data File Format

```csv
full_name,email,cert_no
John Doe,john@example.com,CERT001
Jane Smith,jane@example.com,CERT002
```

---

## 🔐 Security Notes

* API credentials are stored in `config.json`. Keep this file safe.
* Credentials are not sent anywhere other than Mailjet's API during the operation.

---

## 🧰 Troubleshooting

* ❌ Invalid Email: Shown when email does not match the pattern.
* ⚠️ Certificate not found: Shown when a matching PDF is missing.
* ❌ API Error: May indicate wrong Mailjet credentials or issues with the recipient email.

---

## 📜 License

MIT License

---

## 🤝 Acknowledgements

Built with ❤️ by [Lovnish Verma](https://github.com/lovnishverma) using:

* [Mailjet](https://www.mailjet.com/)
* [Tkinter](https://docs.python.org/3/library/tkinter.html)
* [Pandas](https://pandas.pydata.org/)

````

---

## ✅ Add `requirements.txt`

You should also create a `requirements.txt` file:

```txt
mailjet-rest
pandas
tqdm
openpyxl
````

> Note: `openpyxl` is required for reading `.xlsx` files.

![image](https://github.com/user-attachments/assets/74df951f-36ea-4d9e-a66f-b1d069279189)


This as an executable made using with `pyinstaller`.
