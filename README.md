**Bulk Email Sender Using Mailjet - Version 1.3**:

---

````markdown
# 📧 Bulk Email Sender Using Mailjet (v1.3)

Automate sending personalized emails with optional certificate attachments using [Mailjet's API](https://www.mailjet.com/). This version adds support for:
- ✅ CSV or Excel (`.xlsx`) input
- ✅ Optional attachments (configurable)
- ✅ HTML template with personalization
- ✅ Logging & progress bar
- ✅ Summary report to sender

---

## 🚀 Features

- 📄 Supports **CSV** and **Excel** input files
- 📎 Automatically attaches a certificate PDF (based on Cert No.)
- 🛑 Option to **disable attachments** using a single config flag
- 📬 Sends a **summary email** after the batch is processed
- 🔐 Secure configuration via Python `dict`
- ✅ Skips invalid emails automatically
- 📊 Progress bar using `tqdm`
- 📝 Error and delivery logs saved to `email_sending.log`

---

## 🛠️ Requirements

Install dependencies with:

```bash
pip install pandas tqdm mailjet_rest openpyxl
````

---

## 📁 File Structure

```
mailjet/
├── mailjet.py               # Main script
├── TEST.xlsx               # Input Excel/CSV file (see format below)
├── certificates/           # Folder containing certificate PDFs
└── email_sending.log       # Generated log file after execution
```

---

## 📥 Input File Format (Excel or CSV)

Your file should contain the following **three columns**:

| full\_name | email                                                   | cert\_no |
| ---------- | ------------------------------------------------------- | -------- |
| John Doe   | [john.doe@example.com](mailto:john.doe@example.com)     | CERT-001 |
| Jane Smith | [jane.smith@example.com](mailto:jane.smith@example.com) | CERT-002 |

> The certificate file should be named as `CERT-001.pdf`, `CERT-002.pdf`, etc., and placed in the `certificates/` folder.

---

## ⚙️ Configuration

Update this Python dictionary at the top of `mailjet.py`:

```python
config = {
    "from_email": "youremail@example.com",
    "file_type": "excel",  # or 'csv'
    "file_path": "TEST.xlsx",
    "attachments_folder": "certificates",
    "mailjet_api_key": "YOUR_MAILJET_API_KEY",
    "mailjet_api_secret": "YOUR_MAILJET_SECRET",
    "disable_attachments": False  # Set True to skip sending certificates
}
```

---

## 🧪 Running the Script

```bash
python mailjet.py
```

You will see a live progress bar and logs in `email_sending.log`.

---

## 📤 Output

* ✅ Individual emails sent to each recipient
* 📎 Certificate attached (if enabled and found)
* ❗ Logs error if email is invalid or PDF missing
* 📨 Summary email sent to sender with:

  * Total emails processed
  * Emails sent
  * Emails failed

---

## 🛑 Notes & Limitations

* Mailjet **free tier** allows:

  * ✅ 200 emails/day
  * ✅ 6,000 emails/month
* Make sure to **verify your sender email** in your Mailjet account
* Certificate files must be named exactly as `{cert_no}.pdf`

---

## 🤝 Contributing

Suggestions, bug reports, and pull requests are welcome!

You can contribute via:

* Branch: [`mailjetver1.3`](https://github.com/lovnishverma/mailjet/tree/mailjetver1.3)
* Issue: [Suggestion #1](https://github.com/lovnishverma/mailjet/issues/1)

---

## 📧 Sample Email Screenshot

> ![WhatsApp Image 2025-05-15 at 15 15 32_b0c621b2](https://github.com/user-attachments/assets/efee16e0-4355-4b57-a989-6200da5285f2)


---

## 🙋‍♂️ Maintainer

**Lovnish Verma**
🔗 [LinkedIn](https://in.linkedin.com/in/lovnishverma)
📬 [GitHub](https://github.com/lovnishverma)

---

## 📜 License

This project is open-source under the [MIT License](LICENSE).

````

---
