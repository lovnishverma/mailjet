**Certificate Email Sender - Mailjet GUI v1.4**, including instructions for running the Python version and using the standalone Windows `.exe` created with PyInstaller:

---

```markdown
# 🎓 Certificate Email Sender - Mailjet GUI (v1.4)

A user-friendly desktop application built using **Python & Tkinter** for sending personalized certificate emails in bulk using the **Mailjet API**. It supports Excel and CSV data formats, attaches PDF certificates, tracks email delivery progress, and sends a summary report upon completion.

---

## ✨ Features

- ✅ Intuitive graphical interface (built with Tkinter)
- 📂 Supports Excel (`.xlsx`) and CSV (`.csv`) data files
- 📎 Attaches personalized PDF certificates by `cert_no`
- 📈 Real-time logging and progress bar
- 🔐 Secure API key/secret entry with hidden input
- 💾 Auto-saves Mailjet config in `config.json`
- 📬 Sends a final summary email to the sender
- 📊 Email validation to avoid sending to invalid addresses
- 🛑 Option to skip certificate attachment
- 🧪 Built-in error detection for missing files or bad entries
- 🪄 Windows `.exe` version available (no Python required)

---

## 📁 Project Structure

```

certificate\_email\_sender/

├── mailjet.py            # main application code

├── config.json           # Auto-generated config (Mailjet settings)

├── README.md             # Documentation

├── requirements.txt      # Python dependencies

├── dist/

│   └── mailjet.exe          # Standalone executable (Windows)

├── build/

└── mailjet.spec             # PyInstaller build file

````

---

## 🖥️ How to Use (Python Version)

### 🔧 1. Clone this repository

```bash
git clone https://github.com/yourusername/certificate-email-sender.git
cd certificate-email-sender
````

### 🐍 2. (Optional) Create a virtual environment

```bash
python -m venv venv
# On Windows
venv\Scripts\activate
# On Linux/Mac
source venv/bin/activate
```

### 📦 3. Install dependencies

```bash
pip install -r requirements.txt
```

### 🚀 4. Launch the app

```bash
python mailjet.py
```

---

## 📌 Requirements

* Python 3.7+
* A Mailjet account (for API Key & Secret)
* Excel or CSV file containing:

  * `full_name`
  * `email`
  * `cert_no`
* PDF certificates in a folder named exactly as `<cert_no>.pdf`

---

## 📊 Sample Data File (CSV or Excel)

```csv
full_name,email,cert_no
Lovnish Verma,princelv84@gmail.com,CERT101
Prince,prince@gmail.com,CERT102
```

Each certificate should be named as `CERT101.pdf`, `CERT102.pdf`, etc.

---

## 📤 Summary Email

Once the email job completes, a summary email is sent to the sender’s Mailjet account with:

* Total Records
* Emails Sent
* Failed Deliveries (if any)

---

## 🔐 Configuration and Security

* API Key and Secret are securely entered and saved locally to `config.json`
* Inputs are masked using asterisks for safety
* Email data and PDF files are not uploaded anywhere except Mailjet during email delivery

---

## 📦 Using the Windows `.exe` Version

> No Python installation needed!

### ✅ Steps:

1. Download `mailjet.exe` from the [Releases](https://github.com/lovnishverma/mailjet/releases)
2. Double-click to open the GUI
3. Fill in all fields and browse for your data file and certificate folder
4. Click **Start Sending Emails**
5. Wait for the log and summary notification

---

## 🧰 Building the Executable Yourself (Optional)

You can generate your own `.exe` using PyInstaller:

### 1. Install PyInstaller

```bash
pip install pyinstaller
```

### 2. Build the executable

```bash
pyinstaller --noconfirm --onefile --windowed mailjet.py
```

The final `mailjet.exe` will be in the `dist/` directory.

---

## 🧪 Troubleshooting

| Issue                    | Cause / Fix                                        |
| ------------------------ | -------------------------------------------------- |
| ❌ Invalid Email          | Email doesn't match standard format                |
| ⚠️ Certificate not found | `cert_no.pdf` not found in selected folder         |
| ❌ API Error              | Invalid Mailjet credentials or daily limit reached |
| 🛑 GUI not launching     | Missing dependencies or bad Python install         |

---

## 📸 Application Screenshot

![App Screenshot](https://github.com/user-attachments/assets/74df951f-36ea-4d9e-a66f-b1d069279189)

---

## 📋 `requirements.txt`

```
mailjet-rest
pandas
tqdm
openpyxl
```

> ✅ `openpyxl` is required for reading Excel files (`.xlsx`)

---

## 🏷️ Version

**v1.4** — GUI version with:

* Summary email
* Mailjet error handling
* Config persistence
* Executable support

---

## 📜 License

This project is licensed under the MIT License.

---

## 🙌 Acknowledgements

Built with ❤️ by [Lovnish Verma](https://github.com/lovnishverma) using:

* [Mailjet API](https://www.mailjet.com/)
* [Tkinter GUI](https://docs.python.org/3/library/tkinter.html)
* [Pandas](https://pandas.pydata.org/)
* [PyInstaller](https://pyinstaller.org/)

```

---
