# Certificate Email Sender - Mailjet GUI v1.4

A professional desktop application for sending personalized certificate emails in bulk using the Mailjet API. Built with Python and Tkinter, featuring an intuitive GUI, real-time progress tracking, and comprehensive error handling.

## 🚀 Quick Start

**For Windows Users (No Python Required):**
1. Download `mailjet.exe` from [Releases](https://github.com/lovnishverma/certificate-email-sender/releases/tag/v1.4)
2. Double-click to launch the application
3. Follow the setup wizard in the GUI

**For Python Users:**
```bash
git clone https://github.com/lovnishverma/certificate-email-sender.git
cd certificate-email-sender
pip install -r requirements.txt
python mailjet.py
```

## 📋 Table of Contents

- [Features](#features)
- [Prerequisites](#prerequisites)
- [Installation](#installation)
  - [Python Installation](#python-installation)
  - [Windows Executable](#windows-executable)
- [Configuration](#configuration)
- [Usage](#usage)
- [Data Format Requirements](#data-format-requirements)
- [Troubleshooting](#troubleshooting)
- [Building from Source](#building-from-source)
- [Contributing](#contributing)
- [License](#license)

## ✨ Features

### Core Functionality
- **Bulk Email Sending**: Send personalized certificates to multiple recipients
- **Multi-format Support**: Compatible with Excel (`.xlsx`) and CSV (`.csv`) files
- **PDF Attachment**: Automatically attach certificates based on certificate numbers
- **Email Validation**: Built-in validation to prevent sending to invalid addresses

### User Experience
- **Intuitive GUI**: Clean, user-friendly interface built with Tkinter
- **Real-time Progress**: Live progress bar and detailed logging
- **Summary Reports**: Comprehensive email delivery summary sent to sender
- **Secure Configuration**: Encrypted API key storage with masked input fields

### Technical Features
- **Error Handling**: Robust error detection and recovery
- **Configuration Persistence**: Auto-saves settings in `config.json`
- **Cross-platform**: Runs on Windows, macOS, and Linux
- **Standalone Executable**: Windows `.exe` version requires no Python installation

## 📦 Prerequisites

### For Python Installation
- Python 3.7 or higher
- Active internet connection
- Mailjet account with API credentials

### For Windows Executable
- Windows 10 or later
- Active internet connection
- Mailjet account with API credentials

### Mailjet Account Setup
1. Create a free account at [Mailjet.com](https://www.mailjet.com/)
2. Navigate to Account Settings → REST API → API Key Management
3. Generate your API Key and Secret Key
4. Keep these credentials secure - you'll need them for the application

## 🛠️ Installation

### Python Installation

#### Step 1: Clone the Repository
```bash
git clone https://github.com/yourusername/certificate-email-sender.git
cd certificate-email-sender
```

#### Step 2: Create Virtual Environment (Recommended)
```bash
# Create virtual environment
python -m venv certificate_sender_env

# Activate virtual environment
# Windows:
certificate_sender_env\Scripts\activate
# macOS/Linux:
source certificate_sender_env/bin/activate
```

#### Step 3: Install Dependencies
```bash
pip install -r requirements.txt
```

#### Step 4: Launch Application
```bash
python mailjet.py
```

### Windows Executable

1. Download the latest `mailjet.exe` from the [Releases page](https://github.com/lovnishverma/mailjet/releases)
2. Save to your preferred location (no installation required)
3. Double-click `mailjet.exe` to launch
4. Windows may show a security warning - click "More info" → "Run anyway"

## ⚙️ Configuration

### First-Time Setup
1. Launch the application
2. Enter your Mailjet API credentials:
   - **API Key**: Your Mailjet public key
   - **API Secret**: Your Mailjet private key (masked for security)
   - **Sender Name**: Display name for outgoing emails
   - **Sender Email**: Must be verified in your Mailjet account

### Configuration File
Settings are automatically saved to `config.json` in the application directory:
```json
{
  "api_key": "your_api_key",
  "api_secret": "your_api_secret",
  "sender_name": "Your Organization",
  "sender_email": "certificates@yourorg.com"
}
```

**Security Note**: Keep `config.json` secure and never share it publicly.

## 📖 Usage

### Step-by-Step Process

#### 1. Prepare Your Data
Create an Excel or CSV file with the following columns:
- `full_name`: Recipient's complete name
- `email`: Recipient's email address
- `cert_no`: Unique certificate identifier

#### 2. Organize Certificate Files
- Place all PDF certificates in a single folder
- Name each certificate file exactly as: `{cert_no}.pdf`
- Example: If `cert_no` is "CERT101", the file should be named "CERT101.pdf"

#### 3. Configure Email Content
- **Subject Line**: Customize the email subject
- **Email Body**: Write your personalized message
- Use placeholders:
  - `{full_name}` - Replaced with recipient's name
  - `{cert_no}` - Replaced with certificate number

#### 4. Start Sending
1. Click "Browse" to select your data file
2. Click "Browse" to select your certificates folder
3. Verify all settings are correct
4. Click "Start Sending Emails"
5. Monitor progress in the log window

### Example Email Template
```
Subject: Your Certificate - Congratulations {full_name}!

Dear {full_name},

Congratulations on your achievement! Please find your certificate ({cert_no}) attached to this email.

Best regards,
The Certification Team
```

## 📊 Data Format Requirements

### Required Columns
Your data file must contain these exact column names:

| Column | Description | Example |
|--------|-------------|---------|
| `full_name` | Recipient's complete name | "John Smith" |
| `email` | Valid email address | "john@example.com" |
| `cert_no` | Unique certificate identifier | "CERT001" |

### Sample Data File

**CSV Format:**
```csv
full_name,email,cert_no
John Smith,john@example.com,CERT001
Jane Doe,jane@example.com,CERT002
Bob Johnson,bob@example.com,CERT003
```

**Excel Format:**
| full_name | email | cert_no |
|-----------|-------|---------|
| John Smith | john@example.com | CERT001 |
| Jane Doe | jane@example.com | CERT002 |
| Bob Johnson | bob@example.com | CERT003 |

### File Naming Convention
Certificate PDF files must be named exactly as the `cert_no` value:
- CERT001.pdf
- CERT002.pdf
- CERT003.pdf

## 🔍 Progress Tracking & Reporting

### Real-time Monitoring
- **Progress Bar**: Visual indication of completion percentage
- **Live Log**: Detailed status updates for each email
- **Error Reporting**: Immediate notification of any issues

### Summary Email
Upon completion, a summary report is automatically sent to the sender's email containing:
- Total records processed
- Number of emails sent successfully
- Number of failed deliveries
- List of failed recipients (if any)
- Processing time and timestamp

## 🛠️ Troubleshooting

### Common Issues and Solutions

| Issue | Possible Cause | Solution |
|-------|---------------|----------|
| **"Invalid email format"** | Email doesn't match standard format | Verify email addresses in your data file |
| **"Certificate not found"** | PDF file missing or incorrectly named | Ensure `{cert_no}.pdf` exists in certificates folder |
| **"API authentication failed"** | Incorrect Mailjet credentials | Verify API Key and Secret in Mailjet dashboard |
| **"Daily limit exceeded"** | Mailjet sending limit reached | Wait 24 hours or upgrade your Mailjet plan |
| **"Application won't start"** | Missing dependencies or corrupted installation | Reinstall Python dependencies or re-download executable |
| **"Permission denied"** | Insufficient file access rights | Run as administrator or check file permissions |

### Debug Mode
To enable detailed logging:
1. Open command prompt/terminal
2. Navigate to application directory
3. Run: `python mailjet.py --debug` (Python version only)

### Getting Help
If you encounter persistent issues:
1. Check the [Issues page](https://github.com/yourusername/certificate-email-sender/issues)
2. Create a new issue with:
   - Your operating system
   - Application version
   - Error message (if any)
   - Steps to reproduce the problem

## 🔧 Building from Source

### Creating Windows Executable

#### Prerequisites
```bash
pip install pyinstaller
```

#### Build Command
```bash
pyinstaller --noconfirm --onefile --windowed --name="Certificate_Email_Sender" mailjet.py
```

#### Advanced Build Options
```bash
pyinstaller --noconfirm --onefile --windowed \
  --add-data="config.json;." \
  --icon="icon.ico" \
  --name="Certificate_Email_Sender" \
  mailjet.py
```

The executable will be created in the `dist/` directory.

### Project Structure
```
certificate_email_sender/
├── mailjet.py              # Main application code
├── config.json             # Configuration file (auto-generated)
├── requirements.txt        # Python dependencies
├── README.md              # This documentation
├── LICENSE                # License file
├── .gitignore            # Git ignore rules
├── dist/                 # Built executables
│   └── Certificate_Email_Sender.exe
├── build/                # Build artifacts
└── mailjet.spec          # PyInstaller configuration
```

## 📦 Dependencies

```txt
mailjet-rest==1.3.4
pandas>=1.3.0
tqdm>=4.62.0
openpyxl>=3.0.9
```

## 🤝 Contributing

We welcome contributions! Please follow these steps:

1. **Fork the repository**
2. **Create a feature branch**: `git checkout -b feature/your-feature-name`
3. **Make your changes** and test thoroughly
4. **Commit your changes**: `git commit -am 'Add some feature'`
5. **Push to the branch**: `git push origin feature/your-feature-name`
6. **Submit a Pull Request**

### Development Guidelines
- Follow PEP 8 style guidelines
- Add comments for complex logic
- Test on multiple platforms when possible
- Update documentation for new features

## 📄 License

This project is licensed under the MIT License. See the [LICENSE](LICENSE) file for details.

## 🙏 Acknowledgements

### Built With
- **[Mailjet API](https://www.mailjet.com/)** - Email delivery service
- **[Python Tkinter](https://docs.python.org/3/library/tkinter.html)** - GUI framework
- **[Pandas](https://pandas.pydata.org/)** - Data manipulation
- **[PyInstaller](https://pyinstaller.org/)** - Executable creation

### Special Thanks
- Mailjet team for providing reliable email API
- Python community for excellent libraries
- Beta testers for valuable feedback

---

## 📞 Support

- **Documentation**: This README and inline code comments
- **Issues**: [GitHub Issues](https://github.com/yourusername/certificate-email-sender/issues)
- **Discussions**: [GitHub Discussions](https://github.com/yourusername/certificate-email-sender/discussions)

---

<div align="center">

**Built with ❤️ by [Lovnish Verma](https://github.com/lovnishverma)**

[⭐ Star this project](https://github.com/yourusername/certificate-email-sender) | [🐛 Report Bug](https://github.com/yourusername/certificate-email-sender/issues) | [💡 Request Feature](https://github.com/yourusername/certificate-email-sender/issues)

</div>
