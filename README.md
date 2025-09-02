# Bulk Email Sender Using Mailjet API v1.3

A powerful Python script for automating personalized bulk email campaigns with optional PDF certificate attachments. Built with [Mailjet's REST API](https://www.mailjet.com/) for reliable email delivery and comprehensive reporting.

## 🚀 Quick Start

```bash
# Clone the repository
git clone https://github.com/lovnishverma/mailjet.git
cd mailjet

# Install dependencies
pip install -r requirements.txt

# Configure your settings in mailjet.py
# Run the application
python mailjet.py
```

## 📋 Table of Contents

- [Features](#features)
- [Prerequisites](#prerequisites)
- [Installation](#installation)
- [Configuration](#configuration)
- [Data Format](#data-format)
- [Usage](#usage)
- [Email Templates](#email-templates)
- [Logging & Monitoring](#logging--monitoring)
- [Troubleshooting](#troubleshooting)
- [API Limits & Pricing](#api-limits--pricing)
- [Contributing](#contributing)
- [License](#license)

## ✨ Features

### Core Functionality
- **Multi-format Support**: Process CSV and Excel (`.xlsx`) input files
- **Smart Attachments**: Automatically attach PDF certificates based on certificate numbers
- **Flexible Configuration**: Toggle attachments on/off with a single flag
- **Email Validation**: Automatic validation and skipping of invalid email addresses
- **Batch Processing**: Efficiently handle large recipient lists

### Monitoring & Reporting
- **Real-time Progress**: Live progress bar using `tqdm`
- **Comprehensive Logging**: Detailed logs saved to `email_sending.log`
- **Summary Reports**: Automatic delivery summary sent to sender
- **Error Tracking**: Track failed deliveries with detailed error messages

### Customization & Security
- **HTML Templates**: Rich HTML email templates with personalization
- **Secure Configuration**: API credentials stored securely in configuration
- **Personalization**: Dynamic content replacement using recipient data
- **Error Recovery**: Graceful handling of missing files and API errors

## 📦 Prerequisites

### System Requirements
- **Python**: Version 3.7 or higher
- **Internet Connection**: Required for Mailjet API access
- **Storage**: Sufficient space for logs and temporary files

### Mailjet Account Setup
1. **Create Account**: Sign up at [Mailjet.com](https://www.mailjet.com/)
2. **Verify Domain**: Add and verify your sending domain
3. **Get API Keys**: 
   - Navigate to Account Settings → REST API
   - Generate API Key and Secret Key
   - Keep credentials secure

### Required Python Packages
```txt
pandas>=1.3.0
tqdm>=4.62.0
mailjet_rest>=1.3.4
openpyxl>=3.0.9
```

## 🛠️ Installation

### Method 1: Using pip (Recommended)
```bash
# Install all dependencies at once
pip install pandas tqdm mailjet_rest openpyxl
```

### Method 2: Using requirements.txt
```bash
# Create requirements.txt file
echo "pandas>=1.3.0" > requirements.txt
echo "tqdm>=4.62.0" >> requirements.txt
echo "mailjet_rest>=1.3.4" >> requirements.txt
echo "openpyxl>=3.0.9" >> requirements.txt

# Install from requirements file
pip install -r requirements.txt
```

### Method 3: Virtual Environment (Recommended for Production)
```bash
# Create virtual environment
python -m venv mailjet_env

# Activate virtual environment
# Windows:
mailjet_env\Scripts\activate
# macOS/Linux:
source mailjet_env/bin/activate

# Install dependencies
pip install pandas tqdm mailjet_rest openpyxl
```

## ⚙️ Configuration

### Basic Configuration
Edit the configuration dictionary at the top of `mailjet.py`:

```python
config = {
    # Email Settings
    "from_email": "certificates@yourcompany.com",
    "from_name": "Your Organization Name",
    
    # File Settings  
    "file_type": "excel",  # Options: "excel" or "csv"
    "file_path": "recipients.xlsx",
    "attachments_folder": "certificates",
    
    # Mailjet API Credentials
    "mailjet_api_key": "YOUR_MAILJET_API_KEY",
    "mailjet_api_secret": "YOUR_MAILJET_SECRET",
    
    # Feature Toggles
    "disable_attachments": False,  # Set True to skip PDF attachments
    "send_summary": True,         # Set False to disable summary email
    
    # Email Content
    "subject_template": "Your Certificate - {full_name}",
    "html_template": """
    <!DOCTYPE html>
    <html>
    <body style="font-family: Arial, sans-serif; line-height: 1.6; color: #333;">
        <div style="max-width: 600px; margin: 0 auto; padding: 20px;">
            <h2 style="color: #2c3e50;">Congratulations {full_name}!</h2>
            <p>We are pleased to inform you that your certificate has been processed.</p>
            <p><strong>Certificate Number:</strong> {cert_no}</p>
            <p>Please find your certificate attached to this email.</p>
            <hr style="border: 1px solid #eee; margin: 20px 0;">
            <p style="font-size: 12px; color: #666;">
                This is an automated message. Please do not reply to this email.
            </p>
        </div>
    </body>
    </html>
    """
}
```

### Advanced Configuration Options

```python
# Additional configuration for advanced users
advanced_config = {
    # Retry Settings
    "max_retries": 3,
    "retry_delay": 2,  # seconds
    
    # Logging Settings
    "log_level": "INFO",  # DEBUG, INFO, WARNING, ERROR
    "log_file": "email_sending.log",
    
    # Rate Limiting
    "emails_per_batch": 50,
    "batch_delay": 1,  # seconds between batches
    
    # Validation Settings
    "strict_email_validation": True,
    "skip_duplicates": True
}
```

## 📊 Data Format Requirements

### Required Columns
Your input file must contain these exact column headers:

| Column Name | Description | Example | Required |
|-------------|-------------|---------|----------|
| `full_name` | Recipient's complete name | "John Smith" | ✅ Yes |
| `email` | Valid email address | "john@example.com" | ✅ Yes |
| `cert_no` | Unique certificate identifier | "CERT-001" | ✅ Yes |

### Sample Data Files

#### Excel Format (`recipients.xlsx`)
| full_name | email | cert_no |
|-----------|-------|---------|
| Lovnish Verma | technicalboyprince@gmail.com | CERT-001 |
| Prince Verma | princelv84@gmail.com | CERT-002 |
| John Smith | john.smith@example.com | CERT-003 |

#### CSV Format (`recipients.csv`)
```csv
full_name,email,cert_no
Lovnish Verma,technicalboyprince@gmail.com,CERT-001
Prince Verma,princelv84@gmail.com,CERT-002
John Smith,john.smith@example.com,CERT-003
```

### Certificate File Organization
```
certificates/
├── CERT-001.pdf
├── CERT-002.pdf
├── CERT-003.pdf
└── ...
```

**Important**: Certificate filenames must match exactly with the `cert_no` column values.

## 📁 Project Structure

```
mailjet/
├── mailjet.py              # Main application script
├── requirements.txt        # Python dependencies
├── recipients.xlsx         # Input data file (Excel format)
├── recipients.csv          # Input data file (CSV format)  
├── certificates/           # Directory containing PDF certificates
│   ├── CERT-001.pdf
│   ├── CERT-002.pdf
│   └── ...
├── logs/                   # Generated log files
│   ├── email_sending.log
│   └── error.log
├── templates/              # Email templates (optional)
│   └── certificate_template.html
├── README.md              # This documentation
└── LICENSE                # License file
```

## 🚀 Usage

### Step 1: Prepare Your Data
1. Create your recipient list in Excel or CSV format
2. Ensure all required columns are present
3. Validate email addresses for correctness

### Step 2: Organize Certificates
1. Place all PDF certificates in the `certificates/` folder
2. Name files exactly as: `{cert_no}.pdf`
3. Verify file permissions are readable

### Step 3: Configure Settings
1. Open `mailjet.py` in your preferred editor
2. Update the configuration dictionary
3. Add your Mailjet API credentials
4. Customize email templates as needed

### Step 4: Execute Script
```bash
# Basic execution
python mailjet.py

# With verbose output
python mailjet.py --verbose

# Dry run (test without sending)
python mailjet.py --dry-run
```

### Expected Output
```
📧 Bulk Email Sender v1.3 Starting...
📂 Loading data from recipients.xlsx...
✅ Loaded 150 recipients successfully
🔍 Validating email addresses...
✅ 148 valid emails found, 2 invalid emails skipped
📎 Checking certificate files...
⚠️  2 certificates missing, will skip attachments for those
🚀 Starting email delivery...

Sending emails: 100%|████████████| 148/148 [02:35<00:00,  1.05it/s]

📊 Summary:
   • Total recipients: 150
   • Emails sent: 146
   • Failed deliveries: 2
   • Missing certificates: 2
   • Processing time: 2m 35s

✅ Summary report sent to certificates@yourcompany.com
📝 Detailed logs saved to email_sending.log
```

## 📧 Email Templates

### Basic HTML Template
```html
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Certificate Delivery</title>
</head>
<body style="font-family: Arial, sans-serif; margin: 0; padding: 20px; background-color: #f4f4f4;">
    <div style="max-width: 600px; margin: 0 auto; background-color: white; padding: 30px; border-radius: 10px; box-shadow: 0 0 10px rgba(0,0,0,0.1);">
        
        <!-- Header -->
        <div style="text-align: center; margin-bottom: 30px;">
            <h1 style="color: #2c3e50; margin-bottom: 10px;">Certificate Delivery</h1>
            <hr style="border: 2px solid #3498db; width: 100px; margin: 0 auto;">
        </div>
        
        <!-- Main Content -->
        <div style="margin-bottom: 30px;">
            <h2 style="color: #2c3e50;">Dear {full_name},</h2>
            <p style="font-size: 16px; line-height: 1.6; color: #555;">
                Congratulations on your outstanding achievement! We are delighted to present you with your official certificate.
            </p>
            
            <div style="background-color: #ecf0f1; padding: 15px; border-radius: 5px; margin: 20px 0;">
                <p style="margin: 0; font-weight: bold; color: #2c3e50;">Certificate Details:</p>
                <p style="margin: 5px 0 0 0; color: #555;">Number: <strong>{cert_no}</strong></p>
            </div>
            
            <p style="font-size: 16px; line-height: 1.6; color: #555;">
                Please find your certificate attached as a PDF file. Keep this document safe as it serves as official proof of your accomplishment.
            </p>
        </div>
        
        <!-- Footer -->
        <div style="border-top: 1px solid #eee; padding-top: 20px; text-align: center;">
            <p style="font-size: 12px; color: #999; margin: 0;">
                This is an automated message from our certification system.<br>
                Please do not reply to this email address.
            </p>
        </div>
    </div>
</body>
</html>
```

### Template Variables
Available placeholders for personalization:
- `{full_name}` - Recipient's full name
- `{cert_no}` - Certificate number
- `{email}` - Recipient's email address
- `{date}` - Current date (auto-generated)
- `{time}` - Current time (auto-generated)

## 📊 Logging & Monitoring

### Log Levels
- **DEBUG**: Detailed information for troubleshooting
- **INFO**: General information about program execution
- **WARNING**: Something unexpected happened but program continues
- **ERROR**: Serious problem that prevented function execution

### Log File Format
```
2025-09-02 10:15:23,456 - INFO - Starting bulk email sender v1.3
2025-09-02 10:15:23,789 - INFO - Loaded 150 recipients from recipients.xlsx
2025-09-02 10:15:24,123 - WARNING - Invalid email skipped: invalid-email@
2025-09-02 10:15:24,456 - INFO - Email sent successfully to john@example.com
2025-09-02 10:15:24,789 - ERROR - Certificate not found: CERT-999.pdf
2025-09-02 10:15:25,123 - INFO - Batch processing completed
```

### Progress Monitoring
The script provides real-time progress updates:
- Progress bar showing completion percentage
- Current recipient being processed  
- Success/failure indicators
- Estimated time remaining
- Processing rate (emails per second)

## 🛠️ Troubleshooting

### Common Issues and Solutions

| Issue | Symptoms | Solution |
|-------|----------|----------|
| **Authentication Failed** | `401 Unauthorized` error | Verify API Key and Secret are correct |
| **Invalid Email Address** | Emails skipped with validation error | Check email format in your data file |
| **Certificate Not Found** | Missing attachment warnings | Ensure PDF files exist and are named correctly |
| **Rate Limit Exceeded** | `429 Too Many Requests` error | Reduce batch size or add delays between batches |
| **File Not Found** | Cannot load input file | Verify file path and permissions |
| **Memory Error** | Script crashes with large files | Process data in smaller batches |

### Debug Mode
Enable detailed logging for troubleshooting:

```python
import logging
logging.basicConfig(level=logging.DEBUG)
```

### Testing Configuration
Use dry-run mode to test without sending emails:

```python
config["dry_run"] = True  # Add this to your config
```

### Validation Checklist
Before running the script, verify:
- [ ] Mailjet API credentials are valid
- [ ] Sender email is verified in Mailjet
- [ ] Input file exists and has correct format
- [ ] Certificate folder exists and contains PDF files
- [ ] Python dependencies are installed
- [ ] Internet connection is stable

## 📈 API Limits & Pricing

### Mailjet Free Tier Limitations
- **Daily Limit**: 200 emails per day
- **Monthly Limit**: 6,000 emails per month
- **Rate Limit**: 120 emails per minute
- **Attachment Size**: 15MB per email
- **Recipients**: No limit on number of recipients

### Recommended Batch Sizes
| Plan Type | Batch Size | Delay Between Batches |
|-----------|------------|----------------------|
| Free | 10-20 emails | 5-10 seconds |
| Essential | 50-100 emails | 2-5 seconds |
| Premium | 100-200 emails | 1-2 seconds |

### Cost Optimization Tips
1. **Validate Emails**: Remove invalid addresses before sending
2. **Remove Duplicates**: Avoid sending multiple emails to same recipient
3. **Monitor Bounces**: Track and remove bounced email addresses
4. **Use Templates**: Reduce API calls with template-based sending

## 🧪 Testing

### Unit Tests
```bash
# Run basic tests
python -m pytest tests/

# Run with coverage
python -m pytest --cov=mailjet tests/
```

### Integration Tests
```bash
# Test with small dataset
python mailjet.py --test-mode --max-emails=5
```

### Sample Test Data
Create a test file `test_recipients.csv`:
```csv
full_name,email,cert_no
Test User 1,test1@example.com,TEST-001
Test User 2,test2@example.com,TEST-002
```

## 🤝 Contributing

We welcome contributions! Here's how to get started:

### Development Setup
```bash
# Fork the repository
git clone https://github.com/yourusername/mailjet.git
cd mailjet

# Create development branch
git checkout -b feature/your-feature-name

# Install development dependencies
pip install -r requirements-dev.txt

# Make your changes and test
python -m pytest

# Commit and push
git commit -am "Add your feature"
git push origin feature/your-feature-name
```

### Contribution Guidelines
- **Code Style**: Follow PEP 8 guidelines
- **Documentation**: Update README for new features
- **Testing**: Add tests for new functionality
- **Commit Messages**: Use clear, descriptive messages

### Reporting Issues
When reporting bugs, please include:
- Python version
- Operating system
- Error message (full traceback)
- Steps to reproduce
- Sample data (anonymized)

## 🔒 Security Considerations

### API Key Security
- Never commit API keys to version control
- Use environment variables for production
- Rotate keys regularly
- Restrict API key permissions in Mailjet dashboard

### Data Privacy
- Validate email addresses before processing
- Remove sensitive data from logs
- Follow GDPR/privacy regulations
- Secure storage of recipient data

### Best Practices
```python
import os
from dotenv import load_dotenv

load_dotenv()

config = {
    "mailjet_api_key": os.getenv("MAILJET_API_KEY"),
    "mailjet_api_secret": os.getenv("MAILJET_API_SECRET"),
    # ... other config
}
```

## 📧 Sample Email Preview

### Rendered Email Example
![Certificate Email Example](https://github.com/user-attachments/assets/efee16e0-4355-4b57-a989-6200da5285f2)

### Email Headers
```
From: certificates@yourcompany.com
To: recipient@example.com
Subject: Your Certificate - John Smith
Content-Type: text/html; charset=UTF-8
Attachments: CERT-001.pdf (245 KB)
```

## 🔄 Version History

### v1.3 (Current)
- ✅ Added Excel support
- ✅ Configurable attachment toggle
- ✅ Enhanced logging system
- ✅ Summary email reports
- ✅ Better error handling

### v1.2
- ✅ CSV file support
- ✅ Basic attachment functionality
- ✅ Simple progress tracking

### v1.1
- ✅ Initial release
- ✅ Basic email sending
- ✅ Mailjet integration

## 📚 Additional Resources

### Documentation Links
- [Mailjet API Documentation](https://dev.mailjet.com/)
- [Python Mailjet Client](https://github.com/mailjet/mailjet-apiv3-python)
- [Pandas Documentation](https://pandas.pydata.org/docs/)

### Tutorials
- [Setting up Mailjet Account](https://www.mailjet.com/guides/getting-started/)
- [Email Marketing Best Practices](https://www.mailjet.com/blog/news/email-marketing-best-practices/)

## 📄 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

```
MIT License

Copyright (c) 2025 Lovnish Verma

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.
```

---

<div align="center">

## 🙋‍♂️ Maintainer

**Lovnish Verma**

[![LinkedIn](https://img.shields.io/badge/LinkedIn-0077B5?style=for-the-badge&logo=linkedin&logoColor=white)](https://in.linkedin.com/in/lovnishverma)
[![GitHub](https://img.shields.io/badge/GitHub-100000?style=for-the-badge&logo=github&logoColor=white)](https://github.com/lovnishverma)
[![Email](https://img.shields.io/badge/Email-D14836?style=for-the-badge&logo=gmail&logoColor=white)](mailto:technicalboyprince@gmail.com)

---

**⭐ If this project helped you, please consider giving it a star!**

[⭐ Star this project](https://github.com/lovnishverma/mailjet) | [🐛 Report Bug](https://github.com/lovnishverma/mailjet/issues) | [💡 Request Feature](https://github.com/lovnishverma/mailjet/issues/new)

---

*Built with ❤️ by developers, for developers*

</div>
