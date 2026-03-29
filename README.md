# 📧 Smart Email Sender - Batch Email Automation

[![Python 3.7+](https://img.shields.io/badge/python-3.7+-blue.svg)](https://www.python.org/downloads/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

A powerful yet simple Python tool for sending personalized bulk emails with attachments, featuring **batch sending** to respect daily email limits.

Perfect for sending certificates, invoices, reports, or any personalized documents to multiple recipients.

## ✨ Features

- 🎯 **Numbered Column Selection** - Select columns by number or name for easy configuration
- 📊 **Batch Sending** - Send emails in batches to respect daily sending limits
- 📎 **Multiple Attachment Categories** - Certificates, recommendations, invoices, etc.
- 🔤 **Smart Placeholders** - Use `{{name}}`, `{{email}}`, or any column as placeholder
- 📝 **Simple Templates** - Plain text email templates with subject line
- 🔍 **Preview Before Send** - Review the first email before bulk sending
- 📈 **Progress Tracking** - See which row you're on and how many remain
- 🔄 **Resume Capability** - Start from any row, perfect for multi-day campaigns
- ⚠️ **Error Handling** - Gracefully handles missing files and shows warnings

## 🚀 Quick Start

### Installation

```bash
# Clone the repository
git clone https://github.com/yourusername/email-sender.git
cd email-sender

# Install dependencies
pip install pandas openpyxl
```

### Basic Usage

```bash
python email_sender.py
```

## 📋 Requirements

- Python 3.7 or higher
- pandas
- openpyxl (for Excel support)

## 📁 Project Structure

```
email-sender/
├── email_sender.py          # Main script
├── email_template.txt        # Your email template
├── recipients.csv            # Your recipient data
├── certificates/             # Directory with certificates
│   ├── John_Cert.pdf
│   └── Jane_Cert.pdf
└── recommendations/          # Directory with recommendations
    ├── John_Rec.pdf
    └── Jane_Rec.pdf
```

## 🎓 Example Use Case

### 1. Create Your Data File

**recipients.csv**
```csv
Name,Email,Program,Certificate File,Recommendation File
John Doe,john@example.com,Machine Learning,John_ML_Cert.pdf,John_Rec.pdf
Jane Smith,jane@example.com,Web Development,Jane_Web_Cert.pdf,Jane_Rec.pdf
```

### 2. Create Email Template

**email_template.txt**
```
Subject: Congratulations {{name}} on Completing {{program}}!

Dear {{name}},

Congratulations on successfully completing the {{program}} program!

Your certificate and recommendation letter are attached.

Best regards,
Training Team
```

### 3. Run the Script

```bash
python email_sender.py
```

### 4. Follow the Interactive Prompts

```
Path to Excel/CSV file: recipients.csv
✓ Loaded 100 recipients

Available columns:
  1. Name
  2. Email
  3. Program
  4. Certificate File
  5. Recommendation File

Column number or name for EMAIL addresses: 2

Attach documents? (y/n): y

Directory path: ./certificates
✓ Found 100 files

Column number or name containing filenames: 4
✓ Configuration saved: Certificate File → ./certificates

Add another document category? (y/n): y

Directory path: ./recommendations
Column number or name containing filenames: 5
✓ Configuration saved: Recommendation File → ./recommendations

Add another document category? (y/n): n

BATCH SENDING CONFIGURATION
Total recipients in file: 100

Start from row number (1 to 100): 1
How many emails to send (1 to 100): 50

✓ Will send to rows 1 to 50 (50 emails)
```

## 🔄 Batch Sending Examples

### Scenario 1: Gmail Daily Limit (500 emails/day)

**Day 1:**
```
Start from row: 1
Send: 400
→ Sends rows 1-400
💡 Next time, start from row 401
```

**Day 2:**
```
Start from row: 401
Send: 100
→ Sends rows 401-500
✓ All recipients processed!
```

### Scenario 2: Testing Before Full Run

**Test Run:**
```
Start from row: 1
Send: 3
→ Test with first 3 recipients
```

**Full Run (after verification):**
```
Start from row: 1
Send: 500
→ Send to everyone
```

## 📧 SMTP Providers

The script supports all major email providers:

| Provider | SMTP Server | Port | Notes |
|----------|-------------|------|-------|
| Gmail | smtp.gmail.com | 587 | Use [App Password](https://myaccount.google.com/apppasswords) |
| Outlook | smtp-mail.outlook.com | 587 | Regular password works |
| Yahoo | smtp.mail.yahoo.com | 587 | May need app password |
| Custom | Your SMTP server | Usually 587 | Contact your provider |

### Gmail App Password Setup

1. Go to [Google App Passwords](https://myaccount.google.com/apppasswords)
2. Select "Mail" and "Other (Custom name)"
3. Generate and copy the 16-character password
4. Use this password (not your regular Gmail password)

## 🎯 Advanced Features

### Multiple Document Types

Send different document types to different recipients:

```csv
Name,Email,Certificate,Invoice,Contract
John,john@ex.com,john_cert.pdf,john_inv.pdf,john_contract.pdf
Jane,jane@ex.com,jane_cert.pdf,,jane_contract.pdf
```

Jane receives certificate and contract (invoice cell is empty).

### Flexible Placeholders

Use **any column** from your CSV as a placeholder:

```
Subject: {{greeting}} {{name}} - {{document_type}}

Dear {{title}} {{last_name}},

Your {{document_type}} for {{course_name}} (ID: {{student_id}}) is attached.

Completion Date: {{completion_date}}
Grade: {{final_grade}}

Best regards,
{{department}}
{{organization}}
```

### Column Selection Methods

Choose what works best for you:

```
Available columns:
  1. Name
  2. Email
  3. Program

Column number or name for EMAIL addresses: 2    ← Use number
Column number or name for EMAIL addresses: Email ← Or use name
```

## 📊 Output Example

```
SENDING EMAILS

[Row 1] (1/50) john@example.com
   Attachments: 2
      - John_ML_Cert.pdf
      - John_Rec.pdf
   ✓ Sent successfully!

[Row 2] (2/50) jane@example.com
   Attachments: 2
      - Jane_Web_Cert.pdf
      - Jane_Rec.pdf
   ✓ Sent successfully!

...

SUMMARY
Batch: Rows 1 to 50
Total in batch:      50
✓ Successfully sent: 50

💡 Next time, start from row 51 (50 recipients remaining)
```

## ⚠️ Important Notes

### Email Provider Limits

- **Gmail**: 500 emails/day (2000/day for Google Workspace)
- **Outlook**: 300 recipients/day
- **Yahoo**: 500 emails/day
- Always check your provider's current limits

### Security Best Practices

- ✅ Use App Passwords (not regular passwords)
- ✅ Never commit credentials to version control
- ✅ Test with a small batch first
- ✅ Keep your data files secure
- ⚠️ Never share your app passwords

### File Matching

- File names in CSV must **exactly match** actual file names
- Check for extra spaces or hidden characters
- File names are case-sensitive on Linux/Mac

## 🛠️ Troubleshooting

### "File not found" warnings

```
⚠️ File not found: John_Certificate.pdf (in ./certificates)
```

**Solutions:**
- Verify the exact filename in the directory
- Check spelling and case sensitivity
- Look for extra spaces in CSV cells
- Ensure the file exists in the specified directory

### Authentication Failed

```
✗ Failed: Authentication failed
```

**Solutions:**
- Gmail: Use App Password (not regular password)
- Outlook: Enable SMTP in account settings
- Yahoo: Generate app-specific password
- Check username is the full email address

### Invalid Column Number

```
⚠️ Please enter a number between 1 and 5
```

**Solution:** Enter a valid column number from the displayed list or the exact column name.

### Rate Limiting

If you hit daily limits:
1. Note the last successful row number
2. Wait 24 hours
3. Resume from the next row

## 📝 Tips & Tricks

💡 **Keep a log** - Write down which rows you sent each day

💡 **Test first** - Always send 2-3 test emails before the full batch

💡 **Clean data** - Remove extra spaces from CSV columns and filenames

💡 **Organize files** - Use clear directory structure and consistent naming

💡 **Double-check** - Review the preview email carefully before proceeding

💡 **Off-peak sending** - Email delivery is faster during off-peak hours

## 🤝 Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

1. Fork the repository
2. Create your feature branch (`git checkout -b feature/AmazingFeature`)
3. Commit your changes (`git commit -m 'Add some AmazingFeature'`)
4. Push to the branch (`git push origin feature/AmazingFeature`)
5. Open a Pull Request

## 📄 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## 🙏 Acknowledgments

- Built with Python and ❤️
- Uses pandas for data handling
- SMTP for reliable email delivery

## 📞 Support

If you encounter any issues or have questions:

1. Check the [Troubleshooting](#-troubleshooting) section
2. Review the example files in this repository
3. Open an issue on GitHub

## 🔮 Future Features

- [ ] HTML email templates
- [ ] Attachment preview in UI
- [ ] Email scheduling
- [ ] CSV export of send results
- [ ] GUI interface
- [ ] Email tracking/analytics

---

**Made with ❤️ for easy bulk email sending**

⭐ Star this repo if you find it helpful!
