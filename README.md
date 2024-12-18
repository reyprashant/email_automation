# Automated Certificate Email Sender

## Overview
This Python script automates the process of sending participation certificates to event participants via email. It matches participant names with certificates and sends personalized emails with attachments. Here excel file is used.

## Features
- Automatically find and match certificates for participants
- Send personalized emails with certificates
- Support for various certificate filename formats
- Flexible name matching using fuzzy search
- Error handling and logging

## Prerequisites

### System Requirements
- Python 3.7+
- Windows/macOS/Linux

### Required Libraries
- pandas
- openpyxl

## Installation

1. **Clone the Repository**
```bash
git clone (https://github.com/reyprashant/email_automation.git)
cd certificate-email-sender
```

2. **Install Dependencies**
```bash
pip install pandas openpyxl
```

## Configuration

### 1. Prepare Excel File
- Create an Excel file (`participants.xlsx`)
- Columns required:
  - `Name`: Full name of participant
  - `Email`: Participant's email address

**Example:**
| Name | Email |
|------|-------|
| John Doe | john.doe@example.com |
| Jane Smith | jane.smith@example.com |

### 2. Certificates Folder
- Place certificates in a dedicated folder
- Filename should match or be similar to participant names
- Supported formats: .png, .jpg, .pdf

**Filename Examples:**
- `John_Doe.png`
- `Jane_Smith.pdf`
- `Participant_Name.jpg`

### 3. Email Configuration
1. **Google Account Setup**
   - Enable 2-Step Verification
   - Generate App Password
     - Go to Google Account > Security
     - Create App Password for "Mail"

2. **Update Script**
   - Replace email configuration in script
   - Use generated App Password

## Usage

### Running the Script
```bash
python main.py
```

### Command Line Execution
```bash
python main.py
```

## Troubleshooting

### Common Issues
1. **Authentication Errors**
   - Verify App Password
   - Check internet connection
   - Ensure 2-Step Verification is enabled

2. **Certificate Matching**
   - Ensure certificate filenames are close to participant names
   - Use consistent naming conventions

### Logging
- Script provides console output
- Tracks successful and failed email sends

## Security Considerations
- Use App Passwords, not regular passwords
- Do not share sensitive credentials
- Implement additional security as needed

## Contributing
1. Fork the repository
2. Create your feature branch
3. Commit your changes
4. Push to the branch
5. Create a Pull Request
