
import os
import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
import difflib


class CertificateSender:
    def __init__(self, excel_path, certificates_folder, email_config):
        """
        Initialize the certificate sender

        :param excel_path: Path to Excel file with participant details
        :param certificates_folder: Folder containing certificate images/PDFs
        :param email_config: Dictionary with email configuration
        """
        # Read the Excel file
        self.participants_df = pd.read_excel(excel_path)

        # Validate required columns
        required_columns = ['Name', 'Email']
        for col in required_columns:
            if col not in self.participants_df.columns:
                raise ValueError(f"Missing required column: {col}")

        # Store configuration
        self.certificates_folder = certificates_folder
        self.email_config = email_config

        # Preload certificate filenames
        self.certificate_files = os.listdir(certificates_folder)

    def _find_matching_certificate(self, participant_name):
        """
        Find the best matching certificate for a participant name

        :param participant_name: Name of the participant
        :return: Matching certificate filename or None
        """
        # Remove any extra spaces and normalize name
        normalized_name = participant_name.replace(' ', '_')

        # First, try direct match
        direct_matches = [
            f for f in self.certificate_files
            if normalized_name in f or f.startswith(normalized_name)
        ]

        if direct_matches:
            return direct_matches[0]

        # If no direct match, use fuzzy matching
        best_match = None
        best_ratio = 0

        for filename in self.certificate_files:
            # Remove extension for comparison
            name_without_ext = os.path.splitext(filename)[0]

            # Calculate similarity ratio
            ratio = difflib.SequenceMatcher(None, normalized_name, name_without_ext).ratio()

            if ratio > best_ratio and ratio > 0.7:  # 70% similarity threshold
                best_match = filename
                best_ratio = ratio

        return best_match

    def _create_email_message(self, recipient_name, recipient_email, certificate_path):
        """
        Create email message with certificate attachment

        :param recipient_name: Name of the recipient
        :param recipient_email: Email of the recipient
        :param certificate_path: Path to the certificate file
        :return: Prepared email message
        """
        # Create the email message
        msg = MIMEMultipart()
        msg['From'] = self.email_config['sender_email']
        msg['To'] = recipient_email
        msg['Subject'] = 'Tech Talk Event - Participation Certificate'

        # Email body
        body = f"""Dear {recipient_name},

Thank you for participating in our Tech Talk event. Please find attached your certificate of participation.

Best regards,
Project Lead
Prashant Adhikari"""

        msg.attach(MIMEText(body, 'plain'))

        # Attach certificate
        with open(certificate_path, 'rb') as file:
            part = MIMEApplication(file.read(), Name=os.path.basename(certificate_path))
            part['Content-Disposition'] = f'attachment; filename="{os.path.basename(certificate_path)}"'
            msg.attach(part)

        return msg

    def send_certificates(self):
        """
        Send certificates to all participants
        """
        # Setup SMTP connection
        try:
            with smtplib.SMTP(self.email_config['smtp_server'], self.email_config['smtp_port']) as server:
                server.starttls()
                server.login(self.email_config['sender_email'], self.email_config['sender_password'])

                # Process each participant
                for index, row in self.participants_df.iterrows():
                    name = row['Name']
                    email = row['Email']

                    # Find matching certificate
                    certificate_filename = self._find_matching_certificate(name)

                    # Check if certificate exists
                    if not certificate_filename:
                        print(f"Certificate not found for {name}. Skipping.")
                        continue

                    certificate_path = os.path.join(self.certificates_folder, certificate_filename)

                    # Create and send email
                    msg = self._create_email_message(name, email, certificate_path)
                    server.send_message(msg)

                    print(f"Certificate sent to {name} at {email}")

        except Exception as e:
            print(f"An error occurred: {e}")


# Example usage
if __name__ == "__main__":
    # IMPORTANT CONFIGURATION SETTINGS
    # --------------------------------

    # EMAIL CONFIGURATION
    # REPLACE WITH YOUR SPECIFIC DETAILS
    email_config = {
        # Use Gmail SMTP server
        'smtp_server': 'smtp.gmail.com',

        # Standard TLS port for Gmail
        'smtp_port': 587,

        # SENDER EMAIL - ALREADY CONFIGURED
        'sender_email': 'subash96240@gmail.com',

        # CRITICAL: YOU MUST GENERATE AN APP PASSWORD
        # DO NOT USE YOUR REGULAR GMAIL PASSWORD
        # Steps to generate App Password:
        # 1. Go to Google Account
        # 2. Security > 2-Step Verification
        # 3. App Passwords > Select App (Mail) and Device
        # 4. Generate and use THAT password HERE
        'sender_password': 'sxpwcinyjhnyebcb'  # REPLACE THIS IMMEDIATELY
    }

    # FILE PATHS - CONFIGURED AS SPECIFIED
    # Use raw string to handle Windows file paths correctly
    excel_path = r'D:\automated_email_sent\participants.xlsx'
    certificates_folder = r'D:\automated_email_sent\certificates'

    # Create and run the certificate sender
    sender = CertificateSender(excel_path, certificates_folder, email_config)
    sender.send_certificates()



