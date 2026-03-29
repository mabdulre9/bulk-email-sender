#!/usr/bin/env python3
"""
Smart Email Sender - Filename Column Matching
Uses columns containing exact filenames for attachment matching
"""

import smtplib
import os
import sys
import re
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import pandas as pd
from pathlib import Path


class EmailSender:
    def __init__(self):
        self.smtp_server = None
        self.smtp_port = None
        self.sender_email = None
        self.sender_password = None
        self.server = None
        
    def setup_smtp(self):
        """Configure SMTP settings"""
        print("\n" + "=" * 60)
        print("SMTP CONFIGURATION")
        print("=" * 60)
        print("\nCommon providers:")
        print("1. Gmail (smtp.gmail.com:587)")
        print("2. Outlook/Hotmail (smtp-mail.outlook.com:587)")
        print("3. Yahoo (smtp.mail.yahoo.com:587)")
        print("4. Custom SMTP server")
        
        choice = input("\nSelect provider (1-4): ").strip()
        
        if choice == "1":
            self.smtp_server = "smtp.gmail.com"
            self.smtp_port = 587
            print("\n⚠️  Gmail: Use App Password (not regular password)")
            print("   Generate at: https://myaccount.google.com/apppasswords")
        elif choice == "2":
            self.smtp_server = "smtp-mail.outlook.com"
            self.smtp_port = 587
        elif choice == "3":
            self.smtp_server = "smtp.mail.yahoo.com"
            self.smtp_port = 587
        else:
            self.smtp_server = input("SMTP Server: ").strip()
            self.smtp_port = int(input("SMTP Port (usually 587): ").strip())
        
        self.sender_email = input("\nYour email address: ").strip()
        self.sender_password = input("Your password/app password: ").strip()
        
    def connect(self):
        """Establish SMTP connection"""
        try:
            print(f"\nConnecting to {self.smtp_server}:{self.smtp_port}...", end=" ")
            self.server = smtplib.SMTP(self.smtp_server, self.smtp_port)
            self.server.starttls()
            self.server.login(self.sender_email, self.sender_password)
            print("✓ Connected!")
            return True
        except Exception as e:
            print(f"✗ Failed: {e}")
            return False
    
    def disconnect(self):
        """Close SMTP connection"""
        if self.server:
            self.server.quit()
    
    def send_email(self, recipient_email, subject, body, attachments):
        """Send a single email with attachments"""
        try:
            msg = MIMEMultipart()
            msg['From'] = self.sender_email
            msg['To'] = recipient_email
            msg['Subject'] = subject
            
            msg.attach(MIMEText(body, 'plain'))
            
            # Attach files
            for filepath in attachments:
                if not os.path.exists(filepath):
                    print(f"      ⚠️  File not found: {filepath}")
                    continue
                    
                with open(filepath, 'rb') as attachment:
                    part = MIMEBase('application', 'octet-stream')
                    part.set_payload(attachment.read())
                
                encoders.encode_base64(part)
                filename = os.path.basename(filepath)
                part.add_header('Content-Disposition', f'attachment; filename= {filename}')
                msg.attach(part)
            
            self.server.send_message(msg)
            return True
            
        except Exception as e:
            print(f"      ✗ Error: {e}")
            return False


def display_columns(df):
    """Display columns in numbered format"""
    print("\nAvailable columns:")
    for idx, col in enumerate(df.columns, 1):
        print(f"  {idx}. {col}")


def get_column_choice(df, prompt_text):
    """Get column selection from user (accepts number or name)"""
    while True:
        choice = input(f"\n{prompt_text}: ").strip()
        
        # Try as number first
        if choice.isdigit():
            col_num = int(choice)
            if 1 <= col_num <= len(df.columns):
                return df.columns[col_num - 1]
            else:
                print(f"⚠️  Please enter a number between 1 and {len(df.columns)}")
                continue
        
        # Try as column name
        if choice in df.columns:
            return choice
        
        print(f"⚠️  '{choice}' not found. Enter column number (1-{len(df.columns)}) or exact name")


def load_recipients(file_path):
    """Load recipients from Excel or CSV file"""
    try:
        if file_path.lower().endswith('.csv'):
            df = pd.read_csv(file_path)
        else:
            df = pd.read_excel(file_path)
        
        # Clean up column names (remove extra spaces)
        df.columns = df.columns.str.strip()
        
        print(f"\n✓ Loaded {len(df)} recipients")
        
        return df
    except Exception as e:
        print(f"✗ Error loading file: {e}")
        sys.exit(1)


def configure_attachments(df):
    """Configure attachment columns and directories"""
    print("\n" + "=" * 60)
    print("ATTACHMENT CONFIGURATION")
    print("=" * 60)
    
    attachments_config = []
    
    attach = input("\nAttach documents? (y/n): ").strip().lower()
    if attach not in ['y', 'yes']:
        return attachments_config
    
    print("\nYou can configure multiple document categories.")
    print("Each category needs:")
    print("  - Directory containing the files")
    print("  - Column name containing filenames")
    
    while True:
        print("\n" + "-" * 60)
        print("Document Category Configuration")
        print("-" * 60)
        
        # Get directory
        directory = input("\nDirectory path (where files are stored): ").strip()
        if not os.path.exists(directory):
            print(f"✗ Directory not found: {directory}")
            retry = input("Try again? (y/n): ").strip().lower()
            if retry not in ['y', 'yes']:
                break
            continue
        
        # Show files in directory
        files = sorted([f.name for f in Path(directory).iterdir() if f.is_file()])
        print(f"✓ Found {len(files)} files in directory")
        if len(files) <= 10:
            print("Files:", ', '.join(files))
        
        # Get column name
        display_columns(df)
        column_name = get_column_choice(df, "Column number or name containing filenames")
        
        # Add configuration
        attachments_config.append({
            'directory': directory,
            'column': column_name
        })
        
        print(f"✓ Configuration saved: {column_name} → {directory}")
        
        # Ask for another
        another = input("\nAdd another document category? (y/n): ").strip().lower()
        if another not in ['y', 'yes']:
            break
    
    return attachments_config


def get_attachments_for_recipient(row, attachments_config):
    """Get list of attachment files for a specific recipient"""
    all_attachments = []
    
    for config in attachments_config:
        column = config['column']
        directory = config['directory']
        
        # Get filename from row
        if column in row.index:
            filename = str(row[column]).strip()
            
            # Skip if empty or NaN
            if filename and filename.lower() != 'nan':
                filepath = Path(directory) / filename
                if filepath.exists():
                    all_attachments.append(str(filepath))
                else:
                    print(f"      ⚠️  File not found: {filename} (in {directory})")
    
    return all_attachments


def load_email_template(template_path):
    """Load email template from text file"""
    try:
        with open(template_path, 'r', encoding='utf-8') as f:
            content = f.read()
        
        # Extract subject (first line)
        lines = content.strip().split('\n')
        subject = lines[0].replace('Subject:', '').strip()
        body = '\n'.join(lines[1:]).strip()
        
        return subject, body
    except Exception as e:
        print(f"✗ Error loading template: {e}")
        sys.exit(1)


def find_placeholders(text):
    """Find all {{placeholder}} patterns in text"""
    # Use set to automatically remove duplicates
    placeholders = set(re.findall(r'\{\{([^}]+)\}\}', text))
    return sorted(list(placeholders))


def replace_placeholders(text, row, placeholder_mapping):
    """Replace placeholders with values from row"""
    result = text
    for placeholder, column in placeholder_mapping.items():
        if column in row.index:
            value = str(row[column])
            result = result.replace(f'{{{{{placeholder}}}}}', value)
    return result


def configure_placeholders(template_path, df):
    """Configure which columns map to which placeholders"""
    subject, body = load_email_template(template_path)
    
    # Find all placeholders (automatically removes duplicates)
    all_text = subject + '\n' + body
    placeholders = find_placeholders(all_text)
    
    if not placeholders:
        print("\n✓ No placeholders found in template")
        return {}
    
    print("\n" + "=" * 60)
    print("PLACEHOLDER MAPPING")
    print("=" * 60)
    print(f"\nFound placeholders: {', '.join(['{{' + p + '}}' for p in placeholders])}")
    
    mapping = {}
    for placeholder in placeholders:
        display_columns(df)
        column = get_column_choice(df, f"Map {{{{{placeholder}}}}} to column number or name")
        mapping[placeholder] = column
        print(f"✓ {{{{{placeholder}}}}} → {column}")
    
    return mapping


def preview_email(subject, body, row, placeholder_mapping, attachments):
    """Show preview of how email will look"""
    print("\n" + "=" * 60)
    print("EMAIL PREVIEW (First Recipient)")
    print("=" * 60)
    
    preview_subject = replace_placeholders(subject, row, placeholder_mapping)
    preview_body = replace_placeholders(body, row, placeholder_mapping)
    
    print(f"\nSubject: {preview_subject}")
    print(f"\n{preview_body}")
    
    if attachments:
        print(f"\nAttachments ({len(attachments)}):")
        for att in attachments:
            print(f"  - {os.path.basename(att)}")
    else:
        print("\nAttachments: None")
    
    print("=" * 60)


def main():
    print("\n" + "=" * 60)
    print("  SMART EMAIL SENDER - FILENAME COLUMN MATCHING")
    print("=" * 60)
    
    # Load recipients
    print("\n" + "=" * 60)
    print("RECIPIENTS DATA")
    print("=" * 60)
    excel_file = input("\nPath to Excel/CSV file: ").strip()
    if not os.path.exists(excel_file):
        print(f"✗ File not found: {excel_file}")
        sys.exit(1)
    
    df = load_recipients(excel_file)
    
    # Get email column
    display_columns(df)
    email_col = get_column_choice(df, "Column number or name for EMAIL addresses")
    
    # Configure attachments
    attachments_config = configure_attachments(df)
    
    if attachments_config:
        print(f"\n✓ Configured {len(attachments_config)} document category/categories")
    
    # Load email template
    print("\n" + "=" * 60)
    print("EMAIL TEMPLATE")
    print("=" * 60)
    template_path = input("\nPath to email template text file: ").strip()
    if not os.path.exists(template_path):
        print(f"✗ Template file not found: {template_path}")
        sys.exit(1)
    
    subject, body = load_email_template(template_path)
    
    # Configure placeholder mapping
    placeholder_mapping = configure_placeholders(template_path, df)
    
    # Batch configuration
    print("\n" + "=" * 60)
    print("BATCH SENDING CONFIGURATION")
    print("=" * 60)
    print(f"\nTotal recipients in file: {len(df)}")
    print("Rows are numbered starting from 1 (first data row)")
    
    while True:
        try:
            start_row = int(input("\nStart from row number (1 to {}): ".format(len(df))).strip())
            if 1 <= start_row <= len(df):
                break
            else:
                print(f"⚠️  Please enter a number between 1 and {len(df)}")
        except ValueError:
            print("⚠️  Please enter a valid number")
    
    max_emails = len(df) - start_row + 1
    while True:
        try:
            num_emails = int(input(f"How many emails to send (1 to {max_emails}): ").strip())
            if 1 <= num_emails <= max_emails:
                break
            else:
                print(f"⚠️  Please enter a number between 1 and {max_emails}")
        except ValueError:
            print("⚠️  Please enter a valid number")
    
    # Calculate end row (inclusive)
    end_row = start_row + num_emails - 1
    
    print(f"\n✓ Will send to rows {start_row} to {end_row} ({num_emails} emails)")
    
    # Slice dataframe for the batch
    df_batch = df.iloc[start_row-1:end_row].copy()
    
    # Preview with first recipient of the batch
    print("\n" + "=" * 60)
    print("PREVIEW (First Email in Batch)")
    print("=" * 60)
    first_row = df_batch.iloc[0]
    first_attachments = get_attachments_for_recipient(first_row, attachments_config)
    preview_email(subject, body, first_row, placeholder_mapping, first_attachments)
    
    proceed = input("\nLooks good? Proceed with sending? (y/n): ").strip().lower()
    if proceed not in ['y', 'yes']:
        print("\nCancelled.")
        sys.exit(0)
    
    # Setup SMTP
    sender = EmailSender()
    sender.setup_smtp()
    
    if not sender.connect():
        sys.exit(1)
    
    # Send emails
    print("\n" + "=" * 60)
    print("SENDING EMAILS")
    print("=" * 60)
    
    success_count = 0
    failed_count = 0
    skipped_count = 0
    
    for batch_idx, (original_idx, row) in enumerate(df_batch.iterrows()):
        recipient_email = str(row[email_col])
        
        # Calculate actual row number in original file (1-indexed)
        actual_row_num = start_row + batch_idx
        
        print(f"\n[Row {actual_row_num}] ({batch_idx + 1}/{len(df_batch)}) {recipient_email}")
        
        # Get attachments for this recipient
        attachments = get_attachments_for_recipient(row, attachments_config)
        
        # Check if we should skip (only if attachments were configured but none found)
        if attachments_config and not attachments:
            print(f"   ⚠️  No attachments found - skipping")
            skipped_count += 1
            continue
        
        # Prepare personalized content
        personalized_subject = replace_placeholders(subject, row, placeholder_mapping)
        personalized_body = replace_placeholders(body, row, placeholder_mapping)
        
        if attachments:
            print(f"   Attachments: {len(attachments)}")
            for att in attachments:
                print(f"      - {os.path.basename(att)}")
        else:
            print(f"   Attachments: None")
        
        # Send email
        if sender.send_email(recipient_email, personalized_subject, personalized_body, attachments):
            print(f"   ✓ Sent successfully!")
            success_count += 1
        else:
            failed_count += 1
    
    # Disconnect
    sender.disconnect()
    
    # Summary
    print("\n" + "=" * 60)
    print("SUMMARY")
    print("=" * 60)
    print(f"Batch: Rows {start_row} to {end_row}")
    print(f"Total in batch:      {len(df_batch)}")
    print(f"✓ Successfully sent: {success_count}")
    if skipped_count > 0:
        print(f"⚠ Skipped:          {skipped_count}")
    if failed_count > 0:
        print(f"✗ Failed:           {failed_count}")
    print("=" * 60)
    
    # Calculate next starting row
    remaining = len(df) - end_row
    if remaining > 0:
        next_start = end_row + 1
        print(f"\n💡 Next time, start from row {next_start} ({remaining} recipients remaining)")
    else:
        print(f"\n✓ All recipients processed!")
    
    print("=" * 60 + "\n")


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n\n✗ Interrupted by user")
        sys.exit(1)
    except Exception as e:
        print(f"\n✗ Unexpected error: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)
