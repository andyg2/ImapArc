#!/usr/bin/env python3
"""
IMAP Email Archiver - Download old messages from IMAP server for archival
"""

import imaplib
import email
import argparse
import os
import sys
from datetime import datetime, timedelta
import json
import ssl
from pathlib import Path
import getpass

def create_ssl_context():
    """Create SSL context for secure connection"""
    context = ssl.create_default_context()
    # You might need to adjust these settings based on your server
    context.check_hostname = False
    context.verify_mode = ssl.CERT_REQUIRED
    return context

def connect_to_server(server, port, username, password, use_ssl=True):
    """Connect to IMAP server"""
    try:
        if use_ssl:
            context = create_ssl_context()
            mail = imaplib.IMAP4_SSL(server, port, ssl_context=context)
        else:
            mail = imaplib.IMAP4(server, port)
        
        mail.login(username, password)
        print(f"Successfully connected to {server}:{port}")
        return mail
    except Exception as e:
        print(f"Error connecting to server: {e}")
        return None

def get_date_search_criteria(start_date, end_date):
    """Convert date range to IMAP search criteria"""
    criteria = []
    
    if start_date:
        # IMAP date format: DD-Mon-YYYY
        start_str = start_date.strftime("%d-%b-%Y")
        criteria.append(f'SINCE {start_str}')
    
    if end_date:
        end_str = end_date.strftime("%d-%b-%Y")
        criteria.append(f'BEFORE {end_str}')
    
    return ' '.join(criteria)

def search_messages(mail, folder, date_criteria):
    """Search for messages based on criteria"""
    try:
        # Select folder
        status, messages = mail.select(folder)
        if status != 'OK':
            print(f"Error selecting folder {folder}: {messages}")
            return []
        
        # Search for messages
        search_criteria = 'ALL'
        if date_criteria:
            search_criteria = date_criteria
        
        status, message_ids = mail.search(None, search_criteria)
        if status != 'OK':
            print(f"Error searching messages: {message_ids}")
            return []
        
        # Convert to list of message IDs
        msg_ids = message_ids[0].split()
        print(f"Found {len(msg_ids)} messages in {folder}")
        return msg_ids
    
    except Exception as e:
        print(f"Error searching messages: {e}")
        return []

def download_message(mail, msg_id, output_dir):
    """Download a single message"""
    try:
        # Fetch message
        status, msg_data = mail.fetch(msg_id, '(RFC822)')
        if status != 'OK':
            print(f"Error fetching message {msg_id}: {msg_data}")
            return False
        
        # Parse email
        raw_email = msg_data[0][1]
        email_message = email.message_from_bytes(raw_email)
        
        # Generate filename
        subject = email_message.get('Subject', 'No Subject')
        date_str = email_message.get('Date', '')
        from_addr = email_message.get('From', 'unknown')
        
        # Clean filename
        safe_subject = "".join(c for c in subject if c.isalnum() or c in (' ', '-', '_')).rstrip()
        safe_subject = safe_subject[:50]  # Limit length
        
        filename = f"{msg_id.decode()}_{safe_subject}.eml"
        filepath = os.path.join(output_dir, filename)
        
        # Save message
        with open(filepath, 'wb') as f:
            f.write(raw_email)
        
        # Save metadata
        metadata = {
            'message_id': msg_id.decode(),
            'subject': subject,
            'from': from_addr,
            'date': date_str,
            'filename': filename
        }
        
        metadata_file = filepath.replace('.eml', '_metadata.json')
        with open(metadata_file, 'w') as f:
            json.dump(metadata, f, indent=2)
        
        return True
    
    except Exception as e:
        print(f"Error downloading message {msg_id}: {e}")
        return False

def archive_messages(args):
    """Main archiving function"""
    # Create output directory
    output_dir = Path(args.output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    
    # Get password if not provided
    password = args.password
    if not password:
        password = getpass.getpass(f"Password for {args.username}: ")
    
    # Connect to server
    mail = connect_to_server(args.server, args.port, args.username, password, args.ssl)
    if not mail:
        return False
    
    try:
        # Get date criteria
        date_criteria = get_date_search_criteria(args.start_date, args.end_date)
        print(f"Date criteria: {date_criteria if date_criteria else 'All messages'}")
        
        # Process folders
        folders = args.folders if args.folders else ['INBOX']
        
        total_downloaded = 0
        total_errors = 0
        
        for folder in folders:
            print(f"\nProcessing folder: {folder}")
            
            # Create folder-specific output directory
            folder_dir = output_dir / folder.replace('/', '_')
            folder_dir.mkdir(exist_ok=True)
            
            # Search messages
            msg_ids = search_messages(mail, folder, date_criteria)
            
            if not msg_ids:
                print(f"No messages found in {folder}")
                continue
            
            # Download messages
            for i, msg_id in enumerate(msg_ids, 1):
                if args.limit and i > args.limit:
                    print(f"Reached limit of {args.limit} messages")
                    break
                
                print(f"Downloading message {i}/{len(msg_ids)}: {msg_id.decode()}")
                
                if download_message(mail, msg_id, str(folder_dir)):
                    total_downloaded += 1
                else:
                    total_errors += 1
                
                # Progress update
                if i % 10 == 0:
                    print(f"Progress: {i}/{len(msg_ids)} messages processed")
        
        print(f"\nArchiving complete!")
        print(f"Total downloaded: {total_downloaded}")
        print(f"Total errors: {total_errors}")
        
        # Create summary report
        summary = {
            'timestamp': datetime.now().isoformat(),
            'server': args.server,
            'folders': folders,
            'date_range': {
                'start': args.start_date.isoformat() if args.start_date else None,
                'end': args.end_date.isoformat() if args.end_date else None
            },
            'total_downloaded': total_downloaded,
            'total_errors': total_errors
        }
        
        with open(output_dir / 'archive_summary.json', 'w') as f:
            json.dump(summary, f, indent=2)
        
        return True
    
    finally:
        mail.close()
        mail.logout()

def parse_date(date_string):
    """Parse date string to datetime object"""
    try:
        return datetime.strptime(date_string, '%Y-%m-%d')
    except ValueError:
        raise argparse.ArgumentTypeError(f"Invalid date format: {date_string}. Use YYYY-MM-DD")

def main():
    parser = argparse.ArgumentParser(
        description="Download old messages from IMAP server for archival",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Archive all messages from last year
  python imap_archiver.py -s mail.example.com -u user@example.com --start-date 2023-01-01 --end-date 2023-12-31

  # Archive messages from specific folders
  python imap_archiver.py -s mail.example.com -u user@example.com --folders INBOX Sent --limit 100

  # Archive with custom port and no SSL
  python imap_archiver.py -s mail.example.com -p 143 -u user@example.com --no-ssl
        """
    )
    
    # Server configuration
    parser.add_argument('-s', '--server', required=True,
                       help='IMAP server address')
    parser.add_argument('-p', '--port', type=int, default=993,
                       help='IMAP server port (default: 993 for SSL, 143 for non-SSL)')
    parser.add_argument('-u', '--username', required=True,
                       help='Username for IMAP login')
    parser.add_argument('--password',
                       help='Password for IMAP login (will prompt if not provided)')
    parser.add_argument('--no-ssl', action='store_false', dest='ssl',
                       help='Disable SSL connection')
    
    # Date range
    parser.add_argument('--start-date', type=parse_date,
                       help='Start date for message range (YYYY-MM-DD)')
    parser.add_argument('--end-date', type=parse_date,
                       help='End date for message range (YYYY-MM-DD)')
    
    # Folders and limits
    parser.add_argument('--folders', nargs='+', default=['INBOX'],
                       help='Folders to archive (default: INBOX)')
    parser.add_argument('--limit', type=int,
                       help='Maximum number of messages to download per folder')
    
    # Output
    parser.add_argument('-o', '--output-dir', default='email_archive',
                       help='Output directory for archived messages (default: email_archive)')
    
    args = parser.parse_args()
    
    # Validate date range
    if args.start_date and args.end_date and args.start_date > args.end_date:
        print("Error: Start date must be before end date")
        sys.exit(1)
    
    # Adjust default port for non-SSL
    if not args.ssl and args.port == 993:
        args.port = 143
    
    print(f"Starting email archival from {args.server}:{args.port}")
    print(f"Username: {args.username}")
    print(f"SSL: {'Enabled' if args.ssl else 'Disabled'}")
    print(f"Output directory: {args.output_dir}")
    
    if archive_messages(args):
        print("Archival completed successfully!")
    else:
        print("Archival failed!")
        sys.exit(1)

if __name__ == "__main__":
    main()
