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
import zipfile
import shutil
from typing import List, Optional

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

def get_all_folders(mail):
    """Get list of all folders from IMAP server"""
    try:
        status, folders = mail.list()
        if status != 'OK':
            print(f"Error getting folder list: {folders}")
            return []
        
        folder_list = []
        for folder in folders:
            # Parse folder name from IMAP LIST response
            # Format: '(\\HasNoChildren) "/" "INBOX"'
            folder_str = folder.decode('utf-8')
            parts = folder_str.split('"')
            if len(parts) >= 3:
                folder_name = parts[-2]  # Get the folder name
                folder_list.append(folder_name)
        
        print(f"Found {len(folder_list)} folders: {', '.join(folder_list)}")
        return folder_list
    
    except Exception as e:
        print(f"Error getting folders: {e}")
        return []

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

def download_message(mail, msg_id, output_dir, delete_after_download=False):
    """Download a single message and optionally delete it from server"""
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
        
        # Delete from server if requested and download was successful
        if delete_after_download:
            try:
                # Mark message for deletion
                mail.store(msg_id, '+FLAGS', '\\Deleted')
                print(f"  Marked message {msg_id.decode()} for deletion")
            except Exception as e:
                print(f"  Warning: Could not mark message {msg_id.decode()} for deletion: {e}")
                # Don't return False here as the download was successful
        
        return True
    
    except Exception as e:
        print(f"Error downloading message {msg_id}: {e}")
        return False

def expunge_deleted_messages(mail, folder):
    """Expunge (permanently delete) messages marked for deletion"""
    try:
        # Select the folder again to ensure we're in the right context
        mail.select(folder)
        
        # Expunge deleted messages
        mail.expunge()
        print(f"  Expunged deleted messages from {folder}")
        return True
    except Exception as e:
        print(f"  Error expunging messages from {folder}: {e}")
        return False

def get_folder_size(folder_path: Path) -> int:
    """Calculate total size of folder in bytes"""
    total_size = 0
    for dirpath, dirnames, filenames in os.walk(folder_path):
        for filename in filenames:
            filepath = os.path.join(dirpath, filename)
            try:
                total_size += os.path.getsize(filepath)
            except (OSError, IOError):
                pass
    return total_size

def create_multipart_zip(source_folder: Path, output_dir: Path, max_size_mb: int = 100) -> List[str]:
    """Create multi-part zip files from source folder"""
    max_size_bytes = max_size_mb * 1024 * 1024
    
    # Get all files to compress
    files_to_compress = []
    for root, dirs, files in os.walk(source_folder):
        for file in files:
            file_path = os.path.join(root, file)
            relative_path = os.path.relpath(file_path, source_folder)
            file_size = os.path.getsize(file_path)
            files_to_compress.append((file_path, relative_path, file_size))
    
    if not files_to_compress:
        print(f"No files found in {source_folder}")
        return []
    
    # Sort files by size (largest first) for better packing
    files_to_compress.sort(key=lambda x: x[2], reverse=True)
    
    zip_files = []
    current_zip_size = 0
    current_zip_num = 1
    
    # Create base name for zip files
    base_name = source_folder.name
    current_zip_path = output_dir / f"{base_name}_part{current_zip_num:03d}.zip"
    current_zip = zipfile.ZipFile(current_zip_path, 'w', zipfile.ZIP_DEFLATED, compresslevel=6)
    
    print(f"Creating multi-part zip for {source_folder.name}...")
    
    for file_path, relative_path, file_size in files_to_compress:
        # Check if adding this file would exceed the limit
        if current_zip_size + file_size > max_size_bytes and current_zip_size > 0:
            # Close current zip and start new one
            current_zip.close()
            zip_files.append(str(current_zip_path))
            print(f"  Created {current_zip_path.name} ({current_zip_size / (1024*1024):.1f} MB)")
            
            current_zip_num += 1
            current_zip_path = output_dir / f"{base_name}_part{current_zip_num:03d}.zip"
            current_zip = zipfile.ZipFile(current_zip_path, 'w', zipfile.ZIP_DEFLATED, compresslevel=6)
            current_zip_size = 0
        
        # Add file to current zip
        try:
            current_zip.write(file_path, relative_path)
            current_zip_size += file_size
        except Exception as e:
            print(f"  Warning: Could not add {relative_path} to zip: {e}")
    
    # Close the last zip file
    if current_zip_size > 0:
        current_zip.close()
        zip_files.append(str(current_zip_path))
        print(f"  Created {current_zip_path.name} ({current_zip_size / (1024*1024):.1f} MB)")
    
    return zip_files

def compress_folders(archive_dir: Path, max_zip_size_mb: int, keep_uncompressed: bool = False) -> dict:
    """Compress all folder archives into multi-part zip files"""
    print(f"\nCompressing archived folders (max size: {max_zip_size_mb}MB per zip)...")
    
    zip_dir = archive_dir / "compressed"
    zip_dir.mkdir(exist_ok=True)
    
    compression_summary = {
        'timestamp': datetime.now().isoformat(),
        'max_zip_size_mb': max_zip_size_mb,
        'folders_compressed': []
    }
    
    # Find all folder directories to compress
    folder_dirs = [d for d in archive_dir.iterdir() if d.is_dir() and d.name != "compressed"]
    
    for folder_dir in folder_dirs:
        if folder_dir.name == "compressed":
            continue
            
        folder_size = get_folder_size(folder_dir)
        print(f"\nProcessing {folder_dir.name} ({folder_size / (1024*1024):.1f} MB)...")
        
        # Create multi-part zip
        zip_files = create_multipart_zip(folder_dir, zip_dir, max_zip_size_mb)
        
        if zip_files:
            folder_info = {
                'folder_name': folder_dir.name,
                'original_size_mb': round(folder_size / (1024*1024), 2),
                'zip_files': [os.path.basename(zf) for zf in zip_files],
                'zip_count': len(zip_files)
            }
            
            # Calculate total compressed size
            total_compressed_size = sum(os.path.getsize(zf) for zf in zip_files)
            folder_info['compressed_size_mb'] = round(total_compressed_size / (1024*1024), 2)
            folder_info['compression_ratio'] = round((1 - total_compressed_size / folder_size) * 100, 1)
            
            compression_summary['folders_compressed'].append(folder_info)
            
            print(f"  Compressed to {len(zip_files)} parts")
            print(f"  Compression: {folder_info['original_size_mb']}MB → {folder_info['compressed_size_mb']}MB ({folder_info['compression_ratio']}% reduction)")
            
            # Remove original folder if not keeping uncompressed
            if not keep_uncompressed:
                print(f"  Removing original folder {folder_dir.name}")
                shutil.rmtree(folder_dir)
    
    # Save compression summary
    with open(zip_dir / 'compression_summary.json', 'w') as f:
        json.dump(compression_summary, f, indent=2)
    
    return compression_summary

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
        
        # Get folders to process
        if args.all_folders:
            print("Getting all folders from server...")
            folders = get_all_folders(mail)
            if not folders:
                print("No folders found or error getting folders")
                return False
        else:
            folders = args.folders if args.folders else ['INBOX']
        
        print(f"Processing folders: {', '.join(folders)}")
        
        # Confirmation for deletion
        if args.delete_messages:
            print("\nWARNING: Messages will be PERMANENTLY DELETED from the server after successful download!")
            if not args.force_delete:
                confirm = input("Are you sure you want to continue? (yes/no): ").lower()
                if confirm != 'yes':
                    print("Operation cancelled.")
                    return False
        
        total_downloaded = 0
        total_errors = 0
        total_deleted = 0
        
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
            folder_downloaded = 0
            folder_errors = 0
            
            for i, msg_id in enumerate(msg_ids, 1):
                if args.limit and i > args.limit:
                    print(f"Reached limit of {args.limit} messages")
                    break
                
                print(f"Downloading message {i}/{len(msg_ids)}: {msg_id.decode()}")
                
                if download_message(mail, msg_id, str(folder_dir), args.delete_messages):
                    total_downloaded += 1
                    folder_downloaded += 1
                    if args.delete_messages:
                        total_deleted += 1
                else:
                    total_errors += 1
                    folder_errors += 1
                
                # Progress update
                if i % 10 == 0:
                    print(f"Progress: {i}/{len(msg_ids)} messages processed")
            
            # Expunge deleted messages if deletion was enabled
            if args.delete_messages and folder_downloaded > 0:
                print(f"Expunging {folder_downloaded} deleted messages from {folder}...")
                expunge_deleted_messages(mail, folder)
            
            print(f"Folder {folder} complete: {folder_downloaded} downloaded, {folder_errors} errors")
        
        print(f"\nArchiving complete!")
        print(f"Total downloaded: {total_downloaded}")
        print(f"Total errors: {total_errors}")
        if args.delete_messages:
            print(f"Total deleted from server: {total_deleted}")
        
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
            'total_errors': total_errors,
            'total_deleted': total_deleted if args.delete_messages else 0,
            'delete_messages': args.delete_messages
        }
        
        with open(output_dir / 'archive_summary.json', 'w') as f:
            json.dump(summary, f, indent=2)
        
        # Compress folders if requested
        if args.compress:
            compression_summary = compress_folders(output_dir, args.max_zip_size, args.keep_uncompressed)
            print(f"\nCompression complete! Created {len(compression_summary['folders_compressed'])} compressed folder sets.")
        
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
  # Archive all folders with compression and delete from server
  python imap_archiver.py -s mail.example.com -u user@example.com --password mypass --all-folders --delete-messages --compress

  # Archive specific date range from all folders (automated deletion)
  python imap_archiver.py -s mail.example.com -u user@example.com --password mypass --all-folders --start-date 2023-01-01 --end-date 2023-12-31 --delete-messages --force-delete

  # Archive with 50MB zip limit and keep originals
  python imap_archiver.py -s mail.example.com -u user@example.com --password mypass --folders INBOX Sent --compress --max-zip-size 50 --keep-uncompressed

  # Safe archive without deletion (backup mode)
  python imap_archiver.py -s mail.example.com -u user@example.com --password mypass --all-folders --compress
        """
    )
    
    # Server configuration
    parser.add_argument('-s', '--server', required=True,
                       help='IMAP server address')
    parser.add_argument('-p', '--port', type=int, default=993,
                       help='IMAP server port (default: 993 for SSL, 143 for non-SSL)')
    parser.add_argument('-u', '--username', required=True,
                       help='Username for IMAP login')
    parser.add_argument('--password', required=True,
                       help='Password for IMAP login (required for automation)')
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
    parser.add_argument('--all-folders', action='store_true',
                       help='Archive all folders from the server (overrides --folders)')
    parser.add_argument('--limit', type=int,
                       help='Maximum number of messages to download per folder')
    
    # Message deletion options
    parser.add_argument('--delete-messages', action='store_true',
                       help='Delete messages from server after successful download')
    parser.add_argument('--force-delete', action='store_true',
                       help='Skip confirmation prompt for message deletion (use with caution!)')
    
    # Output
    parser.add_argument('-o', '--output-dir', default='email_archive',
                       help='Output directory for archived messages (default: email_archive)')
    
    # Compression options
    parser.add_argument('--compress', action='store_true',
                       help='Compress archived folders into multi-part zip files')
    parser.add_argument('--max-zip-size', type=int, default=100,
                       help='Maximum size for each zip file in MB (default: 100)')
    parser.add_argument('--keep-uncompressed', action='store_true',
                       help='Keep original uncompressed folders after compression')
    
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
