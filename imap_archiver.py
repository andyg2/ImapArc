#!/usr/bin/env python3
"""
IMAP Email Archiver - Download old messages from IMAP server for archival
Now with robust connection handling to survive network timeouts and terminal pauses.
"""

import imaplib
import email
import argparse
import os
import re
import sys
from datetime import datetime, timedelta
import json
import ssl
from pathlib import Path
import getpass
import zipfile
import shutil
from typing import List, Optional, Tuple
from email.utils import parsedate_to_datetime

# --- NEW CONNECTION MANAGER CLASS ---
class IMAPConnectionManager:
    """
    Manages the IMAP connection, credentials, and handles automatic reconnection.
    """
    def __init__(self, server, port, username, password, use_ssl=True):
        self.server = server
        self.port = port
        self.username = username
        self.password = password
        self.use_ssl = use_ssl
        self.mail: Optional[imaplib.IMAP4] = None

    def _create_ssl_context(self):
        context = ssl.create_default_context()
        context.check_hostname = False
        context.verify_mode = ssl.CERT_REQUIRED
        return context

    def connect(self) -> bool:
        """Establishes the initial connection to the IMAP server."""
        try:
            if self.use_ssl:
                context = self._create_ssl_context()
                self.mail = imaplib.IMAP4_SSL(self.server, self.port, ssl_context=context)
            else:
                self.mail = imaplib.IMAP4(self.server, self.port)
            
            self.mail.login(self.username, self.password)
            print(f"Successfully connected to {self.server}:{self.port}")
            return True
        except Exception as e:
            print(f"Error connecting to server: {e}")
            self.mail = None
            return False

    def disconnect(self):
        """Closes the connection cleanly."""
        if self.mail:
            try:
                self.mail.close()
                self.mail.logout()
                print("Disconnected from server.")
            except Exception as e:
                print(f"Error during disconnect: {e}")
            finally:
                self.mail = None

    def reconnect(self) -> bool:
        """Disconnects and then reconnects to the server."""
        print("\nConnection lost or timed out. Attempting to reconnect...")
        self.disconnect() # Ensure old connection is closed
        return self.connect()

def search_messages(manager: IMAPConnectionManager, folder: str, date_criteria: str, retry=True) -> List[bytes]:
    """Search for messages, with built-in retry logic."""
    if not manager.mail:
        print(f"Cannot search, not connected.")
        return []
        
    try:
        # Select folder
        status, messages = manager.mail.select(f'"{folder}"')
        if status != 'OK':
            print(f"Error selecting folder {folder}: {messages}")
            return []
        
        # Search for messages
        search_criteria = 'ALL' if not date_criteria else date_criteria
        status, message_ids = manager.mail.search(None, search_criteria)
        if status != 'OK':
            print(f"Error searching messages: {message_ids}")
            return []
            
        msg_ids = message_ids[0].split()
        print(f"Found {len(msg_ids)} messages in {folder}")
        return msg_ids
    
    except (imaplib.IMAP4.error, ssl.SSLError, BrokenPipeError) as e:
        print(f"  Network error while searching in {folder}: {e}")
        if retry:
            if manager.reconnect():
                # Retry the operation once more
                return search_messages(manager, folder, date_criteria, retry=False)
        return [] # Failed even after retry or retry disabled
    except Exception as e:
        print(f"An unexpected error occurred while searching messages: {e}")
        return []

def download_message(manager: IMAPConnectionManager, msg_id: bytes, output_dir: str, delete_after_download=False, retry=True) -> bool:
    """Download a single message, with built-in retry logic."""
    if not manager.mail:
        print(f"Cannot download message {msg_id.decode()}, not connected.")
        return False
        
    try:
        # Fetch message
        status, msg_data = manager.mail.fetch(msg_id, '(RFC822)')
        if status != 'OK':
            print(f"Error fetching message {msg_id.decode()}: {msg_data}")
            return False
        
        # --- The rest of the function is the same, just wrapped in the try/except ---
        raw_email = msg_data[0][1]
        email_message = email.message_from_bytes(raw_email)
        
        subject = email_message.get('Subject', 'No Subject')
        date_str = email_message.get('Date', '')
        from_addr = email_message.get('From', 'unknown')
        
        safe_subject = "".join(c for c in subject if c.isalnum() or c in (' ', '-', '_')).rstrip()
        safe_subject = safe_subject[:50]
        
        filename = f"{msg_id.decode()}_{safe_subject}.eml"
        filepath = os.path.join(output_dir, filename)
        
        with open(filepath, 'wb') as f:
            f.write(raw_email)
        
        metadata = {'message_id': msg_id.decode(), 'subject': subject, 'from': from_addr, 'date': date_str, 'filename': filename}
        metadata_file = filepath.replace('.eml', '_metadata.json')
        with open(metadata_file, 'w') as f:
            json.dump(metadata, f, indent=2)
        
        if delete_after_download:
            manager.mail.store(msg_id, '+FLAGS', '\\Deleted')
            print(f"  Marked message {msg_id.decode()} for deletion")
        
        return True
    
    except (imaplib.IMAP4.error, ssl.SSLError, BrokenPipeError) as e:
        print(f"  Network error downloading message {msg_id.decode()}: {e}")
        if retry:
            if manager.reconnect():
                # We need to re-select the folder after reconnecting before we can fetch
                manager.mail.select(f'"{Path(output_dir).name}"') 
                return download_message(manager, msg_id, output_dir, delete_after_download, retry=False)
        return False # Failed even after retry or retry disabled
    except Exception as e:
        print(f"Error downloading message {msg_id.decode()}: {e}")
        return False

# --- Other helper functions (get_all_folders, expunge_deleted_messages, etc.) are modified to accept the manager ---

def get_all_folders(manager: IMAPConnectionManager):
    if not manager.mail: return []
    status, folders_raw = manager.mail.list()
    if status != 'OK': return []
    folder_list = []
    for folder in folders_raw:
        if not folder: continue
        parts = folder.decode('utf-8', 'ignore').rsplit(' ', 1)
        if len(parts) == 2:
            folder_list.append(parts[-1].strip('"'))
    print(f"Found {len(folder_list)} folders: {', '.join(folder_list)}")
    return folder_list

def expunge_deleted_messages(manager: IMAPConnectionManager, folder):
    if not manager.mail: return False
    try:
        manager.mail.select(f'"{folder}"')
        manager.mail.expunge()
        print(f"  Expunged deleted messages from {folder}")
        return True
    except Exception as e:
        print(f"  Error expunging messages from {folder}: {e}")
        return False

# --- The compression and date-range functions remain unchanged ---
def get_date_range_from_folder(folder_path: Path) -> Tuple[Optional[datetime], Optional[datetime]]:
    """Scans a folder for metadata files and returns the minimum and maximum email dates."""
    min_date, max_date = None, None
    for metadata_file in folder_path.rglob('*_metadata.json'):
        try:
            with open(metadata_file, 'r') as f:
                metadata = json.load(f)
            date_str = metadata.get('date')
            if not date_str: continue
            email_date = parsedate_to_datetime(date_str)
            if email_date:
                if min_date is None or email_date < min_date: min_date = email_date
                if max_date is None or email_date > max_date: max_date = email_date
        except Exception as e:
            print(f"  Warning: Could not parse date from {metadata_file.name}: {e}")
    return min_date, max_date

def create_multipart_zip(source_folder: Path, output_dir: Path, max_size_mb: int = 100, base_name: str = "email_archive") -> List[str]:
    """Create multi-part (split) zip files from a source folder."""
    max_size_bytes = max_size_mb * 1024 * 1024
    files_to_compress = []
    for root, dirs, files in os.walk(source_folder):
        for file in files:
            file_path = Path(root) / file
            relative_path = file_path.relative_to(source_folder)
            files_to_compress.append((str(file_path), str(relative_path), file_path.stat().st_size))
    if not files_to_compress: return []
    files_to_compress.sort(key=lambda x: x[2], reverse=True)
    zip_files, current_zip_size, current_zip_num = [], 0, 1
    current_zip_path = output_dir / f"{base_name}_part{current_zip_num:03d}.zip"
    current_zip = zipfile.ZipFile(current_zip_path, 'w', zipfile.ZIP_DEFLATED, compresslevel=6)
    print(f"Creating multi-part zip for '{source_folder.name}' (base name: '{base_name}')...")
    for file_path, relative_path, file_size in files_to_compress:
        if current_zip_size + file_size > max_size_bytes and current_zip_size > 0:
            current_zip.close()
            zip_files.append(str(current_zip_path))
            print(f"  Created {current_zip_path.name} ({current_zip_size / (1024*1024):.1f} MB)")
            current_zip_num += 1
            current_zip_path = output_dir / f"{base_name}_part{current_zip_num:03d}.zip"
            current_zip = zipfile.ZipFile(current_zip_path, 'w', zipfile.ZIP_DEFLATED, compresslevel=6)
            current_zip_size = 0
        current_zip.write(file_path, relative_path)
        current_zip_size += file_size
    if current_zip_size > 0:
        current_zip.close()
        zip_files.append(str(current_zip_path))
        print(f"  Created {current_zip_path.name} ({current_zip_size / (1024*1024):.1f} MB)")
    return zip_files

def compress_folders(archive_dir: Path, max_zip_size_mb: int, keep_uncompressed: bool = False) -> dict:
    """Compress all uncompressed folder archives into multi-part zip files with date-ranged names."""
    print(f"\nCompressing archived folders (max size: {max_zip_size_mb}MB per zip part)...")
    zip_dir = archive_dir / "compressed"
    zip_dir.mkdir(exist_ok=True)
    compression_summary = {'timestamp': datetime.now().isoformat(), 'max_zip_size_mb': max_zip_size_mb, 'folders_compressed': []}
    folder_dirs = [d for d in archive_dir.iterdir() if d.is_dir() and d.name != "compressed"]
    for folder_dir in folder_dirs:
        folder_size = sum(f.stat().st_size for f in folder_dir.glob('**/*') if f.is_file())
        print(f"\nProcessing {folder_dir.name} ({folder_size / (1024*1024):.1f} MB)...")
        min_date, max_date = get_date_range_from_folder(folder_dir)
        base_name = folder_dir.name
        if min_date and max_date:
            date_range_str = f"{min_date.strftime('%Y-%m-%d')}_to_{max_date.strftime('%Y-%m-%d')}"
            base_name = f"{date_range_str}_{folder_dir.name}"
            print(f"  Determined content date range: {min_date.date()} to {max_date.date()}")
        zip_files = create_multipart_zip(folder_dir, zip_dir, max_zip_size_mb, base_name)
        if zip_files:
            total_compressed_size = sum(Path(zf).stat().st_size for zf in zip_files)
            folder_info = {
                'folder_name': folder_dir.name,
                'original_size_mb': round(folder_size / (1024*1024), 2),
                'zip_files': [Path(zf).name for zf in zip_files],
                'zip_count': len(zip_files),
                'compressed_size_mb': round(total_compressed_size / (1024*1024), 2),
                'compression_ratio': round((1 - total_compressed_size / folder_size) * 100, 1) if folder_size > 0 else 0
            }
            compression_summary['folders_compressed'].append(folder_info)
            print(f"  Compressed to {len(zip_files)} parts ({folder_info['compression_ratio']:.1f}% reduction)")
            if not keep_uncompressed:
                print(f"  Removing original folder {folder_dir.name}")
                shutil.rmtree(folder_dir)
    with open(zip_dir / 'compression_summary.json', 'w') as f:
        json.dump(compression_summary, f, indent=2)
    return compression_summary
    
def get_date_search_criteria(start_date, end_date):
    criteria = []
    if start_date: criteria.append(f'SINCE {start_date.strftime("%d-%b-%Y")}')
    if end_date: criteria.append(f'BEFORE {end_date.strftime("%d-%b-%Y")}')
    return ' '.join(criteria)

# --- archive_messages is now refactored to use the manager ---
def archive_messages(args):
    """Main archiving function using the connection manager."""
    output_dir = Path(args.output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    
    password = args.password or getpass.getpass(f"Password for {args.username}: ")
    
    manager = IMAPConnectionManager(args.server, args.port, args.username, password, args.ssl)
    if not manager.connect():
        return False
    
    try:
        date_criteria = get_date_search_criteria(args.start_date, args.end_date)
        print(f"Date criteria: {date_criteria or 'All messages'}")
        
        folders = get_all_folders(manager) if args.all_folders else (args.folders or ['INBOX'])
        print(f"Processing folders: {', '.join(folders)}")
        
        if args.delete_messages and not args.force_delete:
            print("\nWARNING: Messages will be PERMANENTLY DELETED from the server!")
            if input("Are you sure? (yes/no): ").lower() != 'yes':
                print("Operation cancelled.")
                return False

        total_downloaded, total_errors, total_deleted = 0, 0, 0
        for folder in folders:
            print(f"\nProcessing folder: {folder}")
            msg_ids = search_messages(manager, folder, date_criteria)
            if not msg_ids: continue

            safe_folder_name = folder.replace('/', '_').replace('\\', '_')
            folder_dir = output_dir / safe_folder_name
            folder_dir.mkdir(exist_ok=True)
            
            folder_downloaded, folder_errors = 0, 0
            for i, msg_id in enumerate(msg_ids, 1):
                if args.limit and total_downloaded >= args.limit:
                    print(f"Reached overall limit of {args.limit} messages"); break
                print(f"Downloading message {i}/{len(msg_ids)} from '{folder}': {msg_id.decode()}", end='\r')
                
                if download_message(manager, msg_id, str(folder_dir), args.delete_messages):
                    total_downloaded += 1; folder_downloaded += 1
                else:
                    total_errors += 1; folder_errors += 1
            print()

            if args.delete_messages and folder_downloaded > 0:
                print(f"Expunging {folder_downloaded} deleted messages from {folder}...")
                if expunge_deleted_messages(manager, folder):
                    total_deleted += folder_downloaded
            
            print(f"Folder '{folder}' complete: {folder_downloaded} downloaded, {folder_errors} errors")
        
        print(f"\nArchiving complete! Total downloaded: {total_downloaded}, Total errors: {total_errors}")
        if args.delete_messages: print(f"Total deleted from server: {total_deleted}")
        
        if total_downloaded > 0 and args.compress:
            compress_folders(output_dir, args.max_zip_size, args.keep_uncompressed)
        
        return True
    
    finally:
        manager.disconnect()

def parse_date(date_string):
    """Parse date string to datetime object"""
    try:
        return datetime.strptime(date_string, '%Y-%m-%d')
    except ValueError:
        raise argparse.ArgumentTypeError(f"Invalid date format: {date_string}. Use YYYY-MM-DD")

def main():
    # Argument parser is unchanged, so it is collapsed for brevity
    parser = argparse.ArgumentParser(description="Download old messages from IMAP server for archival",formatter_class=argparse.RawDescriptionHelpFormatter,epilog="""Examples:\n  # Archive all folders with compression and delete from server\n  python imap_archiver.py -s mail.example.com -u user@example.com --password mypass --all-folders --delete-messages --compress\n\n  # Archive specific date range from all folders (automated deletion)\n  python imap_archiver.py -s mail.example.com -u user@example.com --password mypass --all-folders --start-date 2023-01-01 --end-date 2023-12-31 --delete-messages --force-delete""")
    parser.add_argument('-s', '--server', required=True,help='IMAP server address')
    parser.add_argument('-p', '--port', type=int, default=993,help='IMAP server port (default: 993 for SSL, 143 for non-SSL)')
    parser.add_argument('-u', '--username', required=True,help='Username for IMAP login')
    parser.add_argument('--password',help='Password for IMAP login (will prompt if not provided)')
    parser.add_argument('--no-ssl', action='store_false', dest='ssl',help='Disable SSL connection')
    parser.add_argument('--start-date', type=parse_date,help='Start date for message range (YYYY-MM-DD)')
    parser.add_argument('--end-date', type=parse_date,help='End date for message range (YYYY-MM-DD)')
    parser.add_argument('--folders', nargs='+',help='Folders to archive (e.g., INBOX "Sent Items"). If not provided and --all-folders is not used, defaults to INBOX.')
    parser.add_argument('--all-folders', action='store_true',help='Archive all folders from the server (overrides --folders)')
    parser.add_argument('--limit', type=int,help='Maximum total number of messages to download across all folders')
    parser.add_argument('--delete-messages', action='store_true',help='Delete messages from server after successful download')
    parser.add_argument('--force-delete', action='store_true',help='Skip confirmation prompt for message deletion (use with caution!)')
    parser.add_argument('-o', '--output-dir', default='email_archive',help='Output directory for archived messages (default: email_archive)')
    parser.add_argument('--compress', action='store_true',help='Compress archived folders into multi-part zip files')
    parser.add_argument('--max-zip-size', type=int, default=100,help='Maximum size for each zip file part in MB (default: 100)')
    parser.add_argument('--keep-uncompressed', action='store_true',help='Keep original uncompressed folders after compression')
    args = parser.parse_args()
    if not args.ssl and args.port == 993: args.port = 143
    print(f"Starting email archival from {args.server}:{args.port}")
    if archive_messages(args): print("\nArchival process completed successfully!")
    else: print("\nArchival process failed!"); sys.exit(1)

if __name__ == "__main__":
    main()
