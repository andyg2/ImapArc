# IMAP Archiver

A powerful and flexible command-line tool to archive emails from an IMAP server, optionally compressing the results and deleting messages after successful download.

Supports all standard IMAP servers and securely handles SSL connections, date-based filtering, multi-folder operations, and multi-part ZIP file compression.

---

## Features

✅ Archive emails from one or more IMAP folders  
✅ Filter messages by date range  
✅ Save emails in `.eml` format with metadata  
✅ Optional permanent deletion of downloaded messages  
✅ Compress archive into multi-part ZIP files  
✅ Generate summary reports for archiving and compression  
✅ Compatible with all IMAP servers and email providers

---

## Requirements

- Python 3.6+
- Works on Windows, Linux, and macOS
- No external dependencies

---

## Installation

Clone the repository:

```bash
git clone https://github.com/andyg2/ImapArc.git
cd ImapArc
````

Run the script directly with Python:

```bash
python imap_archiver.py --help
```

---

## Usage

```bash
python imap_archiver.py -s <imap_server> -u <email> --password <password> [OPTIONS...]
```

### Required Parameters

| Flag               | Description                                  |
| ------------------ | -------------------------------------------- |
| `-s`, `--server`   | IMAP server address (e.g., `imap.gmail.com`) |
| `-u`, `--username` | Email/username to log in                     |
| `--password`       | Password for the account (use with care)     |

---

## Common Options

| Option                    | Description                                     |
| ------------------------- | ----------------------------------------------- |
| `--folders FOLDER...`     | List of folders to archive (default: INBOX)     |
| `--all-folders`           | Archive all available folders on the server     |
| `--start-date YYYY-MM-DD` | Start date for message filtering                |
| `--end-date YYYY-MM-DD`   | End date for message filtering                  |
| `--limit N`               | Limit number of messages per folder             |
| `--delete-messages`       | Permanently delete messages after download      |
| `--force-delete`          | Skip confirmation prompt when deleting          |
| `--compress`              | Create compressed ZIP archive(s)                |
| `--max-zip-size MB`       | Max size for each ZIP file (default: 100MB)     |
| `--keep-uncompressed`     | Keep original extracted files after compression |
| `-o`, `--output-dir DIR`  | Output directory (default: `email_archive`)     |
| `--no-ssl`                | Disable SSL (not recommended)                   |

---

## Examples

### 🔐 Download all emails from INBOX (Default)

```bash
python imap_archiver.py -s imap.example.com -u user@example.com --password secret
```

### 📆 Archive emails between specific dates (Inclusive)

```bash
python imap_archiver.py -s imap.example.com -u user@example.com --password secret \
  --start-date 2023-01-01 --end-date 2023-12-31
```

### 📁 Archive specific folders, space separated list, use quotes for folders with spaces

```bash
python imap_archiver.py -s imap.example.com -u user@example.com --password secret \
  --folders INBOX Sent "Archive 2023"
```

### 🧹 Download and delete emails after archiving (with confirmation)

```bash
python imap_archiver.py -s imap.example.com -u user@example.com --password secret \
  --all-folders --delete-messages
```

### ☠️ Delete without prompt `--force-delete` (use with caution)

```bash
python imap_archiver.py -s imap.example.com -u user@example.com --password secret \
  --all-folders --delete-messages --force-delete
```

### 📦 Compress archive to multi-part ZIP files (100MB max each)

```bash
python imap_archiver.py -s imap.example.com -u user@example.com --password secret \
  --all-folders --compress --max-zip-size 100
```

### 🛡️ Safe backup (no deletions, keeps uncompressed folders)

```bash
python imap_archiver.py -s imap.example.com -u user@example.com --password secret \
  --all-folders --compress --keep-uncompressed
```

---

## Output Structure

After a run, the output directory contains:

* `.eml` email files
* `*_metadata.json` per-email metadata
* `archive_summary.json`: summary of downloads
* `compressed/`: ZIP files (if compression enabled)
* `compression_summary.json`: compression stats

---

## Security Note

**Never hardcode passwords in scripts.** Use environment variables or prompt securely if integrating with automation.

For example:

```bash
python imap_archiver.py -s imap.example.com -u user@example.com --password "$EMAIL_PASS"
```

---

## License

MIT License.
Feel free to fork and contribute!

---

## Contributions

Pull requests and issues welcome. Feature ideas include:

* Progress bar
* Attachment extraction
* Retry on failure
* IOmproved Logging

---

## Author

**Andy Gee**

> Built to preserve and organize your digital correspondence reliably and efficiently.
