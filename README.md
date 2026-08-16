# Outlook Bulk Mailer

A Python tool for sending personalised, templated emails through Outlook at scale, with multi-account round-robin, batch rate limiting, and CSV-based status tracking so you never send the same email twice.

Built to run outreach campaigns from a CSV or Excel database of contacts, using an Outlook `.oft` template with placeholder tags that get replaced per recipient.

## Features

- **Multiple sender accounts** — round-robins across as many Outlook accounts as you give it, each with its own daily send cap
- **Batching with cooldowns** — sends in configurable batch sizes with a pause between batches, plus a delay between individual emails, to stay well under Outlook/Exchange sending limits
- **Dynamic personalisation** — swaps `[First name]`, `[Institution Name]`, `[sender name]`, and `[sender mail]` placeholders in the template body, with a fallback for very short first names (e.g. "Dear [First name]," becomes "Greetings," if the name is 3 characters or fewer)
- **CSV status tracking** — after every successful send, the source CSV is updated with a status, the send date, and which account it went out from, so reruns automatically skip anything already sent
- **Pause on demand** — drop a `pause.flag` file in the project folder to pause mid-run, delete it to resume
- **Dry run mode** — open drafts instead of sending, to review before going live
- **Interactive setup** — `main.py` walks you through sender accounts, batch settings, and column mapping at runtime, with sane defaults you can just press Enter to accept

## Requirements

- Windows, with Classic Outlook installed and signed in to the account(s) you want to send from
- Python 3.8+
- `pywin32`

```bash
pip install pywin32
```

Optional, only needed if you plan to load recipients from an Excel file:

```bash
pip install openpyxl
```

## Project structure

```
.
├── main.py                 # interactive entry point — run this
├── send_outlook_emails.py  # core sending engine (Outlook COM automation)
├── load_recipients.py      # loads recipients from CSV / SQLite / Excel
└── pause.flag              # (optional) create this file to pause a run
```

## Setup

1. Have your recipient database as a CSV with, at minimum, a name column and an email column. A status column is used to track what's already been sent — leave it blank for anyone who hasn't been contacted yet.
2. Create your email template in Outlook and save it as an `.oft` file. In the body, use these placeholders wherever you want personalisation:
   - `[First name]`
   - `[Institution Name]`
   - `[sender name]`
   - `[sender mail]`
3. Update the fixed paths at the top of `main.py`:

```python
TEMPLATE_PATH = r"D:\path\to\your\template.oft"
DATABASE_PATH = r"D:\path\to\your\contacts.csv"
```

4. Adjust the defaults if you like (max emails per account, batch size, batch interval, delay between sends inside a batch).

## Usage

```bash
python main.py
```

You'll be prompted for:

- Which account(s) to send from (comma-separated), and a display name for each
- Max emails per account, batch size, and batch interval
- Which CSV columns to use for recipient name, email, and status

The tool then loads unsent recipients, shows you a preview of the first email (with placeholders resolved), and asks for confirmation before sending anything.

To pause a run in progress, create an empty file named `pause.flag` in the project folder. Delete it to resume.

## How recipient loading works

`load_from_csv` only picks up rows where the status column is completely blank — anything with a status already set is treated as already handled. It stops reading as soon as it hits a row with no name, so trailing blank rows in the sheet won't cause issues.

`load_from_sqlite` and `load_from_excel` are also included if your contact list lives somewhere other than a CSV, though `main.py` is currently wired up for CSV only.

## Notes

- This automates Outlook directly through COM, so Outlook needs to be installed and running under the sending account(s) — it isn't using the Microsoft Graph API or SMTP.
- Sending caps and batch timing are there to reduce the risk of being rate-limited or flagged by your mail provider. Tune them conservatively for your organisation's limits.
- Failed sends are logged to the console and skipped rather than stopping the whole run.

## License

MIT — use and adapt freely.
