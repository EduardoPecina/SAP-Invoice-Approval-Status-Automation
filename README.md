# SAP Invoice Approval Status Automation

Automates SAP invoice approval status validation into a shared Excel tracker, with alert-email drafting via Outlook COM automation.

> **Note:** this is a sanitized portfolio copy. All company-specific identifiers (paths, emails, SAP node) have been replaced with environment variables or generic placeholders — see `.env.example`. No real business data is included anywhere in this repository.

## Problem

An analyst manually checked, document by document, whether each invoice had already been authorized in SAP, copying the status by hand into a control Excel file and flagging by email whenever something stayed "stuck" without being resolved.

This script automates the entire flow end to end.

## What it does

1. Opens the actual Excel file and reads the `Invoices` table (columns `DocumentNumber` and `Date`).
2. For each document number, opens the approval report in SAP (a classic text list, not an ALV Grid) and reads all approval rows for that document — there can be several, one per approver in the chain.
3. Validates that the rows read actually correspond to the requested document, as protection against screens left out of sync by unexpected popups.
4. Calculates a score:
   - `0` → Rejected
   - `1` → Authorized (final status)
   - `0.5` → In progress (may have intermediate approvals, but not the final authorization)
5. Writes the score to the `Status` column and the days elapsed since the invoice date into `DaysElapsed`.
6. If there are rejected, stuck (more than N days in progress), or not-found documents, it builds **one draft email** with high priority — never sent automatically, opened for manual review before clicking Send.
7. If everything is in order, no email is generated.

## Safety lock

The email is always generated as a draft (`mail.Display()`), never sent automatically. Switching to automatic sending (`mail.Send()`) is a manual decision made directly in the code, once the report has proven reliable across several runs.

## Requirements

- Windows, with SAP GUI Scripting enabled
- An SAP session already open and logged in (the script does not open a new one)
- Outlook installed (for the email draft)
- Python 3.9+

```bash
pip install -r requirements.txt
```

## Configuration

1. Copy `.env.example` to `.env`.
2. Fill in your real values: Excel path, column names, destination email, SAP node, etc.
3. The `.env` file is never committed to the repo (it's in `.gitignore`).

## Usage

1. Close the Excel file if you have it open.
2. Open SAP Logon, log into your session, and leave it on the initial screen.
3. Run the script:

```bash
python invoice_approval_status_automation.py
```

## Structure

```
.
├── invoice_approval_status_automation.py   # main script
├── .env.example                            # configuration template
├── requirements.txt
└── .gitignore
```

## Technical notes

- SAP reading uses `GuiLabel` controls at fixed row/column coordinates (classic list), not an ALV grid — that's why filtering happens on the selection screen, before execution.
- The script is idempotent: if a document was already `Authorized` in a previous run, it's not queried again in SAP.
- Progress is saved every N documents (configurable) so a mid-run failure doesn't lose all progress.
