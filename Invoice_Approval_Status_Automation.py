"""
AUTOMATION -- BULK INVOICE APPROVAL VALIDATION
==================================================

WHAT IT DOES (in order):
  1. Opens the actual Excel file and reads the "Invoices" table (columns
     DocumentNumber and Date, already present in your file).
  2. For each DocumentNumber, opens the approval report in SAP (a
     CLASSIC text list, not an ALV Grid -- that's why there are no
     column filters here; filtering happens BEFORE execution, on the
     selection screen).
  3. Reads every approval row for that document (there can be several --
     a document goes through several approvers in sequence).
  4. Verifies that the first row read actually belongs to the requested
     document (protection against screens left out of sync by
     unexpected popups -- see validate_correct_document()).
  5. Calculates a "score" based on those rows:
        0    -> Rejected     (any row says "Rejected")
        1    -> Authorized   (AT LEAST ONE row says "Authorized" -- that
                 is the FINAL/definitive status. "Approved" does NOT
                 count as complete on its own, because there can be
                 several intermediate approvers marked "Approved" before
                 the document reaches its final authorization)
        0.5  -> In progress  (no row is Rejected or Authorized yet --
                 there may be one or more rows in "Approved", but the
                 final authorization is still missing)
  6. Writes that score to the "Status" column in Excel, and how many
     days have passed since the invoice's "Date" until today, in a new
     "DaysElapsed" column (created automatically the first time, next to
     the table -- on later runs it's reused, never duplicated).
  7. Builds a SINGLE draft email with High priority (it is NOT sent
     automatically, you review it and click Send) if there are: Rejected
     documents, "In progress" documents that have been stuck for too
     long (more than DAYS_LIMIT_STUCK), documents not found in SAP, or
     documents whose data didn't match what was requested.
  8. If everything is fine (nothing at 0, nothing stuck), NO email is
     built -- so the area owner doesn't get empty emails every week for
     no reason.

IMPORTANT SAFETY LOCK:
  The email is always generated as a DRAFT (mail.Display()), never sent
  on its own. You decide when to click "Send" after reviewing it. Once
  you've seen several weeks of the report always coming out correct,
  you can switch to mail.Send() so it becomes fully automatic -- but
  that decision is yours, not the script's.

Requirements (one time only):
    pip install pywin32 openpyxl python-dotenv --break-system-packages

CONFIGURATION:
    Copy .env.example to .env and fill in your real values (Excel path,
    destination email, SAP node, etc.). .env is NEVER committed to the
    repo (see .gitignore) -- so you can share this script without
    exposing paths, emails, or anything specific to your company.

BEFORE RUNNING:
    - Close the Excel file if you have it open (if it's open, Windows
      won't let the script save it, and it will crash).
    - Open SAP Logon and log into your session manually.
    - Leave the session on the initial screen (the script handles the
      rest of the navigation).
"""

import os
import re
from datetime import datetime

import openpyxl
from dotenv import load_dotenv
from openpyxl.utils import range_boundaries
import win32com.client

load_dotenv()


# =============================================================================
# CONFIGURATION -- everything specific to your company/PC lives in .env,
# never hardcoded here. See .env.example for the full list.
# =============================================================================

EXCEL_PATH = os.environ["EXCEL_PATH"]
EXCEL_TABLE_NAME = os.environ.get("EXCEL_TABLE_NAME", "Invoices")

COL_HEADER_DOCUMENT = os.environ.get("COL_HEADER_DOCUMENT", "DocumentNumber")
COL_HEADER_DATE = os.environ.get("COL_HEADER_DATE", "Date")
COL_HEADER_STATUS = os.environ.get("COL_HEADER_STATUS", "Status")
COL_HEADER_DAYS = os.environ.get("COL_HEADER_DAYS", "DaysElapsed")

DAYS_LIMIT_STUCK = int(os.environ.get("DAYS_LIMIT_STUCK", 20))  # calendar days; < limit = normal "in progress" (no alert), >= limit = "stuck" (alert)
DESTINATION_EMAIL = os.environ["DESTINATION_EMAIL"]

SAVE_EVERY_N_DOCUMENTS = int(os.environ.get("SAVE_EVERY_N_DOCUMENTS", 15))  # saves the Excel file every N
                                                                              # processed rows, to avoid losing
                                                                              # all progress if something
                                                                              # crashes mid-run

SAP_FAVORITE_NODE = os.environ["SAP_FAVORITE_NODE"]  # id of the favorite/folder node in the SAP tree
                                                        # that opens the approval report

# Column positions (col, row) on the SAP screen, taken from the screen
# diagnostic. This is a CLASSIC list (not an ALV Grid), so each "cell"
# is actually a label (GuiLabel) drawn at a fixed coordinate -- that's
# why there are no interactive column filters here, unlike the ALV
# lists used in other projects.
COL_DOCUMENT = 1
COL_SEQ = 31
COL_PARTNER = 35
COL_NAME = 44
COL_STATUS = 75
COL_KEY = 86

DATA_START_ROW = 6       # the first row of actual data (everything before that is headers)
MAX_ROWS_PER_DOCUMENT = 100  # safety cap, in case a document had an abnormal number of
                               # approvers -- prevents an infinite loop


# =============================================================================
# STEP 1: SAP connection and reading
# =============================================================================

def connect_to_sap():
    """Connects to the SAP session you ALREADY have open, meaning you need to open SAP
    before running the script (it doesn't open a new one)."""
    sap_gui_auto = win32com.client.GetObject("SAPGUI")
    application = sap_gui_auto.GetScriptingEngine
    connection = application.Children(0)
    session = connection.Children(0)
    return session


def close_extra_popups(session):
    """
    If an invalid document number or something unexpected makes SAP pop
    up an error window (wnd[1], wnd[2]...), that window "steals" the
    focus and can bring down the rest of the process if we don't close
    it.

    This function checks how many windows are currently open, and
    closes every one that is NOT the main window (wnd[0]) -- from the
    last one back to the first, in case several are stacked.

    Called after each document, as "preventive cleanup", so a
    problematic document number doesn't drag garbage into the next one.
    """
    try:
        window_count = session.Children.Count
        for i in range(window_count - 1, 0, -1):  # from the last one down to wnd[1] (never wnd[0])
            try:
                session.findById(f"wnd[{i}]").Close()
            except Exception:
                pass
    except Exception:
        pass


def read_document_rows(session):
    """
    Reads, row by row, the full approval history of ONE document
    (already filtered beforehand on the selection screen). A document
    can have several rows -- one for each approver it went through.

    Stops when it finds 2 consecutive empty rows (a sign there's no more
    data), or when it hits the MAX_ROWS_PER_DOCUMENT cap (as a safety
    measure, so it can never get stuck in an infinite loop).
    """
    rows = []
    row = DATA_START_ROW
    empty_attempts = 0

    while row < DATA_START_ROW + MAX_ROWS_PER_DOCUMENT:
        try:
            document = session.findById(
                f"wnd[0]/usr/lbl[{COL_DOCUMENT},{row}]"
            ).Text.strip()
        except Exception:
            # There isn't even a control at that position anymore -- the
            # list has run out of rows, so we stop here.
            break

        if document == "":
            # Blank row -- could be a gap between sections, not
            # necessarily the end. We give it ONE more chance before
            # giving up.
            empty_attempts += 1
            if empty_attempts > 1:
                break
            row += 1
            continue

        empty_attempts = 0  # reset as soon as we find data again

        def read_cell(col):
            """Reads the text of a specific cell in this same row, with error handling."""
            try:
                return session.findById(f"wnd[0]/usr/lbl[{col},{row}]").Text.strip()
            except Exception:
                return ""

        rows.append({
            "document": document,
            "seq": read_cell(COL_SEQ),
            "partner": read_cell(COL_PARTNER),
            "name": read_cell(COL_NAME),
            "status": read_cell(COL_STATUS),
            "key": read_cell(COL_KEY),
        })
        row += 1

    return rows


def validate_correct_document(rows, expected_document):
    """
    Confirms that the first row read actually belongs to the requested
    document. Protection against the case where an unexpected popup left
    the screen out of sync -- if that happens, Python could
    unknowingly read the rows still VISIBLE from the previous run,
    calculating a score for the wrong document without raising any
    error.

    Returns True if it matches (or if the row list is empty, which is
    already handled elsewhere as "not found"), False if there's a real
    mismatch.
    """
    if not rows:
        return True  # empty list = document not found, not a mismatch

    return rows[0]["document"] == expected_document


def calculate_score(document_rows):
    """
    Translates the list of approval rows into a single number.

    IMPORTANT: "Approved" is NOT the final status -- a document can go
    through SEVERAL intermediate approvers whose row is marked
    "Approved" before it reaches final authorization. The status that
    actually closes the document is "Authorized". So:

      None -> the document has no rows at all (not found in SAP)
      0    -> any row says "Rejected" (wins over any other status)
      1    -> AT LEAST ONE row says "Authorized" (that's the definitive
               close, no matter how many rows only say "Approved")
      0.5  -> any other case (for example, one or more rows in
               "Approved" but none in "Authorized" yet -- still in
               progress, not closed)
    """
    if not document_rows:
        return None

    statuses = [r["status"].lower() for r in document_rows]

    if any(s == "rejected" for s in statuses):
        return 0
    if any(s == "authorized" for s in statuses):
        return 1
    return 0.5


# =============================================================================
# STEP 2: Read/prepare the Excel file
# =============================================================================

def open_excel_table():
    """
    Opens the Excel file and locates the "Invoices" table inside it
    (regardless of which tab it's on). Returns the workbook, the sheet,
    the table object, and the boundaries (rows/columns) it occupies.
    """
    wb = openpyxl.load_workbook(EXCEL_PATH)

    sheet = None
    table = None
    for s in wb.worksheets:
        if EXCEL_TABLE_NAME in s.tables:
            sheet = s
            table = s.tables[EXCEL_TABLE_NAME]
            break

    if table is None:
        raise ValueError(f"Table '{EXCEL_TABLE_NAME}' not found in {EXCEL_PATH}")

    min_col, min_row, max_col, max_row = range_boundaries(table.ref)
    return wb, sheet, table, (min_col, min_row, max_col, max_row)


def locate_days_column(sheet, headers, min_row, max_col):
    """
    Looks for the "DaysElapsed" column in two ways, in this order:

      1. By NAME inside the official table (the `headers` dict, already
         built in main() from the columns Excel recognizes as part of
         the "Invoices" table). This covers the case where you absorbed
         the column into the table manually from Excel (Table Design >
         Resize Table).

      2. If it's not there, falls back to the previous behavior: look
         for (or create) the column right next to the table's current
         border, outside the Excel Table object.

    Without this double check, if you ever extend the table to include
    this column, the script would think "the column next to it" moved
    further right, and would create a DUPLICATE DaysElapsed column
    instead of reusing the one you already have.
    """
    if COL_HEADER_DAYS in headers:
        return headers[COL_HEADER_DAYS]  # already an official part of the table

    candidate_column = max_col + 1
    current_text = sheet.cell(row=min_row, column=candidate_column).value

    if current_text == COL_HEADER_DAYS:
        return candidate_column  # already existed outside the table, we reuse it

    # Didn't exist anywhere -- create it for the first time, outside the table
    sheet.cell(row=min_row, column=candidate_column).value = COL_HEADER_DAYS
    return candidate_column


# =============================================================================
# STEP 3: Main process
# =============================================================================

def main():
    wb, sheet, table, bounds = open_excel_table()
    min_col, min_row, max_col, max_row = bounds

    # Build a dictionary {column_name: column_number}, so we don't have
    # to remember fixed positions -- we search by the header text, so if
    # you ever reorder columns in Excel, the script still finds them
    # correctly.
    headers = {}
    for col in range(min_col, max_col + 1):
        text = sheet.cell(row=min_row, column=col).value
        if text:
            headers[text] = col

    col_document = headers[COL_HEADER_DOCUMENT]
    col_date = headers[COL_HEADER_DATE]
    col_status = headers[COL_HEADER_STATUS]
    col_days = locate_days_column(sheet, headers, min_row, max_col)

    session = connect_to_sap()
    session.findById("wnd[0]").maximize()
    session.findById(
        "wnd[0]/usr/cntlIMAGE_CONTAINER/shellcont/shell/shellcont[0]/shell"
    ).doubleClickNode(SAP_FAVORITE_NODE)

    # Here we collect everything that needs to go in the alert email.
    rejected = []
    stuck = []
    not_found = []

    documents_processed_since_last_save = 0
    total_rows = max_row - min_row  # how many data rows there are in total (excluding header)

    for index, excel_row in enumerate(range(min_row + 1, max_row + 1), start=1):
        document_raw = sheet.cell(row=excel_row, column=col_document).value
        if document_raw in (None, ""):
            continue  # empty row in the table, just skip it

        # Excel sometimes stores numbers like "4107023215.0" (float)
        # instead of plain text -- this line strips the trailing ".0" if
        # present.
        document = re.sub(r"\.0$", "", str(document_raw).strip())

        print(f"[{index}/{total_rows}] {document} ...", end=" ")

        # Idempotency lock: if this document was ALREADY authorized
        # (status == 1) in a previous run, we don't query SAP again for
        # it -- this is what keeps the script fast even as the invoice
        # list grows month over month, instead of having to review the
        # ENTIRE history every time.
        current_status = sheet.cell(row=excel_row, column=col_status).value
        if current_status == 1:
            print("already authorized in a previous run, skipping")
            continue

        invoice_date = sheet.cell(row=excel_row, column=col_date).value

        # try/except around the ENTIRE processing of this document: if
        # anything fails (SAP freezes, a weird popup shows up that we
        # can't close, whatever), we log it as an error and move on to
        # the NEXT document -- one problematic document number no longer
        # brings down the whole run.
        try:
            session.findById("wnd[0]/usr/ctxtS_BELNR-LOW").text = document
            session.findById("wnd[0]").sendVKey(8)

            rows = read_document_rows(session)

            if not validate_correct_document(rows, document):
                # The first row read doesn't correspond to the requested
                # document -- likely a popup left the screen out of
                # sync. We don't trust this data: it's flagged for
                # manual review and its previous Status is NOT touched.
                print("⚠️  data doesn't match the requested document, flagging for manual review")
                not_found.append(document)
            else:
                score = calculate_score(rows)

                days = None
                if isinstance(invoice_date, datetime):
                    days = (datetime.now() - invoice_date).days

                sheet.cell(row=excel_row, column=col_status).value = score
                sheet.cell(row=excel_row, column=col_days).value = days

                if score is None:
                    not_found.append(document)
                elif score == 0:
                    rejected.append({"document": document, "rows": rows})
                elif score == 0.5:
                    if days is not None and days >= DAYS_LIMIT_STUCK:
                        stuck.append({
                            "document": document,
                            "days": days,
                            "rows": rows,
                        })
                # score == 1 -> document authorized, nothing else to do with it

                print(f"score={score}  days={days}")

        except Exception as e:
            print(f"⚠️  error while processing -- {e}")

        finally:
            # No matter what happens (success or error), try to return
            # to a clean selection screen for the next document.
            close_extra_popups(session)
            try:
                session.findById("wnd[0]/tbar[0]/btn[3]").press()
            except Exception:
                pass

        documents_processed_since_last_save += 1
        if documents_processed_since_last_save >= SAVE_EVERY_N_DOCUMENTS:
            try:
                wb.save(EXCEL_PATH)
                print(f"  (progress saved after {documents_processed_since_last_save} documents)")
            except PermissionError:
                print("  ⚠️  Could not save progress -- the Excel file is still open in another program.")
            documents_processed_since_last_save = 0

    # Final save, once all processing is done
    try:
        wb.save(EXCEL_PATH)
    except PermissionError:
        print("\n⚠️  Could not save the final Excel file -- close it if it's open and run again.")

    if rejected or stuck or not_found:
        prepare_alert_email(rejected, stuck, not_found)
        print(f"\n✅ Email draft prepared. "
              f"Rejected: {len(rejected)}, Stuck: {len(stuck)}, "
              f"Not found: {len(not_found)}")
    else:
        print("\nNo irregularities. No email was prepared.")

    print(f"Excel file updated: {EXCEL_PATH}")

# =============================================================================
# EMAIL SENDING -- email delivery configuration
# =============================================================================


def prepare_alert_email(rejected, stuck, not_found):
    """
    Builds the High priority alert email and leaves it on screen as a
    DRAFT (mail.Display()), never sending it on its own. Once you've had
    several runs in a row that you trust, switch .Display() to .Send()
    below to make it fully automatic.
    """
    outlook = win32com.client.Dispatch("Outlook.Application")
    mail = outlook.CreateItem(0)  # 0 = new blank email
    mail.To = DESTINATION_EMAIL
    mail.Subject = f"🚨 INVOICE FOLLOW-UP ALERT 🚨 - {datetime.now().strftime('%d/%m/%Y')}"
    mail.Importance = 2  # 0 = Low, 1 = Normal, 2 = High (olImportanceHigh)

    body = "Automatic Validation Results:\n\n"

    if rejected:
        body += f"REJECTED DOCUMENTS ({len(rejected)}):\n"
        for r in rejected:
            last_row = r["rows"][-1]
            body += (
                f"  - {r['document']} | Rejection detected at: "
                f"{last_row['name']}\n"
            )
        body += "\n"

    if stuck:
        body += f"STUCK DOCUMENTS (more than {DAYS_LIMIT_STUCK} days) ({len(stuck)}):\n"
        for s in stuck:
            body += f"  - {s['document']} | {s['days']} days unresolved\n"
        body += "\n"

    if not_found:
        body += f"NOT FOUND / PENDING MANUAL REVIEW ({len(not_found)}):\n"
        for n in not_found:
            body += f"  - {n}\n"

    mail.Body = body

    # ---------------------------------------------------------------
    # SEND SWITCH -- keep ONLY ONE of these two lines active:
    #
    #   mail.Display()  ->  opens the email for
    #                       you to review and click Send.
    #   mail.Send()      -> sends it automatically, without anyone
    #                       reviewing it. Only enable this after
    #                       confirming several correct runs in a row.
    # ---------------------------------------------------------------
    mail.Display()
    # mail.Send()


if __name__ == "__main__":
    main()
