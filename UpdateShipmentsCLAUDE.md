# UpdateShipments — Context & Memory

Context for future Claude sessions working on `Updateshipments.py` in `c:\DV\Utilities`.

## Purpose
Takes the daily courier dispatch report (`Dispatched_Order_Report.xlsx`) and appends new orders to the shipment-tracking **Google Sheet**. Previously the target was a local `Shipments.xlsx`; it was moved to Google Sheets (that local file and `Shipments - Copy.xlsx` are now legacy).

## Run
```powershell
cd c:\DV\Utilities
pip install -r requirements.txt
python Updateshipments.py
```
Run from `c:\DV\Utilities` because file paths are relative to the working directory.

## Inputs / targets
| Item | Value |
|---|---|
| Source file | `Dispatched_Order_Report.xlsx`, sheet `AWB wise Details`, header in the first row |
| Target sheet ID | `1mFvlvN0wIF9sTIPJrldAoyBfolS1rodzUqG_gDtG1hQ` |
| Target tab | `Sheet1` (`SHEET_TAB`) |
| Auth | Service-account key at `service_account.json`, or the path in env var `GOOGLE_SERVICE_ACCOUNT_FILE` |

The target sheet must be shared (Editor) with the `client_email` inside `service_account.json`. The key file is git-ignored and must never be committed.

## Source columns used
`Order Dispatched Date`, `AWB NO.`, `Invoice Value`, `Amount to Collect`, `Customer Name`, `Customer Number`, `Customer Address`, `Track Link`, `Courier Partner`.
Rows with no AWB are dropped, and duplicate AWBs within the report are dropped.

## Output row layout (15 columns, 1-indexed)
| Col | Content |
|---|---|
| 1 | Customer Name |
| 2 | AWB (numeric via `to_num`; stays text if it has letters, e.g. DTDC `7D137810927`) |
| 3 | Courier Partner (stored in the "CP_ID" column) |
| 4–9 | Blank (filled by `TrackShipments.py`: status, location, remark, last-scan time, last-checked time, EDD) |
| 10 | Track Link, with `https://shopdeck.clickpost.ai/` replaced by `https://track.clickpost.in/` |
| 11 | Dispatch date as `YYYY-MM-DD` (source format `%d/%m/%Y  %H:%M`, two spaces before the time) |
| 12 | Address: last two comma-separated parts, minus the last 9 characters (intended to strip the pincode) |
| 13 | `prepaid` if Amount to Collect is 0, otherwise `COD` |
| 14 | Invoice Value (numeric) |
| 15 | Customer Number (numeric) |

## Behavior
1. Validate that the source file and the service-account file exist.
2. Read and clean the source report and build the rows.
3. Open the sheet with `gspread.service_account(...)`.
4. **Skip AWBs already in the sheet.** It reads column 2 with `UNFORMATTED_VALUE` and compares as text after `to_num`, which makes reruns idempotent.
5. `append_rows(..., value_input_option="USER_ENTERED", table_range="A1")` appends below the existing data. `USER_ENTERED` makes Sheets parse the ISO date.
6. Print a skipped count, an appended count, and the appended rows.

## History / decisions
- Switched the target from local `Shipments.xlsx` (openpyxl, first blank row in col A) to Google Sheets (gspread), at the user's request.
- Added the duplicate-AWB skip, because reruns on the same report previously double-appended.
- Tested: the first run appended 11 rows (dispatch date 2026-10-01); the second run skipped all 11 and appended 0.
- The old `extract_cp_id` helper is commented out. The courier partner name is now used instead of the `cp_id` from the track link.
- `.gitignore` covers `*.xlsx`, `.env`, `__pycache__` and `service_account.json`.

## Known limitations / gotchas
- `datetime.strptime` on the dispatch date crashes if a date is missing or in a different format.
- `last_two_tokens` removes the last 9 characters of any result longer than 9 characters. That is wrong for addresses that don't end in a pincode-like suffix.
- Dedup is by AWB only, so an AWB already in the sheet is never updated, even if the report has changed data for it.
- Only the AWB column (col 2) is read for dedup. If someone changes the column order in the sheet, update the `col_values(2, ...)` call and the row layout.
- If the tab isn't called `Sheet1`, the script fails with `WorksheetNotFound`; change `SHEET_TAB`.
- `TrackShipments.py` reads and writes the same sheet and depends on this column layout (cols 4–9 and 17–19 are its columns). Check it before changing the layout. Its archiving **deletes** delivered rows from `Sheet1`, so no blank rows are left; keep appending with `append_rows` (never write to a fixed row number).

## Related files in `c:\DV\Utilities`
- `TrackShipments.py` — tracks every shipment in the same Google Sheet, writes status columns, sends WhatsApp messages and archives delivered rows to the "Delivered Earlier" tab. Documented in [TrackShipmentsCLAUDE.md](TrackShipmentsCLAUDE.md).
- `whatsapp_CloudAPISender.py` — WhatsApp Cloud API sender.
- `.env` — secrets (git-ignored).

## Conventions
- Match the existing style: plain script, section banner comments, `sys.exit(1)` with an `ERROR:` message on failures.
- Keep the user informed of anything that writes to the live Google Sheet. Running the script appends real data.
