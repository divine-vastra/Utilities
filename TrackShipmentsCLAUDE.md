# TrackShipments â€” Context & Memory

Context for future Claude sessions working on `TrackShipments.py` in `c:\DV\Utilities`.
Companion doc: [UpdateShipmentsCLAUDE.md](UpdateShipmentsCLAUDE.md) (the script that appends new orders to the same sheet).

## Purpose
Polls the tracking page of every shipment in the Google Sheet, writes the latest status back, sends WhatsApp template messages on key status changes, and moves delivered rows to a "Delivered Earlier" tab. Loops forever, one full run every `POLL_SECONDS` (3600 = 60 min).

Pipeline: `Updateshipments.py` appends new dispatched orders â†’ `TrackShipments.py` tracks them, notifies customers and archives delivered ones.

## Run
```powershell
cd c:\DV\Utilities
pip install gspread openpyxl playwright   # then: playwright install chromium
python TrackShipments.py                  # runs forever (poll loop)
```
One-off run, or a dry run without editing the file:
```powershell
python -c "import asyncio, TrackShipments as t; t.DRY_RUN=True; asyncio.run(t.run())"
```
Importing the module does not start the loop (guarded by `__main__`). Dry runs still write scraped status columns to the sheet (see Gotchas).

## Configuration (top of file)
| Setting | Value / meaning |
|---|---|
| `BACKEND` | `"gsheet"` (default) or `"excel"` (legacy local `shipments.xlsx`, kept as a switch) |
| `SHEET_ID` / `SHEET_TAB` | `1mFvlvN0wIF9sTIPJrldAoyBfolS1rodzUqG_gDtG1hQ` / `Sheet1` |
| Google credentials | Priority: env `GOOGLE_SERVICE_ACCOUNT_JSON` (the full key JSON as text, for GitHub Actions: store it as a repo secret and map it to this env var), else the key file `SERVICE_ACCOUNT_FILE` (env `GOOGLE_SERVICE_ACCOUNT_FILE`, default `service_account.json` next to the script). The sheet must be shared (Editor) with the key's `client_email`. |
| `DRY_RUN` | Read from env var / `.env` key `DRY_RUN` (`1`/`true`/`yes`/`on` = on; unset or anything else = off). On = log what would be sent/moved; nothing is sent, nothing archived |
| `.env` | Loaded from the script's folder by `load_env_file` (no python-dotenv). Needs `WA_TOKEN` and `WA_PHONE_NUMBER_ID`. Real env vars take priority. |
| WhatsApp templates | `order_dispatched`, `out_for_delivery` (`WA_TEMPLATE_NAME`), `order_delivered`; language `en`; `WA_FIXED_CP_ID = "105"` |
| `OFD_STATUS` / `DELIVERED_STATUS` | `"out for delivery"` / `"delivered"` (compared case-insensitively) |
| `ARCHIVE_SHEET` | `Delivered Earlier` |
| `NCOLS` | 19: columns A..S are loaded into memory |

## Sheet column map (1-based; shared with Updateshipments.py)
| Col | Name | Written by |
|---|---|---|
| 1 A | Customer Name | Updateshipments |
| 2 B | AWB | Updateshipments |
| 3 C | Courier Partner (header "CP_ID") | Updateshipments |
| 4 D | Status | **Track** |
| 5 E | Last scan location | **Track** |
| 6 F | Last scan remark | **Track** |
| 7 G | Time of last scan | **Track** |
| 8 H | When the script last scraped the row | **Track** |
| 9 I | EDD (only overwritten when the scrape returns one) | **Track** |
| 10 J | Track URL | Updateshipments |
| 11 K | Dispatch date | Updateshipments |
| 12 L | Destination | Updateshipments |
| 13 M | Order type (prepaid/COD) | Updateshipments |
| 14 N | Invoice value | Updateshipments |
| 15 O | Customer phone (10 digits; `91` is prefixed) | Updateshipments |
| 16 P | unused | â€” |
| 17 Q | WA Dispatched Sent timestamp | **Track** |
| 18 R | WA OFD Sent timestamp | **Track** |
| 19 S | WA Delivered Sent timestamp | **Track** |

Headers for Q/R/S are created if blank. Changing the layout means updating the `COL_*` constants here **and** `Updateshipments.py`.

## How a run works
1. `GSheetStore()` / `ExcelStore()` opens the backend. `store.load()` reads the whole sheet into an in-memory grid (padded to 19 columns, raw values, so phones/AWBs come back as numbers).
2. For each row with an AWB and courier partner, `scrape_shipment` opens the track URL in headless Chromium (Playwright). It tries network-response interception first (`parse_intercepted_json`) and falls back to HTML selectors (`scrape_via_html`). It sleeps 2 s between rows.
3. Changes go through `setc(r, c, value)`, which updates the grid and records the cell in `dirty`.
4. **WhatsApp events**, each sent at most once per shipment (guarded by the Q/R/S timestamp):
   - old status blank â†’ non-blank: `order_dispatched` (tracking-link button)
   - status becomes "Out for Delivery": `out_for_delivery` (tracking-link button)
   - status becomes "Delivered": `order_delivered` (no components)
   - Button value is `{awb}&cp_id=105`. An invalid or missing phone logs a warning and retries on the next poll.
5. Delivered rows are collected for archiving **regardless of whether the WhatsApp send succeeded** (no retry).
6. `finally:` always flushes `dirty` to the backend (also on Ctrl+C or a crash).
7. Archive step: `store.archive(rows, awbs)`, skipped in dry run.
8. Prints a summary with the number updated, the number of WhatsApp messages sent and the time taken.

## Storage layer
- `ExcelStore` (legacy): saves after **every** WhatsApp send (`save_after_send = True`). Archives with `archive_rows` (openpyxl, copies styles).
- `GSheetStore`: writes **once at the end** (`save_after_send = False`).
  - `flush` groups dirty cells into contiguous per-row ranges and does a `batch_update` with `RAW` values, in chunks of 400.
  - `archive` re-reads the sheet and re-locates delivered rows **by AWB** (the sheet may have changed during the scrape). It copies each row with its formatting (`copyPaste`) to `Delivered Earlier` and `deleteDimension`s the originals, all in one atomic `batch_update`. The tab and its header row are created if missing. Deleting (not clearing) rows avoids blank gaps, which matters because `Updateshipments.py` appends below the last row.
  - `_retry` waits and retries on 429/500/503 (30 s Ã— attempt).

## History / decisions
- Originally ran on a local `shipments.xlsx`. Migrated to the Google Sheet (same sheet as Updateshipments) on 2026-10-03 and kept an `excel` backend switch.
- User chose **batch write at the end** (accepting the risk that a hard kill after a WhatsApp send loses the sent-timestamp and may re-send). A `finally` flush mitigates normal crashes and Ctrl+C.
- Tested on 2026-10-03:
  - Dry run over 37 shipments in about 5.7 min: scraping, WhatsApp-event detection and phone normalisation all worked.
  - Live archive test: 4 delivered rows were copied to `Delivered Earlier` and deleted from `Sheet1`, with no blank rows left.
- Credentials: `GSheetStore` now uses `GOOGLE_SERVICE_ACCOUNT_JSON` (key content) when set, otherwise the local key file, so it works in GitHub Actions without a file. `Updateshipments.py` still uses the file only.
- `DRY_RUN` is no longer hard-coded; it is read from the `DRY_RUN` env var or `.env` (default off).
- Scraped `location` now has newlines and repeated whitespace collapsed to single spaces before it is written (applied right after the scrape in `run()`). Remark/status are unchanged.
- Real WhatsApp sending through the Google Sheet backend has **not** been run end to end.

## Known gotchas
- **Dry runs are not read-only:** status columns Dâ€“I are written. Because the old status then equals the new one, the next live run will **not** fire WhatsApp for those transitions. To re-trigger, set the status cell back to the previous value.
- Not writing per send on the gsheet backend means a hard kill can duplicate WhatsApp messages. If that becomes a problem, flush the Q/R/S cells after each send (cost: more API calls and quota).
- Writes use row numbers read at the start of the run. If someone inserts or deletes rows in `Sheet1` mid-run, the writes can land on the wrong rows. Appending (Updateshipments) is safe. Archiving re-locates by AWB.
- The HTML fallback relies on CSS selectors for the ClickPost page. If statuses stop updating, check `selectors_to_try` first. The network-intercept path currently usually returns empty, so HTML is the common path.
- The "Timed out waiting for status element" warning is non-fatal; the HTML scrape still usually succeeds.
- Archive column widths are not copied to a newly created archive tab.
- `track_url` is trusted as stored. Updateshipments rewrites `shopdeck.clickpost.ai` to `track.clickpost.in`.
- `NCOLS = 19` is the columns loaded. Adding columns beyond S needs a bump.

## Conventions
- Match the existing style: plain script, banner comments, `ERROR:` prefix with `sys.exit(1)` for fatal setup problems.
- Be careful with anything that touches the live sheet or WhatsApp: a real run messages real customers. Prefer a dry run or calling store methods directly for tests, and tell the user what will change.
- Keep this file and `UpdateShipmentsCLAUDE.md` updated with every change to either script (add a line to History and fix any affected sections).
