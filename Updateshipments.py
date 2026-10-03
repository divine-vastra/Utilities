import pandas as pd
import gspread
from datetime import date, datetime
import re
import sys
import os

# ── Paths ────────────────────────────────────────────────────────────────────
DISPATCHED_FILE = "Dispatched_Order_Report.xlsx"

# Target Google Sheet (ID from the sheet URL) and the tab to append to
SHEET_ID  = "1mFvlvN0wIF9sTIPJrldAoyBfolS1rodzUqG_gDtG1hQ"
SHEET_TAB = "Sheet1"
# Service-account key JSON; the sheet must be shared (Editor) with its client_email
SERVICE_ACCOUNT_FILE = os.environ.get("GOOGLE_SERVICE_ACCOUNT_FILE", "service_account.json")

# ── Validate files exist ──────────────────────────────────────────────────────
missing = [f for f in [DISPATCHED_FILE, SERVICE_ACCOUNT_FILE] if not os.path.exists(f)]
if missing:
    for f in missing:
        print(f"ERROR: File not found — '{f}'")
    sys.exit(1)

# ── Read "AWB wise Details" sheet ────────────────────────────────────────────
try:
    raw = pd.read_excel(DISPATCHED_FILE, sheet_name="AWB wise Details", header=None)
except Exception as e:
    print(f"ERROR: Could not read '{DISPATCHED_FILE}' — {e}")
    sys.exit(1)

raw.columns = raw.iloc[0].tolist()
df = raw.iloc[1:].reset_index(drop=True)
df.columns = [str(c).strip() for c in df.columns]

needed = ["Order Dispatched Date", "AWB NO.", "Invoice Value", "Amount to Collect",
          "Customer Name", "Customer Number", "Customer Address", "Track Link", "Courier Partner"]
df = df[needed].copy()
df = df[df["AWB NO."].notna()].reset_index(drop=True)
df = df.drop_duplicates(subset=["AWB NO."]).reset_index(drop=True)

# ── Helpers ──────────────────────────────────────────────────────────────────
def to_num(val):
    """Convert to int if whole number, float otherwise; return original on failure."""
    try:
        f = float(val)
        return int(f) if f == int(f) else f
    except (ValueError, TypeError):
        return val

def last_two_tokens(address: str) -> str:
    if not isinstance(address, str):
        return str(address) if address else ""
    parts = [p.strip() for p in address.split(",")]
    parts = [p for p in parts if p]
    result = ", ".join(parts[-2:]) if len(parts) >= 2 else address.strip()
    return result[:-9].rstrip(", ") if len(result) > 9 else result

#def extract_cp_id(link: str) -> str:
#    if not isinstance(link, str):
#        return ""
#    m = re.search(r"cp_id=(\w+)", link)
#    return m.group(1) if m else ""

def fix_track_link(link: str) -> str:
    if not isinstance(link, str):
        return ""
    
    return link.replace("https://shopdeck.clickpost.ai/", "https://track.clickpost.in/")

def order_type(amount) -> str:
    try:
        return "prepaid" if float(amount) == 0 else "COD"
    except (ValueError, TypeError):
        return "prepaid"

# ── Build rows to append ─────────────────────────────────────────────────────
#today_str = str(date.today())
new_rows = []
for _, row in df.iterrows():
    track_raw   = str(row["Track Link"]).strip() if pd.notna(row["Track Link"]) else ""
    track_fixed = fix_track_link(track_raw)
    cp_id = row["Courier Partner"]
    #cp_id       = extract_cp_id(track_raw)
    date_object = datetime.strptime(row["Order Dispatched Date"], "%d/%m/%Y  %H:%M")
    formatted_date = date_object.strftime("%Y-%m-%d")

    new_rows.append([
        row["Customer Name"],           # col 1  - Customer Name
        to_num(row["AWB NO."]),         # col 2  - AWB (numeric)
        cp_id,                          # col 3  - CP_ID (text)
        "", "", "", "", "", "",         # col 4-9 - blank
        track_fixed,                    # col 10 - Track Link
        formatted_date,                 # col 11 - Dispatch Date
        last_two_tokens(row["Customer Address"]),  # col 12 - Address
        order_type(row["Amount to Collect"]),      # col 13 - Order Type
        to_num(row["Invoice Value"]),   # col 14 - Invoice Value (numeric)
        to_num(row["Customer Number"]), # col 15 - Customer Number (numeric)
    ])

# ── Open Google Sheet ─────────────────────────────────────────────────────────
try:
    gc = gspread.service_account(filename=SERVICE_ACCOUNT_FILE)
    ws = gc.open_by_key(SHEET_ID).worksheet(SHEET_TAB)
except Exception as e:
    print(f"ERROR: Could not open Google Sheet — {e}")
    sys.exit(1)

# ── Skip AWBs already in the sheet (col 2) ────────────────────────────────────
try:
    existing = {str(to_num(v)).strip() for v in
                ws.col_values(2, value_render_option="UNFORMATTED_VALUE")[1:] if v != ""}
except Exception as e:
    print(f"ERROR: Could not read existing AWBs from Google Sheet — {e}")
    sys.exit(1)

total = len(new_rows)
new_rows = [r for r in new_rows if str(r[1]).strip() not in existing]
print(f"Skipped {total - len(new_rows)} AWBs already in the sheet.")

# ── Append ────────────────────────────────────────────────────────────────────
# Numbers are sent as real numbers; USER_ENTERED lets Sheets parse the ISO date.
try:
    if new_rows:
        ws.append_rows(new_rows, value_input_option="USER_ENTERED", table_range="A1")
except Exception as e:
    print(f"ERROR: Could not append to Google Sheet — {e}")
    sys.exit(1)

print(f"Done. Appended {len(new_rows)} rows → Google Sheet '{SHEET_TAB}'")
for r in new_rows:
    print(r)
