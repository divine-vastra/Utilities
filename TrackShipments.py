import asyncio
import sys
import openpyxl
from datetime import datetime
from playwright.async_api import async_playwright
import time
import json
import re
import os
from pathlib import Path
from whatsapp_CloudAPISender import WhatsAppCloudAPI

def load_env_file(path: Path) -> None:
    """Minimal .env loader (no extra package needed). Lines: KEY=VALUE, '#' comments,
    optional quotes and 'export ' prefix. Real environment variables take priority."""
    if not path.is_file():
        print(f"[INFO] No .env file found at {path} — using system environment variables only")
        return
    for line in path.read_text(encoding="utf-8-sig").splitlines():
        line = line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        if line.lower().startswith("export "):
            line = line[7:].strip()
        key, value = line.split("=", 1)
        key, value = key.strip(), value.strip()
        if len(value) >= 2 and value[0] == value[-1] and value[0] in ("'", '"'):
            value = value[1:-1]
        os.environ.setdefault(key, value)

# .env lives in the same folder as this script (not the current working directory)
load_env_file(Path(__file__).resolve().parent / ".env")


# ─── CONFIG ───────────────────────────────────────────────────────────────────
EXCEL_FILE   = "shipments.xlsx"
POLL_SECONDS = 3600  # 60 minutes between full runs

# Excel column positions (1-based)
COL_CUSTOMER_NAME = 1   # A: Customer Name         ← NEW
COL_AWB           = 2   # B: AWB / waybill number  ← shifted
COL_CP_NAME       = 3   # C: Courier Partner       ← shifted
COL_STATUS        = 4   # D: Status description    ← shifted
COL_LOCATION      = 5   # E: Last scan location    ← shifted
COL_REMARK        = 6   # F: Last scan remark      ← shifted
COL_TIMESTAMP     = 7   # G: Time of last scan     ← shifted
COL_LAST_UPDATED  = 8   # H: When script last detected a change ← shifted
COL_EDD           = 9   # I: EDD
COL_TRACK_URL     = 10  # J: Tracking URL
COL_DEST          = 12  # L: Destination
COL_ORDER_TYPE    = 13  # M: Order Type - Prepaid or COD
COL_PHONE         = 15  # O: Customer phone number (10 digits, 91 gets prefixed)
COL_WA_DISPATCHED = 17  # Q: Timestamp when Order-Dispatched WhatsApp was sent
COL_WA_OFD        = 18  # R: Timestamp when Out-for-Delivery WhatsApp was sent
COL_WA_DELIVERED  = 19  # S: Timestamp when Order-Delivered WhatsApp was sent


# ─── WHATSAPP CONFIG ──────────────────────────────────────────────────────────
DRY_RUN           = False          # True = only log what would be sent; nothing is sent or marked
WA_TOKEN          = os.environ.get("WA_TOKEN", "")
WA_PHONE_NUMBER_ID = os.environ.get("WA_PHONE_NUMBER_ID", "")
WA_TEMPLATE_NAME  = "out_for_delivery"
WA_TEMPLATE_DISPATCHED = "order_dispatched"
WA_TEMPLATE_DELIVERED  = "order_delivered"
WA_LANG           = "en"
WA_FIXED_CP_ID    = "105"
OFD_STATUS        = "out for delivery"
DELIVERED_STATUS  = "delivered"
# ──────────────────────────────────────────────────────────────────────────────

# def build_url(awb: str, cp_id: str) -> str:
#     return f"https://track.clickpost.in/page?waybill={awb}&cp_id={cp_id}"

async def scrape_via_network_intercept(page, url: str) -> dict | None:
    captured = {}

    async def handle_response(response):
        try:
            resp_url = response.url
            if ("clickpost" in resp_url or "track" in resp_url) and response.status == 200:
                content_type = response.headers.get("content-type", "")
                if "json" in content_type:
                    body = await response.json()
                    if any(k in str(body) for k in ["status", "waybill", "scan", "shipment"]):
                        captured["data"] = body
        except Exception:
            pass

    page.on("response", handle_response)
    await page.goto(url, wait_until="networkidle", timeout=30000)
    await page.wait_for_timeout(3000)

    if not captured.get("data"):
        return None

    return parse_intercepted_json(captured["data"])

def parse_intercepted_json(data: dict) -> dict | None:
    try:
        result = data.get("result", {})
        if result:
            shipment = next(iter(result.values()), {})
            if shipment.get("valid"):
                latest = shipment.get("latest_status", {})
                additional = shipment.get("additional", {})
                return {
                    "status":    latest.get("clickpost_status_description")
                                 or latest.get("status", ""),
                    "location":  latest.get("location", ""),
                    "remark":    latest.get("remark", ""),
                    "timestamp": latest.get("timestamp", ""),
                    "edd":       additional.get("courier_partner_edd", ""),
                }

        status = (data.get("status") or data.get("current_status")
                  or data.get("shipment_status") or "")
        if status:
            return {
                "status":    str(status),
                "location":  str(data.get("location") or data.get("city") or ""),
                "remark":    str(data.get("remark") or data.get("message") or ""),
                "timestamp": str(data.get("timestamp") or data.get("updated_at") or ""),
                "edd":       str(data.get("edd") or data.get("expected_delivery") or ""),
            }
    except Exception as e:
        print(f"    [WARN] JSON parse error: {e}")

    return None

async def scrape_via_html(page) -> dict | None:
    try:
        await page.wait_for_selector(
            "[class*='status'], [class*='clickpost_status_bucket_description'], [class*='track'], [class*='Track']",
            timeout=15000
        )
    except Exception:
        print("    [WARN] Timed out waiting for status element on page")

    html = await page.content()

    selectors_to_try = [
        #("[class*='current-status']",  "status"),
        ("[id*='currentStatus']",   "status"),
        #("[class*='status-Text']",     "status"),
        #("[class*='statusText']",      "status"),
        #("[class*='shipment-status']", "status"),
        #("[class*='tracking-status']", "status"),
        ("[class*='location']",        "location"),
        #("[class*='city']",            "location"),
        ("[class*='remark']",          "remark"),
        #("[class*='message']",         "remark"),
        #("[class*='description']",     "remark"),
        ("[id*='currentTimeStamp']",     "timestamp"),
        ("[class*='expected-date pull-left']",     "edd"),
        
        
    ]

    result = {"status": "", "location": "", "remark": "", "timestamp": "", "edd": ""}

    for selector, field in selectors_to_try:
        if result[field]:
            continue
        try:
            el = await page.query_selector(selector)
            if el:
                text = (await el.inner_text()).strip()
                if text and len(text) < 200:
                    result[field] = text
        except Exception:
            pass

    if not result["status"]:
        patterns = [
            r'"current_status"\s*:\s*"([^"]+)"',
            r'"status"\s*:\s*"([^"]+)"',
            r'"clickpost_status_description"\s*:\s*"([^"]+)"',
            r'status["\s:]+([A-Za-z ]{4,30})',
        ]
        for pattern in patterns:
            match = re.search(pattern, html, re.IGNORECASE)
            if match:
                result["status"] = match.group(1).strip()
                break

    return result if result["status"] else None

async def scrape_shipment(browser, awb: str, track_url: str) -> dict | None:
    #url = build_url(awb, cp_id)
    context = await browser.new_context(
        user_agent=(
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) "
            "Chrome/120.0.0.0 Safari/537.36"
        )
    )
    page = await context.new_page()

    try:
        print(f"    → Fetching: {track_url}")

        data = await scrape_via_network_intercept(page, track_url)

        if data:
            print(f"    ✓ Got data via network intercept")
        else:
            print(f"    → Network intercept empty, falling back to HTML scrape...")
            data = await scrape_via_html(page)
            if data:
                print(f"    ✓ Got data via HTML scrape")
            else:
                print(f"    ✗ Could not extract status — check selectors for this page")

        return data

    except Exception as e:
        print(f"    [ERROR] {e}")
        return None
    finally:
        await context.close()

def normalize_phone(raw) -> str | None:
    """Return 91XXXXXXXXXX or None if the number is not usable."""
    if raw is None:
        return None
    if isinstance(raw, float):
        raw = int(raw)
    digits = re.sub(r"\D", "", str(raw))
    if len(digits) == 12 and digits.startswith("91"):
        return digits
    digits = digits.lstrip("0")
    if len(digits) == 10:
        return "91" + digits
    return None

def is_ofd(status: str) -> bool:
    return (status or "").strip().lower() == OFD_STATUS

def is_delivered(status: str) -> bool:
    return (status or "").strip().lower() == DELIVERED_STATUS

def send_whatsapp(client, template: str, phone: str, awb: str) -> bool:
    """Send a template (tracking-link button, except order_delivered which has none). Returns True on success."""
    try:
        if template == WA_TEMPLATE_DELIVERED:
            components = []
        else:
            components = [client.build_url_button(index=0, value=f"{awb}&cp_id={WA_FIXED_CP_ID}")]
        client.send_template(phone, template, components, WA_LANG)
        return True
    except Exception as e:
        print(f"    ⚠ WhatsApp failed for {phone} (AWB {awb}): {e}")
        return False

async def run():
    try:
        wb = openpyxl.load_workbook(EXCEL_FILE)
    except Exception as e:
        print(f"ERROR: Could not read '{EXCEL_FILE}' — {e}")
        sys.exit(1)

    
    ws = wb['Sheet1']
    #ws = wb.active
    updated_count = 0
    total_rows = 0
    wa_sent_count = 0

    for col, header in ((COL_WA_OFD, "WA OFD Sent"),
                        (COL_WA_DISPATCHED, "WA Dispatched Sent"),
                        (COL_WA_DELIVERED, "WA Delivered Sent")):
        if not ws.cell(row=1, column=col).value:
            ws.cell(row=1, column=col).value = header

    wa_client = None
    if not DRY_RUN:
        if not WA_TOKEN or not WA_PHONE_NUMBER_ID:
            print("ERROR: set WA_TOKEN and WA_PHONE_NUMBER_ID environment variables (or set DRY_RUN = True).")
            sys.exit(1)
        wa_client = WhatsAppCloudAPI(token=WA_TOKEN, phone_number_id=WA_PHONE_NUMBER_ID)
    else:
        print("  *** DRY_RUN is ON — no WhatsApp messages will be sent ***")

    print(f"\n{'='*60}")
    print(f"  ClickPost Page Scraper — {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"{'='*60}")

    async with async_playwright() as pw:
        browser = await pw.chromium.launch(headless=True)

        for row in ws.iter_rows(min_row=2):  # Row 1 = headers
            customer_name = str(row[COL_CUSTOMER_NAME - 1].value).strip() if row[COL_CUSTOMER_NAME - 1].value else ""
            awb           = str(row[COL_AWB   - 1].value).strip() if row[COL_AWB   - 1].value else None
            cp_name         = str(row[COL_CP_NAME - 1].value).strip() if row[COL_CP_NAME - 1].value else None
            dest          = str(row[COL_DEST - 1].value).strip() if row[COL_DEST - 1].value else None
            order_type    = str(row[COL_ORDER_TYPE - 1].value).strip() if row[COL_ORDER_TYPE - 1].value else None
            track_url     = str(row[COL_TRACK_URL - 1].value).strip() if row[COL_TRACK_URL - 1].value else None

            if not awb or not cp_name:
                continue

            total_rows += 1
            print(f"\n{total_rows:02}. → Customer: {customer_name or 'N/A'} | Dest: {dest} | TYPE: {order_type} | CP_NAME: {cp_name}")

            new = await scrape_shipment(browser, awb, track_url)

            if not new:
                print(f"  Skipping — no data returned")
                continue

            old_status = row[COL_STATUS - 1].value or ""

            if new["status"] and new["status"] != old_status:
                print(f"    ✅ STATUS CHANGED: [{old_status}] → [{new['status']}], EDD: {new['edd']}, {new['location'][:30]} ")
                updated_count += 1
            else:
                print(f"    — No change: {new['status']},  EDD: {new['edd']}, {new['location'][:30]}")

            row[COL_STATUS       - 1].value = new["status"]
            row[COL_LOCATION     - 1].value = new["location"]
            row[COL_REMARK       - 1].value = new["remark"]
            row[COL_TIMESTAMP    - 1].value = new["timestamp"]
            row[COL_LAST_UPDATED - 1].value = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            if new["edd"]:
                row[COL_EDD - 1].value = new["edd"]

            # ── WhatsApp events (each sent once per shipment, tracked by its own column) ──
            #   blank -> non-blank  : order_dispatched
            #   any   -> Out for Delivery : out_for_delivery
            #   any   -> Delivered  : order_delivered
            r = row[0].row
            events = []
            if not old_status.strip() and new["status"].strip():
                events.append(("Dispatched", WA_TEMPLATE_DISPATCHED, COL_WA_DISPATCHED))
            if is_ofd(new["status"]) and not is_ofd(old_status):
                events.append(("OFD", WA_TEMPLATE_NAME, COL_WA_OFD))
            if is_delivered(new["status"]) and not is_delivered(old_status):
                events.append(("Delivered", WA_TEMPLATE_DELIVERED, COL_WA_DELIVERED))

            for label, template, sent_col in events:
                sent_cell = ws.cell(row=r, column=sent_col)
                if sent_cell.value:
                    continue
                phone = normalize_phone(ws.cell(row=r, column=COL_PHONE).value)
                if not phone:
                    print(f"    ⚠ {label} but phone invalid/missing in column O — will retry next poll")
                elif DRY_RUN:
                    button = "no components" if template == WA_TEMPLATE_DELIVERED else f"button value '{awb}&cp_id={WA_FIXED_CP_ID}'"
                    print(f"    [DRY_RUN] Would send '{template}' to {phone} (AWB {awb}, {button})")
                elif send_whatsapp(wa_client, template, phone, awb):
                    sent_cell.value = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    wa_sent_count += 1
                    print(f"    📲 WhatsApp '{template}' sent to {phone} (AWB {awb})")
                    try:
                        wb.save(EXCEL_FILE)  # persist immediately to avoid duplicates after a crash
                    except Exception as e:
                        print(f"    [WARN] Could not save after send (is the Excel file open?): {e}")

            await asyncio.sleep(2)

        await browser.close()

    wb.save(EXCEL_FILE)
    print(f"\n{'='*60}")
    print(f"  Done. {updated_count}/{total_rows} shipments updated. WhatsApp sent: {wa_sent_count}")
    print(f"{'='*60}\n")

if __name__ == "__main__":
    while True:
        asyncio.run(run())
        print(f"Sleeping {POLL_SECONDS // 60} minutes before next poll...\n")
        time.sleep(POLL_SECONDS)