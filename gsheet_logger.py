"""
Google Sheet Logger
Appends a new row to the RFQ Report Google Sheet after a successful RFQ submission.
Requires: gsheet_credentials.json (Service Account key) in the same folder.
"""

import os
import datetime

# Spreadsheet config
CONFIG_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "config.txt")

def get_config_val(key, default=""):
    if not os.path.exists(CONFIG_PATH): return default
    try:
        with open(CONFIG_PATH, "r", encoding="utf-8") as f:
            for line in f:
                if f"{key}=" in line:
                    return line.split("=", 1)[1].strip()
    except: pass
    return default

SPREADSHEET_ID = get_config_val("SPREADSHEET_ID", "1zVrlKMmbvrmdfnEzXfrkbCqB-ihvPla2YOswEqi_NTo")
SHEET_NAME     = "Sheet1"

CREDENTIALS_FILE = os.path.join(
    os.path.dirname(os.path.abspath(__file__)),
    "gsheet_credentials.json"
)


def _get_sheet():
    """Returns the gspread worksheet, or raises RuntimeError on failure."""
    try:
        import gspread
        from google.oauth2.service_account import Credentials
    except ImportError:
        raise RuntimeError("Missing packages. Run: pip install gspread google-auth")

    if not os.path.exists(CREDENTIALS_FILE):
        raise RuntimeError(
            f"gsheet_credentials.json not found at {CREDENTIALS_FILE}. "
            "Please download it from Google Cloud Console and place it there."
        )

    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    creds  = Credentials.from_service_account_file(CREDENTIALS_FILE, scopes=scopes)
    client = gspread.authorize(creds)
    return client.open_by_key(SPREADSHEET_ID).worksheet(SHEET_NAME)


def log_rfq_submission(order_row: dict, rfq_no: str) -> tuple[bool, str]:
    """
    Appends one row to the Google Sheet after a successful RFQ submission.

    Columns written (matching the RFQ Report layout):
      A: GX RFQ No.
      B: Customer - Item Description
      C: QT No.       (blank at submission time)
      D: Customer Email
      E: Markup %     (blank at submission time)
      F: Sales Price  (blank)
      G: Landed Cost  (blank)
      H: Profit       (blank)
      I: Ratio        (blank)
    """
    try:
        ws = _get_sheet()

        customer   = str(order_row.get("Customer Name", ""))
        item       = str(order_row.get("Item Name", ""))
        qty        = str(order_row.get("QTY", ""))
        email      = str(order_row.get("Customer Email", ""))
        date_str   = datetime.datetime.now().strftime("%Y-%m-%d")

        row = [
            rfq_no,                            # A: GX RFQ No.
            f"{customer} - {item} x{qty}",     # B: Customer - Item Desc
            "",                                 # C: QT No. (fill later)
            email,                             # D: Customer Email
            "20.00%",                          # E: Default markup (can edit on sheet)
            "",                                 # F: Sales Price
            "",                                 # G: Landed Cost
            "",                                 # H: Profit
            "",                                 # I: Ratio
        ]

        ws.append_row(row, value_input_option="USER_ENTERED")
        return True, f"✅ Logged to Google Sheet: {rfq_no}"

    except Exception as e:
        return False, f"Google Sheet log failed: {e}"

def update_rfq_quotation(rfq_no: str, cost: float, selling_price: float, profit: float, markup_pct: float, order_row: dict = None) -> tuple[bool, str]:
    """
    Finds the row by GX RFQ No. (Column A) and updates:
      E: Markup %
      F: Sales Price (Selling)
      G: Landed Cost (Cost)
      H: Profit
    If rfq_no is not found and order_row is provided, appends a new row.
    """
    try:
        import math
        ws = _get_sheet()
        
        # Find the row matching rfq_no
        col_a_values = ws.col_values(1)  # Getting column A
        row_idx = -1
        try:
            row_idx = col_a_values.index(rfq_no) + 1  # 1-based index
        except ValueError:
            # If not found, and we have order_row info, we append as a new row
            if order_row:
                # Helper to handle nan and variants from Pandas
                def clean_val(val):
                    if val is None or (isinstance(val, float) and math.isnan(val)) or str(val).lower() == "nan":
                        return ""
                    return str(val).strip()
                
                # Extremely robust extraction
                def get_val(d, *keys):
                    # 1. Exact match
                    for k in keys:
                        if k in d: return clean_val(d[k])
                    
                    # 2. Case-insensitive and stripped match
                    for k in keys:
                        search_k = str(k).lower().strip()
                        for dk in d.keys():
                            if str(dk).lower().strip() == search_k:
                                return clean_val(d[dk])
                                
                    # 3. Partial match (if key contains the search term)
                    for k in keys:
                        search_k = str(k).lower().strip()
                        for dk in d.keys():
                            if search_k in str(dk).lower():
                                return clean_val(d[dk])
                    return ""

                customer = get_val(order_row, "Customer Name", "Customer")
                item     = get_val(order_row, "Item Name", "Item", "Description")
                qty      = get_val(order_row, "QTY", "Qty", "Quantity")
                email    = get_val(order_row, "Customer Email", "Email", "Sender")
                
                # Format: Customer - Item xQty (Ensuring Name is in front)
                # If everything else fails, try to grab whatever we can
                if not customer and not item:
                    # Last resort: Try to see if there's ANY column with content
                    vals = [v for v in order_row.values() if clean_val(v)]
                    if vals: item = str(vals[0])

                desc_label = f"{customer} - {item}" if customer else item
                if qty and str(qty) != "1":
                    desc_label += f" x{qty}"
                
                if not desc_label: desc_label = "New Order (Details not found)"

                new_row = [
                    rfq_no,                            # A: GX RFQ No.
                    desc_label,                        # B: Customer - Item Desc
                    "",                                 # C: QT No.
                    email,                             # D: Customer Email
                    f"{markup_pct:.2f}%",              # E: Markup %
                    selling_price,                     # F: Sales Price
                    cost,                              # G: Landed Cost
                    profit,                            # H: Profit
                    ""                                 # I: Ratio
                ]
                ws.append_row(new_row, value_input_option="USER_ENTERED")
                return True, f"✅ RFQ {rfq_no} added as NEW row with info: {desc_label}"
            else:
                return False, f"RFQ {rfq_no} not found and no order info provided."
            
        # 4. Update existing row
        # We also check if Column B (Description) or Column D (Email) are missing and refill them if possible
        update_list = [
            {'range': f"E{row_idx}", 'values': [[f"{markup_pct:.2f}%"]]},
            {'range': f"F{row_idx}", 'values': [[selling_price]]}, # F: Sales Price
            {'range': f"G{row_idx}", 'values': [[cost]]},          # G: Landed Cost
            {'range': f"H{row_idx}", 'values': [[profit]]},
        ]
        
        if order_row:
            # Helper to handle nan and variants from pandas
            def get_clean_val(val):
                if val is None or (isinstance(val, float) and math.isnan(val)) or str(val).lower() == "nan":
                    return ""
                return str(val).strip()
            
            def find_val(d, *keys):
                for k in keys:
                    if k in d: return get_clean_val(d[k])
                search_keys = [str(k).lower().strip() for k in keys]
                for dk in d.keys():
                    if str(dk).lower().strip() in search_keys: return get_clean_val(d[dk])
                return ""

            # Check if Column B (Description) is currently empty
            try:
                curr_desc = str(ws.cell(row=row_idx, column=2).value or "").strip()
                if not curr_desc or curr_desc.lower() == "nan":
                    cust = find_val(order_row, "Customer Name", "Customer")
                    itm = find_val(order_row, "Item Name", "Item")
                    qty = find_val(order_row, "QTY", "Qty")
                    lbl = f"{cust} - {itm}" if cust else itm
                    if qty and str(qty) != "1": lbl += f" x{qty}"
                    if lbl:
                        update_list.append({'range': f"B{row_idx}", 'values': [[lbl]]})
                
                # Check Column D (Email)
                curr_email = str(ws.cell(row=row_idx, column=4).value or "").strip()
                if not curr_email or curr_email.lower() == "nan":
                    em = find_val(order_row, "Customer Email", "Email", "Sender")
                    if em:
                        update_list.append({'range': f"D{row_idx}", 'values': [[em]]})
            except:
                pass # Fallback: don't crash if individual cell fetch fails

        ws.batch_update(update_list, value_input_option="USER_ENTERED")
        return True, f"✅ Updated prices and missing info for {rfq_no} (Row {row_idx})"
        
    except Exception as e:
        return False, f"Failed to update Google Sheet: {e}"
