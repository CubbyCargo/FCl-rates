import openpyxl
import json
import re
from datetime import datetime
from datetime import timezone

wb = openpyxl.load_workbook("Customer Rate Tariff Template_Week 21_2026.xlsx", read_only=True, data_only=True)

def clean(val):
    if val is None:
        return None
    if isinstance(val, float):
        return val if val == val else None  # filter NaN
    s = str(val).strip()
    return s if s and s not in ("-", "N/A", "") else None

def to_num(val):
    if val is None:
        return None
    try:
        s = str(val).replace(",", "").strip()
        if "Included" in s or "included" in s:
            return "Included"
        return float(s) if "." in s else int(s)
    except:
        return None

def parse_date(val):
    if val is None:
        return None
    if isinstance(val, datetime):
        return val.strftime("%d/%m/%Y")
    s = str(val).strip()
    # Excel serial dates sometimes slip through as integers
    if s.isdigit():
        try:
            from datetime import date, timedelta
            d = date(1899, 12, 30) + timedelta(days=int(s))
            return d.strftime("%d/%m/%Y")
        except:
            pass
    return s

def rows(ws):
    return list(ws.iter_rows(values_only=True))

def find_header_row(data, keyword):
    """Find the row index where a tariff section header appears."""
    for i, row in enumerate(data):
        for cell in row:
            if cell and keyword.lower() in str(cell).lower():
                return i
    return None

def extract_section(data, start_idx):
    """Extract rate rows from a section starting after the header row."""
    # Row after section title is usually the column header row
    # Then rate rows until blank or next section header
    results = []
    col_header_row = None
    headers = []
    
    # Find column header row (contains POL, POD, OF/Bunker etc.)
    for i in range(start_idx, min(start_idx + 5, len(data))):
        row = data[i]
        row_vals = [str(c).strip() if c else "" for c in row]
        if any(kw in " ".join(row_vals).upper() for kw in ["POL", "OF/BUNKER", "OCEAN FREIGHT"]):
            col_header_row = i
            headers = row_vals
            break
    
    if col_header_row is None:
        return results
    
    # Find key column indices
    def col(keywords):
        for kw in keywords:
            for j, h in enumerate(headers):
                if kw.lower() in h.lower():
                    return j
        return None

    idx_pol = col(["POL"])
    idx_pod = col(["POD"])
    idx_size = col(["20ft", "40ft", "container", "size"])  # sometimes blank col
    idx_of = col(["OF/Bunker", "Ocean Freight"])
    idx_total_no_ins = col(["Total w/out", "Total without"])
    idx_total_ins = col(["Total with"])
    idx_insurance = col(["Insurance"])
    idx_transit = col(["Transit"])
    idx_validity = col(["Validity"])
    idx_carrier = col(["Carrier"])
    idx_comment = col(["Comment"])
    idx_agent = col(["Agent", "Free Days"])

    current_pol = None
    current_pod = None

    for i in range(col_header_row + 1, len(data)):
        row = data[i]
        row_vals = [str(c).strip() if c else "" for c in row]
        joined = " ".join(row_vals)

        # Stop at Notes row or next major section
        if any(kw in joined for kw in ["Rate are subject", "Notes"]):
            break
        if all(v == "" for v in row_vals):
            continue
        # New section header (all caps tariff title)
        if any(kw in joined.upper() for kw in ["SHIPPING TARIFF", "TARIFF"]) and joined.upper() == joined:
            break

        # Carry forward POL/POD when cells are merged/blank
        if idx_pol is not None and row[idx_pol]:
            current_pol = clean(row[idx_pol])
        if idx_pod is not None and row[idx_pod]:
            current_pod = clean(row[idx_pod])

        # Detect container size — look for "20ft" / "40ft" pattern
        size = None
        for cell in row:
            if cell and str(cell).strip() in ("20ft", "40ft"):
                size = str(cell).strip()
                break

        if size is None:
            continue

        def g(idx):
            if idx is None or idx >= len(row):
                return None
            return row[idx]

        entry = {
            "pol": current_pol,
            "pod": current_pod,
            "size": size,
            "of_bunker": to_num(g(idx_of)),
            "total_without_insurance": to_num(g(idx_total_no_ins)),
            "insurance": to_num(g(idx_insurance)) if idx_insurance != idx_total_ins else 200,
            "total_with_insurance": to_num(g(idx_total_ins)),
            "transit_time": clean(g(idx_transit)),
            "validity": parse_date(g(idx_validity)),
            "carrier": clean(g(idx_carrier)),
            "comment": clean(g(idx_comment)),
        }

        if entry["pol"] or entry["pod"]:
            results.append(entry)

    return results

# ── Main parse ──────────────────────────────────────────────────────────────
output = {
    "generated_at": datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ"),
    "source_file": "Customer_Rate_Tariff_Template_Week_21_2026.xlsx",
    "destinations": {}
}

SECTION_KEYWORDS = [
    "USA", "BRAZIL", "CHINA", "KOREA", "TAIWAN", "JAPAN",
    "VIETNAM", "INDIA", "MALAYSIA", "THAILAND", "PANAMA",
    "TURKEY", "COLOMBIA", "TRINIDAD TO", "FAR EAST"
]

for sheet_name in wb.sheetnames:
    ws = wb[sheet_name]
    data = rows(ws)
    dest_key = sheet_name.strip()
    output["destinations"][dest_key] = {}

    # Find all tariff section headers in this sheet
    i = 0
    while i < len(data):
        row = data[i]
        for cell in row:
            if cell:
                s = str(cell).strip().upper()
                if "SHIPPING TARIFF" in s:
                    # e.g. "USA / TRINIDAD SHIPPING TARIFF"
                    label = str(cell).strip()
                    section_key = label.replace(" SHIPPING TARIFF", "").strip()
                    rates = extract_section(data, i)
                    if rates:
                        output["destinations"][dest_key][section_key] = rates
                    break
        i += 1

print(json.dumps(output, indent=2))
