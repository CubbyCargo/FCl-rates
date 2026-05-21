"""
convert_rates.py
Reads the Cubby Cargo FCL rate tariff Excel file and outputs rates.json
for GitHub Pages / Respond.io knowledge source.

Excel column mapping (per tariff template):
  A  - POL
  B  - POD
  C  - Container Size
  D  - OF/Bunker
  E  - THC"""
convert_rates.py
Reads the Cubby Cargo FCL rate tariff Excel file and outputs rates.json
for GitHub Pages / Respond.io knowledge source.

Excel column mapping (per tariff template):
  A  - POL
  B  - POD
  C  - Container Size
  D  - OF/Bunker
  E  - THC
  F  - LAC
  G  - Dredging
  H  - Terminal Lease Surcharge
  I  - GRI
  J  - Total (without insurance)
  K  - Transit Time
  L  - Validity
  M  - Carrier
  N  - Comment (optional)
"""

import json
import os
import glob
import re
from datetime import datetime, timezone
import openpyxl

# ---------------------------------------------------------------------------
# Configuration — update these if filenames or sheet names change
# ---------------------------------------------------------------------------
EXCEL_PATTERN = "Customer Rate Tariff Template_*.xlsx"   # wildcards OK
DATA_START_ROW = 10      # rows 1-9 are title/header blocks; row 10 = first data row
INSURANCE_AMOUNT = 200   # fixed USD amount added for total_with_insurance

# Column indexes (1-indexed, matching openpyxl default)
COL_POL            = 1   # A
COL_POD            = 2   # B
COL_SIZE           = 3   # C
COL_OF_BUNKER      = 4   # D
COL_THC            = 5   # E
COL_LAC            = 6   # F
COL_ISPS           = 7   # G
COL_CONTAINER_INSP = 8   # H  Container Inspection
COL_GRI            = 9   # I
COL_TOTAL          = 10  # J  Total without Insurance
COL_INSURANCE      = 11  # K
COL_TOTAL_WITH_INS = 12  # L  Total with Insurance
COL_TRANSIT        = 13  # M
COL_VALIDITY       = 14  # N
COL_CARRIER        = 15  # O
COL_COMMENT        = 16  # P  (optional)
# ---------------------------------------------------------------------------


def find_excel_file():
    matches = sorted(glob.glob(EXCEL_PATTERN))
    if not matches:
        raise FileNotFoundError(
            f"No Excel file matching '{EXCEL_PATTERN}' found in {os.getcwd()}"
        )
    return matches[-1]   # use the most recent if multiple exist


def safe_float(value):
    """Return float or None - handles empty cells, strings, and numbers."""
    if value is None:
        return None
    try:
        return float(value)
    except (ValueError, TypeError):
        return None


def safe_str(value):
    if value is None:
        return None
    v = str(value).strip()
    return v if v else None


def parse_sheet(ws):
    """
    Parse a single worksheet.
    Returns a list of dicts, one per data row.
    Skips rows where POL or POD is empty.
    Also returns the section heading (sheet name or first non-empty header row).
    """
    rows = []
    for row in ws.iter_rows(min_row=DATA_START_ROW, values_only=True):
        pol       = safe_str(row[COL_POL - 1])
        pod       = safe_str(row[COL_POD - 1])
        if not pol or not pod:
            continue  # skip blank / subtotal rows

        size      = safe_str(row[COL_SIZE - 1])
        of_bunker = safe_float(row[COL_OF_BUNKER - 1])
        thc       = safe_float(row[COL_THC - 1])
        lac       = safe_float(row[COL_LAC - 1])
        isps      = safe_float(row[COL_ISPS - 1])
        cont_insp = safe_float(row[COL_CONTAINER_INSP - 1])
        gri       = safe_float(row[COL_GRI - 1])
        total     = safe_float(row[COL_TOTAL - 1])
        transit   = safe_str(row[COL_TRANSIT - 1])
        validity  = safe_str(row[COL_VALIDITY - 1])
        carrier   = safe_str(row[COL_CARRIER - 1])
        comment   = safe_str(row[COL_COMMENT - 1]) if len(row) >= COL_COMMENT else None

        # Normalise validity date to dd/mm/yyyy string
        if hasattr(total, 'strftime'):
            total = None
        if validity and hasattr(row[COL_VALIDITY - 1], 'strftime'):
            validity = row[COL_VALIDITY - 1].strftime('%d/%m/%Y')

        entry = {
            "pol":                    pol,
            "pod":                    pod,
            "size":                   size,
            "of_bunker":              of_bunker,
            "thc":                    thc,
            "lac":                    lac,
            "isps":                   isps,
            "container_inspection":   cont_insp,
            "gri":                    gri,
            "total":                  total,
            "total_with_insurance":   safe_float(row[COL_TOTAL_WITH_INS - 1]),
            "insurance":              INSURANCE_AMOUNT,
            "transit_time":           transit,
            "validity":               validity,
            "carrier":                carrier,
            "comment":                comment,
        }
        rows.append(entry)
    return rows


def group_by_destination_and_lane(all_rows, sheet_name):
    """
    Derive destination and lane from sheet name.
    Expected sheet naming convention: 'ORIGIN / DESTINATION' or 'DEST - LANE'
    Falls back to sheet_name as the lane key.

    Destination codes:
      TT  — Trinidad
      GUY — Guyana
      SUR — Suriname
      COL — Colombia
      (anything else stored under its own key)
    """
    dest_keywords = {
        "TRINIDAD": "TT",
        "T&T": "TT",
        "TT": "TT",
        "GUYANA": "GUY",
        "GUY": "GUY",
        "GY": "GUY",
        "SURINAME": "SUR",
        "SUR": "SUR",
        "COLOMBIA": "COL",
        "COL": "COL",
        "TRINIDAD EXPORTS": "Trinidad Exports",
    }
    name_upper = sheet_name.strip().upper()
    dest_code = dest_keywords.get(name_upper)
    if not dest_code:
        dest_code = sheet_name   # fallback: use raw sheet name

    return dest_code, sheet_name


def convert(excel_path):
    print(f"Reading: {excel_path}")
    wb = openpyxl.load_workbook(excel_path, data_only=True)

    destinations = {}
    total_rows = 0

    print(f"\nDEBUG sheet names found in workbook:")
    for s in wb.sheetnames:
        print(f"  repr: {repr(s)}")

    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        rows = parse_sheet(ws)
        print(f"  Sheet '{sheet_name}' → {len(rows)} rows parsed")
        if not rows:
            print(f"  Skipped (no data): {sheet_name}")
            continue

        dest_code, lane = group_by_destination_and_lane(rows, sheet_name)

        if dest_code not in destinations:
            destinations[dest_code] = {}
        destinations[dest_code][lane] = rows
        total_rows += len(rows)
        print(f"  Parsed sheet '{sheet_name}' → dest '{dest_code}' | {len(rows)} rows")

    output = {
        "generated_at": datetime.now(timezone.utc).isoformat().replace("+00:00", "Z"),
        "source_file": os.path.basename(excel_path),
        "destinations": destinations,
    }

    with open("rates.json", "w", encoding="utf-8") as f:
        json.dump(output, f, ensure_ascii=False, indent=2, default=str)

    size_kb = round(os.path.getsize("rates.json") / 1024, 1)
    print(f"\nDone — {total_rows} rate rows → rates.json ({size_kb} KB)")


if __name__ == "__main__":
    excel_file = find_excel_file()
    convert(excel_file)
  F  - LAC
  G  - Dredging
  H  - Terminal Lease Surcharge
  I  - GRI
  J  - Total (without insurance)
  K  - Transit Time
  L  - Validity
  M  - Carrier
  N  - Comment (optional)
"""

import json
import os
import glob
import re
from datetime import datetime, timezone
import openpyxl

# ---------------------------------------------------------------------------
# Configuration — update these if filenames or sheet names change
# ---------------------------------------------------------------------------
EXCEL_PATTERN = "Customer Rate Tariff Template_*.xlsx"   # wildcards OK
DATA_START_ROW = 2       # first row of data (1-indexed); row 1 = headers
INSURANCE_AMOUNT = 200   # fixed USD amount added for total_with_insurance

# Column indexes (1-indexed, matching openpyxl default)
COL_POL            = 1   # A
COL_POD            = 2   # B
COL_SIZE           = 3   # C
COL_OF_BUNKER      = 4   # D
COL_THC            = 5   # E
COL_LAC            = 6   # F
COL_DREDGING       = 7   # G
COL_TLS            = 8   # H  Terminal Lease Surcharge
COL_GRI            = 9   # I
COL_TOTAL          = 10  # J  Total (excl. insurance)
COL_TRANSIT        = 11  # K
COL_VALIDITY       = 12  # L
COL_CARRIER        = 13  # M
COL_COMMENT        = 14  # N  (optional)
# ---------------------------------------------------------------------------


def find_excel_file():
    matches = sorted(glob.glob(EXCEL_PATTERN))
    if not matches:
        raise FileNotFoundError(
            f"No Excel file matching '{EXCEL_PATTERN}' found in {os.getcwd()}"
        )
    return matches[-1]   # use the most recent if multiple exist


def safe_float(value):
    """Return float or None — handles empty cells, strings, and numbers."""
    if value is None:
        return None
    try:
        return float(value)
    except (ValueError, TypeError):
        return None


def safe_str(value):
    if value is None:
        return None
    v = str(value).strip()
    return v if v else None


def parse_sheet(ws):
    """
    Parse a single worksheet.
    Returns a list of dicts, one per data row.
    Skips rows where POL or POD is empty.
    Also returns the section heading (sheet name or first non-empty header row).
    """
    rows = []
    for row in ws.iter_rows(min_row=DATA_START_ROW, values_only=True):
        pol       = safe_str(row[COL_POL - 1])
        pod       = safe_str(row[COL_POD - 1])
        if not pol or not pod:
            continue  # skip blank / subtotal rows

        size      = safe_str(row[COL_SIZE - 1])
        of_bunker = safe_float(row[COL_OF_BUNKER - 1])
        thc       = safe_float(row[COL_THC - 1])
        lac       = safe_float(row[COL_LAC - 1])
        dredging  = safe_float(row[COL_DREDGING - 1])
        tls       = safe_float(row[COL_TLS - 1])
        gri       = safe_float(row[COL_GRI - 1])
        total     = safe_float(row[COL_TOTAL - 1])
        transit   = safe_str(row[COL_TRANSIT - 1])
        validity  = safe_str(row[COL_VALIDITY - 1])
        carrier   = safe_str(row[COL_CARRIER - 1])
        comment   = safe_str(row[COL_COMMENT - 1]) if len(row) >= COL_COMMENT else None

        # Normalise validity date to dd/mm/yyyy string
        if hasattr(total, 'strftime'):
            total = None   # guard: if Excel puts a date in wrong column
        if validity and hasattr(row[COL_VALIDITY - 1], 'strftime'):
            validity = row[COL_VALIDITY - 1].strftime('%d/%m/%Y')

        entry = {
            "pol":                    pol,
            "pod":                    pod,
            "size":                   size,
            "of_bunker":              of_bunker,
            "thc":                    thc,
            "lac":                    lac,
            "dredging":               dredging,
            "terminal_lease_surcharge": tls,
            "gri":                    gri,
            "total":                  total,
            "total_with_insurance":   round(total + INSURANCE_AMOUNT, 2) if total is not None else None,
            "insurance":              INSURANCE_AMOUNT,
            "transit_time":           transit,
            "validity":               validity,
            "carrier":                carrier,
            "comment":                comment,
        }
        rows.append(entry)
    return rows


def group_by_destination_and_lane(all_rows, sheet_name):
    """
    Derive destination and lane from sheet name.
    Expected sheet naming convention: 'ORIGIN / DESTINATION' or 'DEST - LANE'
    Falls back to sheet_name as the lane key.

    Destination codes:
      TT  — Trinidad
      GUY — Guyana
      SUR — Suriname
      COL — Colombia
      (anything else stored under its own key)
    """
    dest_keywords = {
        "TRINIDAD": "TT",
        "T&T": "TT",
        "TT": "TT",
        "GUYANA": "GUY",
        "GUY": "GUY",
        "GY": "GUY",
        "SURINAME": "SUR",
        "SUR": "SUR",
        "COLOMBIA": "COL",
        "COL": "COL",
        "TRINIDAD EXPORTS": "Trinidad Exports",
    }
    name_upper = sheet_name.strip().upper()
    dest_code = dest_keywords.get(name_upper)
    if not dest_code:
        dest_code = sheet_name   # fallback: use raw sheet name

    return dest_code, sheet_name


def convert(excel_path):
    print(f"Reading: {excel_path}")
    wb = openpyxl.load_workbook(excel_path, data_only=True)

    destinations = {}
    total_rows = 0

    print(f"\nDEBUG sheet names found in workbook:")
    for s in wb.sheetnames:
        print(f"  repr: {repr(s)}")

    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        rows = parse_sheet(ws)
        print(f"  Sheet '{sheet_name}' → {len(rows)} rows parsed")
        if not rows:
            print(f"  Skipped (no data): {sheet_name}")
            continue

        dest_code, lane = group_by_destination_and_lane(rows, sheet_name)

        if dest_code not in destinations:
            destinations[dest_code] = {}
        destinations[dest_code][lane] = rows
        total_rows += len(rows)
        print(f"  Parsed sheet '{sheet_name}' → dest '{dest_code}' | {len(rows)} rows")

    output = {
        "generated_at": datetime.now(timezone.utc).isoformat().replace("+00:00", "Z"),
        "source_file": os.path.basename(excel_path),
        "destinations": destinations,
    }

    with open("rates.json", "w", encoding="utf-8") as f:
        json.dump(output, f, ensure_ascii=False, indent=2, default=str)

    size_kb = round(os.path.getsize("rates.json") / 1024, 1)
    print(f"\nDone — {total_rows} rate rows → rates.json ({size_kb} KB)")


if __name__ == "__main__":
    excel_file = find_excel_file()
    convert(excel_file)
