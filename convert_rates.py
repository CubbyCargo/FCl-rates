import json
import os
import glob
import re
from datetime import datetime, timezone
import openpyxl

# ---------------------------------------------------------------------------
# Configuration - update these if filenames or sheet names change
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


def unmerge_and_fill(ws):
    """
    Build a dict of {(row, col): value} with merged cell values filled down/across.
    This resolves merged cells so every logical cell has the correct value.
    """
    # Collect the top-left value for every merged range
    merge_values = {}
    for merge_range in ws.merged_cells.ranges:
        top_left = ws.cell(merge_range.min_row, merge_range.min_col).value
        for row in range(merge_range.min_row, merge_range.max_row + 1):
            for col in range(merge_range.min_col, merge_range.max_col + 1):
                merge_values[(row, col)] = top_left

    # Build full cell map
    cell_map = {}
    for row in ws.iter_rows():
        for cell in row:
            if (cell.row, cell.column) in merge_values:
                cell_map[(cell.row, cell.column)] = merge_values[(cell.row, cell.column)]
            else:
                cell_map[(cell.row, cell.column)] = cell.value
    return cell_map


def get_cell(cell_map, row, col):
    """Get a value from the cell map by row and 1-indexed column."""
    return cell_map.get((row, col))


def parse_sheet(ws):
    """
    Parse a single worksheet, handling merged cells.
    Returns a list of dicts, one per data row.
    Skips rows where size is empty (merged POL/POD filled in, size is never merged).
    """
    cell_map = unmerge_and_fill(ws)
    rows = []

    max_row = ws.max_row
    for r in range(DATA_START_ROW, max_row + 1):
        # Use size as the anchor - it's never merged, so empty = not a data row
        size = safe_str(get_cell(cell_map, r, COL_SIZE))
        if not size:
            continue

        pol       = safe_str(get_cell(cell_map, r, COL_POL))
        pod       = safe_str(get_cell(cell_map, r, COL_POD))
        of_bunker = safe_float(get_cell(cell_map, r, COL_OF_BUNKER))
        thc       = safe_float(get_cell(cell_map, r, COL_THC))
        lac       = safe_float(get_cell(cell_map, r, COL_LAC))
        isps      = safe_float(get_cell(cell_map, r, COL_ISPS))
        cont_insp = safe_float(get_cell(cell_map, r, COL_CONTAINER_INSP))
        gri_raw   = get_cell(cell_map, r, COL_GRI)
        gri       = safe_float(gri_raw) if safe_str(gri_raw) not in ('N/A', 'n/a', None) else None
        total     = safe_float(get_cell(cell_map, r, COL_TOTAL))
        total_ins = safe_float(get_cell(cell_map, r, COL_TOTAL_WITH_INS))
        transit_raw = get_cell(cell_map, r, COL_TRANSIT)
        transit   = safe_str(transit_raw)
        validity_raw = get_cell(cell_map, r, COL_VALIDITY)
        # Normalise validity date
        if hasattr(validity_raw, 'strftime'):
            validity = validity_raw.strftime('%d/%m/%Y')
        else:
            validity = safe_str(validity_raw)
        carrier   = safe_str(get_cell(cell_map, r, COL_CARRIER))
        comment_raw = get_cell(cell_map, r, COL_COMMENT)
        comment   = safe_str(comment_raw)

        if not pol or not pod:
            continue

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
            "total_with_insurance":   total_ins,
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
      TT  - Trinidad
      GUY - Guyana
      SUR - Suriname
      COL - Colombia
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
    print(f"\nDone - {total_rows} rate rows → rates.json ({size_kb} KB)")


if __name__ == "__main__":
    excel_file = find_excel_file()
    convert(excel_file)
