import json
import os
import glob
from datetime import datetime, timezone
import openpyxl

EXCEL_PATTERN = "Customer Rate Tariff Template_*.xlsx"
INSURANCE_AMOUNT = 200

COL_POL  = 2
COL_POD  = 3
COL_SIZE = 4
COL_OF_BUNKER      = 5
COL_THC            = 6
COL_LAC            = 7
COL_ISPS           = 8
COL_CONTAINER_INSP = 9
COL_GRI            = 10
COL_TOTAL          = 11
COL_INSURANCE      = 12
COL_TOTAL_WITH_INS = 13
COL_TRANSIT        = 14
COL_VALIDITY       = 15
COL_CARRIER        = 16
COL_COMMENT        = 17

DEST_MAP = {
    "TT": "TT",
    "GUY": "GUY",
    "GY": "GUY",
    "SUR": "SUR",
    "TRINIDAD EXPORTS": "Trinidad Exports",
    "PRINT FE-TT": "Print FE-TT",
    "COL": "COL",
}

def find_excel_file():
    matches = sorted(glob.glob(EXCEL_PATTERN))
    if not matches:
        raise FileNotFoundError(f"No Excel file matching '{EXCEL_PATTERN}' found in {os.getcwd()}")
    return matches[-1]

def safe_float(value):
    if value is None:
        return None
    if isinstance(value, str) and value.strip().upper() in ('N/A', '-', ''):
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

def build_cell_map(ws):
    merge_values = {}
    for merge_range in ws.merged_cells.ranges:
        top_left = ws.cell(merge_range.min_row, merge_range.min_col).value
        for row in range(merge_range.min_row, merge_range.max_row + 1):
            for col in range(merge_range.min_col, merge_range.max_col + 1):
                merge_values[(row, col)] = top_left
    cell_map = {}
    for row in ws.iter_rows():
        for cell in row:
            key = (cell.row, cell.column)
            cell_map[key] = merge_values.get(key, cell.value)
    return cell_map

def get(cell_map, r, c):
    return cell_map.get((r, c))

def is_section_title(cell_map, r):
    val = get(cell_map, r, COL_POL)
    if val and isinstance(val, str) and 'TARIFF' in val.upper():
        return val.strip()
    return None

def extract_lane(title):
    # "USA / TRINIDAD SHIPPING TARIFF" -> "USA / TRINIDAD"
    title = title.replace(' SHIPPING TARIFF', '').replace(' TARIFF', '').strip()
    return title

def parse_sheet(ws, dest_code):
    cell_map = build_cell_map(ws)
    max_row = ws.max_row
    sections = {}
    current_lane = None

    for r in range(1, max_row + 1):
        title = is_section_title(cell_map, r)
        if title:
            current_lane = extract_lane(title)
            if current_lane not in sections:
                sections[current_lane] = []
            continue

        if current_lane is None:
            continue

        size = safe_str(get(cell_map, r, COL_SIZE))
        if not size or size.upper() not in ('20FT', '40FT', '20GP', '40GP', '20HC', '40HC'):
            continue

        pol = safe_str(get(cell_map, r, COL_POL))
        pod = safe_str(get(cell_map, r, COL_POD))
        if not pol or not pod:
            continue

        validity_raw = get(cell_map, r, COL_VALIDITY)
        if hasattr(validity_raw, 'strftime'):
            validity = validity_raw.strftime('%d/%m/%Y')
        else:
            validity = safe_str(validity_raw)

        entry = {
            "pol":                   pol,
            "pod":                   pod,
            "size":                  size,
            "of_bunker":             safe_float(get(cell_map, r, COL_OF_BUNKER)),
            "thc":                   safe_float(get(cell_map, r, COL_THC)),
            "lac":                   safe_float(get(cell_map, r, COL_LAC)),
            "isps":                  safe_float(get(cell_map, r, COL_ISPS)),
            "container_inspection":  safe_float(get(cell_map, r, COL_CONTAINER_INSP)),
            "gri":                   safe_float(get(cell_map, r, COL_GRI)),
            "total":                 safe_float(get(cell_map, r, COL_TOTAL)),
            "total_with_insurance":  safe_float(get(cell_map, r, COL_TOTAL_WITH_INS)),
            "insurance":             INSURANCE_AMOUNT,
            "transit_time":          safe_str(get(cell_map, r, COL_TRANSIT)),
            "validity":              validity,
            "carrier":               safe_str(get(cell_map, r, COL_CARRIER)),
            "comment":               safe_str(get(cell_map, r, COL_COMMENT)),
        }
        sections[current_lane].append(entry)

    return sections

def convert(excel_path):
    print(f"Reading: {excel_path}")
    wb = openpyxl.load_workbook(excel_path, data_only=True)
    destinations = {}
    total_rows = 0

    print(f"\nSheets found: {wb.sheetnames}")

    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        dest_code = DEST_MAP.get(sheet_name.strip().upper(), sheet_name)
        sections = parse_sheet(ws, dest_code)

        row_count = sum(len(v) for v in sections.values())
        print(f"  Sheet '{sheet_name}' -> dest '{dest_code}' | {len(sections)} lanes | {row_count} rows")

        if not sections or row_count == 0:
            continue

        if dest_code not in destinations:
            destinations[dest_code] = {}
        destinations[dest_code].update(sections)
        total_rows += row_count

    output = {
        "generated_at": datetime.now(timezone.utc).isoformat().replace("+00:00", "Z"),
        "source_file": os.path.basename(excel_path),
        "destinations": destinations,
    }

    with open("rates.json", "w", encoding="utf-8") as f:
        json.dump(output, f, ensure_ascii=False, indent=2, default=str)

    size_kb = round(os.path.getsize("rates.json") / 1024, 1)
    print(f"\nDone - {total_rows} rows -> rates.json ({size_kb} KB)")

if __name__ == "__main__":
    excel_file = find_excel_file()
    convert(excel_file)
