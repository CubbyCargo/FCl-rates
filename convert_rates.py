import openpyxl
import json
from datetime import datetime, timezone

DEST_LABELS = {
    "TT": "Trinidad",
    "GUY": "Guyana",
    "SUR": "Suriname",
    "Trinidad Exports": "Trinidad Exports",
    "Print FE-TT": "Far East to Trinidad",
    "COL": "Colombia"
}

def clean(val):
    if val is None:
        return None
    if isinstance(val, float):
        return val if val == val else None
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

def extract_section(data, start_idx):
    results = []
    col_header_row = None
    headers = []
    for i in range(start_idx, min(start_idx + 5, len(data))):
        row = data[i]
        row_vals = [str(c).strip() if c else "" for c in row]
        if any(kw in " ".join(row_vals).upper() for kw in ["POL", "OF/BUNKER", "OCEAN FREIGHT"]):
            col_header_row = i
            headers = row_vals
            break
    if col_header_row is None:
        return results

    def col(keywords):
        for kw in keywords:
            for j, h in enumerate(headers):
                if kw.lower() in h.lower():
                    return j
        return None

    idx_pol = col(["POL"])
    idx_pod = col(["POD"])
    idx_of = col(["OF/Bunker", "Ocean Freight"])
    idx_total_no_ins = col(["Total w/out", "Total without"])
    idx_total_ins = col(["Total with"])
    idx_insurance = col(["Insurance"])
    idx_transit = col(["Transit"])
    idx_validity = col(["Validity"])
    idx_carrier = col(["Carrier"])
    idx_comment = col(["Comment"])

    current_pol = None
    current_pod = None

    for i in range(col_header_row + 1, len(data)):
        row = data[i]
        row_vals = [str(c).strip() if c else "" for c in row]
        joined = " ".join(row_vals)
        if any(kw in joined for kw in ["Rate are subject", "Notes"]):
            break
        if all(v == "" for v in row_vals):
            continue
        if "SHIPPING TARIFF" in joined.upper() and joined.upper() == joined:
            break
        if idx_pol is not None and row[idx_pol]:
            current_pol = clean(row[idx_pol])
        if idx_pod is not None and row[idx_pod]:
            current_pod = clean(row[idx_pod])
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
            "insurance": 200,
            "total_with_insurance": to_num(g(idx_total_ins)),
            "transit_time": clean(g(idx_transit)),
            "validity": parse_date(g(idx_validity)),
            "carrier": clean(g(idx_carrier)),
            "comment": clean(g(idx_comment)),
        }
        if entry["pol"] or entry["pod"]:
            results.append(entry)
    return results

# ── Parse Excel ──────────────────────────────────────────────────────────────
import glob, os

# Find whatever .xlsx file is in the repo root
xlsx_files = glob.glob("*.xlsx")
if not xlsx_files:
    raise FileNotFoundError("No .xlsx file found in repo root")
xlsx_file = xlsx_files[0]
print(f"Reading: {xlsx_file}")

wb = openpyxl.load_workbook(xlsx_file, read_only=True, data_only=True)

output = {
    "generated_at": datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ"),
    "source_file": xlsx_file,
    "destinations": {}
}

for sheet_name in wb.sheetnames:
    ws = wb[sheet_name]
    data = rows(ws)
    dest_key = sheet_name.strip()
    output["destinations"][dest_key] = {}
    i = 0
    while i < len(data):
        row = data[i]
        for cell in row:
            if cell:
                s = str(cell).strip().upper()
                if "SHIPPING TARIFF" in s:
                    label = str(cell).strip()
                    section_key = label.replace(" SHIPPING TARIFF", "").strip()
                    rates = extract_section(data, i)
                    if rates:
                        output["destinations"][dest_key][section_key] = rates
                    break
        i += 1

# ── Write rates.json ─────────────────────────────────────────────────────────
with open("rates.json", "w") as f:
    json.dump(output, f, indent=2)

# ── Generate index.html ──────────────────────────────────────────────────────
lines = []
lines.append(f"<p><strong>Last updated:</strong> {output['generated_at']} &nbsp;|&nbsp; <strong>Source:</strong> {xlsx_file}</p>")

for dest_key, sections in output["destinations"].items():
    dest_label = DEST_LABELS.get(dest_key, dest_key)
    lines.append(f"<h2>{dest_label}</h2>")
    for section, entries in sections.items():
        lines.append(f"<h3>{section}</h3>")
        lines.append("<table>")
        lines.append("<tr><th>POL</th><th>POD</th><th>Size</th><th>Carrier</th><th>Total (with ins.)</th><th>Transit</th><th>Validity</th><th>Comment</th></tr>")
        for e in entries:
            lines.append(
                f"<tr>"
                f"<td>{e.get('pol') or ''}</td>"
                f"<td>{e.get('pod') or ''}</td>"
                f"<td>{e.get('size') or ''}</td>"
                f"<td>{e.get('carrier') or ''}</td>"
                f"<td>USD {e.get('total_with_insurance') or ''}</td>"
                f"<td>{e.get('transit_time') or ''}</td>"
                f"<td>{e.get('validity') or ''}</td>"
                f"<td>{e.get('comment') or ''}</td>"
                f"</tr>"
            )
        lines.append("</table>")

body = "\n".join(lines)

html = f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<title>Ramps Logistics FCL Rates</title>
<style>
body {{ font-family: Arial, sans-serif; padding: 20px; max-width: 1200px; margin: 0 auto; }}
h1 {{ color: #1a1a1a; }}
h2 {{ color: #1a1a1a; border-bottom: 2px solid #ccc; padding-bottom: 4px; margin-top: 40px; }}
h3 {{ color: #333; margin-top: 20px; }}
table {{ border-collapse: collapse; width: 100%; margin-bottom: 20px; }}
th, td {{ border: 1px solid #ddd; padding: 6px 10px; text-align: left; font-size: 13px; }}
th {{ background: #f5f5f5; font-weight: bold; }}
tr:nth-child(even) {{ background: #fafafa; }}
</style>
</head>
<body>
<h1>Ramps Logistics FCL Shipping Rates</h1>
{body}
</body>
</html>"""

with open("index.html", "w") as f:
    f.write(html)

print(f"Done — {len([e for s in output['destinations'].values() for entries in s.values() for e in entries])} rate entries written to index.html and rates.json")
