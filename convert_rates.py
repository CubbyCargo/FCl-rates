"""
build_html.py  —  FCl-rates repo (old Excel format)
Reads the old-style Customer Rate Tariff Template Excel file and generates:
  docs/index.html   — full tariff page (new card/table style)
  docs/rates.json   — data feed consumed by quote.html
  docs/quote.html   — quote generator (new style)

Old Excel structure (per sheet, one section per trade lane):
  - Sections separated by a bold "ORIGIN / DEST SHIPPING TARIFF" header row
  - Next row: column headers  (POL, POD, OF/Bunker, THC, LAC/Local Charges, ISPS, ...)
  - Data rows: col 1=POL (may be blank = carry-forward), col 2=POD (may be blank),
               col 3=container size, cols 4..N=surcharge values,
               penultimate cols = Total without Insurance, Insurance, Total with Insurance,
               then Transit Time, Validity, Carrier, Agent
  - Sections end at a "Notes" row
  - Sheets: TT, GUY, SUR, COL  (Trinidad Exports and Print FE-TT are skipped)

Column indices (0-based) are FIXED per section header row — we detect them dynamically.
"""

import pandas as pd
import json
import sys
from pathlib import Path
from datetime import datetime, timezone
from collections import defaultdict

# ── Paths ──────────────────────────────────────────────────────────────────
# Match both "Customer Rate Tariff Template_Week 23_2026.xlsx" (spaces)
# and "Customer_Rate_Tariff_Template_Week_23_2026.xlsx" (underscores)
_candidates = list(Path(".").glob("Customer*Tariff*Template*Week*.xlsx")) + \
              list(Path(".").glob("Customer Rate Tariff Template_Week *.xlsx"))
if not _candidates:
    print("ERROR: No Excel file found. Expected 'Customer Rate Tariff Template_Week *.xlsx'", file=sys.stderr)
    sys.exit(1)
EXCEL_PATH = _candidates[0]
OUTPUT_PATH       = Path("docs/index.html")
OUTPUT_JSON_PATH  = Path("docs/rates.json")
OUTPUT_QUOTE_PATH = Path("docs/quote.html")

# Sheets to process (skip Trinidad Exports, Print FE-TT)
SHEETS_TO_PROCESS = ["TT", "GUY", "SUR", "COL"]

# Map sheet name → destination base label (used if POD is ambiguous)
SHEET_DEST_HINT = {
    "TT":  "Trinidad & Tobago",
    "GUY": "Guyana",
    "SUR": "Suriname",
    "COL": "Colombia",
}

# ── Helpers ────────────────────────────────────────────────────────────────
def clean_val(v):
    """Return float if numeric and > 0, else None."""
    try:
        f = float(v)
        return f if f > 0 else None
    except (TypeError, ValueError):
        return None

def clean_str(v):
    s = str(v).strip()
    return s if s not in ("nan", "None", "") else ""

def fmt_usd(v):
    if v is None:
        return "-"
    s = f"${v:,.2f}"
    # Strip trailing zeros after decimal
    if "." in s:
        s = s.rstrip("0").rstrip(".")
    return s

def get_destination_label(pod):
    pod = clean_str(pod)
    tt = {"Port of Spain", "Point Lisas", "Port-of-Spain", "Port of Spain/Point Lisas"}
    gy = {"Georgetown"}
    sr = {"Paramaribo"}
    bw = {"Bridgetown"}
    kn = {"Kingston"}
    ca = {"Caucedo"}
    co = {"Buenaventura", "Buenaventua", "Barranquilla", "Cartagena", "Buenaventura/Barranquilla"}
    if pod in tt or "Port of Spain" in pod or "Point Lisas" in pod:
        return "Trinidad & Tobago"
    if pod in gy:
        return "Guyana"
    if pod in sr:
        return "Suriname"
    if pod in bw:
        return "Bridgetown"
    if pod in kn:
        return "Kingston"
    if pod in ca:
        return "Caucedo"
    if any(p in pod for p in co):
        return "Colombia"
    return pod

# ── Excel reader ───────────────────────────────────────────────────────────
def parse_sheet(sheet_name: str) -> list[dict]:
    """
    Parse one sheet and return a list of rate-row dicts.
    Each dict has keys: origin, pol, pod, container, surcharges,
                        total_no_ins, total_with_ins, transit, validity, carrier, agent
    """
    df = pd.read_excel(EXCEL_PATH, sheet_name=sheet_name, header=None)

    rates = []
    section_origin = ""       # e.g. "USA"
    col_map = {}              # column name → col index (rebuilt per section header)
    current_pol = ""
    current_pod = ""

    # Column name aliases — maps various header spellings → canonical key
    SURCHARGE_ALIASES = {
        "of/bunker":                    "Ocean Freight",
        "ocean freight":                "Ocean Freight",
        "thc":                          "THC",
        "thc ":                         "THC",
        "lac":                          "LAC",
        "lac/local charges":            "LAC",
        "local charges":                "LAC",
        "local charges ":               "LAC",
        "isps":                         "ISPS",
        "container inspection":         "Container Inspection",
        "gri":                          "GRI",
        "other port charges":           "Other Port Charges",
        "local handling":               "Local Handling",
        "admin":                        "Admin",
        "dthc":                         "THC",
        "doc fee":                      "DOC Fee",
        "eir (vat incl)":               "EIR",
        "dredging ":                    "Dredging Fee",
        "dredging fee":                 "Dredging Fee",
        "terminal lease surchage":      "Terminal Lease",
        "terminal lease surcharge":     "Terminal Lease",
        "port charges":                 "Port Charges",
        "handling fee":                 "Handling Fee",
    }

    def parse_header_row(row) -> dict:
        """Given a raw pandas row, build col_map."""
        cm = {}
        for j, v in enumerate(row):
            if pd.isna(v):
                continue
            key = str(v).strip().lower()
            if key in ("pol",):
                cm["POL"] = j
            elif key in ("pod",):
                cm["POD"] = j
            elif key in ("total without insurance", "total w/out insurance", "total w/out ins"):
                cm["total_no_ins"] = j
            elif key in ("insurance ", "insurance"):
                cm["insurance"] = j
            elif key in ("total with insurance", "total w/ insurance"):
                cm["total_with_ins"] = j
            elif key in ("transit time ", "transit time"):
                cm["transit"] = j
            elif key in ("validity ", "validity"):
                cm["validity"] = j
            elif key in ("carrier",):
                cm["carrier"] = j
            elif key in ("agent",):
                cm["agent"] = j
            elif key in SURCHARGE_ALIASES:
                label = SURCHARGE_ALIASES[key]
                if "surcharges" not in cm:
                    cm["surcharges"] = {}
                cm["surcharges"][j] = label
        return cm

    def is_section_header(row) -> str:
        """If this row is a 'ORIGIN / DEST SHIPPING TARIFF' line, return origin country."""
        non_null = [clean_str(v) for v in row if pd.notna(v) and clean_str(v)]
        if len(non_null) == 1:
            val = non_null[0].upper()
            if "SHIPPING TARIFF" in val or "TARIFF" in val:
                # Extract origin country (everything before the first '/')
                origin = val.split("/")[0].strip().title()
                # Clean up "Usa" → "USA"
                origin = origin.replace("Usa", "USA").replace("Uk", "UK")
                return origin
        return ""

    def is_column_header(row) -> bool:
        """Check if this row contains 'POL' in col 1."""
        for j, v in enumerate(row):
            if pd.notna(v) and clean_str(v).upper() == "POL":
                return True
        return False

    def is_notes_row(row) -> bool:
        non_null = [clean_str(v) for v in row if pd.notna(v) and clean_str(v)]
        if non_null and non_null[0].startswith("Notes"):
            return True
        return False

    for i, row in df.iterrows():
        row_vals = list(row)

        # Detect section header
        origin = is_section_header(row_vals)
        if origin:
            section_origin = origin
            current_pol = ""
            current_pod = ""
            continue

        # Detect column header row
        if is_column_header(row_vals):
            col_map = parse_header_row(row_vals)
            current_pol = ""
            current_pod = ""
            continue

        # Skip notes
        if is_notes_row(row_vals):
            continue

        # Skip if we don't have a col_map yet
        if not col_map:
            continue

        # Try to parse a data row
        # POL — col 1 (index 1), carry forward if blank
        pol_col = col_map.get("POL", 1)
        pod_col = col_map.get("POD", 2)
        size_col = 3   # always col index 3

        pol_val = clean_str(row_vals[pol_col]) if pol_col < len(row_vals) else ""
        pod_val = clean_str(row_vals[pod_col]) if pod_col < len(row_vals) else ""
        size_val = clean_str(row_vals[size_col]) if size_col < len(row_vals) else ""

        if pol_val:
            current_pol = pol_val
        if pod_val:
            current_pod = pod_val

        # Must have a container size to be a data row
        if size_val not in ("20ft", "40ft", "20 ft", "40 ft"):
            continue
        # Normalise
        size_val = "20ft" if "20" in size_val else "40ft"

        if not current_pol or not current_pod:
            continue

        # Surcharges
        surcharges = {}
        for col_idx, label in col_map.get("surcharges", {}).items():
            v = clean_val(row_vals[col_idx]) if col_idx < len(row_vals) else None
            if v:
                surcharges[label] = v

        total_no_ins  = clean_val(row_vals[col_map["total_no_ins"]])  if "total_no_ins"  in col_map and col_map["total_no_ins"]  < len(row_vals) else None
        total_with_ins = clean_val(row_vals[col_map["total_with_ins"]]) if "total_with_ins" in col_map and col_map["total_with_ins"] < len(row_vals) else None

        # Derive missing totals
        if total_no_ins is None and surcharges:
            total_no_ins = round(sum(surcharges.values()), 2)
        if total_with_ins is None and total_no_ins is not None:
            total_with_ins = round(total_no_ins + 200, 2)

        transit  = clean_str(row_vals[col_map["transit"]])  if "transit"  in col_map and col_map["transit"]  < len(row_vals) else ""
        validity = clean_str(row_vals[col_map["validity"]]) if "validity" in col_map and col_map["validity"] < len(row_vals) else ""
        carrier  = clean_str(row_vals[col_map["carrier"]])  if "carrier"  in col_map and col_map["carrier"]  < len(row_vals) else ""
        agent    = clean_str(row_vals[col_map["agent"]])    if "agent"    in col_map and col_map["agent"]    < len(row_vals) else ""

        # Format validity date nicely if it's a date object
        if hasattr(validity, "strftime"):
            validity = validity.strftime("%d/%m/%Y")

        dest_label = get_destination_label(current_pod)

        rates.append({
            "sheet":         sheet_name,
            "origin":        section_origin,
            "pol":           current_pol,
            "pod":           current_pod,
            "dest_label":    dest_label,
            "container":     size_val,
            "commodity":     "",   # old format has no commodity column
            "carrier":       carrier,
            "agent":         agent,
            "transit":       transit,
            "validity":      validity,
            "surcharges":    surcharges,
            "total_no_ins":  total_no_ins,
            "total_with_ins": total_with_ins,
        })

    return rates


# ── Group into cards ────────────────────────────────────────────────────────
def build_cards(all_rates: list[dict]) -> list[dict]:
    lane_map = defaultdict(list)
    for r in all_rates:
        lane_key = f"{r['origin']} → {r['dest_label']}"
        lane_map[lane_key].append(r)

    def lane_sort_key(lane):
        dest_order = ["Trinidad & Tobago", "Guyana", "Suriname", "Colombia", "Bridgetown", "Kingston", "Caucedo"]
        dest = lane.split(" → ", 1)[1] if " → " in lane else lane
        try:
            return (dest_order.index(dest), lane)
        except ValueError:
            return (99, lane)

    cards = []
    for lane in sorted(lane_map.keys(), key=lane_sort_key):
        rows = lane_map[lane]
        origin, dest = lane.split(" → ", 1)
        carriers  = sorted(set(r["carrier"] for r in rows if r["carrier"]))
        validities = sorted(set(r["validity"] for r in rows if r["validity"]))
        transits   = sorted(set(r["transit"]  for r in rows if r["transit"]))
        cards.append({
            "lane":      lane,
            "origin":    origin,
            "dest":      dest,
            "carriers":  carriers,
            "validities": validities,
            "transits":  transits,
            "rate_rows": rows,
        })
    return cards


# ── HTML renderer ────────────────────────────────────────────────────────────
def render_html(cards: list[dict]) -> str:
    generated = datetime.now(timezone.utc).strftime("%d %b %Y %H:%M UTC")
    dests = sorted(set(c["dest"] for c in cards))
    dest_options = "".join(f'<button class="filter-btn" data-dest="{d}">{d}</button>' for d in dests)

    cards_html = ""
    for c in cards:
        dest_slug = c["dest"].lower().replace(" ", "-").replace("&", "and").replace("(", "").replace(")", "")

        table_rows = ""
        for r in c["rate_rows"]:
            surcharge_breakdown = "".join(
                f'<div class="surcharge-item"><span class="surcharge-label">{k}</span>'
                f'<span class="surcharge-val">{fmt_usd(v)}</span></div>'
                for k, v in r["surcharges"].items()
            )
            agent_note = f'<div class="agent-note">{r["agent"]}</div>' if r["agent"] else ""
            commodity_cell = r["commodity"] if r["commodity"] else "FAK"
            table_rows += f"""
            <tr>
              <td><span class="tag">{r["container"]}</span></td>
              <td class="pol-cell">{r["pol"]}<div class="pod-sub">→ {r["pod"]}</div></td>
              <td class="commodity-cell">{commodity_cell}</td>
              <td class="carrier-cell">{r["carrier"]}{agent_note}</td>
              <td class="transit-cell">{r["transit"]}</td>
              <td class="validity-cell">{r["validity"]}</td>
              <td class="surcharge-cell"><div class="surcharge-grid">{surcharge_breakdown}</div></td>
              <td class="total-cell">
                <div class="rate-block">
                  <div class="rate-no-ins">{fmt_usd(r["total_no_ins"])}<span class="rate-label">excl. insurance</span></div>
                  <div class="rate-with-ins">{fmt_usd(r["total_with_ins"])}<span class="rate-label ins-label">🛡 insured</span></div>
                </div>
              </td>
            </tr>"""

        carriers_badges = " ".join(f'<span class="badge">{cr}</span>' for cr in c["carriers"])
        transit_range   = " / ".join(sorted(set(c["transits"])))  or "—"
        validity_range  = " / ".join(sorted(set(c["validities"]))) or "—"

        cards_html += f"""
    <div class="card" data-dest="{c['dest']}" data-dest-slug="{dest_slug}">
      <div class="card-header">
        <div class="lane-title">
          <span class="origin-label">{c['origin']}</span>
          <span class="arrow">→</span>
          <span class="dest-label">{c['dest']}</span>
        </div>
        <div class="card-meta">
          <span class="meta-item">🚢 {carriers_badges}</span>
          <span class="meta-item">⏱ {transit_range}</span>
          <span class="meta-item">📅 {validity_range}</span>
        </div>
      </div>
      <div class="card-body">
        <div class="table-wrapper">
          <table>
            <thead>
              <tr>
                <th>Container</th>
                <th>POL → POD</th>
                <th>Commodity</th>
                <th>Carrier</th>
                <th>Transit</th>
                <th>Validity</th>
                <th>Surcharge Breakdown</th>
                <th>Rate</th>
              </tr>
            </thead>
            <tbody>{table_rows}</tbody>
          </table>
        </div>
      </div>
    </div>"""

    return f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8"/>
<meta name="viewport" content="width=device-width, initial-scale=1.0"/>
<title>Cubby Cargo – FCL Rate Tariff</title>
<style>
  :root {{
    --purple: #6b22d3; --purple-dark: #4e12a8; --purple-light: #f3eeff;
    --green: #8bea98; --green-dark: #2db84b; --green-bg: #f0fdf3;
    --white: #ffffff; --bg: #f8f7fc; --border: #e4ddf5;
    --text: #1a1030; --muted: #7a6e8a;
    --card-shadow: 0 2px 16px rgba(107,34,211,0.08);
  }}
  * {{ box-sizing: border-box; margin: 0; padding: 0; }}
  body {{ font-family: 'Segoe UI', Arial, sans-serif; background: var(--bg); color: var(--text); font-size: 13px; }}
  header {{ background: var(--purple); color: white; padding: 18px 32px; display: flex; align-items: center; justify-content: space-between; }}
  header h1 {{ font-size: 1.25rem; font-weight: 800; }}
  header .generated {{ font-size: 11px; opacity: 0.65; }}
  .subtitle {{ background: linear-gradient(90deg, var(--purple-dark), var(--purple)); color: var(--green); padding: 5px 32px; font-size: 11px; font-weight: 700; letter-spacing: 1px; text-transform: uppercase; }}
  .controls {{ padding: 14px 32px; background: var(--white); border-bottom: 1px solid var(--border); display: flex; gap: 8px; flex-wrap: wrap; align-items: center; }}
  .controls label {{ font-size: 10px; font-weight: 700; color: var(--muted); text-transform: uppercase; letter-spacing: 0.6px; margin-right: 4px; }}
  .filter-btn {{ background: var(--white); border: 1.5px solid var(--border); border-radius: 20px; padding: 5px 14px; font-size: 12px; cursor: pointer; color: var(--muted); transition: all 0.15s; font-weight: 500; }}
  .filter-btn:hover {{ border-color: var(--purple); color: var(--purple); }}
  .filter-btn.active {{ background: var(--purple); color: white; border-color: var(--purple); font-weight: 600; }}
  .filter-btn.all-btn.active {{ background: var(--purple-dark); }}
  .main {{ padding: 24px 32px; display: flex; flex-direction: column; gap: 18px; }}
  .card {{ background: var(--white); border-radius: 12px; box-shadow: var(--card-shadow); overflow: hidden; border: 1.5px solid var(--border); }}
  .card-header {{ background: linear-gradient(135deg, var(--purple) 0%, var(--purple-dark) 100%); color: white; padding: 14px 20px; display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap; gap: 8px; }}
  .lane-title {{ display: flex; align-items: center; gap: 10px; font-size: 1rem; font-weight: 800; }}
  .dest-label {{ color: var(--green); }}
  .arrow {{ color: rgba(255,255,255,0.5); }}
  .card-meta {{ display: flex; gap: 16px; font-size: 11px; opacity: 0.9; flex-wrap: wrap; align-items: center; }}
  .badge {{ background: rgba(255,255,255,0.15); border-radius: 4px; padding: 2px 8px; font-size: 10px; font-weight: 700; }}
  .table-wrapper {{ overflow-x: auto; }}
  table {{ width: 100%; border-collapse: collapse; font-size: 12px; }}
  thead tr {{ background: var(--purple-light); }}
  th {{ padding: 9px 12px; text-align: left; font-size: 10px; font-weight: 700; text-transform: uppercase; letter-spacing: 0.5px; color: var(--purple); border-bottom: 2px solid var(--border); white-space: nowrap; }}
  td {{ padding: 10px 12px; border-bottom: 1px solid #f0ecfa; vertical-align: top; }}
  tr:last-child td {{ border-bottom: none; }}
  tr:hover td {{ background: #faf8ff; }}
  .tag {{ background: var(--purple); color: white; border-radius: 5px; padding: 3px 9px; font-size: 10px; font-weight: 700; white-space: nowrap; }}
  .pol-cell {{ font-weight: 600; }}
  .pod-sub {{ font-size: 10px; color: var(--muted); font-weight: 400; margin-top: 2px; }}
  .commodity-cell {{ color: var(--muted); font-size: 11px; }}
  .carrier-cell {{ font-weight: 700; color: var(--purple); }}
  .agent-note {{ font-size: 10px; color: var(--muted); font-weight: 400; margin-top: 2px; }}
  .transit-cell, .validity-cell {{ white-space: nowrap; color: var(--muted); font-size: 11px; }}
  .surcharge-cell {{ min-width: 200px; max-width: 300px; }}
  .surcharge-grid {{ display: flex; flex-direction: column; gap: 3px; }}
  .surcharge-item {{ display: flex; justify-content: space-between; background: var(--purple-light); border-radius: 4px; padding: 3px 8px; font-size: 11px; }}
  .surcharge-label {{ color: var(--muted); font-weight: 500; }}
  .surcharge-val {{ color: var(--purple); font-weight: 700; margin-left: 8px; }}
  .rate-block {{ display: flex; flex-direction: column; gap: 6px; min-width: 120px; }}
  .rate-no-ins {{ display: flex; flex-direction: column; font-size: 14px; font-weight: 700; color: var(--muted); background: #f5f3fa; border-radius: 6px; padding: 6px 10px; }}
  .rate-with-ins {{ display: flex; flex-direction: column; font-size: 16px; font-weight: 800; color: var(--green-dark); background: var(--green-bg); border: 1.5px solid var(--green); border-radius: 6px; padding: 6px 10px; }}
  .rate-label {{ font-size: 9px; font-weight: 600; text-transform: uppercase; letter-spacing: 0.5px; color: var(--muted); margin-top: 2px; }}
  .ins-label {{ color: var(--green-dark); }}
  .notes {{ background: var(--purple-light); border: 1.5px solid var(--border); border-radius: 10px; padding: 16px 22px; margin: 0 32px 28px; font-size: 11px; color: var(--purple-dark); line-height: 1.7; }}
  .notes strong {{ display: block; margin-bottom: 6px; font-size: 12px; color: var(--purple); }}
  .hidden {{ display: none !important; }}
  @media (max-width: 700px) {{ header {{ padding: 14px 16px; }} .main {{ padding: 14px; }} .controls {{ padding: 12px 16px; }} .notes {{ margin: 0 14px 20px; }} }}
</style>
</head>
<body>
<header>
  <h1>🚢 Cubby Cargo — FCL Rate Tariff</h1>
  <span class="generated">Generated: {generated}</span>
</header>
<div class="subtitle">All-in rates (USD) · FCL only · Subject to space &amp; equipment availability</div>
<div class="controls">
  <label>Filter by Destination:</label>
  <button class="filter-btn all-btn active" data-dest="all">All Destinations</button>
  {dest_options}
</div>
<div class="main" id="cards-container">
{cards_html}
</div>
<div class="notes">
  <strong>📋 Important Notes</strong>
  Rates are subject to space and equipment validity. Cargo must be ingated on or before the specified validity date; updated rates may apply thereafter.<br/>
  Marine Insurance covers a C&amp;F value of USD $30,000 at an additional <strong>$200</strong>. Values greater than USD $30,000 and restricted commodities must be quoted on a case-by-case basis.<br/>
  Should client decline Marine Insurance provided by Ramps Logistics, Ramps Logistics shall not be held liable for any claims, loss or damages arising from the execution of services.
</div>
<script>
  const btns = document.querySelectorAll('.filter-btn');
  const cards = document.querySelectorAll('.card');
  btns.forEach(btn => {{
    btn.addEventListener('click', () => {{
      btns.forEach(b => b.classList.remove('active'));
      btn.classList.add('active');
      const dest = btn.dataset.dest;
      cards.forEach(card => {{
        card.classList.toggle('hidden', dest !== 'all' && card.dataset.dest !== dest);
      }});
    }});
  }});
</script>
</body>
</html>"""


# ── JSON builder ─────────────────────────────────────────────────────────────
def build_json(cards: list[dict]):
    generated = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
    dest_key_map = {
        "Trinidad & Tobago": "TT",
        "Guyana":            "GUY",
        "Suriname":          "SUR",
        "Colombia":          "COL",
        "Bridgetown":        "BDG",
        "Kingston":          "KIN",
        "Caucedo":           "CAU",
    }

    destinations = {}
    for c in cards:
        dest_key = dest_key_map.get(c["dest"], c["dest"])
        lane_key = c["lane"].replace(" → ", " / ").upper()
        if dest_key not in destinations:
            destinations[dest_key] = {}

        rate_list = []
        for r in c["rate_rows"]:
            rate_list.append({
                "pol":                    r["pol"],
                "pod":                    r["pod"],
                "size":                   r["container"],
                "commodity":              r["commodity"] or "FAK",
                "carrier":                r["carrier"],
                "agent":                  r["agent"],
                "transit_time":           r["transit"],
                "validity":               r["validity"],
                "surcharges":             r["surcharges"],
                "total_without_insurance": r["total_no_ins"],
                "insurance":              200,
                "total_with_insurance":   r["total_with_ins"],
            })

        if lane_key not in destinations[dest_key]:
            destinations[dest_key][lane_key] = []
        destinations[dest_key][lane_key].extend(rate_list)

    payload = {
        "generated_at": generated,
        "source_file":  str(EXCEL_PATH.name),
        "destinations": destinations,
    }
    OUTPUT_JSON_PATH.parent.mkdir(parents=True, exist_ok=True)
    OUTPUT_JSON_PATH.write_text(json.dumps(payload, indent=2), encoding="utf-8")
    total = sum(len(v) for d in destinations.values() for v in d.values())
    print(f"✅ Written to {OUTPUT_JSON_PATH} ({len(destinations)} destinations, {total} rate rows)")


# ── Quote HTML builder ────────────────────────────────────────────────────────
def build_quote(cards: list[dict]):
    lanes_js = json.dumps([
        {
            "lane":   c["lane"],
            "origin": c["origin"],
            "dest":   c["dest"],
            "rates": [{
                "pol":            r["pol"],
                "pod":            r["pod"],
                "container":      r["container"],
                "commodity":      r["commodity"] or "FAK",
                "carrier":        r["carrier"],
                "agent":          r["agent"],
                "transit":        r["transit"],
                "validity":       r["validity"],
                "surcharges":     r["surcharges"],
                "total_no_ins":   r["total_no_ins"],
                "total_with_ins": r["total_with_ins"],
            } for r in c["rate_rows"]],
        }
        for c in cards
    ], indent=2)

    dests = sorted(set(c["dest"] for c in cards))

    html = f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8"/>
<meta name="viewport" content="width=device-width, initial-scale=1.0"/>
<title>Cubby Cargo — FCL Quote Generator</title>
<style>
  :root {{
    --purple: #6b22d3; --purple-dark: #4e12a8; --purple-light: #f3eeff;
    --green: #8bea98; --green-dark: #2db84b; --green-bg: #f0fdf3;
    --white: #ffffff; --bg: #f8f7fc; --border: #e4ddf5;
    --text: #1a1030; --muted: #7a6e8a; --shadow: 0 2px 16px rgba(107,34,211,0.10);
  }}
  * {{ box-sizing: border-box; margin: 0; padding: 0; }}
  body {{ font-family: 'Segoe UI', Arial, sans-serif; background: var(--bg); color: var(--text); font-size: 13px; }}

  @media screen {{
    header {{ background: var(--purple); color: white; padding: 16px 32px; display: flex; align-items: center; justify-content: space-between; }}
    header h1 {{ font-size: 1.2rem; font-weight: 800; }}
    header a {{ color: var(--green); font-size: 12px; text-decoration: none; }}
    .subtitle {{ background: linear-gradient(90deg, var(--purple-dark), var(--purple)); color: var(--green); padding: 5px 32px; font-size: 11px; font-weight: 700; letter-spacing: 1px; text-transform: uppercase; }}
    .main {{ max-width: 900px; margin: 28px auto; padding: 0 20px; display: flex; flex-direction: column; gap: 20px; }}
    .form-card {{ background: var(--white); border-radius: 12px; border: 1.5px solid var(--border); box-shadow: var(--shadow); padding: 24px; }}
    .form-card h2 {{ font-size: 14px; font-weight: 700; color: var(--purple); margin-bottom: 18px; }}
    .form-grid {{ display: grid; grid-template-columns: 1fr 1fr; gap: 14px; }}
    .field {{ display: flex; flex-direction: column; gap: 5px; }}
    .field label {{ font-size: 10px; font-weight: 700; text-transform: uppercase; letter-spacing: 0.5px; color: var(--muted); }}
    .field input, .field select {{ border: 1.5px solid var(--border); border-radius: 7px; padding: 8px 11px; font-size: 13px; color: var(--text); background: var(--white); outline: none; transition: border-color 0.15s; }}
    .field input:focus, .field select:focus {{ border-color: var(--purple); }}
    .field select:disabled {{ background: #f5f3fa; color: var(--muted); }}
    .btn-search {{ margin-top: 6px; background: var(--purple); color: white; border: none; border-radius: 8px; padding: 10px 28px; font-size: 13px; font-weight: 700; cursor: pointer; }}
    .btn-search:hover {{ background: var(--purple-dark); }}
    .results-card {{ background: var(--white); border-radius: 12px; border: 1.5px solid var(--border); box-shadow: var(--shadow); padding: 24px; display: none; }}
    .results-card h2 {{ font-size: 14px; font-weight: 700; color: var(--purple); margin-bottom: 16px; }}
    .rate-row {{ border: 1.5px solid var(--border); border-radius: 10px; padding: 14px 16px; margin-bottom: 10px; display: grid; grid-template-columns: 1fr 1fr 1.5fr auto; gap: 12px; align-items: start; cursor: pointer; }}
    .rate-row:hover {{ border-color: var(--purple); background: var(--purple-light); }}
    .rate-row.selected {{ border-color: var(--green-dark); background: var(--green-bg); }}
    .btn-pdf {{ background: var(--purple); color: white; border: none; border-radius: 8px; padding: 10px 24px; font-size: 13px; font-weight: 700; cursor: pointer; }}
    .btn-pdf:hover {{ background: var(--purple-dark); }}
    .quote-card {{ background: var(--white); border-radius: 12px; border: 1.5px solid var(--border); box-shadow: var(--shadow); padding: 28px; display: none; }}
  }}

  @media print {{
    @page {{ size: A4; margin: 18mm 18mm 14mm 18mm; }}
    header, .subtitle, .form-card, .results-card, .btn-pdf {{ display: none !important; }}
    body {{ background: white; font-size: 12px; }}
    .main {{ margin: 0; padding: 0; max-width: 100%; gap: 0; }}
    .quote-card {{ display: block !important; border: none; box-shadow: none; padding: 0; border-radius: 0; }}
    .ins-push, .quote-notes, .quote-lane-banner, .total-box, .detail-block {{ -webkit-print-color-adjust: exact; print-color-adjust: exact; }}
    .print-logo {{ display: block !important; }}
  }}

  .print-logo {{ display: none; margin-bottom: 16px; }}
  .print-logo h2 {{ font-size: 1.1rem; font-weight: 800; color: var(--purple); }}
  .print-logo p {{ font-size: 10px; color: var(--muted); margin-top: 2px; }}
  .rate-meta {{ font-size: 11px; color: var(--muted); margin-top: 3px; }}
  .rate-carrier {{ font-weight: 700; color: var(--purple); }}
  .surcharge-list {{ font-size: 11px; }}
  .surcharge-line {{ display: flex; justify-content: space-between; padding: 2px 0; border-bottom: 1px dashed #ede8f8; }}
  .surcharge-line:last-child {{ border-bottom: none; }}
  .surcharge-line span:last-child {{ font-weight: 600; color: var(--purple); }}
  .totals {{ text-align: right; }}
  .total-no {{ font-size: 13px; color: var(--muted); margin-bottom: 6px; }}
  .total-with {{ font-size: 18px; font-weight: 800; color: var(--green-dark); background: var(--green-bg); border: 1.5px solid var(--green); border-radius: 7px; padding: 6px 12px; display: inline-block; }}
  .ins-note {{ font-size: 9px; color: var(--green-dark); display: block; margin-top: 2px; }}
  .tag {{ background: var(--purple); color: white; border-radius: 4px; padding: 2px 8px; font-size: 10px; font-weight: 700; }}
  .no-results {{ color: var(--muted); font-size: 13px; padding: 12px 0; }}
  .quote-header {{ display: flex; justify-content: space-between; align-items: flex-start; margin-bottom: 16px; flex-wrap: wrap; gap: 10px; }}
  .quote-title {{ font-size: 1.1rem; font-weight: 800; color: var(--purple); }}
  .quote-meta {{ font-size: 11px; color: var(--muted); margin-top: 4px; }}
  .quote-ref {{ font-size: 11px; color: var(--muted); text-align: right; }}
  .quote-lane-banner {{ background: linear-gradient(135deg, var(--purple), var(--purple-dark)); color: white; border-radius: 8px; padding: 11px 16px; display: flex; align-items: center; gap: 10px; font-size: 1rem; font-weight: 800; margin-bottom: 16px; }}
  .quote-lane-banner .dest {{ color: var(--green); }}
  .quote-details {{ display: grid; grid-template-columns: 1fr 1fr; gap: 14px; margin-bottom: 16px; }}
  .detail-block {{ background: var(--purple-light); border-radius: 8px; padding: 11px 13px; }}
  .detail-block h4 {{ font-size: 10px; font-weight: 700; text-transform: uppercase; letter-spacing: 0.5px; color: var(--purple); margin-bottom: 7px; }}
  .detail-line {{ display: flex; justify-content: space-between; font-size: 11px; padding: 3px 0; border-bottom: 1px solid #e4ddf5; }}
  .detail-line:last-child {{ border-bottom: none; }}
  .detail-line .val {{ font-weight: 600; }}
  .quote-totals {{ display: flex; gap: 12px; margin-bottom: 14px; }}
  .total-box {{ flex: 1; border-radius: 9px; padding: 12px 16px; text-align: center; }}
  .total-box.no-ins {{ background: #f5f3fa; border: 1.5px solid var(--border); }}
  .total-box.with-ins {{ background: var(--green-bg); border: 2px solid var(--green); }}
  .total-box .label {{ font-size: 10px; font-weight: 700; text-transform: uppercase; letter-spacing: 0.5px; color: var(--muted); margin-bottom: 5px; }}
  .total-box.with-ins .label {{ color: var(--green-dark); }}
  .total-box .amount {{ font-size: 20px; font-weight: 800; }}
  .total-box.no-ins .amount {{ color: var(--muted); }}
  .total-box.with-ins .amount {{ color: var(--green-dark); }}
  .ins-push {{ background: linear-gradient(90deg, var(--green-bg), #e8fded); border: 1.5px solid var(--green); border-radius: 8px; padding: 10px 14px; font-size: 11px; color: var(--green-dark); margin-bottom: 14px; }}
  .ins-push strong {{ font-size: 12px; }}
  .quote-notes {{ font-size: 10px; color: var(--muted); line-height: 1.6; background: var(--purple-light); border-radius: 8px; padding: 10px 14px; }}
  @media (max-width: 600px) {{ .form-grid {{ grid-template-columns: 1fr; }} .rate-row {{ grid-template-columns: 1fr; }} .quote-details {{ grid-template-columns: 1fr; }} .quote-totals {{ flex-direction: column; }} }}
</style>
</head>
<body>
<header>
  <h1>🚢 Cubby Cargo — FCL Quote Generator</h1>
  <a href="index.html">← View Full Tariff</a>
</header>
<div class="subtitle">FCL All-in Rates (USD) · Subject to space &amp; equipment availability</div>

<div class="main">
  <div class="form-card">
    <h2>📋 Shipment Details</h2>
    <div class="form-grid">
      <div class="field">
        <label>Prepared for</label>
        <input type="text" id="customer" placeholder="Customer name"/>
      </div>
      <div class="field">
        <label>Quote Reference</label>
        <input type="text" id="quoteref" placeholder="Auto-generated"/>
      </div>
      <div class="field">
        <label>Destination</label>
        <select id="sel-dest">
          <option value="">Select destination</option>
          {"".join(f'<option value="{d}">{d}</option>' for d in dests)}
        </select>
      </div>
      <div class="field">
        <label>Trade Lane</label>
        <select id="sel-lane" disabled><option value="">Select lane</option></select>
      </div>
      <div class="field">
        <label>Container Size</label>
        <select id="sel-size" disabled><option value="">Select size</option></select>
      </div>
    </div>
    <button class="btn-search" onclick="doSearch()">Search Rates</button>
  </div>

  <div class="results-card" id="results-card">
    <h2>Available Rates <span id="results-count" style="font-weight:400;color:var(--muted);font-size:12px;"></span></h2>
    <div id="results-list"></div>
  </div>

  <div class="quote-card" id="quote-card">
    <div class="print-logo">
      <h2>🚢 Cubby Cargo — FCL Shipping Quote</h2>
      <p>Ramps Logistics Ltd · ramps.co.tt · All rates in USD</p>
    </div>
    <div class="quote-header">
      <div>
        <div class="quote-title">FCL Shipping Quote</div>
        <div class="quote-meta" id="q-customer"></div>
      </div>
      <div class="quote-ref" id="q-ref"></div>
    </div>
    <div class="quote-lane-banner">
      <span id="q-origin"></span>
      <span style="color:rgba(255,255,255,0.5)">→</span>
      <span class="dest" id="q-dest"></span>
    </div>
    <div class="quote-details">
      <div class="detail-block">
        <h4>Shipment Info</h4>
        <div id="q-shipment-lines"></div>
      </div>
      <div class="detail-block">
        <h4>Surcharge Breakdown</h4>
        <div id="q-surcharge-lines"></div>
      </div>
    </div>
    <div class="quote-totals">
      <div class="total-box no-ins">
        <div class="label">Base Rate (excl. insurance)</div>
        <div class="amount" id="q-total-no"></div>
      </div>
      <div class="total-box with-ins">
        <div class="label">🛡 Total with Insurance</div>
        <div class="amount" id="q-total-with"></div>
      </div>
    </div>
    <div class="ins-push">
      🛡 <strong>Protect your cargo for just USD $200.</strong> Marine Insurance covers a C&amp;F value of up to USD $30,000. Ask your Cubby representative to include it in your booking.
    </div>
    <div class="quote-notes">
      Rates are subject to space and equipment validity. Cargo must be ingated on or before the validity date; updated rates may apply thereafter.
      Values greater than USD $30,000 and restricted commodities must be quoted on a case-by-case basis.
      Should client decline Marine Insurance, Ramps Logistics shall not be held liable for any claims, loss or damages arising from the execution of services.
    </div>
    <button class="btn-pdf" onclick="window.print()">🖨 Print / Save as PDF</button>
  </div>
</div>

<script>
const LANES = {lanes_js};

(function() {{
  const now = new Date();
  const yy = now.getFullYear();
  const mm = String(now.getMonth()+1).padStart(2,'0');
  const dd = String(now.getDate()).padStart(2,'0');
  document.getElementById('quoteref').value = `QT-${{yy}}${{mm}}${{dd}}-${{String(Math.floor(Math.random()*900)+100)}}`;
}})();

document.getElementById('sel-dest').addEventListener('change', function() {{
  const dest = this.value;
  const laneEl = document.getElementById('sel-lane');
  const sizeEl = document.getElementById('sel-size');
  laneEl.innerHTML = '<option value="">Select lane</option>';
  sizeEl.innerHTML = '<option value="">Select size</option>';
  laneEl.disabled = true; sizeEl.disabled = true;
  if (!dest) return;
  LANES.filter(l => l.dest === dest).forEach(l => {{
    const opt = document.createElement('option');
    opt.value = l.lane; opt.textContent = l.lane;
    laneEl.appendChild(opt);
  }});
  laneEl.disabled = false;
}});

document.getElementById('sel-lane').addEventListener('change', function() {{
  const lane = this.value;
  const sizeEl = document.getElementById('sel-size');
  sizeEl.innerHTML = '<option value="">All sizes</option>';
  sizeEl.disabled = true;
  if (!lane) return;
  const laneData = LANES.find(l => l.lane === lane);
  if (!laneData) return;
  [...new Set(laneData.rates.map(r => r.container))].sort().forEach(s => {{
    const opt = document.createElement('option');
    opt.value = s.startsWith('20') ? '20' : '40';
    opt.textContent = s;
    sizeEl.appendChild(opt);
  }});
  sizeEl.disabled = false;
}});

let selectedRate = null;

function doSearch() {{
  const dest = document.getElementById('sel-dest').value;
  const lane = document.getElementById('sel-lane').value;
  const size = document.getElementById('sel-size').value;
  let rates = [];
  LANES.forEach(l => {{
    if (dest && l.dest !== dest) return;
    if (lane && l.lane !== lane) return;
    l.rates.forEach(r => {{
      if (size === '20' && !r.container.startsWith('20')) return;
      if (size === '40' && !r.container.startsWith('40')) return;
      rates.push({{ ...r, _lane: l.lane, _origin: l.origin, _dest: l.dest }});
    }});
  }});
  const rc = document.getElementById('results-card');
  const list = document.getElementById('results-list');
  document.getElementById('results-count').textContent = `(${{rates.length}} result${{rates.length !== 1 ? 's' : ''}})`;
  rc.style.display = 'block';
  if (!rates.length) {{ list.innerHTML = '<div class="no-results">No rates found.</div>'; return; }}
  list.innerHTML = rates.map((r, i) => {{
    const lines = Object.entries(r.surcharges || {{}}).map(([k,v]) =>
      `<div class="surcharge-line"><span>${{k}}</span><span>$${{v.toFixed(2)}}</span></div>`).join('');
    return `<div class="rate-row" id="rr-${{i}}" onclick="selectRate(${{i}})">
      <div><span class="tag">${{r.container}}</span><div style="font-weight:600;margin-top:4px;">${{r.pol}}</div><div class="rate-meta">→ ${{r.pod}}</div></div>
      <div><div class="rate-carrier">${{r.carrier}}</div>${{r.agent ? `<div class="rate-meta">${{r.agent}}</div>` : ''}}<div class="rate-meta" style="margin-top:4px;">⏱ ${{r.transit}}</div><div class="rate-meta">📅 ${{r.validity}}</div></div>
      <div class="surcharge-list">${{lines}}</div>
      <div class="totals">
        <div class="total-no">$${{(r.total_no_ins||0).toLocaleString('en-US',{{minimumFractionDigits:2,maximumFractionDigits:2}})}}<br><small>excl. ins.</small></div>
        <div class="total-with">$${{(r.total_with_ins||0).toLocaleString('en-US',{{minimumFractionDigits:2,maximumFractionDigits:2}})}}<span class="ins-note">🛡 insured</span></div>
      </div></div>`;
  }}).join('');
  window._searchRates = rates;
  selectedRate = null;
  document.getElementById('quote-card').style.display = 'none';
}}

function selectRate(i) {{
  document.querySelectorAll('.rate-row').forEach(el => el.classList.remove('selected'));
  document.getElementById('rr-' + i).classList.add('selected');
  selectedRate = window._searchRates[i];
  renderQuote(selectedRate);
}}

function fmt(v) {{ return v != null ? '$' + v.toLocaleString('en-US', {{minimumFractionDigits:2,maximumFractionDigits:2}}) : '—'; }}

function renderQuote(r) {{
  const customer = document.getElementById('customer').value || '—';
  const ref = document.getElementById('quoteref').value;
  const today = new Date().toLocaleDateString('en-GB', {{day:'2-digit',month:'short',year:'numeric'}});
  document.getElementById('q-customer').textContent = `Prepared for: ${{customer}} · Date: ${{today}}`;
  document.getElementById('q-ref').textContent = `Quote Ref: ${{ref}}`;
  document.getElementById('q-origin').textContent = r._origin;
  document.getElementById('q-dest').textContent = r._dest;
  document.getElementById('q-shipment-lines').innerHTML = [
    ['POL', r.pol], ['POD', r.pod], ['Container', r.container],
    ['Commodity', r.commodity || 'FAK'],
    ['Carrier', r.carrier + (r.agent ? ` / ${{r.agent}}` : '')],
    ['Transit Time', r.transit || '—'],
    ['Validity', r.validity || '—'],
  ].map(([k,v]) => `<div class="detail-line"><span>${{k}}</span><span class="val">${{v}}</span></div>`).join('');
  document.getElementById('q-surcharge-lines').innerHTML =
    Object.entries(r.surcharges || {{}}).map(([k,v]) =>
      `<div class="detail-line"><span>${{k}}</span><span class="val">$${{v.toFixed(2)}}</span></div>`
    ).join('') || '<div style="color:var(--muted);font-size:11px;">No breakdown available</div>';
  document.getElementById('q-total-no').textContent = fmt(r.total_no_ins);
  document.getElementById('q-total-with').textContent = fmt(r.total_with_ins);
  const qc = document.getElementById('quote-card');
  qc.style.display = 'block';
  qc.scrollIntoView({{behavior:'smooth', block:'start'}});
}}
</script>
</body>
</html>"""

    OUTPUT_QUOTE_PATH.parent.mkdir(parents=True, exist_ok=True)
    OUTPUT_QUOTE_PATH.write_text(html, encoding="utf-8")
    print(f"✅ Written to {OUTPUT_QUOTE_PATH}")


# ── Main ─────────────────────────────────────────────────────────────────────
def build():
    print(f"Reading {EXCEL_PATH}...")
    all_rates = []
    for sheet in SHEETS_TO_PROCESS:
        sheet_rates = parse_sheet(sheet)
        print(f"  {sheet}: {len(sheet_rates)} rate rows")
        all_rates.extend(sheet_rates)

    cards = build_cards(all_rates)

    OUTPUT_PATH.parent.mkdir(parents=True, exist_ok=True)
    html = render_html(cards)
    OUTPUT_PATH.write_text(html, encoding="utf-8")
    print(f"✅ Written to {OUTPUT_PATH} ({len(cards)} trade lanes, {sum(len(c['rate_rows']) for c in cards)} rate rows)")

    build_json(cards)
    build_quote(cards)


if __name__ == "__main__":
    build()
