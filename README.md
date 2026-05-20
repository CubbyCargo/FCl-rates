# FCL Rate Publisher

Converts the weekly FCL Excel tariff into a live JSON endpoint that Cubby (Respond.io AI Agent) can query in real time.

---

## How to update rates

1. Open the repo in GitHub
2. Replace `Customer_Rate_Tariff_Template_Week_XX_XXXX.xlsx` with your new file  
   *(keep the filename or update the filename reference in `convert_rates.py` line 11)*
3. Go to **Actions → Publish FCL Rates → Run workflow**
4. Optionally add a note (e.g. "Week 22 update") and click **Run workflow**
5. Wait ~60 seconds — rates are live ✅

---

## Live rates URL

Once GitHub Pages is enabled, your rates will be available at:

```
https://<your-username>.github.io/<repo-name>/rates.json
```

Give this URL to Cubby as its knowledge source.

---

## JSON structure

```json
{
  "generated_at": "2026-05-19T18:00:00Z",
  "source_file": "Customer_Rate_Tariff_Template_Week_21_2026.xlsx",
  "destinations": {
    "TT": {
      "USA / TRINIDAD": [
        {
          "pol": "Port Everglades",
          "pod": "Point Lisas",
          "size": "20ft",
          "of_bunker": 1800,
          "total_without_insurance": 2310.5,
          "insurance": 200,
          "total_with_insurance": 2510.5,
          "transit_time": "7 days",
          "validity": "31/05/2026",
          "carrier": "King Ocean",
          "comment": null
        }
      ]
    }
  }
}
```

**Sheets / destination keys:**
| Key | Destination |
|-----|-------------|
| `TT` | Trinidad |
| `GUY` | Guyana |
| `SUR` | Suriname |
| `Trinidad Exports` | TT outbound (Guyana, Suriname, Barbados, Jamaica, DR) |
| `Print FE-TT` | Far East → Trinidad |
| `COL` | Colombia |

---

## Files

| File | Purpose |
|------|---------|
| `convert_rates.py` | Parses Excel → `rates.json` |
| `rates.json` | Published output (auto-generated, do not edit manually) |
| `.github/workflows/publish_rates.yml` | GitHub Actions workflow |
