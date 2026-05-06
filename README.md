# Excel → JSON Converter

Converts a clinic performance scoreboard `.xlsx` into a meaningful, self-describing JSON file that an LLM, developer, or dashboard can use without needing the original spreadsheet.

## How to run

```bash
pip install -r requirements.txt
python convert.py "Scoreboard Test.xlsx" output.json
```

## JSON shape

```
output.json
├── about          — one sentence describing what this data is
├── schema         — what every metric IS: grouped by focus, with descriptions and units
└── records        — what every metric IS WORTH, week by week, grouped by the same focus areas
```

### `schema`

Organised by focus area. Each section has a human-readable `description` and lists every metric with its `source` system, `unit`, and `owner`. Defined once — not repeated in records.

```json
"schema": {
  "Financial": {
    "description": "Revenue billed across all service lines (PT, RMT, Chiropractic, Pelvic Health). Values in CAD.",
    "metrics": {
      "Total Revenue - All Services":  { "source": "EMR",   "unit": "CAD",          "owner": "J" },
      "PT Total Revenue":              { "source": "EMR",   "unit": "CAD",          "owner": "J" },
      "RMT Total Revenue":             { "source": "EMR",   "unit": "CAD",          "owner": "J" }
    }
  },
  "Phone": {
    "description": "Call centre performance via CallHero. Answer Rate and Booking Rate are 0–1 ratios (e.g. 0.93 = 93%). NBO = New Bookings (Online), DNB = Did Not Book.",
    "metrics": {
      "Total Calls (5530)": { "source": "CallHero", "unit": "count",       "owner": "J" },
      "Answer Rate":        { "source": "CallHero", "unit": "ratio (0–1)", "owner": "J" },
      "Booking Rate":       { "source": "CallHero", "unit": "ratio (0–1)", "owner": "J" },
      "NBO":                { "source": "CallHero", "unit": "count",       "owner": "J" },
      "DNB":                { "source": "CallHero", "unit": "count",       "owner": "J" }
    }
  },
  "AHS": {
    "description": "Alberta Health Services (government-funded) caseload — AHS assessments and treatments including Hip & Knee class.",
    "metrics": { ... }
  }
}
```

### `records`

One entry per week, sorted by date. Each record is grouped by the same focus area names as the schema. **Only metrics that have a value appear** — `null` (not entered that week) is omitted. Zero (`0`) is a real value and is always kept.

```json
{
  "date": "2026-02-02",
  "Financial": {
    "Total Revenue - All Services": 39202.17,
    "PT Total Revenue":             30394,
    "RMT Total Revenue":            775.5,
    "CHIRO Total Revenue":          5649.9,
    "Pelvic Health Total Revenue":  1215
  },
  "Collection": {
    "Total Payments - All Services": 40068.6,
    "AR > 90 days":                  60,
    "Total AR":                      57239.65
  },
  "Phone": {
    "Total Calls (5530)": 282,
    "Missed Calls":       16,
    "Answer Rate":        0.94,
    "Booking Rate":       0.78,
    "NBO":                33,
    "DNB":                7
  }
}
```

To understand what `Answer Rate: 0.94` means → look up `schema["Phone"]["description"]` or `schema["Phone"]["metrics"]["Answer Rate"]["unit"]` → `"ratio (0–1)"` → 94%.

### Example queries — one line each

```python
# Revenue trend across all weeks
[(r["date"], r["Financial"]["Total Revenue - All Services"]) for r in data["records"]]

# All phone metrics for a specific week
next(r for r in data["records"] if r["date"] == "2026-02-09").get("Phone", {})

# What does NBO mean and what unit is it?
data["schema"]["Phone"]["description"]                     # → "... NBO = New Bookings (Online) ..."
data["schema"]["Phone"]["metrics"]["NBO"]["unit"]          # → "count"

# Which focus areas had data entered on 2026-02-16?
[k for k in data["records"][-1] if k != "date"]

# All metrics in the AHS section and what they track
data["schema"]["AHS"]
```

## Decisions on the messy bits

**Abbreviations** — the source spreadsheet uses clinic shorthand throughout: `Ax` = Assessment, `Tx` = Treatment, `NAR` = New Assessment Request, `NBO` = New Bookings Online, `DNB` = Did Not Book, `PVA` = Per Visit Average, `AHS` = Alberta Health Services, `AR` = Accounts Receivable. These are explained in each focus area's `description` in the schema.

**Units** — inferred from metric names: metrics containing "Revenue", "AR", or "Payments" → `CAD`; metrics containing "Rate" or "%" → `ratio (0–1)`; everything else → `count`. This lets any consumer know whether `0.93` means 93% or 0.93 of some other scale.

**Merged cells** — `openpyxl` only exposes the value at the top-left of a merged range. The script broadcasts that value to all covered columns before reading headers, so no label is lost.

**Spacer columns** — columns with no name in Row 2 are visual separators. Skipped.

**Null omission** — 67 of 115 metrics were unfilled in the latest week. Keeping them as `null` everywhere made records unreadable. Only filled values appear in records; the full metric catalogue is in `schema`.

**Numeric strings** — some formula results came out as `"2.87"` (string). Cast to `float` so callers never encounter mixed types for the same metric.

**Phone sub-groups** — the spreadsheet split `Answer Rate` and `Booking Rate` into their own 1-metric focus labels. Both share `category: PHONE PERFORMANCE`, so they are merged under `Phone`.

**Formula results** — `data_only=True` returns the last cached formula result. `Revenue Collected (4 wk avg)` comes out as a number, not `=B8/B7`.

**Excel errors (`#REF!`, `#DIV/0!`)** — treated as null (omitted from records).

## With another two hours

- Parse the target row and embed `"target"` per metric so the JSON supports alerting without the spreadsheet.
- Replace inferred units with explicit ones mapped per metric name for higher accuracy.
- Add `--sheet` and `--start-row` CLI flags for different layouts.

## Round-trip: JSON → Excel

The two limitations of any Excel → JSON conversion are formula strings (we store computed results, not `=B8/B7`) and visual formatting (colors, borders, merged cell styles).

Both are fully solvable. Keep the original `.xlsx` as a formatting template alongside the JSON. A companion `json_to_excel.py` script pours the JSON data back into that template — you get a 100% exact replica of the original spreadsheet, formatting and all, with the added benefit that the data passed through a clean, queryable JSON layer in between.
