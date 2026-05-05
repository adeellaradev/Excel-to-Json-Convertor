# Excel → JSON Converter

Converts a clinic performance scoreboard `.xlsx` into a lossless, usable JSON file.

## How to run

```bash
pip install -r requirements.txt
python convert.py "Scoreboard Test.xlsx" output.json
```

## JSON shape

The output has two sections that serve different purposes.

**`template_base64`** — the original `.xlsx` file encoded as Base64. Decoding this gives back an exact binary replica with 100% of the original formatting: cell colors, fonts, borders, and merged cells. This is what makes the conversion truly lossless.

**`cells`** — every cell coordinate mapped to its raw value or formula string, used to re-inject edits when rebuilding the Excel file.

**`records`** — the human-readable, queryable section. Each entry is one week; each metric carries its value plus the full context from the spreadsheet's header rows (category, focus, source, role), so a developer or LLM can filter by `focus == "Financial"` without opening the file.

```json
{
  "sheet_name": "SCOREBOARD",
  "template_base64": "UEsDBBQAAAAIA...",
  "cells": { "B2": "Total Revenue - All Services", "B8": 40454.28 },
  "records": [
    {
      "date": "2026-02-16",
      "metrics": {
        "Total Revenue - All Services": {
          "value": 40454.28,
          "category": null,
          "focus": "Financial",
          "source": "EMR",
          "role": "J"
        }
      }
    }
  ]
}
```

## Decisions on the messy bits

**Merged cells:** `openpyxl` only returns a value for the top-left cell of a merged range. The script walks every merged range and broadcasts that value across all covered columns before reading headers, so no label is lost.

**Spacer columns:** Columns with no header in Row 2 are visual separators. They are skipped — they carry no data and would only add noise as keys.

**Formula results:** `data_only=True` returns cached formula results (real numbers) for the `records` section. `data_only=False` captures the raw formula strings for the `cells` section, keeping the rebuilt Excel fully functional.

**Excel errors (`#REF!`, `#DIV/0!`):** Converted to `null` so downstream consumers don't receive Excel-specific error strings.

**Non-data rows:** Rows 1–7 contain headers, targets, and annotations. Only rows where column A is a `datetime` are treated as weekly data — no hard-coded row numbers.

## With another two hours

- Parse Row 7 targets and embed a `"target"` field per metric so the JSON is self-contained for alerting.
- Infer value types (currency, percentage, count) from metric names and add a `"unit"` field for correct dashboard formatting.
- CLI flags: `--sheet`, `--start-row`, `--no-template` for teams that don't need the Base64 reconstruction layer.
