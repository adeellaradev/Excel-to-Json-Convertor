import base64
import json
import sys
from datetime import datetime, date

import openpyxl


def broadcast_merged(ws):
    merged = {}
    for rng in ws.merged_cells.ranges:
        val = ws.cell(rng.min_row, rng.min_col).value
        for r in range(rng.min_row, rng.max_row + 1):
            for c in range(rng.min_col, rng.max_col + 1):
                merged[(r, c)] = val
    return merged


def get(ws, row, col, merged):
    return merged.get((row, col), ws.cell(row, col).value)


def clean(val):
    if isinstance(val, str):
        return " ".join(val.split()) or None
    return val


def serialise(obj):
    if isinstance(obj, (datetime, date)):
        return {"__datetime__": obj.isoformat()}
    raise TypeError(type(obj))


def convert(file_path):
    with open(file_path, "rb") as f:
        template_b64 = base64.b64encode(f.read()).decode()

    wb_raw = openpyxl.load_workbook(file_path, data_only=False)
    ws_raw = wb_raw.active

    cells = {}
    for row in ws_raw.iter_rows():
        for cell in row:
            cells[cell.coordinate] = cell.value

    wb = openpyxl.load_workbook(file_path, data_only=True)
    ws = wb.active
    merged = broadcast_merged(ws)

    schema = {}
    for col in range(2, ws.max_column + 1):
        name = clean(get(ws, 2, col, merged))
        if not name:
            continue
        schema[name] = {
            "category": clean(get(ws, 1, col, merged)),
            "focus":    clean(get(ws, 3, col, merged)),
            "source":   clean(get(ws, 4, col, merged)),
            "role":     clean(get(ws, 5, col, merged)),
            "col":      col,
        }

    records = []
    for row in range(1, ws.max_row + 1):
        date_val = ws.cell(row, 1).value
        if not isinstance(date_val, datetime):
            continue
        record = {"date": date_val.strftime("%Y-%m-%d"), "metrics": {}}
        for name, meta in schema.items():
            raw = ws.cell(row, meta["col"]).value
            if isinstance(raw, str) and raw.startswith("#"):
                raw = None
            record["metrics"][name] = {
                "value":    raw,
                "category": meta["category"],
                "focus":    meta["focus"],
                "source":   meta["source"],
                "role":     meta["role"],
            }
        records.append(record)

    return {
        "sheet_name":      ws_raw.title,
        "template_base64": template_b64,
        "cells":           cells,
        "records":         records,
    }


if __name__ == "__main__":
    input_file  = sys.argv[1] if len(sys.argv) > 1 else "Scoreboard Test.xlsx"
    output_file = sys.argv[2] if len(sys.argv) > 2 else "output.json"

    data = convert(input_file)

    with open(output_file, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2, default=serialise)

    print(f"Done: {len(data['records'])} records, {len(data['cells'])} cells -> {output_file}")
