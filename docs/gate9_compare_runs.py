import json
import sys
from pathlib import Path
from openpyxl import load_workbook

first, second = map(Path, sys.argv[1:])

required = [
    "accounting_analysis.xlsx",
    "analysis.log",
    "tests.log",
    "validation.json",
    "validation.log",
    "ai_report.md",
    "codex.log",
]

files_ok = all(
    (folder / name).is_file()
    for folder in (first, second)
    for name in required
)

def workbook_snapshot(path):
    book = load_workbook(path, data_only=False)
    try:
        result = []
        for ws in book.worksheets:
            cells = tuple(
                (c.coordinate, c.value, c.data_type,
                 c.style_id, c.number_format)
                for row in ws for c in row
            )
            widths = tuple(
                (key, dim.width)
                for key, dim in sorted(ws.column_dimensions.items())
            )
            result.append(
                (ws.title, ws.freeze_panes,
                 ws.auto_filter.ref, cells, widths)
            )
        return result
    finally:
        book.close()

json1 = json.loads(
    (first / "validation.json").read_text(encoding="utf-8")
)
json2 = json.loads(
    (second / "validation.json").read_text(encoding="utf-8")
)

json_match = json1 == json2
excel_match = workbook_snapshot(
    first / "accounting_analysis.xlsx"
) == workbook_snapshot(
    second / "accounting_analysis.xlsx"
)

print("Both run folders have required files:", files_ok)
print("Validation JSON matches:", json_match)
print("Excel values and styles match:", excel_match)

if not (files_ok and json_match and excel_match):
    sys.exit(1)

print("GATE 9-2: PASS")