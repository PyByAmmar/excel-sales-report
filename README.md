# excel-sales-report

A Python script that reads a messy Excel sales file, skips and reports bad rows, and writes a clean total-per-category report.

## What it does
Reads `sales_messy.xlsx`, adds up the amounts per category, and writes `report.xlsx`.
Rows with problems are skipped and listed, not guessed.

## How to run
```
python sales_report.py
```

## Example
Input (`sales_messy.xlsx`):
```
date category amount
2026-10-01 Food 12,75
2026-10-02 Sport
2026-10-03 13
2026-10-04 food 20
2026-10-05 Transport abc
```

Output (terminal):
```
Done
Processed 1 rows
Skipped 4 rows: row 2 (invalid amount), row 3 (empty amount), row 4 (missing category), row 6 (invalid amount)
```

Output (`report.xlsx`):
```
category total
Food 20.0
```

## Notes
- Category names are cleaned (`food` and `Food` are the same category).
- Amounts that are empty, not numbers, or ambiguous (like `12,75`) are skipped and reported, not guessed.
