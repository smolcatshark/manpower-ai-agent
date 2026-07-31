# v0.8.4 Regression Report

Tested application:
`manpower_flexible_parser_v0_8_4_final_export_review.py`

## Source checks

- Python syntax: passed
- Stable interface retained: passed
- No v0.8.5 competition banner included: passed

## 30 July 2026 PDF set

| Trade | Today Total | Worker | Detail total | Result |
|---|---:|---:|---:|---|
| AC | 29 | 23 | 23 | Passed |
| EL | 53 | 46 | 46 | Passed |
| FS | 24 | 19 | 19 | Passed |
| PD | 63 | 54 | 54 | Passed |

## 31 July 2026 PDF set

| Trade | Today Total | Worker | Detail total | Result |
|---|---:|---:|---:|---|
| AC | 33 | 27 | 27 | Passed |
| EL | 45 | 38 | 38 | Passed |
| FS | 25 | 20 | 20 | Passed |
| PD | 63 | 54 | 54 | Passed |

## 31 July FS location checks

- T1 floor unspecified: 1
- T2 floor unspecified: 3
- T3 floor unspecified: 1
- Podium 3F: 3
- Basement floor unspecified: 7
- T2 + T3 distribution unspecified: 2
- Basement + T3 distribution unspecified: 2
- Basement + Podium distribution unspecified: 1
- Total FS Worker: 20

## Excel checks

- Create a new workbook from the built-in template: passed
- Update the same workbook with a later date: passed
- Preserve the earlier date after update: passed
- Required six worksheets present: passed
- Daily Master formulas present: passed
- Location Detail rows generated: passed
- Cross-F and distribution-U rows generated: passed
- Duplicate date/trade key detection: passed
- Existing history preserved while adding 31 July: passed

## Result

v0.8.4 is accepted as the stable application base for the Windows
desktop packaging phase.
