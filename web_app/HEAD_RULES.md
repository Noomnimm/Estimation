# Confirmed head rules (2026-08-28)

Insulator counts are informational, per entered head, summed across pages. They
do not alter BaseData, SET contents or material exports.

| Head | Upright | Horizontal |
|---|---:|---:|
| BA / BA.AL | 4 | 12 |
| DE / DE.AL | 0 | 12 |
| DDE / DDE.AL | 6 | 24 |
| DDE.BL | 0 | 24 |
| SP (including บน,ล่าง) | 3 | 0 |
| DP | 6 | 0 |
| CCB / paired CCB | 3 / 6 | 0 |
| BA 1-P / BA.AL 1P | 2 | 8 |
| DDE 1-P | 4 | 16 |
| DDE.BL 1-P | 0 | 16 |
| DE.CON 1-P | 4 | 8 |
| SP 1-P | 2 | 0 |
| BA.SLK | 6 | 0 |
| LAT.SLK (บน or ล่าง) | 6 | 12 |
| CTB (including paired) / CSC | 0 | 0 |
| DDE,DP.st 3.0m | 12 | 24 |
| DDE.st 3m, LAT.SLK | 12 | 36 |
| 2BA.st 4.5m+DE.CON | 12 | 36 |
| 2DE.st4.5 + DE.CON | 6 | 36 |

2BA, 2DE, 2DDE, 2SP, 2DP (including .st) use twice the base head's
insulator rate. Exact combined-head and single-phase overrides take precedence.

Wire-dependent additions: 2BA/2DE/2DDE use a multiplier of 2; 1-P/1P
use 2/3. These factors apply to Preform/Strain, Clevis, BA connectors,
and DDE tensionless/tapes, not to BaseData material quantities.
The two confirmed DDE mixed heads use the ordinary DDE wire rule.

Pending, intentionally unchanged:
- Wire additions for +DE.CON double-head assemblies remain absent until confirmed.
- LAT.SLK has no separate wire additions; BA.SLK retains its existing BA wire rule.
- SET insulator quantities can differ from the informational totals. No workbook
  data was changed to reconcile them.
- BaseData contains separately spelled 2SP names with comma-spacing differences;
  these are recognized for informational counts but their source rows are not merged.

Tests: `python -m unittest web_app.test_head_rules` and
`node web_app/test_insulators.cjs`.
