# Performance and composition results

| Fixture | Nodes / edges | Slides | Time (s) | Primary backward | Crossings | Result |
|---|---:|---:|---:|---:|---:|---|
| Employee leave | 7 / 7 | 1 | 0.374 | 0 | 0 | Success |
| QES overview | 17 / 19 | 2 | 6.912 | 0 | 1 | Success |
| QES detailed | 28 / 30 | 3 | 21.931 | 0 | 19 | Success |
| Expense overview | 15 / 17 | 1 | 6.363 | 0 | 7 | Success |
| Expense detailed | 24 / 25 | 2 | 5.307 | 0 | 11 | Success |
| Supplier overview | 22 / 27 | 3 | 28.741 | 0 | 3 | Success |
| Supplier detailed | 32 / 37 | 3 | 17.335 | 0 | 11 | Success |
| Banking 34 | 34 / 45 | 3 | 25.305 | 0 | 26 | Success |
| Banking 51 | 51 / 70 | 6 | 40.591 | 0 | 15 | Success |
| Credit-card operational | 68 / 94 | — | 110.148 | — | — | Bounded failure |
| Credit-card business | 24 / 30 | 3 | 6.748 | 0 | 17 | Success |
| Corporate expansion | 28 / 35 | 6 | 3.614 | 0 | 6 | Success |

The exact machine-readable results, warnings, source hashes, and coverage records are in `acceptance-results.json`.

