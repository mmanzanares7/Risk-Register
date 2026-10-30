# Excel Risk Register with Heat Map

## Project Description

This project is a formula-driven risk register in Excel that consolidates 15 risks from three sources: the exceptions found in my ITGC testing, the threats identified in my vulnerability assessment, and an enterprise risk list. Each risk is scored on a 5x5 likelihood and impact matrix before and after mitigation. The workbook calculates scores and ratings, flags overdue mitigations, plots inherent and residual heat maps, and ranks the top risks, so the register can be maintained and reported to management without rebuilding anything by hand. This is the kind of register a GRC or IT audit analyst keeps to track risks from finding through remediation.

## Files

| File | Purpose |
|---|---|
| [`risk_register.xlsx`](risk_register.xlsx) | The workbook: Register, Heat Map, Summary, Scale and Legend sheets (248 formulas, no errors) |
| [`risk_register_export.csv`](risk_register_export.csv) | The register with calculated values, so it can be read on GitHub without opening Excel |
| [`images/heat_map.png`](images/heat_map.png) | Heat maps taken from the workbook values |
| [`scripts/build_workbook.py`](scripts/build_workbook.py) | Python (openpyxl) script that generates the workbook |

## How the workbook works

- **Register:** one row per risk. Blue cells are inputs. Inherent score, residual score, ratings, risk reduction, days to target, the OVERDUE flag and a ranking helper are formulas. Dropdowns restrict scores to 1 to 5 and status to Open, In Progress or Closed.
- **Heat Map:** two 5x5 grids that count risks at each likelihood and impact, using COUNTIFS. Cell colours follow the rating thresholds through conditional formatting.
- **Summary:** counts by rating, status, source and category, average scores, overdue count, the top 5 risks (INDEX, MATCH and LARGE), and a chart of inherent against residual ratings.
- **Scale:** rating thresholds, the report as-of date, and definitions for each likelihood and impact level. Changing a threshold or the date updates every sheet.

Rating bands: Low 1 to 4, Medium 5 to 9, High 10 to 14, Critical 15 to 25. These match [`risk-register-tracker`](https://github.com/mmanzanares7/risk-register-tracker).

## Where the risks come from

- **R-01 to R-04** are the four ITGC control failures from [`itgc-controls-testing`](https://github.com/mmanzanares7/itgc-controls-testing): terminated users with active access, overdue access reviews, production changes without independent approval, and an AP and Purchasing segregation of duties conflict. Each description cites the user IDs or change IDs found in testing.
- **R-05 to R-08** come from [`vulnerability-assessment-report`](https://github.com/mmanzanares7/vulnerability-assessment-report): the publicly reachable customer database, automated scanning, insider misuse and deprecated SSL. That assessment scored on a 1 to 3 scale, so I rescored these on 1 to 5 and kept the original scores in each description. The new scores are my judgment, not measured data.
- **R-09 to R-15** are the enterprise risks from [`risk-register-tracker`](https://github.com/mmanzanares7/risk-register-tracker).

## Register

| ID | Risk | Source | Inherent | Residual | Status |
|---|---|---|---|---|---|
| R-01 | Terminated users retain active ERP access | ITGC Testing | 16 (Critical) | 4 (Low) | In Progress |
| R-02 | Periodic access reviews overdue | ITGC Testing | 9 (Medium) | 3 (Low) | In Progress |
| R-03 | Production changes without independent approval | ITGC Testing | 16 (Critical) | 8 (Medium) | Open |
| R-04 | AP and Purchasing roles held by one user | ITGC Testing | 12 (High) | 4 (Low) | Open |
| R-05 | Customer database reachable from the public internet | Vulnerability Assessment | 20 (Critical) | 5 (Medium) | Open |
| R-06 | Automated scanners find the exposed database | Vulnerability Assessment | 15 (Critical) | 3 (Low) | Open |
| R-07 | Employee alters or deletes customer records | Vulnerability Assessment | 6 (Medium) | 3 (Low) | Open |
| R-08 | Database connections use deprecated SSL | Vulnerability Assessment | 9 (Medium) | 3 (Low) | Open |
| R-09 | Unpatched ERP server vulnerability | Enterprise Register | 20 (Critical) | 10 (High) | In Progress |
| R-10 | No MFA on remote VPN access | Enterprise Register | 15 (Critical) | 5 (Medium) | Open |
| R-11 | No tested ransomware response runbook | Enterprise Register | 15 (Critical) | 10 (High) | Open |
| R-12 | Shared admin credentials on legacy file server | Enterprise Register | 12 (High) | 3 (Low) | Open |
| R-13 | Backup restoration untested for over a year | Enterprise Register | 10 (High) | 5 (Medium) | Closed |
| R-14 | Payroll data in unencrypted spreadsheets | Enterprise Register | 12 (High) | 4 (Low) | Open |
| R-15 | Critical vendor risk assessments out of date | Enterprise Register | 9 (Medium) | 6 (Medium) | Open |

## Heat maps

![Inherent and residual heat maps](images/heat_map.png)

## Results

- Inherent exposure is concentrated at the top: 7 Critical, 4 High and 4 Medium risks, with the highest scores (20) on the exposed customer database (R-05) and the unpatched ERP server (R-09).
- If every planned mitigation is completed, no risk remains Critical, 2 stay High (R-09 and R-11) and the average score falls from 13.1 to 5.1, an average reduction of 60%.
- One mitigation is overdue as of the report date: R-01, removing terminated users' access, due 2026-10-15.
- Four of the 15 risks trace directly to control failures I tested, which shows how audit exceptions feed a risk register.

## Limitations

All data is synthetic. Likelihood and impact scores are qualitative judgments, and residual scores are targets that assume mitigations are fully implemented. The report as-of date is fixed at 2026-10-30 so results can be reproduced; replace it with `=TODAY()` on the Scale sheet for a live view.

## Summary

I built an Excel risk register that turns audit and assessment findings into scored, ranked and tracked risks, with heat maps and a management summary driven entirely by formulas. It shows how I connect control testing, risk assessment and remediation tracking in one working tool.
