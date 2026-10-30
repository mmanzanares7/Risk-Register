# Risk Register and Vulnerability Tracker

## Project Description

This project is a risk register tracking ten identified risks across vulnerability management, access control, segregation of duties, third-party risk, incident response, business continuity, data protection, and change management. Each risk is scored using a standard 5x5 likelihood/impact matrix, both before and after mitigation, and a Python script summarizes the register by risk rating band. This reflects the kind of risk tracking and prioritization a GRC analyst maintains to report risk posture to management and track remediation to closure.

## Risk register

**File:** [`risk_register.csv`](risk_register.csv)

Each risk includes:

- **Inherent risk** - likelihood and impact scored 1-5 before any mitigation, multiplied for an inherent risk score (1-25)
- **Mitigation** - the action being taken or planned to reduce the risk
- **Residual risk** - likelihood and impact re-scored assuming the mitigation is fully in place, giving a residual risk score
- **Owner and target date** - who is accountable and when the mitigation is due
- **Status** - Open, In Progress, or Closed

## Scoring methodology

Risk score = likelihood x impact, using a 1-5 scale for each, giving a score from 1 to 25. Scores are grouped into four rating bands:

| Score range | Rating |
|---|---|
| 15-25 | Critical |
| 10-14 | High |
| 5-9 | Medium |
| 1-4 | Low |

This matches the likelihood/impact scoring approach commonly used in enterprise risk management and GRC platforms, where risk owners and remediation deadlines are tracked against each rated risk until it's closed or accepted.

## Summary script

**File:** [`risk_summary.py`](risk_summary.py)

Run with:

```
python3 risk_summary.py
```

The script reads the register and reports:

- How many risks fall into each rating band, before and after mitigation
- The shift in overall risk posture once planned mitigations are complete (in this register, inherent risk shows 3 Critical and 4 High risks, dropping to 0 Critical and 2 High residual risks once mitigations land)
- A callout list of any risk still Open or In Progress that is rated Critical or High, which is what a risk owner or auditor would prioritize first

## Connection to other projects

Three risks in this register (R-002 terminated user access, R-004 segregation of duties, R-010 change management approval) correspond directly to the exceptions identified in the [ITGC controls testing project](../itgc-controls-testing), showing how an audit finding flows into a tracked, owned, and dated remediation item rather than stopping at the finding itself.

## Summary

This project tracks ten risks across common GRC domains using a standard likelihood/impact scoring matrix, capturing both inherent and residual risk alongside ownership and remediation deadlines. The accompanying script summarizes the register by rating band and surfaces which open risks carry the highest priority, the kind of reporting a GRC analyst would run ahead of a risk committee meeting or an audit status update.

## Related project

An Excel version of this register, with formulas, heat maps and a summary dashboard, is in [`risk-register-excel`](https://github.com/mmanzanares7/risk-register-excel).
