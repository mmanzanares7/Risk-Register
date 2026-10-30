# IT General Controls (ITGC) Testing: Access Management and Change Management

## Project Description

This project simulates an IT general controls audit over a company's financial ERP environment, covering user access management, periodic access review, change management, and segregation of duties. Using synthetic data and SQL, I designed and executed control tests, documented control objectives and test steps in workpaper format, and identified exceptions that would be reported to management. This reflects the type of testing performed in a SOX ITGC audit or an internal IT risk assessment.

## Scope

The audit covers four controls over a company's ERP system (AP, GL, and Procurement modules) and payroll system, based on a sample population of 16 users and 8 production changes over a two-month period.

| Data file | Description |
|---|---|
| [`data/user_access_list.csv`](data/user_access_list.csv) | Active and terminated users, their role, access level, and last access review date |
| [`data/dual_role_assignments.csv`](data/dual_role_assignments.csv) | Users holding more than one system role |
| [`data/sod_conflict_matrix.csv`](data/sod_conflict_matrix.csv) | Role combinations considered a segregation of duties conflict |
| [`data/change_management_log.csv`](data/change_management_log.csv) | Production changes, requester, approver, approval date, and implementation date |

## Control 1: Timely Termination of Access

**Control objective:** Access to financial systems is removed timely upon termination.

**Test script:** [`scripts/01_terminated_user_access_test.sql`](scripts/01_terminated_user_access_test.sql)

**Test steps:** Query the user access list for any record with a termination date populated where the account status still shows Active.

**Result:** 2 exceptions identified.

| user_id | user_name | termination_date | account_status |
|---|---|---|---|
| U1003 | R. Foster | 2026-05-15 | Active |
| U1010 | B. Kessler | 2026-06-02 | Active |

**Conclusion:** Control did not operate effectively. Two terminated employees retained active system access at the time of testing, representing unauthorized access risk to the AP module.

## Control 2: Periodic Access Review

**Control objective:** User access is reviewed and recertified by department management at least semi-annually.

**Test script:** [`scripts/02_stale_access_review_test.sql`](scripts/02_stale_access_review_test.sql)

**Test steps:** Query active users whose last access review date is more than six months prior to the audit date (2026-06-15).

**Result:** 2 exceptions identified (U1010, review dated 2025-11-01; U1016, review dated 2025-09-01).

**Conclusion:** Control did not operate effectively for all users in the population. Recommend management complete overdue reviews and reinforce the review cadence.

## Control 3: Change Management Approval

**Control objective:** All production changes require independent approval prior to implementation, and no individual approves their own change.

**Test script:** [`scripts/03_unapproved_change_test.sql`](scripts/03_unapproved_change_test.sql)

**Test steps:** Query the change log for changes missing an approver or approval date, changes approved after they were already implemented, or changes where the requester and approver are the same person.

**Result:** 4 exceptions identified.

| change_id | issue |
|---|---|
| CH-2203 | Implemented with no approver or approval date on file |
| CH-2206 | Approval date (2026-05-18) is after the implementation date (2026-05-14) |
| CH-2207 | Implemented with no approver or approval date on file |
| CH-2208 | Requester and approver are the same individual (self-approved) |

**Conclusion:** Control did not operate effectively. Recommend enforcing a system-level control that blocks implementation until an independent approval is logged, and restricting users from approving their own change requests.

## Control 4: Segregation of Duties

**Control objective:** No single user holds two roles that together allow them to both initiate and complete a financial transaction without independent review.

**Test script:** [`scripts/04_segregation_of_duties_test.sql`](scripts/04_segregation_of_duties_test.sql)

**Test steps:** Join the dual-role assignment table against the segregation of duties conflict matrix to identify any user whose combination of roles is flagged as conflicting.

**Result:** 1 exception identified.

| user_id | user_name | primary_role | secondary_role | conflict_reason |
|---|---|---|---|---|
| U1016 | N. Castillo | AP Clerk | Purchasing Agent | Can create a purchase order and process the resulting payment, enabling fictitious vendor payments |

**Conclusion:** Control did not operate effectively. Recommend removing the Purchasing Agent role from this user or implementing a compensating control, such as a supervisor review of all payments processed by users with dual AP/Procurement access.

## Summary

Across four ITGC tests covering access termination, periodic access review, change management approval, and segregation of duties, I identified exceptions in every control area: two terminated users with active access, two overdue access reviews, four change management approval gaps, and one segregation of duties conflict. Each finding was documented with the control objective, test steps, query used, and results, reflecting the workpaper format used in a SOX or internal ITGC audit, and each includes a recommendation a company could act on to remediate the underlying risk.
