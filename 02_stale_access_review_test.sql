-- Control objective: User access is reviewed and recertified on a periodic
-- (at least semi-annual) basis by department management.
-- Test: Identify active users whose last_access_review is more than 6 months
-- before the audit date (2026-06-15). Any row returned is a control exception.

SELECT user_id, user_name, department, role, last_access_review
FROM user_access_list
WHERE account_status = 'Active'
AND last_access_review < DATE('2026-06-15', '-6 months');
