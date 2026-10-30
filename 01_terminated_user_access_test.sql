-- Control objective: Access to financial systems is removed timely upon termination.
-- Test: Identify any user with a termination_date populated whose account_status
-- is still Active. Any row returned is a control exception.

SELECT user_id, user_name, department, role, account_status,
       termination_date, last_access_review
FROM user_access_list
WHERE termination_date IS NOT NULL
AND account_status = 'Active';
