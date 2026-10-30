-- Control objective: All production changes require independent approval
-- prior to implementation, and no individual approves their own change.
-- Test: Identify changes with a missing approver, a missing approval date,
-- an approval date after the implementation date, or a case where the
-- requester and approver are the same person.

SELECT change_id, system, description, requested_by, approved_by,
       approval_date, implementation_date, status
FROM change_management_log
WHERE approved_by IS NULL
OR approval_date IS NULL
OR approval_date > implementation_date
OR requested_by = approved_by;
