-- Control objective: No single user holds two roles that together allow
-- them to both initiate and complete a financial transaction without
-- independent review (segregation of duties).
-- Test: Join each user's role(s) against the SoD conflict matrix, checking
-- both the primary role on the access list and any secondary role on the
-- dual-role assignment table, and return any user whose role combination
-- appears in the conflict matrix.

SELECT d.user_id, d.user_name, d.primary_role, d.secondary_role,
       m.conflict_reason
FROM dual_role_assignments d
JOIN sod_conflict_matrix m
ON (d.primary_role = m.role_a AND d.secondary_role = m.role_b)
OR (d.primary_role = m.role_b AND d.secondary_role = m.role_a);
