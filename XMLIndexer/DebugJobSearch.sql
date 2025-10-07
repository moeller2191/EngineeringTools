-- Quick test to check if IK3NC-0000 is in the database
SELECT 'Checking for IK3NC-0000...' as Test;

SELECT COUNT(*) as TotalJobs FROM MrpPriorityList;

SELECT * FROM MrpPriorityList WHERE JobNumber LIKE '%IK3NC%';

SELECT * FROM MrpPriorityList WHERE JobNumber = 'IK3NC-0000';

SELECT 'All jobs containing 0000:' as Test;
SELECT JobNumber, PartNumber, Description FROM MrpPriorityList WHERE JobNumber LIKE '%0000%';