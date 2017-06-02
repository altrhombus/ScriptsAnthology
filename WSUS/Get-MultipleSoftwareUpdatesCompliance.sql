-- ConfigMgr WSUS Query to pull information about multiple updates at the same time
-- Faster than running the "Compliance 2 - Specific software update" report over and over again :)
-- Written by Jacob Thornberry
-- Made with ❤️ in MKE

-- How to use:
-- Edit the collections you'd like to filter by (COLLECTION_NAME FILTER %). For example, if your update collections that you're interested in all start with 'Workstation', enter 'Workstation %'
-- Edit the v_updateinfo.title with the full name of the update. Add as many as you'd like!
-- Execute the query, and enjoy all that extra free time.

SELECT 	dbo.v_updateinfo.title AS 'Title', 
		dbo.v_update_deploymentSummary_Live.CollectionName AS 'Collection Name', 
		dbo.v_update_deploymentSummary_Live.NumTotal AS 'Total', 
		dbo.v_update_deploymentSummary_Live.NumPresent AS 'Installed', 
		dbo.v_update_deploymentSummary_Live.NumMissing AS 'Required', 
		dbo.v_update_deploymentSummary_Live.NumNotApplicable AS 'Not Required', 
		dbo.v_update_deploymentSummary_Live.NumUnknown AS 'Unknown', 
		ROUND ((CAST (dbo.v_update_deploymentSummary_Live.NumPresent AS FLOAT) + CAST (dbo.v_update_deploymentSummary_Live.NumNotApplicable AS FLOAT)) / CAST (dbo.v_update_deploymentSummary_Live.NumTotal AS FLOAT) * 100, 2) AS '% Compliant',
		ROUND (CAST (dbo.v_update_deploymentSummary_Live.NumMissing AS FLOAT) / CAST (dbo.v_update_deploymentSummary_Live.NumTotal AS FLOAT) * 100, 2) AS '% Not Compliant',
		ROUND (CAST (dbo.v_update_deploymentSummary_Live.NumUnknown AS FLOAT) / CAST (dbo.v_update_deploymentSummary_Live.NumTotal AS FLOAT) * 100, 2) AS '% Unknown'
FROM dbo.v_update_deploymentSummary_Live
INNER JOIN dbo.v_updateinfo ON dbo.v_UpdateInfo.CI_ID = v_update_deploymentSummary_Live.CI_ID
WHERE dbo.v_update_deploymentSummary_Live.CollectionName LIKE 'COLLECTION_NAME FILTER %' 
	AND (dbo.v_updateinfo.title = 'AN_UPDATE (KB1337)'
	OR dbo.v_updateinfo.title = 'ANOTHER_UPDATE (KB1338)'
	OR dbo.v_updateinfo.title = 'AND_SO_ON (KB1339)')