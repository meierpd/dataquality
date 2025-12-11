-- Migration script to update views to include berichtsjahr column
-- Database: GBI_REPORTING
-- Schema: gbi
--
-- This script recreates the views vw_orsa_analysis_latest and vw_orsa_analysis_summary
-- to include the berichtsjahr column in their output.

USE GBI_REPORTING;
GO

-- Recreate the view for latest versions only (with berichtsjahr)
CREATE OR ALTER VIEW gbi.vw_orsa_analysis_latest AS
WITH LatestVersions AS (
    SELECT 
        institute_id,
        file_name,
        MAX(version) as max_version
    FROM gbi.orsa_analysis_data
    GROUP BY institute_id, file_name
)
SELECT 
    o.*
FROM gbi.orsa_analysis_data o
INNER JOIN LatestVersions lv
    ON o.institute_id = lv.institute_id
    AND o.file_name = lv.file_name
    AND o.version = lv.max_version;
GO

-- Recreate the view for summary statistics (with berichtsjahr)
CREATE OR ALTER VIEW gbi.vw_orsa_analysis_summary AS
SELECT 
    institute_id,
    file_name,
    version,
    file_hash,
    geschaeft_nr,
    berichtsjahr,
    COUNT(*) as total_checks,
    SUM(CASE WHEN outcome_bool = 1 THEN 1 ELSE 0 END) as checks_passed,
    SUM(CASE WHEN outcome_bool = 0 THEN 1 ELSE 0 END) as checks_failed,
    CAST(SUM(CASE WHEN outcome_bool = 1 THEN 1 ELSE 0 END) AS FLOAT) / COUNT(*) as pass_rate,
    MAX(processed_timestamp) as last_processed
FROM gbi.orsa_analysis_data
GROUP BY institute_id, file_name, version, file_hash, geschaeft_nr, berichtsjahr;
GO

PRINT 'Views vw_orsa_analysis_latest and vw_orsa_analysis_summary updated successfully';
PRINT 'Both views now include the berichtsjahr column';
GO
