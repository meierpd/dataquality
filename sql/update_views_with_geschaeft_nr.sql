-- Update views to include geschaeft_nr column
-- Database: GBI_REPORTING
-- Schema: gbi

USE GBI_REPORTING;
GO

-- Update the view for latest versions to include geschaeft_nr
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

-- Update the summary view to include geschaeft_nr
CREATE OR ALTER VIEW gbi.vw_orsa_analysis_summary AS
SELECT 
    institute_id,
    file_name,
    version,
    file_hash,
    geschaeft_nr,
    COUNT(*) as total_checks,
    SUM(CASE WHEN outcome_bool = 1 THEN 1 ELSE 0 END) as checks_passed,
    SUM(CASE WHEN outcome_bool = 0 THEN 1 ELSE 0 END) as checks_failed,
    CAST(SUM(CASE WHEN outcome_bool = 1 THEN 1 ELSE 0 END) AS FLOAT) / COUNT(*) as pass_rate,
    MAX(processed_timestamp) as last_processed
FROM gbi.orsa_analysis_data
GROUP BY institute_id, file_name, version, file_hash, geschaeft_nr;
GO

PRINT 'Views updated successfully with geschaeft_nr column';
GO
