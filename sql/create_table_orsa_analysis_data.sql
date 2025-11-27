-- Table schema for ORSA Analysis Quality Control Results
-- Database: GBI_REPORTING
-- Schema: gbi
-- Table: orsa_analysis_data

USE GBI_REPORTING;
GO

-- Drop table if exists (for development/testing)
-- IF OBJECT_ID('gbi.orsa_analysis_data', 'U') IS NOT NULL
--     DROP TABLE gbi.orsa_analysis_data;
-- GO

-- Create schema if it doesn't exist
IF NOT EXISTS (SELECT * FROM sys.schemas WHERE name = 'gbi')
BEGIN
    EXEC('CREATE SCHEMA gbi');
END;
GO

-- Create the orsa_analysis_data table
CREATE TABLE gbi.orsa_analysis_data (
    -- Primary key
    id BIGINT IDENTITY(1,1) PRIMARY KEY,
    
    -- File identification
    institute_id NVARCHAR(50) NOT NULL,
    file_name NVARCHAR(255) NOT NULL,
    file_hash VARCHAR(64) NOT NULL,  -- SHA-256 hash
    version INT NOT NULL,
    geschaeft_nr NVARCHAR(50) NULL,  -- Business case number from source document
    
    -- Check information
    check_name NVARCHAR(100) NOT NULL,
    check_description NVARCHAR(500) NULL,
    
    -- Check outcomes
    outcome_bool BIT NOT NULL,  -- Pass (1) or Fail (0)
    outcome_numeric FLOAT NULL,  -- Optional numeric outcome
    
    -- Metadata
    processed_timestamp DATETIME2 NOT NULL DEFAULT GETDATE(),
    
    -- Indexes for common queries
    INDEX idx_institute_id (institute_id),
    INDEX idx_file_hash (file_hash),
    INDEX idx_check_name (check_name),
    INDEX idx_processed_timestamp (processed_timestamp),
    INDEX idx_file_version (institute_id, file_name, version),
    INDEX idx_geschaeft_nr (geschaeft_nr)
);
GO

-- Create a view for latest versions only
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

-- Create a view for summary statistics
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

-- Add comments (SQL Server extended properties)
EXEC sp_addextendedproperty 
    @name = N'MS_Description', 
    @value = N'Stores quality control check results for ORSA documents', 
    @level0type = N'SCHEMA', @level0name = 'gbi',
    @level1type = N'TABLE', @level1name = 'orsa_analysis_data';
GO

EXEC sp_addextendedproperty 
    @name = N'MS_Description', 
    @value = N'Unique identifier for each check result', 
    @level0type = N'SCHEMA', @level0name = 'gbi',
    @level1type = N'TABLE', @level1name = 'orsa_analysis_data',
    @level2type = N'COLUMN', @level2name = 'id';
GO

EXEC sp_addextendedproperty 
    @name = N'MS_Description', 
    @value = N'Identifier for the financial institution', 
    @level0type = N'SCHEMA', @level0name = 'gbi',
    @level1type = N'TABLE', @level1name = 'orsa_analysis_data',
    @level2type = N'COLUMN', @level2name = 'institute_id';
GO

EXEC sp_addextendedproperty 
    @name = N'MS_Description', 
    @value = N'Name of the Excel file being checked', 
    @level0type = N'SCHEMA', @level0name = 'gbi',
    @level1type = N'TABLE', @level1name = 'orsa_analysis_data',
    @level2type = N'COLUMN', @level2name = 'file_name';
GO

EXEC sp_addextendedproperty 
    @name = N'MS_Description', 
    @value = N'SHA-256 hash of the file for version tracking', 
    @level0type = N'SCHEMA', @level0name = 'gbi',
    @level1type = N'TABLE', @level1name = 'orsa_analysis_data',
    @level2type = N'COLUMN', @level2name = 'file_hash';
GO

EXEC sp_addextendedproperty 
    @name = N'MS_Description', 
    @value = N'Version number (increments when same file changes)', 
    @level0type = N'SCHEMA', @level0name = 'gbi',
    @level1type = N'TABLE', @level1name = 'orsa_analysis_data',
    @level2type = N'COLUMN', @level2name = 'version';
GO

EXEC sp_addextendedproperty 
    @name = N'MS_Description', 
    @value = N'Business case number (Geschäftsnummer) from source document', 
    @level0type = N'SCHEMA', @level0name = 'gbi',
    @level1type = N'TABLE', @level1name = 'orsa_analysis_data',
    @level2type = N'COLUMN', @level2name = 'geschaeft_nr';
GO

EXEC sp_addextendedproperty 
    @name = N'MS_Description', 
    @value = N'Name/identifier of the quality check', 
    @level0type = N'SCHEMA', @level0name = 'gbi',
    @level1type = N'TABLE', @level1name = 'orsa_analysis_data',
    @level2type = N'COLUMN', @level2name = 'check_name';
GO

EXEC sp_addextendedproperty 
    @name = N'MS_Description', 
    @value = N'Detailed description of what the check verified', 
    @level0type = N'SCHEMA', @level0name = 'gbi',
    @level1type = N'TABLE', @level1name = 'orsa_analysis_data',
    @level2type = N'COLUMN', @level2name = 'check_description';
GO

EXEC sp_addextendedproperty 
    @name = N'MS_Description', 
    @value = N'Boolean outcome: 1 = Pass, 0 = Fail', 
    @level0type = N'SCHEMA', @level0name = 'gbi',
    @level1type = N'TABLE', @level1name = 'orsa_analysis_data',
    @level2type = N'COLUMN', @level2name = 'outcome_bool';
GO

EXEC sp_addextendedproperty 
    @name = N'MS_Description', 
    @value = N'Optional numeric outcome (e.g., count, percentage)', 
    @level0type = N'SCHEMA', @level0name = 'gbi',
    @level1type = N'TABLE', @level1name = 'orsa_analysis_data',
    @level2type = N'COLUMN', @level2name = 'outcome_numeric';
GO

EXEC sp_addextendedproperty 
    @name = N'MS_Description', 
    @value = N'Timestamp when the check was processed', 
    @level0type = N'SCHEMA', @level0name = 'gbi',
    @level1type = N'TABLE', @level1name = 'orsa_analysis_data',
    @level2type = N'COLUMN', @level2name = 'processed_timestamp';
GO

-- Grant permissions (adjust as needed for your environment)
-- GRANT SELECT, INSERT, UPDATE ON gbi.orsa_analysis_data TO [YourUserOrRole];
-- GRANT SELECT ON gbi.vw_orsa_analysis_latest TO [YourUserOrRole];
-- GRANT SELECT ON gbi.vw_orsa_analysis_summary TO [YourUserOrRole];
-- GO

PRINT 'Table gbi.orsa_analysis_data created successfully';
PRINT 'Views gbi.vw_orsa_analysis_latest and gbi.vw_orsa_analysis_summary created successfully';
GO
