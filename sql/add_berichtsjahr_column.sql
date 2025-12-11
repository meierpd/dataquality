-- Migration script to add berichtsjahr column to orsa_analysis_data table
-- Database: GBI_REPORTING
-- Schema: gbi
-- Table: orsa_analysis_data
--
-- This script adds the berichtsjahr (reporting year) column to track which
-- reporting year each document belongs to. This is a critical piece of metadata
-- that was previously not stored in the database.

USE GBI_REPORTING;
GO

-- Add berichtsjahr column to the existing table
-- Using INT data type as berichtsjahr is a year (e.g., 2026, 2027)
-- Making it nullable initially to allow for existing data
ALTER TABLE gbi.orsa_analysis_data
ADD berichtsjahr INT NULL;
GO

-- Create an index on berichtsjahr for efficient filtering
CREATE INDEX idx_berichtsjahr ON gbi.orsa_analysis_data(berichtsjahr);
GO

-- Add extended property (documentation) for the new column
EXEC sp_addextendedproperty 
    @name = N'MS_Description', 
    @value = N'Reporting year (Berichtsjahr) for the ORSA document', 
    @level0type = N'SCHEMA', @level0name = 'gbi',
    @level1type = N'TABLE', @level1name = 'orsa_analysis_data',
    @level2type = N'COLUMN', @level2name = 'berichtsjahr';
GO

PRINT 'Column berichtsjahr added successfully to gbi.orsa_analysis_data';
PRINT 'Index idx_berichtsjahr created successfully';
GO
