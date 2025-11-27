-- Add GeschaeftNr column to orsa_analysis_data table
-- Database: GBI_REPORTING
-- Schema: gbi
-- Table: orsa_analysis_data

USE GBI_REPORTING;
GO

-- Add geschaeft_nr column to the table
IF NOT EXISTS (
    SELECT * FROM sys.columns 
    WHERE object_id = OBJECT_ID('gbi.orsa_analysis_data') 
    AND name = 'geschaeft_nr'
)
BEGIN
    ALTER TABLE gbi.orsa_analysis_data
    ADD geschaeft_nr NVARCHAR(50) NULL;
    
    PRINT 'Column geschaeft_nr added to gbi.orsa_analysis_data';
END
ELSE
BEGIN
    PRINT 'Column geschaeft_nr already exists in gbi.orsa_analysis_data';
END
GO

-- Add index on geschaeft_nr for better query performance
IF NOT EXISTS (
    SELECT * FROM sys.indexes 
    WHERE name = 'idx_geschaeft_nr' 
    AND object_id = OBJECT_ID('gbi.orsa_analysis_data')
)
BEGIN
    CREATE INDEX idx_geschaeft_nr ON gbi.orsa_analysis_data(geschaeft_nr);
    PRINT 'Index idx_geschaeft_nr created on gbi.orsa_analysis_data';
END
ELSE
BEGIN
    PRINT 'Index idx_geschaeft_nr already exists on gbi.orsa_analysis_data';
END
GO

-- Add extended property (documentation) for the new column
IF NOT EXISTS (
    SELECT * FROM sys.extended_properties
    WHERE major_id = OBJECT_ID('gbi.orsa_analysis_data')
    AND minor_id = (SELECT column_id FROM sys.columns 
                    WHERE object_id = OBJECT_ID('gbi.orsa_analysis_data') 
                    AND name = 'geschaeft_nr')
    AND name = 'MS_Description'
)
BEGIN
    EXEC sp_addextendedproperty 
        @name = N'MS_Description', 
        @value = N'Business case number (Geschäftsnummer) from source document', 
        @level0type = N'SCHEMA', @level0name = 'gbi',
        @level1type = N'TABLE', @level1name = 'orsa_analysis_data',
        @level2type = N'COLUMN', @level2name = 'geschaeft_nr';
    
    PRINT 'Extended property added for geschaeft_nr column';
END
ELSE
BEGIN
    PRINT 'Extended property already exists for geschaeft_nr column';
END
GO

PRINT 'GeschaeftNr column setup completed successfully';
GO
