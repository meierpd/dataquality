-- Query to retrieve ORSA document metadata from the database
-- This is a placeholder query - replace with actual SQL based on your database schema

SELECT 
    DokumentID,
    DokumentName,
    DokumentLink,
    ErstellungsDatum,
    BerichtPeriode,
    InstitutID,
    InstitutName,
    GeschaeftNr
FROM 
    dbo.ORSA_Dokumente
WHERE 
    Aktiv = 1
    AND DokumentTyp = 'ORSA-Formular'
ORDER BY 
    BerichtPeriode DESC,
    InstitutName ASC;
