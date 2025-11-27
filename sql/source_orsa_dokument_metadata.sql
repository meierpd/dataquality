-- Query to retrieve ORSA document metadata from the database
-- This query filters for relevant ORSA documents for a specific reporting year (Berichtsjahr)
-- The {berichtsjahr} placeholder will be replaced with the actual year value

SELECT 
    DokumentID,
    DokumentName,
    DokumentLink,
    ErstellungsDatum,
    BerichtPeriode,
    InstitutID,
    InstitutName,
    GeschaeftNr,
    Berichtsjahr
FROM 
    dbo.ORSA_Dokumente
WHERE 
    Aktiv = 1
    AND DokumentTyp = 'ORSA-Formular'
    AND Berichtsjahr = {berichtsjahr}
    AND DokumentName LIKE '%_ORSA-Formular%'
ORDER BY 
    BerichtPeriode DESC,
    InstitutName ASC;
