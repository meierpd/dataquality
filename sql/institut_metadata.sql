
    SELECT 
    MitarbeiterNummer,
    MitarbeiterKuerzel,
    MitarbeiterName,
    MitarbeiterOrgEinheit,
    FinmaObjektNr AS FINMAID,
    FinmaObjektName,
    ZulassungName,
    SachbearbeiterTypName
    FROM DWHMart.dbo.Sachbearbeiter s
    where SachbearbeiterTypID = 'D684317D-98A3-4A31-B353-84D64E5D8A4C'
    and StatusCode = 1