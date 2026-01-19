
    SELECT 
    s.MitarbeiterNummer,
    s.MitarbeiterKuerzel,
    s.MitarbeiterName,
    s.MitarbeiterOrgEinheit,
    s.FinmaObjektNr AS FINMAID,
    s.FinmaObjektName,
    s.ZulassungName,
    s.SachbearbeiterTypName,
    stamm.Aufsichtskategorie
    FROM DWHMart.dbo.Sachbearbeiter s
    left join [GBV_Reporting].[gbv].[RV_GBV_Stammdaten] stamm
    on s.FinmaObjektId = stamm.FinmaObjID
    where s.SachbearbeiterTypID = 'D684317D-98A3-4A31-B353-84D64E5D8A4C'
    and s.StatusCode = 1

