SELECT Maksut.PVM, Maksut.Maksutapa, Maksut.Summa, Yhteystiedot.Etunimi, Yhteystiedot.Sukunimi
FROM (Yhteystiedot INNER JOIN Kortit ON Yhteystiedot.[UID] = Kortit.[Omistaja]) INNER JOIN Maksut ON Kortit.[CID] = Maksut.[Kortti]
WHERE (((Maksut.PVM) Between [Forms]![Tervetuloa]![RaportitAlku] And [Forms]![Tervetuloa]![RaportitLoppu]))
ORDER BY Maksut.PVM;
