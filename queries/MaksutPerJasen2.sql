SELECT Maksut.Summa, Maksut.PVM, Maksut.Maksutapa, Yhteystiedot.Sukunimi + ", " + Yhteystiedot.Etunimi AS Nimi
FROM (Yhteystiedot INNER JOIN Kortit ON Yhteystiedot.[UID] = Kortit.[Omistaja]) INNER JOIN Maksut ON Kortit.[CID] = Maksut.[Kortti]
WHERE (((Maksut.PVM) Between [Forms]![Tervetuloa]![RaportitAlku] And [Forms]![Tervetuloa]![RaportitLoppu]));
