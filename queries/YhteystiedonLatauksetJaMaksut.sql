SELECT Yhteystiedot.Etunimi, Yhteystiedot.Sukunimi, Lataukset.Ajankohta, Lataukset.Korttityyppi, Lataukset.KortinArvo, Maksut.PVM, Maksut.Summa
FROM (Yhteystiedot INNER JOIN (Kortit INNER JOIN Maksut ON Kortit.[CID] = Maksut.[Kortti]) ON Yhteystiedot.[UID] = Kortit.[Omistaja]) INNER JOIN Lataukset ON Kortit.[CID] = Lataukset.[Kortti];
