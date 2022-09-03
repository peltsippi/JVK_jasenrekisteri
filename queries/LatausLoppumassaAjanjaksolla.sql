SELECT Kortit.Kortti, Lataukset.Korttityyppi, Yhteystiedot.Sukunimi, Yhteystiedot.Etunimi
FROM Yhteystiedot INNER JOIN (Kortit INNER JOIN Lataukset ON Kortit.[CID] = Lataukset.[Kortti]) ON Yhteystiedot.UID = Kortit.Omistaja
WHERE (((Lataukset.Voimassa)>=[Lomakkeet]![Tervetuloa]![RaportitAlku]) AND ((Lataukset.Voimassa)<=[Lomakkeet]![Tervetuloa]![RaportitLoppu]));
