SELECT Lataukset.Voimassa, Kortit.Kortti, Lataukset.Lataaja, Lataukset.Korttityyppi, Lataukset.KortinArvo, Lataukset.Ajankohta
FROM Kortit INNER JOIN Lataukset ON Kortit.[CID] = Lataukset.[Kortti]
WHERE (((Kortit.Kortti)=[Lomakkeet]![Tervetuloa]![Korttivalinta]));
