SELECT Lataukset.Voimassa, Kortit.Kortti, Lataukset.[Puumerkki], Lataukset.Korttityyppi, Lataukset.KortinArvo, Lataukset.Ajankohta
FROM Kortit INNER JOIN Lataukset ON Kortit.[CID] = Lataukset.[Kortti]
WHERE (((Kortit.Kortti)=[Lomakkeet]![Tervetuloa]![Korttivalinta]));
