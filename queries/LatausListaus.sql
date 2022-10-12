SELECT Lataukset.Ajankohta, Lataukset.Korttityyppi, Lataukset.Voimassa, Lataukset.Puumerkki, Lataukset.KortinArvo
FROM Kortit INNER JOIN Lataukset ON Kortit.[CID] = Lataukset.[Kortti]
WHERE (((Kortit.Kortti)=[Forms]![Tervetuloa]![Korttivalinta]));
