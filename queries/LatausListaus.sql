SELECT Lataukset.Ajankohta, Lataukset.Korttityyppi, Lataukset.Voimassa, Lataukset.Lataaja, Lataukset.KortinArvo
FROM Kortit INNER JOIN Lataukset ON Kortit.[CID] = Lataukset.[Kortti]
WHERE (((Kortit.Kortti)=[Lomakkeet]![Tervetuloa]![Korttivalinta]));
