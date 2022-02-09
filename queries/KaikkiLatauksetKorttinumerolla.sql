SELECT Kortit.Kortti, Lataukset.Ajankohta, Lataukset.Lataaja, Lataukset.Korttityyppi, Lataukset.Voimassa, Lataukset.KortinArvo
FROM Kortit INNER JOIN Lataukset ON Kortit.[CID] = Lataukset.[Kortti];
