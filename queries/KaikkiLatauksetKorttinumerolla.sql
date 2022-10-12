SELECT Kortit.Kortti, Lataukset.Ajankohta, Lataukset.Lataaja AS Lauseke1, Lataukset.Korttityyppi, Lataukset.Voimassa, Lataukset.KortinArvo
FROM Kortit INNER JOIN Lataukset ON Kortit.[CID] = Lataukset.[Kortti];
