SELECT Kortit.Kortti
FROM Kortit INNER JOIN Lataukset ON Kortit.[CID] = Lataukset.[Kortti]
WHERE (Lataukset.Ajankohta <= [Lomakkeet]![Tervetuloa]![RaportitAlku]) AND (Lataukset.Voimassa >= [Lomakkeet]![Tervetuloa]![RaportitLoppu]);
