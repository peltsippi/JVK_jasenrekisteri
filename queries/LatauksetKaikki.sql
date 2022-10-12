SELECT Lataukset.Korttityyppi, Count(Lataukset.Korttityyppi) AS Kpl, Sum(Lataukset.KortinArvo) AS Myynti
FROM Lataukset
WHERE (((Lataukset.Ajankohta) Between [Forms]![Tervetuloa]![RaportitAlku] And DateAdd("d",1,[Forms]![Tervetuloa]![RaportitLoppu])))
GROUP BY Lataukset.Korttityyppi
ORDER BY Sum(Lataukset.KortinArvo) DESC;
