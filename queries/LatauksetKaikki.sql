SELECT Korttityyppi, Count(Lataukset.Korttityyppi) AS Kpl, Sum(KortinArvo) AS Myynti
FROM Lataukset
WHERE (Lataukset.[Ajankohta]) Between Lomakkeet!Tervetuloa!RaportitAlku And DateAdd("d",1,Lomakkeet!Tervetuloa!RaportitLoppu)
GROUP BY Korttityyppi
ORDER BY Sum(KortinArvo) DESC;
