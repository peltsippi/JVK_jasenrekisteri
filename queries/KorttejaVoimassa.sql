SELECT Count([Lataukset.Kortti]) AS Kaikki, SUM(IIF([Lataukset.Korttityyppi] LIKE "*kk",1,0)) AS Kuukausikortit, SUM(IIF([Lataukset.Korttityyppi] LIKE "*ap",1,0)) AS Aamupvkortit, SUM(IIF([Lataukset.Korttityyppi] LIKE "*opisk",1,0)) AS Opiskelijakortit, SUM(IIF([Lataukset.Korttityyppi] LIKE "*krt",1,0)) AS Kertakortit, (Kaikki - Kuukausikortit - Aamupvkortit - Opiskelijakortit - Kertakortit) AS Muut
FROM Lataukset
WHERE ([Lataukset.Voimassa]>=Now());
