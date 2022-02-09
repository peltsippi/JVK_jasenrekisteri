SELECT DISTINCTROW k.Kortti, m.Maksettu, l.Ladattu, (l.Ladattu - m.Maksettu) AS Maksamatta
FROM (Kortit AS k LEFT JOIN (SELECT Kortti, Sum(Summa) AS Maksettu FROM Maksut GROUP BY Kortti)  AS m ON k.[CID] = m.Kortti) LEFT JOIN (SELECT Kortti, SUM(KortinArvo) AS Ladattu FROM Lataukset GROUP BY Kortti)  AS l ON k.[CID] = l.Kortti;
