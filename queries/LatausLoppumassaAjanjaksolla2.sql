SELECT LAST(Voimassa) AS Voimassaolo, LAST(Korttityyppi) AS Kortintyyppi, LAST(Kortti) AS KorttiID
FROM (SELECT Voimassa, Korttityyppi, Kortti FROM Lataukset GROUP BY Kortti, Voimassa, Korttityyppi ORDER BY Kortti, Voimassa DESC)  AS [%$##@_Alias]
GROUP BY Kortti;
