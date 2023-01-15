SELECT k.Kortti & " - " & Tyyppi.Korttityyppi & " - " & Lataus.Voimassaolo AS Korttilistaus, k.Kortti
FROM (Kortit AS k LEFT JOIN (SELECT Kortti, Max(Voimassa) AS Voimassaolo FROM Lataukset GROUP BY Kortti)  AS Lataus ON k.[CID] = Lataus.Kortti) LEFT JOIN (SELECT Kortti, Korttityyppi, Voimassa FROM Lataukset)  AS Tyyppi ON (Lataus.[Voimassaolo] = Tyyppi.[Voimassa]) AND (Lataus.[Kortti] = Tyyppi.[Kortti])
WHERE (((k.Omistaja)=Forms!Tervetuloa!Yhteystietovalinta)) Or (((Forms!Tervetuloa!Yhteystietovalinta) Is Null))
ORDER BY k.Kortti;
