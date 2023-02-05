SELECT k.Kortti & " - " & Tyyppi.Korttityyppi & " - " & Tyyppi.Voimassa AS Korttilistaus, k.Kortti
FROM (Kortit AS k LEFT JOIN (SELECT Kortti, Max(Lataukset.Ajankohta) AS LatausPV FROM Lataukset GROUP BY Kortti)  AS Lataus ON k.[CID] = Lataus.Kortti) LEFT JOIN (SELECT Kortti, Korttityyppi, Ajankohta, Voimassa FROM Lataukset)  AS Tyyppi ON (Lataus.[Kortti] = Tyyppi.[Kortti]) AND (Lataus.[LatausPV] = Tyyppi.[Ajankohta])
WHERE (((k.Omistaja)=Forms!Tervetuloa!Yhteystietovalinta)) Or (((Forms!Tervetuloa!Yhteystietovalinta) Is Null))
ORDER BY k.Kortti;
