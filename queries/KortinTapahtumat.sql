SELECT k.Kortti AS Card, Ajankohta AS PVM, (-1 * KortinArvo) AS Balance
FROM Lataukset
INNER JOIN Kortit AS k ON k.[CID] = Lataukset.Kortti

WHERE k.Kortti = Lomakkeet!Tervetuloa!Korttivalinta

UNION SELECT kk.Kortti AS Card, Maksut.PVM, Summa As Balance
FROM Maksut
INNER JOIN Kortit As kk ON kk.[CID] = Maksut.Kortti 
WHERE kk.Kortti = Lomakkeet!Tervetuloa!Korttivalinta
ORDER BY PVM;
