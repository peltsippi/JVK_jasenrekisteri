SELECT DISTINCT Kortit.Kortti
FROM Kortit LEFT JOIN Lataukset ON Kortit.[CID] = Lataukset.[Kortti]
WHERE (((Lataukset.Voimassa)>=Now()) AND ((Kortit.Omistaja)=0));
