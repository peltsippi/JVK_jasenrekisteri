SELECT DISTINCT Count(Lataukset.Kortti) AS Kuukausikortit
FROM Lataukset
WHERE (((Lataukset.Voimassa)>=Now()) AND ((Lataukset.Korttityyppi) Like "*kk"));
