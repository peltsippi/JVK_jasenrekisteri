SELECT Count(Lataukset.Kortti) AS Kertakortit
FROM Lataukset
WHERE (Lataukset.Voimassa)>=Now() AND (Lataukset.Korttityyppi) Like "*krt";
