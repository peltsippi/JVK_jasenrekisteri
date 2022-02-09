SELECT Kortit.Kortti, Maksut.PVM, Maksut.Summa, Maksut.Maksutapa, Maksut.Puumerkki
FROM Kortit INNER JOIN Maksut ON Kortit.[CID] = Maksut.[Kortti];
