SELECT Maksut.PVM, Maksut.Summa, Maksut.Puumerkki, Maksut.Maksutapa
FROM Kortit INNER JOIN Maksut ON Kortit.[CID] = Maksut.[Kortti]
WHERE (((Kortit.Kortti)=[Lomakkeet]![Tervetuloa]![Korttivalinta]));
