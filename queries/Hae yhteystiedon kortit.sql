SELECT Kortit.Kortti, Kortit.PVM, Kortit.Puumerkki, Kortit.Muistiinpanot, Kortit.Omistaja
FROM Kortit
GROUP BY Kortit.Kortti, Kortit.PVM, Kortit.Puumerkki, Kortit.Muistiinpanot, Kortit.Omistaja
HAVING ((Kortit.Omistaja)=[Forms]![Tervetuloa]![Yhteystietovalinta] Or (Kortit.Omistaja)=IsNull([Forms]![Tervetuloa]![Yhteystietovalinta]));
