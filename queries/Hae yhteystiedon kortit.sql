SELECT Kortit.Kortti, Kortit.PVM, Kortit.Puumerkki, Kortit.Muistiinpanot, Kortit.Omistaja
FROM Kortit
GROUP BY Kortit.Kortti, Kortit.PVM, Kortit.Puumerkki, Kortit.Muistiinpanot, Kortit.Omistaja
HAVING (((Kortit.Omistaja)=[Lomakkeet]![Tervetuloa]![Yhteystietovalinta] Or (Kortit.Omistaja)=IsNull([Lomakkeet]![Tervetuloa]![Yhteystietovalinta]) Or (Kortit.Omistaja)="isNull[Lomakkeet]![Tervetuloa]![Yhteystietovalinta]" Or (Kortit.Omistaja)=IsNull([Lomakkeet]![Tervetuloa]![Yhteystietovalinta])));
