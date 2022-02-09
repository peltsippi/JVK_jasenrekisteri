SELECT Hinnasto.Hinta
FROM Hinnasto
WHERE (((Hinnasto.Tyyppi)=[Lomakkeet]![RekisteroiLataus]![Korttityyppi])) OR ((([Lomakkeet]![RekisteroiLataus]![Korttityyppi]) Is Null));
