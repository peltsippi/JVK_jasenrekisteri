SELECT Hinnasto.Hinta
FROM Hinnasto
WHERE (((Hinnasto.Tyyppi)=Forms!RekisteroiLataus!Korttityyppi)) Or (((Forms!RekisteroiLataus!Korttityyppi) Is Null));
