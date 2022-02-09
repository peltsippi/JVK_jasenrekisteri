SELECT Yhteystiedot.Sukunimi, Yhteystiedot.Etunimi
FROM Yhteystiedot
WHERE (((Yhteystiedot.UID)=[Lomakkeet]![Tervetuloa]![Yhteystietovalinta])) OR ((([Lomakkeet]![Tervetuloa]![Yhteystietovalinta]) Is Null));
