SELECT Yhteystiedot.Sukunimi, Yhteystiedot.Etunimi, Yhteystiedot.UID
FROM Yhteystiedot
WHERE (((Yhteystiedot.UID)=[Lomakkeet]![Tervetuloa]![Yhteystietovalinta] Or [Lomakkeet]![Tervetuloa]![Yhteystietovalinta] Is Null));
