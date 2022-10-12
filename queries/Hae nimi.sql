SELECT Yhteystiedot.Sukunimi, Yhteystiedot.Etunimi
FROM Yhteystiedot
WHERE (((Yhteystiedot.UID)=Forms!Tervetuloa!Yhteystietovalinta)) Or (((Forms!Tervetuloa!Yhteystietovalinta) Is Null));
