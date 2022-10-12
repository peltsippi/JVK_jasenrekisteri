SELECT Yhteystiedot.Sukunimi, Yhteystiedot.Etunimi, Yhteystiedot.UID
FROM Yhteystiedot
WHERE (((Yhteystiedot.UID)=Forms!Tervetuloa!Yhteystietovalinta)) Or (((Forms!Tervetuloa!Yhteystietovalinta) Is Null));
