SELECT Yhteystiedot.Etunimi, Yhteystiedot.Sukunimi, Kortit.Kortti, Yhteystiedot.Jäsenyys
FROM Yhteystiedot INNER JOIN Kortit ON Yhteystiedot.[UID] = Kortit.[Omistaja]
WHERE (((Yhteystiedot.Jäsenyys)<>"Jäsen"));
