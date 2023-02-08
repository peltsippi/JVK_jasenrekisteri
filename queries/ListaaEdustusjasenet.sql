SELECT Yhteystiedot.Etunimi, Yhteystiedot.Sukunimi, Kortit.Kortti, Lataukset.Ajankohta, Lataukset.Korttityyppi
FROM (Yhteystiedot INNER JOIN Kortit ON Kortit.Omistaja = Yhteystiedot.[UID]) LEFT JOIN Lataukset ON Kortit.[CID] = Lataukset.Kortti
WHERE (((Yhteystiedot.Edustusjasen)=True));
