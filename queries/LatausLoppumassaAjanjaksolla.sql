SELECT MAX(Lataukset.Voimassa) AS Voimassaolo, Kortit.Kortti, Yhteystiedot.Sukunimi & ", " & Yhteystiedot.Etunimi AS Nimi
FROM (Lataukset INNER JOIN Kortit ON Kortit.CID = Lataukset.Kortti) INNER JOIN Yhteystiedot ON Yhteystiedot.UID = Kortit.Omistaja
GROUP BY Kortit.Kortti, Yhteystiedot.Sukunimi, Yhteystiedot.Etunimi
HAVING MAX(Lataukset.Voimassa)>=Forms!Tervetuloa!RaportitAlku
AND
MAX(Lataukset.Voimassa)<=Forms!Tervetuloa!RaportitLoppu;
