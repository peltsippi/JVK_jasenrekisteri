SELECT Yhteystiedot.Etunimi, Yhteystiedot.Sukunimi, Kortit.Kortti, Lataukset.Ajankohta, Lataukset.Korttityyppi
FROM (Yhteystiedot INNER JOIN Kortit ON Kortit.Omistaja = Yhteystiedot.[UID]) LEFT JOIN Lataukset ON Kortit.[CID] = Lataukset.Kortti And Lataukset.Ajankohta>=Forms!Tervetuloa!RaportitAlku And Lataukset.Ajankohta<=Forms!Tervetuloa!RaportitLoppu
WHERE ((Yhteystiedot.Edustusjasen)=True);
