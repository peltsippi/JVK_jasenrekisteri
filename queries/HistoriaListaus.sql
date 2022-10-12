SELECT Historia.Aika, Historia.Kirjaus
FROM Historia
WHERE (((Historia.[Aika]) Between [Forms]![Tervetuloa]![RaportitAlku] And DateAdd("d",1,[Forms]![Tervetuloa]![RaportitLoppu])));
