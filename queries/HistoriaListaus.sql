SELECT Historia.Aika, Historia.Kirjaus
FROM Historia
WHERE [Aika] Between Lomakkeet!Tervetuloa!RaportitAlku And DateAdd("d",1,Lomakkeet!Tervetuloa!RaportitLoppu);
