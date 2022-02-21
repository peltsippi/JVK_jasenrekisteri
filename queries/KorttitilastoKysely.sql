SELECT Korttitilasto.[PVM], Korttitilasto.[Kaikki], Korttitilasto.[KkKortit], Korttitilasto.[ApKortit], Korttitilasto.[KrtKortit], Korttitilasto.[OpiskKortit], Korttitilasto.[MuuKortit]
FROM Korttitilasto
WHERE (Korttitilasto.[PVM] Between [Lomakkeet]![Tervetuloa]![RaportitAlku] And DateAdd("d",1,[Lomakkeet]![Tervetuloa]![RaportitLoppu]));
