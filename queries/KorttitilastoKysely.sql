SELECT Korttitilasto.[PVM], Korttitilasto.[Kaikki], Korttitilasto.[KkKortit], Korttitilasto.[ApKortit], Korttitilasto.[KrtKortit], Korttitilasto.[OpiskKortit], Korttitilasto.[MuuKortit]
FROM Korttitilasto
WHERE (Korttitilasto.[PVM] Between [Forms]![Tervetuloa]![RaportitAlku] And DateAdd("d",1,[Forms]![Tervetuloa]![RaportitLoppu]));
