dbMemo "SQL" ="SELECT Korttitilasto.[PVM], Korttitilasto.[Kaikki], Korttitilasto.[KkKortit], Ko"
    "rttitilasto.[ApKortit], Korttitilasto.[KrtKortit], Korttitilasto.[OpiskKortit], "
    "Korttitilasto.[MuuKortit]\015\012FROM Korttitilasto\015\012WHERE (Korttitilasto."
    "[PVM] Between [Lomakkeet]![Tervetuloa]![RaportitAlku] And DateAdd(\"d\",1,[Lomak"
    "keet]![Tervetuloa]![RaportitLoppu]));\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
