﻿dbMemo "SQL" ="SELECT Count(Lataukset.Kortti) AS Kertakortit\015\012FROM Lataukset\015\012WHERE"
    " (Lataukset.Voimassa)>=Now() AND (Lataukset.Korttityyppi) Like \"*krt\";\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
