﻿dbMemo "SQL" ="SELECT Korttityyppi, Count(Lataukset.Korttityyppi) AS Kpl, Sum(KortinArvo) AS My"
    "ynti\015\012FROM Lataukset\015\012WHERE (Lataukset.[Ajankohta]) Between Lomakkee"
    "t!Tervetuloa!RaportitAlku And DateAdd(\"d\",1,Lomakkeet!Tervetuloa!RaportitLoppu"
    ")\015\012GROUP BY Korttityyppi\015\012ORDER BY Sum(KortinArvo) DESC;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "OrderByOn" ="0"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
