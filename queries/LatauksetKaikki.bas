﻿dbMemo "SQL" ="SELECT Lataukset.Korttityyppi, Count(Lataukset.Korttityyppi) AS Kpl, Sum(Latauks"
    "et.KortinArvo) AS Myynti\015\012FROM Lataukset\015\012WHERE (((Lataukset.Ajankoh"
    "ta) Between [Forms]![Tervetuloa]![RaportitAlku] And DateAdd(\"d\",1,[Forms]![Ter"
    "vetuloa]![RaportitLoppu])))\015\012GROUP BY Lataukset.Korttityyppi\015\012ORDER "
    "BY Sum(Lataukset.KortinArvo) DESC;\015\012"
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
