﻿dbMemo "SQL" ="SELECT DISTINCT Count(Lataukset.Kortti) AS Kuukausikortit\015\012FROM Lataukset\015"
    "\012WHERE (((Lataukset.Voimassa)>=Now()) AND ((Lataukset.Korttityyppi) Like \"*k"
    "k\"));\015\012"
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