﻿dbMemo "SQL" ="SELECT k.Kortti & \" - \" & Tyyppi.Korttityyppi & \" - \" & Tyyppi.Voimassa AS K"
    "orttilistaus, k.Kortti\015\012FROM (Kortit AS k LEFT JOIN (SELECT Kortti, Max(La"
    "taukset.Ajankohta) AS LatausPV FROM Lataukset GROUP BY Kortti)  AS Lataus ON k.["
    "CID] = Lataus.Kortti) LEFT JOIN (SELECT Kortti, Korttityyppi, Ajankohta, Voimass"
    "a FROM Lataukset)  AS Tyyppi ON (Lataus.[Kortti] = Tyyppi.[Kortti]) AND (Lataus."
    "[LatausPV] = Tyyppi.[Ajankohta])\015\012WHERE (((k.Omistaja)=Forms!Tervetuloa!Yh"
    "teystietovalinta)) Or (((Forms!Tervetuloa!Yhteystietovalinta) Is Null))\015\012O"
    "RDER BY k.Kortti;\015\012"
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
