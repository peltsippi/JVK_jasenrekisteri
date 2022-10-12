﻿dbMemo "SQL" ="SELECT k.Kortti, Tyyppi.Korttityyppi, Lataus.Voimassaolo\015\012FROM (Kortit AS "
    "k LEFT JOIN (SELECT Kortti, Max(Voimassa) AS Voimassaolo FROM Lataukset GROUP BY"
    " Kortti)  AS Lataus ON k.[CID] = Lataus.Kortti) LEFT JOIN (SELECT Kortti, Kortti"
    "tyyppi, Voimassa FROM Lataukset)  AS Tyyppi ON (Lataus.[Voimassaolo] = Tyyppi.[V"
    "oimassa]) AND (Lataus.[Kortti] = Tyyppi.[Kortti])\015\012WHERE (((k.Omistaja)=Fo"
    "rms!Tervetuloa!Yhteystietovalinta)) Or (((Forms!Tervetuloa!Yhteystietovalinta) I"
    "s Null))\015\012ORDER BY k.Kortti;\015\012"
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
