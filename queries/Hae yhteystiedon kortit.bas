﻿dbMemo "SQL" ="SELECT Kortit.Kortti, Kortit.PVM, Kortit.Puumerkki, Kortit.Muistiinpanot, Kortit"
    ".Omistaja\015\012FROM Kortit\015\012GROUP BY Kortit.Kortti, Kortit.PVM, Kortit.P"
    "uumerkki, Kortit.Muistiinpanot, Kortit.Omistaja\015\012HAVING ((Kortit.Omistaja)"
    "=[Forms]![Tervetuloa]![Yhteystietovalinta] Or (Kortit.Omistaja)=IsNull([Forms]!["
    "Tervetuloa]![Yhteystietovalinta]));\015\012"
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
