﻿dbMemo "SQL" ="SELECT Kortit.Kortti, Yhteystiedot.Sukunimi, Yhteystiedot.Etunimi\015\012FROM Yh"
    "teystiedot INNER JOIN (Kortit INNER JOIN Lataukset ON Kortit.[CID] = Lataukset.["
    "Kortti]) ON Yhteystiedot.UID = Kortit.Omistaja\015\012WHERE (((Lataukset.Ajankoh"
    "ta)<=Forms!Tervetuloa!RaportitAlku) And ((Lataukset.Voimassa)>=Forms!Tervetuloa!"
    "RaportitLoppu));\015\012"
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
