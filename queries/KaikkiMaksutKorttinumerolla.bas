﻿Operation =1
Option =0
Begin InputTables
    Name ="Kortit"
    Name ="Maksut"
End
Begin OutputColumns
    Expression ="Kortit.Kortti"
    Expression ="Maksut.PVM"
    Expression ="Maksut.Summa"
    Expression ="Maksut.Maksutapa"
    Expression ="Maksut.Puumerkki"
End
Begin Joins
    LeftTable ="Kortit"
    RightTable ="Maksut"
    Expression ="Kortit.[CID] = Maksut.[Kortti]"
    Flag =1
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="[Kortit].[Kortti]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Maksut].[PVM]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Maksut].[Summa]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Maksut].[Maksutapa]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Maksut].[Puumerkki]"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =0
    Right =1583
    Bottom =709
    Left =-1
    Top =-1
    Right =1563
    Bottom =339
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =60
        Top =15
        Right =240
        Bottom =195
        Top =0
        Name ="Kortit"
        Name =""
    End
    Begin
        Left =300
        Top =15
        Right =480
        Bottom =195
        Top =0
        Name ="Maksut"
        Name =""
    End
End
