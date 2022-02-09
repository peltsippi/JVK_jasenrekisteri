Operation =1
Option =0
Having ="(((Kortit.Omistaja)=[Lomakkeet]![Tervetuloa]![Yhteystietovalinta] Or (Kortit.Omi"
    "staja)=IsNull([Lomakkeet]![Tervetuloa]![Yhteystietovalinta]) Or (Kortit.Omistaja"
    ")=\"isNull[Lomakkeet]![Tervetuloa]![Yhteystietovalinta]\" Or (Kortit.Omistaja)=I"
    "sNull([Lomakkeet]![Tervetuloa]![Yhteystietovalinta])))"
Begin InputTables
    Name ="Kortit"
End
Begin OutputColumns
    Expression ="Kortit.Kortti"
    Expression ="Kortit.PVM"
    Expression ="Kortit.Puumerkki"
    Expression ="Kortit.Muistiinpanot"
    Expression ="Kortit.Omistaja"
End
Begin Groups
    Expression ="Kortit.Kortti"
    GroupLevel =0
    Expression ="Kortit.PVM"
    GroupLevel =0
    Expression ="Kortit.Puumerkki"
    GroupLevel =0
    Expression ="Kortit.Muistiinpanot"
    GroupLevel =0
    Expression ="Kortit.Omistaja"
    GroupLevel =0
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
        dbText "Name" ="Kortit.PVM"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Kortit.Puumerkki"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Kortit.Muistiinpanot"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Kortit.Kortti"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Kortit.Omistaja"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =0
    Right =1583
    Bottom =708
    Left =-1
    Top =-1
    Right =1563
    Bottom =102
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =60
        Top =15
        Right =240
        Bottom =195
        Top =0
        Name ="Kortit"
        Name =""
    End
End
