Operation =1
Option =0
Begin InputTables
    Name ="Hinnasto"
End
Begin OutputColumns
    Expression ="Hinnasto.Tyyppi"
End
Begin Groups
    Expression ="Hinnasto.Tyyppi"
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
        dbText "Name" ="Hinnasto.Tyyppi"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =0
    Right =1525
    Bottom =708
    Left =-1
    Top =-1
    Right =1505
    Bottom =297
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =60
        Top =15
        Right =240
        Bottom =195
        Top =0
        Name ="Hinnasto"
        Name =""
    End
End
