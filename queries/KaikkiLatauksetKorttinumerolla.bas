Operation =1
Option =0
Begin InputTables
    Name ="Kortit"
    Name ="Lataukset"
End
Begin OutputColumns
    Expression ="Kortit.Kortti"
    Expression ="Lataukset.Ajankohta"
    Expression ="Lataukset.Lataaja"
    Expression ="Lataukset.Korttityyppi"
    Expression ="Lataukset.Voimassa"
    Expression ="Lataukset.KortinArvo"
End
Begin Joins
    LeftTable ="Kortit"
    RightTable ="Lataukset"
    Expression ="Kortit.[CID] = Lataukset.[Kortti]"
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
        dbText "Name" ="[Lataukset].[KortinArvo]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Kortit].[Kortti]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Lataukset].[Ajankohta]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Lataukset].[Lataaja]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Lataukset].[Korttityyppi]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Lataukset].[Voimassa]"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =0
    Right =1083
    Bottom =708
    Left =-1
    Top =-1
    Right =1063
    Bottom =318
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
        Name ="Lataukset"
        Name =""
    End
End
