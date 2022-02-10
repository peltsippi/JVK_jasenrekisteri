Operation =1
Option =0
Where ="(((Maksut.PVM) Between [Forms]![Tervetuloa]![RaportitAlku] And [Forms]![Tervetul"
    "oa]![RaportitLoppu]))"
Begin InputTables
    Name ="Yhteystiedot"
    Name ="Kortit"
    Name ="Maksut"
End
Begin OutputColumns
    Expression ="Maksut.PVM"
    Expression ="Maksut.Maksutapa"
    Expression ="Maksut.Summa"
    Expression ="Yhteystiedot.Etunimi"
    Expression ="Yhteystiedot.Sukunimi"
End
Begin Joins
    LeftTable ="Yhteystiedot"
    RightTable ="Kortit"
    Expression ="Yhteystiedot.[UID] = Kortit.[Omistaja]"
    Flag =1
    LeftTable ="Kortit"
    RightTable ="Maksut"
    Expression ="Kortit.[CID] = Maksut.[Kortti]"
    Flag =1
End
Begin OrderBy
    Expression ="Maksut.PVM"
    Flag =0
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
        dbText "Name" ="[Maksut].[Summa]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Maksut].[PVM]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Maksut].[Maksutapa]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Yhteystiedot].[Etunimi]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Yhteystiedot].[Sukunimi]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Yhteystiedot.Sukunimi"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Maksut.Summa"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Maksut.PVM"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Maksut.Maksutapa"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Yhteystiedot.Etunimi"
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
    Right =1147
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
        Name ="Yhteystiedot"
        Name =""
    End
    Begin
        Left =300
        Top =15
        Right =480
        Bottom =195
        Top =0
        Name ="Kortit"
        Name =""
    End
    Begin
        Left =540
        Top =15
        Right =720
        Bottom =195
        Top =0
        Name ="Maksut"
        Name =""
    End
End
