Operation =1
Option =0
Where ="(((Lataukset.Ajankohta)<=[Lomakkeet]![Tervetuloa]![RaportitAlku]) AND ((Lataukse"
    "t.Voimassa)>=[Lomakkeet]![Tervetuloa]![RaportitLoppu]))"
Begin InputTables
    Name ="Kortit"
    Name ="Lataukset"
    Name ="Yhteystiedot"
End
Begin OutputColumns
    Expression ="Kortit.Kortti"
    Expression ="Yhteystiedot.Sukunimi"
    Expression ="Yhteystiedot.Etunimi"
End
Begin Joins
    LeftTable ="Kortit"
    RightTable ="Lataukset"
    Expression ="Kortit.[CID] = Lataukset.[Kortti]"
    Flag =1
    LeftTable ="Yhteystiedot"
    RightTable ="Kortit"
    Expression ="Yhteystiedot.UID = Kortit.Omistaja"
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
        dbText "Name" ="Kortti"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Lataukset_Kortti"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Kortit_Kortti"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Kortit.Kortti"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Yhteystiedot.Sukunimi"
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
    Right =1075
    Bottom =708
    Left =-1
    Top =-1
    Right =1055
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
    Begin
        Left =610
        Top =119
        Right =790
        Bottom =299
        Top =0
        Name ="Yhteystiedot"
        Name =""
    End
End
