Operation =1
Option =0
Where ="(((Yhteystiedot.Edustusjasen)=True))"
Begin InputTables
    Name ="Yhteystiedot"
End
Begin OutputColumns
    Expression ="Yhteystiedot.Etunimi"
    Expression ="Yhteystiedot.Sukunimi"
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
        dbText "Name" ="Yhteystiedot.[Etunimi]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Yhteystiedot.[Sukunimi]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Etunimi]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Sukunimi]"
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
    Bottom =709
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
End
