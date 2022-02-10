Attribute VB_Name = "Common"
' See jvk_jasenrekisteri_notes.txt for tracking info
'
'   Just backup of the most important query..
'
'   SELECT k.Kortti, Tyyppi.Korttityyppi, Lataus.Voimassaolo
'   FROM (Kortit AS k LEFT JOIN (SELECT Kortti, Max(Voimassa) AS Voimassaolo FROM Lataukset GROUP BY Kortti)  AS Lataus ON k.[CID] = Lataus.Kortti) LEFT JOIN (SELECT Kortti, Korttityyppi, Voimassa FROM Lataukset)  AS Tyyppi ON Lataus.[Voimassaolo] = Tyyppi.[Voimassa] AND Lataus.[Kortti] = Tyyppi.[Kortti]
'   WHERE (([k].Omistaja) = [Lomakkeet]![Tervetuloa]![Yhteystietovalinta]) Or (([Lomakkeet]![Tervetuloa]![Yhteystietovalinta]) Is Null)
'   ORDER BY k.Kortti;f






Option Compare Database

Public Function SendMessageToMainScreen(message As String)
    [Form_Tervetuloa].Status.Caption = Now() & " - " & message
    [Form_Tervetuloa].Refresh
End Function

Public Function FetchCardID(cardnumber As String) As Integer
    Dim querystring As String
    querystring = "SELECT CID FROM Kortit WHERE Kortti = '" & cardnumber & "'"
    Set sqlRecords = CurrentDb.OpenRecordset(querystring)
    If (sqlRecords.RecordCount = 1) Then
        FetchCardID = sqlRecords.Fields.Item(0).Value
    Else
        FetchCardID = 0
    End If
End Function


Public Function FetchGeneralID(table As String, desiredID As String, criteria As String) As Integer
    Dim querystring As String
    Dim sqlRecords As DAO.Recordset
    
    querystring = "SELECT " & desiredID & " FROM " & table & " WHERE " & criteria
    Set sqlRecords = CurrentDb.OpenRecordset(querystring)
    
    If (sqlRecords.RecordCount = 1) Then
        'MsgBox (sqlRecords.Fields.Item(0).Value)
        FetchGeneralID = sqlRecords.Fields.Item(0).Value
        
        
    Else
        FetchGeneralID = 0
    End If
    
    sqlRecords.Close
    
    
End Function

Public Function CheckIfRecordFound(table As String, criteria As String) As Integer
    Dim querystring As String
    querystring = "SELECT * FROM " & table & " WHERE " & criteria
    
    Dim sqlRecords As DAO.Recordset
    Set sqlRecords = CurrentDb.OpenRecordset(querystring)
    
    CheckIfRecordFound = sqlRecords.RecordCount
    
    sqlRecords.Close
 
    End Function


Public Function SQLQuery(query As String) 'popup only when errors

    Dim dbs As DAO.Database
    
    Set dbs = CurrentDb()
    
    dbs.Execute query, dbFailOnError
    
    dbs.Close
    Set dbs = Nothing

End Function

Public Function SaveToLog(message As String)
    Dim querystring As String

    querystring = "INSERT INTO Historia " _
        & "(Aika, Kirjaus) " _
        & "VALUES ('" _
        & Now() & "', '" & message & "')"
    
    'MsgBox (querystring)
    Dim success As Integer
    success = Common.SQLQuery(querystring)

End Function

Public Function InsertOrUpdate(table As String, values As String, target As String) As Boolean

'   Readme: use always [[ key = value, key = value, key = value ]] syntax!
'   This shit takes care of the rest!


    Dim toInsert As Boolean
    Dim success As Boolean
    
    toInsert = False ' prefer update
    success = True  '   success unless there is some edge cases later on
    
    Dim querystring As String
    
    'if target not defined : insert
    If (target = Null) Or (target = "") Then
        toInsert = True
    Else
        Dim checkforrows As Integer
        checkforrows = Common.CheckIfRecordFound(table, target)
        
        If (checkforrows > 1) Or (checkforrows < 1) Then
            toInsert = True
            success = False ' i have no idea what i am doing...
            
        End If
        
    End If
      
    
    If (toInsert) Then
    
        Dim array1() As String
        Dim array2() As String
        
        Dim part1 As String
        Dim part2 As String
        
        part1 = "( "
        part2 = " VALUES ( "
        
        Dim insertValues As String
        
        array1 = Split(values, ", ") ' separate each value pair as its own unit
        
        Dim first As Boolean
        first = True
        
        For Each row In array1
        
            array2 = Split(row, " = ") 'split value pair and put it to 2 different parts
            If (first) Then
                part1 = part1 & array2(0)
                part2 = part2 & array2(1)
                first = False
                
            Else
                part1 = part1 & " , " & array2(0)
                part2 = part2 & " , " & array2(1)
            
            End If
        
        Next
        
        part1 = part1 & " ) "
        part2 = part2 & " ) "
        insertValues = part1 & part2
        
        querystring = "INSERT INTO " & table & insertValues
        
    Else
        querystring = "UPDATE " & table & " SET " & values & " WHERE " & target

    End If
    
    'DoCmd.RunSQL (querystring)
    Dim success2 As Integer
    success2 = Common.SQLQuery(querystring)
    
    InsertOrUpdate = success
    'need to fix this later on! some kind of error checking?!?

End Function

Public Function EnableDisableButtons()
    'MsgBox ("EnableDisableButtons called")
    If IsNull(Form_Tervetuloa.Korttivalinta) Then
        Form_Tervetuloa.poistalinkitys.Enabled = False
        Form_Tervetuloa.RegisterPayment.Enabled = False
        Form_Tervetuloa.Korttilataus.Enabled = False
        'Form_Tervetuloa.KorjaaTietoja.Enabled = False
        '[Form_Tervetuloa].Korttikorjaukset.Visible = False
        
        
        
    Else
        Form_Tervetuloa.poistalinkitys.Enabled = True
        Form_Tervetuloa.RegisterPayment.Enabled = True
        Form_Tervetuloa.Korttilataus.Enabled = True
        'Form_Tervetuloa.KorjaaTietoja.Enabled = True
        
    End If

    If IsNull(Form_Tervetuloa.Yhteystietovalinta) Then
        Form_Tervetuloa.Lisääkortti.Enabled = False
        Form_Tervetuloa.poistalinkitys.Enabled = False
    Else
        Form_Tervetuloa.Lisääkortti.Enabled = True
        If Not IsNull(Form_Tervetuloa.Korttivalinta) Then
        
            Form_Tervetuloa.poistalinkitys.Enabled = True
        End If
        
    
    End If
    
    If ([Form_Tervetuloa].KorjaaTietoja.Value) Then
        'MsgBox ("Jee enabled korjaa kortteja")
        [Form_Tervetuloa].Korttikorjaukset.Visible = True
        [Form_Tervetuloa].Puumerkki.Visible = True
        [Form_Tervetuloa].Hinnat.Visible = True
        
    Else
        [Form_Tervetuloa].Korttikorjaukset.Visible = False
        [Form_Tervetuloa].Puumerkki.Visible = False
        [Form_Tervetuloa].MuokkaaLatauksia.Visible = False
        [Form_Tervetuloa].MuokkaaMaksuja.Visible = False
        [Form_Tervetuloa].Hinnat.Visible = False
        [Form_Tervetuloa].Puumerkki.Value = ""
        [Form_Tervetuloa].Historia.Visible = False
        [Form_Tervetuloa].Raportit.Visible = False
        [Form_Tervetuloa].RaporttiMaksamatta.Visible = False
        [Form_Tervetuloa].KortinTapahtumat.Visible = False
        [Form_Tervetuloa].IlmaiseksiLadattavat.Visible = False
        Form_Tervetuloa.PaymentMethods.Visible = False
    
    End If
    
    If IsNull([Form_Tervetuloa].Puumerkki) Or ([Form_Tervetuloa].Puumerkki.Value = "") Then
        [Form_Tervetuloa].MuokkaaLatauksia.Visible = False
        [Form_Tervetuloa].MuokkaaMaksuja.Visible = False
        [Form_Tervetuloa].Hinnat.Visible = False
        [Form_Tervetuloa].Historia.Visible = False
        [Form_Tervetuloa].RaportitAlku.Visible = False
        [Form_Tervetuloa].RaportitLoppu.Visible = False
        [Form_Tervetuloa].LatauksetKaikki.Visible = False
        [Form_Tervetuloa].ListaaEdustusj.Visible = False
        [Form_Tervetuloa].ListaaKaikkiMaksut.Visible = False
        
    Else
        If ([Form_Tervetuloa].KorjaaTietoja.Value) Then
            If Not (IsNull([Form_Tervetuloa].Korttivalinta) Or ([Form_Tervetuloa].Korttivalinta.Value = "")) Then
                [Form_Tervetuloa].MuokkaaLatauksia.Visible = True
                [Form_Tervetuloa].MuokkaaMaksuja.Visible = True
                [Form_Tervetuloa].KortinTapahtumat.Visible = True
                
            End If
            [Form_Tervetuloa].Hinnat.Visible = True
            [Form_Tervetuloa].Historia.Visible = True
            [Form_Tervetuloa].Raportit.Visible = True
            [Form_Tervetuloa].RaporttiMaksamatta.Visible = True
            [Form_Tervetuloa].RaportitAlku.Visible = True
            [Form_Tervetuloa].RaportitLoppu.Visible = True
            [Form_Tervetuloa].LatauksetKaikki.Visible = True
            [Form_Tervetuloa].ListaaEdustusj.Visible = True
            [Form_Tervetuloa].IlmaiseksiLadattavat.Visible = True
            [Form_Tervetuloa].PaymentMethods.Visible = True
            [Form_Tervetuloa].ListaaKaikkiMaksut.Visible = True
        End If
    End If
    
    
    
End Function

Public Function FillCardChargeData(months As Double, cardType As Integer)

'type 1 = regular
'type 2 = morning
'type 3 = student
'type 4 = kertako
'type 5 = other

Dim expirationDate As Date
Dim dateRound As Integer

Select Case cardType

    Case 1: [Form_RekisteroiLataus].Korttityyppi.Value = months & "kk"
    
    Case 2:
    If (months <= 1) Then
        months = 1
        [Form_RekisteroiLataus].Korttityyppi.Value = "1kk ap"
    ElseIf (months <= 6) Then
        months = 6
        [Form_RekisteroiLataus].Korttityyppi.Value = "6kk ap"
    Else
        months = 12
        [Form_RekisteroiLataus].Korttityyppi.Value = "12kk ap"
    End If
    
    Case 3: [Form_RekisteroiLataus].Korttityyppi.Value = months & "kk opisk"
    
    Case 4:
    [Form_RekisteroiLataus].Korttityyppi.Value = [Form_RekisteroiLataus].KERMaara & "krt"
    months = 24 'always 24 months to these cards as default
    
    Case 5:
    [Form_RekisteroiLataus].Korttityyppi.Value = "muu"

End Select

expirationDate = DateAdd("m", months, Now())

'and round up
If (Day(expirationDate) < 15) Then
    Dim difference As Integer
    Dim yy As Integer
    Dim mm As Integer
    Dim dd As Integer
    
    dd = Day(expirationDate)
    mm = Month(expirationDate)
    yy = Year(expirationDate)
    
    difference = 15 - dd
    
    dd = dd + difference
    'MsgBox ("Difference for date, day less than 15: " & difference)
    
    expirationDate = DateSerial(yy, mm, dd)
    ' round to 15

Else
    ' round to month end, how?
    Dim lastDayOfMonth As Date
    
    lastDayOfMonth = DateAdd("d", -1, (DateAdd("m", 1, expirationDate)))
    'MsgBox ("Last day of month: " & lastDayOfMonth)
    
    expirationDate = lastDayOfMonth
End If


'just add expirationDate to form and other stuff also!
[Form_RekisteroiLataus].Voimassa = expirationDate

Dim listahinta As Currency
Dim querystring As String

querystring = "SELECT Hinta FROM Hinnasto WHERE Tyyppi = '" & [Form_RekisteroiLataus].Korttityyppi.Value & "'"
Dim sqlRecords As DAO.Recordset
    Set sqlRecords = CurrentDb.OpenRecordset(querystring)
    listahinta = sqlRecords!Hinta
    
    sqlRecords.Close

[Form_RekisteroiLataus].Hinta.Value = listahinta


End Function



