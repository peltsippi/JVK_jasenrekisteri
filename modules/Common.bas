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

Public Function WriteStats()
    Dim succs
    Dim arvoParit As String
    Dim targetString As String
    Dim KKKortti As Integer
    Dim opiskKortti As Integer
    Dim kertaKortti As Integer
    Dim APKortti As Integer
    Dim MUUKortti As Integer
    Dim korttejaYht As Integer
    
    KKKortti = [Form_Tervetuloa].kortitKK.Value
    opiskKortti = [Form_Tervetuloa].kortitOpisk.Value
    kertaKortti = [Form_Tervetuloa].kortitKrt.Value
    APKortti = [Form_Tervetuloa].kortitAP.Value
    MUUKortti = [Form_Tervetuloa].kortitMuu.Value
    korttejaYht = [Form_Tervetuloa].kortitKaikki.Value
    
    arvoParit = "Kaikki = " & korttejaYht & " , " _
    & "KkKortit = " & KKKortti & " , " _
    & "ApKortit = " & APKortti & " , " _
    & "KrtKortit = " & kertaKortti & " , " _
    & "OpiskKortit = " & opiskKortti & " , " _
    & "MuuKortit = " & MUUKortti & " , " _
    & "PVM = '" & Date & "'"
    
    targetString = "PVM LIKE '" & Date & "'"
    
    'MsgBox (targetString)

    Common.InsertOrUpdate "Korttitilasto", arvoParit, targetString
    
        
End Function


Public Function DoBackup(treshold As Integer)

    Dim succs
    
    Dim okString As String
    
    okString = "Varmuuskopio ok!"

    'check latest backup
    Dim timeStamp As Date
    Dim queryString As String
    queryString = "SELECT Max(Aika) AS Viimeisin FROM Historia WHERE Kirjaus Like '" & okString & "'"
    Dim sqlRecords As DAO.Recordset
    Set sqlRecords = CurrentDb.OpenRecordset(queryString)
    If (sqlRecords.RecordCount = 1) Then
        If (IsNull(sqlRecords.Fields.Item(0).Value)) Then
            timeStamp = 0
        Else
            timeStamp = sqlRecords.Fields.Item(0).Value
        End If
    Else
        timeStamp = 0
    End If
    
    Dim difference As Integer 'as days
    difference = (Year(Now()) - Year(timeStamp))
    
    If (difference < 2) Then ' calculate years first just to prevent overflow
    difference = difference * 365
    difference = difference + (Month(Now()) - Month(timeStamp)) * 30
    difference = difference + (Day(Now()) - Day(timeStamp))
    
    End If
    
    
    If (difference < treshold) Then
    
        'MsgBox ("Too early, not backing up")
        succs = Common.SendMessageToMainScreen("Varmuuskopio on jo otettu hiljattain!")
        succs = Common.SaveToLog("Varmuuskopiointi peruutettu. Varmuuskopio otettu " & difference & " päivää sitten, raja: " & treshold)
        
        Exit Function
    
    End If
    
    

'thank you http://justin-hampton.com/microsoft-office-tips/access-tips/automate-backing-database-vba/
    Dim Source As String
    Dim Target As String
    Dim retval As Integer
    
    'Source = CurrentDb.Name
    Source = CurrentDb.TableDefs("Kortit").Connect
    
    Dim splitString
    
    splitString = Split(Source, "=")
    Source = splitString(1)
    
    MsgBox (Source)
    
    'This is the only thing to change - add the path of where you want the file to save here
    Target = Application.CurrentProject.Path & "\Jasenrekisteri-backup-"
    Target = Target & Format(Date, "yyyy-mm-dd") & ".accdb"
    'MsgBox (Target)
    ' create the backup
    retval = 0
    Dim objFSO As Object
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    retval = objFSO.CopyFile(Source, Target, True)
    'MsgBox (retval)
    Set objFSO = Nothing
    
    
    
    'Opens the folder of the file you just created
    'Application.FollowHyperlink Application.CurrentProject.Path
    'this is just confusing...
    succs = Common.SaveToLog(okString)
    succs = Common.SendMessageToMainScreen("Varmuuskopio tehty!")

End Function


Public Function SendMessageToMainScreen(message As String)
    [Form_Tervetuloa].Status.Caption = Now() & " - " & message
    [Form_Tervetuloa].Requery
    [Form_Tervetuloa].Refresh
End Function

Public Function FetchCardID(cardNumber As String) As Integer
    Dim queryString As String
    Dim sqlRecords As DAO.Recordset
    queryString = "SELECT CID FROM Kortit WHERE Kortti = '" & cardNumber & "'"
    Set sqlRecords = CurrentDb.OpenRecordset(queryString)
    If (sqlRecords.RecordCount = 1) Then
        FetchCardID = sqlRecords.Fields.Item(0).Value
    Else
        FetchCardID = 0
    End If
End Function


Public Function FetchExiprationDate(card As String) As Date
    'Fetches expiration date for a card (largest date from valid until -column where card id matches
    'if it is older than current date, current date will be used
    
    Dim cardID As Integer
    cardID = Common.FetchCardID(card)
    Dim query As String
    Dim endDate As Date
    query = "SELECT Max(Voimassa) As MaxDate FROM Lataukset WHERE Kortti = " & cardID
    Dim result As DAO.Recordset
    Set result = CurrentDb.OpenRecordset(query)
    If (IsNull(result!MaxDate)) Then
        endDate = Date
    Else
    endDate = result!MaxDate
    End If
    
    'MsgBox ("end date before closing: " & endDate)
    result.Close
    'MsgBox ("end date after closing: " & endDate)
 
    If (endDate < Date) Then
        endDate = Date
    End If
    
    FetchExiprationDate = endDate
    
     
    
End Function

Public Function FetchGeneralID(Table As String, desiredID As String, criteria As String) As Integer
    Dim queryString As String
    Dim sqlRecords As DAO.Recordset
    
    queryString = "SELECT " & desiredID & " FROM " & Table & " WHERE " & criteria
    Set sqlRecords = CurrentDb.OpenRecordset(queryString)
    
    If (sqlRecords.RecordCount = 1) Then
        'MsgBox (sqlRecords.Fields.Item(0).Value)
        FetchGeneralID = sqlRecords.Fields.Item(0).Value
        
        
    Else
        FetchGeneralID = 0
    End If
    
    sqlRecords.Close
    
    
End Function

Public Function CheckIfRecordFound(Table As String, criteria As String) As Integer
    Dim queryString As String
    queryString = "SELECT * FROM " & Table & " WHERE " & criteria
    
    Dim sqlRecords As DAO.Recordset
    Set sqlRecords = CurrentDb.OpenRecordset(queryString)
    
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
    Dim queryString As String

    queryString = "INSERT INTO Historia " _
        & "(Aika, Kirjaus) " _
        & "VALUES ('" _
        & Now() & "', '" & message & "')"
    
    'MsgBox (querystring)
    Dim success As Integer
    success = Common.SQLQuery(queryString)

End Function

Public Function InsertOrUpdate(Table As String, Values As String, Target As String) As Boolean

'   Readme: use always [[ key = value, key = value, key = value ]] syntax!
'   note: spaces are extremely important!
'   This shit takes care of the rest!


    Dim toInsert As Boolean
    Dim success As Boolean
    
    toInsert = False ' prefer update
    success = True  '   success unless there is some edge cases later on
    
    Dim queryString As String
    
    'if target not defined : insert
    If (Target = Null) Or (Target = "") Then
        toInsert = True
    Else
        Dim checkforrows As Integer
        checkforrows = Common.CheckIfRecordFound(Table, Target)
        'MsgBox (checkforrows)
        
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
        
        array1 = Split(Values, ", ") ' separate each value pair as its own unit
        
        Dim first As Boolean
        first = True
        
        For Each row In array1
            'MsgBox (row)
        
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
        
        queryString = "INSERT INTO " & Table & " " & insertValues
        
    Else
        queryString = "UPDATE " & Table & " SET " & Values & " WHERE " & Target

    End If
    
    'MsgBox ("Querystring: " & queryString)
    
    'DoCmd.RunSQL (querystring)
    Dim success2 As Integer
    success2 = Common.SQLQuery(queryString)
    
    InsertOrUpdate = success
    'need to fix this later on! some kind of error checking?!?

End Function

Public Function EnableDisableButtons()
    
    're-query stats to main form...
    'MsgBox ("EnableDisableButtons called")
    Dim cardSelected As Boolean
    Dim userSelected As Boolean
    Dim adminModeOn As Boolean
    Dim adminInitialsOk As Boolean
    
    If IsNull([Form_Tervetuloa].Korttivalinta) Then
        cardSelected = False
    Else
        cardSelected = True
    End If
    
    If IsNull(Form_Tervetuloa.Yhteystietovalinta) Then
        userSelected = False
    Else
        userSelected = True
    End If
    
    If ([Form_Tervetuloa].KorjaaTietoja.Value) Then
        adminModeOn = True
    Else
        adminModeOn = False
    End If
    
    If IsNull([Form_Tervetuloa].Puumerkki) Or ([Form_Tervetuloa].Puumerkki.Value = "") Then
        adminInitialsOk = False
    Else
        adminInitialsOk = True
    End If
    
    ' then actual logic
    
    If (userSelected) Then
        Form_Tervetuloa.Lisääkortti.Enabled = True
    Else
        Form_Tervetuloa.Lisääkortti.Enabled = False
    End If
    
    
    
    If (cardSelected) Then
         Form_Tervetuloa.poistalinkitys.Enabled = True
        Form_Tervetuloa.RegisterPayment.Enabled = True
        Form_Tervetuloa.Korttilataus.Enabled = True
        Form_Tervetuloa.KorvaaRikkinainenKortti.Enabled = True
    Else
        Form_Tervetuloa.poistalinkitys.Enabled = False
        Form_Tervetuloa.RegisterPayment.Enabled = False
        Form_Tervetuloa.Korttilataus.Enabled = False
        Form_Tervetuloa.KorvaaRikkinainenKortti.Enabled = False
    
    End If
    
    If (userSelected) And (cardSelected) Then
        Form_Tervetuloa.poistalinkitys.Enabled = True
    Else
        Form_Tervetuloa.poistalinkitys.Enabled = False
    End If
    
    
    If (adminModeOn) Then
        [Form_Tervetuloa].Puumerkki.Visible = True
    Else
        [Form_Tervetuloa].Puumerkki.Value = ""
        [Form_Tervetuloa].Puumerkki.Visible = False
    End If
    
    If (adminModeOn) And (adminInitialsOk) Then
        [Form_Tervetuloa].Bulldog.Visible = False
    Else
        [Form_Tervetuloa].Bulldog.Visible = True
    End If
    
    If (adminModeOn) And (adminInitialsOk) And (cardSelected) Then
         [Form_Tervetuloa].MuokkaaLatauksia.Enabled = True
         [Form_Tervetuloa].MuokkaaMaksuja.Enabled = True
         [Form_Tervetuloa].KortinTapahtumat.Enabled = True
    Else
        [Form_Tervetuloa].MuokkaaLatauksia.Enabled = False
        [Form_Tervetuloa].MuokkaaMaksuja.Enabled = False
        [Form_Tervetuloa].KortinTapahtumat.Enabled = False
    End If
    
End Function

Public Function FillCardChargeData(months As Double, cardType As Integer)

'type 1 = regular
'type 2 = morning
'type 3 = student
'type 4 = kertako
'type 5 = other

Dim expirationDate As Date

Dim setDate As Date

If IsNull([Form_RekisteroiLataus].aloituspvm) Then
    setDate = Date
Else
    setDate = [Form_RekisteroiLataus].aloituspvm.Value
End If


Dim dateRound As Integer

Select Case cardType

    Case 1: [Form_RekisteroiLataus].KorttiTyyppi.Value = months & "kk"
    
    Case 2:
    If (months <= 1) Then
        months = 1
        [Form_RekisteroiLataus].KorttiTyyppi.Value = "1kk ap"
    ElseIf (months <= 6) Then
        months = 6
        [Form_RekisteroiLataus].KorttiTyyppi.Value = "6kk ap"
    Else
        months = 12
        [Form_RekisteroiLataus].KorttiTyyppi.Value = "12kk ap"
    End If
    
    Case 3: [Form_RekisteroiLataus].KorttiTyyppi.Value = months & "kk opisk"
    
    Case 4:
    [Form_RekisteroiLataus].KorttiTyyppi.Value = [Form_RekisteroiLataus].KERMaara & "krt"
    months = 24 'always 24 months to these cards as default
    
    Case 5:
    [Form_RekisteroiLataus].KorttiTyyppi.Value = "muu"

End Select

expirationDate = DateAdd("m", months, setDate)
'MsgBox ("Expiration date before rounding: " & expirationDate)

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
    
    lastDayOfMonth = DateSerial(Year(expirationDate), Month(expirationDate) + 1, 1) ' add 1 extra month, 1st day.
    'then just substract 1 day to get it right
    expirationDate = DateAdd("d", -1, lastDayOfMonth)
    

End If


'just add expirationDate to form and other stuff also!
[Form_RekisteroiLataus].Voimassa = expirationDate

Dim listahinta As Currency
Dim queryString As String

queryString = "SELECT Hinta FROM Hinnasto WHERE Tyyppi = '" & [Form_RekisteroiLataus].KorttiTyyppi.Value & "'"
Dim sqlRecords As DAO.Recordset
    Set sqlRecords = CurrentDb.OpenRecordset(queryString)
    listahinta = sqlRecords!Hinta
    
    sqlRecords.Close

[Form_RekisteroiLataus].Hinta.Value = listahinta


End Function

Public Function IsCardLinkedAlready(cardNumber As String) As Boolean
    If (Common.GetCardOwner(cardNumber) < 1) Then
        IsCardLinkedAlready = False
    Else
        IsCardLinkedAlready = True
    End If

End Function

Public Function GetCardOwner(cardNumber As String) As Integer
    Dim cardOwner As Integer
    Dim Table As String
    Dim wantedColumn As String
    Dim criteria As String
    
    Table = "Kortit"
    wantedColumn = "Omistaja"
    criteria = "Kortti = '" & cardNumber & "'"
    
    If (Common.CheckIfRecordFound(Table, criteria) < 1) Then
        GetCardOwner = -1
    Else
        GetCardOwner = Common.FetchGeneralID(Table, wantedColumn, criteria)
    End If

End Function

