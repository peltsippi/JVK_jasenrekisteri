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
    
    If (Not (IsNull([Form_Tervetuloa].kortitKK))) Then
    
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
    
    End If
        
End Function


Public Function DoBackup(treshold As Integer)

    'todo skip that treshold part completely...

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
    
    
    If (difference <= treshold) Then
    
        'MsgBox ("Too early, not backing up")
        succs = Common.SendMessageToMainScreen("Varmuuskopio on jo otettu hiljattain!")
        succs = Common.SaveToLog("Varmuuskopiointi peruutettu. Varmuuskopio otettu " & difference & " päivää sitten, raja: " & treshold)
        
        Exit Function
    
    End If
    
    

'thank you http://justin-hampton.com/microsoft-office-tips/access-tips/automate-backing-database-vba/
    Dim source As String
    Dim Target As String
    Dim retval As Integer
    
    'Source = CurrentDb.Name
    source = CurrentDb.TableDefs("Kortit").Connect
    
    Dim splitString
    
    splitString = Split(source, "=")
    source = splitString(1)
    
    'MsgBox (Source)
    
    'This is the only thing to change - add the path of where you want the file to save here
    Target = Application.CurrentProject.path & "\Jasenrekisteri-backup-"
    Target = Target & Format(Date, "yyyy-mm-dd") & ".accdb"
    'MsgBox (Target)
    ' create the backup
    retval = 0
    Dim objFSO As Object
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    retval = objFSO.CopyFile(source, Target, True)
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
    
    result.Close
    
    FetchExiprationDate = endDate
        
End Function

Public Function FetchGeneralID(table As String, desiredID As String, criteria As String) As Integer
    Dim queryString As String
    Dim sqlRecords As DAO.Recordset
    
    queryString = "SELECT " & desiredID & " FROM " & table & " WHERE " & criteria
    'MsgBox (queryString)
    Set sqlRecords = CurrentDb.OpenRecordset(queryString)
    
    If (sqlRecords.RecordCount = 1) Then
        'MsgBox (sqlRecords.Fields.Item(0).Value)
        FetchGeneralID = sqlRecords.Fields.Item(0).Value
        
        
    Else
        FetchGeneralID = 0
    End If
    
    sqlRecords.Close
    
    
End Function

Public Function CheckIfRecordFound(table As String, criteria As String) As Integer
    Dim queryString As String
    queryString = "SELECT * FROM " & table & " WHERE " & criteria
    
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

Public Function InsertOrUpdate(table As String, values As String, Target As String) As Boolean

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
        checkforrows = Common.CheckIfRecordFound(table, Target)
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
        
        array1 = Split(values, ", ") ' separate each value pair as its own unit
        
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
        
        queryString = "INSERT INTO " & table & " " & insertValues
        
    Else
        queryString = "UPDATE " & table & " SET " & values & " WHERE " & Target

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

Public Function CalculateEndingDate(months As Integer, startDate As Date)
    Dim expirationDate As Date
    expirationDate = DateAdd("m", months, startDate)
    
    'and round up
    If (Day(expirationDate) < 15) Then
        Dim difference As Integer
        Dim yy As Integer
        Dim mm As Integer
        Dim dd As Integer
    
        dd = 15
        mm = Month(expirationDate)
        yy = Year(expirationDate)
        expirationDate = DateSerial(yy, mm, dd)
    Else
        Dim lastDayOfMonth As Date
        lastDayOfMonth = DateSerial(Year(expirationDate), Month(expirationDate) + 1, 1) ' add 1 extra month, 1st day.
        expirationDate = DateAdd("d", -1, lastDayOfMonth) 'and remove 1 day to get to last day of actual month
    End If

    CalculateEndingDate = expirationDate

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
    '[Form_RekisteroiLataus].KorttiTyyppi.Value = [Form_RekisteroiLataus].KERMaara & "krt"
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
    'listahinta = sqlRecords!Hinta
    If (IsNull(sqlRecords!Hinta)) Then
        listahinta = 999
    Else
        'listahinta = sqlRecords!Hinta
    End If
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
    Dim table As String
    Dim wantedColumn As String
    Dim criteria As String
    
    table = "Kortit"
    wantedColumn = "Omistaja"
    criteria = "Kortti = '" & cardNumber & "'"
    
    If (Common.CheckIfRecordFound(table, criteria) < 1) Then
        GetCardOwner = -1
    Else
        GetCardOwner = Common.FetchGeneralID(table, wantedColumn, criteria)
    End If

End Function

Public Function GetCardType(cardNumber As String) As String
    Dim queryString As String
    Dim dateStamp As Date
    
    dateStamp = Common.FetchExiprationDate(cardNumber)
    
    cardID = Common.FetchCardID(cardNumber)
    
    
    Dim queryString2 As String
    Dim sqlRecords2 As DAO.Recordset
   '
    queryString2 = "SELECT Korttityyppi FROM Lataukset WHERE Kortti = " & cardID & " AND Voimassa = CDATE('" & dateStamp & "')"
    

    Set sqlRecords2 = CurrentDb.OpenRecordset(queryString2)
   
    'Dim i As Integer
   
    If (IsNull(sqlRecords2.Fields.Item(0).Value)) Then 'catch potential null stuff here
        GetCardType = ""
    Else
        GetCardType = sqlRecords2.Fields.Item(0).Value
    End If
    sqlRecords2.Close
        

End Function

Public Function GetPriceForCard(cardType As String, cardTime As Integer)
    
    Dim price As String
    
    Dim queryString3 As String
    Dim sqlRecords3 As DAO.Recordset
    
    'queryString3 = "SELECT Hinta FROM Hinnasto WHERE Aika='" & cardTime & "' AND Tyyppi='" & cardType & ""
    queryString3 = "SELECT Hinta FROM Hinnasto WHERE Aika=" & cardTime & " AND Tyyppi='" & cardType & "'"
    
    
    Set sqlRecords3 = CurrentDb.OpenRecordset(queryString3)
    
    'MsgBox (sqlRecords3.Fields.Count)
    'MsgBox (sqlRecords3.Fields.Item(0).Value)
    
    If (IsNull(sqlRecords3.Fields.Item(0).Value)) Then
        price = "999,99 €"
    Else
        price = sqlRecords3.Fields.Item(0).Value & " €"
    End If
    
    GetPriceForCard = price
        

End Function

Public Function Reconnect()
'**************************************************************
'*     START YOUR APPLICATION (MACRO: AUTOEXEC) WITH THIS FUNCTION
'*     AND THIS PROGRAM WILL CHANGE THE CONNECTIONS AUTOMATICALLY
'*     WHEN THE 'DATA.MDB'  AND THE 'PRG.MDB'
'*     ARE IN THE SAME DIRECTORY!!!
'*                  PROGRAMMING BY PETER VUKOVIC, Germany
'*                  100700.1262@compuserve.com
'* ************************************************************
Dim db As Database, source As String, path As String
Dim dbsource As String, i As Integer, j As Integer

Set db = DBEngine.Workspaces(0).Databases(0)
'*************************************************************
'*                     RECOGNIZE THE PATH                    *
'*************************************************************

For i = Len(db.Name) To 1 Step -1
    If Mid(db.Name, i, 1) = Chr(92) Then
        path = Mid(db.Name, 1, i)
        'MsgBox (path)
        Exit For
    End If
Next
'*************************************************************
'*              CHANGE THE PATH   AND   CONNECT  AGAIN       *
'*************************************************************

For i = 0 To db.TableDefs.Count - 1
    If db.TableDefs(i).Connect <> " " Then
        source = Mid(db.TableDefs(i).Connect, 11)
        'Debug.Print source
        For j = Len(source) To 1 Step -1
            If Mid(source, j, 1) = Chr(92) Then
               dbsource = Mid(source, j + 1, Len(source))
               source = Mid(source, 1, j)
                   If source <> path Then
                        db.TableDefs(i).Connect = ";Database=" + path + dbsource
                        db.TableDefs(i).RefreshLink
                        'Debug.Print ";Database=" + path + dbsource
                    End If
                Exit For
            End If
         Next
    End If
Next
End Function

Public Function ReplaceDatabaseFile() As Integer

Const msoFileDialogFilePicker As Long = 3
    Dim FD As Object
    Dim File As Variant
    
    'Dim varFile As Variant
    'Me.FileList.RowSource = ""
    'Set fDialog = Application.FileDialog(msoFileDialogFilePicker)
    Set FD = Application.FileDialog(msoFileDialogFilePicker)
    With FD
 
      ' Allow user to make multiple selections in dialog box
      .AllowMultiSelect = False
             
      ' Set the title of the dialog box.
      .Title = "Valitse tietokantatiedosto"
 
      ' Clear out the current filters, and add our own.
      .Filters.Clear
      .Filters.Add "Access Databases", "*.accdb"
 
      ' Show the dialog box. If the .Show method returns True, the
      ' user picked at least one file. If the .Show method returns
      ' False, the user clicked Cancel.
      If .Show = True Then
      
        Dim response
        For Each entry In .selectedItems  ' Grab the path/name of the selected file
            File = entry 'only last file is picked... not sure how to hande this more sensibly..
        Next

        response = MsgBox("Valittu tiedosto: " & File & ", oletko varma?", vbOKCancel, "Vahvista tietokannan valinta")
        
        'MsgBox (response)
        If response = 1 Then
        
            'then just refresh db link to actual database...
            
            Dim connectionString As String
            Dim tbl As TableDef
            Dim db As Database
            
            connectionString = ("MS Access;DATABASE=" & File)
            Set db = CurrentDb
            For Each tbl In db.TableDefs
                If Len(tbl.Connect) > 0 Then
                    'MsgBox tbl.Connect 'If you're getting errors, uncomment this to see connection string syntax
                    tbl.Connect = connectionString
                    tbl.RefreshLink
                End If
            Next
            ReplaceDatabaseFile = 0
        Else
            'peruutettu
            ReplaceDatabaseFile = 1
            
        End If
 
      Else
         ReplaceDatabaseFile = 2
      End If
   End With
End Function

Public Function CheckIfFileExists(ByVal path_ As String) As Boolean
    CheckIfFileExists = (Len(Dir(path_)) > 0)

End Function

Public Function CheckDatabaseFile()
    
     Dim DBFile As String
     
     DBFile = Common.GetDBPath
     
     If Not (CheckIfFileExists(DBFile)) Then
        MsgBox ("Tietokantatiedosto hukassa, tiedosto on etsittävä ennen kuin voit jatkaa. " & vbNewLine _
        & "Alkuperäinen tiedosto: " & DBFile)
        Dim succs
        succs = ReplaceDatabaseFile
        MsgBox ("Tietokantatiedosto vaihdettu, käynnistä jäsenrekisteri uusiksi että muutokset astuu voimaan")
        'DoCmd.CloseDatabase
        DoCmd.Quit
        'CurrentProject.CloseConnection
        'Dim project As String
        'project = CurrentProject.FullName
        'CurrentProject.CloseConnection
        'CurrentProject.OpenConnection
        'DoCmd.OpenForm ([Form_Tervetuloa])
     End If
     'MsgBox (Common.GetDBPath)
    'If Not (CheckIfFileExists(Application.CurrentProject.FullName)) Then
    '    MsgBox ("Tietokantatiedosto hukassa, tiedosto on etsittävä ennen kuin voit jatkaa.")
    '    Dim succs
    '    succs = ReplaceDatabaseFile
    'End If
    
End Function

Public Function GetDBPath() As String

    'thx http://www.ammara.com/access_image_faq/get_mdb_database_path.html
    'and
    '---------------------------------------------------------------------------------------
    ' Procedure : GetLinkedTablePath
    ' Author    : Daniel Pineault, CARDA Consultants Inc.
    ' Website   : http://www.cardaconsultants.com
    'plus:
    '---------------------------------------------------------------------------------------
    ' Procedure : GetCurrentPath
    ' DateTime  : 08/23/2010
    ' Author    : Rx
    ' Purpose   : Returns Current Path of a Linked Table in Access
    '---------------------------------------------------------------------------------------
    'https://www.access-programmers.co.uk/forums/threads/get-current-path-of-linked-table.198057/
    
    Dim table As String
    Dim db As DAO.Database
    Dim tdf As DAO.TableDef
    Dim i As Long
    
    table = "Kortit" ' just a fixed table name
    
    Set db = DBEngine.Workspaces(0).Databases(0)
    Set tdf = db.TableDefs(table)
    
    GetDBPath = Mid(tdf.Connect, InStr(1, tdf.Connect, "=") + 1)

End Function
