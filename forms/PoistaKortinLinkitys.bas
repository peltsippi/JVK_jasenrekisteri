Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    RecordSelectors = NotDefault
    NavigationButtons = NotDefault
    CloseButton = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    PictureType =2
    GridY =10
    Width =7880
    DatasheetFontHeight =11
    ItemSuffix =53
    Left =2556
    Top =3468
    Right =22788
    Bottom =11712
    Picture ="bulldog_pienempi"
    RecSrcDt = Begin
        0xa22cdf047cc5e540
    End
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Calibri"
    AllowDatasheetView =0
    FilterOnLoad =0
    ShowPageMargins =0
    DisplayOnSharePointSite =1
    DatasheetAlternateBackColor =15921906
    DatasheetGridlinesColor12 =0
    FitToScreen =1
    DatasheetBackThemeColorIndex =1
    BorderThemeColorIndex =3
    ThemeFontIndex =1
    ForeThemeColorIndex =0
    AlternateBackThemeColorIndex =1
    AlternateBackShade =95.0
    Begin
        Begin Label
            BackStyle =0
            FontSize =11
            FontName ="Segoe UI"
            ThemeFontIndex =1
            BackThemeColorIndex =1
            BorderThemeColorIndex =0
            BorderTint =50.0
            ForeThemeColorIndex =0
            ForeTint =60.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin Rectangle
            SpecialEffect =3
            BackStyle =0
            BorderLineStyle =0
            Width =850
            Height =850
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin Image
            BackStyle =0
            OldBorderStyle =0
            BorderLineStyle =0
            SizeMode =3
            PictureAlignment =2
            Width =1701
            Height =1701
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin CommandButton
            Width =1701
            Height =283
            FontSize =11
            FontWeight =400
            FontName ="Segoe UI"
            ForeThemeColorIndex =0
            ForeTint =75.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
            UseTheme =1
            Shape =1
            Gradient =12
            BackThemeColorIndex =4
            BackTint =60.0
            BorderLineStyle =0
            BorderThemeColorIndex =4
            BorderTint =60.0
            ThemeFontIndex =1
            HoverThemeColorIndex =4
            HoverTint =40.0
            PressedThemeColorIndex =4
            PressedShade =75.0
            HoverForeThemeColorIndex =0
            HoverForeTint =75.0
            PressedForeThemeColorIndex =0
            PressedForeTint =75.0
        End
        Begin OptionButton
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
            BorderThemeColorIndex =1
            BorderShade =65.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin CheckBox
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
            BorderThemeColorIndex =1
            BorderShade =65.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin TextBox
            AddColon = NotDefault
            FELineBreak = NotDefault
            BorderLineStyle =0
            Width =1701
            LabelX =-1701
            FontSize =11
            FontName ="Segoe UI"
            AsianLineBreak =1
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            ThemeFontIndex =1
            ForeThemeColorIndex =0
            ForeTint =75.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin EmptyCell
            Height =240
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin FormHeader
            Height =660
            Name ="LomakkeenYlätunniste"
            AutoHeight =1
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =2
            BackTint =20.0
            Begin
                Begin Label
                    BackStyle =1
                    OverlapFlags =85
                    TextFontCharSet =177
                    TextAlign =1
                    TextFontFamily =0
                    Left =300
                    Top =60
                    Width =3156
                    Height =460
                    FontSize =18
                    BackColor =15921906
                    Name ="Automaattinen_ylätunniste0"
                    Caption ="Poista kortin linkitys"
                    FontName ="Calibri Light"
                    GroupTable =1
                    GridlineColor =10921638
                    LayoutCachedLeft =300
                    LayoutCachedTop =60
                    LayoutCachedWidth =3456
                    LayoutCachedHeight =520
                    LayoutGroup =1
                    ThemeFontIndex =0
                    BackShade =95.0
                    BorderThemeColorIndex =2
                    BorderTint =100.0
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                    GroupTable =1
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =93
                    TextFontCharSet =177
                    TextAlign =3
                    TextFontFamily =0
                    IMESentenceMode =3
                    Left =5889
                    Top =60
                    Width =1635
                    Height =300
                    BackColor =15921906
                    BorderColor =10921638
                    Name ="Automaattinen_päivämäärä"
                    ControlSource ="=Date()"
                    Format ="Short Date"
                    FontName ="Calibri"
                    GridlineColor =10921638

                    LayoutCachedLeft =5889
                    LayoutCachedTop =60
                    LayoutCachedWidth =7524
                    LayoutCachedHeight =360
                    BackShade =95.0
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =87
                    TextFontCharSet =177
                    TextAlign =3
                    TextFontFamily =0
                    IMESentenceMode =3
                    Left =5889
                    Top =360
                    Width =1635
                    Height =300
                    TabIndex =1
                    BackColor =15921906
                    BorderColor =10921638
                    Name ="Automaattinen_aika"
                    ControlSource ="=Time()"
                    Format ="Long Time"
                    FontName ="Calibri"
                    GridlineColor =10921638

                    LayoutCachedLeft =5889
                    LayoutCachedTop =360
                    LayoutCachedWidth =7524
                    LayoutCachedHeight =660
                    BackShade =95.0
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                End
            End
        End
        Begin Section
            Height =5102
            Name ="Tiedot"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin Rectangle
                    BackStyle =1
                    OverlapFlags =93
                    Left =283
                    Top =56
                    Width =7317
                    Height =3857
                    BackColor =14277081
                    BorderColor =10921638
                    Name ="Ruutu42"
                    GridlineColor =10921638
                    LayoutCachedLeft =283
                    LayoutCachedTop =56
                    LayoutCachedWidth =7600
                    LayoutCachedHeight =3913
                    BackShade =85.0
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =215
                    TextFontCharSet =177
                    TextAlign =1
                    TextFontFamily =0
                    IMESentenceMode =3
                    Left =3648
                    Top =360
                    Width =3276
                    Height =300
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="korttinro"
                    ControlSource ="=[Forms]![Tervetuloa]![Korttivalinta]"
                    FontName ="Calibri"
                    GroupTable =2
                    GridlineColor =10921638

                    LayoutCachedLeft =3648
                    LayoutCachedTop =360
                    LayoutCachedWidth =6924
                    LayoutCachedHeight =660
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =2
                    GroupTable =2
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =215
                            TextFontCharSet =177
                            TextAlign =1
                            TextFontFamily =0
                            Left =360
                            Top =360
                            Width =3204
                            Height =300
                            BackColor =15921906
                            BorderColor =8355711
                            ForeColor =6710886
                            Name ="Selite5"
                            Caption ="Olet poistamassa tämän kortin linkitystä: "
                            FontName ="Calibri"
                            GroupTable =2
                            GridlineColor =10921638
                            LayoutCachedLeft =360
                            LayoutCachedTop =360
                            LayoutCachedWidth =3564
                            LayoutCachedHeight =660
                            LayoutGroup =2
                            BackShade =95.0
                            GroupTable =2
                        End
                    End
                End
                Begin TextBox
                    BorderWidth =2
                    OverlapFlags =215
                    TextFontCharSet =177
                    TextAlign =1
                    TextFontFamily =0
                    IMESentenceMode =3
                    Left =3648
                    Top =744
                    Width =3276
                    Height =336
                    TabIndex =1
                    ForeColor =4210752
                    Name ="Puumerkki"
                    FontName ="Calibri"
                    OnChange ="[Event Procedure]"
                    GroupTable =2
                    GridlineColor =10921638

                    LayoutCachedLeft =3648
                    LayoutCachedTop =744
                    LayoutCachedWidth =6924
                    LayoutCachedHeight =1080
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =2
                    BorderThemeColorIndex =5
                    BorderShade =100.0
                    GroupTable =2
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =215
                            TextFontCharSet =177
                            TextAlign =1
                            TextFontFamily =0
                            Left =360
                            Top =744
                            Width =3204
                            Height =336
                            BackColor =15921906
                            BorderColor =8355711
                            ForeColor =6710886
                            Name ="Selite15"
                            Caption ="Puumerkki"
                            FontName ="Calibri"
                            GroupTable =2
                            GridlineColor =10921638
                            LayoutCachedLeft =360
                            LayoutCachedTop =744
                            LayoutCachedWidth =3564
                            LayoutCachedHeight =1080
                            RowStart =1
                            RowEnd =1
                            LayoutGroup =2
                            BackShade =95.0
                            GroupTable =2
                        End
                    End
                End
                Begin TextBox
                    BorderWidth =2
                    OverlapFlags =215
                    TextFontCharSet =177
                    TextAlign =1
                    TextFontFamily =0
                    IMESentenceMode =3
                    Left =3648
                    Top =1164
                    Width =3276
                    Height =1380
                    TabIndex =2
                    ForeColor =4210752
                    Name ="Muistiinpano"
                    FontName ="Calibri"
                    OnChange ="[Event Procedure]"
                    GroupTable =2
                    GridlineColor =10921638

                    LayoutCachedLeft =3648
                    LayoutCachedTop =1164
                    LayoutCachedWidth =6924
                    LayoutCachedHeight =2544
                    RowStart =2
                    RowEnd =2
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =2
                    BorderThemeColorIndex =5
                    BorderShade =100.0
                    GroupTable =2
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =215
                            TextFontCharSet =177
                            TextAlign =1
                            TextFontFamily =0
                            Left =360
                            Top =1164
                            Width =3204
                            Height =1380
                            BackColor =15921906
                            BorderColor =8355711
                            ForeColor =6710886
                            Name ="Muistiinpano_selite"
                            Caption ="Muistiinpanot\015\012(ei jää talteen jos kortti merkataan\015\012rikkinäiseksi/k"
                                "adonneeksi!\015\012Kirjoita tähän silti jotain että nappi ilmestyy)"
                            FontName ="Calibri"
                            GroupTable =2
                            GridlineColor =10921638
                            LayoutCachedLeft =360
                            LayoutCachedTop =1164
                            LayoutCachedWidth =3564
                            LayoutCachedHeight =2544
                            RowStart =2
                            RowEnd =2
                            LayoutGroup =2
                            BackShade =95.0
                            GroupTable =2
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontCharSet =177
                    TextFontFamily =0
                    Left =360
                    Top =4056
                    Width =3204
                    Height =576
                    TabIndex =4
                    ForeColor =4210752
                    Name ="Poista"
                    Caption ="Poista linkitys"
                    OnClick ="[Event Procedure]"
                    FontName ="Calibri"
                    GroupTable =2
                    GridlineColor =10921638

                    LayoutCachedLeft =360
                    LayoutCachedTop =4056
                    LayoutCachedWidth =3564
                    LayoutCachedHeight =4632
                    RowStart =4
                    RowEnd =4
                    LayoutGroup =2
                    BackColor =15123357
                    BorderColor =15123357
                    HoverColor =15652797
                    PressedColor =11957550
                    HoverForeColor =4210752
                    PressedForeColor =4210752
                    GroupTable =2
                    WebImagePaddingLeft =3
                    WebImagePaddingTop =3
                    WebImagePaddingRight =2
                    WebImagePaddingBottom =2
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontCharSet =177
                    TextFontFamily =0
                    Left =3648
                    Top =4056
                    Width =3276
                    Height =576
                    TabIndex =5
                    ForeColor =4210752
                    Name ="Komento35"
                    Caption ="Sulje"
                    FontName ="Calibri"
                    ControlTipText ="Sulje lomake"
                    GroupTable =2
                    GridlineColor =10921638
                    OnClickEmMacro = Begin
                        Version =196611
                        ColumnsShown =8
                        Begin
                            Action ="Close"
                            Argument ="-1"
                            Argument =""
                            Argument ="0"
                        End
                        Begin
                            Comment ="_AXL:<?xml version=\"1.0\" encoding=\"UTF-16\" standalone=\"no\"?>\015\012<UserI"
                                "nterfaceMacro For=\"Komento35\" xmlns=\"http://schemas.microsoft.com/office/acce"
                                "ssservices/2009/11/application\"><Statements><Action Name=\"CloseWindow\"/></Sta"
                                "tements></UserInterfaceMacro>"
                        End
                    End

                    LayoutCachedLeft =3648
                    LayoutCachedTop =4056
                    LayoutCachedWidth =6924
                    LayoutCachedHeight =4632
                    RowStart =4
                    RowEnd =4
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =2
                    BackColor =15123357
                    BorderColor =15123357
                    HoverColor =15652797
                    PressedColor =11957550
                    HoverForeColor =4210752
                    PressedForeColor =4210752
                    GroupTable =2
                    WebImagePaddingLeft =3
                    WebImagePaddingTop =3
                    WebImagePaddingRight =2
                    WebImagePaddingBottom =2
                    Overlaps =1
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =215
                    TextFontCharSet =177
                    TextFontFamily =0
                    Left =340
                    Top =56
                    Width =7030
                    Height =284
                    BackColor =62207
                    BorderColor =8355711
                    ForeColor =6710886
                    Name ="Selite40"
                    Caption ="Paina enter, tab tai klikkaa toista kenttää päästäksesi eteenpäin!"
                    FontName ="Calibri"
                    GridlineColor =10921638
                    LayoutCachedLeft =340
                    LayoutCachedTop =56
                    LayoutCachedWidth =7370
                    LayoutCachedHeight =340
                    BackThemeColorIndex =-1
                End
                Begin CheckBox
                    OverlapFlags =215
                    Left =3648
                    Top =2604
                    Width =3276
                    Height =1380
                    TabIndex =3
                    BorderColor =10921638
                    Name ="discard"
                    DefaultValue ="=False"
                    GroupTable =2
                    GridlineColor =10921638

                    LayoutCachedLeft =3648
                    LayoutCachedTop =2604
                    LayoutCachedWidth =6924
                    LayoutCachedHeight =3984
                    RowStart =3
                    RowEnd =3
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =2
                    GroupTable =2
                    Begin
                        Begin Label
                            OverlapFlags =215
                            TextFontCharSet =177
                            TextAlign =1
                            TextFontFamily =0
                            Left =360
                            Top =2604
                            Width =3204
                            Height =1380
                            BorderColor =8355711
                            ForeColor =6710886
                            Name ="Selite46"
                            Caption ="Kortti rikki/kadonnut?\015\012Täppä: kortti on rikki/kadonnut ja sitä ei voi enä"
                                "ä käyttää. Jää jäsenen tietoihin talteen esim siltä varalta, että kadonnut kortt"
                                "i löytyy myöhemmin!"
                            FontName ="Calibri"
                            GroupTable =2
                            GridlineColor =10921638
                            LayoutCachedLeft =360
                            LayoutCachedTop =2604
                            LayoutCachedWidth =3564
                            LayoutCachedHeight =3984
                            RowStart =3
                            RowEnd =3
                            LayoutGroup =2
                            GroupTable =2
                        End
                    End
                End
            End
        End
        Begin FormFooter
            Height =0
            Name ="LomakkeenAlatunniste"
            AutoHeight =1
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
        End
    End
End
CodeBehindForm
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database


    
Private Sub Form_Open(Cancel As Integer)
    [Form_PoistaKortinLinkitys].Poista.Visible = False
    [Form_PoistaKortinLinkitys].Muistiinpano.Visible = False
    
End Sub


Private Sub Muistiinpano_Change()
    [Form_PoistaKortinLinkitys].Poista.Visible = True
End Sub

Private Sub Poista_Click()
    Dim cardnumber As String
    Dim cardID As Integer
    Dim SQLQuery As String
    Dim Puumerkki As String
    Dim Muistiinpano As String
    
    If IsNull(Form_Tervetuloa.Korttivalinta) Then
        MsgBox ("Korttia ei valittu. Valitse kortti pääikkunassa!")
        Exit Sub
    Else
        cardnumber = Form_Tervetuloa.Korttivalinta.Value
    End If
    
    If IsNull([Form_PoistaKortinLinkitys].Puumerkki) Then
        MsgBox ("Puumerkki ei voi linkitystä poistaessa olla tyhjä!")
        Exit Sub
    Else
        Puumerkki = [Form_PoistaKortinLinkitys].Puumerkki.Value
    End If
    
    
    If IsNull([Form_PoistaKortinLinkitys].Muistiinpano) Then
        MsgBox ("Muistiinpanokenttä ei voi linkitystä poistaessa olla tyhjä!")
        Exit Sub
    Else
        Muistiinpano = [Form_PoistaKortinLinkitys].Muistiinpano.Value
    End If
        
    
    Dim newOwner As Integer
    newOwner = 0 ' kortille vaan määritellään omistajaksi 0 eli nobody...
    
    Dim korttiID As Integer
    'korttiID = Common.FetchGeneralID("Kortit", "CID", "Kortti = '" & cardnumber & "'")
    korttiID = Common.FetchCardID(cardnumber)
    
    Dim deletebool As Boolean
    deletebool = [Form_PoistaKortinLinkitys].discard.Value
    
    If deletebool Then
        'Dim query As String
        Dim largestDate As Date
        largestDate = Common.FetchExiprationDate(cardnumber)
        'MsgBox ("Suurin päivämäärä on " & largestDate)
        
        largestDate = DateAdd("d", 1, largestDate) 'add 1 more than previous charge!
        Dim feedback As Integer
        Dim table2 As String
        table2 = "Lataukset"
        Dim values2 As String
        
        values2 = " Kortti = '" & korttiID & "'" _
        & ", Voimassa = '" & DateAdd("d", 1, largestDate) & "'" _
        & ", Lataaja = '" & Puumerkki & "'" _
        & ", Korttityyppi = 'RIKKI/KADONNUT'" _
        & ", KortinArvo = '0 €'" _
        & ", Ajankohta = '" & largestDate & "' "
        
        feedback = Common.InsertOrUpdate(table2, values2, "")
        
        If (Not (feedback)) Then
            MsgBox ("Kortin poistossa tapahtui virhe, käy muokkaamassa manuaalisesti kortin tietoihin että se on rikki!")
        Else
            MsgBox ("Kortti merkattu poistetuksi!")
        End If
        
        
        
    End If
    
    
    Dim success As Boolean
    Dim table As String
    Dim values As String
    Dim Target As String
    
    table = "Kortit"
    values = "Omistaja = '" & newOwner & "' , " _
    & "PVM = '" & Date & "' ," _
    & "Puumerkki = '" & Puumerkki & "' ," _
    & "Muistiinpanot = '" & Muistiinpano & "' "
    
    Target = "Kortti = '" & cardnumber & "'"
    
    'Jätetäänpä nämä sittenkin talteen käyttäjän alle heh heh...
    
    If (Not (deletebool)) Then
        success = Common.InsertOrUpdate(table, values, Target)
    
        If Not (success) Then
            MsgBox ("Jotain meni pieleen sori siitä!")
        
    
        End If
    
    End If
    
    Dim logOutput As String
    logOutput = "Puumerkki " & Puumerkki & " poisti kortin " & cardnumber & " linkityksen, muistiinpanot: " & Muistiinpano & " ja kortti merkattu poistetuksi: " & deletebool
    success = Common.SaveToLog(logOutput)
    
    success = Common.SendMessageToMainScreen("Kortin " & cardnumber & " linkitys poistettu!")
    
    DoCmd.Close
    
End Sub





Private Sub Puumerkki_Change()
    [Form_PoistaKortinLinkitys].Muistiinpano.Visible = True
End Sub
