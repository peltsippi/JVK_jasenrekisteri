Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    Modal = NotDefault
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
    Width =5592
    DatasheetFontHeight =11
    ItemSuffix =49
    Left =4740
    Top =3456
    Right =22788
    Bottom =11712
    Picture ="bulldog_pienempi"
    RecSrcDt = Begin
        0xe76ced057dc5e540
    End
    RecordSource ="Hae nimi"
    Caption ="Linkitä kortti yhteystiedolle"
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
                    Left =36
                    Top =60
                    Width =2892
                    Height =480
                    FontSize =18
                    BackColor =14277081
                    Name ="Automaattinen_ylätunniste0"
                    Caption ="Korvaa kortti"
                    FontName ="Calibri Light"
                    GroupTable =1
                    GridlineColor =10921638
                    LayoutCachedLeft =36
                    LayoutCachedTop =60
                    LayoutCachedWidth =2928
                    LayoutCachedHeight =540
                    LayoutGroup =1
                    ThemeFontIndex =0
                    BackShade =85.0
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
                    Left =4305
                    Top =60
                    Width =1287
                    Height =300
                    BackColor =14277081
                    BorderColor =10921638
                    Name ="Automaattinen_päivämäärä"
                    ControlSource ="=Date()"
                    Format ="Short Date"
                    FontName ="Calibri"
                    GridlineColor =10921638

                    LayoutCachedLeft =4305
                    LayoutCachedTop =60
                    LayoutCachedWidth =5592
                    LayoutCachedHeight =360
                    BackShade =85.0
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
                    Left =4305
                    Top =360
                    Width =1287
                    Height =300
                    TabIndex =1
                    BackColor =14277081
                    BorderColor =10921638
                    Name ="Automaattinen_aika"
                    ControlSource ="=Time()"
                    Format ="Long Time"
                    FontName ="Calibri"
                    GridlineColor =10921638

                    LayoutCachedLeft =4305
                    LayoutCachedTop =360
                    LayoutCachedWidth =5592
                    LayoutCachedHeight =660
                    BackShade =85.0
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                End
            End
        End
        Begin Section
            Height =5442
            Name ="Tiedot"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin Rectangle
                    SpecialEffect =0
                    BackStyle =1
                    OverlapFlags =93
                    Left =226
                    Top =56
                    Width =4599
                    Height =5208
                    BackColor =12566463
                    Name ="Ruutu32"
                    GridlineColor =10921638
                    LayoutCachedLeft =226
                    LayoutCachedTop =56
                    LayoutCachedWidth =4825
                    LayoutCachedHeight =5264
                    BackShade =75.0
                    BorderThemeColorIndex =0
                    BorderShade =100.0
                End
                Begin TextBox
                    BorderWidth =3
                    OverlapFlags =215
                    TextFontCharSet =177
                    TextAlign =1
                    TextFontFamily =0
                    IMESentenceMode =3
                    Left =2136
                    Top =2172
                    Width =1704
                    Height =1380
                    TabIndex =1
                    ForeColor =4210752
                    Name ="Korttinro"
                    ValidationRule ="Like \"????\""
                    ValidationText ="Katso ohje vasemmalta!"
                    FontName ="Calibri"
                    OnChange ="[Event Procedure]"
                    GroupTable =2
                    BottomPadding =150
                    GridlineColor =10921638

                    LayoutCachedLeft =2136
                    LayoutCachedTop =2172
                    LayoutCachedWidth =3840
                    LayoutCachedHeight =3552
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
                            Left =336
                            Top =2172
                            Width =1704
                            Height =1380
                            BackColor =14277081
                            BorderColor =8355711
                            ForeColor =6710886
                            Name ="Selite3"
                            Caption ="Korvaavan kortin numero\015\012HUOM! \015\012Aina 4 numeroa!!!\015\012Esim: 0056"
                            FontName ="Calibri"
                            GroupTable =2
                            BottomPadding =150
                            GridlineColor =10921638
                            LayoutCachedLeft =336
                            LayoutCachedTop =2172
                            LayoutCachedWidth =2040
                            LayoutCachedHeight =3552
                            RowStart =1
                            RowEnd =1
                            LayoutGroup =2
                            BackShade =85.0
                            GroupTable =2
                        End
                    End
                End
                Begin TextBox
                    BorderWidth =3
                    OverlapFlags =215
                    TextFontCharSet =177
                    TextAlign =1
                    TextFontFamily =0
                    IMESentenceMode =3
                    Left =2136
                    Top =3768
                    Width =1704
                    Height =336
                    TabIndex =2
                    ForeColor =4210752
                    Name ="Puumerkki"
                    FontName ="Calibri"
                    OnChange ="[Event Procedure]"
                    GroupTable =2
                    BottomPadding =150
                    GridlineColor =10921638

                    LayoutCachedLeft =2136
                    LayoutCachedTop =3768
                    LayoutCachedWidth =3840
                    LayoutCachedHeight =4104
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
                            Left =336
                            Top =3768
                            Width =1704
                            Height =336
                            BackColor =14277081
                            BorderColor =8355711
                            ForeColor =6710886
                            Name ="Selite7"
                            Caption ="Puumerkki"
                            FontName ="Calibri"
                            GroupTable =2
                            BottomPadding =150
                            GridlineColor =10921638
                            LayoutCachedLeft =336
                            LayoutCachedTop =3768
                            LayoutCachedWidth =2040
                            LayoutCachedHeight =4104
                            RowStart =2
                            RowEnd =2
                            LayoutGroup =2
                            BackShade =85.0
                            GroupTable =2
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =215
                    TextFontCharSet =177
                    TextFontFamily =0
                    Left =340
                    Top =4195
                    Width =1764
                    Height =852
                    TabIndex =3
                    ForeColor =4210752
                    Name ="Korvaa"
                    Caption ="Korvaa"
                    OnClick ="[Event Procedure]"
                    FontName ="Calibri"
                    GridlineColor =10921638

                    LayoutCachedLeft =340
                    LayoutCachedTop =4195
                    LayoutCachedWidth =2104
                    LayoutCachedHeight =5047
                    BackColor =15123357
                    BorderColor =15123357
                    HoverColor =15652797
                    PressedColor =11957550
                    HoverForeColor =4210752
                    PressedForeColor =4210752
                    WebImagePaddingLeft =3
                    WebImagePaddingTop =3
                    WebImagePaddingRight =2
                    WebImagePaddingBottom =2
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =215
                    TextFontCharSet =177
                    TextFontFamily =0
                    Left =2267
                    Top =4195
                    Width =1764
                    Height =912
                    TabIndex =4
                    ForeColor =4210752
                    Name ="sulje"
                    Caption ="Peruuta"
                    FontName ="Calibri"
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
                                "nterfaceMacro For=\"sulje\" xmlns=\"http://schemas.microsoft.com/office/accessse"
                                "rvices/2009/11/application\"><Statements><Action Name=\"CloseWindow\"/></Stateme"
                                "nts></UserInterfaceMacro>"
                        End
                    End

                    LayoutCachedLeft =2267
                    LayoutCachedTop =4195
                    LayoutCachedWidth =4031
                    LayoutCachedHeight =5107
                    BackColor =15123357
                    BorderColor =15123357
                    HoverColor =15652797
                    PressedColor =11957550
                    HoverForeColor =4210752
                    PressedForeColor =4210752
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
                    Left =396
                    Top =170
                    Width =4082
                    Height =680
                    BackColor =62207
                    BorderColor =8355711
                    ForeColor =6710886
                    Name ="Selite30"
                    Caption ="Paina enter, tab tai klikkaa toiseen kenttään jos ei meinaa edetä!"
                    FontName ="Calibri"
                    GridlineColor =10921638
                    LayoutCachedLeft =396
                    LayoutCachedTop =170
                    LayoutCachedWidth =4478
                    LayoutCachedHeight =850
                    BackThemeColorIndex =-1
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =215
                    TextFontCharSet =177
                    TextAlign =1
                    TextFontFamily =0
                    IMESentenceMode =3
                    Left =2136
                    Top =1128
                    Width =1704
                    Height =828
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Muokkaus33"
                    DefaultValue ="=[Forms]![Tervetuloa]![Korttivalinta]"
                    FontName ="Calibri"
                    GroupTable =2
                    BottomPadding =150
                    GridlineColor =10921638

                    LayoutCachedLeft =2136
                    LayoutCachedTop =1128
                    LayoutCachedWidth =3840
                    LayoutCachedHeight =1956
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
                            Left =336
                            Top =1128
                            Width =1704
                            Height =828
                            BorderColor =8355711
                            ForeColor =6710886
                            Name ="Selite34"
                            Caption ="Vanhan kortin numero"
                            FontName ="Calibri"
                            GroupTable =2
                            BottomPadding =150
                            GridlineColor =10921638
                            LayoutCachedLeft =336
                            LayoutCachedTop =1128
                            LayoutCachedWidth =2040
                            LayoutCachedHeight =1956
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
    [Form_KorvaaKortti].Korvaa.Visible = False
    [Form_KorvaaKortti].Puumerkki.Visible = False
    
End Sub


Private Sub Puumerkki_Change()
    [Form_KorvaaKortti].Korvaa.Visible = True
End Sub


Private Sub Korttinro_Change()
    [Form_KorvaaKortti].Puumerkki.Visible = True
End Sub

Public Sub Korvaa_Click()

    Dim userNumber As Integer
    Dim oldCard As String
    Dim newCard As String
    Dim Puumerkki As String
    
    
    '1. Make sure all information is available
    
    If IsNull(Form_Tervetuloa.Yhteystietovalinta) Then
        MsgBox ("Yhteystietoa ei valittu. Valitse yhteystieto pääikkunassa!")
        Exit Sub
    Else
        userNumber = Form_Tervetuloa.Yhteystietovalinta.Value
    End If
    
    If IsNull(Form_Tervetuloa.Korttivalinta) Then
        MsgBox ("Vanhaa korttia ei valittu. Valitse korvattava kortti pääikkunassa!")
        Exit Sub
    Else
        oldCard = Form_Tervetuloa.Korttivalinta.Value
    End If
    
    If IsNull([Form_KorvaaKortti].Korttinro) Then
        MsgBox ("Korttinumeroa ei annettu, yritä uudestaan!")
        Exit Sub
    Else
        newCard = [Form_KorvaaKortti].Korttinro.Value
        
        If (Common.IsCardLinkedAlready(newCard)) Then
            MsgBox ("Kortti on jo linkitetty jollekulle!" & vbNewLine & "Käy tarvittaessa poistamassa kortin linkitys ensin." & vbNewLine & "Valitse kortti pääikkunassa -> Poista kortti")
            Exit Sub
        End If
        'check if already linked newCard!!
    End If
    
    If IsNull([Form_KorvaaKortti].Puumerkki) Then
        MsgBox ("Puumerkki ei voi olla tyhjä, yritä uudestaan!")
        Exit Sub
    Else
        Puumerkki = [Form_KorvaaKortti].Puumerkki.Value
    End If
    
    Dim succs
    succs = Common.SaveToLog(Puumerkki & " aloitti kortin " & oldCard & " korvaamisen kortilla " & newCard)
    
    
    '2. Ask final confirmation
    
    If MsgBox("Siirretään kaikki mahdollinen, mm. lataukset ja maksut " _
    & vbNewLine & "kortilta: " & oldCard & " kortille: " & newCard & vbNewLine & "Meneehän varmasti oikein?" _
    & vbNewLine & vbNewLline & "Tämä toiminto tekee kaiken mahdollisen automaattisesti kerralla.", vbYesNo) = vbNo Then Exit Sub
    
    '3. Link new card
    DoCmd.OpenForm "LisaaKortinLinkitys"
    Form_LisaaKortinLinkitys.Korttinro.Value = newCard
    Form_LisaaKortinLinkitys.Puumerkki.Value = Puumerkki
    'Form_LisaaKortinLinkitys.Puumerkki.Visible = True
    'Form_LisaaKortinLinkitys.Linkita.Visible = True
    Form_LisaaKortinLinkitys.Linkita_Click
    succs = Common.SaveToLog("Kortin korvaus - uusi kortti linkitetty")
    'MsgBox ("Uuden kortin linkitys ok")
        
    '4. prepare move of charges and payments
    
    Dim oldCardID As Integer
    Dim newCardID As Integer
    
    oldCardID = Common.FetchCardID(oldCard)
    newCardID = Common.FetchCardID(newCard)
    
    Dim table As String
    Dim values As String
    values = "Kortti = " & newCardID & ", Puumerkki = '" & Puumerkki & "'"
    Dim Target As String
    Target = "Kortti = " & oldCardID
    '5. move charges
    'MsgBox ("Move charges")
    table = "Lataukset"
    If (Common.CheckIfRecordFound(table, Target) > 0) Then 'do only when there is something to be moved..
        succs = Common.InsertOrUpdate(table, values, Target)
    End If
    
    succs = Common.SaveToLog("Kortin korvaus - lataukset siirretty kortilta " & oldCard & " kortille " & newCard & ".")
    
    '5.5: instructions for charging the new card
    DoCmd.OpenForm "LatausOhje"
    Form_LatausOhje.Save.Visible = False
    Form_LatausOhje.Cancel.Visible = False
    Form_LatausOhje.KorttiNumero.Value = "A" & newCard
    
    Dim expiration As Date
    expiration = Common.FetchExiprationDate(newCard)
    
    Dim chargeType As String
     
    Dim visitsLeft As Integer
    
    chargeType = Common.GetCardType(newCard)
    'MsgBox (chargeType)
    
    Form_LatausOhje.Voimassa.Value = expiration
    
    If (InStr(1, chargeType, "krt")) Then
        cardType = "Määräkortti"
        visitsLeft = InputBox("Kuinka monta käyntikertaa kortilla " & oldCard & " jäljellä?")
        'add query for how many left!!!
        succs = Common.SaveToLog("Kortilla " & oldCard & " oli " & visitsLeft & " latausta jäljellä.")
        Form_LatausOhje.Maara.Value = visitsLeft
    Else
        cardType = "Kausikortti"
    End If
    
    If (InStr(1, chargeType, "ap")) Then
        timeGroup = "Aamupäivä ma-su"
    Else
        timeGroup = "Normaali"
       
    End If
    
    'Dim cardType As String ' how to get this ?!?
    'Dim timeGroup As String 'how to get this ?!?!
    
    Form_LatausOhje.KorttiTyyppi.Value = cardType
    Form_LatausOhje.AikaRyhma.Value = timeGroup
    
    
    'jos kertakortti, kysy vanhalla kortilla oleva jäljellä olevien käyntien määrä!!
    'Form_LatausOhje.Maara.Value = "" 'how to get this ?!?!
    MsgBox ("Ohje uuden kortin lataamiseksi on auki kunnes painat OK tästä!" & vbNewLine & "Se uusi kortti kannattaa ladata oikeasti nyt")
    DoCmd.Close
    
    succs = Common.SaveToLog("Kortin korvaus - latausohje kuitattu luetuksi")
    
    '6. move payments
    'MsgBox ("Move payments")
    table = "Maksut"
    'MsgBox ("Table: " & Table & " values : " & Values & " target: " & Target)
    If (Common.CheckIfRecordFound(table, Target) > 0) Then 'do only if there are something to be moved..
        succs = Common.InsertOrUpdate(table, values, Target)
    End If
    'MsgBox ("Lataukset ja maksut siirretty vanhalta kortilta uudelle")
    succs = Common.SaveToLog("Kortin korvaus - maksut siirretty kortilta " & oldCard & " kortille " & newCard & ".")
    
    
    '7. mark old card as missing/broken
    
    DoCmd.OpenForm "PoistaKortinLinkitys"
    Form_PoistaKortinLinkitys.Puumerkki.Value = Puumerkki
    Form_PoistaKortinLinkitys.discard.Value = True
    Form_PoistaKortinLinkitys.Muistiinpano.Value = "Automaattinen kortin korvaustoiminto teki"
    Form_Tervetuloa.Korttivalinta.Value = oldCard 'Just in case card number updates for some reason..
    Form_PoistaKortinLinkitys.Poista_Click
    
    succs = Common.SaveToLog("Kortin korvaus - kortin " & oldCard & " linkitys poistettu. Valmista tuli!")
    
    DoCmd.Close
    Common.SendMessageToMainScreen ("Kortti " & oldCard & " korvattu kortilla " & newCard & ".")
    
    
End Sub
