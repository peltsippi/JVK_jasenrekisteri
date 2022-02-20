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
    Width =12018
    DatasheetFontHeight =11
    ItemSuffix =61
    Left =4044
    Top =3468
    Right =17484
    Bottom =11712
    Picture ="bulldog_pienempi"
    RecSrcDt = Begin
        0xdfeefa493cc6e540
    End
    RecordSource ="SELECT Lataukset.Voimassa FROM Lataukset; "
    Caption ="Kirjaa maksu kortille"
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
        Begin ListBox
            BorderLineStyle =0
            Width =1701
            Height =1417
            LabelX =-1701
            FontSize =11
            FontName ="Segoe UI"
            AllowValueListEdits =1
            InheritValueList =1
            ThemeFontIndex =1
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            ForeThemeColorIndex =0
            ForeTint =75.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin Subform
            BorderLineStyle =0
            Width =1701
            Height =1701
            BorderThemeColorIndex =1
            GridlineThemeColorIndex =1
            GridlineShade =65.0
            BorderShade =65.0
            ShowPageHeaderAndPageFooter =1
        End
        Begin EmptyCell
            Height =240
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin NavigationControl
            BorderWidth =1
            BorderLineStyle =0
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin NavigationButton
            Width =283
            Height =283
            ForeColor =-2
            GridlineThemeColorIndex =1
            GridlineShade =65.0
            HoverColor =-2
            HoverThemeColorIndex =2
            HoverTint =20.0
            PressedColor =-2
            PressedThemeColorIndex =2
            PressedTint =60.0
            HoverForeColor =-2
            HoverForeThemeColorIndex =0
            HoverForeTint =75.0
            PressedForeColor =-2
            PressedForeThemeColorIndex =0
            PressedForeTint =75.0
            BackThemeColorIndex =1
            OldBorderStyle =0
            BorderLineStyle =0
            BorderThemeColorIndex =3
            BorderShade =90.0
            ThemeFontIndex =1
            FontName ="Segoe UI"
            FontWeight =400
            FontSize =11
            ForeThemeColorIndex =0
            ForeTint =75.0
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
                    Width =2856
                    Height =460
                    FontSize =18
                    BackColor =15921906
                    Name ="Automaattinen_ylätunniste0"
                    Caption ="Rekisteröi maksu"
                    FontName ="Calibri Light"
                    GroupTable =2
                    GridlineColor =10921638
                    LayoutCachedLeft =300
                    LayoutCachedTop =60
                    LayoutCachedWidth =3156
                    LayoutCachedHeight =520
                    LayoutGroup =2
                    ThemeFontIndex =0
                    BackShade =95.0
                    BorderThemeColorIndex =2
                    BorderTint =100.0
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                    GroupTable =2
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
                    Left =8841
                    Top =60
                    Width =2523
                    Height =300
                    BorderColor =10921638
                    Name ="Automaattinen_päivämäärä"
                    ControlSource ="=Date()"
                    Format ="Long Date"
                    FontName ="Calibri"
                    GridlineColor =10921638

                    LayoutCachedLeft =8841
                    LayoutCachedTop =60
                    LayoutCachedWidth =11364
                    LayoutCachedHeight =360
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
                    Left =8841
                    Top =360
                    Width =2523
                    Height =300
                    TabIndex =1
                    BorderColor =10921638
                    Name ="Automaattinen_aika"
                    ControlSource ="=Time()"
                    Format ="Long Time"
                    FontName ="Calibri"
                    GridlineColor =10921638

                    LayoutCachedLeft =8841
                    LayoutCachedTop =360
                    LayoutCachedWidth =11364
                    LayoutCachedHeight =660
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                End
            End
        End
        Begin Section
            Height =6973
            Name ="Tiedot"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin Rectangle
                    BackStyle =1
                    OverlapFlags =93
                    Left =283
                    Top =172
                    Width =11280
                    Height =5446
                    BackColor =14277081
                    BorderColor =10921638
                    Name ="Ruutu59"
                    GridlineColor =10921638
                    LayoutCachedLeft =283
                    LayoutCachedTop =172
                    LayoutCachedWidth =11563
                    LayoutCachedHeight =5618
                    BackShade =85.0
                End
                Begin TextBox
                    Locked = NotDefault
                    BorderWidth =3
                    OverlapFlags =215
                    TextFontCharSet =177
                    TextAlign =1
                    TextFontFamily =0
                    IMESentenceMode =3
                    Left =1692
                    Top =336
                    Width =9660
                    Height =348
                    LeftMargin =44
                    TopMargin =22
                    RightMargin =44
                    BottomMargin =22
                    ForeColor =4210752
                    Name ="Kortti"
                    ControlSource ="=[Forms]![Tervetuloa]![Korttivalinta]"
                    StatusBarText ="Kortti, haetaan Kortit -taulusta id"
                    DefaultValue ="\"\""
                    FontName ="Calibri"
                    GroupTable =1
                    BottomPadding =150
                    GridlineColor =10921638

                    LayoutCachedLeft =1692
                    LayoutCachedTop =336
                    LayoutCachedWidth =11352
                    LayoutCachedHeight =684
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    BorderThemeColorIndex =5
                    BorderShade =100.0
                    GroupTable =1
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =215
                            TextFontCharSet =177
                            TextAlign =1
                            TextFontFamily =0
                            Left =336
                            Top =336
                            Width =1260
                            Height =348
                            LeftMargin =44
                            TopMargin =22
                            RightMargin =44
                            BottomMargin =22
                            BorderColor =8355711
                            ForeColor =6710886
                            Name ="Selite3"
                            Caption ="Kortti"
                            FontName ="Calibri"
                            GroupTable =1
                            BottomPadding =150
                            GridlineColor =10921638
                            LayoutCachedLeft =336
                            LayoutCachedTop =336
                            LayoutCachedWidth =1596
                            LayoutCachedHeight =684
                            LayoutGroup =1
                            BackThemeColorIndex =-1
                            GroupTable =1
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
                    Left =1692
                    Top =900
                    Width =9660
                    Height =348
                    TabIndex =1
                    LeftMargin =44
                    TopMargin =22
                    RightMargin =44
                    BottomMargin =22
                    ForeColor =4210752
                    Name ="Summa"
                    Format ="#,##0.00 €;-#,##0.00 €"
                    StatusBarText ="Paljoin maksettu"
                    FontName ="Calibri"
                    OnChange ="[Event Procedure]"
                    GroupTable =1
                    BottomPadding =150
                    GridlineColor =10921638

                    LayoutCachedLeft =1692
                    LayoutCachedTop =900
                    LayoutCachedWidth =11352
                    LayoutCachedHeight =1248
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    BorderThemeColorIndex =5
                    BorderShade =100.0
                    GroupTable =1
                    CurrencySymbol ="€"
                    ColLCID =1035
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =215
                            TextFontCharSet =177
                            TextAlign =1
                            TextFontFamily =0
                            Left =336
                            Top =900
                            Width =1260
                            Height =348
                            LeftMargin =44
                            TopMargin =22
                            RightMargin =44
                            BottomMargin =22
                            BorderColor =8355711
                            ForeColor =6710886
                            Name ="Selite6"
                            Caption ="Summa"
                            FontName ="Calibri"
                            GroupTable =1
                            BottomPadding =150
                            GridlineColor =10921638
                            LayoutCachedLeft =336
                            LayoutCachedTop =900
                            LayoutCachedWidth =1596
                            LayoutCachedHeight =1248
                            RowStart =1
                            RowEnd =1
                            LayoutGroup =1
                            BackThemeColorIndex =-1
                            GroupTable =1
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
                    Left =1692
                    Top =4512
                    Width =9660
                    Height =564
                    TabIndex =4
                    LeftMargin =44
                    TopMargin =22
                    RightMargin =44
                    BottomMargin =22
                    ForeColor =4210752
                    Name ="Puumerkki"
                    StatusBarText ="Kuittaajan puumerkit"
                    FontName ="Calibri"
                    OnChange ="[Event Procedure]"
                    GroupTable =1
                    BottomPadding =150
                    GridlineColor =10921638

                    LayoutCachedLeft =1692
                    LayoutCachedTop =4512
                    LayoutCachedWidth =11352
                    LayoutCachedHeight =5076
                    RowStart =4
                    RowEnd =4
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    BorderThemeColorIndex =5
                    BorderShade =100.0
                    GroupTable =1
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =215
                            TextFontCharSet =177
                            TextAlign =1
                            TextFontFamily =0
                            Left =336
                            Top =4512
                            Width =1260
                            Height =564
                            LeftMargin =44
                            TopMargin =22
                            RightMargin =44
                            BottomMargin =22
                            BorderColor =8355711
                            ForeColor =6710886
                            Name ="Selite9"
                            Caption ="Puumerkki"
                            FontName ="Calibri"
                            GroupTable =1
                            BottomPadding =150
                            GridlineColor =10921638
                            LayoutCachedLeft =336
                            LayoutCachedTop =4512
                            LayoutCachedWidth =1596
                            LayoutCachedHeight =5076
                            RowStart =4
                            RowEnd =4
                            LayoutGroup =1
                            BackThemeColorIndex =-1
                            GroupTable =1
                        End
                    End
                End
                Begin EmptyCell
                    Left =336
                    Top =1464
                    Width =1260
                    Height =2268
                    Name ="TyhjäSolu26"
                    GroupTable =1
                    BottomPadding =150
                    GridlineColor =10921638
                    LayoutCachedLeft =336
                    LayoutCachedTop =1464
                    LayoutCachedWidth =1596
                    LayoutCachedHeight =3732
                    RowStart =2
                    RowEnd =2
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin ListBox
                    OverlapFlags =215
                    TextFontCharSet =177
                    TextFontFamily =0
                    BorderWidth =3
                    IMESentenceMode =3
                    Left =1692
                    Top =1464
                    Width =9660
                    Height =2268
                    TabIndex =2
                    ForeColor =4210752
                    Name ="Maksutapa"
                    RowSourceType ="Table/Query"
                    RowSource ="HaeMaksutavat"
                    ColumnWidths ="1440"
                    FontName ="Calibri"
                    OnClick ="[Event Procedure]"
                    GroupTable =1
                    BottomPadding =150
                    GridlineColor =10921638
                    AllowValueListEdits =0

                    LayoutCachedLeft =1692
                    LayoutCachedTop =1464
                    LayoutCachedWidth =11352
                    LayoutCachedHeight =3732
                    RowStart =2
                    RowEnd =2
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    BorderThemeColorIndex =5
                    BorderShade =100.0
                    GroupTable =1
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =215
                    TextFontCharSet =177
                    TextFontFamily =0
                    Left =396
                    Top =1814
                    Width =1116
                    Height =314
                    BorderColor =8355711
                    ForeColor =6710886
                    Name ="Maksutapa_Selite"
                    Caption ="Maksutapa"
                    FontName ="Calibri"
                    GridlineColor =10921638
                    LayoutCachedLeft =396
                    LayoutCachedTop =1814
                    LayoutCachedWidth =1512
                    LayoutCachedHeight =2128
                    BackThemeColorIndex =-1
                End
                Begin TextBox
                    BorderWidth =3
                    OverlapFlags =215
                    TextFontCharSet =177
                    TextAlign =1
                    TextFontFamily =0
                    IMESentenceMode =3
                    Left =1692
                    Top =3948
                    Width =9660
                    Height =348
                    TabIndex =3
                    LeftMargin =44
                    TopMargin =22
                    RightMargin =44
                    BottomMargin =22
                    ForeColor =4210752
                    Name ="PVM"
                    Format ="Short Date"
                    StatusBarText ="Kuittaajan puumerkit"
                    DefaultValue ="=Date()"
                    FontName ="Calibri"
                    GroupTable =1
                    BottomPadding =150
                    GridlineColor =10921638

                    LayoutCachedLeft =1692
                    LayoutCachedTop =3948
                    LayoutCachedWidth =11352
                    LayoutCachedHeight =4296
                    RowStart =3
                    RowEnd =3
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    BorderThemeColorIndex =5
                    BorderShade =100.0
                    GroupTable =1
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =215
                            TextFontCharSet =177
                            TextAlign =1
                            TextFontFamily =0
                            Left =336
                            Top =3948
                            Width =1260
                            Height =348
                            LeftMargin =44
                            TopMargin =22
                            RightMargin =44
                            BottomMargin =22
                            BorderColor =8355711
                            ForeColor =6710886
                            Name ="Selite47"
                            Caption ="PVM"
                            FontName ="Calibri"
                            GroupTable =1
                            BottomPadding =150
                            GridlineColor =10921638
                            LayoutCachedLeft =336
                            LayoutCachedTop =3948
                            LayoutCachedWidth =1596
                            LayoutCachedHeight =4296
                            RowStart =3
                            RowEnd =3
                            LayoutGroup =1
                            BackThemeColorIndex =-1
                            GroupTable =1
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontCharSet =177
                    TextFontFamily =0
                    Left =566
                    Top =5669
                    Width =5783
                    Height =1247
                    TabIndex =5
                    ForeColor =4210752
                    Name ="Tallenna"
                    Caption ="OK"
                    OnClick ="[Event Procedure]"
                    FontName ="Calibri"
                    GridlineColor =10921638

                    LayoutCachedLeft =566
                    LayoutCachedTop =5669
                    LayoutCachedWidth =6349
                    LayoutCachedHeight =6916
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
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontCharSet =177
                    TextFontFamily =0
                    Left =6803
                    Top =5669
                    Width =4875
                    Height =1247
                    TabIndex =6
                    ForeColor =4210752
                    Name ="Cancel"
                    Caption ="Sulje"
                    FontName ="Calibri"
                    GridlineColor =10921638
                    OnClickEmMacro = Begin
                        Version =196611
                        ColumnsShown =0
                        Begin
                            Action ="Close"
                            Argument ="2"
                            Argument ="RekisteroiMaksu"
                            Argument ="2"
                        End
                        Begin
                            Comment ="_AXL:<?xml version=\"1.0\" encoding=\"UTF-16\" standalone=\"no\"?>\015\012<UserI"
                                "nterfaceMacro For=\"Cancel\" Event=\"OnClick\" xmlns=\"http://schemas.microsoft."
                                "com/office/accessservices/2009/11/application\"><Statements><Action Name=\"Close"
                                "Window\"><Argument Name=\"ObjectTy"
                        End
                        Begin
                            Comment ="_AXL:pe\">Form</Argument><Argument Name=\"ObjectName\">RekisteroiMaksu</Argument"
                                "><Argument Name=\"Save\">No</Argument></Action></Statements></UserInterfaceMacro"
                                ">"
                        End
                    End

                    LayoutCachedLeft =6803
                    LayoutCachedTop =5669
                    LayoutCachedWidth =11678
                    LayoutCachedHeight =6916
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
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =215
                    TextFontCharSet =177
                    TextFontFamily =0
                    Left =566
                    Top =5159
                    Width =10545
                    Height =396
                    BackColor =62207
                    BorderColor =8355711
                    ForeColor =6710886
                    Name ="Selite56"
                    Caption ="Paina enter, tab tai klikkaa toista kenttää jos et meinaa päästä eteenpäin tieto"
                        "jen syöttämisen jälkeen!"
                    FontName ="Calibri"
                    GridlineColor =10921638
                    LayoutCachedLeft =566
                    LayoutCachedTop =5159
                    LayoutCachedWidth =11111
                    LayoutCachedHeight =5555
                    BackThemeColorIndex =-1
                End
                Begin Label
                    OverlapFlags =215
                    Left =396
                    Top =2381
                    Width =1104
                    Height =972
                    FontSize =8
                    BorderColor =8355711
                    ForeColor =6710886
                    Name ="Selite60"
                    Caption ="Voit muokata maksutapoja admin moden kautta!"
                    FontName ="Calibri"
                    GridlineColor =10921638
                    LayoutCachedLeft =396
                    LayoutCachedTop =2381
                    LayoutCachedWidth =1500
                    LayoutCachedHeight =3353
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
    [Form_RekisteroiMaksu].Maksutapa.Visible = False
    [Form_RekisteroiMaksu].PVM.Visible = False
    [Form_RekisteroiMaksu].Puumerkki.Visible = False
    [Form_RekisteroiMaksu].Tallenna.Visible = False
End Sub


Private Sub Maksutapa_Click()
    [Form_RekisteroiMaksu].PVM.Visible = True
    [Form_RekisteroiMaksu].Puumerkki.Visible = True

End Sub


Private Sub Puumerkki_Change()
    [Form_RekisteroiMaksu].Tallenna.Visible = True
End Sub



Private Sub Summa_Change()
    [Form_RekisteroiMaksu].Maksutapa.Visible = True
End Sub




Private Sub Tallenna_Click()
    Dim card As String
    Dim payment As Currency
    Dim method As String
    Dim dateStamp As String
    Dim Puumerkki As String
        
    
    If IsNull([Form_Tervetuloa].Korttivalinta) Then
        MsgBox ("Korttinumero ei voi olla tyhjä, valitse kortti pääikkunasta!")
        Exit Sub
    Else
        card = [Form_RekisteroiMaksu].Kortti.Value
    End If
    
    
    If IsNull([Form_RekisteroiMaksu].Summa) Or ([Form_RekisteroiMaksu].Summa.Value = "") Then
        MsgBox ("Summa ei voi olla tyhjä!")
        Exit Sub
    Else
        payment = [Form_RekisteroiMaksu].Summa.Value
    End If
    
    If IsNull([Form_RekisteroiMaksu].Maksutapa) Or ([Form_RekisteroiMaksu].Maksutapa.Value = "") Then
        MsgBox ("Maksutapa ei voi olla tyhjä! Valitse joku maksutapa")
        Exit Sub
    Else
        method = [Form_RekisteroiMaksu].Maksutapa.Value
    End If
    
    dateStamp = [Form_RekisteroiMaksu].PVM.Value

    If IsNull([Form_RekisteroiMaksu].Puumerkki) Or ([Form_RekisteroiMaksu].Puumerkki.Value = "") Then
        MsgBox ("Puumerkki ei voi olla tyhjä!")
        Exit Sub
    Else
        Puumerkki = [Form_RekisteroiMaksu].Puumerkki.Value
    End If
    
    'Form query
    Dim querystring As String
    
    Dim cardID As Integer
    
    cardID = Common.FetchCardID(card)
    
    querystring = " Puumerkki = '" & Puumerkki & "' , " _
    & " Kortti = " & cardID & " , " _
    & " Summa = " & payment & " , " _
    & " Maksutapa = '" & method & "' , " _
    & " PVM = '" & dateStamp & "' "
    
    Dim success As Boolean
    success = Common.InsertOrUpdate("Maksut", querystring, "")
    
    Common.SaveToLog (Puumerkki & " päivitti maksun kortille " & card & ", maksutapa: " & method & " ja summa: " & payment)
    
    
    Dim retval
    retval = Common.SendMessageToMainScreen("Maksu kortille " & card & " rekisteröity!")
    DoCmd.Close
    

    
End Sub
