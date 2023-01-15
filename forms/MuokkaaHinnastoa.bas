Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    Modal = NotDefault
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    CloseButton = NotDefault
    DividingLines = NotDefault
    OrderByOn = NotDefault
    AllowDesignChanges = NotDefault
    ViewsAllowed =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    PictureType =2
    GridY =10
    Width =10602
    DatasheetFontHeight =11
    ItemSuffix =15
    Left =4740
    Top =3468
    Right =18432
    Bottom =11712
    OrderBy ="[Hinnasto].[Tyyppi], [Hinnasto].[Aika] DESC"
    Picture ="bulldog_pienempi"
    RecSrcDt = Begin
        0xa085a184d2c6e540
    End
    RecordSource ="Hinnasto"
    Caption ="MuokkaaHinnastoa"
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
            Height =2325
            Name ="LomakkeenYlätunniste"
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
                    Left =1927
                    Top =1927
                    Width =1416
                    Height =396
                    FontSize =16
                    BackColor =15921906
                    BorderColor =8355711
                    ForeColor =6710886
                    Name ="Tyyppi_Selite"
                    Caption ="Korttityypi"
                    FontName ="Calibri"
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    GridlineColor =10921638
                    LayoutCachedLeft =1927
                    LayoutCachedTop =1927
                    LayoutCachedWidth =3343
                    LayoutCachedHeight =2323
                    BackShade =95.0
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =85
                    TextFontCharSet =177
                    TextAlign =3
                    TextFontFamily =0
                    Left =4422
                    Top =1927
                    Width =1536
                    Height =396
                    FontSize =16
                    BorderColor =8355711
                    ForeColor =6710886
                    Name ="Hinta_Selite"
                    Caption ="Hinta"
                    FontName ="Calibri"
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    GridlineColor =10921638
                    LayoutCachedLeft =4422
                    LayoutCachedTop =1927
                    LayoutCachedWidth =5958
                    LayoutCachedHeight =2323
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =85
                    TextFontCharSet =177
                    TextFontFamily =0
                    Left =170
                    Top =56
                    Width =3283
                    Height =513
                    FontSize =20
                    BackColor =15921906
                    BorderColor =8355711
                    ForeColor =6710886
                    Name ="Selite4"
                    Caption ="Muokkaa hinnastoa"
                    FontName ="Calibri"
                    GridlineColor =10921638
                    LayoutCachedLeft =170
                    LayoutCachedTop =56
                    LayoutCachedWidth =3453
                    LayoutCachedHeight =569
                    BackShade =95.0
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =93
                    TextFontCharSet =177
                    TextAlign =3
                    TextFontFamily =0
                    BackStyle =0
                    IMESentenceMode =3
                    Left =6861
                    Width =3411
                    Height =300
                    BorderColor =10921638
                    Name ="Automaattinen_päivämäärä"
                    ControlSource ="=Date()"
                    Format ="Short Date"
                    FontName ="Calibri"
                    GridlineColor =10921638

                    LayoutCachedLeft =6861
                    LayoutCachedWidth =10272
                    LayoutCachedHeight =300
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =95
                    TextFontCharSet =177
                    TextAlign =3
                    TextFontFamily =0
                    BackStyle =0
                    IMESentenceMode =3
                    Left =5169
                    Top =300
                    Width =5103
                    Height =300
                    TabIndex =1
                    BorderColor =10921638
                    Name ="Automaattinen_aika"
                    ControlSource ="=Time()"
                    Format ="Long Time"
                    FontName ="Calibri"
                    GridlineColor =10921638

                    LayoutCachedLeft =5169
                    LayoutCachedTop =300
                    LayoutCachedWidth =10272
                    LayoutCachedHeight =600
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                End
                Begin CommandButton
                    OverlapFlags =247
                    TextFontCharSet =177
                    TextFontFamily =0
                    Left =7143
                    Top =398
                    Width =3231
                    Height =1132
                    TabIndex =2
                    ForeColor =4210752
                    Name ="Sulje"
                    Caption ="Sulje"
                    OnClick ="[Event Procedure]"
                    FontName ="Calibri"
                    GridlineColor =10921638

                    LayoutCachedLeft =7143
                    LayoutCachedTop =398
                    LayoutCachedWidth =10374
                    LayoutCachedHeight =1530
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
                    OverlapFlags =85
                    TextFontCharSet =177
                    TextFontFamily =0
                    Left =226
                    Top =680
                    Width =6360
                    Height =1224
                    FontSize =16
                    FontWeight =700
                    BackColor =62207
                    BorderColor =8355711
                    ForeColor =255
                    Name ="Selite9"
                    Caption ="HUOM! Muutokset tallentuu aina välittömästi!\015\012Olethan varovainen! Korttity"
                        "ypin avulla voit lisätä tai poistaa korttityyppejä hinnastosta!"
                    FontName ="Calibri"
                    GridlineColor =10921638
                    LayoutCachedLeft =226
                    LayoutCachedTop =680
                    LayoutCachedWidth =6586
                    LayoutCachedHeight =1904
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =85
                    TextFontCharSet =177
                    TextAlign =1
                    TextFontFamily =0
                    Left =170
                    Top =1927
                    Width =1596
                    Height =396
                    FontSize =16
                    BackColor =15921906
                    BorderColor =8355711
                    ForeColor =6710886
                    Name ="Selite10"
                    Caption ="Kuukautta"
                    FontName ="Calibri"
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    GridlineColor =10921638
                    LayoutCachedLeft =170
                    LayoutCachedTop =1927
                    LayoutCachedWidth =1766
                    LayoutCachedHeight =2323
                    BackShade =95.0
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontCharSet =177
                    TextFontFamily =0
                    Left =7143
                    Top =1757
                    Width =3231
                    Height =568
                    TabIndex =3
                    ForeColor =4210752
                    Name ="JarjestaLista"
                    Caption ="Järjestä lista"
                    OnClick ="[Event Procedure]"
                    FontName ="Calibri"
                    GridlineColor =10921638

                    LayoutCachedLeft =7143
                    LayoutCachedTop =1757
                    LayoutCachedWidth =10374
                    LayoutCachedHeight =2325
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
            End
        End
        Begin Section
            Height =440
            Name ="Tiedot"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin Rectangle
                    SpecialEffect =0
                    BackStyle =1
                    BorderWidth =2
                    OverlapFlags =93
                    Left =113
                    Width =10196
                    Height =440
                    Name ="Ruutu12"
                    GridlineColor =10921638
                    LayoutCachedLeft =113
                    LayoutCachedWidth =10309
                    LayoutCachedHeight =440
                    BorderThemeColorIndex =0
                    BorderShade =100.0
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =215
                    TextFontCharSet =177
                    TextFontFamily =0
                    IMESentenceMode =3
                    Left =1984
                    Width =1584
                    Height =396
                    ColumnWidth =3000
                    FontSize =16
                    TabIndex =1
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Tyyppi"
                    ControlSource ="Tyyppi"
                    FontName ="Calibri"
                    GridlineColor =10921638

                    LayoutCachedLeft =1984
                    LayoutCachedWidth =3568
                    LayoutCachedHeight =396
                End
                Begin TextBox
                    OverlapFlags =215
                    TextFontCharSet =177
                    TextFontFamily =0
                    IMESentenceMode =3
                    Left =4478
                    Width =2664
                    Height =396
                    ColumnWidth =3000
                    FontSize =16
                    TabIndex =2
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Hinta"
                    ControlSource ="Hinta"
                    Format ="#,##0.00 €;-#,##0.00 €"
                    FontName ="Calibri"
                    GridlineColor =10921638

                    LayoutCachedLeft =4478
                    LayoutCachedWidth =7142
                    LayoutCachedHeight =396
                    CurrencySymbol ="€"
                    ColLCID =1035
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =215
                    TextFontCharSet =177
                    TextFontFamily =0
                    IMESentenceMode =3
                    Left =113
                    Width =960
                    Height =396
                    FontSize =16
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Muokkaus11"
                    ControlSource ="Aika"
                    FontName ="Calibri"
                    GridlineColor =10921638

                    LayoutCachedLeft =113
                    LayoutCachedWidth =1073
                    LayoutCachedHeight =396
                End
                Begin CommandButton
                    OverlapFlags =215
                    TextFontCharSet =177
                    TextFontFamily =0
                    Left =7426
                    Top =56
                    Width =2891
                    Height =340
                    TabIndex =3
                    ForeColor =4210752
                    Name ="PoistaRivi"
                    Caption ="Poista"
                    OnClick ="[Event Procedure]"
                    FontName ="Calibri"
                    GridlineColor =10921638

                    LayoutCachedLeft =7426
                    LayoutCachedTop =56
                    LayoutCachedWidth =10317
                    LayoutCachedHeight =396
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

Private Sub JarjestaLista_Click()
    Me.Requery
End Sub

Private Sub PoistaRivi_Click()
    DoCmd.RunCommand acCmdSelectRecord
    DoCmd.RunCommand acCmdDeleteRecord
End Sub

Private Sub SortList_Click()
    'DoCmd.SetOrderBy ("{Hinnasto].[Tyyppi], [Hinnasto].[Aika] DESC")
    Me.Requery
    Me.Refresh
    Me.Repaint
    MsgBox ("en helvetti tiiä")
End Sub

Private Sub Sulje_Click()
    Dim succs
    succs = Common.SaveToLog("Hinnaston muokkaaminen päättynyt")
    succs = Common.SendMessageToMainScreen("Hinnaston muokkaaminen valmis")
    DoCmd.Close
    
End Sub
