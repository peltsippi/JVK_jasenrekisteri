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
    AllowDesignChanges = NotDefault
    ViewsAllowed =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    PictureType =2
    GridY =10
    Width =10602
    DatasheetFontHeight =11
    ItemSuffix =13
    Left =4740
    Top =3468
    Right =18432
    Bottom =11712
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
            Height =2307
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
                    TextAlign =1
                    Left =340
                    Top =1587
                    Width =2316
                    Height =300
                    BackColor =15921906
                    BorderColor =8355711
                    ForeColor =6710886
                    Name ="Tyyppi_Selite"
                    Caption ="Tyyppi"
                    FontName ="Calibri"
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    GridlineColor =10921638
                    LayoutCachedLeft =340
                    LayoutCachedTop =1587
                    LayoutCachedWidth =2656
                    LayoutCachedHeight =1887
                    BackShade =95.0
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =85
                    TextAlign =3
                    Left =6009
                    Top =1587
                    Width =3288
                    Height =300
                    BorderColor =8355711
                    ForeColor =6710886
                    Name ="Hinta_Selite"
                    Caption ="Hinta"
                    FontName ="Calibri"
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    GridlineColor =10921638
                    LayoutCachedLeft =6009
                    LayoutCachedTop =1587
                    LayoutCachedWidth =9297
                    LayoutCachedHeight =1887
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =85
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
                    TextAlign =3
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
                    TextAlign =3
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
                    Left =226
                    Top =680
                    Width =6355
                    Height =779
                    FontSize =16
                    FontWeight =700
                    BackColor =62207
                    BorderColor =8355711
                    ForeColor =255
                    Name ="Selite9"
                    Caption ="HUOM! Muutokset tallentuu aina välittömästi!\015\012Olethan varovainen!"
                    FontName ="Calibri"
                    GridlineColor =10921638
                    LayoutCachedLeft =226
                    LayoutCachedTop =680
                    LayoutCachedWidth =6581
                    LayoutCachedHeight =1459
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =85
                    TextAlign =1
                    Left =3231
                    Top =1644
                    Width =2316
                    Height =300
                    BackColor =15921906
                    BorderColor =8355711
                    ForeColor =6710886
                    Name ="Selite10"
                    Caption ="Kesto"
                    FontName ="Calibri"
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    GridlineColor =10921638
                    LayoutCachedLeft =3231
                    LayoutCachedTop =1644
                    LayoutCachedWidth =5547
                    LayoutCachedHeight =1944
                    BackShade =95.0
                End
            End
        End
        Begin Section
            Height =333
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
                    Height =332
                    Name ="Ruutu12"
                    GridlineColor =10921638
                    LayoutCachedLeft =113
                    LayoutCachedWidth =10309
                    LayoutCachedHeight =332
                    BorderThemeColorIndex =0
                    BorderShade =100.0
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =215
                    IMESentenceMode =3
                    Left =342
                    Top =57
                    Width =1584
                    Height =276
                    ColumnWidth =3000
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Tyyppi"
                    ControlSource ="Tyyppi"
                    FontName ="Calibri"
                    GridlineColor =10921638

                    LayoutCachedLeft =342
                    LayoutCachedTop =57
                    LayoutCachedWidth =1926
                    LayoutCachedHeight =333
                End
                Begin TextBox
                    OverlapFlags =215
                    IMESentenceMode =3
                    Left =3685
                    Top =56
                    Width =3288
                    Height =228
                    ColumnWidth =3000
                    TabIndex =1
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Hinta"
                    ControlSource ="Hinta"
                    Format ="#,##0.00 €;-#,##0.00 €"
                    FontName ="Calibri"
                    GridlineColor =10921638

                    LayoutCachedLeft =3685
                    LayoutCachedTop =56
                    LayoutCachedWidth =6973
                    LayoutCachedHeight =284
                    CurrencySymbol ="€"
                    ColLCID =1035
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =215
                    IMESentenceMode =3
                    Left =2494
                    Top =56
                    Width =960
                    Height =276
                    TabIndex =2
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Muokkaus11"
                    ControlSource ="Aika"
                    FontName ="Calibri"
                    GridlineColor =10921638

                    LayoutCachedLeft =2494
                    LayoutCachedTop =56
                    LayoutCachedWidth =3454
                    LayoutCachedHeight =332
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

Private Sub Sulje_Click()
    Dim succs
    succs = Common.SaveToLog("Hinnaston muokkaaminen päättynyt")
    succs = Common.SendMessageToMainScreen("Hinnaston muokkaaminen valmis")
    DoCmd.Close
    
End Sub
