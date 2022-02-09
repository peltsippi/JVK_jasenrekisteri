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
    ItemSuffix =10
    Left =4044
    Top =3456
    Right =17796
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
            Height =1984
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
                    Left =2834
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
                    LayoutCachedLeft =2834
                    LayoutCachedTop =1587
                    LayoutCachedWidth =6122
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
                    OverlapFlags =87
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
                    OverlapFlags =85
                    Left =7143
                    Top =1190
                    Width =3231
                    Height =340
                    TabIndex =2
                    ForeColor =4210752
                    Name ="Sulje"
                    Caption ="Sulje"
                    OnClick ="[Event Procedure]"
                    FontName ="Calibri"
                    GridlineColor =10921638

                    LayoutCachedLeft =7143
                    LayoutCachedTop =1190
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
            End
        End
        Begin Section
            Height =690
            Name ="Tiedot"
            AutoHeight =1
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =342
                    Top =57
                    Width =2316
                    Height =576
                    ColumnWidth =3000
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Tyyppi"
                    ControlSource ="Tyyppi"
                    FontName ="Calibri"
                    GridlineColor =10921638

                    LayoutCachedLeft =342
                    LayoutCachedTop =57
                    LayoutCachedWidth =2658
                    LayoutCachedHeight =633
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2721
                    Top =56
                    Width =3288
                    Height =564
                    ColumnWidth =3000
                    TabIndex =1
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Hinta"
                    ControlSource ="Hinta"
                    Format ="#,##0.00 €;-#,##0.00 €"
                    FontName ="Calibri"
                    GridlineColor =10921638

                    LayoutCachedLeft =2721
                    LayoutCachedTop =56
                    LayoutCachedWidth =6009
                    LayoutCachedHeight =620
                    CurrencySymbol ="€"
                    ColLCID =1035
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
