Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    Modal = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    CloseButton = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    ViewsAllowed =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    PictureType =2
    GridY =10
    Width =14116
    DatasheetFontHeight =11
    ItemSuffix =10
    Left =4044
    Top =3468
    Right =17484
    Bottom =11712
    Picture ="bulldog_pienempi"
    RecSrcDt = Begin
        0xa8077c0f5dc7e540
    End
    RecordSource ="MaksuListaus"
    Caption ="Korjaa kortin maksuja"
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
        Begin FormHeader
            Height =1026
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
                    Left =57
                    Top =57
                    Width =3787
                    Height =501
                    FontSize =20
                    BackColor =15921906
                    BorderColor =8355711
                    ForeColor =6710886
                    Name ="Selite4"
                    Caption ="Maksujen korjaaminen"
                    FontName ="Calibri"
                    GridlineColor =10921638
                    LayoutCachedLeft =57
                    LayoutCachedTop =57
                    LayoutCachedWidth =3844
                    LayoutCachedHeight =558
                    BackShade =95.0
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =93
                    TextAlign =3
                    IMESentenceMode =3
                    Left =12417
                    Width =1695
                    Height =300
                    BackColor =15921906
                    BorderColor =10921638
                    Name ="Automaattinen_päivämäärä"
                    ControlSource ="=Date()"
                    Format ="Short Date"
                    FontName ="Calibri"
                    GridlineColor =10921638

                    LayoutCachedLeft =12417
                    LayoutCachedWidth =14112
                    LayoutCachedHeight =300
                    BackShade =95.0
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =87
                    TextAlign =3
                    IMESentenceMode =3
                    Left =12465
                    Top =300
                    Width =1647
                    Height =300
                    TabIndex =1
                    BackColor =15921906
                    BorderColor =10921638
                    Name ="Automaattinen_aika"
                    ControlSource ="=Time()"
                    Format ="Long Time"
                    FontName ="Calibri"
                    GridlineColor =10921638

                    LayoutCachedLeft =12465
                    LayoutCachedTop =300
                    LayoutCachedWidth =14112
                    LayoutCachedHeight =600
                    BackShade =95.0
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                End
            End
        End
        Begin Section
            Height =5642
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
                    Left =2640
                    Top =342
                    Width =1668
                    Height =576
                    ColumnWidth =3000
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Kortti"
                    ControlSource ="=[Forms]![Tervetuloa]![Korttivalinta]"
                    StatusBarText ="Kortin numero niin kuin se on kirjoitettu, esim 0285"
                    FontName ="Calibri"
                    GridlineColor =10921638

                    LayoutCachedLeft =2640
                    LayoutCachedTop =342
                    LayoutCachedWidth =4308
                    LayoutCachedHeight =918
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =85
                            Left =342
                            Top =342
                            Width =2208
                            Height =312
                            BackColor =15921906
                            BorderColor =8355711
                            ForeColor =6710886
                            Name ="Kortti_Selite"
                            Caption ="Kortti"
                            FontName ="Calibri"
                            GridlineColor =10921638
                            LayoutCachedLeft =342
                            LayoutCachedTop =342
                            LayoutCachedWidth =2550
                            LayoutCachedHeight =654
                            BackShade =95.0
                        End
                    End
                End
                Begin Subform
                    OverlapFlags =215
                    Left =453
                    Top =1424
                    Width =13485
                    Height =4218
                    TabIndex =1
                    BorderColor =10921638
                    Name ="MaksuLuetteo"
                    SourceObject ="Form.MaksuListaus"
                    GridlineColor =10921638
                    FilterOnEmptyMaster =0

                    LayoutCachedLeft =453
                    LayoutCachedTop =1424
                    LayoutCachedWidth =13938
                    LayoutCachedHeight =5642
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =93
                            Left =532
                            Top =1133
                            Width =2208
                            Height =312
                            BackColor =15921906
                            BorderColor =8355711
                            ForeColor =6710886
                            Name ="MaksuLuetteo_Selite"
                            Caption ="Maksulistaus:"
                            FontName ="Calibri"
                            GridlineColor =10921638
                            LayoutCachedLeft =532
                            LayoutCachedTop =1133
                            LayoutCachedWidth =2740
                            LayoutCachedHeight =1445
                            BackShade =95.0
                        End
                    End
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =85
                    Left =6349
                    Top =226
                    Width =7711
                    Height =851
                    FontSize =16
                    FontWeight =700
                    BackColor =62207
                    BorderColor =8355711
                    ForeColor =255
                    Name ="Selite9"
                    Caption ="HUOM! Muutokset tallentuu aina välittömästi!\015\012Olethan varovainen!"
                    FontName ="Calibri"
                    GridlineColor =10921638
                    LayoutCachedLeft =6349
                    LayoutCachedTop =226
                    LayoutCachedWidth =14060
                    LayoutCachedHeight =1077
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =4478
                    Top =340
                    Width =1758
                    Height =680
                    TabIndex =2
                    ForeColor =4210752
                    Name ="Sulje"
                    Caption ="Sulje"
                    OnClick ="[Event Procedure]"
                    FontName ="Calibri"
                    GridlineColor =10921638

                    LayoutCachedLeft =4478
                    LayoutCachedTop =340
                    LayoutCachedWidth =6236
                    LayoutCachedHeight =1020
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
    succs = Common.SaveToLog("Kortin maksujen korjaaminen päättyi")
    succs = Common.SendMessageToMainScreen("Maksujen korjaus valmis")
    DoCmd.Close
    
End Sub
