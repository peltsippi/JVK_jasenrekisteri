Version =20
VersionRequired =20
Begin Form
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    OrderByOn = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =2
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =9225
    DatasheetFontHeight =11
    ItemSuffix =10
    Left =264
    Top =876
    Right =14952
    Bottom =5076
    OrderBy ="[Lataukset].[Ajankohta]"
    RecSrcDt = Begin
        0xcf4c6b05d1c6e540
    End
    RecordSource ="Lataukset"
    Caption ="LatausListaus"
    DatasheetFontName ="Calibri"
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
        Begin FormHeader
            Height =0
            Name ="LomakkeenYlätunniste"
            AutoHeight =1
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =2
            BackTint =20.0
        End
        Begin Section
            Height =2934
            Name ="Tiedot"
            AutoHeight =1
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2640
                    Top =342
                    Width =1452
                    Height =312
                    ColumnWidth =1452
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Ajankohta"
                    ControlSource ="Ajankohta"
                    StatusBarText ="Miloin tehty"
                    FontName ="Calibri"
                    GridlineColor =10921638

                    LayoutCachedLeft =2640
                    LayoutCachedTop =342
                    LayoutCachedWidth =4092
                    LayoutCachedHeight =654
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =342
                            Top =342
                            Width =2208
                            Height =312
                            BorderColor =8355711
                            ForeColor =6710886
                            Name ="Ajankohta_Selite"
                            Caption ="Ajankohta"
                            FontName ="Calibri"
                            GridlineColor =10921638
                            LayoutCachedLeft =342
                            LayoutCachedTop =342
                            LayoutCachedWidth =2550
                            LayoutCachedHeight =654
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2640
                    Top =741
                    Width =6528
                    Height =576
                    ColumnWidth =3000
                    TabIndex =1
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Lataaja"
                    ControlSource ="Lataaja"
                    StatusBarText ="Lataajan puumerkit"
                    FontName ="Calibri"
                    GridlineColor =10921638

                    LayoutCachedLeft =2640
                    LayoutCachedTop =741
                    LayoutCachedWidth =9168
                    LayoutCachedHeight =1317
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =342
                            Top =741
                            Width =2208
                            Height =312
                            BorderColor =8355711
                            ForeColor =6710886
                            Name ="Lataaja_Selite"
                            Caption ="Lataaja"
                            FontName ="Calibri"
                            GridlineColor =10921638
                            LayoutCachedLeft =342
                            LayoutCachedTop =741
                            LayoutCachedWidth =2550
                            LayoutCachedHeight =1053
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2640
                    Top =1425
                    Width =6528
                    Height =576
                    ColumnWidth =3000
                    TabIndex =2
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Korttityyppi"
                    ControlSource ="Korttityyppi"
                    StatusBarText ="Mikä korttityyppi ladattiin"
                    FontName ="Calibri"
                    GridlineColor =10921638

                    LayoutCachedLeft =2640
                    LayoutCachedTop =1425
                    LayoutCachedWidth =9168
                    LayoutCachedHeight =2001
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =342
                            Top =1425
                            Width =2208
                            Height =312
                            BorderColor =8355711
                            ForeColor =6710886
                            Name ="Korttityyppi_Selite"
                            Caption ="Korttityyppi"
                            FontName ="Calibri"
                            GridlineColor =10921638
                            LayoutCachedLeft =342
                            LayoutCachedTop =1425
                            LayoutCachedWidth =2550
                            LayoutCachedHeight =1737
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2640
                    Top =2109
                    Width =1452
                    Height =312
                    ColumnWidth =1452
                    TabIndex =3
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Voimassa"
                    ControlSource ="Voimassa"
                    StatusBarText ="Kortin voimassaoloaika"
                    FontName ="Calibri"
                    GridlineColor =10921638

                    LayoutCachedLeft =2640
                    LayoutCachedTop =2109
                    LayoutCachedWidth =4092
                    LayoutCachedHeight =2421
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =342
                            Top =2109
                            Width =2208
                            Height =312
                            BorderColor =8355711
                            ForeColor =6710886
                            Name ="Voimassa_Selite"
                            Caption ="Voimassa"
                            FontName ="Calibri"
                            GridlineColor =10921638
                            LayoutCachedLeft =342
                            LayoutCachedTop =2109
                            LayoutCachedWidth =2550
                            LayoutCachedHeight =2421
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2640
                    Top =2508
                    Width =3288
                    Height =312
                    ColumnWidth =3000
                    TabIndex =4
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="KortinArvo"
                    ControlSource ="KortinArvo"
                    Format ="#,##0.00 €;-#,##0.00 €"
                    StatusBarText ="Minkä arvoinen korttilataus on tehty"
                    FontName ="Calibri"
                    GridlineColor =10921638

                    LayoutCachedLeft =2640
                    LayoutCachedTop =2508
                    LayoutCachedWidth =5928
                    LayoutCachedHeight =2820
                    CurrencySymbol ="€"
                    ColLCID =1035
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =342
                            Top =2508
                            Width =2208
                            Height =312
                            BorderColor =8355711
                            ForeColor =6710886
                            Name ="KortinArvo_Selite"
                            Caption ="KortinArvo"
                            FontName ="Calibri"
                            GridlineColor =10921638
                            LayoutCachedLeft =342
                            LayoutCachedTop =2508
                            LayoutCachedWidth =2550
                            LayoutCachedHeight =2820
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
