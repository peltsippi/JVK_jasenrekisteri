Version =20
VersionRequired =20
Begin Form
    NavigationButtons = NotDefault
    CloseButton = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =2
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =9225
    DatasheetFontHeight =11
    ItemSuffix =8
    Right =12744
    Bottom =8244
    RecSrcDt = Begin
        0xc1709afed1c6e540
    End
    RecordSource ="Maksut"
    Caption ="MaksuLuetteo"
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
            Height =2514
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
                    Name ="PVM"
                    ControlSource ="PVM"
                    StatusBarText ="Ajankohta"
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
                            Name ="PVM_Selite"
                            Caption ="PVM"
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
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2640
                    Top =741
                    Width =3288
                    Height =312
                    ColumnWidth =3000
                    TabIndex =1
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Summa"
                    ControlSource ="Summa"
                    Format ="#,##0.00 €;-#,##0.00 €"
                    StatusBarText ="Paljoin maksettu"
                    FontName ="Calibri"
                    GridlineColor =10921638

                    LayoutCachedLeft =2640
                    LayoutCachedTop =741
                    LayoutCachedWidth =5928
                    LayoutCachedHeight =1053
                    CurrencySymbol ="€"
                    ColLCID =1035
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =342
                            Top =741
                            Width =2208
                            Height =312
                            BorderColor =8355711
                            ForeColor =6710886
                            Name ="Summa_Selite"
                            Caption ="Summa"
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
                    Top =1140
                    Width =6528
                    Height =576
                    ColumnWidth =3000
                    TabIndex =2
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Maksutapa"
                    ControlSource ="Maksutapa"
                    FontName ="Calibri"
                    GridlineColor =10921638

                    LayoutCachedLeft =2640
                    LayoutCachedTop =1140
                    LayoutCachedWidth =9168
                    LayoutCachedHeight =1716
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =342
                            Top =1140
                            Width =2208
                            Height =312
                            BorderColor =8355711
                            ForeColor =6710886
                            Name ="Maksutapa_Selite"
                            Caption ="Maksutapa"
                            FontName ="Calibri"
                            GridlineColor =10921638
                            LayoutCachedLeft =342
                            LayoutCachedTop =1140
                            LayoutCachedWidth =2550
                            LayoutCachedHeight =1452
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2640
                    Top =1824
                    Width =6528
                    Height =576
                    ColumnWidth =3000
                    TabIndex =3
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Puumerkki"
                    ControlSource ="Puumerkki"
                    StatusBarText ="Kuittaajan puumerkit"
                    FontName ="Calibri"
                    GridlineColor =10921638

                    LayoutCachedLeft =2640
                    LayoutCachedTop =1824
                    LayoutCachedWidth =9168
                    LayoutCachedHeight =2400
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =342
                            Top =1824
                            Width =2208
                            Height =312
                            BorderColor =8355711
                            ForeColor =6710886
                            Name ="Puumerkki_Selite"
                            Caption ="Puumerkki"
                            FontName ="Calibri"
                            GridlineColor =10921638
                            LayoutCachedLeft =342
                            LayoutCachedTop =1824
                            LayoutCachedWidth =2550
                            LayoutCachedHeight =2136
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
