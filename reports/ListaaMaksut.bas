Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    PopUp = NotDefault
    Modal = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DateGrouping =1
    GrpKeepTogether =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =15120
    DatasheetFontHeight =11
    ItemSuffix =16
    Left =4740
    Top =3468
    RecSrcDt = Begin
        0x8f0bb9bd5ec7e540
    End
    RecordSource ="ListaaMaksut"
    Caption ="ListaaMaksut"
    DatasheetFontName ="Calibri"
    FilterOnLoad =0
    FitToPage =1
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
        Begin Image
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
        Begin TextBox
            AddColon = NotDefault
            FELineBreak = NotDefault
            BorderLineStyle =0
            Width =1701
            LabelX =-1701
            FontSize =11
            FontName ="Segoe UI"
            AsianLineBreak =1
            ShowDatePicker =0
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            ThemeFontIndex =1
            ForeThemeColorIndex =0
            ForeTint =75.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin BreakLevel
            ControlSource ="PVM"
        End
        Begin FormHeader
            KeepTogether = NotDefault
            Height =927
            Name ="RaportinYlätunniste"
            AutoHeight =1
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            Begin
                Begin Label
                    Left =1360
                    Top =113
                    Width =4260
                    Height =528
                    FontSize =20
                    BorderColor =8355711
                    ForeColor =6710886
                    Name ="Selite10"
                    Caption ="Listaa maksut ajanjaksolla"
                    FontName ="Calibri"
                    GridlineColor =10921638
                    LayoutCachedLeft =1360
                    LayoutCachedTop =113
                    LayoutCachedWidth =5620
                    LayoutCachedHeight =641
                End
                Begin Image
                    PictureType =2
                    Left =113
                    Top =113
                    Width =1020
                    Height =794
                    BorderColor =10921638
                    Name ="Kuva13"
                    Picture ="punttilogo_pieni_invert"
                    GridlineColor =10921638

                    LayoutCachedLeft =113
                    LayoutCachedTop =113
                    LayoutCachedWidth =1133
                    LayoutCachedHeight =907
                    TabIndex =1
                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =6916
                    Top =113
                    Width =4706
                    Height =576
                    FontWeight =700
                    BackColor =14277081
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Muokkaus14"
                    ControlSource ="=[Forms]![Tervetuloa]![RaportitAlku] & \" - \" & [Forms]![Tervetuloa]![RaportitL"
                        "oppu]"
                    FontName ="Calibri"
                    GridlineColor =10921638

                    LayoutCachedLeft =6916
                    LayoutCachedTop =113
                    LayoutCachedWidth =11622
                    LayoutCachedHeight =689
                    BackShade =85.0
                End
            End
        End
        Begin PageHeader
            Height =414
            BackColor =15527148
            Name ="SivunYlätunniste"
            AutoHeight =1
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            Begin
                Begin Label
                    TextAlign =3
                    Left =342
                    Top =57
                    Width =1197
                    Height =300
                    BorderColor =8355711
                    ForeColor =6710886
                    Name ="PVM_Selite"
                    Caption ="PVM"
                    FontName ="Calibri"
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    GridlineColor =10921638
                    LayoutCachedLeft =342
                    LayoutCachedTop =57
                    LayoutCachedWidth =1539
                    LayoutCachedHeight =357
                End
                Begin Label
                    TextAlign =1
                    Left =1596
                    Top =57
                    Width =5415
                    Height =300
                    BorderColor =8355711
                    ForeColor =6710886
                    Name ="Maksutapa_Selite"
                    Caption ="Maksutapa"
                    FontName ="Calibri"
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    GridlineColor =10921638
                    LayoutCachedLeft =1596
                    LayoutCachedTop =57
                    LayoutCachedWidth =7011
                    LayoutCachedHeight =357
                End
                Begin Label
                    TextAlign =3
                    Left =7068
                    Top =57
                    Width =2736
                    Height =300
                    BorderColor =8355711
                    ForeColor =6710886
                    Name ="Summa_Selite"
                    Caption ="Summa"
                    FontName ="Calibri"
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    GridlineColor =10921638
                    LayoutCachedLeft =7068
                    LayoutCachedTop =57
                    LayoutCachedWidth =9804
                    LayoutCachedHeight =357
                End
                Begin Label
                    TextAlign =1
                    Left =9861
                    Top =57
                    Width =2736
                    Height =300
                    BorderColor =8355711
                    ForeColor =6710886
                    Name ="Etunimi_Selite"
                    Caption ="Etunimi"
                    FontName ="Calibri"
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    GridlineColor =10921638
                    LayoutCachedLeft =9861
                    LayoutCachedTop =57
                    LayoutCachedWidth =12597
                    LayoutCachedHeight =357
                End
                Begin Label
                    TextAlign =1
                    Left =12654
                    Top =57
                    Width =2409
                    Height =300
                    BorderColor =8355711
                    ForeColor =6710886
                    Name ="Sukunimi_Selite"
                    Caption ="Sukunimi"
                    FontName ="Calibri"
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    GridlineColor =10921638
                    LayoutCachedLeft =12654
                    LayoutCachedTop =57
                    LayoutCachedWidth =15063
                    LayoutCachedHeight =357
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            Height =426
            Name ="Tiedot"
            AutoHeight =1
            AlternateBackColor =14211288
            Begin
                Begin TextBox
                    OldBorderStyle =0
                    IMESentenceMode =3
                    Left =342
                    Top =57
                    Width =1197
                    Height =312
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="PVM"
                    ControlSource ="PVM"
                    StatusBarText ="Ajankohta"
                    FontName ="Calibri"
                    GridlineColor =10921638
                    ShowDatePicker =1

                    LayoutCachedLeft =342
                    LayoutCachedTop =57
                    LayoutCachedWidth =1539
                    LayoutCachedHeight =369
                End
                Begin TextBox
                    OldBorderStyle =0
                    IMESentenceMode =3
                    Left =1596
                    Top =57
                    Width =5415
                    Height =312
                    TabIndex =1
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Maksutapa"
                    ControlSource ="Maksutapa"
                    FontName ="Calibri"
                    GridlineColor =10921638

                    LayoutCachedLeft =1596
                    LayoutCachedTop =57
                    LayoutCachedWidth =7011
                    LayoutCachedHeight =369
                End
                Begin TextBox
                    OldBorderStyle =0
                    IMESentenceMode =3
                    Left =7068
                    Top =57
                    Width =2736
                    Height =312
                    TabIndex =2
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Summa"
                    ControlSource ="Summa"
                    Format ="#,##0.00 €;-#,##0.00 €"
                    StatusBarText ="Paljoin maksettu"
                    FontName ="Calibri"
                    GridlineColor =10921638

                    LayoutCachedLeft =7068
                    LayoutCachedTop =57
                    LayoutCachedWidth =9804
                    LayoutCachedHeight =369
                    CurrencySymbol ="€"
                    ColLCID =1035
                End
                Begin TextBox
                    OldBorderStyle =0
                    Left =9861
                    Top =57
                    Width =2736
                    Height =312
                    TabIndex =3
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Etunimi"
                    ControlSource ="Etunimi"
                    FontName ="Calibri"
                    GridlineColor =10921638

                    LayoutCachedLeft =9861
                    LayoutCachedTop =57
                    LayoutCachedWidth =12597
                    LayoutCachedHeight =369
                End
                Begin TextBox
                    OldBorderStyle =0
                    Left =12654
                    Top =57
                    Width =2409
                    Height =312
                    TabIndex =4
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Sukunimi"
                    ControlSource ="Sukunimi"
                    FontName ="Calibri"
                    GridlineColor =10921638

                    LayoutCachedLeft =12654
                    LayoutCachedTop =57
                    LayoutCachedWidth =15063
                    LayoutCachedHeight =369
                End
            End
        End
        Begin PageFooter
            Height =540
            BackColor =15527148
            Name ="SivunAlatunniste"
            AutoHeight =1
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            Begin
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =1
                    IMESentenceMode =3
                    Left =57
                    Top =228
                    Width =5040
                    Height =312
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Muokkaus11"
                    ControlSource ="=Now()"
                    Format ="Long Date"
                    FontName ="Calibri"
                    GridlineColor =10921638

                    LayoutCachedLeft =57
                    LayoutCachedTop =228
                    LayoutCachedWidth =5097
                    LayoutCachedHeight =540
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =10023
                    Top =228
                    Width =5040
                    Height =312
                    TabIndex =1
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Muokkaus12"
                    ControlSource ="=\"Sivu \" & [Page] & \"/\" & [Pages]"
                    FontName ="Calibri"
                    GridlineColor =10921638

                    LayoutCachedLeft =10023
                    LayoutCachedTop =228
                    LayoutCachedWidth =15063
                    LayoutCachedHeight =540
                End
            End
        End
        Begin FormFooter
            KeepTogether = NotDefault
            Height =0
            Name ="RaportinAlatunniste"
            AutoHeight =1
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
        End
    End
End
