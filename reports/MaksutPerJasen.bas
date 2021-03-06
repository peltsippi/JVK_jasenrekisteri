Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    PopUp = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DateGrouping =1
    GrpKeepTogether =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =11520
    DatasheetFontHeight =11
    ItemSuffix =17
    Left =2556
    Top =3468
    RecSrcDt = Begin
        0x6afb30f83cd6e540
    End
    RecordSource ="ListaaMaksut"
    Caption ="Maksut per jäsen"
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
            GroupHeader = NotDefault
            ControlSource ="Sukunimi"
        End
        Begin BreakLevel
            GroupHeader = NotDefault
            ControlSource ="Maksutapa"
        End
        Begin BreakLevel
            GroupHeader = NotDefault
            GroupOn =6
            ControlSource ="PVM"
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
            BackThemeColorIndex =2
            BackTint =20.0
            Begin
                Begin Label
                    Left =57
                    Top =57
                    Width =4428
                    Height =528
                    FontSize =20
                    BorderColor =8355711
                    ForeColor =6710886
                    Name ="Selite12"
                    Caption ="Maksut jäsenittäin"
                    FontName ="Calibri"
                    GridlineColor =10921638
                    LayoutCachedLeft =57
                    LayoutCachedTop =57
                    LayoutCachedWidth =4485
                    LayoutCachedHeight =585
                End
                Begin Image
                    PictureType =2
                    Left =4818
                    Top =113
                    Width =735
                    Height =737
                    BorderColor =10921638
                    Name ="Kuva16"
                    Picture ="punttilogo_pieni_invert"
                    GridlineColor =10921638

                    LayoutCachedLeft =4818
                    LayoutCachedTop =113
                    LayoutCachedWidth =5553
                    LayoutCachedHeight =850
                End
            End
        End
        Begin PageHeader
            Visible = NotDefault
            Height =414
            Name ="SivunYlätunniste"
            AutoHeight =1
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin Label
                    TextAlign =1
                    Left =342
                    Top =57
                    Width =2052
                    Height =300
                    BorderColor =8355711
                    ForeColor =6710886
                    Name ="Sukunimi_Selite"
                    Caption ="Sukunimi"
                    FontName ="Calibri"
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    GridlineColor =10921638
                    LayoutCachedLeft =342
                    LayoutCachedTop =57
                    LayoutCachedWidth =2394
                    LayoutCachedHeight =357
                End
                Begin Label
                    TextAlign =1
                    Left =2451
                    Top =57
                    Width =2052
                    Height =300
                    BorderColor =8355711
                    ForeColor =6710886
                    Name ="Etunimi_Selite"
                    Caption ="Etunimi"
                    FontName ="Calibri"
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    GridlineColor =10921638
                    LayoutCachedLeft =2451
                    LayoutCachedTop =57
                    LayoutCachedWidth =4503
                    LayoutCachedHeight =357
                End
                Begin Label
                    TextAlign =1
                    Left =4845
                    Top =57
                    Width =2052
                    Height =300
                    BorderColor =8355711
                    ForeColor =6710886
                    Name ="Maksutapa_Selite"
                    Caption ="Maksutapa"
                    FontName ="Calibri"
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    GridlineColor =10921638
                    LayoutCachedLeft =4845
                    LayoutCachedTop =57
                    LayoutCachedWidth =6897
                    LayoutCachedHeight =357
                End
                Begin Label
                    TextAlign =3
                    Left =8673
                    Top =57
                    Width =1872
                    Height =300
                    BorderColor =8355711
                    ForeColor =6710886
                    Name ="PVM_Selite"
                    Caption ="PVM"
                    FontName ="Calibri"
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    GridlineColor =10921638
                    LayoutCachedLeft =8673
                    LayoutCachedTop =57
                    LayoutCachedWidth =10545
                    LayoutCachedHeight =357
                End
                Begin Label
                    TextAlign =3
                    Left =10602
                    Top =57
                    Width =861
                    Height =300
                    BorderColor =8355711
                    ForeColor =6710886
                    Name ="Summa_Selite"
                    Caption ="Summa"
                    FontName ="Calibri"
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    GridlineColor =10921638
                    LayoutCachedLeft =10602
                    LayoutCachedTop =57
                    LayoutCachedWidth =11463
                    LayoutCachedHeight =357
                End
            End
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            Height =369
            BackColor =-2147483611
            Name ="RyhmänYlätunniste0"
            AutoHeight =1
            AlternateBackColor =-2147483611
            Begin
                Begin TextBox
                    OldBorderStyle =0
                    Left =342
                    Width =2052
                    Height =312
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Sukunimi"
                    ControlSource ="Sukunimi"
                    FontName ="Calibri"
                    GridlineColor =10921638

                    LayoutCachedLeft =342
                    LayoutCachedWidth =2394
                    LayoutCachedHeight =312
                End
                Begin TextBox
                    OldBorderStyle =0
                    Left =2451
                    Width =2052
                    Height =312
                    TabIndex =1
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Etunimi"
                    ControlSource ="Etunimi"
                    FontName ="Calibri"
                    GridlineColor =10921638

                    LayoutCachedLeft =2451
                    LayoutCachedWidth =4503
                    LayoutCachedHeight =312
                End
            End
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            Visible = NotDefault
            Height =369
            BreakLevel =1
            Name ="RyhmänYlätunniste1"
            AutoHeight =1
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            Visible = NotDefault
            Height =369
            BreakLevel =2
            Name ="RyhmänYlätunniste2"
            AutoHeight =1
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
        End
        Begin Section
            Height =369
            BackColor =-2147483610
            Name ="Tiedot"
            AutoHeight =1
            AlternateBackColor =-2147483610
            Begin
                Begin TextBox
                    OldBorderStyle =0
                    IMESentenceMode =3
                    Left =8673
                    Width =1872
                    Height =312
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="PVM"
                    ControlSource ="PVM"
                    StatusBarText ="Ajankohta"
                    FontName ="Calibri"
                    GridlineColor =10921638
                    ShowDatePicker =1

                    LayoutCachedLeft =8673
                    LayoutCachedWidth =10545
                    LayoutCachedHeight =312
                End
                Begin TextBox
                    OldBorderStyle =0
                    IMESentenceMode =3
                    Left =10602
                    Width =861
                    Height =312
                    TabIndex =1
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Summa"
                    ControlSource ="Summa"
                    Format ="#,##0.00 €;-#,##0.00 €"
                    StatusBarText ="Paljoin maksettu"
                    FontName ="Calibri"
                    GridlineColor =10921638

                    LayoutCachedLeft =10602
                    LayoutCachedWidth =11463
                    LayoutCachedHeight =312
                    CurrencySymbol ="€"
                    ColLCID =1035
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    IMESentenceMode =3
                    Left =4988
                    Width =2052
                    Height =312
                    TabIndex =2
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Maksutapa"
                    ControlSource ="Maksutapa"
                    FontName ="Calibri"
                    GridlineColor =10921638

                    LayoutCachedLeft =4988
                    LayoutCachedWidth =7040
                    LayoutCachedHeight =312
                End
            End
        End
        Begin PageFooter
            Height =540
            Name ="SivunAlatunniste"
            AutoHeight =1
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
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
                    Name ="Muokkaus13"
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
                    Left =6423
                    Top =228
                    Width =5040
                    Height =312
                    TabIndex =1
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Muokkaus14"
                    ControlSource ="=\"Sivu \" & [Page] & \"/\" & [Pages]"
                    FontName ="Calibri"
                    GridlineColor =10921638

                    LayoutCachedLeft =6423
                    LayoutCachedTop =228
                    LayoutCachedWidth =11463
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
