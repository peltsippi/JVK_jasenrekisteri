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
    Width =11520
    DatasheetFontHeight =11
    ItemSuffix =13
    Left =4740
    Top =3468
    RecSrcDt = Begin
        0x7c17dc9c5cdae540
    End
    RecordSource ="LataustaAjanjaksolla"
    Caption ="Kortit joissa latausta ajanjaksolla"
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
            ControlSource ="Kortti"
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
                    Width =5352
                    Height =528
                    FontSize =20
                    BorderColor =8355711
                    ForeColor =6710886
                    Name ="Selite2"
                    Caption ="Kortit joissa latausta ajanjaksolla"
                    FontName ="Calibri"
                    GridlineColor =10921638
                    LayoutCachedLeft =57
                    LayoutCachedTop =57
                    LayoutCachedWidth =5409
                    LayoutCachedHeight =585
                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =7313
                    Top =113
                    Height =300
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Muokkaus5"
                    ControlSource ="=[Forms]![Tervetuloa]![RaportitAlku]"
                    FontName ="Calibri"
                    GridlineColor =10921638

                    LayoutCachedLeft =7313
                    LayoutCachedTop =113
                    LayoutCachedWidth =9014
                    LayoutCachedHeight =413
                    Begin
                        Begin Label
                            Left =6066
                            Top =113
                            Width =1176
                            Height =336
                            BorderColor =8355711
                            ForeColor =6710886
                            Name ="Selite6"
                            Caption ="Alkaen"
                            FontName ="Calibri"
                            GridlineColor =10921638
                            LayoutCachedLeft =6066
                            LayoutCachedTop =113
                            LayoutCachedWidth =7242
                            LayoutCachedHeight =449
                        End
                    End
                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =7313
                    Top =569
                    Height =300
                    TabIndex =1
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Muokkaus7"
                    ControlSource ="=[Forms]![Tervetuloa]![RaportitLoppu]"
                    FontName ="Calibri"
                    GridlineColor =10921638

                    LayoutCachedLeft =7313
                    LayoutCachedTop =569
                    LayoutCachedWidth =9014
                    LayoutCachedHeight =869
                    Begin
                        Begin Label
                            Left =6066
                            Top =569
                            Width =1176
                            Height =336
                            BorderColor =8355711
                            ForeColor =6710886
                            Name ="Selite8"
                            Caption ="Päättyen"
                            FontName ="Calibri"
                            GridlineColor =10921638
                            LayoutCachedLeft =6066
                            LayoutCachedTop =569
                            LayoutCachedWidth =7242
                            LayoutCachedHeight =905
                        End
                    End
                End
            End
        End
        Begin PageHeader
            Height =357
            Name ="SivunYlätunniste"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin Label
                    TextAlign =1
                    Left =342
                    Top =57
                    Width =1356
                    Height =300
                    BorderColor =8355711
                    ForeColor =6710886
                    Name ="Kortti_Selite"
                    Caption ="Kortti"
                    FontName ="Calibri"
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    GridlineColor =10921638
                    LayoutCachedLeft =342
                    LayoutCachedTop =57
                    LayoutCachedWidth =1698
                    LayoutCachedHeight =357
                End
                Begin Label
                    OverlapFlags =4
                    TextAlign =1
                    Left =1814
                    Top =56
                    Width =3180
                    Height =300
                    BorderColor =8355711
                    ForeColor =6710886
                    Name ="Selite9"
                    Caption ="Sukunimi"
                    FontName ="Calibri"
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    GridlineColor =10921638
                    LayoutCachedLeft =1814
                    LayoutCachedTop =56
                    LayoutCachedWidth =4994
                    LayoutCachedHeight =356
                End
                Begin Label
                    OverlapFlags =4
                    TextAlign =1
                    Left =5102
                    Top =56
                    Width =3180
                    Height =300
                    BorderColor =8355711
                    ForeColor =6710886
                    Name ="Selite10"
                    Caption ="Etunimi"
                    FontName ="Calibri"
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    GridlineColor =10921638
                    LayoutCachedLeft =5102
                    LayoutCachedTop =56
                    LayoutCachedWidth =8282
                    LayoutCachedHeight =356
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            Height =453
            Name ="Tiedot"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin TextBox
                    OldBorderStyle =0
                    IMESentenceMode =3
                    Left =342
                    Top =57
                    Width =1368
                    Height =312
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Kortti"
                    ControlSource ="Kortti"
                    StatusBarText ="Kortin numero niin kuin se on kirjoitettu, esim 0285"
                    FontName ="Calibri"
                    GridlineColor =10921638

                    LayoutCachedLeft =342
                    LayoutCachedTop =57
                    LayoutCachedWidth =1710
                    LayoutCachedHeight =369
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =4
                    IMESentenceMode =3
                    Left =1814
                    Top =56
                    Width =3240
                    Height =312
                    TabIndex =1
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Muokkaus11"
                    ControlSource ="Sukunimi"
                    StatusBarText ="Kortin numero niin kuin se on kirjoitettu, esim 0285"
                    FontName ="Calibri"
                    GridlineColor =10921638

                    LayoutCachedLeft =1814
                    LayoutCachedTop =56
                    LayoutCachedWidth =5054
                    LayoutCachedHeight =368
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =4
                    IMESentenceMode =3
                    Left =5102
                    Top =56
                    Width =3240
                    Height =312
                    TabIndex =2
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Muokkaus12"
                    ControlSource ="Etunimi"
                    StatusBarText ="Kortin numero niin kuin se on kirjoitettu, esim 0285"
                    FontName ="Calibri"
                    GridlineColor =10921638

                    LayoutCachedLeft =5102
                    LayoutCachedTop =56
                    LayoutCachedWidth =8342
                    LayoutCachedHeight =368
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
                    Name ="Muokkaus3"
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
                    Name ="Muokkaus4"
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
