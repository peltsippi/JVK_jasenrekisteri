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
    ItemSuffix =15
    Left =4740
    Top =3468
    RecSrcDt = Begin
        0xc80ec84976f8e540
    End
    RecordSource ="LatausLoppumassaAjanjaksolla"
    Caption ="Kortit"
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
            ControlSource ="Kortti"
        End
        Begin FormHeader
            KeepTogether = NotDefault
            Height =1020
            Name ="RaportinYlätunniste"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =2
            BackTint =20.0
            Begin
                Begin Label
                    Left =57
                    Top =57
                    Width =7164
                    Height =528
                    FontSize =20
                    BorderColor =8355711
                    ForeColor =6710886
                    Name ="Selite8"
                    Caption ="Kortit, joissa lataus loppumassa ajanjaksolla"
                    FontName ="Calibri"
                    GridlineColor =10921638
                    LayoutCachedLeft =57
                    LayoutCachedTop =57
                    LayoutCachedWidth =7221
                    LayoutCachedHeight =585
                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =9354
                    Top =113
                    Height =300
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Muokkaus5"
                    ControlSource ="=[Forms]![Tervetuloa]![RaportitAlku]"
                    FontName ="Calibri"
                    GridlineColor =10921638

                    LayoutCachedLeft =9354
                    LayoutCachedTop =113
                    LayoutCachedWidth =11055
                    LayoutCachedHeight =413
                    Begin
                        Begin Label
                            Left =8107
                            Top =113
                            Width =1176
                            Height =336
                            BorderColor =8355711
                            ForeColor =6710886
                            Name ="Selite6"
                            Caption ="Alkaen"
                            FontName ="Calibri"
                            GridlineColor =10921638
                            LayoutCachedLeft =8107
                            LayoutCachedTop =113
                            LayoutCachedWidth =9283
                            LayoutCachedHeight =449
                        End
                    End
                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =9354
                    Top =569
                    Height =300
                    TabIndex =1
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Muokkaus7"
                    ControlSource ="=[Forms]![Tervetuloa]![RaportitLoppu]"
                    FontName ="Calibri"
                    GridlineColor =10921638

                    LayoutCachedLeft =9354
                    LayoutCachedTop =569
                    LayoutCachedWidth =11055
                    LayoutCachedHeight =869
                    Begin
                        Begin Label
                            Left =8107
                            Top =569
                            Width =1176
                            Height =336
                            BorderColor =8355711
                            ForeColor =6710886
                            Name ="Selite11"
                            Caption ="Päättyen"
                            FontName ="Calibri"
                            GridlineColor =10921638
                            LayoutCachedLeft =8107
                            LayoutCachedTop =569
                            LayoutCachedWidth =9283
                            LayoutCachedHeight =905
                        End
                    End
                End
                Begin CommandButton
                    Left =170
                    Top =623
                    Width =7483
                    Height =397
                    TabIndex =2
                    ForeColor =4210752
                    Name ="HideNames"
                    Caption ="Piilota nimet"
                    OnClick ="[Event Procedure]"
                    FontName ="Calibri"
                    GridlineColor =10921638

                    LayoutCachedLeft =170
                    LayoutCachedTop =623
                    LayoutCachedWidth =7653
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
                    Overlaps =1
                End
            End
        End
        Begin PageHeader
            Height =510
            Name ="SivunYlätunniste"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin Label
                    TextAlign =1
                    Left =342
                    Top =57
                    Width =1986
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
                    LayoutCachedWidth =2328
                    LayoutCachedHeight =357
                End
                Begin Label
                    TextAlign =1
                    Left =6292
                    Top =56
                    Width =741
                    Height =300
                    BorderColor =8355711
                    ForeColor =6710886
                    Name ="Sukunimi_Selite"
                    Caption ="Nimi"
                    FontName ="Calibri"
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    GridlineColor =10921638
                    LayoutCachedLeft =6292
                    LayoutCachedTop =56
                    LayoutCachedWidth =7033
                    LayoutCachedHeight =356
                End
                Begin Label
                    TextAlign =1
                    Left =2437
                    Top =56
                    Width =1530
                    Height =300
                    BorderColor =8355711
                    ForeColor =6710886
                    Name ="Selite13"
                    Caption ="Voimassa asti"
                    FontName ="Calibri"
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    GridlineColor =10921638
                    LayoutCachedLeft =2437
                    LayoutCachedTop =56
                    LayoutCachedWidth =3967
                    LayoutCachedHeight =356
                End
            End
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            Height =0
            Name ="RyhmänYlätunniste0"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
        End
        Begin Section
            KeepTogether = NotDefault
            Height =312
            Name ="Tiedot"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin TextBox
                    OldBorderStyle =0
                    IMESentenceMode =3
                    Left =170
                    Width =2094
                    Height =312
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Kortti"
                    ControlSource ="Kortti"
                    StatusBarText ="Kortin numero niin kuin se on kirjoitettu, esim 0285"
                    FontName ="Calibri"
                    GridlineColor =10921638

                    LayoutCachedLeft =170
                    LayoutCachedWidth =2264
                    LayoutCachedHeight =312
                End
                Begin TextBox
                    OldBorderStyle =0
                    Left =6236
                    Width =5061
                    Height =312
                    TabIndex =1
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Nimi"
                    ControlSource ="Nimi"
                    FontName ="Calibri"
                    GridlineColor =10921638

                    LayoutCachedLeft =6236
                    LayoutCachedWidth =11297
                    LayoutCachedHeight =312
                End
                Begin TextBox
                    OldBorderStyle =0
                    IMESentenceMode =3
                    Left =2324
                    Width =1758
                    Height =312
                    TabIndex =2
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Voimassaolo"
                    ControlSource ="Voimassaolo"
                    StatusBarText ="Kortin numero niin kuin se on kirjoitettu, esim 0285"
                    FontName ="Calibri"
                    GridlineColor =10921638

                    LayoutCachedLeft =2324
                    LayoutCachedWidth =4082
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
                    Name ="Muokkaus9"
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
                    Name ="Muokkaus10"
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
CodeBehindForm
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub HideNames_Click()
Dim previousstate As Boolean
previousstate = [Report_LatausLoppumassaAjanjaksolla].Nimi.Visible
previousstate = Not previousstate
'MsgBox ("Previous status: " & previousstate)
[Report_LatausLoppumassaAjanjaksolla].Nimi.Visible = previousstate
End Sub
