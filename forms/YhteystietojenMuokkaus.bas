Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    Modal = NotDefault
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    CloseButton = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    ViewsAllowed =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    PictureType =2
    GridY =10
    Width =9071
    DatasheetFontHeight =11
    ItemSuffix =146
    Left =4044
    Top =3468
    Right =17796
    Bottom =11712
    Filter ="[UID]=[Forms]![Tervetuloa]![Yhteystietovalinta]"
    Picture ="bulldog_pienempi"
    RecSrcDt = Begin
        0xaa33e4155dc6e540
    End
    RecordSource ="Yhteystiedot"
    Caption ="Yhteystietojen muokkaus"
    DatasheetFontName ="Calibri"
    AllowDatasheetView =0
    FilterOnLoad =0
    OnOpenEmMacro = Begin
        Version =196611
        ColumnsShown =0
        Begin
            Action ="ApplyFilter"
            Argument ="Suodatin"
            Argument ="[UID]=[Forms]![Tervetuloa]![Yhteystietovalinta]"
        End
        Begin
            Comment ="_AXL:<?xml version=\"1.0\" encoding=\"UTF-16\" standalone=\"no\"?>\015\012<UserI"
                "nterfaceMacro Event=\"OnOpen\" xmlns=\"http://schemas.microsoft.com/office/acces"
                "sservices/2009/11/application\"><Statements><Action Name=\"ApplyFilter\"><Argume"
                "nt Name=\"FilterName\">Suodatin</"
        End
        Begin
            Comment ="_AXL:Argument><Argument Name=\"WhereCondition\">[UID]=[Forms]![Tervetuloa]![Yhte"
                "ystietovalinta]</Argument></Action></Statements></UserInterfaceMacro>"
        End
    End
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
        Begin CheckBox
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
            BorderThemeColorIndex =1
            BorderShade =65.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin OptionGroup
            SpecialEffect =3
            BorderLineStyle =0
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
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            ThemeFontIndex =1
            ForeThemeColorIndex =0
            ForeTint =75.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin ListBox
            BorderLineStyle =0
            Width =1701
            Height =1417
            LabelX =-1701
            FontSize =11
            FontName ="Segoe UI"
            AllowValueListEdits =1
            InheritValueList =1
            ThemeFontIndex =1
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            ForeThemeColorIndex =0
            ForeTint =75.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin ComboBox
            AddColon = NotDefault
            BorderLineStyle =0
            Width =1701
            LabelX =-1701
            FontSize =11
            FontName ="Segoe UI"
            AllowValueListEdits =1
            InheritValueList =1
            ThemeFontIndex =1
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            ForeThemeColorIndex =2
            ForeShade =50.0
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
        Begin NavigationControl
            BorderWidth =1
            BorderLineStyle =0
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin NavigationButton
            Width =283
            Height =283
            ForeColor =-2
            GridlineThemeColorIndex =1
            GridlineShade =65.0
            HoverColor =-2
            HoverThemeColorIndex =2
            HoverTint =20.0
            PressedColor =-2
            PressedThemeColorIndex =2
            PressedTint =60.0
            HoverForeColor =-2
            HoverForeThemeColorIndex =0
            HoverForeTint =75.0
            PressedForeColor =-2
            PressedForeThemeColorIndex =0
            PressedForeTint =75.0
            BackThemeColorIndex =1
            OldBorderStyle =0
            BorderLineStyle =0
            BorderThemeColorIndex =3
            BorderShade =90.0
            ThemeFontIndex =1
            FontName ="Segoe UI"
            FontWeight =400
            FontSize =11
            ForeThemeColorIndex =0
            ForeTint =75.0
        End
        Begin FormHeader
            Height =980
            Name ="LomakkeenYlätunniste"
            AutoHeight =1
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =2
            BackTint =20.0
            Begin
                Begin Label
                    OverlapFlags =85
                    TextFontCharSet =177
                    TextFontFamily =0
                    Left =57
                    Top =57
                    Width =3977
                    Height =561
                    FontSize =20
                    BorderColor =8355711
                    Name ="Selite16"
                    Caption ="Yhteystiedon muokkaus"
                    FontName ="Calibri"
                    GridlineColor =10921638
                    LayoutCachedLeft =57
                    LayoutCachedTop =57
                    LayoutCachedWidth =4034
                    LayoutCachedHeight =618
                    ForeTint =100.0
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
                    Left =7316
                    Top =396
                    Width =1755
                    Height =300
                    BackColor =12566463
                    BorderColor =10921638
                    Name ="Automaattinen_päivämäärä"
                    ControlSource ="=Date()"
                    Format ="Short Date"
                    FontName ="Calibri"
                    GridlineColor =10921638

                    LayoutCachedLeft =7316
                    LayoutCachedTop =396
                    LayoutCachedWidth =9071
                    LayoutCachedHeight =696
                    BackShade =75.0
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =247
                    TextFontCharSet =177
                    TextAlign =3
                    TextFontFamily =0
                    BackStyle =0
                    IMESentenceMode =3
                    Left =7316
                    Top =680
                    Width =1755
                    Height =300
                    TabIndex =1
                    BackColor =12566463
                    BorderColor =10921638
                    Name ="Automaattinen_aika"
                    ControlSource ="=Time()"
                    Format ="Long Time"
                    FontName ="Calibri"
                    GridlineColor =10921638

                    LayoutCachedLeft =7316
                    LayoutCachedTop =680
                    LayoutCachedWidth =9071
                    LayoutCachedHeight =980
                    BackShade =75.0
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                End
            End
        End
        Begin Section
            CanGrow = NotDefault
            Height =9250
            Name ="Tiedot"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin Rectangle
                    BackStyle =1
                    OverlapFlags =93
                    Left =510
                    Top =1020
                    Width =5726
                    Height =5678
                    BackColor =14277081
                    BorderColor =10921638
                    Name ="Ruutu140"
                    GridlineColor =10921638
                    LayoutCachedLeft =510
                    LayoutCachedTop =1020
                    LayoutCachedWidth =6236
                    LayoutCachedHeight =6698
                    BackShade =85.0
                End
                Begin Rectangle
                    SpecialEffect =0
                    BackStyle =1
                    BorderWidth =2
                    OverlapFlags =93
                    Left =6689
                    Top =4478
                    Width =2211
                    Height =907
                    BackColor =62207
                    Name ="Ruutu139"
                    GridlineColor =10921638
                    LayoutCachedLeft =6689
                    LayoutCachedTop =4478
                    LayoutCachedWidth =8900
                    LayoutCachedHeight =5385
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =5
                    BorderShade =100.0
                End
                Begin EmptyCell
                    Left =3288
                    Top =924
                    Width =2940
                    Height =1128
                    Name ="TyhjäSolu84"
                    GroupTable =2
                    GridlineColor =10921638
                    LayoutCachedLeft =3288
                    LayoutCachedTop =924
                    LayoutCachedWidth =6228
                    LayoutCachedHeight =2052
                    RowStart =1
                    RowEnd =1
                    LayoutGroup =1
                    GroupTable =2
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    BorderWidth =2
                    OverlapFlags =215
                    TextFontCharSet =177
                    TextAlign =1
                    TextFontFamily =0
                    IMESentenceMode =3
                    Left =2304
                    Top =1020
                    Width =3840
                    Height =312
                    TabIndex =4
                    ForeColor =4210752
                    Name ="Tunnus"
                    ControlSource ="UID"
                    StatusBarText ="UID"
                    DefaultValue ="=\"\""
                    FontName ="Calibri"
                    Tag ="Tunnus_arvo"
                    GroupTable =8
                    BottomPadding =150
                    GridlineColor =10921638
                    BeforeUpdateEmMacro = Begin
                        Version =196611
                        ColumnsShown =0
                        Begin
                            Action ="SetFilter"
                            Argument ="Hae tietty käyttäjä"
                            Argument ="[Forms]![Tervetuloa]![Yhteystietovalinta]"
                            Argument ="UID"
                        End
                        Begin
                            Comment ="_AXL:<?xml version=\"1.0\" encoding=\"UTF-16\" standalone=\"no\"?>\015\012<UserI"
                                "nterfaceMacro For=\"Tunnus\" Event=\"BeforeUpdate\" xmlns=\"http://schemas.micro"
                                "soft.com/office/accessservices/2009/11/application\"><Statements><Action Name=\""
                                "SetFilter\"><Argument Name=\"Filte"
                        End
                        Begin
                            Comment ="_AXL:rName\">Hae tietty käyttäjä</Argument><Argument Name=\"WhereCondition\">[Fo"
                                "rms]![Tervetuloa]![Yhteystietovalinta]</Argument><Argument Name=\"ControlName\">"
                                "UID</Argument></Action></Statements></UserInterfaceMacro>"
                        End
                    End

                    LayoutCachedLeft =2304
                    LayoutCachedTop =1020
                    LayoutCachedWidth =6144
                    LayoutCachedHeight =1332
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =2
                    BorderThemeColorIndex =0
                    BorderShade =100.0
                    GroupTable =8
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =215
                            TextFontCharSet =177
                            TextAlign =1
                            TextFontFamily =0
                            Left =564
                            Top =1020
                            Width =1656
                            Height =312
                            FontWeight =700
                            BorderColor =8355711
                            ForeColor =6710886
                            Name ="Tunnus_Selite"
                            Caption ="ID"
                            FontName ="Calibri"
                            GroupTable =8
                            BottomPadding =150
                            GridlineColor =10921638
                            LayoutCachedLeft =564
                            LayoutCachedTop =1020
                            LayoutCachedWidth =2220
                            LayoutCachedHeight =1332
                            LayoutGroup =2
                            GroupTable =8
                        End
                    End
                End
                Begin TextBox
                    BorderWidth =2
                    OverlapFlags =223
                    TextFontCharSet =177
                    TextAlign =1
                    TextFontFamily =0
                    Left =2304
                    Top =2028
                    Width =3840
                    Height =312
                    ColumnWidth =3000
                    TabIndex =6
                    ForeColor =4210752
                    Name ="Sukunimi"
                    ControlSource ="Sukunimi"
                    FontName ="Calibri"
                    GroupTable =8
                    BottomPadding =150
                    GridlineColor =10921638

                    LayoutCachedLeft =2304
                    LayoutCachedTop =2028
                    LayoutCachedWidth =6144
                    LayoutCachedHeight =2340
                    RowStart =2
                    RowEnd =2
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =2
                    BorderThemeColorIndex =5
                    BorderShade =100.0
                    GroupTable =8
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =215
                            TextFontCharSet =177
                            TextAlign =1
                            TextFontFamily =0
                            Left =564
                            Top =2028
                            Width =1656
                            Height =312
                            FontWeight =700
                            BorderColor =8355711
                            ForeColor =6710886
                            Name ="Sukunimi_Selite"
                            Caption ="Sukunimi"
                            FontName ="Calibri"
                            GroupTable =8
                            BottomPadding =150
                            GridlineColor =10921638
                            LayoutCachedLeft =564
                            LayoutCachedTop =2028
                            LayoutCachedWidth =2220
                            LayoutCachedHeight =2340
                            RowStart =2
                            RowEnd =2
                            LayoutGroup =2
                            GroupTable =8
                        End
                    End
                End
                Begin TextBox
                    BorderWidth =2
                    OverlapFlags =215
                    TextFontCharSet =177
                    TextAlign =1
                    TextFontFamily =0
                    Left =2304
                    Top =1512
                    Width =3840
                    Height =312
                    ColumnWidth =3000
                    TabIndex =5
                    ForeColor =4210752
                    Name ="Etunimi"
                    ControlSource ="Etunimi"
                    FontName ="Calibri"
                    GroupTable =8
                    BottomPadding =150
                    GridlineColor =10921638

                    LayoutCachedLeft =2304
                    LayoutCachedTop =1512
                    LayoutCachedWidth =6144
                    LayoutCachedHeight =1824
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =2
                    BorderThemeColorIndex =5
                    BorderShade =100.0
                    GroupTable =8
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =215
                            TextFontCharSet =177
                            TextAlign =1
                            TextFontFamily =0
                            Left =564
                            Top =1512
                            Width =1656
                            Height =312
                            FontWeight =700
                            BorderColor =8355711
                            ForeColor =6710886
                            Name ="Etunimi_Selite"
                            Caption ="Etunimi"
                            FontName ="Calibri"
                            GroupTable =8
                            BottomPadding =150
                            GridlineColor =10921638
                            LayoutCachedLeft =564
                            LayoutCachedTop =1512
                            LayoutCachedWidth =2220
                            LayoutCachedHeight =1824
                            RowStart =1
                            RowEnd =1
                            LayoutGroup =2
                            GroupTable =8
                        End
                    End
                End
                Begin TextBox
                    BorderWidth =2
                    OverlapFlags =223
                    TextFontCharSet =177
                    TextAlign =1
                    TextFontFamily =0
                    Left =2304
                    Top =2544
                    Width =3840
                    Height =312
                    ColumnWidth =3000
                    TabIndex =7
                    ForeColor =4210752
                    Name ="Sähköpostiosoite"
                    ControlSource ="Sähköpostiosoite"
                    FontName ="Calibri"
                    GroupTable =8
                    BottomPadding =150
                    GridlineColor =10921638

                    LayoutCachedLeft =2304
                    LayoutCachedTop =2544
                    LayoutCachedWidth =6144
                    LayoutCachedHeight =2856
                    RowStart =3
                    RowEnd =3
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =2
                    BorderThemeColorIndex =5
                    BorderShade =100.0
                    GroupTable =8
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =215
                            TextFontCharSet =177
                            TextAlign =1
                            TextFontFamily =0
                            Left =564
                            Top =2544
                            Width =1656
                            Height =312
                            FontWeight =700
                            BorderColor =8355711
                            ForeColor =6710886
                            Name ="Sähköpostiosoite_Selite"
                            Caption ="Sähköpostiosoite"
                            FontName ="Calibri"
                            GroupTable =8
                            BottomPadding =150
                            GridlineColor =10921638
                            LayoutCachedLeft =564
                            LayoutCachedTop =2544
                            LayoutCachedWidth =2220
                            LayoutCachedHeight =2856
                            RowStart =3
                            RowEnd =3
                            LayoutGroup =2
                            GroupTable =8
                        End
                    End
                End
                Begin TextBox
                    BorderWidth =2
                    OverlapFlags =223
                    TextFontCharSet =177
                    TextAlign =1
                    TextFontFamily =0
                    IMEMode =2
                    Left =2304
                    Top =3060
                    Width =3840
                    Height =312
                    ColumnWidth =2748
                    TabIndex =8
                    ForeColor =4210752
                    Name ="Matkapuhelin"
                    ControlSource ="Matkapuhelin"
                    FontName ="Calibri"
                    GroupTable =8
                    BottomPadding =150
                    GridlineColor =10921638

                    LayoutCachedLeft =2304
                    LayoutCachedTop =3060
                    LayoutCachedWidth =6144
                    LayoutCachedHeight =3372
                    RowStart =4
                    RowEnd =4
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =2
                    BorderThemeColorIndex =5
                    BorderShade =100.0
                    GroupTable =8
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =215
                            TextFontCharSet =177
                            TextAlign =1
                            TextFontFamily =0
                            Left =564
                            Top =3060
                            Width =1656
                            Height =312
                            FontWeight =700
                            BorderColor =8355711
                            ForeColor =6710886
                            Name ="Matkapuhelin_Selite"
                            Caption ="Matkapuhelin"
                            FontName ="Calibri"
                            GroupTable =8
                            BottomPadding =150
                            GridlineColor =10921638
                            LayoutCachedLeft =564
                            LayoutCachedTop =3060
                            LayoutCachedWidth =2220
                            LayoutCachedHeight =3372
                            RowStart =4
                            RowEnd =4
                            LayoutGroup =2
                            GroupTable =8
                        End
                    End
                End
                Begin TextBox
                    BorderWidth =2
                    OverlapFlags =215
                    TextFontCharSet =177
                    TextAlign =1
                    TextFontFamily =0
                    Left =2304
                    Top =3576
                    Width =3840
                    Height =312
                    ColumnWidth =3000
                    TabIndex =9
                    ForeColor =4210752
                    Name ="Kaupunki"
                    ControlSource ="Kaupunki"
                    FontName ="Calibri"
                    GroupTable =8
                    BottomPadding =150
                    GridlineColor =10921638

                    LayoutCachedLeft =2304
                    LayoutCachedTop =3576
                    LayoutCachedWidth =6144
                    LayoutCachedHeight =3888
                    RowStart =5
                    RowEnd =5
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =2
                    BorderThemeColorIndex =5
                    BorderShade =100.0
                    GroupTable =8
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =215
                            TextFontCharSet =177
                            TextAlign =1
                            TextFontFamily =0
                            Left =564
                            Top =3576
                            Width =1656
                            Height =312
                            FontWeight =700
                            BorderColor =8355711
                            ForeColor =6710886
                            Name ="Kaupunki_Selite"
                            Caption ="Kaupunki"
                            FontName ="Calibri"
                            GroupTable =8
                            BottomPadding =150
                            GridlineColor =10921638
                            LayoutCachedLeft =564
                            LayoutCachedTop =3576
                            LayoutCachedWidth =2220
                            LayoutCachedHeight =3888
                            RowStart =5
                            RowEnd =5
                            LayoutGroup =2
                            GroupTable =8
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    RowSourceTypeInt =1
                    BorderWidth =2
                    OverlapFlags =215
                    TextFontCharSet =177
                    TextAlign =1
                    TextFontFamily =0
                    IMESentenceMode =3
                    ListWidth =1440
                    Left =2304
                    Top =4092
                    Width =3840
                    Height =312
                    ColumnWidth =3000
                    TabIndex =10
                    ForeColor =3484194
                    Name ="Jäsenyys"
                    ControlSource ="Jäsenyys"
                    RowSourceType ="Value List"
                    RowSource ="\"Jäsen\";\"Edustusjäsen\";\"Hallitus\";\"Siivous\";\"Kirjanpito\";\"Vuokrananta"
                        "ja\";\"Muu\""
                    ColumnWidths ="1440"
                    StatusBarText ="Edustusjäsenyys, hallitus yms yms yms"
                    FontName ="Calibri"
                    GroupTable =8
                    BottomPadding =150
                    GridlineColor =10921638

                    LayoutCachedLeft =2304
                    LayoutCachedTop =4092
                    LayoutCachedWidth =6144
                    LayoutCachedHeight =4404
                    RowStart =6
                    RowEnd =6
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =2
                    BorderThemeColorIndex =5
                    BorderShade =100.0
                    GroupTable =8
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =215
                            TextFontCharSet =177
                            TextAlign =1
                            TextFontFamily =0
                            Left =564
                            Top =4092
                            Width =1656
                            Height =312
                            FontWeight =700
                            BorderColor =8355711
                            ForeColor =6710886
                            Name ="Jäsenyys_Selite"
                            Caption ="Jäsenyys"
                            FontName ="Calibri"
                            GroupTable =8
                            BottomPadding =150
                            GridlineColor =10921638
                            LayoutCachedLeft =564
                            LayoutCachedTop =4092
                            LayoutCachedWidth =2220
                            LayoutCachedHeight =4404
                            RowStart =6
                            RowEnd =6
                            LayoutGroup =2
                            GroupTable =8
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    BorderWidth =2
                    OverlapFlags =215
                    TextFontCharSet =177
                    TextAlign =1
                    TextFontFamily =0
                    IMESentenceMode =3
                    Left =2304
                    Top =4608
                    Width =3840
                    Height =1692
                    ColumnWidth =3000
                    TabIndex =11
                    ForeColor =4210752
                    Name ="Muistiinpanot"
                    ControlSource ="Muistiinpanot"
                    StatusBarText ="Sekalaiset muistiinpanot"
                    FontName ="Calibri"
                    GroupTable =8
                    BottomPadding =150
                    GridlineColor =10921638

                    LayoutCachedLeft =2304
                    LayoutCachedTop =4608
                    LayoutCachedWidth =6144
                    LayoutCachedHeight =6300
                    RowStart =7
                    RowEnd =7
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =2
                    BorderThemeColorIndex =5
                    BorderShade =100.0
                    GroupTable =8
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =215
                            TextFontCharSet =177
                            TextAlign =1
                            TextFontFamily =0
                            Left =564
                            Top =4608
                            Width =1656
                            Height =1692
                            FontWeight =700
                            BorderColor =8355711
                            ForeColor =6710886
                            Name ="Muistiinpanot_Selite"
                            Caption ="Muistiinpanot"
                            FontName ="Calibri"
                            GroupTable =8
                            BottomPadding =150
                            GridlineColor =10921638
                            LayoutCachedLeft =564
                            LayoutCachedTop =4608
                            LayoutCachedWidth =2220
                            LayoutCachedHeight =6300
                            RowStart =7
                            RowEnd =7
                            LayoutGroup =2
                            GroupTable =8
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontCharSet =177
                    TextFontFamily =0
                    Left =6300
                    Top =924
                    Width =2724
                    Height =1128
                    TabIndex =2
                    ForeColor =4210752
                    Name ="Tallennus"
                    Caption ="Tallenna"
                    OnClick ="[Event Procedure]"
                    FontName ="Calibri"
                    GroupTable =2
                    GridlineColor =10921638

                    LayoutCachedLeft =6300
                    LayoutCachedTop =924
                    LayoutCachedWidth =9024
                    LayoutCachedHeight =2052
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    BackColor =15123357
                    BorderColor =15123357
                    HoverColor =15652797
                    PressedColor =11957550
                    HoverForeColor =4210752
                    PressedForeColor =4210752
                    GroupTable =2
                    WebImagePaddingLeft =3
                    WebImagePaddingTop =3
                    WebImagePaddingRight =2
                    WebImagePaddingBottom =2
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontCharSet =177
                    TextFontFamily =0
                    Left =6300
                    Top =2124
                    Width =2724
                    Height =1128
                    TabIndex =3
                    ForeColor =4210752
                    Name ="Cancelli"
                    Caption ="Sulje tallentamatta"
                    FontName ="Calibri"
                    GroupTable =2
                    GridlineColor =10921638
                    OnClickEmMacro = Begin
                        Version =196611
                        ColumnsShown =0
                        Begin
                            Action ="OnError"
                            Argument ="0"
                        End
                        Begin
                            Condition ="[MacroError]<>0"
                            Action ="MsgBox"
                            Argument ="=[MacroError].[Description]"
                            Argument ="-1"
                            Argument ="0"
                        End
                        Begin
                            Action ="Close"
                            Argument ="-1"
                            Argument =""
                            Argument ="0"
                        End
                        Begin
                            Comment ="_AXL:<?xml version=\"1.0\" encoding=\"UTF-16\" standalone=\"no\"?>\015\012<UserI"
                                "nterfaceMacro For=\"Cancelli\" Event=\"OnClick\" xmlns=\"http://schemas.microsof"
                                "t.com/office/accessservices/2009/11/application\"><Statements><Action Name=\"OnE"
                                "rror\"/><ConditionalBlock><If><Co"
                        End
                        Begin
                            Comment ="_AXL:ndition>[MacroError]&lt;&gt;0</Condition><Statements><Action Name=\"Message"
                                "Box\"><Argument Name=\"Message\">=[MacroError].[Description]</Argument></Action>"
                                "</Statements></If></ConditionalBlock><Action Name=\"CloseWindow\"/></Statements>"
                                "</UserInterfaceMacr"
                        End
                        Begin
                            Comment ="_AXL:o>"
                        End
                    End

                    LayoutCachedLeft =6300
                    LayoutCachedTop =2124
                    LayoutCachedWidth =9024
                    LayoutCachedHeight =3252
                    RowStart =2
                    RowEnd =2
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    BackColor =15123357
                    BorderColor =15123357
                    HoverColor =15652797
                    PressedColor =11957550
                    HoverForeColor =4210752
                    PressedForeColor =4210752
                    GroupTable =2
                    WebImagePaddingLeft =3
                    WebImagePaddingTop =3
                    WebImagePaddingRight =2
                    WebImagePaddingBottom =2
                    Overlaps =1
                End
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =177
                    TextFontFamily =0
                    Left =3288
                    Top =288
                    Width =2940
                    Height =564
                    ForeColor =4210752
                    Name ="uusijasen"
                    Caption ="Uusi jäsen"
                    OnClick ="[Event Procedure]"
                    FontName ="Calibri"
                    GroupTable =2
                    GridlineColor =10921638

                    LayoutCachedLeft =3288
                    LayoutCachedTop =288
                    LayoutCachedWidth =6228
                    LayoutCachedHeight =852
                    LayoutGroup =1
                    BackColor =15123357
                    BorderColor =15123357
                    HoverColor =15652797
                    PressedColor =11957550
                    HoverForeColor =4210752
                    PressedForeColor =4210752
                    GroupTable =2
                    WebImagePaddingLeft =3
                    WebImagePaddingTop =3
                    WebImagePaddingRight =2
                    WebImagePaddingBottom =2
                    Overlaps =1
                End
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =177
                    TextFontFamily =0
                    Left =6300
                    Top =288
                    Width =2724
                    Height =564
                    TabIndex =1
                    ForeColor =4210752
                    Name ="deletejäsen"
                    Caption ="Poista jäsen"
                    OnClick ="[Event Procedure]"
                    FontName ="Calibri"
                    GroupTable =2
                    GridlineColor =10921638

                    LayoutCachedLeft =6300
                    LayoutCachedTop =288
                    LayoutCachedWidth =9024
                    LayoutCachedHeight =852
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    BackColor =15123357
                    BorderColor =15123357
                    HoverColor =15652797
                    PressedColor =11957550
                    HoverForeColor =4210752
                    PressedForeColor =4210752
                    GroupTable =2
                    WebImagePaddingLeft =3
                    WebImagePaddingTop =3
                    WebImagePaddingRight =2
                    WebImagePaddingBottom =2
                    Overlaps =1
                End
                Begin ListBox
                    ColumnHeads = NotDefault
                    OverlapFlags =87
                    TextFontCharSet =177
                    TextFontFamily =0
                    BorderWidth =3
                    IMESentenceMode =3
                    ColumnCount =3
                    Left =793
                    Top =6689
                    Width =6180
                    Height =1928
                    TabIndex =13
                    ForeColor =4210752
                    Name ="Korttilistaus"
                    RowSourceType ="Table/Query"
                    RowSource ="ListaaKortitJaLataukset"
                    ColumnWidths ="1440;1440"
                    OnDblClick ="[Event Procedure]"
                    FontName ="Calibri"
                    GridlineColor =10921638
                    AllowValueListEdits =0

                    LayoutCachedLeft =793
                    LayoutCachedTop =6689
                    LayoutCachedWidth =6973
                    LayoutCachedHeight =8617
                    BorderThemeColorIndex =0
                    BorderShade =100.0
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =223
                            TextFontCharSet =177
                            TextFontFamily =0
                            Left =793
                            Top =6349
                            Width =1812
                            Height =314
                            FontWeight =700
                            BorderColor =8355711
                            ForeColor =6710886
                            Name ="Yhteystiedon kortit_Selite"
                            Caption ="Yhteystiedon kortit"
                            FontName ="Calibri"
                            EventProcPrefix ="Yhteystiedon_kortit_Selite"
                            GridlineColor =10921638
                            LayoutCachedLeft =793
                            LayoutCachedTop =6349
                            LayoutCachedWidth =2605
                            LayoutCachedHeight =6663
                        End
                    End
                End
                Begin EmptyCell
                    Left =3288
                    Top =2124
                    Width =2940
                    Height =1128
                    Name ="TyhjäSolu87"
                    GroupTable =2
                    GridlineColor =10921638
                    LayoutCachedLeft =3288
                    LayoutCachedTop =2124
                    LayoutCachedWidth =6228
                    LayoutCachedHeight =3252
                    RowStart =2
                    RowEnd =2
                    LayoutGroup =1
                    GroupTable =2
                End
                Begin CheckBox
                    OverlapFlags =215
                    Left =6973
                    Top =4792
                    Width =276
                    Height =262
                    TabIndex =12
                    BorderColor =10921638
                    Name ="Edustusj"
                    ControlSource ="Edustusjasen"
                    GridlineColor =10921638

                    LayoutCachedLeft =6973
                    LayoutCachedTop =4792
                    LayoutCachedWidth =7249
                    LayoutCachedHeight =5054
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =247
                            TextFontCharSet =177
                            TextFontFamily =0
                            Left =7203
                            Top =4762
                            Width =1560
                            Height =336
                            FontWeight =700
                            BorderColor =8355711
                            ForeColor =6710886
                            Name ="Selite138"
                            Caption ="Edustusjäsen"
                            FontName ="Calibri"
                            GridlineColor =10921638
                            LayoutCachedLeft =7203
                            LayoutCachedTop =4762
                            LayoutCachedWidth =8763
                            LayoutCachedHeight =5098
                        End
                    End
                End
                Begin Label
                    OverlapFlags =85
                    TextFontCharSet =177
                    TextFontFamily =0
                    Left =793
                    Top =8674
                    Width =6228
                    Height =576
                    BorderColor =8355711
                    ForeColor =6710886
                    Name ="Selite141"
                    Caption ="Vinkki: sulje lomake ja siirry muokkaamaan korttia tuplaklikkaamalla!"
                    FontName ="Calibri"
                    GridlineColor =10921638
                    LayoutCachedLeft =793
                    LayoutCachedTop =8674
                    LayoutCachedWidth =7021
                    LayoutCachedHeight =9250
                End
            End
        End
        Begin FormFooter
            Height =0
            Name ="LomakkeenAlatunniste"
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


Private Sub deletejäsen_Click()
    If ([Form_YhteystietojenMuokkaus].Korttilistaus.ListCount > 0) Then
        MsgBox ("Yhteystiedolla on linkitettyjä kortteja, ei voida jatkaa!")
        Exit Sub
    Else
        'do delete stuff here
        Dim succs
        succs = Common.SaveToLog("Poistetaan käyttäjä ID: " & [Form_YhteystietojenMuokkaus].UID.Value)
        DoCmd.RunCommand acCmdDeleteRecord
    End If
    
End Sub



Private Sub Korttilistaus_DblClick(Cancel As Integer)
    [Form_Tervetuloa].Korttivalinta.Value = [Form_YhteystietojenMuokkaus].Korttilistaus.Value
    DoCmd.Close
    [Form_Tervetuloa].Korttivalinta.SetFocus
    [Form_Tervetuloa].Korttivalinta.Dropdown
    
    
End Sub

Private Sub Tallennus_Click()

DoCmd.RunCommand acCmdSaveRecord
Dim succs
succs = Common.SaveToLog("Tallennettu muutokset käyttäjälle ID: " & [Form_YhteystietojenMuokkaus].UID.Value)
succs = Common.SendMessageToMainScreen("Muutokset yhteystietoon tallennettu!")

'DoCmd.Close do not do this, it is not ok.

End Sub

Private Sub uusijasen_Click()
    Dim succs
    succs = Common.SaveToLog("Ollaan luomassa uutta yhteystietoa")
    DoCmd.GoToRecord , , acNewRec
End Sub
