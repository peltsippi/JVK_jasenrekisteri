Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    DefaultView =0
    ScrollBars =0
    ViewsAllowed =1
    BorderStyle =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =9544
    DatasheetFontHeight =11
    ItemSuffix =211
    Left =4044
    Top =3456
    Right =17796
    Bottom =11712
    RecSrcDt = Begin
        0x377662ec10c5e540
    End
    AfterUpdate ="[Event Procedure]"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Calibri"
    OnGotFocus ="[Event Procedure]"
    AllowDatasheetView =0
    OnSelectionChange ="[Event Procedure]"
    FilterOnLoad =0
    ShowPageMargins =0
    DisplayOnSharePointSite =1
    AllowLayoutView =0
    DatasheetAlternateBackColor =16053492
    DatasheetGridlinesColor12 =15062992
    FitToScreen =1
    Begin
        Begin Label
            BackStyle =0
            FontSize =11
            FontName ="Calibri"
        End
        Begin Rectangle
            SpecialEffect =3
            BackStyle =0
            BorderLineStyle =0
        End
        Begin Line
            BorderLineStyle =0
        End
        Begin Image
            BackStyle =0
            OldBorderStyle =0
            BorderLineStyle =0
            SizeMode =3
            PictureAlignment =2
        End
        Begin CommandButton
            FontSize =11
            FontWeight =400
            ForeColor =-2147483630
            FontName ="Calibri"
            BorderLineStyle =0
        End
        Begin CheckBox
            SpecialEffect =2
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
        End
        Begin BoundObjectFrame
            AddColon = NotDefault
            SizeMode =3
            SpecialEffect =2
            BorderLineStyle =0
            Width =4536
            Height =2835
            LabelX =-1701
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
        Begin UnboundObjectFrame
            SpecialEffect =2
            OldBorderStyle =1
        End
        Begin CustomControl
            SpecialEffect =2
        End
        Begin ToggleButton
            Width =283
            Height =283
            ForeThemeColorIndex =0
            ForeTint =75.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
            UseTheme =1
            Shape =2
            Bevel =1
            BackColor =-1
            BackThemeColorIndex =4
            BackTint =60.0
            OldBorderStyle =0
            BorderLineStyle =0
            BorderColor =-1
            BorderThemeColorIndex =4
            BorderTint =60.0
            ThemeFontIndex =1
            HoverThemeColorIndex =4
            HoverTint =40.0
            PressedThemeColorIndex =4
            PressedShade =75.0
            HoverForeThemeColorIndex =0
            HoverForeTint =75.0
            PressedForeThemeColorIndex =1
        End
        Begin EmptyCell
            Height =240
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin FormHeader
            Height =2527
            BackColor =8421504
            Name ="FormHeader"
            Begin
                Begin Label
                    OverlapFlags =93
                    TextAlign =1
                    Left =2480
                    Top =188
                    Width =4644
                    Height =1524
                    FontSize =28
                    ForeColor =16777215
                    Name ="Auto_Title0"
                    Caption ="JVK:n jäsenrekisteri \015\012ja korttilataukset \015\012"
                    FontName ="Segoe UI Light"
                    GridlineColor =-2147483609
                    LayoutCachedLeft =2480
                    LayoutCachedTop =188
                    LayoutCachedWidth =7124
                    LayoutCachedHeight =1712
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
                    Left =4299
                    Top =1842
                    Width =5103
                    Height =300
                    BorderColor =10921638
                    ForeColor =16777215
                    Name ="Automaattinen_päivämäärä"
                    ControlSource ="=Date()"
                    Format ="Short Date"
                    FontName ="Calibri"
                    GridlineColor =10921638

                    LayoutCachedLeft =4299
                    LayoutCachedTop =1842
                    LayoutCachedWidth =9402
                    LayoutCachedHeight =2142
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =255
                    TextFontCharSet =177
                    TextAlign =3
                    TextFontFamily =0
                    BackStyle =0
                    IMESentenceMode =3
                    Left =4299
                    Top =2078
                    Width =5103
                    Height =300
                    TabIndex =1
                    BorderColor =10921638
                    ForeColor =16777215
                    Name ="Automaattinen_aika"
                    ControlSource ="=Time()"
                    Format ="Long Time"
                    FontName ="Calibri"
                    GridlineColor =10921638

                    LayoutCachedLeft =4299
                    LayoutCachedTop =2078
                    LayoutCachedWidth =9402
                    LayoutCachedHeight =2378
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Image
                    PictureType =2
                    Left =236
                    Top =141
                    Width =2173
                    Height =1937
                    Name ="Kuva178"
                    Picture ="punttilogo_pieni_invert"

                    LayoutCachedLeft =236
                    LayoutCachedTop =141
                    LayoutCachedWidth =2409
                    LayoutCachedHeight =2078
                    TabIndex =2
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =1
                    Left =6944
                    Top =94
                    Width =2568
                    Height =1524
                    FontSize =20
                    FontWeight =700
                    ForeColor =255
                    Name ="Selite179"
                    Caption ="VAIN HALLITUKSEN KÄYTTÖÖN!!! \015\012"
                    FontName ="Segoe UI Light"
                    GridlineColor =-2147483609
                    LayoutCachedLeft =6944
                    LayoutCachedTop =94
                    LayoutCachedWidth =9512
                    LayoutCachedHeight =1618
                End
                Begin Label
                    OverlapFlags =247
                    Left =2574
                    Top =2055
                    Width =3213
                    Height =472
                    Name ="copyrightteksti"
                    Caption ="(C) Timo Pelkonen, 2022"
                    LayoutCachedLeft =2574
                    LayoutCachedTop =2055
                    LayoutCachedWidth =5787
                    LayoutCachedHeight =2527
                End
            End
        End
        Begin Section
            Height =7370
            Name ="Detail"
            BackThemeColorIndex =1
            Begin
                Begin Image
                    PictureType =2
                    Left =614
                    Top =3685
                    Width =8434
                    Height =3612
                    Name ="Bulldog"
                    Picture ="bulldog_pienempi"

                    LayoutCachedLeft =614
                    LayoutCachedTop =3685
                    LayoutCachedWidth =9048
                    LayoutCachedHeight =7297
                    TabIndex =21
                End
                Begin Rectangle
                    Visible = NotDefault
                    BackStyle =1
                    OverlapFlags =255
                    Left =307
                    Top =3377
                    Width =9000
                    Height =3103
                    BackColor =62207
                    Name ="Korttikorjaukset"
                    LayoutCachedLeft =307
                    LayoutCachedTop =3377
                    LayoutCachedWidth =9307
                    LayoutCachedHeight =6480
                End
                Begin Rectangle
                    SpecialEffect =0
                    BorderWidth =1
                    OverlapFlags =85
                    Left =9450
                    Top =450
                    Width =0
                    Height =4725
                    BorderColor =14869218
                    Name ="Box71"
                    LayoutCachedLeft =9450
                    LayoutCachedTop =450
                    LayoutCachedWidth =9450
                    LayoutCachedHeight =5175
                End
                Begin ComboBox
                    OverlapFlags =215
                    TextFontCharSet =177
                    TextFontFamily =0
                    IMESentenceMode =3
                    ColumnCount =3
                    ListWidth =4320
                    Left =566
                    Top =473
                    Width =5092
                    Height =516
                    TabIndex =1
                    BoundColumn =2
                    BorderColor =10921638
                    ForeColor =3484194
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"\";\"\";\"10\";\"100\""
                    Name ="Yhteystietovalinta"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT [ListaaKaikkiYhteystiedot].Sukunimi, [ListaaKaikkiYhteystiedot].Etunimi, "
                        "[ListaaKaikkiYhteystiedot].UID FROM ListaaKaikkiYhteystiedot ORDER BY [Sukunimi]"
                        ", [Etunimi]; "
                    ColumnWidths ="1440;1440;1440"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Calibri"
                    OnChange ="[Event Procedure]"
                    GridlineColor =10921638
                    AllowValueListEdits =0
                    ListItemsEditForm ="d"

                    LayoutCachedLeft =566
                    LayoutCachedTop =473
                    LayoutCachedWidth =5658
                    LayoutCachedHeight =989
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =566
                            Top =165
                            Width =1764
                            Height =314
                            Name ="Valitse yhteystieto_Selite"
                            Caption ="Valitse yhteystieto"
                            EventProcPrefix ="Valitse_yhteystieto_Selite"
                            LayoutCachedLeft =566
                            LayoutCachedTop =165
                            LayoutCachedWidth =2330
                            LayoutCachedHeight =479
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =6236
                    Top =165
                    Width =2832
                    Height =564
                    Name ="Main-jäsenmuokkaus"
                    Caption ="Yhteystiedon muokkaus"
                    OnClick ="[Event Procedure]"
                    EventProcPrefix ="Main_jäsenmuokkaus"

                    LayoutCachedLeft =6236
                    LayoutCachedTop =165
                    LayoutCachedWidth =9068
                    LayoutCachedHeight =729
                    WebImagePaddingLeft =3
                    WebImagePaddingTop =3
                    WebImagePaddingRight =2
                    WebImagePaddingBottom =2
                    Overlaps =1
                End
                Begin ComboBox
                    OverlapFlags =85
                    TextFontCharSet =177
                    TextFontFamily =0
                    IMESentenceMode =3
                    ColumnCount =3
                    ListWidth =4320
                    Left =566
                    Top =1534
                    Width =5092
                    Height =504
                    TabIndex =4
                    BorderColor =10921638
                    ForeColor =3484194
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"\";\"\";\"10\";\"510\""
                    Name ="Korttivalinta"
                    RowSourceType ="Table/Query"
                    RowSource ="ListaaKortitJaLataukset"
                    ColumnWidths ="1440;1440;1440"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Calibri"
                    GridlineColor =10921638

                    LayoutCachedLeft =566
                    LayoutCachedTop =1534
                    LayoutCachedWidth =5658
                    LayoutCachedHeight =2038
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =566
                            Top =1110
                            Width =1224
                            Height =314
                            Name ="Korttivalinta_Selite"
                            Caption ="Valitse kortti"
                            LayoutCachedLeft =566
                            LayoutCachedTop =1110
                            LayoutCachedWidth =1790
                            LayoutCachedHeight =1424
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =93
                    Left =6236
                    Top =803
                    Width =1414
                    Height =685
                    TabIndex =2
                    Name ="Lisääkortti"
                    Caption ="Lisää kortti"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =6236
                    LayoutCachedTop =803
                    LayoutCachedWidth =7650
                    LayoutCachedHeight =1488
                    WebImagePaddingLeft =3
                    WebImagePaddingTop =3
                    WebImagePaddingRight =2
                    WebImagePaddingBottom =2
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =87
                    Left =7653
                    Top =803
                    Width =1428
                    Height =684
                    TabIndex =3
                    Name ="poistalinkitys"
                    Caption ="Poista kortti"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =7653
                    LayoutCachedTop =803
                    LayoutCachedWidth =9081
                    LayoutCachedHeight =1487
                    WebImagePaddingLeft =3
                    WebImagePaddingTop =3
                    WebImagePaddingRight =2
                    WebImagePaddingBottom =2
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =590
                    Top =2173
                    Width =5073
                    Height =408
                    TabIndex =7
                    Name ="Tyhjennä"
                    Caption ="Tyhjennä valinnat ja päivitä ikkuna"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =590
                    LayoutCachedTop =2173
                    LayoutCachedWidth =5663
                    LayoutCachedHeight =2581
                    UseTheme =1
                    BackThemeColorIndex =1
                    BackShade =85.0
                    HoverThemeColorIndex =1
                    HoverShade =85.0
                    PressedThemeColorIndex =1
                    PressedShade =85.0
                    WebImagePaddingLeft =3
                    WebImagePaddingTop =3
                    WebImagePaddingRight =2
                    WebImagePaddingBottom =2
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =6236
                    Top =2078
                    Width =2833
                    Height =506
                    TabIndex =6
                    Name ="RegisterPayment"
                    Caption ="Kirjaa maksu kortille"
                    OnClickEmMacro = Begin
                        Version =196611
                        ColumnsShown =0
                        Begin
                            Action ="OpenForm"
                            Argument ="RekisteroiMaksu"
                            Argument ="0"
                            Argument =""
                            Argument =""
                            Argument ="-1"
                            Argument ="0"
                        End
                        Begin
                            Comment ="_AXL:<?xml version=\"1.0\" encoding=\"UTF-16\" standalone=\"no\"?>\015\012<UserI"
                                "nterfaceMacro For=\"RegisterPayment\" Event=\"OnClick\" xmlns=\"http://schemas.m"
                                "icrosoft.com/office/accessservices/2009/11/application\"><Statements><Action Nam"
                                "e=\"OpenForm\"><Argument Name=\"Fo"
                        End
                        Begin
                            Comment ="_AXL:rmName\">RekisteroiMaksu</Argument></Action></Statements></UserInterfaceMac"
                                "ro>"
                        End
                    End

                    LayoutCachedLeft =6236
                    LayoutCachedTop =2078
                    LayoutCachedWidth =9069
                    LayoutCachedHeight =2584
                    WebImagePaddingLeft =3
                    WebImagePaddingTop =3
                    WebImagePaddingRight =2
                    WebImagePaddingBottom =2
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =6236
                    Top =1559
                    Width =2833
                    Height =482
                    TabIndex =5
                    Name ="Korttilataus"
                    Caption ="Kirjaa lataus kortille"
                    OnClickEmMacro = Begin
                        Version =196611
                        ColumnsShown =0
                        Begin
                            Action ="OpenForm"
                            Argument ="RekisteroiLataus"
                            Argument ="0"
                            Argument =""
                            Argument =""
                            Argument ="-1"
                            Argument ="0"
                        End
                        Begin
                            Comment ="_AXL:<?xml version=\"1.0\" encoding=\"UTF-16\" standalone=\"no\"?>\015\012<UserI"
                                "nterfaceMacro For=\"Korttilataus\" Event=\"OnClick\" xmlns=\"http://schemas.micr"
                                "osoft.com/office/accessservices/2009/11/application\"><Statements><Action Name=\""
                                "OpenForm\"><Argument Name=\"FormN"
                        End
                        Begin
                            Comment ="_AXL:ame\">RekisteroiLataus</Argument></Action></Statements></UserInterfaceMacro"
                                ">"
                        End
                    End

                    LayoutCachedLeft =6236
                    LayoutCachedTop =1559
                    LayoutCachedWidth =9069
                    LayoutCachedHeight =2041
                    WebImagePaddingLeft =3
                    WebImagePaddingTop =3
                    WebImagePaddingRight =2
                    WebImagePaddingBottom =2
                End
                Begin Label
                    OverlapFlags =247
                    Left =566
                    Top =2716
                    Width =8511
                    Height =735
                    FontSize =14
                    ForeColor =5026082
                    Name ="Status"
                    Caption ="6.2.2022 13.46.59 - Maksujen korjaus valmis"
                    LayoutCachedLeft =566
                    LayoutCachedTop =2716
                    LayoutCachedWidth =9077
                    LayoutCachedHeight =3451
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =247
                    TextFontCharSet =177
                    TextFontFamily =0
                    IMESentenceMode =3
                    Left =5408
                    Top =3590
                    Width =1744
                    Height =300
                    TabIndex =9
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Puumerkki"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Calibri"
                    GridlineColor =10921638

                    LayoutCachedLeft =5408
                    LayoutCachedTop =3590
                    LayoutCachedWidth =7152
                    LayoutCachedHeight =3890
                    Begin
                        Begin Label
                            OverlapFlags =255
                            Left =3755
                            Top =3590
                            Width =1656
                            Height =300
                            Name ="Selite156"
                            Caption ="Puumerkki + enter:"
                            LayoutCachedLeft =3755
                            LayoutCachedTop =3590
                            LayoutCachedWidth =5411
                            LayoutCachedHeight =3890
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =247
                    Left =564
                    Top =4116
                    Width =1572
                    Height =588
                    TabIndex =11
                    Name ="MuokkaaLatauksia"
                    Caption ="Muokkaa latauksia"
                    OnClick ="[Event Procedure]"
                    GroupTable =6

                    LayoutCachedLeft =564
                    LayoutCachedTop =4116
                    LayoutCachedWidth =2136
                    LayoutCachedHeight =4704
                    LayoutGroup =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =6
                    WebImagePaddingLeft =3
                    WebImagePaddingTop =3
                    WebImagePaddingRight =2
                    WebImagePaddingBottom =2
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =247
                    Left =2208
                    Top =4116
                    Height =588
                    TabIndex =12
                    Name ="MuokkaaMaksuja"
                    Caption ="Muokkaa maksuja"
                    OnClick ="[Event Procedure]"
                    GroupTable =6

                    LayoutCachedLeft =2208
                    LayoutCachedTop =4116
                    LayoutCachedWidth =3648
                    LayoutCachedHeight =4704
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =6
                    WebImagePaddingLeft =3
                    WebImagePaddingTop =3
                    WebImagePaddingRight =2
                    WebImagePaddingBottom =2
                    Overlaps =1
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =247
                    Left =7606
                    Top =3472
                    Width =1464
                    Height =473
                    FontSize =16
                    FontWeight =700
                    BackColor =8435191
                    Name ="Raportit"
                    Caption ="Raportit:"
                    LayoutCachedLeft =7606
                    LayoutCachedTop =3472
                    LayoutCachedWidth =9070
                    LayoutCachedHeight =3945
                End
                Begin ToggleButton
                    OverlapFlags =247
                    TextFontCharSet =177
                    TextFontFamily =0
                    Left =543
                    Top =3496
                    Width =2997
                    Height =560
                    TabIndex =8
                    ForeColor =4210752
                    Name ="KorjaaTietoja"
                    Caption ="Admin -moodi"
                    FontName ="Calibri"
                    OnClick ="[Event Procedure]"
                    GridlineColor =10921638

                    LayoutCachedLeft =543
                    LayoutCachedTop =3496
                    LayoutCachedWidth =3540
                    LayoutCachedHeight =4056
                    BackColor =62207
                    BackThemeColorIndex =-1
                    BackTint =100.0
                    BorderColor =15123357
                    HoverColor =15652797
                    PressedColor =11957550
                    HoverForeColor =4210752
                    PressedForeColor =16777215
                    WebImagePaddingLeft =3
                    WebImagePaddingTop =2
                    WebImagePaddingRight =3
                    WebImagePaddingBottom =3
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =247
                    Left =3720
                    Top =4116
                    Width =1512
                    Height =588
                    TabIndex =13
                    Name ="Hinnat"
                    Caption ="Muokkaa korttihintoja"
                    OnClick ="[Event Procedure]"
                    GroupTable =6

                    LayoutCachedLeft =3720
                    LayoutCachedTop =4116
                    LayoutCachedWidth =5232
                    LayoutCachedHeight =4704
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =6
                    WebImagePaddingLeft =3
                    WebImagePaddingTop =3
                    WebImagePaddingRight =2
                    WebImagePaddingBottom =2
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =247
                    Left =2220
                    Top =4908
                    Width =1584
                    Height =804
                    TabIndex =16
                    Name ="RaporttiMaksamatta"
                    Caption ="Tarkasta maksamattomat kortit"
                    OnClick ="[Event Procedure]"
                    GroupTable =1

                    LayoutCachedLeft =2220
                    LayoutCachedTop =4908
                    LayoutCachedWidth =3804
                    LayoutCachedHeight =5712
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =8435191
                    HoverColor =8435191
                    PressedColor =8435191
                    GroupTable =1
                    WebImagePaddingLeft =3
                    WebImagePaddingTop =3
                    WebImagePaddingRight =2
                    WebImagePaddingBottom =2
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =247
                    Left =3118
                    Top =5858
                    Width =3266
                    Height =550
                    TabIndex =20
                    Name ="Historia"
                    Caption ="Sovelluksen täysi historia"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =3118
                    LayoutCachedTop =5858
                    LayoutCachedWidth =6384
                    LayoutCachedHeight =6408
                    UseTheme =1
                    BackColor =8435191
                    HoverColor =8435191
                    PressedColor =8435191
                    WebImagePaddingLeft =3
                    WebImagePaddingTop =3
                    WebImagePaddingRight =2
                    WebImagePaddingBottom =2
                    Overlaps =1
                End
                Begin TextBox
                    OverlapFlags =247
                    TextFontCharSet =177
                    TextFontFamily =0
                    IMESentenceMode =3
                    Left =7110
                    Top =4015
                    Width =1961
                    Height =300
                    TabIndex =10
                    BackColor =8435191
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="RaportitAlku"
                    Format ="Short Date"
                    DefaultValue ="=DateSerial(Year(Now()),Month(Now())-6,Day(Now()))"
                    FontName ="Calibri"
                    GridlineColor =10921638

                    LayoutCachedLeft =7110
                    LayoutCachedTop =4015
                    LayoutCachedWidth =9071
                    LayoutCachedHeight =4315
                    BackThemeColorIndex =-1
                    Begin
                        Begin Label
                            OverlapFlags =255
                            Left =6283
                            Top =4015
                            Width =1284
                            Height =300
                            FontWeight =700
                            Name ="Selite166"
                            Caption ="Alkaen"
                            LayoutCachedLeft =6283
                            LayoutCachedTop =4015
                            LayoutCachedWidth =7567
                            LayoutCachedHeight =4315
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =247
                    TextFontCharSet =177
                    TextFontFamily =0
                    IMESentenceMode =3
                    Left =7109
                    Top =4346
                    Width =1961
                    Height =300
                    TabIndex =14
                    BackColor =8435191
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="RaportitLoppu"
                    Format ="Short Date"
                    DefaultValue ="=Date()"
                    FontName ="Calibri"
                    GridlineColor =10921638

                    LayoutCachedLeft =7109
                    LayoutCachedTop =4346
                    LayoutCachedWidth =9070
                    LayoutCachedHeight =4646
                    BackThemeColorIndex =-1
                    Begin
                        Begin Label
                            OverlapFlags =255
                            Left =6236
                            Top =4346
                            Width =1284
                            Height =300
                            FontWeight =700
                            Name ="Selite168"
                            Caption ="Loppuen"
                            LayoutCachedLeft =6236
                            LayoutCachedTop =4346
                            LayoutCachedWidth =7520
                            LayoutCachedHeight =4646
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =247
                    Left =3876
                    Top =4908
                    Width =1728
                    Height =804
                    TabIndex =17
                    Name ="LatauksetKaikki"
                    Caption ="Korttilatausten kokonaisraportti"
                    OnClick ="[Event Procedure]"
                    GroupTable =1

                    LayoutCachedLeft =3876
                    LayoutCachedTop =4908
                    LayoutCachedWidth =5604
                    LayoutCachedHeight =5712
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =8435191
                    HoverColor =8435191
                    PressedColor =8435191
                    GroupTable =1
                    WebImagePaddingLeft =3
                    WebImagePaddingTop =3
                    WebImagePaddingRight =2
                    WebImagePaddingBottom =2
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =247
                    Left =588
                    Top =4908
                    Width =1560
                    Height =804
                    TabIndex =15
                    Name ="KortinTapahtumat"
                    Caption ="Valitun kortin tapahtumat"
                    OnClick ="[Event Procedure]"
                    GroupTable =1

                    LayoutCachedLeft =588
                    LayoutCachedTop =4908
                    LayoutCachedWidth =2148
                    LayoutCachedHeight =5712
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =8435191
                    HoverColor =8435191
                    PressedColor =8435191
                    GroupTable =1
                    WebImagePaddingLeft =3
                    WebImagePaddingTop =3
                    WebImagePaddingRight =2
                    WebImagePaddingBottom =2
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =247
                    Left =5676
                    Top =4908
                    Height =804
                    TabIndex =18
                    Name ="ListaaEdustusj"
                    Caption ="Listaa edustusjäsenet"
                    OnClick ="[Event Procedure]"
                    GroupTable =1

                    LayoutCachedLeft =5676
                    LayoutCachedTop =4908
                    LayoutCachedWidth =7116
                    LayoutCachedHeight =5712
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =8435191
                    HoverColor =8435191
                    PressedColor =8435191
                    GroupTable =1
                    WebImagePaddingLeft =3
                    WebImagePaddingTop =3
                    WebImagePaddingRight =2
                    WebImagePaddingBottom =2
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =247
                    Left =7188
                    Top =4908
                    Height =804
                    TabIndex =19
                    Name ="IlmaiseksiLadattavat"
                    Caption ="Listaa ilmaiseksi ladattavat"
                    OnClick ="[Event Procedure]"
                    GroupTable =1

                    LayoutCachedLeft =7188
                    LayoutCachedTop =4908
                    LayoutCachedWidth =8628
                    LayoutCachedHeight =5712
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =8435191
                    HoverColor =8435191
                    PressedColor =8435191
                    GroupTable =1
                    WebImagePaddingLeft =3
                    WebImagePaddingTop =3
                    WebImagePaddingRight =2
                    WebImagePaddingBottom =2
                    Overlaps =1
                End
            End
        End
        Begin FormFooter
            Height =0
            Name ="FormFooter"
        End
    End
End
CodeBehindForm
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database


Public Function Paivita_korttiluettelo()
    
    Dim ctlCombo As Control
    Set ctlCombo = Form_Tervetuloa.Korttivalinta
    ctlCombo.Requery
    

End Function




Private Sub Form_AfterUpdate()
    Common.EnableDisableButtons
    
End Sub



Private Sub Form_GotFocus()
    [Form_Tervetuloa].Refresh
    
    'DoCmd.Requery "[Form].[Tervetuloa]"
    
End Sub


Private Sub Form_Open(Cancel As Integer)
    Dim succs
    succs = Common.EnableDisableButtons()
    succs = Common.SendMessageToMainScreen("Tervetuloa!")
End Sub

Private Sub Form_SelectionChange()
    Common.EnableDisableButtons
    
End Sub


Private Sub Hinnat_Click()
    Dim succs
    succs = Common.SaveToLog([Form_Tervetuloa].Puumerkki.Value & " aloitti hinnaston muokkaamisen.")
    succs = Common.SendMessageToMainScreen("Hinnaston muokkaaminen aloitettu")
    DoCmd.OpenForm "MuokkaaHinnastoa"
End Sub

Private Sub Historia_Click()
    Dim succs
    succs = Common.SaveToLog([Form_Tervetuloa].Puumerkki.Value & " avasi historiaraportin")
    'MsgBox ("[Aika] Between #" & [Form_Tervetuloa].RaportitAlku.Value & " 00:00:00# AND #" & [Form_Tervetuloa].RaportitLoppu.Value & " 23:59:59#")
    DoCmd.OpenReport "HistoriaListaus", acViewPreview, "HistoriaListaus"
End Sub

Private Sub IlmaiseksiLadattavat_Click()
    Dim succs
    succs = Common.SaveToLog([Form_Tervetuloa].Puumerkki.Value & " pyysi listauksen ilmaiseksi ladattavista korteista")
    DoCmd.OpenReport "HaeIlmaiseksiLadattavat", acViewPreview, "HaeIlmaiseksiLadattavat"
End Sub

Private Sub KorjaaTietoja_Click()
    'MsgBox ([Form_Tervetuloa].KorjaaTietoja.Value)
    Dim upd
    upd = Common.EnableDisableButtons()
    
End Sub

Private Sub KortinTapahtumat_Click()
    Dim succs
    succs = Common.SaveToLog([Form_Tervetuloa].Puumerkki.Value & " avasi kortin " & [Form_Tervetuloa].Korttivalinta.Value & " tapahtumaraportin")
    'MsgBox ("[Aika] Between #" & [Form_Tervetuloa].RaportitAlku.Value & " 00:00:00# AND #" & [Form_Tervetuloa].RaportitLoppu.Value & " 23:59:59#")
    DoCmd.OpenReport "KortinTapahtumat", acViewPreview
End Sub

Private Sub Korttivalinta_AfterUpdate()
Common.EnableDisableButtons
    
    Dim criteria As String
    
    criteria = "Kortti = '" & [Form_Tervetuloa].Korttivalinta.Value & "'"
    
    Dim user As Integer
    user = Common.FetchGeneralID("Kortit", "Omistaja", criteria)
    'MsgBox (user)
    [Form_Tervetuloa].Yhteystietovalinta.Value = Common.FetchGeneralID("Kortit", "Omistaja", criteria)
    [Form_Tervetuloa].Refresh
End Sub


Private Sub LatauksetKaikki_Click()
    Dim succs
    succs = Common.SaveToLog([Form_Tervetuloa].Puumerkki.Value & " avasi kaikkien latausten raportin")
    DoCmd.OpenReport "LatauksetKaikki", acViewPreview

End Sub

Private Sub Lisääkortti_Click()
    If IsNull(Form_Tervetuloa.Yhteystietovalinta) Or (Form_Tervetuloa.Yhteystietovalinta.Value = "") Then
        MsgBox ("Yhteystieto tulee olla valittuna tätä toimintoa varten!")
    Else
        DoCmd.OpenForm ("LisaaKortinLinkitys")
        
    End If
    
End Sub

Private Sub ListaaEdustusj_Click()
    Dim succs
    succs = Common.SaveToLog([Form_Tervetuloa].Puumerkki.Value & " pyysi edustusjäsenlistauksen")
    DoCmd.OpenReport "ListaaEdustusjasenet", acViewPreview
End Sub

Private Sub Main_jäsenmuokkaus_Click()
Dim succs
succs = Common.SaveToLog("Yhteystietojen muokkaushommat aloitettu")
'DoCmd.OpenForm ("Form_YhteystietojenMuokkaus")
DoCmd.OpenForm "YhteystietojenMuokkaus"

End Sub

Private Sub MuokkaaLatauksia_Click()
    Dim succs
    succs = Common.SaveToLog(Puumerkki & " aloitti kortin " & [Form_Tervetuloa].Korttivalinta.Value & " latausten muokkaamisen")
    DoCmd.OpenForm "KorjaaKortinLatauksia"
    succs = Common.SendMessageToMainScreen("Kortin " & [Form_Tervetuloa].Korttivalinta.Value & " latausmuokkaus aloitettu")
    
End Sub

Private Sub MuokkaaMaksuja_Click()
    Dim succs
    succs = Common.SaveToLog(Puumerkki & " aloitti kortin " & [Form_Tervetuloa].Korttivalinta.Value & " maksujen muokkaamisen")
    DoCmd.OpenForm "KorjaaKortinMaksuja"
    succs = Common.SendMessageToMainScreen("Kortin " & [Form_Tervetuloa].Korttivalinta.Value & " maksumuokkaus aloitettu")
    
End Sub

Private Sub poistalinkitys_Click()
    If IsNull(Form_Tervetuloa.Yhteystietovalinta) Or (Form_Tervetuloa.Yhteystietovalinta.Value = "") Or IsNull(Form_Tervetuloa.Korttivalinta) Or (Form_Tervetuloa.Korttivalinta.Value = "") Then
        MsgBox ("Kortti ja yhteystieto tulee olla valittuna tätä toimintoa varten!")
    Else
        DoCmd.OpenForm ("PoistaKortinLinkitys")
    End If
    
End Sub

Private Sub Puumerkki_AfterUpdate()
    Dim succs
    succs = Common.EnableDisableButtons()
End Sub

Private Sub RaporttiMaksamatta_Click()
    Dim succs
    succs = Common.SaveToLog([Form_Tervetuloa].Puumerkki.Value & " avasi maksamattomien korttien raportin")
    DoCmd.OpenReport "KortitMaksamatta", acViewPreview
End Sub

Private Sub Tyhjennä_Click()

    Form_Tervetuloa.Yhteystietovalinta = Null
    Form_Tervetuloa.Korttivalinta = Null
    Form_Tervetuloa.Paivita_korttiluettelo
    Common.EnableDisableButtons
    [Form_Tervetuloa].Refresh
    

End Sub

Private Sub Yhteystietovalinta_AfterUpdate()

    'Form_Tervetuoa.Paivita_korttiluettelo()
    Form_Tervetuloa.Paivita_korttiluettelo
    
    
End Sub


Private Sub Yhteystietovalinta_Change()
    Common.EnableDisableButtons
    
End Sub
