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
    Width =9708
    DatasheetFontHeight =11
    ItemSuffix =294
    Left =4740
    Top =3468
    Right =22788
    Bottom =11712
    RecSrcDt = Begin
        0x23fa53ee5dc7e540
    End
    RecordSource ="KorttejaVoimassa"
    AfterUpdate ="[Event Procedure]"
    OnOpen ="[Event Procedure]"
    OnClose ="[Event Procedure]"
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
            Height =2828
            BackColor =8421504
            Name ="FormHeader"
            Begin
                Begin Label
                    OverlapFlags =93
                    TextAlign =1
                    Left =2480
                    Top =188
                    Width =4836
                    Height =1776
                    FontSize =28
                    ForeColor =16777215
                    Name ="Auto_Title0"
                    Caption ="JVK:n jäsenrekisteri \015\012ja korttilataukset \015\012"
                    FontName ="Segoe UI Light"
                    GridlineColor =-2147483609
                    LayoutCachedLeft =2480
                    LayoutCachedTop =188
                    LayoutCachedWidth =7316
                    LayoutCachedHeight =1964
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
                    Left =8409
                    Top =1795
                    Width =1119
                    Height =300
                    BorderColor =10921638
                    ForeColor =16777215
                    Name ="Automaattinen_päivämäärä"
                    ControlSource ="=Date()"
                    Format ="Short Date"
                    FontName ="Calibri"
                    GridlineColor =10921638

                    LayoutCachedLeft =8409
                    LayoutCachedTop =1795
                    LayoutCachedWidth =9528
                    LayoutCachedHeight =2095
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =177
                    TextAlign =3
                    TextFontFamily =0
                    BackStyle =0
                    IMESentenceMode =3
                    Left =8598
                    Top =2173
                    Width =1083
                    Height =288
                    TabIndex =1
                    BorderColor =10921638
                    ForeColor =16777215
                    Name ="Automaattinen_aika"
                    ControlSource ="=Time()"
                    Format ="Long Time"
                    FontName ="Calibri"
                    GridlineColor =10921638

                    LayoutCachedLeft =8598
                    LayoutCachedTop =2173
                    LayoutCachedWidth =9681
                    LayoutCachedHeight =2461
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
                    TabIndex =8
                End
                Begin Label
                    OverlapFlags =255
                    TextAlign =1
                    Left =6944
                    Top =94
                    Width =2712
                    Height =1776
                    FontSize =16
                    FontWeight =700
                    ForeColor =255
                    Name ="Selite179"
                    Caption ="VAIN HALLITUKSEN KÄYTTÖÖN!!! \015\012"
                    FontName ="Segoe UI Light"
                    GridlineColor =-2147483609
                    LayoutCachedLeft =6944
                    LayoutCachedTop =94
                    LayoutCachedWidth =9656
                    LayoutCachedHeight =1870
                End
                Begin Label
                    OverlapFlags =215
                    Left =2409
                    Top =47
                    Width =3213
                    Height =472
                    Name ="copyrightteksti"
                    Caption ="(C) Timo Pelkonen, 2022"
                    LayoutCachedLeft =2409
                    LayoutCachedTop =47
                    LayoutCachedWidth =5622
                    LayoutCachedHeight =519
                End
                Begin Label
                    OverlapFlags =93
                    Left =165
                    Top =2149
                    Width =2076
                    Height =300
                    Name ="Selite216"
                    Caption ="Kortteja aktiivisena:"
                    LayoutCachedLeft =165
                    LayoutCachedTop =2149
                    LayoutCachedWidth =2241
                    LayoutCachedHeight =2449
                    ForeThemeColorIndex =7
                    ForeTint =40.0
                End
                Begin TextBox
                    OverlapFlags =223
                    TextFontCharSet =177
                    TextFontFamily =0
                    IMESentenceMode =3
                    Left =3543
                    Top =1818
                    Width =1229
                    Height =300
                    TabIndex =2
                    BackColor =10921638
                    BorderColor =10921638
                    Name ="kortitKK"
                    ControlSource ="Kuukausikortit"
                    Format ="General Number"
                    FontName ="Calibri"
                    GridlineColor =10921638

                    LayoutCachedLeft =3543
                    LayoutCachedTop =1818
                    LayoutCachedWidth =4772
                    LayoutCachedHeight =2118
                    BackShade =65.0
                    ForeThemeColorIndex =7
                    ForeTint =60.0
                End
                Begin Label
                    OverlapFlags =223
                    Left =3023
                    Top =1818
                    Width =492
                    Height =300
                    Name ="Selite245"
                    Caption ="KK:"
                    LayoutCachedLeft =3023
                    LayoutCachedTop =1818
                    LayoutCachedWidth =3515
                    LayoutCachedHeight =2118
                    ForeThemeColorIndex =7
                    ForeTint =40.0
                End
                Begin TextBox
                    OverlapFlags =247
                    TextFontCharSet =177
                    TextFontFamily =0
                    IMESentenceMode =3
                    Left =3543
                    Top =2101
                    Width =1229
                    Height =300
                    TabIndex =3
                    BackColor =10921638
                    BorderColor =10921638
                    Name ="kortitAP"
                    ControlSource ="Aamupvkortit"
                    Format ="General Number"
                    FontName ="Calibri"
                    GridlineColor =10921638

                    LayoutCachedLeft =3543
                    LayoutCachedTop =2101
                    LayoutCachedWidth =4772
                    LayoutCachedHeight =2401
                    BackShade =65.0
                    ForeThemeColorIndex =7
                    ForeTint =60.0
                End
                Begin Label
                    OverlapFlags =215
                    Left =3023
                    Top =2101
                    Width =492
                    Height =300
                    Name ="Selite247"
                    Caption ="AP:"
                    LayoutCachedLeft =3023
                    LayoutCachedTop =2101
                    LayoutCachedWidth =3515
                    LayoutCachedHeight =2401
                    ForeThemeColorIndex =7
                    ForeTint =40.0
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =177
                    TextFontFamily =0
                    IMESentenceMode =3
                    Left =3543
                    Top =2456
                    Width =1229
                    Height =300
                    TabIndex =4
                    BackColor =10921638
                    BorderColor =10921638
                    Name ="kortitOpisk"
                    ControlSource ="Opiskelijakortit"
                    Format ="General Number"
                    FontName ="Calibri"
                    GridlineColor =10921638

                    LayoutCachedLeft =3543
                    LayoutCachedTop =2456
                    LayoutCachedWidth =4772
                    LayoutCachedHeight =2756
                    BackShade =65.0
                    ForeThemeColorIndex =7
                    ForeTint =60.0
                End
                Begin Label
                    OverlapFlags =85
                    Left =2787
                    Top =2479
                    Width =648
                    Height =300
                    Name ="Selite249"
                    Caption ="OPISK:"
                    LayoutCachedLeft =2787
                    LayoutCachedTop =2479
                    LayoutCachedWidth =3435
                    LayoutCachedHeight =2779
                    ForeThemeColorIndex =7
                    ForeTint =40.0
                End
                Begin TextBox
                    OverlapFlags =247
                    TextFontCharSet =177
                    TextFontFamily =0
                    IMESentenceMode =3
                    Left =5858
                    Top =1771
                    Width =1229
                    Height =300
                    TabIndex =5
                    BackColor =10921638
                    BorderColor =10921638
                    Name ="kortitKrt"
                    ControlSource ="Kertakortit"
                    Format ="General Number"
                    FontName ="Calibri"
                    GridlineColor =10921638

                    LayoutCachedLeft =5858
                    LayoutCachedTop =1771
                    LayoutCachedWidth =7087
                    LayoutCachedHeight =2071
                    BackShade =65.0
                    ForeThemeColorIndex =7
                    ForeTint =60.0
                End
                Begin Label
                    OverlapFlags =215
                    Left =5149
                    Top =1795
                    Width =696
                    Height =300
                    Name ="Selite251"
                    Caption ="Kertak:"
                    LayoutCachedLeft =5149
                    LayoutCachedTop =1795
                    LayoutCachedWidth =5845
                    LayoutCachedHeight =2095
                    ForeThemeColorIndex =7
                    ForeTint =40.0
                End
                Begin TextBox
                    OverlapFlags =215
                    TextFontCharSet =177
                    TextFontFamily =0
                    IMESentenceMode =3
                    Left =1937
                    Top =2149
                    Width =941
                    Height =300
                    TabIndex =6
                    BackColor =10921638
                    BorderColor =10921638
                    Name ="kortitKaikki"
                    ControlSource ="Kaikki"
                    Format ="General Number"
                    FontName ="Calibri"
                    GridlineColor =10921638

                    LayoutCachedLeft =1937
                    LayoutCachedTop =2149
                    LayoutCachedWidth =2878
                    LayoutCachedHeight =2449
                    BackShade =65.0
                    ForeThemeColorIndex =7
                    ForeTint =60.0
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =177
                    TextFontFamily =0
                    IMESentenceMode =3
                    Left =5858
                    Top =2149
                    Width =1229
                    Height =300
                    TabIndex =7
                    BackColor =10921638
                    BorderColor =10921638
                    Name ="kortitMuu"
                    ControlSource ="Muut"
                    Format ="General Number"
                    FontName ="Calibri"
                    GridlineColor =10921638

                    LayoutCachedLeft =5858
                    LayoutCachedTop =2149
                    LayoutCachedWidth =7087
                    LayoutCachedHeight =2449
                    BackShade =65.0
                    ForeThemeColorIndex =7
                    ForeTint =60.0
                End
                Begin Label
                    OverlapFlags =85
                    Left =5102
                    Top =2125
                    Width =696
                    Height =300
                    Name ="Selite254"
                    Caption ="Muu:"
                    LayoutCachedLeft =5102
                    LayoutCachedTop =2125
                    LayoutCachedWidth =5798
                    LayoutCachedHeight =2425
                    ForeThemeColorIndex =7
                    ForeTint =40.0
                End
            End
        End
        Begin Section
            Height =8187
            Name ="Detail"
            BackThemeColorIndex =1
            Begin
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
                    OnGotFocus ="[Event Procedure]"
                    OnChange ="[Event Procedure]"
                    GridlineColor =10921638
                    AllowValueListEdits =0
                    ListItemsEditForm ="YhteystietojenMuokkaus"

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
                    OnGotFocus ="[Event Procedure]"
                    GridlineColor =10921638
                    ListItemsEditForm ="RekisteroiLataus"

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
                    TabIndex =8
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
                    OverlapFlags =93
                    Left =7688
                    Top =2078
                    Width =1381
                    Height =662
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

                    LayoutCachedLeft =7688
                    LayoutCachedTop =2078
                    LayoutCachedWidth =9069
                    LayoutCachedHeight =2740
                    WebImagePaddingLeft =3
                    WebImagePaddingTop =3
                    WebImagePaddingRight =2
                    WebImagePaddingBottom =2
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =93
                    Left =6259
                    Top =2102
                    Width =1393
                    Height =638
                    TabIndex =7
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

                    LayoutCachedLeft =6259
                    LayoutCachedTop =2102
                    LayoutCachedWidth =7652
                    LayoutCachedHeight =2740
                    WebImagePaddingLeft =3
                    WebImagePaddingTop =3
                    WebImagePaddingRight =2
                    WebImagePaddingBottom =2
                    Overlaps =1
                End
                Begin Label
                    OverlapFlags =223
                    Left =566
                    Top =2716
                    Width =8487
                    Height =735
                    FontSize =14
                    ForeColor =5026082
                    Name ="Status"
                    Caption ="6.2.2022 13.46.59 - Maksujen korjaus valmis"
                    LayoutCachedLeft =566
                    LayoutCachedTop =2716
                    LayoutCachedWidth =9053
                    LayoutCachedHeight =3451
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =177
                    TextFontFamily =0
                    IMESentenceMode =3
                    Left =5550
                    Top =3519
                    Width =1744
                    Height =300
                    TabIndex =10
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Puumerkki"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Calibri"
                    GridlineColor =10921638

                    LayoutCachedLeft =5550
                    LayoutCachedTop =3519
                    LayoutCachedWidth =7294
                    LayoutCachedHeight =3819
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3732
                            Top =3543
                            Width =1656
                            Height =300
                            Name ="Selite156"
                            Caption ="Puumerkki + enter:"
                            LayoutCachedLeft =3732
                            LayoutCachedTop =3543
                            LayoutCachedWidth =5388
                            LayoutCachedHeight =3843
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =93
                    Left =564
                    Top =4116
                    Width =1332
                    Height =588
                    TabIndex =11
                    Name ="MuokkaaLatauksia"
                    Caption ="Muokkaa latauksia"
                    OnClick ="[Event Procedure]"
                    GroupTable =6

                    LayoutCachedLeft =564
                    LayoutCachedTop =4116
                    LayoutCachedWidth =1896
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
                    OverlapFlags =93
                    Left =1968
                    Top =4116
                    Width =1212
                    Height =588
                    TabIndex =12
                    Name ="MuokkaaMaksuja"
                    Caption ="Muokkaa maksuja"
                    OnClick ="[Event Procedure]"
                    GroupTable =6

                    LayoutCachedLeft =1968
                    LayoutCachedTop =4116
                    LayoutCachedWidth =3180
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
                    OverlapFlags =93
                    Left =6897
                    Top =4251
                    Width =1464
                    Height =473
                    FontSize =16
                    FontWeight =700
                    BackColor =8435191
                    Name ="Raportit"
                    Caption ="Aikaväli:"
                    LayoutCachedLeft =6897
                    LayoutCachedTop =4251
                    LayoutCachedWidth =8361
                    LayoutCachedHeight =4724
                End
                Begin ToggleButton
                    OverlapFlags =215
                    TextFontCharSet =177
                    TextFontFamily =0
                    Left =472
                    Top =3425
                    Width =2997
                    Height =560
                    TabIndex =9
                    ForeColor =4210752
                    Name ="KorjaaTietoja"
                    Caption ="Admin -moodi"
                    FontName ="Calibri"
                    OnClick ="[Event Procedure]"
                    GridlineColor =10921638

                    LayoutCachedLeft =472
                    LayoutCachedTop =3425
                    LayoutCachedWidth =3469
                    LayoutCachedHeight =3985
                    BackColor =62207
                    BackThemeColorIndex =-1
                    BackTint =100.0
                    BorderColor =15123357
                    HoverColor =15652797
                    PressedColor =11957550
                    HoverForeColor =4210752
                    WebImagePaddingLeft =3
                    WebImagePaddingTop =2
                    WebImagePaddingRight =3
                    WebImagePaddingBottom =3
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =93
                    Left =3252
                    Top =4116
                    Width =1236
                    Height =588
                    TabIndex =13
                    Name ="Hinnat"
                    Caption ="Muokkaa korttihintoja"
                    OnClick ="[Event Procedure]"
                    GroupTable =6

                    LayoutCachedLeft =3252
                    LayoutCachedTop =4116
                    LayoutCachedWidth =4488
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
                    OverlapFlags =93
                    Left =1872
                    Top =4908
                    Width =1224
                    Height =804
                    TabIndex =17
                    Name ="RaporttiMaksamatta"
                    Caption ="Tarkasta maksamattomat kortit"
                    OnClick ="[Event Procedure]"
                    GroupTable =1

                    LayoutCachedLeft =1872
                    LayoutCachedTop =4908
                    LayoutCachedWidth =3096
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
                    OverlapFlags =93
                    Left =1872
                    Top =5784
                    Width =1224
                    Height =804
                    TabIndex =20
                    Name ="Historia"
                    Caption ="Sovelluksen täysi historia"
                    OnClick ="[Event Procedure]"
                    GroupTable =1

                    LayoutCachedLeft =1872
                    LayoutCachedTop =5784
                    LayoutCachedWidth =3096
                    LayoutCachedHeight =6588
                    RowStart =1
                    RowEnd =1
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
                Begin TextBox
                    OverlapFlags =223
                    TextFontCharSet =177
                    TextAlign =1
                    TextFontFamily =0
                    IMESentenceMode =3
                    Left =7181
                    Top =4889
                    Width =1961
                    Height =300
                    TabIndex =15
                    BackColor =8435191
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="RaportitAlku"
                    Format ="Short Date"
                    DefaultValue ="=DateSerial(Year(Now()),Month(Now())-6,Day(Now()))"
                    FontName ="Calibri"
                    GridlineColor =10921638

                    LayoutCachedLeft =7181
                    LayoutCachedTop =4889
                    LayoutCachedWidth =9142
                    LayoutCachedHeight =5189
                    BackThemeColorIndex =-1
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =1
                            Left =6354
                            Top =4889
                            Width =1284
                            Height =300
                            FontWeight =700
                            Name ="Selite166"
                            Caption ="Alkaen"
                            LayoutCachedLeft =6354
                            LayoutCachedTop =4889
                            LayoutCachedWidth =7638
                            LayoutCachedHeight =5189
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =223
                    TextFontCharSet =177
                    TextFontFamily =0
                    IMESentenceMode =3
                    Left =7180
                    Top =5291
                    Width =1961
                    Height =300
                    TabIndex =29
                    BackColor =8435191
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="RaportitLoppu"
                    Format ="Short Date"
                    DefaultValue ="=Date()"
                    FontName ="Calibri"
                    GridlineColor =10921638

                    LayoutCachedLeft =7180
                    LayoutCachedTop =5291
                    LayoutCachedWidth =9141
                    LayoutCachedHeight =5591
                    BackThemeColorIndex =-1
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =6307
                            Top =5291
                            Width =1284
                            Height =300
                            FontWeight =700
                            Name ="Selite168"
                            Caption ="Loppuen"
                            LayoutCachedLeft =6307
                            LayoutCachedTop =5291
                            LayoutCachedWidth =7591
                            LayoutCachedHeight =5591
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =93
                    Left =3168
                    Top =4908
                    Width =1392
                    Height =804
                    TabIndex =18
                    Name ="LatauksetKaikki"
                    Caption ="Korttilatausten kokonais-raportti"
                    OnClick ="[Event Procedure]"
                    GroupTable =1

                    LayoutCachedLeft =3168
                    LayoutCachedTop =4908
                    LayoutCachedWidth =4560
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
                    OverlapFlags =93
                    Left =588
                    Top =4908
                    Width =1212
                    Height =804
                    TabIndex =16
                    Name ="KortinTapahtumat"
                    Caption ="Kortin tapahtumat"
                    OnClick ="[Event Procedure]"
                    GroupTable =1

                    LayoutCachedLeft =588
                    LayoutCachedTop =4908
                    LayoutCachedWidth =1800
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
                    OverlapFlags =93
                    Left =3168
                    Top =5784
                    Width =1392
                    Height =804
                    TabIndex =21
                    Name ="ListaaEdustusj"
                    Caption ="Listaa edustusjäsenet"
                    OnClick ="[Event Procedure]"
                    GroupTable =1

                    LayoutCachedLeft =3168
                    LayoutCachedTop =5784
                    LayoutCachedWidth =4560
                    LayoutCachedHeight =6588
                    RowStart =1
                    RowEnd =1
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
                    OverlapFlags =93
                    Left =588
                    Top =6660
                    Width =1212
                    Height =804
                    TabIndex =24
                    Name ="IlmaiseksiLadattavat"
                    Caption ="Listaa ilmaiseksi ladattavat"
                    OnClick ="[Event Procedure]"
                    GroupTable =1

                    LayoutCachedLeft =588
                    LayoutCachedTop =6660
                    LayoutCachedWidth =1800
                    LayoutCachedHeight =7464
                    RowStart =2
                    RowEnd =2
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
                    OverlapFlags =93
                    Left =4560
                    Top =4116
                    Width =1248
                    Height =588
                    TabIndex =14
                    Name ="PaymentMethods"
                    Caption ="Muokkaa maksutapoja"
                    OnClick ="[Event Procedure]"
                    GroupTable =6

                    LayoutCachedLeft =4560
                    LayoutCachedTop =4116
                    LayoutCachedWidth =5808
                    LayoutCachedHeight =4704
                    ColumnStart =3
                    ColumnEnd =3
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
                    OverlapFlags =93
                    Left =588
                    Top =5784
                    Width =1212
                    Height =804
                    TabIndex =19
                    Name ="ListaaKaikkiMaksut"
                    Caption ="Listaa kaikki maksut"
                    OnClick ="[Event Procedure]"
                    GroupTable =1

                    LayoutCachedLeft =588
                    LayoutCachedTop =5784
                    LayoutCachedWidth =1800
                    LayoutCachedHeight =6588
                    RowStart =1
                    RowEnd =1
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
                Begin EmptyCell
                    Left =4632
                    Top =4908
                    Height =804
                    Name ="TyhjäSolu244"
                    GroupTable =1
                    GridlineColor =10921638
                    LayoutCachedLeft =4632
                    LayoutCachedTop =4908
                    LayoutCachedWidth =6072
                    LayoutCachedHeight =5712
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin CommandButton
                    OverlapFlags =93
                    Left =1872
                    Top =6660
                    Width =1224
                    Height =804
                    TabIndex =25
                    Name ="kortitIlmanOmistajaa"
                    Caption ="Aktiiviset kortit ilman omistajaa"
                    OnClick ="[Event Procedure]"
                    GroupTable =1

                    LayoutCachedLeft =1872
                    LayoutCachedTop =6660
                    LayoutCachedWidth =3096
                    LayoutCachedHeight =7464
                    RowStart =2
                    RowEnd =2
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
                    OverlapFlags =93
                    Left =3168
                    Top =6660
                    Width =1392
                    Height =804
                    TabIndex =26
                    Name ="maksettuEiLadattu"
                    Caption ="Tarkasta lataamattomat kortit"
                    OnClick ="[Event Procedure]"
                    GroupTable =1

                    LayoutCachedLeft =3168
                    LayoutCachedTop =6660
                    LayoutCachedWidth =4560
                    LayoutCachedHeight =7464
                    RowStart =2
                    RowEnd =2
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
                    OverlapFlags =93
                    Left =4632
                    Top =6660
                    Height =804
                    TabIndex =27
                    Name ="otaVarmuuskopio"
                    Caption ="Ota varmuuskopio"
                    OnClick ="[Event Procedure]"
                    GroupTable =1

                    LayoutCachedLeft =4632
                    LayoutCachedTop =6660
                    LayoutCachedWidth =6072
                    LayoutCachedHeight =7464
                    RowStart =2
                    RowEnd =2
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
                    OverlapFlags =93
                    Left =6144
                    Top =6660
                    Height =804
                    TabIndex =28
                    Name ="korttiTilastot"
                    Caption ="Näytä korttitilastot"
                    OnClick ="[Event Procedure]"
                    GroupTable =1

                    LayoutCachedLeft =6144
                    LayoutCachedTop =6660
                    LayoutCachedWidth =7584
                    LayoutCachedHeight =7464
                    RowStart =2
                    RowEnd =2
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
                Begin EmptyCell
                    Left =6144
                    Top =4908
                    Height =804
                    Name ="TyhjäSolu266"
                    GroupTable =1
                    GridlineColor =10921638
                    LayoutCachedLeft =6144
                    LayoutCachedTop =4908
                    LayoutCachedWidth =7584
                    LayoutCachedHeight =5712
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin CommandButton
                    OverlapFlags =93
                    Left =4632
                    Top =5784
                    Height =804
                    TabIndex =22
                    Name ="MaksutPerJasen"
                    Caption ="Listaa maksut per jäsen"
                    OnClick ="[Event Procedure]"
                    GroupTable =1

                    LayoutCachedLeft =4632
                    LayoutCachedTop =5784
                    LayoutCachedWidth =6072
                    LayoutCachedHeight =6588
                    RowStart =1
                    RowEnd =1
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
                    OverlapFlags =93
                    Left =6144
                    Top =5784
                    Height =804
                    TabIndex =23
                    Name ="voimassaOlleetKorit"
                    Caption ="Aikavälillä voimassa olleet kortit"
                    OnClick ="[Event Procedure]"
                    GroupTable =1

                    LayoutCachedLeft =6144
                    LayoutCachedTop =5784
                    LayoutCachedWidth =7584
                    LayoutCachedHeight =6588
                    RowStart =1
                    RowEnd =1
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
                Begin CommandButton
                    OverlapFlags =85
                    Left =6236
                    Top =1608
                    Width =2758
                    Height =385
                    TabIndex =5
                    Name ="KorvaaRikkinainenKortti"
                    Caption ="Korvaa rikkinäinen kortti"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =6236
                    LayoutCachedTop =1608
                    LayoutCachedWidth =8994
                    LayoutCachedHeight =1993
                    WebImagePaddingLeft =3
                    WebImagePaddingTop =3
                    WebImagePaddingRight =2
                    WebImagePaddingBottom =2
                End
                Begin CommandButton
                    OverlapFlags =93
                    Left =7632
                    Top =5760
                    Height =804
                    TabIndex =30
                    Name ="umpeenMenevatRaportti"
                    Caption ="Aikavälillä umpeen menevät kortit"
                    OnClick ="[Event Procedure]"
                    GroupTable =7

                    LayoutCachedLeft =7632
                    LayoutCachedTop =5760
                    LayoutCachedWidth =9072
                    LayoutCachedHeight =6564
                    LayoutGroup =3
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =8435191
                    HoverColor =8435191
                    PressedColor =8435191
                    GroupTable =7
                    WebImagePaddingLeft =3
                    WebImagePaddingTop =3
                    WebImagePaddingRight =2
                    WebImagePaddingBottom =2
                    Overlaps =1
                End
                Begin Image
                    PictureType =2
                    Left =519
                    Top =4039
                    Width =8794
                    Height =3792
                    Name ="Bulldog"
                    Picture ="bulldog_pienempi"

                    LayoutCachedLeft =519
                    LayoutCachedTop =4039
                    LayoutCachedWidth =9313
                    LayoutCachedHeight =7831
                    TabIndex =31
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



Private Sub Form_Close()
    Dim succs
    succs = Common.DoBackup(1)  'do not back up unless latest is over 1 days old!
    succs = Common.WriteStats 'write stats to own table
    succs = Common.SaveToLog("Sovellus suljettu")
End Sub

Private Sub Form_GotFocus()
    [Form_Tervetuloa].Refresh
    
    'DoCmd.Requery "[Form].[Tervetuloa]"
    
End Sub


Private Sub Form_Open(Cancel As Integer)
    Dim succs
    succs = Common.EnableDisableButtons()
    succs = Common.SendMessageToMainScreen("Tervetuloa!")
    succs = Common.SaveToLog("Jäsenrekisteri avattiin")
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

Private Sub kortitIlmanOmistajaa_Click()
    Dim succs
    succs = Common.SaveToLog([Form_Tervetuloa].Puumerkki.Value & " avasi aktiiviset kortit ilman omistajaa -raportin")
    DoCmd.OpenReport "VoimassaolevatIlmanOmistajaa", acViewPreview
    
End Sub

Private Sub korttiTilastot_Click()
    Dim succs
    succs = Common.SaveToLog([Form_Tervetuloa].Puumerkki.Value & " avasi korttitilasto-raportin")
    DoCmd.OpenReport "Korttitilasto", acViewPreview
    
End Sub

Private Sub Korttivalinta_AfterUpdate()
    Common.EnableDisableButtons
    
    Dim criteria As String
    
    criteria = "Kortti = '" & [Form_Tervetuloa].Korttivalinta.Value & "'"
    
    Dim user As Integer
    user = Common.FetchGeneralID("Kortit", "Omistaja", criteria)
    'MsgBox (user)
    ''[Form_Tervetuloa].Yhteystietovalinta.
    '[Form_Tervetuloa].Yhteystietovalinta.Value = Common.FetchGeneralID("Kortit", "Omistaja", criteria)
    [Form_Tervetuloa].Yhteystietovalinta = Common.FetchGeneralID("Kortit", "Omistaja", criteria)
    [Form_Tervetuloa].Yhteystietovalinta.Requery
    
    [Form_Tervetuloa].Refresh
    [Form_Tervetuloa].Yhteystietovalinta.SetFocus
    'MsgBox ([Form_Tervetuloa].Yhteystietovalinta.Value)
End Sub


Private Sub Korttivalinta_GotFocus()
    Common.EnableDisableButtons
    
End Sub

Private Sub KorvaaRikkinainenKortti_Click()
    DoCmd.OpenForm "KorvaaKortti"
    
    'MsgBox ("Humpan juoni: " & vbNewLine _
    '& "1. tehdään uusi kortti " & vbNewLine _
    '& "2. tehdään uudelle kortille lataus " & vbNewLine _
    '& "3. korjataan kortin maksuihin liittyvät jutut " & vbNewLine _
    '& "4. merkataan vanha kortti rikkinäiseksi/kadonneeksi." & vbNewLine _
    '& "Täytä vaan ruutuja sitä mukaa, kun niitä hyppii ja kaikki menee hyvin!")
    'vai: varmista, että ruuduissa olevat tiedot on ok ja paina ok?!?
    
    
    
    'Dim oldCard As String
    'oldCard = Form_Tervetuloa.Korttivalinta.Value
    
    'Dim initials As String
    'initials = InputBox("Nimikirjaimet tai jotain vastaavaa", "Anna puumerkkisi", "NN")
    
    'Dim newCard As String
    'newCard = InputBox("Kortti muodossa 4 numeroa", "Anna uuden kortin numero", 1234)
    
    'Dim succs
    'succs = Common.SaveToLog(initials & " aloitti kortin " & oldCard & " korvaamisen kortilla " & newCard)
    'logitus jo tästä pisteestä asti ihan vaan siksi..
    
    'If MsgBox("Tehdään uusi kortti " & newCard & vbNewLine & "Meneehän varmasti oikein?", vbYesNo) = vbNo Then Exit Sub

    
    'DoCmd.OpenForm "LisaaKortinLinkitys"
    'Form_LisaaKortinLinkitys.Korttinro.Value = newCard
    'Form_LisaaKortinLinkitys.Puumerkki.Value = initials
    'Form_LisaaKortinLinkitys.Refresh
    'Form_LisaaKortinLinkitys.Puumerkki.Visible = True
    'Form_LisaaKortinLinkitys.Linkita.Visible = True
    'Form_LisaaKortinLinkitys.Linkita_Click
    
    
    'If MsgBox("Siirretään kaikki lataukset ja maksut kortilta " & oldCard & " kortille " & newCard & vbNewLine & "Tätä on hankala korjata jälkikäteen, oletko aivan varma?", vbYesNo) = vbNo Then Exit Sub
    
    'Dim oldCardID As Integer
    'Dim newCardID As Integer
    
    'oldCardID = Common.FetchCardID(oldCard)
    'newCardID = Common.FetchCardID(newCard)
    
    'Dim Table As String
    'Table = "Lataukset"
    'Dim Values As String
    'Values = "Kortti=" & newCardID
    'Dim Target As String
    'Target = "Kortti=" & oldCardID
    'succs = Common.InsertOrUpdate(Table, Values, Target)
    'Table = "Maksut"
    'succs = Common.InsertOrUpdate(Table, Values, Target)
    
    'MsgBox ("Vanhan kortin lataukset ja suoritetut maksut siirretty uudelle kortille")
    
    'DoCmd.OpenForm "PoistaKortinLinkitys"
    'Form_PoistaKortinLinkitys.Puumerkki.Value = initials
    'Form_PoistaKortinLinkitys.discard.Value = True
    'Form_PoistaKortinLinkitys.Muistiinpano.Value = "Kortin korvaus"
    'Form_PoistaKortinLinkitys.Poista_Click
    
    'MsgBox ("Valmista tuli!")
    
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

Private Sub ListaaKaikkiMaksut_Click()
    Dim succs
    succus = Common.SaveToLog([Form_Tervetuloa].Puumerkki.Value & " pyysi maksulistauksen")
    DoCmd.OpenReport "ListaaMaksut", acViewPreview
End Sub

Private Sub Main_jäsenmuokkaus_Click()
    Dim succs
    succs = Common.SaveToLog("Yhteystietojen muokkaushommat aloitettu")
    'DoCmd.OpenForm ("Form_YhteystietojenMuokkaus")
    DoCmd.OpenForm "YhteystietojenMuokkaus"

End Sub

Private Sub maksettuEiLadattu_Click()
    Dim succs
    succs = Common.SaveToLog([Form_Tervetuloa].Puumerkki.Value & " pyysi listauksen maksetuista mutta ei vielä ladatuista korteista.")
    DoCmd.OpenReport "MaksettuEiLadattu", acViewPreview
    
End Sub

Private Sub MaksutPerJasen_Click()
    Dim succs
    succs = Common.SaveToLog([Form_Tervetuloa].Puumerkki.Value & " pyysi listauksen kaikista maksuista jäsenittäin.")
    DoCmd.OpenReport "MaksutPerJasen", acViewPreview
End Sub

Private Sub MuokkaaLatauksia_Click()
    Dim succs
    succs = Common.SaveToLog([Form_Tervetuloa].Puumerkki.Value & " aloitti kortin " & [Form_Tervetuloa].Korttivalinta.Value & " latausten muokkaamisen")
    DoCmd.OpenForm "KorjaaKortinLatauksia"
    succs = Common.SendMessageToMainScreen("Kortin " & [Form_Tervetuloa].Korttivalinta.Value & " latausmuokkaus aloitettu")
    
End Sub

Private Sub MuokkaaMaksuja_Click()
    Dim succs
    succs = Common.SaveToLog([Form_Tervetuloa].Puumerkki.Value & " aloitti kortin " & [Form_Tervetuloa].Korttivalinta.Value & " maksujen muokkaamisen")
    DoCmd.OpenForm "KorjaaKortinMaksuja"
    succs = Common.SendMessageToMainScreen("Kortin " & [Form_Tervetuloa].Korttivalinta.Value & " maksumuokkaus aloitettu")
    
End Sub


Private Sub otaVarmuuskopio_Click()
    Dim succs
    succs = Common.SaveToLog([Form_Tervetuloa].Puumerkki.Value & " käski ottaa varmuuskopion.")
    succs = Common.DoBackup(0) 'just to make sure you get the newest backup for the day..
End Sub

Private Sub PaymentMethods_Click()
    Dim succs
    succs = Common.SaveToLog([Form_Tervetuloa].Puumerkki.Value & " avasi maksutapojen muokkausikkunan.")
    DoCmd.OpenForm "MuokkaaMaksutapoja"
    succs = Common.SendMessageToMainScreen("Maksutapojen muokkaus aloitettu")
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

Public Sub Tyhjennä_Click()

    Form_Tervetuloa.Yhteystietovalinta = Null
    Form_Tervetuloa.Korttivalinta = Null
    Form_Tervetuloa.Paivita_korttiluettelo
    Common.EnableDisableButtons
    [Form_Tervetuloa].Refresh
    

End Sub

Private Sub umpeenMenevatRaportti_Click()
    Dim succs
    succs = Common.SaveToLog([Form_Tervetuloa].Puumerkki.Value & " avasi listauksen tietyllä ajanjaksolla umpeen menevistä korteista")
    DoCmd.OpenReport "LatausLoppumassaAjanjaksolla", acViewPreview
End Sub

Private Sub voimassaOlleetKorit_Click()
    Dim succs
    succs = Common.SaveToLog([Form_Tervetuloa].Puumerkki.Value & " avasi listauksen tietyllä ajanjaksolla voimassa olleista korteista")
    DoCmd.OpenReport "Kortit joissa latausta ajanjaksolla", acViewPreview
    
End Sub

Private Sub Yhteystietovalinta_AfterUpdate()

    'Form_Tervetuoa.Paivita_korttiluettelo()
    Form_Tervetuloa.Paivita_korttiluettelo
    Common.EnableDisableButtons
    'MsgBox ([Form_Tervetuloa].Yhteystietovalinta.Value)
    
End Sub


Private Sub Yhteystietovalinta_Change()
    Common.EnableDisableButtons
    
End Sub

Private Sub Yhteystietovalinta_GotFocus()
    Common.EnableDisableButtons
End Sub
