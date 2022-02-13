Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    Modal = NotDefault
    NavigationButtons = NotDefault
    CloseButton = NotDefault
    DividingLines = NotDefault
    DataEntry = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    AllowUpdating =2
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    PictureType =2
    GridY =10
    Width =6973
    DatasheetFontHeight =11
    ItemSuffix =381
    Left =4044
    Top =3468
    Right =17484
    Bottom =11712
    Picture ="bulldog_pienempi"
    RecSrcDt = Begin
        0x84756ccb9ec6e540
    End
    Caption ="Rekisteröi lataus"
    OnOpen ="[Event Procedure]"
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
        Begin Line
            BorderLineStyle =0
            Width =1701
            BorderThemeColorIndex =0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin Image
            BackStyle =0
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
        Begin OptionButton
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
            BorderThemeColorIndex =1
            BorderShade =65.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
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
        Begin EmptyCell
            Height =240
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin FormHeader
            Height =660
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
                    TextFontCharSet =177
                    TextAlign =1
                    TextFontFamily =0
                    Left =300
                    Top =36
                    Width =3576
                    Height =480
                    FontSize =18
                    Name ="Automaattinen_ylätunniste0"
                    Caption ="Rekisteröi lataus kortille"
                    FontName ="Calibri Light"
                    GroupTable =1
                    GridlineColor =10921638
                    HorizontalAnchor =2
                    LayoutCachedLeft =300
                    LayoutCachedTop =36
                    LayoutCachedWidth =3876
                    LayoutCachedHeight =516
                    LayoutGroup =1
                    ThemeFontIndex =0
                    BorderThemeColorIndex =2
                    BorderTint =100.0
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                    GroupTable =1
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =93
                    TextFontCharSet =177
                    TextAlign =3
                    TextFontFamily =0
                    IMESentenceMode =3
                    Left =5324
                    Top =56
                    Width =1503
                    Height =300
                    BorderColor =10921638
                    Name ="Automaattinen_päivämäärä"
                    ControlSource ="=Date()"
                    Format ="Short Date"
                    FontName ="Calibri"
                    GridlineColor =10921638

                    LayoutCachedLeft =5324
                    LayoutCachedTop =56
                    LayoutCachedWidth =6827
                    LayoutCachedHeight =356
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
                    IMESentenceMode =3
                    Left =5324
                    Top =340
                    Width =1503
                    Height =300
                    TabIndex =1
                    BorderColor =10921638
                    Name ="Automaattinen_aika"
                    ControlSource ="=Time()"
                    Format ="Long Time"
                    FontName ="Calibri"
                    GridlineColor =10921638

                    LayoutCachedLeft =5324
                    LayoutCachedTop =340
                    LayoutCachedWidth =6827
                    LayoutCachedHeight =640
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                End
            End
        End
        Begin Section
            Height =7842
            Name ="Tiedot"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin Rectangle
                    BackStyle =1
                    OverlapFlags =93
                    Left =283
                    Top =113
                    Width =6123
                    Height =6066
                    BackColor =14277081
                    BorderColor =10921638
                    Name ="Ruutu379"
                    GridlineColor =10921638
                    LayoutCachedLeft =283
                    LayoutCachedTop =113
                    LayoutCachedWidth =6406
                    LayoutCachedHeight =6179
                    BackShade =85.0
                End
                Begin TextBox
                    OverlapFlags =215
                    TextFontCharSet =177
                    TextAlign =1
                    TextFontFamily =0
                    IMESentenceMode =3
                    Left =2412
                    Top =108
                    Width =3780
                    Height =336
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Kortti"
                    ControlSource ="=[Forms]![Tervetuloa]![Korttivalinta]"
                    FontName ="Calibri"
                    GroupTable =2
                    BottomPadding =150
                    GridlineColor =10921638

                    LayoutCachedLeft =2412
                    LayoutCachedTop =108
                    LayoutCachedWidth =6192
                    LayoutCachedHeight =444
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =2
                    GroupTable =2
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =215
                            TextFontCharSet =177
                            TextAlign =1
                            TextFontFamily =0
                            Left =336
                            Top =108
                            Width =1980
                            Height =336
                            BackColor =15921906
                            BorderColor =8355711
                            ForeColor =6710886
                            Name ="Selite5"
                            Caption ="Kortti:"
                            FontName ="Calibri"
                            GroupTable =2
                            BottomPadding =150
                            GridlineColor =10921638
                            LayoutCachedLeft =336
                            LayoutCachedTop =108
                            LayoutCachedWidth =2316
                            LayoutCachedHeight =444
                            LayoutGroup =2
                            BackShade =95.0
                            GroupTable =2
                        End
                    End
                End
                Begin TextBox
                    BorderWidth =3
                    OverlapFlags =215
                    TextFontCharSet =177
                    TextAlign =1
                    TextFontFamily =0
                    IMESentenceMode =3
                    Left =2412
                    Top =2844
                    Width =3780
                    Height =336
                    TabIndex =2
                    ForeColor =4210752
                    Name ="Hinta"
                    Format ="#,##0.00 €;-#,##0.00 €"
                    FontName ="Calibri"
                    OnChange ="[Event Procedure]"
                    GroupTable =2
                    BottomPadding =150
                    GridlineColor =10921638

                    LayoutCachedLeft =2412
                    LayoutCachedTop =2844
                    LayoutCachedWidth =6192
                    LayoutCachedHeight =3180
                    RowStart =6
                    RowEnd =6
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =2
                    BorderThemeColorIndex =5
                    BorderShade =100.0
                    GroupTable =2
                    CurrencySymbol ="€"
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =215
                            TextFontCharSet =177
                            TextAlign =1
                            TextFontFamily =0
                            Left =336
                            Top =2844
                            Width =1980
                            Height =336
                            BorderColor =8355711
                            ForeColor =6710886
                            Name ="Selite75"
                            Caption ="Hinta"
                            FontName ="Calibri"
                            GroupTable =2
                            BottomPadding =150
                            GridlineColor =10921638
                            LayoutCachedLeft =336
                            LayoutCachedTop =2844
                            LayoutCachedWidth =2316
                            LayoutCachedHeight =3180
                            RowStart =6
                            RowEnd =6
                            LayoutGroup =2
                            GroupTable =2
                        End
                    End
                End
                Begin TextBox
                    BorderWidth =3
                    OverlapFlags =215
                    TextFontCharSet =177
                    TextAlign =1
                    TextFontFamily =0
                    IMESentenceMode =3
                    Left =2412
                    Top =3396
                    Width =3780
                    Height =336
                    TabIndex =3
                    ForeColor =4210752
                    Name ="Voimassa"
                    Format ="Short Date"
                    DefaultValue ="=Date()"
                    FontName ="Calibri"
                    OnChange ="[Event Procedure]"
                    GroupTable =2
                    BottomPadding =150
                    GridlineColor =10921638

                    LayoutCachedLeft =2412
                    LayoutCachedTop =3396
                    LayoutCachedWidth =6192
                    LayoutCachedHeight =3732
                    RowStart =7
                    RowEnd =7
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =2
                    BorderThemeColorIndex =5
                    BorderShade =100.0
                    GroupTable =2
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =215
                            TextFontCharSet =177
                            TextAlign =1
                            TextFontFamily =0
                            Left =336
                            Top =3396
                            Width =1980
                            Height =336
                            BorderColor =8355711
                            ForeColor =6710886
                            Name ="Selite90"
                            Caption ="Voimassa asti"
                            FontName ="Calibri"
                            GroupTable =2
                            BottomPadding =150
                            GridlineColor =10921638
                            LayoutCachedLeft =336
                            LayoutCachedTop =3396
                            LayoutCachedWidth =2316
                            LayoutCachedHeight =3732
                            RowStart =7
                            RowEnd =7
                            LayoutGroup =2
                            GroupTable =2
                        End
                    End
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =215
                    TextFontCharSet =177
                    TextAlign =1
                    TextFontFamily =0
                    Left =396
                    Top =4536
                    Width =5844
                    Height =612
                    BackColor =62207
                    BorderColor =8355711
                    ForeColor =6710886
                    Name ="Selite144"
                    Caption ="HUOM! Voimassaolo pyöristetään aina seuraavaan 15. päivään tai kuun loppuun hall"
                        "ituksen linjauksen mukaan!"
                    FontName ="Calibri"
                    GroupTable =13
                    BottomPadding =150
                    GridlineColor =10921638
                    LayoutCachedLeft =396
                    LayoutCachedTop =4536
                    LayoutCachedWidth =6240
                    LayoutCachedHeight =5148
                    LayoutGroup =3
                    BackThemeColorIndex =-1
                    GroupTable =13
                End
                Begin TextBox
                    BorderWidth =3
                    OverlapFlags =215
                    TextFontCharSet =177
                    TextAlign =1
                    TextFontFamily =0
                    IMESentenceMode =3
                    Left =2412
                    Top =2340
                    Width =3780
                    Height =288
                    TabIndex =1
                    ForeColor =4210752
                    Name ="Puumerkki"
                    FontName ="Calibri"
                    OnChange ="[Event Procedure]"
                    GroupTable =2
                    BottomPadding =150
                    GridlineColor =10921638

                    LayoutCachedLeft =2412
                    LayoutCachedTop =2340
                    LayoutCachedWidth =6192
                    LayoutCachedHeight =2628
                    RowStart =5
                    RowEnd =5
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =2
                    BorderThemeColorIndex =5
                    BorderShade =100.0
                    GroupTable =2
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =215
                            TextFontCharSet =177
                            TextAlign =1
                            TextFontFamily =0
                            Left =336
                            Top =2340
                            Width =1980
                            Height =288
                            Name ="Selite176"
                            Caption ="Lataajan puumerkki"
                            FontName ="Calibri"
                            GroupTable =2
                            BottomPadding =150
                            GridlineColor =10921638
                            LayoutCachedLeft =336
                            LayoutCachedTop =2340
                            LayoutCachedWidth =2316
                            LayoutCachedHeight =2628
                            RowStart =5
                            RowEnd =5
                            LayoutGroup =2
                            BorderThemeColorIndex =-1
                            ForeThemeColorIndex =-1
                            GroupTable =2
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =215
                    TextFontCharSet =177
                    TextFontFamily =0
                    Left =396
                    Top =6084
                    Width =5844
                    Height =696
                    TabIndex =8
                    ForeColor =4210752
                    Name ="Save"
                    Caption ="Tallenna"
                    OnClick ="[Event Procedure]"
                    FontName ="Calibri"
                    GroupTable =13
                    BottomPadding =150
                    GridlineColor =10921638

                    LayoutCachedLeft =396
                    LayoutCachedTop =6084
                    LayoutCachedWidth =6240
                    LayoutCachedHeight =6780
                    RowStart =2
                    RowEnd =2
                    LayoutGroup =3
                    BackColor =15123357
                    BorderColor =15123357
                    HoverColor =15652797
                    PressedColor =11957550
                    HoverForeColor =4210752
                    PressedForeColor =4210752
                    GroupTable =13
                    WebImagePaddingLeft =3
                    WebImagePaddingTop =3
                    WebImagePaddingRight =2
                    WebImagePaddingBottom =12
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontCharSet =177
                    TextFontFamily =0
                    Left =396
                    Top =6972
                    Width =5844
                    Height =720
                    TabIndex =9
                    ForeColor =4210752
                    Name ="ragequit"
                    Caption ="Sulje"
                    FontName ="Calibri"
                    GroupTable =13
                    BottomPadding =150
                    GridlineColor =10921638
                    OnClickEmMacro = Begin
                        Version =196611
                        ColumnsShown =0
                        Begin
                            Action ="Close"
                            Argument ="2"
                            Argument ="RekisteroiLataus"
                            Argument ="2"
                        End
                        Begin
                            Comment ="_AXL:<?xml version=\"1.0\" encoding=\"UTF-16\" standalone=\"no\"?>\015\012<UserI"
                                "nterfaceMacro For=\"ragequit\" Event=\"OnClick\" xmlns=\"http://schemas.microsof"
                                "t.com/office/accessservices/2009/11/application\"><Statements><Action Name=\"Clo"
                                "seWindow\"><Argument Name=\"Object"
                        End
                        Begin
                            Comment ="_AXL:Type\">Form</Argument><Argument Name=\"ObjectName\">RekisteroiLataus</Argum"
                                "ent><Argument Name=\"Save\">No</Argument></Action></Statements></UserInterfaceMa"
                                "cro>"
                        End
                    End

                    LayoutCachedLeft =396
                    LayoutCachedTop =6972
                    LayoutCachedWidth =6240
                    LayoutCachedHeight =7692
                    RowStart =3
                    RowEnd =3
                    LayoutGroup =3
                    BackColor =15123357
                    BorderColor =15123357
                    HoverColor =15652797
                    PressedColor =11957550
                    HoverForeColor =4210752
                    PressedForeColor =4210752
                    GroupTable =13
                    WebImagePaddingLeft =3
                    WebImagePaddingTop =3
                    WebImagePaddingRight =2
                    WebImagePaddingBottom =12
                    Overlaps =1
                End
                Begin EmptyCell
                    Left =336
                    Top =1884
                    Width =1980
                    Name ="TyhjäSolu301"
                    GroupTable =2
                    BottomPadding =150
                    GridlineColor =10921638
                    LayoutCachedLeft =336
                    LayoutCachedTop =1884
                    LayoutCachedWidth =2316
                    LayoutCachedHeight =2124
                    RowStart =4
                    RowEnd =4
                    LayoutGroup =2
                    GroupTable =2
                End
                Begin EmptyCell
                    Left =2412
                    Top =1884
                    Width =3780
                    Name ="TyhjäSolu302"
                    GroupTable =2
                    BottomPadding =150
                    GridlineColor =10921638
                    LayoutCachedLeft =2412
                    LayoutCachedTop =1884
                    LayoutCachedWidth =6192
                    LayoutCachedHeight =2124
                    RowStart =4
                    RowEnd =4
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =2
                    GroupTable =2
                End
                Begin EmptyCell
                    Left =336
                    Top =1464
                    Width =1980
                    Name ="TyhjäSolu303"
                    GroupTable =2
                    BottomPadding =150
                    GridlineColor =10921638
                    LayoutCachedLeft =336
                    LayoutCachedTop =1464
                    LayoutCachedWidth =2316
                    LayoutCachedHeight =1704
                    RowStart =3
                    RowEnd =3
                    LayoutGroup =2
                    GroupTable =2
                End
                Begin EmptyCell
                    Left =2412
                    Top =1464
                    Width =3780
                    Name ="TyhjäSolu304"
                    GroupTable =2
                    BottomPadding =150
                    GridlineColor =10921638
                    LayoutCachedLeft =2412
                    LayoutCachedTop =1464
                    LayoutCachedWidth =6192
                    LayoutCachedHeight =1704
                    RowStart =3
                    RowEnd =3
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =2
                    GroupTable =2
                End
                Begin EmptyCell
                    Left =336
                    Top =1044
                    Width =1980
                    Name ="TyhjäSolu305"
                    GroupTable =2
                    BottomPadding =150
                    GridlineColor =10921638
                    LayoutCachedLeft =336
                    LayoutCachedTop =1044
                    LayoutCachedWidth =2316
                    LayoutCachedHeight =1284
                    RowStart =2
                    RowEnd =2
                    LayoutGroup =2
                    GroupTable =2
                End
                Begin EmptyCell
                    Left =2412
                    Top =1044
                    Width =3780
                    Name ="TyhjäSolu306"
                    GroupTable =2
                    BottomPadding =150
                    GridlineColor =10921638
                    LayoutCachedLeft =2412
                    LayoutCachedTop =1044
                    LayoutCachedWidth =6192
                    LayoutCachedHeight =1284
                    RowStart =2
                    RowEnd =2
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =2
                    GroupTable =2
                End
                Begin OptionGroup
                    SpecialEffect =0
                    BackStyle =1
                    BorderWidth =2
                    OverlapFlags =223
                    Left =453
                    Top =630
                    Width =3384
                    Height =1351
                    TabIndex =5
                    BackColor =15921906
                    Name ="Valinta"
                    OnClick ="[Event Procedure]"
                    GridlineColor =10921638

                    LayoutCachedLeft =453
                    LayoutCachedTop =630
                    LayoutCachedWidth =3837
                    LayoutCachedHeight =1981
                    BackShade =95.0
                    BorderThemeColorIndex =5
                    BorderShade =100.0
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =223
                            TextFontCharSet =177
                            TextFontFamily =0
                            Left =573
                            Top =510
                            Width =1116
                            Height =300
                            BorderColor =8355711
                            ForeColor =6710886
                            Name ="Selite282"
                            Caption ="Korttityyppi"
                            FontName ="Calibri"
                            GridlineColor =10921638
                            LayoutCachedLeft =573
                            LayoutCachedTop =510
                            LayoutCachedWidth =1689
                            LayoutCachedHeight =810
                        End
                        Begin OptionButton
                            OverlapFlags =87
                            Left =510
                            Top =937
                            OptionValue =1
                            BorderColor =10921638
                            Name ="KKKortti"
                            GridlineColor =10921638

                            LayoutCachedLeft =510
                            LayoutCachedTop =937
                            LayoutCachedWidth =770
                            LayoutCachedHeight =1177
                            Begin
                                Begin Label
                                    OverlapFlags =255
                                    TextFontCharSet =177
                                    TextFontFamily =0
                                    Left =740
                                    Top =907
                                    Width =1056
                                    Height =336
                                    BorderColor =8355711
                                    ForeColor =6710886
                                    Name ="Selite284"
                                    Caption ="Normaali"
                                    FontName ="Calibri"
                                    GridlineColor =10921638
                                    LayoutCachedLeft =740
                                    LayoutCachedTop =907
                                    LayoutCachedWidth =1796
                                    LayoutCachedHeight =1243
                                End
                            End
                        End
                        Begin OptionButton
                            OverlapFlags =87
                            Left =2040
                            Top =937
                            TabIndex =1
                            OptionValue =4
                            BorderColor =10921638
                            Name ="KERKortti"
                            GridlineColor =10921638

                            LayoutCachedLeft =2040
                            LayoutCachedTop =937
                            LayoutCachedWidth =2300
                            LayoutCachedHeight =1177
                            Begin
                                Begin Label
                                    OverlapFlags =247
                                    TextFontCharSet =177
                                    TextFontFamily =0
                                    Left =2270
                                    Top =907
                                    Width =1056
                                    Height =336
                                    BorderColor =8355711
                                    ForeColor =6710886
                                    Name ="Selite290"
                                    Caption ="Kertakortti"
                                    FontName ="Calibri"
                                    GridlineColor =10921638
                                    LayoutCachedLeft =2270
                                    LayoutCachedTop =907
                                    LayoutCachedWidth =3326
                                    LayoutCachedHeight =1243
                                End
                            End
                        End
                        Begin OptionButton
                            OverlapFlags =87
                            Left =510
                            Top =1299
                            Width =284
                            Height =252
                            TabIndex =2
                            OptionValue =2
                            BorderColor =10921638
                            Name ="APKortti"
                            GridlineColor =10921638

                            LayoutCachedLeft =510
                            LayoutCachedTop =1299
                            LayoutCachedWidth =794
                            LayoutCachedHeight =1551
                            Begin
                                Begin Label
                                    OverlapFlags =255
                                    TextFontCharSet =177
                                    TextFontFamily =0
                                    Left =736
                                    Top =1247
                                    Width =1056
                                    Height =336
                                    BorderColor =8355711
                                    ForeColor =6710886
                                    Name ="Selite286"
                                    Caption ="Aamupäivä"
                                    FontName ="Calibri"
                                    GridlineColor =10921638
                                    LayoutCachedLeft =736
                                    LayoutCachedTop =1247
                                    LayoutCachedWidth =1792
                                    LayoutCachedHeight =1583
                                End
                            End
                        End
                        Begin OptionButton
                            OverlapFlags =87
                            Left =2040
                            Top =1334
                            TabIndex =3
                            OptionValue =5
                            BorderColor =10921638
                            Name ="MUUKortti"
                            GridlineColor =10921638

                            LayoutCachedLeft =2040
                            LayoutCachedTop =1334
                            LayoutCachedWidth =2300
                            LayoutCachedHeight =1574
                            Begin
                                Begin Label
                                    OverlapFlags =247
                                    TextFontCharSet =177
                                    TextFontFamily =0
                                    Left =2270
                                    Top =1304
                                    Width =1056
                                    Height =336
                                    BorderColor =8355711
                                    ForeColor =6710886
                                    Name ="Selite292"
                                    Caption ="Muu"
                                    FontName ="Calibri"
                                    GridlineColor =10921638
                                    LayoutCachedLeft =2270
                                    LayoutCachedTop =1304
                                    LayoutCachedWidth =3326
                                    LayoutCachedHeight =1640
                                End
                            End
                        End
                        Begin OptionButton
                            OverlapFlags =87
                            Left =510
                            Top =1618
                            TabIndex =4
                            OptionValue =3
                            BorderColor =10921638
                            Name ="OPISKortti"
                            GridlineColor =10921638

                            LayoutCachedLeft =510
                            LayoutCachedTop =1618
                            LayoutCachedWidth =770
                            LayoutCachedHeight =1858
                            Begin
                                Begin Label
                                    OverlapFlags =247
                                    TextFontCharSet =177
                                    TextFontFamily =0
                                    Left =740
                                    Top =1588
                                    Width =1056
                                    Height =336
                                    BorderColor =8355711
                                    ForeColor =6710886
                                    Name ="Selite288"
                                    Caption ="Opiskelija"
                                    FontName ="Calibri"
                                    GridlineColor =10921638
                                    LayoutCachedLeft =740
                                    LayoutCachedTop =1588
                                    LayoutCachedWidth =1796
                                    LayoutCachedHeight =1924
                                End
                            End
                        End
                    End
                End
                Begin EmptyCell
                    Left =336
                    Top =624
                    Width =1980
                    Name ="TyhjäSolu323"
                    GroupTable =2
                    BottomPadding =150
                    GridlineColor =10921638
                    LayoutCachedLeft =336
                    LayoutCachedTop =624
                    LayoutCachedWidth =2316
                    LayoutCachedHeight =864
                    RowStart =1
                    RowEnd =1
                    LayoutGroup =2
                    GroupTable =2
                End
                Begin EmptyCell
                    Left =2412
                    Top =624
                    Width =3780
                    Name ="TyhjäSolu324"
                    GroupTable =2
                    BottomPadding =150
                    GridlineColor =10921638
                    LayoutCachedLeft =2412
                    LayoutCachedTop =624
                    LayoutCachedWidth =6192
                    LayoutCachedHeight =864
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =2
                    GroupTable =2
                End
                Begin ComboBox
                    Enabled = NotDefault
                    RowSourceTypeInt =1
                    BorderWidth =3
                    OverlapFlags =215
                    TextFontCharSet =177
                    TextFontFamily =0
                    IMESentenceMode =3
                    Left =3932
                    Top =1071
                    Width =1190
                    Height =300
                    TabIndex =6
                    Name ="KKmaara"
                    RowSourceType ="Value List"
                    RowSource ="\"18\";\"17\";\"16\";\"15\";\"14\";\"13\";\"12\";\"11\";\"10\";\"9\";\"8\";\"7\""
                        ";\"6\";\"5\";\"4\";\"3\";\"2\";\"1\""
                    DefaultValue ="\"6\""
                    FontName ="Calibri"
                    OnChange ="[Event Procedure]"
                    GridlineColor =10921638

                    LayoutCachedLeft =3932
                    LayoutCachedTop =1071
                    LayoutCachedWidth =5122
                    LayoutCachedHeight =1371
                    BorderThemeColorIndex =5
                    BorderShade =100.0
                    ForeThemeColorIndex =0
                    ForeTint =75.0
                    ForeShade =100.0
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =215
                            TextFontCharSet =177
                            TextFontFamily =0
                            Left =3932
                            Top =732
                            Width =1056
                            Height =276
                            BackColor =15921906
                            BorderColor =8355711
                            ForeColor =6710886
                            Name ="Selite340"
                            Caption ="Kuukautta"
                            FontName ="Calibri"
                            GridlineColor =10921638
                            LayoutCachedLeft =3932
                            LayoutCachedTop =732
                            LayoutCachedWidth =4988
                            LayoutCachedHeight =1008
                            BackShade =95.0
                        End
                    End
                End
                Begin ComboBox
                    Enabled = NotDefault
                    RowSourceTypeInt =1
                    BorderWidth =3
                    OverlapFlags =215
                    TextFontCharSet =177
                    TextFontFamily =0
                    IMESentenceMode =3
                    Left =3968
                    Top =1756
                    Width =1190
                    Height =300
                    TabIndex =7
                    Name ="KERMaara"
                    RowSourceType ="Value List"
                    RowSource ="\"50\";\"40\";\"30\";\"20\";\"10\""
                    DefaultValue ="\"10\""
                    FontName ="Calibri"
                    OnChange ="[Event Procedure]"
                    GridlineColor =10921638

                    LayoutCachedLeft =3968
                    LayoutCachedTop =1756
                    LayoutCachedWidth =5158
                    LayoutCachedHeight =2056
                    BorderThemeColorIndex =5
                    BorderShade =100.0
                    ForeThemeColorIndex =0
                    ForeTint =75.0
                    ForeShade =100.0
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =215
                            TextFontCharSet =177
                            TextFontFamily =0
                            Left =3968
                            Top =1417
                            Width =1056
                            Height =276
                            BackColor =15921906
                            BorderColor =8355711
                            ForeColor =6710886
                            Name ="Selite342"
                            Caption ="Kertaa"
                            FontName ="Calibri"
                            GridlineColor =10921638
                            LayoutCachedLeft =3968
                            LayoutCachedTop =1417
                            LayoutCachedWidth =5024
                            LayoutCachedHeight =1693
                            BackShade =95.0
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    BorderWidth =3
                    OverlapFlags =215
                    TextFontCharSet =177
                    TextAlign =1
                    TextFontFamily =0
                    IMESentenceMode =3
                    Left =2412
                    Top =3948
                    Width =3780
                    Height =336
                    TabIndex =4
                    ForeColor =4210752
                    Name ="Korttityyppi"
                    FontName ="Calibri"
                    GroupTable =2
                    BottomPadding =150
                    GridlineColor =10921638

                    LayoutCachedLeft =2412
                    LayoutCachedTop =3948
                    LayoutCachedWidth =6192
                    LayoutCachedHeight =4284
                    RowStart =8
                    RowEnd =8
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =2
                    BorderThemeColorIndex =5
                    BorderShade =100.0
                    GroupTable =2
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =215
                            TextFontCharSet =177
                            TextAlign =1
                            TextFontFamily =0
                            Left =336
                            Top =3948
                            Width =1980
                            Height =336
                            BorderColor =8355711
                            ForeColor =6710886
                            Name ="Selite345"
                            Caption ="Korttityyppi:"
                            FontName ="Calibri"
                            GroupTable =2
                            BottomPadding =150
                            GridlineColor =10921638
                            LayoutCachedLeft =336
                            LayoutCachedTop =3948
                            LayoutCachedWidth =2316
                            LayoutCachedHeight =4284
                            RowStart =8
                            RowEnd =8
                            LayoutGroup =2
                            GroupTable =2
                        End
                    End
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =215
                    TextFontCharSet =177
                    TextAlign =1
                    TextFontFamily =0
                    Left =396
                    Top =5328
                    Width =5844
                    Height =564
                    BackColor =62207
                    BorderColor =8355711
                    ForeColor =6710886
                    Name ="Selite56"
                    Caption ="Paina enter, tab tai klikkaa toista kenttää jos et meinaa päästä eteenpäin tieto"
                        "jen syöttämisen jälkeen!"
                    FontName ="Calibri"
                    GroupTable =13
                    BottomPadding =150
                    GridlineColor =10921638
                    LayoutCachedLeft =396
                    LayoutCachedTop =5328
                    LayoutCachedWidth =6240
                    LayoutCachedHeight =5892
                    RowStart =1
                    RowEnd =1
                    LayoutGroup =3
                    BackThemeColorIndex =-1
                    GroupTable =13
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



Private Sub Form_Load()
    Dim checksum As Integer
    
    checksum = Paivita_UI()
End Sub

Private Sub Form_Open(Cancel As Integer)
Dim checksum As Integer
    
    checksum = Paivita_UI()
    
End Sub

Private Sub Hinta_Change()
    Dim checksum As Integer
    checksum = Paivita_UI()
End Sub

Private Sub KERMaara_Change()
    Dim checksum As Integer
    checksum = Paivita_UI()
End Sub

Private Sub KKmaara_Change()
    Dim checksum As Integer
    checksum = Paivita_UI()
End Sub

Private Sub Puumerkki_Change()
    Dim checksum As Integer
    checksum = Paivita_UI()
    'this needs to be moved to under Paivita_UI()!
    [Form_RekisteroiLataus].Hinta.Visible = True
    [Form_RekisteroiLataus].Voimassa.Visible = True
    [Form_RekisteroiLataus].Korttityyppi.Visible = True
    [Form_RekisteroiLataus].Save.Visible = True
    
End Sub


Private Sub Save_Click()
    DoCmd.OpenForm "LatausOhje"
End Sub

Public Sub SaveStuff()
    Dim Kortti As String
    Dim Korttityyppi As String
    Dim Puumerkki As String
    Dim Hinta As Currency
    Dim Voimassaolo As Date
    Dim arvot As String
    
    
    'check that all required info is ok
    
    
    If IsNull([Form_Tervetuloa].Korttivalinta) Then
        MsgBox ("Korttia ei valittu pääikkunasta, ei voida jatkaa!")
        Exit Sub
    Else
        Kortti = [Form_RekisteroiLataus].Kortti.Value
    End If



    If ([Form_RekisteroiLataus].Korttityyppi.Value = "") Or IsNull([Form_RekisteroiLataus].Korttityyppi) Then
        MsgBox ("Korttityyppiä ei määritelty!")
        Exit Sub
    Else
        Korttityyppi = [Form_RekisteroiLataus].Korttityyppi.Value
    End If
    
    If ([Form_RekisteroiLataus].Puumerkki.Value = "") Or IsNull([Form_RekisteroiLataus].Puumerkki) Then
        MsgBox ("Puumerkki ei voi olla tyhjä!")
        Exit Sub
    Else
        Puumerkki = [Form_RekisteroiLataus].Puumerkki.Value
    End If
    
    If ([Form_RekisteroiLataus].Hinta.Value = "") Or IsNull([Form_RekisteroiLataus].Hinta) Then
        MsgBox ("Hintaa ei määritelty!")
        Exit Sub
    Else
        Hinta = [Form_RekisteroiLataus].Hinta.Value
    End If
    
    If ([Form_RekisteroiLataus].Voimassa.Value = "") Or IsNull([Form_RekisteroiLataus].Voimassa) Then
        MsgBox ("Voimassaoloa ei määritelty!")
        Exit Sub
    Else
        Voimassaolo = [Form_RekisteroiLataus].Voimassa.Value
    End If
    
    
    Dim kortti_id As Integer
    kortti_id = Common.FetchCardID(Kortti)
    'MsgBox (kortti_id)
    arvot = ("Kortti = " & kortti_id & " , Voimassa = '" & Voimassaolo & "' , Lataaja = '" & Puumerkki & "' , Korttityyppi = '" & Korttityyppi & "' , KortinArvo = '" & Hinta & "' , Ajankohta = '" & Date & "'")
    'MsgBox (arvot)
    
    'Dim preventDuplicates As String
    
    'preventDuplicates = "Kortti " & kortti_id & " , Voimassa = '" & Voimassaolo & "'"
    'note to self = you need to make sure there are no existing values there so it means fixing the function at common
    Dim success As Boolean
    
    success = Common.InsertOrUpdate("Lataukset", arvot, "")

    Common.SaveToLog (Puumerkki & " päivitti lataukset kortille " & Kortti & ", tyyppi: " & Korttityyppi & " , voimassa: " & Voimassaolo & " ja hinta: " & Hinta)

    Dim retval
    retval = Common.SendMessageToMainScreen("Lataus kortille " & Kortti & " rekisteröity!")


    DoCmd.Close

End Sub

Private Sub Valinta_Click()
Dim checksum As Integer

checksum = Paivita_UI()

End Sub


Private Function Paivita_UI()

'No inputs yet
If (IsNull([Form_RekisteroiLataus].Valinta)) Then
    [Form_RekisteroiLataus].KKmaara.Enabled = False
    [Form_RekisteroiLataus].KKmaara.Visible = False
    [Form_RekisteroiLataus].KERMaara.Visible = False
    [Form_RekisteroiLataus].Puumerkki.Visible = False
    [Form_RekisteroiLataus].Hinta.Visible = False
    [Form_RekisteroiLataus].Voimassa.Visible = False
    [Form_RekisteroiLataus].Save.Visible = False
    [Form_RekisteroiLataus].Korttityyppi.Visible = False

Else
    [Form_RekisteroiLataus].Puumerkki.Visible = True
    
    Dim student As Boolean
    Dim morning As Boolean
    Dim kertak As Boolean
    
    student = False
    morning = False
    kertak = False
    
    Dim output As Integer
    
    Select Case [Form_RekisteroiLataus].Valinta.Value
        
        Case 1:
            output = HideOrShowQty(True)

        Case 2:
            output = HideOrShowQty(True)
            morning = True
        Case 3:
            output = HideOrShowQty(True)
            student = True
        Case 4:
            output = HideOrShowQty(False)
            kertak = True
        
        Case 5:
            output = HideOrShowQty(True)
            
    
End Select




Dim feedback As Integer


feedback = Common.FillCardChargeData([Form_RekisteroiLataus].KKmaara.Value, [Form_RekisteroiLataus].Valinta.Value)


End If



End Function

Function HideOrShowQty(months As Boolean)
If (months) Then
    [Form_RekisteroiLataus].KKmaara.Enabled = True
    [Form_RekisteroiLataus].KKmaara.Visible = True
    [Form_RekisteroiLataus].KERMaara.Enabled = False
    [Form_RekisteroiLataus].KERMaara.Visible = False
Else
    [Form_RekisteroiLataus].KKmaara.Enabled = False
    [Form_RekisteroiLataus].KKmaara.Visible = False
    [Form_RekisteroiLataus].KERMaara.Enabled = True
    [Form_RekisteroiLataus].KERMaara.Visible = True
End If


End Function

Private Sub Voimassa_Change()
    Dim succs
    succs = Paivita_UI()
End Sub
