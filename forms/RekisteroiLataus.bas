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
    Width =6827
    DatasheetFontHeight =11
    ItemSuffix =421
    Left =4740
    Top =3468
    Right =22788
    Bottom =11712
    Picture ="bulldog_pienempi"
    RecSrcDt = Begin
        0x84756ccb9ec6e540
    End
    Caption ="Rekisteröi lataus"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Calibri"
    OnLoad ="[Event Procedure]"
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
            Height =640
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
            Height =8277
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
                    Width =3768
                    Height =336
                    TabIndex =5
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
                    LayoutCachedWidth =6180
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
                    Visible = NotDefault
                    BorderWidth =3
                    OverlapFlags =215
                    TextFontCharSet =177
                    TextAlign =1
                    TextFontFamily =0
                    IMESentenceMode =3
                    Left =2412
                    Top =3888
                    Width =3768
                    Height =336
                    TabIndex =6
                    ForeColor =4210752
                    Name ="Hinta"
                    Format ="#,##0.00 €;-#,##0.00 €"
                    FontName ="Calibri"
                    GroupTable =2
                    BottomPadding =150
                    GridlineColor =10921638

                    LayoutCachedLeft =2412
                    LayoutCachedTop =3888
                    LayoutCachedWidth =6180
                    LayoutCachedHeight =4224
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
                            Top =3888
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
                            LayoutCachedTop =3888
                            LayoutCachedWidth =2316
                            LayoutCachedHeight =4224
                            RowStart =6
                            RowEnd =6
                            LayoutGroup =2
                            GroupTable =2
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    BorderWidth =3
                    OverlapFlags =215
                    TextFontCharSet =177
                    TextAlign =1
                    TextFontFamily =0
                    IMESentenceMode =3
                    Left =2412
                    Top =4440
                    Width =3768
                    Height =336
                    TabIndex =7
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
                    LayoutCachedTop =4440
                    LayoutCachedWidth =6180
                    LayoutCachedHeight =4776
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
                            Top =4440
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
                            LayoutCachedTop =4440
                            LayoutCachedWidth =2316
                            LayoutCachedHeight =4776
                            RowStart =7
                            RowEnd =7
                            LayoutGroup =2
                            GroupTable =2
                        End
                    End
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =223
                    TextFontCharSet =177
                    TextAlign =1
                    TextFontFamily =0
                    Left =396
                    Top =4932
                    Width =5952
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
                    LayoutCachedTop =4932
                    LayoutCachedWidth =6348
                    LayoutCachedHeight =5544
                    LayoutGroup =3
                    BackThemeColorIndex =-1
                    GroupTable =13
                End
                Begin TextBox
                    Visible = NotDefault
                    BorderWidth =3
                    OverlapFlags =215
                    TextFontCharSet =177
                    TextAlign =1
                    TextFontFamily =0
                    IMESentenceMode =3
                    Left =2412
                    Top =3384
                    Width =3768
                    Height =288
                    TabIndex =3
                    ForeColor =4210752
                    Name ="Puumerkki"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Calibri"
                    OnChange ="[Event Procedure]"
                    GroupTable =2
                    BottomPadding =150
                    GridlineColor =10921638

                    LayoutCachedLeft =2412
                    LayoutCachedTop =3384
                    LayoutCachedWidth =6180
                    LayoutCachedHeight =3672
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
                            Top =3384
                            Width =1980
                            Height =288
                            Name ="Selite176"
                            Caption ="Lataajan puumerkki"
                            FontName ="Calibri"
                            GroupTable =2
                            BottomPadding =150
                            GridlineColor =10921638
                            LayoutCachedLeft =336
                            LayoutCachedTop =3384
                            LayoutCachedWidth =2316
                            LayoutCachedHeight =3672
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
                    Visible = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =177
                    TextFontFamily =0
                    Left =396
                    Top =6480
                    Width =5952
                    Height =696
                    TabIndex =4
                    ForeColor =4210752
                    Name ="Save"
                    Caption ="Tallenna"
                    OnClick ="[Event Procedure]"
                    FontName ="Calibri"
                    GroupTable =13
                    BottomPadding =150
                    GridlineColor =10921638

                    LayoutCachedLeft =396
                    LayoutCachedTop =6480
                    LayoutCachedWidth =6348
                    LayoutCachedHeight =7176
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
                    Top =7368
                    Width =5952
                    Height =672
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
                    LayoutCachedTop =7368
                    LayoutCachedWidth =6348
                    LayoutCachedHeight =8040
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
                    Top =2388
                    Width =1980
                    Name ="TyhjäSolu301"
                    GroupTable =2
                    BottomPadding =150
                    GridlineColor =10921638
                    LayoutCachedLeft =336
                    LayoutCachedTop =2388
                    LayoutCachedWidth =2316
                    LayoutCachedHeight =2628
                    RowStart =3
                    RowEnd =3
                    LayoutGroup =2
                    GroupTable =2
                End
                Begin EmptyCell
                    Left =2412
                    Top =2388
                    Width =3768
                    Name ="TyhjäSolu302"
                    GroupTable =2
                    BottomPadding =150
                    GridlineColor =10921638
                    LayoutCachedLeft =2412
                    LayoutCachedTop =2388
                    LayoutCachedWidth =6180
                    LayoutCachedHeight =2628
                    RowStart =3
                    RowEnd =3
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =2
                    GroupTable =2
                End
                Begin ComboBox
                    RowSourceTypeInt =1
                    BorderWidth =3
                    OverlapFlags =215
                    TextFontCharSet =177
                    TextAlign =1
                    TextFontFamily =0
                    IMESentenceMode =3
                    Left =2412
                    Top =1656
                    Width =3768
                    Height =552
                    FontSize =16
                    TabIndex =2
                    Name ="KKmaara"
                    RowSourceType ="Value List"
                    RowSource ="\"0\""
                    DefaultValue ="\"0\""
                    FontName ="Calibri"
                    OnGotFocus ="[Event Procedure]"
                    OnChange ="[Event Procedure]"
                    GroupTable =2
                    BottomPadding =150
                    GridlineColor =10921638
                    AllowValueListEdits =0

                    LayoutCachedLeft =2412
                    LayoutCachedTop =1656
                    LayoutCachedWidth =6180
                    LayoutCachedHeight =2208
                    RowStart =2
                    RowEnd =2
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =2
                    BorderThemeColorIndex =5
                    BorderShade =100.0
                    ForeThemeColorIndex =0
                    ForeTint =75.0
                    ForeShade =100.0
                    GroupTable =2
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =215
                            TextFontCharSet =177
                            TextAlign =1
                            TextFontFamily =0
                            Left =336
                            Top =1656
                            Width =1980
                            Height =552
                            BackColor =15921906
                            BorderColor =8355711
                            ForeColor =6710886
                            Name ="pituusSelite"
                            Caption ="Kuukautta"
                            FontName ="Calibri"
                            GroupTable =2
                            BottomPadding =150
                            GridlineColor =10921638
                            LayoutCachedLeft =336
                            LayoutCachedTop =1656
                            LayoutCachedWidth =2316
                            LayoutCachedHeight =2208
                            RowStart =2
                            RowEnd =2
                            LayoutGroup =2
                            BackShade =95.0
                            GroupTable =2
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    Locked = NotDefault
                    OverlapFlags =215
                    TextFontCharSet =177
                    TextAlign =1
                    TextFontFamily =0
                    IMESentenceMode =3
                    Left =2412
                    Top =4968
                    Width =3768
                    Height =336
                    TabIndex =8
                    ForeColor =4210752
                    Name ="Korttityyppi"
                    FontName ="Calibri"
                    GroupTable =2
                    BottomPadding =150
                    GridlineColor =10921638

                    LayoutCachedLeft =2412
                    LayoutCachedTop =4968
                    LayoutCachedWidth =6180
                    LayoutCachedHeight =5304
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
                            Top =4968
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
                            LayoutCachedTop =4968
                            LayoutCachedWidth =2316
                            LayoutCachedHeight =5304
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
                    Top =5724
                    Width =5952
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
                    LayoutCachedTop =5724
                    LayoutCachedWidth =6348
                    LayoutCachedHeight =6288
                    RowStart =1
                    RowEnd =1
                    LayoutGroup =3
                    BackThemeColorIndex =-1
                    GroupTable =13
                End
                Begin TextBox
                    BorderWidth =2
                    OverlapFlags =215
                    TextFontCharSet =177
                    TextAlign =1
                    TextFontFamily =0
                    IMESentenceMode =3
                    Left =2412
                    Top =2832
                    Width =3768
                    Height =336
                    BorderColor =5167783
                    ForeColor =4210752
                    Name ="aloituspvm"
                    Format ="Short Date"
                    FontName ="Calibri"
                    OnChange ="[Event Procedure]"
                    GroupTable =2
                    BottomPadding =150
                    GridlineColor =10921638

                    LayoutCachedLeft =2412
                    LayoutCachedTop =2832
                    LayoutCachedWidth =6180
                    LayoutCachedHeight =3168
                    RowStart =4
                    RowEnd =4
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =2
                    BorderThemeColorIndex =-1
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
                            Top =2832
                            Width =1980
                            Height =336
                            BorderColor =8355711
                            ForeColor =6710886
                            Name ="Selite382"
                            Caption ="Aloituspvm"
                            FontName ="Calibri"
                            GroupTable =2
                            BottomPadding =150
                            GridlineColor =10921638
                            LayoutCachedLeft =336
                            LayoutCachedTop =2832
                            LayoutCachedWidth =2316
                            LayoutCachedHeight =3168
                            RowStart =4
                            RowEnd =4
                            LayoutGroup =2
                            GroupTable =2
                        End
                    End
                End
                Begin ComboBox
                    BorderWidth =3
                    OverlapFlags =215
                    TextFontCharSet =177
                    TextAlign =1
                    TextFontFamily =0
                    IMESentenceMode =3
                    Left =2412
                    Top =660
                    Width =3768
                    Height =780
                    FontSize =16
                    TabIndex =1
                    BorderColor =967423
                    ForeColor =3484194
                    ColumnInfo ="\"\";\"\";\"10\";\"510\""
                    Name ="korttiTyyppiValinta"
                    RowSourceType ="Table/Query"
                    RowSource ="haeKorttiTyypit"
                    FontName ="Calibri"
                    OnGotFocus ="[Event Procedure]"
                    OnChange ="[Event Procedure]"
                    GroupTable =2
                    BottomPadding =150
                    GridlineColor =10921638

                    LayoutCachedLeft =2412
                    LayoutCachedTop =660
                    LayoutCachedWidth =6180
                    LayoutCachedHeight =1440
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =2
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    GroupTable =2
                    Begin
                        Begin Label
                            OverlapFlags =215
                            TextFontCharSet =177
                            TextAlign =1
                            TextFontFamily =0
                            Left =336
                            Top =660
                            Width =1980
                            Height =780
                            BorderColor =8355711
                            ForeColor =6710886
                            Name ="Selite392"
                            Caption ="Korttityyppi"
                            FontName ="Calibri"
                            GroupTable =2
                            BottomPadding =150
                            GridlineColor =10921638
                            LayoutCachedLeft =336
                            LayoutCachedTop =660
                            LayoutCachedWidth =2316
                            LayoutCachedHeight =1440
                            RowStart =1
                            RowEnd =1
                            LayoutGroup =2
                            GroupTable =2
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
CodeBehindForm
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database



Private Sub aloituspvm_Change()
    Dim checksum As Integer
    checksum = Paivita_UI()
End Sub



Private Sub Form_Load()
    Dim checksum As Integer
    checksum = Paivita_UI()
End Sub

Private Sub Form_Open(Cancel As Integer)
    If IsNull([Form_Tervetuloa].Korttivalinta) Then
        MsgBox ("Korttia ei valittu pääikkunasta, ei voida jatkaa!")
    End If
    
    Dim checksum As Integer
    checksum = Paivita_UI()
    checksum = GetDefaultDate()
    
End Sub

Private Sub Hinta_Change()
    Dim checksum As Integer
    checksum = Paivita_UI()
End Sub


Private Sub KKmaara_Change()
    Dim checksum As Integer
    checksum = Paivita_UI()
    [Form_RekisteroiLataus].Puumerkki.SetFocus
    
End Sub

Private Sub KKmaara_GotFocus()
    [Form_RekisteroiLataus].KKmaara.Dropdown
End Sub

Private Sub korttiTyyppiValinta_Change()
    'add code to switch stuff here...
    [Form_RekisteroiLataus].KKmaara.RowSourceType = "Table/Query"
    Dim queryString As String
    
    queryString = "SELECT Aika FROM Hinnasto WHERE Tyyppi = '" & [Form_RekisteroiLataus].korttiTyyppivalinta.Value & "' ORDER BY Aika DESC"
    
    'MsgBox (queryString)
    [Form_RekisteroiLataus].KKmaara.RowSource = queryString
    [Form_RekisteroiLataus].KKmaara.Value = [Form_RekisteroiLataus].KKmaara.ItemData(6)
    If ([Form_RekisteroiLataus].korttiTyyppivalinta.Value = "krt") Then
        [Form_RekisteroiLataus].KKmaara.Value = [Form_RekisteroiLataus].KKmaara.ItemData([Form_RekisteroiLataus].KKmaara.ListCount - 1)
        [Form_RekisteroiLataus].pituusSelite.Caption = "Kertaa"
    Else
        [Form_RekisteroiLataus].KKmaara.Value = [Form_RekisteroiLataus].KKmaara.ItemData([Form_RekisteroiLataus].KKmaara.ListCount / 2)
        [Form_RekisteroiLataus].pituusSelite.Caption = "Kuukautta"
    End If
    [Form_RekisteroiLataus].Puumerkki.Visible = True
    
    [Form_RekisteroiLataus].KKmaara.SetFocus
End Sub

Private Sub korttiTyyppiValinta_GotFocus()
    [Form_RekisteroiLataus].korttiTyyppivalinta.Dropdown
End Sub

Private Sub Puumerkki_AfterUpdate()
    Dim checksum As Integer
    checksum = Paivita_UI()
End Sub

Private Sub Puumerkki_Change()
    Dim checksum As Integer
    checksum = Paivita_UI()
    '[Form_RekisteroiLataus].Save.SetFocus 'this does not work! And seems to be hard to implement...
    
    
End Sub


Private Sub Save_Click()
    DoCmd.OpenForm "LatausOhje"
End Sub

Public Sub SaveStuff()
    Dim Kortti As String
    Dim KorttiTyyppi As String
    Dim Puumerkki As String
    Dim Hinta As Currency
    Dim Voimassaolo As Date
    Dim arvot As String
    Dim vanhanKortinVoimassaOlo As Date
    Dim startDate As Date
    

    'check that all required info is ok
    
    
    If IsNull([Form_RekisteroiLataus].aloituspvm) Then
        MsgBox ("Päivämäärää ei asetettu, peruutetaan toiminto!")
        Exit Sub
    Else
        startDate = [Form_RekisteroiLataus].aloituspvm.Value
    End If
    
    If IsNull([Form_Tervetuloa].Korttivalinta) Then
        MsgBox ("Korttia ei valittu pääikkunasta, ei voida jatkaa!")
        Exit Sub
    Else
        Kortti = [Form_RekisteroiLataus].Kortti.Value
        vanhanKortinVoimassaOlo = Common.FetchExiprationDate(Kortti)
    End If



    If ([Form_RekisteroiLataus].KorttiTyyppi.Value = "") Or IsNull([Form_RekisteroiLataus].KorttiTyyppi) Then
        MsgBox ("Korttityyppiä ei määritelty!")
        Exit Sub
    Else
        KorttiTyyppi = [Form_RekisteroiLataus].KorttiTyyppi.Value
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
    arvot = ("Kortti = " & kortti_id & " , Voimassa = '" & Voimassaolo & "' , Puumerkki = '" & Puumerkki & "' , Korttityyppi = '" & KorttiTyyppi & "' , KortinArvo = '" & Hinta & "' , Ajankohta = '" & startDate & "'")
    'MsgBox (arvot)
    
    'Dim preventDuplicates As String
    
    'preventDuplicates = "Kortti " & kortti_id & " , Voimassa = '" & Voimassaolo & "'"
    'note to self = you need to make sure there are no existing values there so it means fixing the function at common
    Dim success As Boolean
    
    success = Common.InsertOrUpdate("Lataukset", arvot, "")

    Common.SaveToLog (Puumerkki & " päivitti lataukset kortille " & Kortti & ", tyyppi: " & KorttiTyyppi & " , voimassa: " & Voimassaolo & " ja hinta: " & Hinta)
    
    'If old charge date is newer than new charge start date do stuff
    If (startDate > Date) Then 'and this is so if you add a charge retroactively that things will not go badly sour...
    If (vanhanKortinVoimassaOlo > startDate) Then
        Dim succs
        
        Dim criteria, table As String
        
        query = "UPDATE Lataukset SET Voimassa = '" & startDate & "' WHERE Voimassa = DateValue('" & vanhanKortinVoimassaOlo & " ') And Kortti = " & kortti_id
        
        succs = Common.SQLQuery(query)
        
        succs = Common.SaveToLog("Kortin " & card_id & " edellisen latauksen voimassaoloa lyhennetty samalla")
        
        
    End If
    End If
    Dim retval
    retval = Common.SendMessageToMainScreen("Lataus kortille " & Kortti & " rekisteröity!")


    DoCmd.Close

End Sub

Private Sub Valinta_Click()
Dim checksum As Integer

checksum = Paivita_UI()

End Sub


Private Function Paivita_UI()

Dim checksum2 As Integer

'No inputs yet
If (IsNull([Form_RekisteroiLataus].korttiTyyppivalinta)) Then
    [Form_RekisteroiLataus].Puumerkki.Visible = False
    [Form_RekisteroiLataus].Hinta.Visible = False
    [Form_RekisteroiLataus].Voimassa.Visible = False
    [Form_RekisteroiLataus].Save.Visible = False
    [Form_RekisteroiLataus].KorttiTyyppi.Visible = False

Else 'card type selected
    [Form_RekisteroiLataus].Puumerkki.Visible = True
End If

If (IsNull([Form_RekisteroiLataus].Puumerkki)) Then
    [Form_RekisteroiLataus].Hinta.Visible = False
    [Form_RekisteroiLataus].Voimassa.Visible = False
    [Form_RekisteroiLataus].Save.Visible = False
    [Form_RekisteroiLataus].KorttiTyyppi.Visible = False

Else
    checksum2 = UpdateCardPrice()
    [Form_RekisteroiLataus].Hinta.Visible = True
    
    checksum2 = UpdateDueDate()
    [Form_RekisteroiLataus].Voimassa.Visible = True
    [Form_RekisteroiLataus].Save.Visible = True
    
    checksum2 = FormCardType()
    [Form_RekisteroiLataus].KorttiTyyppi.Visible = True

End If


'and update calculations if enough data in form
    'if
If Not (IsNull([Form_RekisteroiLataus].korttiTyyppivalinta.Value) Or IsNull([Form_RekisteroiLataus].aloituspvm.Value) Or IsNull([Form_RekisteroiLataus].korttiTyyppivalinta) Or IsNull([Form_RekisteroiLataus].KKmaara)) Then
    'MsgBox ("All data properly filled")
    Dim succs
    succs = UpdateCardPrice()
    succs = UpdateDueDate()
End If
    

End Function



Private Sub Voimassa_Change()
    Dim succs
    succs = Paivita_UI()
End Sub

Private Function GetDefaultDate()
    Dim initialDate As Date
    Dim Kortti As String
    Kortti = [Form_RekisteroiLataus].Kortti.Value
    'MsgBox ("Kortti: " & kortti)
    initialDate = Common.FetchExiprationDate(Kortti)
    If (initialDate < Date) Then
        initialDate = Date
    End If
    'MsgBox ("Initial date : " & initialDate)
    
    'TODO: add logic to handle situation where start date goes years to future...
    
    Dim naggingLogic As Date
    
    naggingLogic = DateAdd("m", 3, Date)
    
    If (naggingLogic < initialDate) Then
        Dim answer As Integer
        answer = MsgBox("Kortin edellinen lataus on voimassa vielä yli 3kk," & vbNewLine & "kirjataanko uusi lataus alkamaan tästä päivästä?", vbQuestion + vbYesNo, "Edellisen latauksen päättymiseen yli 3kk")
        
        If (answer = vbYes) Then
            initialDate = Date
        End If
        
    End If
    
    [Form_RekisteroiLataus].aloituspvm.Value = initialDate
    
End Function

Public Function UpdateCardPrice()
    Dim cardType As String
    Dim cardTime As Integer
    Dim price As String
    cardType = [Form_RekisteroiLataus].korttiTyyppivalinta
    cardTime = [Form_RekisteroiLataus].KKmaara
    price = Common.GetPriceForCard(cardType, cardTime)
    'MsgBox (price)
    [Form_RekisteroiLataus].Hinta.Value = price

End Function

Public Function UpdateDueDate()
    Dim dueDate As Date
    Dim months As Integer
    
    If ([Form_RekisteroiLataus].korttiTyyppivalinta.Value = "krt") Then
        months = 24
        '[Form_RekisteroiLataus].aloituspvm.Value = Date
    Else
        months = [Form_RekisteroiLataus].KKmaara.Value
    End If
    'MsgBox ("Lisättäviä kuukausia: " & months)
    dueDate = Common.CalculateEndingDate(months, [Form_RekisteroiLataus].aloituspvm.Value)
    
    [Form_RekisteroiLataus].Voimassa = dueDate
    
End Function

Public Function FormCardType()
    Dim cardType As String
    cardType = [Form_RekisteroiLataus].KKmaara.Value & [Form_RekisteroiLataus].korttiTyyppivalinta.Value
    [Form_RekisteroiLataus].KorttiTyyppi.Value = cardType
End Function
