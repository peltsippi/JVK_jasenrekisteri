Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    Modal = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    CloseButton = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    ViewsAllowed =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    PictureType =2
    GridY =10
    Width =15312
    DatasheetFontHeight =11
    ItemSuffix =11
    Left =4044
    Top =3456
    Right =17796
    Bottom =11712
    Picture ="bulldog_pienempi"
    RecSrcDt = Begin
        0xa2045f05d1c6e540
    End
    RecordSource ="Kortit"
    Caption ="Korjaa kortin latauksia"
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
        Begin FormHeader
            Height =1026
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
                    TextFontFamily =0
                    Left =57
                    Top =57
                    Width =3857
                    Height =561
                    FontSize =20
                    BackColor =15921906
                    BorderColor =8355711
                    ForeColor =6710886
                    Name ="Selite4"
                    Caption ="Latausten korjaaminen"
                    FontName ="Calibri"
                    GridlineColor =10921638
                    LayoutCachedLeft =57
                    LayoutCachedTop =57
                    LayoutCachedWidth =3914
                    LayoutCachedHeight =618
                    BackShade =95.0
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
                    Left =13713
                    Width =1599
                    Height =300
                    BackColor =15921906
                    BorderColor =10921638
                    Name ="Automaattinen_päivämäärä"
                    ControlSource ="=Date()"
                    Format ="Short Date"
                    FontName ="Calibri"
                    GridlineColor =10921638

                    LayoutCachedLeft =13713
                    LayoutCachedWidth =15312
                    LayoutCachedHeight =300
                    BackShade =95.0
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =87
                    TextFontCharSet =177
                    TextAlign =3
                    TextFontFamily =0
                    IMESentenceMode =3
                    Left =13713
                    Top =300
                    Width =1599
                    Height =300
                    TabIndex =1
                    BackColor =15921906
                    BorderColor =10921638
                    Name ="Automaattinen_aika"
                    ControlSource ="=Time()"
                    Format ="Long Time"
                    FontName ="Calibri"
                    GridlineColor =10921638

                    LayoutCachedLeft =13713
                    LayoutCachedTop =300
                    LayoutCachedWidth =15312
                    LayoutCachedHeight =600
                    BackShade =95.0
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                End
            End
        End
        Begin Section
            Height =5586
            Name ="Tiedot"
            AutoHeight =1
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin Subform
                    OverlapFlags =85
                    Left =427
                    Top =1247
                    Width =14709
                    Height =4218
                    TabIndex =1
                    BorderColor =10921638
                    Name ="LatausListaus"
                    SourceObject ="Form.LatausListaus"
                    LinkChildFields ="Kortti"
                    LinkMasterFields ="CID"
                    GridlineColor =10921638

                    LayoutCachedLeft =427
                    LayoutCachedTop =1247
                    LayoutCachedWidth =15136
                    LayoutCachedHeight =5465
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =85
                            TextFontCharSet =177
                            TextFontFamily =0
                            Left =396
                            Top =850
                            Width =2208
                            Height =312
                            BorderColor =8355711
                            ForeColor =6710886
                            Name ="LatausListaus_Selite"
                            Caption ="Kortin lataukset listana:"
                            FontName ="Calibri"
                            GridlineColor =10921638
                            LayoutCachedLeft =396
                            LayoutCachedTop =850
                            LayoutCachedWidth =2604
                            LayoutCachedHeight =1162
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =177
                    TextFontFamily =0
                    IMESentenceMode =3
                    Left =2640
                    Top =342
                    Width =2292
                    Height =300
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Kortti"
                    ControlSource ="=[Forms]![Tervetuloa]![Korttivalinta]"
                    StatusBarText ="Kortin numero niin kuin se on kirjoitettu, esim 0285"
                    FontName ="Calibri"
                    GridlineColor =10921638

                    LayoutCachedLeft =2640
                    LayoutCachedTop =342
                    LayoutCachedWidth =4932
                    LayoutCachedHeight =642
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =85
                            TextFontCharSet =177
                            TextFontFamily =0
                            Left =342
                            Top =342
                            Width =2208
                            Height =312
                            BorderColor =8355711
                            ForeColor =6710886
                            Name ="Kortti_Selite"
                            Caption ="Kortti:"
                            FontName ="Calibri"
                            GridlineColor =10921638
                            LayoutCachedLeft =342
                            LayoutCachedTop =342
                            LayoutCachedWidth =2550
                            LayoutCachedHeight =654
                        End
                    End
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =85
                    TextFontCharSet =177
                    TextFontFamily =0
                    Left =7540
                    Top =170
                    Width =7711
                    Height =851
                    FontSize =16
                    FontWeight =700
                    BackColor =62207
                    BorderColor =8355711
                    ForeColor =255
                    Name ="Selite9"
                    Caption ="HUOM! Muutokset tallentuu aina välittömästi!\015\012Olethan varovainen!"
                    FontName ="Calibri"
                    GridlineColor =10921638
                    LayoutCachedLeft =7540
                    LayoutCachedTop =170
                    LayoutCachedWidth =15251
                    LayoutCachedHeight =1021
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontCharSet =177
                    TextFontFamily =0
                    Left =5045
                    Top =226
                    Width =2383
                    Height =794
                    TabIndex =2
                    ForeColor =4210752
                    Name ="Close"
                    Caption ="Sulje"
                    OnClick ="[Event Procedure]"
                    FontName ="Calibri"
                    GridlineColor =10921638

                    LayoutCachedLeft =5045
                    LayoutCachedTop =226
                    LayoutCachedWidth =7428
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

Private Sub Close_Click()
    Dim succs
    succs = Common.SaveToLog("Latausmuokkausten teko lopetettu")
    succs = Common.SendMessageToMainScreen("Latausmuokkaukset tehty")
    DoCmd.Close
    
End Sub
