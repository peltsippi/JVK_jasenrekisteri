Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    Modal = NotDefault
    CloseButton = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =9524
    DatasheetFontHeight =11
    ItemSuffix =60
    Left =4044
    Top =3456
    Right =17484
    Bottom =11712
    RecSrcDt = Begin
        0xf995cdd7bcc7e540
    End
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
        Begin EmptyCell
            Height =240
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin Section
            Height =9467
            Name ="Tiedot"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin Image
                    PictureAlignment =0
                    PictureType =2
                    Left =60
                    Top =36
                    Width =9360
                    Height =8460
                    BorderColor =10921638
                    Name ="Kuva0"
                    Picture ="JVK_jasenrekisteri_img"
                    GroupTable =1
                    GridlineColor =10921638

                    LayoutCachedLeft =60
                    LayoutCachedTop =36
                    LayoutCachedWidth =9420
                    LayoutCachedHeight =8496
                    TabIndex =9
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    Locked = NotDefault
                    BorderWidth =3
                    OverlapFlags =247
                    TextAlign =1
                    IMESentenceMode =3
                    Left =2100
                    Top =1867
                    Width =2205
                    Height =408
                    FontSize =14
                    ForeColor =4210752
                    Name ="KorttiNumero"
                    FontName ="Calibri"
                    GridlineColor =10921638

                    LayoutCachedLeft =2100
                    LayoutCachedTop =1867
                    LayoutCachedWidth =4305
                    LayoutCachedHeight =2275
                    BorderThemeColorIndex =5
                    BorderShade =100.0
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =170
                    Top =8617
                    Width =4308
                    Height =680
                    TabIndex =3
                    ForeColor =4210752
                    Name ="Save"
                    Caption ="Tehty, painoin molempia talleta -nappeja!"
                    OnClick ="[Event Procedure]"
                    FontName ="Calibri"
                    GridlineColor =10921638

                    LayoutCachedLeft =170
                    LayoutCachedTop =8617
                    LayoutCachedWidth =4478
                    LayoutCachedHeight =9297
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
                Begin CommandButton
                    OverlapFlags =85
                    Left =4592
                    Top =8617
                    Width =4643
                    Height =683
                    TabIndex =4
                    ForeColor =4210752
                    Name ="Cancel"
                    Caption ="Peruuta!"
                    OnClick ="[Event Procedure]"
                    FontName ="Calibri"
                    GridlineColor =10921638

                    LayoutCachedLeft =4592
                    LayoutCachedTop =8617
                    LayoutCachedWidth =9235
                    LayoutCachedHeight =9300
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
                Begin TextBox
                    Locked = NotDefault
                    BorderWidth =3
                    OverlapFlags =247
                    IMESentenceMode =3
                    Left =1644
                    Top =5839
                    Width =3515
                    Height =336
                    FontSize =14
                    TabIndex =1
                    ForeColor =4210752
                    Name ="KorttiTyyppi"
                    FontName ="Calibri"
                    GridlineColor =10921638

                    LayoutCachedLeft =1644
                    LayoutCachedTop =5839
                    LayoutCachedWidth =5159
                    LayoutCachedHeight =6175
                    BorderThemeColorIndex =5
                    BorderShade =100.0
                End
                Begin TextBox
                    Locked = NotDefault
                    BorderWidth =3
                    OverlapFlags =247
                    IMESentenceMode =3
                    Left =1644
                    Top =6746
                    Width =3458
                    Height =396
                    FontSize =14
                    TabIndex =2
                    ForeColor =4210752
                    Name ="AikaRyhma"
                    FontName ="Calibri"
                    GridlineColor =10921638

                    LayoutCachedLeft =1644
                    LayoutCachedTop =6746
                    LayoutCachedWidth =5102
                    LayoutCachedHeight =7142
                    BorderThemeColorIndex =5
                    BorderShade =100.0
                End
                Begin TextBox
                    Locked = NotDefault
                    BorderWidth =3
                    OverlapFlags =247
                    IMESentenceMode =3
                    Left =1644
                    Top =6292
                    Width =1810
                    Height =348
                    FontSize =14
                    TabIndex =5
                    ForeColor =4210752
                    Name ="Tanaan"
                    FontName ="Calibri"
                    GridlineColor =10921638

                    LayoutCachedLeft =1644
                    LayoutCachedTop =6292
                    LayoutCachedWidth =3454
                    LayoutCachedHeight =6640
                    BorderThemeColorIndex =5
                    BorderShade =100.0
                End
                Begin TextBox
                    Locked = NotDefault
                    BorderWidth =3
                    OverlapFlags =247
                    IMESentenceMode =3
                    Left =3911
                    Top =6292
                    Width =1814
                    Height =348
                    FontSize =14
                    TabIndex =6
                    ForeColor =4210752
                    Name ="Voimassa"
                    FontName ="Calibri"
                    GridlineColor =10921638

                    LayoutCachedLeft =3911
                    LayoutCachedTop =6292
                    LayoutCachedWidth =5725
                    LayoutCachedHeight =6640
                    BorderThemeColorIndex =5
                    BorderShade =100.0
                End
                Begin TextBox
                    OverlapFlags =247
                    IMESentenceMode =3
                    Left =4422
                    Top =1927
                    Width =2267
                    Height =300
                    TabIndex =7
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Muokkaus56"
                    FontName ="Calibri"
                    GridlineColor =10921638

                    LayoutCachedLeft =4422
                    LayoutCachedTop =1927
                    LayoutCachedWidth =6689
                    LayoutCachedHeight =2227
                End
                Begin TextBox
                    BorderWidth =3
                    OverlapFlags =247
                    IMESentenceMode =3
                    Left =1644
                    Top =7196
                    Width =850
                    Height =396
                    FontSize =14
                    TabIndex =8
                    ForeColor =4210752
                    Name ="Maara"
                    FontName ="Calibri"
                    GridlineColor =10921638

                    LayoutCachedLeft =1644
                    LayoutCachedTop =7196
                    LayoutCachedWidth =2494
                    LayoutCachedHeight =7592
                    BorderThemeColorIndex =5
                    BorderShade =100.0
                End
            End
        End
    End
End
CodeBehindForm
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Cancel_Click()
    DoCmd.Close
    
End Sub

Private Sub Form_Open(Cancel As Integer)

    [Form_LatausOhje].KorttiNumero.Value = "A" & [Form_RekisteroiLataus].Kortti
    
    
    If ([Form_RekisteroiLataus].Valinta = 4) Then
        [Form_LatausOhje].Korttityyppi.Value = "Määräkortti"
        [Form_LatausOhje].Maara.Value = [Form_RekisteroiLataus].KERMaara.Value
    Else
        [Form_LatausOhje].Korttityyppi.Value = "Kausikortti"
    End If
    
    If ([Form_RekisteroiLataus].Valinta = 2) Then
        [Form_LatausOhje].AikaRyhma.Value = "aamupäivä ma-su"
    Else
        [Form_LatausOhje].AikaRyhma.Value = "Normaali"
    End If
    
    [Form_LatausOhje].Voimassa.Value = [Form_RekisteroiLataus].Voimassa.Value
    
    [Form_LatausOhje].Tanaan.Value = Date
    
End Sub

Private Sub Save_Click()
    DoCmd.Close
    
    [Form_RekisteroiLataus].SaveStuff
    
    
End Sub
