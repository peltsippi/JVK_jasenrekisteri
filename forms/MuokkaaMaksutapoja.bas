Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    Modal = NotDefault
    AutoCenter = NotDefault
    CloseButton = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    ViewsAllowed =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    PictureType =2
    GridY =10
    Width =7711
    DatasheetFontHeight =11
    ItemSuffix =4
    Left =4044
    Top =3468
    Right =17796
    Bottom =11712
    Picture ="bulldog_pienempi"
    RecSrcDt = Begin
        0xd6ff39645cc7e540
    End
    RecordSource ="Maksutavat"
    Caption ="MuokkaaMaksutapoja"
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
        Begin FormHeader
            Height =1073
            Name ="LomakkeenYlätunniste"
            AutoHeight =1
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =2
            BackTint =20.0
            Begin
                Begin Label
                    BackStyle =1
                    OverlapFlags =93
                    Left =57
                    Top =57
                    Width =3737
                    Height =561
                    FontSize =20
                    BackColor =14277081
                    BorderColor =8355711
                    ForeColor =6710886
                    Name ="Selite2"
                    Caption ="Muokkaa maksutapoja"
                    FontName ="Calibri"
                    GridlineColor =10921638
                    LayoutCachedLeft =57
                    LayoutCachedTop =57
                    LayoutCachedWidth =3794
                    LayoutCachedHeight =618
                    BackShade =85.0
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =215
                    Left =3231
                    Top =510
                    Width =4423
                    Height =563
                    FontSize =10
                    FontWeight =700
                    BackColor =62207
                    BorderColor =8355711
                    ForeColor =255
                    Name ="Selite9"
                    Caption ="HUOM! Muutokset tallentuu aina välittömästi!\015\012Olethan varovainen!"
                    FontName ="Calibri"
                    GridlineColor =10921638
                    LayoutCachedLeft =3231
                    LayoutCachedTop =510
                    LayoutCachedWidth =7654
                    LayoutCachedHeight =1073
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
            End
        End
        Begin Section
            Height =793
            Name ="Tiedot"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =342
                    Top =57
                    Width =4296
                    Height =480
                    FontSize =18
                    FontWeight =700
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Maksutapa"
                    ControlSource ="Maksutapa"
                    FontName ="Calibri"
                    GridlineColor =10921638

                    LayoutCachedLeft =342
                    LayoutCachedTop =57
                    LayoutCachedWidth =4638
                    LayoutCachedHeight =537
                End
            End
        End
        Begin FormFooter
            Height =1190
            Name ="LomakkeenAlatunniste"
            AutoHeight =1
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin CommandButton
                    OverlapFlags =85
                    Left =170
                    Top =283
                    Width =7303
                    Height =794
                    ForeColor =4210752
                    Name ="Close"
                    Caption ="Sulje"
                    OnClick ="[Event Procedure]"
                    FontName ="Calibri"
                    GridlineColor =10921638

                    LayoutCachedLeft =170
                    LayoutCachedTop =283
                    LayoutCachedWidth =7473
                    LayoutCachedHeight =1077
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
    succs = Common.SaveToLog("Maksutapojen muokkaus lopetettu")
    succs = Common.SendMessageToMainScreen("Valmis!")
    DoCmd.Close
    
End Sub
