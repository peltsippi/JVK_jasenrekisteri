Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    RecordSelectors = NotDefault
    NavigationButtons = NotDefault
    CloseButton = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    PictureType =2
    GridY =10
    Width =7880
    DatasheetFontHeight =11
    ItemSuffix =43
    Left =4044
    Top =3468
    Right =17796
    Bottom =11712
    Picture ="bulldog_pienempi"
    RecSrcDt = Begin
        0xa22cdf047cc5e540
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
                    Top =60
                    Width =3156
                    Height =460
                    FontSize =18
                    BackColor =15921906
                    Name ="Automaattinen_ylätunniste0"
                    Caption ="Poista kortin linkitys"
                    FontName ="Calibri Light"
                    GroupTable =1
                    GridlineColor =10921638
                    LayoutCachedLeft =300
                    LayoutCachedTop =60
                    LayoutCachedWidth =3456
                    LayoutCachedHeight =520
                    LayoutGroup =1
                    ThemeFontIndex =0
                    BackShade =95.0
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
                    TextAlign =3
                    IMESentenceMode =3
                    Left =5889
                    Top =60
                    Width =1635
                    Height =300
                    BackColor =15921906
                    BorderColor =10921638
                    Name ="Automaattinen_päivämäärä"
                    ControlSource ="=Date()"
                    Format ="Short Date"
                    FontName ="Calibri"
                    GridlineColor =10921638

                    LayoutCachedLeft =5889
                    LayoutCachedTop =60
                    LayoutCachedWidth =7524
                    LayoutCachedHeight =360
                    BackShade =95.0
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =87
                    TextAlign =3
                    IMESentenceMode =3
                    Left =5889
                    Top =360
                    Width =1635
                    Height =300
                    TabIndex =1
                    BackColor =15921906
                    BorderColor =10921638
                    Name ="Automaattinen_aika"
                    ControlSource ="=Time()"
                    Format ="Long Time"
                    FontName ="Calibri"
                    GridlineColor =10921638

                    LayoutCachedLeft =5889
                    LayoutCachedTop =360
                    LayoutCachedWidth =7524
                    LayoutCachedHeight =660
                    BackShade =95.0
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                End
            End
        End
        Begin Section
            Height =3066
            Name ="Tiedot"
            AutoHeight =1
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin Rectangle
                    BackStyle =1
                    OverlapFlags =93
                    Left =283
                    Top =56
                    Width =7257
                    Height =3005
                    BackColor =14277081
                    BorderColor =10921638
                    Name ="Ruutu42"
                    GridlineColor =10921638
                    LayoutCachedLeft =283
                    LayoutCachedTop =56
                    LayoutCachedWidth =7540
                    LayoutCachedHeight =3061
                    BackShade =85.0
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =215
                    TextFontCharSet =177
                    TextAlign =1
                    TextFontFamily =0
                    IMESentenceMode =3
                    Left =4224
                    Top =360
                    Width =3276
                    Height =300
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="korttinro"
                    ControlSource ="=[Forms]![Tervetuloa]![Korttivalinta]"
                    FontName ="Calibri"
                    GroupTable =2
                    GridlineColor =10921638

                    LayoutCachedLeft =4224
                    LayoutCachedTop =360
                    LayoutCachedWidth =7500
                    LayoutCachedHeight =660
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
                            Left =360
                            Top =360
                            Width =3780
                            Height =300
                            BackColor =15921906
                            BorderColor =8355711
                            ForeColor =6710886
                            Name ="Selite5"
                            Caption ="Olet poistamassa tämän kortin linkitystä: "
                            FontName ="Calibri"
                            GroupTable =2
                            GridlineColor =10921638
                            LayoutCachedLeft =360
                            LayoutCachedTop =360
                            LayoutCachedWidth =4140
                            LayoutCachedHeight =660
                            LayoutGroup =2
                            BackShade =95.0
                            GroupTable =2
                        End
                    End
                End
                Begin TextBox
                    BorderWidth =2
                    OverlapFlags =215
                    TextFontCharSet =177
                    TextAlign =1
                    TextFontFamily =0
                    IMESentenceMode =3
                    Left =4224
                    Top =744
                    Width =3276
                    Height =336
                    TabIndex =1
                    ForeColor =4210752
                    Name ="Puumerkki"
                    FontName ="Calibri"
                    OnChange ="[Event Procedure]"
                    GroupTable =2
                    GridlineColor =10921638

                    LayoutCachedLeft =4224
                    LayoutCachedTop =744
                    LayoutCachedWidth =7500
                    LayoutCachedHeight =1080
                    RowStart =1
                    RowEnd =1
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
                            Left =360
                            Top =744
                            Width =3780
                            Height =336
                            BackColor =15921906
                            BorderColor =8355711
                            ForeColor =6710886
                            Name ="Selite15"
                            Caption ="Puumerkki"
                            FontName ="Calibri"
                            GroupTable =2
                            GridlineColor =10921638
                            LayoutCachedLeft =360
                            LayoutCachedTop =744
                            LayoutCachedWidth =4140
                            LayoutCachedHeight =1080
                            RowStart =1
                            RowEnd =1
                            LayoutGroup =2
                            BackShade =95.0
                            GroupTable =2
                        End
                    End
                End
                Begin TextBox
                    BorderWidth =2
                    OverlapFlags =215
                    TextFontCharSet =177
                    TextAlign =1
                    TextFontFamily =0
                    IMESentenceMode =3
                    Left =4224
                    Top =1164
                    Width =3276
                    Height =1224
                    TabIndex =2
                    ForeColor =4210752
                    Name ="Muistiinpano"
                    FontName ="Calibri"
                    OnChange ="[Event Procedure]"
                    GroupTable =2
                    GridlineColor =10921638

                    LayoutCachedLeft =4224
                    LayoutCachedTop =1164
                    LayoutCachedWidth =7500
                    LayoutCachedHeight =2388
                    RowStart =2
                    RowEnd =2
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
                            Left =360
                            Top =1164
                            Width =3780
                            Height =1224
                            BackColor =15921906
                            BorderColor =8355711
                            ForeColor =6710886
                            Name ="Muistiinpano_selite"
                            Caption ="Muistiinpanot"
                            FontName ="Calibri"
                            GroupTable =2
                            GridlineColor =10921638
                            LayoutCachedLeft =360
                            LayoutCachedTop =1164
                            LayoutCachedWidth =4140
                            LayoutCachedHeight =2388
                            RowStart =2
                            RowEnd =2
                            LayoutGroup =2
                            BackShade =95.0
                            GroupTable =2
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =215
                    TextFontCharSet =177
                    TextFontFamily =0
                    Left =360
                    Top =2460
                    Width =3780
                    Height =576
                    TabIndex =3
                    ForeColor =4210752
                    Name ="Poista"
                    Caption ="Poista linkitys"
                    OnClick ="[Event Procedure]"
                    FontName ="Calibri"
                    GroupTable =2
                    GridlineColor =10921638

                    LayoutCachedLeft =360
                    LayoutCachedTop =2460
                    LayoutCachedWidth =4140
                    LayoutCachedHeight =3036
                    RowStart =3
                    RowEnd =3
                    LayoutGroup =2
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
                    OverlapFlags =215
                    TextFontCharSet =177
                    TextFontFamily =0
                    Left =4224
                    Top =2460
                    Width =3276
                    Height =576
                    TabIndex =4
                    ForeColor =4210752
                    Name ="Komento35"
                    Caption ="Sulje"
                    FontName ="Calibri"
                    ControlTipText ="Sulje lomake"
                    GroupTable =2
                    GridlineColor =10921638
                    OnClickEmMacro = Begin
                        Version =196611
                        ColumnsShown =8
                        Begin
                            Action ="Close"
                            Argument ="-1"
                            Argument =""
                            Argument ="0"
                        End
                        Begin
                            Comment ="_AXL:<?xml version=\"1.0\" encoding=\"UTF-16\" standalone=\"no\"?>\015\012<UserI"
                                "nterfaceMacro For=\"Komento35\" xmlns=\"http://schemas.microsoft.com/office/acce"
                                "ssservices/2009/11/application\"><Statements><Action Name=\"CloseWindow\"/></Sta"
                                "tements></UserInterfaceMacro>"
                        End
                    End

                    LayoutCachedLeft =4224
                    LayoutCachedTop =2460
                    LayoutCachedWidth =7500
                    LayoutCachedHeight =3036
                    RowStart =3
                    RowEnd =3
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =2
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
                Begin Label
                    BackStyle =1
                    OverlapFlags =215
                    Left =340
                    Top =56
                    Width =7030
                    Height =284
                    BackColor =62207
                    BorderColor =8355711
                    ForeColor =6710886
                    Name ="Selite40"
                    Caption ="Paina enter, tab tai klikkaa toista kenttää päästäksesi eteenpäin!"
                    FontName ="Calibri"
                    GridlineColor =10921638
                    LayoutCachedLeft =340
                    LayoutCachedTop =56
                    LayoutCachedWidth =7370
                    LayoutCachedHeight =340
                    BackThemeColorIndex =-1
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


    
Private Sub Form_Open(Cancel As Integer)
    [Form_PoistaKortinLinkitys].Poista.Visible = False
    [Form_PoistaKortinLinkitys].Muistiinpano.Visible = False
    
End Sub


Private Sub Muistiinpano_Change()
    [Form_PoistaKortinLinkitys].Poista.Visible = True
End Sub

Private Sub Poista_Click()
    Dim cardnumber As String
    Dim cardID As Integer
    Dim SQLQuery As String
    Dim Puumerkki As String
    Dim Muistiinpano As String
    
    If IsNull(Form_Tervetuloa.Korttivalinta) Then
        MsgBox ("Korttia ei valittu. Valitse kortti pääikkunassa!")
        Exit Sub
    Else
        cardnumber = Form_Tervetuloa.Korttivalinta.Value
    End If
    
    If IsNull([Form_PoistaKortinLinkitys].Puumerkki) Then
        MsgBox ("Puumerkki ei voi linkitystä poistaessa olla tyhjä!")
        Exit Sub
    Else
        Puumerkki = [Form_PoistaKortinLinkitys].Puumerkki.Value
    End If
    
    
    If IsNull([Form_PoistaKortinLinkitys].Muistiinpano) Then
        MsgBox ("Muistiinpanokenttä ei voi linkitystä poistaessa olla tyhjä!")
        Exit Sub
    Else
        Muistiinpano = [Form_PoistaKortinLinkitys].Muistiinpano.Value
    End If
        
    
    Dim newOwner As Integer
    newOwner = 0 ' kortille vaan määritellään omistajaksi 0 eli nobody...
    
    Dim korttiID As Integer
    korttiID = Common.FetchGeneralID("Kortit", "CID", "Kortti = '" & cardnumber & "'")
    
    Dim success As Boolean
    Dim table As String
    Dim values As String
    Dim target As String
    
    table = "Kortit"
    values = "Omistaja = " & newOwner & ", " _
    & "PVM = '" & Now() & "', " _
    & "Puumerkki = '" & Puumerkki & "', " _
    & "Muistiinpanot = '" & Muistiinpano & "' "
    
    target = "Kortti = '" & cardnumber & "'"
    
    success = Common.InsertOrUpdate(table, values, target)
    
    If Not (success) Then
        MsgBox ("Jotain meni pieleen sori siitä!")
    End If
    
    
    Dim logOutput As String
    logOutput = "Puumerkki " & Puumerkki & " poisti kortin " & cardnumber & " linkityksen, muistiinpanot: " & Muistiinpano
    success = Common.SaveToLog(logOutput)
    
    success = Common.SendMessageToMainScreen("Kortin " & cardnumber & " linkitys poistettu!")
    
    DoCmd.Close
    
End Sub





Private Sub Puumerkki_Change()
    [Form_PoistaKortinLinkitys].Muistiinpano.Visible = True
End Sub
