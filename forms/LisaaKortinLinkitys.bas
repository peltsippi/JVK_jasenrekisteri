Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    Modal = NotDefault
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
    Width =5592
    DatasheetFontHeight =11
    ItemSuffix =33
    Left =4044
    Top =3468
    Right =17484
    Bottom =11712
    Picture ="bulldog_pienempi"
    RecSrcDt = Begin
        0xe76ced057dc5e540
    End
    RecordSource ="Hae nimi"
    Caption ="Linkitä kortti yhteystiedolle"
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
                    Left =36
                    Top =60
                    Width =2856
                    Height =456
                    FontSize =18
                    BackColor =14277081
                    Name ="Automaattinen_ylätunniste0"
                    Caption ="Lisää kortin linkitys"
                    FontName ="Calibri Light"
                    GroupTable =1
                    GridlineColor =10921638
                    LayoutCachedLeft =36
                    LayoutCachedTop =60
                    LayoutCachedWidth =2892
                    LayoutCachedHeight =516
                    LayoutGroup =1
                    ThemeFontIndex =0
                    BackShade =85.0
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
                    Left =4305
                    Top =60
                    Width =1287
                    Height =300
                    BackColor =14277081
                    BorderColor =10921638
                    Name ="Automaattinen_päivämäärä"
                    ControlSource ="=Date()"
                    Format ="Short Date"
                    FontName ="Calibri"
                    GridlineColor =10921638

                    LayoutCachedLeft =4305
                    LayoutCachedTop =60
                    LayoutCachedWidth =5592
                    LayoutCachedHeight =360
                    BackShade =85.0
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
                    Left =4305
                    Top =360
                    Width =1287
                    Height =300
                    TabIndex =1
                    BackColor =14277081
                    BorderColor =10921638
                    Name ="Automaattinen_aika"
                    ControlSource ="=Time()"
                    Format ="Long Time"
                    FontName ="Calibri"
                    GridlineColor =10921638

                    LayoutCachedLeft =4305
                    LayoutCachedTop =360
                    LayoutCachedWidth =5592
                    LayoutCachedHeight =660
                    BackShade =85.0
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                End
            End
        End
        Begin Section
            Height =4251
            Name ="Tiedot"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin Rectangle
                    SpecialEffect =0
                    BackStyle =1
                    OverlapFlags =93
                    Left =283
                    Top =166
                    Width =4479
                    Height =3972
                    BackColor =12566463
                    Name ="Ruutu32"
                    GridlineColor =10921638
                    LayoutCachedLeft =283
                    LayoutCachedTop =166
                    LayoutCachedWidth =4762
                    LayoutCachedHeight =4138
                    BackShade =75.0
                    BorderThemeColorIndex =0
                    BorderShade =100.0
                End
                Begin TextBox
                    BorderWidth =3
                    OverlapFlags =215
                    TextFontCharSet =177
                    TextAlign =1
                    TextFontFamily =0
                    IMESentenceMode =3
                    Left =2136
                    Top =1128
                    Width =1704
                    Height =1116
                    ForeColor =4210752
                    Name ="Korttinro"
                    ValidationRule ="Like \"????\""
                    ValidationText ="Katso ohje vasemmalta!"
                    FontName ="Calibri"
                    OnChange ="[Event Procedure]"
                    GroupTable =2
                    BottomPadding =150
                    GridlineColor =10921638

                    LayoutCachedLeft =2136
                    LayoutCachedTop =1128
                    LayoutCachedWidth =3840
                    LayoutCachedHeight =2244
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
                            Top =1128
                            Width =1704
                            Height =1116
                            BackColor =14277081
                            BorderColor =8355711
                            ForeColor =6710886
                            Name ="Selite3"
                            Caption ="Kortin numero\015\012HUOM! \015\012Aina 4 numeroa!!!\015\012Esim: 0056"
                            FontName ="Calibri"
                            GroupTable =2
                            BottomPadding =150
                            GridlineColor =10921638
                            LayoutCachedLeft =336
                            LayoutCachedTop =1128
                            LayoutCachedWidth =2040
                            LayoutCachedHeight =2244
                            LayoutGroup =2
                            BackShade =85.0
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
                    Left =2136
                    Top =2460
                    Width =1704
                    Height =336
                    TabIndex =1
                    ForeColor =4210752
                    Name ="Puumerkki"
                    FontName ="Calibri"
                    OnChange ="[Event Procedure]"
                    GroupTable =2
                    BottomPadding =150
                    GridlineColor =10921638

                    LayoutCachedLeft =2136
                    LayoutCachedTop =2460
                    LayoutCachedWidth =3840
                    LayoutCachedHeight =2796
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
                            Left =336
                            Top =2460
                            Width =1704
                            Height =336
                            BackColor =14277081
                            BorderColor =8355711
                            ForeColor =6710886
                            Name ="Selite7"
                            Caption ="Puumerkki"
                            FontName ="Calibri"
                            GroupTable =2
                            BottomPadding =150
                            GridlineColor =10921638
                            LayoutCachedLeft =336
                            LayoutCachedTop =2460
                            LayoutCachedWidth =2040
                            LayoutCachedHeight =2796
                            RowStart =1
                            RowEnd =1
                            LayoutGroup =2
                            BackShade =85.0
                            GroupTable =2
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =215
                    TextFontCharSet =177
                    TextFontFamily =0
                    Left =340
                    Top =3118
                    Width =1764
                    Height =852
                    TabIndex =2
                    ForeColor =4210752
                    Name ="Linkita"
                    Caption ="Linkitä"
                    OnClick ="[Event Procedure]"
                    FontName ="Calibri"
                    GridlineColor =10921638

                    LayoutCachedLeft =340
                    LayoutCachedTop =3118
                    LayoutCachedWidth =2104
                    LayoutCachedHeight =3970
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
                Begin CommandButton
                    OverlapFlags =215
                    TextFontCharSet =177
                    TextFontFamily =0
                    Left =2834
                    Top =3118
                    Width =1764
                    Height =912
                    TabIndex =3
                    ForeColor =4210752
                    Name ="sulje"
                    Caption ="Sulje"
                    FontName ="Calibri"
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
                                "nterfaceMacro For=\"sulje\" xmlns=\"http://schemas.microsoft.com/office/accessse"
                                "rvices/2009/11/application\"><Statements><Action Name=\"CloseWindow\"/></Stateme"
                                "nts></UserInterfaceMacro>"
                        End
                    End

                    LayoutCachedLeft =2834
                    LayoutCachedTop =3118
                    LayoutCachedWidth =4598
                    LayoutCachedHeight =4030
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
                Begin Label
                    BackStyle =1
                    OverlapFlags =215
                    TextFontCharSet =177
                    TextFontFamily =0
                    Left =396
                    Top =170
                    Width =4082
                    Height =680
                    BackColor =62207
                    BorderColor =8355711
                    ForeColor =6710886
                    Name ="Selite30"
                    Caption ="Paina enter, tab tai klikkaa toiseen kenttään jos ei meinaa edetä!"
                    FontName ="Calibri"
                    GridlineColor =10921638
                    LayoutCachedLeft =396
                    LayoutCachedTop =170
                    LayoutCachedWidth =4478
                    LayoutCachedHeight =850
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
    [Form_LisaaKortinLinkitys].Linkita.Visible = False
    [Form_LisaaKortinLinkitys].Puumerkki.Visible = False
    
End Sub


Private Sub Puumerkki_Change()
    [Form_LisaaKortinLinkitys].Linkita.Visible = True
End Sub


Private Sub Korttinro_Change()
    [Form_LisaaKortinLinkitys].Puumerkki.Visible = True
End Sub

Private Sub Linkita_Click()

    Dim userNumber As Integer
    Dim kortinNro As String
    Dim Puumerkki As String
    
    If IsNull(Form_Tervetuloa.Yhteystietovalinta) Then
        MsgBox ("Yhteystietoa ei valittu. Valitse yhteystieto pääikkunassa!")
        Exit Sub
    Else
        userNumber = Form_Tervetuloa.Yhteystietovalinta.Value
    End If
    
    If IsNull([Form_LisaaKortinLinkitys].korttinro) Then
        MsgBox ("Korttinumeroa ei annettu, yritä uudestaan!")
        Exit Sub
    Else
        kortinNro = [Form_LisaaKortinLinkitys].korttinro.Value
    '    MsgBox (kortinNro)
    End If
    
    If IsNull([Form_LisaaKortinLinkitys].Puumerkki) Then
        MsgBox ("Puumerkki ei voi olla tyhjä, yritä uudestaan!")
        Exit Sub
    Else
        Puumerkki = [Form_LisaaKortinLinkitys].Puumerkki.Value
    End If
    
    
    
    'Check if there is entry for the card
    Dim criteria As String
    criteria = "Kortti = '" & kortinNro & "'"
    Dim recordQty As Integer
    recordQty = Common.CheckIfRecordFound("Kortit", criteria)
    'MsgBox (recordQty)
    
    
    Dim valuelist As String
    Dim update As Boolean
    
    valuelist = "Kortti = '" & kortinNro & "' , " _
    & "Omistaja = " & userNumber & " , " _
    & "PVM = '" & Now() & "' , " _
    & "Puumerkki = '" & Puumerkki & "'"
    
    
    If (recordQty = 0) Then
        valuelist = valuelist & " , Muistiinpanot = 'Kortti lisätty ensimmäistä kertaa'"
        update = False
        
    Else
        Dim cardID As Integer
        cardID = Common.FetchCardID(kortinNro)
        
        Dim previousOwner As Integer
        previousOwner = Common.FetchGeneralID("Kortit", "Omistaja", criteria)
        'MsgBox (previousOwner)
        
        If Not (previousOwner = 0) Then
            MsgBox ("Kortilla on jo omistaja, poista kortin linkitys ensin vanhalta omistajalta jos numero meni oikein!")
            Exit Sub
        Else
        
            valuelist = valuelist & ", Muistiinpanot = 'Lisätty vanha omistajaton kortti uudelle omistajalle'"
            update = True
            
        End If
        
 

            
        
    End If
    'MsgBox (valuelist)
    Dim succs
    If (update) Then
        succs = Common.InsertOrUpdate("Kortit", valuelist, criteria)
    Else
        succs = Common.InsertOrUpdate("Kortit", valuelist, "")
    End If
    
    succs = Common.SaveToLog("Kortti " & kortinNro & " linkitetty omistajalle: " & userNumber)
    succs = Common.SendMessageToMainScreen("Kortti " & kortinNro & " linkitetty uudelle omistajalle!")
    DoCmd.Close
    
    

End Sub
