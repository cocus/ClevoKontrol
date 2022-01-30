VERSION 5.00
Begin VB.Form frmPopup 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3384
   ClientLeft      =   -36
   ClientTop       =   -384
   ClientWidth     =   5880
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   8.4
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   282
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   490
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picBrightnessLevel 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   372
      Left            =   5040
      ScaleHeight     =   31
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   51
      TabIndex        =   5
      Top             =   1680
      Width           =   612
   End
   Begin ClevoKontrol.ucMetroSlider ucSliderBrightness 
      Height          =   372
      Left            =   840
      TabIndex        =   1
      Top             =   1704
      Width           =   3972
      _ExtentX        =   7006
      _ExtentY        =   656
      BackColor       =   -2147483635
      ForeColor       =   -2147483630
      Value           =   0
   End
   Begin ClevoKontrol.ucButtonMetro ucMetroAllColorsLinked 
      Height          =   972
      Left            =   1440
      TabIndex        =   7
      Top             =   2280
      Width           =   1332
      _ExtentX        =   2350
      _ExtentY        =   1715
      BackColor       =   -2147483633
      Caption         =   "Link Colors"
      CheckStyle      =   -1  'True
      CheckValue      =   0   'False
   End
   Begin ClevoKontrol.ucPalette ucPalette 
      Height          =   1332
      Index           =   3
      Left            =   4440
      TabIndex        =   6
      Top             =   120
      Width           =   1332
      _ExtentX        =   2350
      _ExtentY        =   2350
   End
   Begin VB.Image imgBrightnessIcon 
      Height          =   384
      Left            =   240
      Top             =   1680
      Width           =   384
   End
   Begin ClevoKontrol.ucPalette ucPalette 
      Height          =   1332
      Index           =   2
      Left            =   3000
      TabIndex        =   4
      Top             =   120
      Width           =   1332
      _ExtentX        =   2350
      _ExtentY        =   2350
   End
   Begin ClevoKontrol.ucPalette ucPalette 
      Height          =   1332
      Index           =   1
      Left            =   1560
      TabIndex        =   3
      Top             =   120
      Width           =   1332
      _ExtentX        =   2350
      _ExtentY        =   2350
   End
   Begin ClevoKontrol.ucPalette ucPalette 
      Height          =   1332
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1332
      _ExtentX        =   2350
      _ExtentY        =   2350
   End
   Begin ClevoKontrol.ucButtonMetro ucMetroMaxFan 
      Height          =   972
      Left            =   120
      TabIndex        =   0
      Top             =   2280
      Width           =   1332
      _ExtentX        =   2350
      _ExtentY        =   1715
      BackColor       =   -2147483633
      Caption         =   "Max Fan"
      CheckStyle      =   -1  'True
      CheckValue      =   0   'False
   End
End
Attribute VB_Name = "frmPopup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private f_cRenderer                 As New stdPicEx2
Private f_objGlass                  As cVistaGlass
Private f_objText                   As cVistaText

Private WithEvents f_objListener    As clsSettingChangedListener
Attribute f_objListener.VB_VarHelpID = -1

Private Sub LoadPalettes()
    Dim lIndex          As Long
    For lIndex = 0 To ucPalette.UBound
        With ucPalette(lIndex)
            .LoadImagePaletteFromFile App.path & "\colorwheel.png"
            '.Color = vbBlue
        End With
    Next lIndex
End Sub

Private Sub SetImmersiveColor()
    With f_objGlass
        If .IsGlassEnabled Then
            .EnableAcrylicBlurForWin10 Me.hwnd, GetImmersiveColorByEnum([ImmersiveStartBackground]) And &HFFFFFF, 205 '&HD9

            Me.BackColor = vbBlack
            picBrightnessLevel.BackColor = vbBlack
            ucSliderBrightness.BackColor = vbBlack
        End If
    End With


    ucMetroMaxFan.BackColor = GetImmersiveColorByEnum([ImmersiveStartBackground]) And &HFFFFFF
    ucMetroAllColorsLinked.BackColor = ucMetroMaxFan.BackColor
    ucSliderBrightness.ForeColor = GetImmersiveColorByEnum([ImmersiveLightHighlight]) And &HFFFFFF
End Sub






Private Sub f_objListener_SettingChanged(ByVal lParam As Long, ByVal wParam As Long)
    Call SetImmersiveColor
End Sub






Private Sub Form_Load()

    Set f_objListener = New clsSettingChangedListener
    Set f_objGlass = New cVistaGlass
    Set f_objText = New cVistaText
   


    ucMetroMaxFan.SetImageFromFromFile App.path & "\fan.png"
    ucMetroAllColorsLinked.SetImageFromFromFile App.path & "\link20.png"
    
    Set imgBrightnessIcon.Picture = f_cRenderer.LoadPictureEx(App.path & "\brightness.png")
    
    Call LoadPalettes
    Call SetImmersiveColor

    With f_objText
        .Attach picBrightnessLevel.hwnd
        .SetData "0", picBrightnessLevel.ForeColor, picBrightnessLevel.BackColor, 5, 0, 0, (picBrightnessLevel.ScaleWidth), picBrightnessLevel.ScaleHeight, picBrightnessLevel.Font, ALIGN_CENTER
    End With

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    'Call DragForm(Me)
End Sub






Private Sub imgBrightnessIcon_Click()
    Static dPreviousValue As Double
    
    Dim dTemp       As Double
    
    dTemp = ucSliderBrightness.Value
    ucSliderBrightness.Value = dPreviousValue
    dPreviousValue = dTemp
End Sub

Private Sub picBrightnessLevel_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call DragForm(Me)
End Sub

Private Sub ucMetroAllColorsLinked_Click()
    Static oColors(3)   As OLE_COLOR
    Dim lIndex          As Long
    
    For lIndex = 1 To ucPalette.UBound
        ucPalette(lIndex).Visible = Not ucMetroAllColorsLinked.CheckValue
    Next lIndex
    
    If ucMetroAllColorsLinked.CheckValue Then
        For lIndex = 0 To ucPalette.UBound
            oColors(lIndex) = ucPalette(lIndex).Color
        Next lIndex
        
        Dim oColor              As OLE_COLOR
        oColor = ucPalette(0).Color
        frmControl.f_objClevo.KeyboardColor(KEYBOARD_PART_LEFT) = oColor
        frmControl.f_objClevo.KeyboardColor(KEYBOARD_PART_MID) = oColor
        frmControl.f_objClevo.KeyboardColor(KEYBOARD_PART_RIGHT) = oColor
        frmControl.f_objClevo.KeyboardColor(KEYBOARD_PART_LIGHTBAR) = oColor
    Else
        For lIndex = 0 To ucPalette.UBound
            Call ucPalette_ColorSelect(CInt(lIndex), oColors(lIndex))
        Next lIndex
    End If
End Sub

Private Sub ucMetroMaxFan_Click()
    frmControl.f_objClevo.FanMax = ucMetroMaxFan.CheckValue
End Sub

Private Sub ucSliderBrightness_Scroll()
    Debug.Print ucSliderBrightness.Value * 255
    frmControl.f_objClevo.KeyboardBrightness = ucSliderBrightness.Value * 255
    f_objText.SetData Round(ucSliderBrightness.Value * 100, 0), picBrightnessLevel.ForeColor, picBrightnessLevel.BackColor, 5, 0, 0, (picBrightnessLevel.ScaleWidth), picBrightnessLevel.ScaleHeight, picBrightnessLevel.Font, ALIGN_CENTER
End Sub

Private Sub ucPalette_ColorSelect(Index As Integer, oColor As stdole.OLE_COLOR)
    If ucMetroAllColorsLinked.CheckValue Then
        frmControl.f_objClevo.KeyboardColor(KEYBOARD_PART_LEFT) = oColor
        frmControl.f_objClevo.KeyboardColor(KEYBOARD_PART_MID) = oColor
        frmControl.f_objClevo.KeyboardColor(KEYBOARD_PART_RIGHT) = oColor
        frmControl.f_objClevo.KeyboardColor(KEYBOARD_PART_LIGHTBAR) = oColor
    Else
        Select Case Index
            Case 0: frmControl.f_objClevo.KeyboardColor(KEYBOARD_PART_LEFT) = oColor
            Case 1: frmControl.f_objClevo.KeyboardColor(KEYBOARD_PART_MID) = oColor
            Case 2: frmControl.f_objClevo.KeyboardColor(KEYBOARD_PART_RIGHT) = oColor
            Case 3: frmControl.f_objClevo.KeyboardColor(KEYBOARD_PART_LIGHTBAR) = oColor
        End Select
    End If
End Sub

