VERSION 5.00
Begin VB.Form frmControl 
   Caption         =   "Form2"
   ClientHeight    =   2325
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   3630
   LinkTopic       =   "Form2"
   ScaleHeight     =   2325
   ScaleWidth      =   3630
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu mnuPop 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu mnuClose 
         Caption         =   "Salir"
      End
   End
End
Attribute VB_Name = "frmControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public f_objClevo                   As clsClevo
Private f_objSettings               As clsDumbParser

Private f_objSystrayPopup           As ClsPopUpTray

Private WithEvents f_objSystray     As cSystray
Attribute f_objSystray.VB_VarHelpID = -1


Private Sub Form_Load()
    Set f_objClevo = New clsClevo
    Set f_objSettings = New clsDumbParser


    Set f_objSystray = New cSystray
    Set f_objSystrayPopup = New ClsPopUpTray
    With f_objSystray
        .SysTrayToolTip = "Cocus Keyboard LED Control"
        .SysTrayIconFromFile App.path & "\lightbulbon.ico"
        Call .SysTrayShow(True)
    End With

    'Me.Visible = False
    Me.Hide
    'Me.Visible = True
    
    Load frmPopup
    
    
    Call LoadSettings
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveSettings
    Unload frmPopup
End Sub

Private Sub f_objSystray_MouseUp(Button As Integer)
    Static bShowingPopup                As Boolean
    If Button = 1 Then
        If bShowingPopup Then
            Exit Sub
        End If

        bShowingPopup = True
        f_objSystrayPopup.ShowPopUp frmPopup.hwnd ' Me.Visible = Not Me.Visible
        frmPopup.Visible = False
        'Call Unload(frmPopup)
        
        bShowingPopup = False
    ElseIf Button = 2 Then
        Call Me.PopupMenu(mnuPop)
    End If
End Sub

Private Sub mnuClose_Click()
    Call Unload(Me)
End Sub









Private Sub LoadSettings()
    With f_objSettings
        If Not .LoadFile("config.ini") Then
            Call SetDefaults
            Exit Sub
        End If

        frmPopup.ucPalette(0).Color = CLng("&H" & .DumbString(.Key("keyboard_leds", "section_left")))
        frmPopup.ucPalette(1).Color = CLng("&H" & .DumbString(.Key("keyboard_leds", "section_mid")))
        frmPopup.ucPalette(2).Color = CLng("&H" & .DumbString(.Key("keyboard_leds", "section_right")))
        frmPopup.ucPalette(3).Color = CLng("&H" & .DumbString(.Key("keyboard_leds", "section_lightbar")))
        
        If (frmPopup.ucSliderBrightness.Value = .DumbLong(.Key("keyboard_leds", "brightness")) / 100) Then
            f_objClevo.KeyboardBrightness = .DumbLong(.Key("keyboard_leds", "brightness")) / 100
        End If
        
        frmPopup.ucSliderBrightness.Value = .DumbLong(.Key("keyboard_leds", "brightness")) / 100
        
        '// Just toggle the values if the value is true
        If .DumbBool(.Key("keyboard_leds", "linked")) Then
            frmPopup.ucMetroAllColorsLinked.CheckValue = True
        End If
        
        frmPopup.ucMetroMaxFan.CheckValue = .DumbBool(.Key("fan", "max"))
    End With
End Sub

Private Sub SaveSettings()
    With f_objSettings
        .Key("keyboard_leds", "section_left") = .StringDumb(Hex(frmPopup.ucPalette(0).Color))
        .Key("keyboard_leds", "section_mid") = .StringDumb(Hex(frmPopup.ucPalette(1).Color))
        .Key("keyboard_leds", "section_right") = .StringDumb(Hex(frmPopup.ucPalette(2).Color))
        .Key("keyboard_leds", "section_lightbar") = .StringDumb(Hex(frmPopup.ucPalette(3).Color))

        .Key("keyboard_leds", "brightness") = .LongDumb(Round(frmPopup.ucSliderBrightness.Value * 100, 0))
        
        .Key("keyboard_leds", "linked") = .BoolDumb(frmPopup.ucMetroAllColorsLinked.CheckValue)
        
        .Key("fan", "max") = .BoolDumb(frmPopup.ucMetroMaxFan.CheckValue)

        Call .SaveFile("config.ini")
    End With
End Sub

Private Sub SetDefaults()
    Dim oColor          As OLE_COLOR
    oColor = vbBlue

    With f_objClevo
        .KeyboardBrightness = 0
        .KeyboardColor(KEYBOARD_PART_LEFT) = oColor
        .KeyboardColor(KEYBOARD_PART_MID) = oColor
        .KeyboardColor(KEYBOARD_PART_RIGHT) = oColor
        .KeyboardColor(KEYBOARD_PART_LIGHTBAR) = oColor
        .FanMax = False
    End With
End Sub
