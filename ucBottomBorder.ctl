VERSION 5.00
Begin VB.UserControl ucBottomBorder 
   Alignable       =   -1  'True
   BackColor       =   &H00FBF5F1&
   ClientHeight    =   210
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   210
   ControlContainer=   -1  'True
   ScaleHeight     =   210
   ScaleWidth      =   210
   Begin VB.Line linLines 
      BorderColor     =   &H00F9F2ED&
      Index           =   3
      X1              =   0
      X2              =   3120
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line linLines 
      BorderColor     =   &H00F7EEE8&
      Index           =   2
      X1              =   0
      X2              =   3120
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line linLines 
      BorderColor     =   &H00F0E3D9&
      Index           =   1
      X1              =   0
      X2              =   3120
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line linLines 
      BorderColor     =   &H00EAD9CC&
      Index           =   0
      X1              =   0
      X2              =   3120
      Y1              =   0
      Y2              =   0
   End
End
Attribute VB_Name = "ucBottomBorder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Sub UserControl_Resize()
    Dim lngIndex As Long
    
    For lngIndex = 0 To 3
        With linLines(lngIndex)
            .X1 = 0
            .X2 = UserControl.ScaleWidth
            .Y1 = lngIndex * 15
            .Y2 = .Y1
        End With
    Next lngIndex
End Sub
