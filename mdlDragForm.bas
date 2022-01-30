Attribute VB_Name = "mdlDragForm"
Option Explicit

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal iparam As Long) As Long
Private Declare Sub ReleaseCapture Lib "user32" ()

Public Sub DragForm(ByRef fForm As Form)
    Call ReleaseCapture
    Call SendMessage(fForm.hwnd, &HA1, 2, 0&)
End Sub
