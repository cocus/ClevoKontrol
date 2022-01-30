Attribute VB_Name = "mdlMain"
Option Explicit

Private Declare Function LoadLibraryA Lib "kernel32" (ByVal szLib As String) As Long
Private Declare Function CallWindowProcA Lib "user32" (ByVal adr As Long, ByVal p1 As Long, ByVal p2 As Long, ByVal p3 As Long, ByVal p4 As Long) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal szFnc As String) As Long
Private Declare Function GetModuleHandleA Lib "kernel32" (ByVal szModule As String) As Long

Public Sub Main()
    Dim lngLib As Long
    Dim lngAdd As Long
    
    lngLib = GetModuleHandleA("comctl32.dll")
    
    If lngLib = 0 Then
        lngLib = LoadLibraryA("comctl32.dll")
    End If
    
    lngAdd = GetProcAddress(lngLib, "InitCommonControls")
    If lngAdd = 0 Then
        Exit Sub
    End If
    Call CallWindowProcA(lngAdd, 0, 0, 0, 0)
    
    Load frmControl
End Sub
