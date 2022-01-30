VERSION 5.00
Begin VB.UserControl ucLabelItem 
   BackStyle       =   0  'Transparent
   ClientHeight    =   1650
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2940
   ClipBehavior    =   0  'None
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H8000000D&
   HasDC           =   0   'False
   ScaleHeight     =   110
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   196
   Windowless      =   -1  'True
   Begin VB.Timer Timer1 
      Left            =   840
      Top             =   960
   End
End
Attribute VB_Name = "ucLabelItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Autor: Leandro Ascierto
'Web:   www.leandroascierto.com.ar
'Date:  21/06/2011
'Nota:  Este usercontrol no tiene en cuenta muchas cosas solo esta echo a medidas, tener en cuenta que DrawPopUpItemTheme es valido solo para window Vista y Seven
        'en caso de correr en xp se pude aplicar este tipo de selecion http://www.leandroascierto.com.ar/categoria/Gr%C3%A1ficos/articulo/DrawAlphaSelection.php
        
Private Declare Function SetRect Lib "user32.dll" (ByRef lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function DrawText Lib "user32.dll" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, ByRef lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function DrawFocusRect Lib "user32.dll" (ByVal hdc As Long, ByRef lpRect As RECT) As Long
Private Declare Function DrawIcon Lib "user32.dll" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal hIcon As Long) As Long
Private Declare Function GetIconInfo Lib "user32.dll" (ByVal hIcon As Long, ByRef piconinfo As ICONINFO) As Long
Private Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long
Private Declare Function DrawIconEx Lib "user32.dll" (ByVal hdc As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long
Private Declare Function DestroyIcon Lib "user32.dll" (ByVal hIcon As Long) As Long
Private Declare Function WindowFromPoint Lib "user32.dll" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function GetCursorPos Lib "user32.dll" (ByRef lpPoint As POINTAPI) As Long
Private Declare Function ClientToScreen Lib "user32.dll" (ByVal hwnd As Long, ByRef lpPoint As POINTAPI) As Long
Private Declare Function PtInRect Lib "user32.dll" (ByRef lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function DrawThemeBackground Lib "uxtheme.dll" (ByVal hTheme As Long, ByVal lhDC As Long, ByVal iPartId As Long, ByVal iStateId As Long, pRect As RECT, pClipRect As Any) As Long
Private Declare Function OpenThemeData Lib "uxtheme.dll" (ByVal hwnd As Long, ByVal pszClassList As Long) As Long
Private Declare Function CloseThemeData Lib "uxtheme.dll" (ByVal hTheme As Long) As Long
Private Declare Function LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long
Private Declare Function DestroyCursor Lib "user32" (ByVal hCursor As Long) As Long
Private Declare Function SetCursor Lib "user32.dll" (ByVal hCursor As Long) As Long
Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function CopyImage Lib "user32" (ByVal handle As Long, ByVal imageType As Long, ByVal newWidth As Long, ByVal newHeight As Long, ByVal lFlags As Long) As Long

Private Type ICONINFO
    fIcon As Long
    xHotspot As Long
    yHotspot As Long
    hbmMask As Long
    hbmColor As Long
End Type

Private Type BITMAP
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Const IDC_HAND          As Long = 32649
Private Const IMAGE_ICON        As Long = 1
Private Const LR_COPYFROMRESOURCE As Long = &H4000

Private Const MPI_HOT           As Long = 2
Private Const MENU_POPUPITEM    As Long = 14

Private Const DI_MASK           As Long = &H1
Private Const DI_IMAGE          As Long = &H2
Private Const DI_NORMAL         As Long = DI_MASK Or DI_IMAGE

Private Const DT_CALCRECT       As Long = &H400
Private Const DT_END_ELLIPSIS   As Long = &H8000
Private Const DT_SINGLELINE     As Long = &H20
Private Const DT_TOP            As Long = &H0
Private Const DT_VCENTER        As Long = &H4
Private Const DT_BOTTOM         As Long = &H8
Private Const DT_NOCLIP         As Long = &H100
Private Const DT_WORDBREAK      As Long = &H10

Public Enum EnuAlignment
    DT_LEFT = &H0
    DT_CENTER = &H1
    DT_RIGHT = &H2
End Enum

Public Event Click()
Public Event DblClick()
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event KeyPress(KeyAscii As Integer)
Public Event KeyUp(KeyCode As Integer, Shift As Integer)
Public Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event MouseEnter()
Public Event MouseLeave()

Private m_Caption       As String
Private m_Note          As String
Private m_Icon          As Long
Private II              As ICONINFO
Private m_IconWidth     As Long
Private m_IconHeight    As Long
Private c_tPT           As POINTAPI
Private c_lhWnd         As Long
Private bFocus          As Boolean
Private m_AutoSize      As Boolean
Private m_WordWrap      As Boolean
Private m_Alignment     As EnuAlignment
Private m_LinkStyle     As Boolean
Private hHandCursor     As Long
Private m_CenterIcon    As Boolean
Private m_ForeColor     As OLE_COLOR

Private Sub Timer1_Timer()
    If Not IsMouseInArea Then
        Timer1.Interval = 0
        UserControl.Refresh
        RaiseEvent MouseLeave
    End If
End Sub

Public Function IsMouseInArea() As Boolean
    Dim PT As POINTAPI
    Dim CPT As POINTAPI
    Dim TR As RECT
    Dim bArea As Boolean
    
    Call GetCursorPos(PT)
    Call ClientToScreen(c_lhWnd, CPT)
    
    CPT.x = PT.x - CPT.x - c_tPT.x
    CPT.y = PT.y - CPT.y - c_tPT.y


    Call SetRect(TR, 0, 0, UserControl.Width / 15, UserControl.Height / 15)
    bArea = PtInRect(TR, CPT.x, CPT.y)

    
    If bArea And WindowFromPoint(PT.x, PT.y) = c_lhWnd Then
        IsMouseInArea = True
    End If

End Function

Private Sub UserControl_HitTest(x As Single, y As Single, HitResult As Integer)
    HitResult = vbHitResultHit

    If Ambient.UserMode Then
        Dim PT  As POINTAPI

        Call GetCursorPos(c_tPT)
        Call ClientToScreen(c_lhWnd, PT)
        c_tPT.x = c_tPT.x - PT.x - x
        c_tPT.y = c_tPT.y - PT.y - y
        If Timer1.Interval = 0 Then
            UserControl.Refresh
            Timer1.Interval = 50
            RaiseEvent MouseEnter
        End If
        If m_LinkStyle Then
            SetCursor hHandCursor
        End If
    End If
End Sub

Private Sub UserControl_InitProperties()
    m_AutoSize = True
    m_WordWrap = True
    m_Alignment = DT_LEFT
    m_ForeColor = vbHighlight
    m_Caption = Ambient.DisplayName
    UserControl.Font.Name = "Segoe UI"
    UserControl.Font.Size = 8
End Sub

Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub UserControl_EnterFocus()
    bFocus = True
    UserControl.Refresh
End Sub

Private Sub UserControl_ExitFocus()
    bFocus = False
    UserControl.Refresh
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseDown(Button, Shift, x, y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseMove(Button, Shift, x, y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseUp(Button, Shift, x, y)
End Sub

Private Sub UserControl_Paint()
    Dim Rec As RECT
    Dim RecNote As RECT
    Dim MarginLeft As Long
    Dim MarginRight As Long
    Dim MarginTop As Long
    Dim MarginBottom As Long
    Dim lLeft As Long
    Dim lTop As Long
    Dim lHeight As Long
    
    If IsMouseInArea Then
        If m_LinkStyle Then
            UserControl.Font.Underline = True
        Else
            Call DrawPopUpItemTheme(UserControl.hdc, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight)
        End If
    Else
        If m_LinkStyle Then UserControl.Font.Underline = False
    End If

    MarginLeft = IIf(m_Icon, 8 + m_IconWidth, 4)
    MarginRight = 4
    MarginTop = 4
    MarginBottom = IIf(Len(m_Note), 8, 4)

    Rec.Right = UserControl.ScaleWidth - MarginLeft - MarginRight
    RecNote.Right = Rec.Right
    
    DrawText UserControl.hdc, m_Caption, -1, Rec, DT_CALCRECT Or DT_WORDBREAK
    DrawText UserControl.hdc, m_Note, -1, RecNote, DT_CALCRECT Or DT_WORDBREAK
    
    
    lHeight = IIf(Rec.Bottom > m_IconHeight, Rec.Bottom, m_IconHeight)
    
    lTop = ((UserControl.ScaleHeight - lHeight - RecNote.Bottom) / 2)
    
    Select Case m_Alignment
        Case DT_LEFT
            
            
            If m_CenterIcon Then
                DrawIconEx UserControl.hdc, 4, (UserControl.ScaleHeight - m_IconHeight) / 2, m_Icon, m_IconWidth, m_IconHeight, 0, 0, DI_NORMAL
                lTop = ((UserControl.ScaleHeight - Rec.Bottom - RecNote.Bottom) / 2)
                lHeight = Rec.Bottom
                SetRect Rec, MarginLeft, lTop, UserControl.ScaleWidth - MarginRight, lTop + Rec.Bottom
                lTop = lTop + lHeight
            Else
                DrawIconEx UserControl.hdc, 4, lTop, m_Icon, m_IconWidth, m_IconHeight, 0, 0, DI_NORMAL
                lTop = ((UserControl.ScaleHeight - Rec.Bottom - RecNote.Bottom) / 2)
                SetRect Rec, MarginLeft, lTop, UserControl.ScaleWidth - MarginRight, lTop + Rec.Bottom
                lTop = lHeight + ((UserControl.ScaleHeight - lHeight - RecNote.Bottom) / 2) + MarginTop
            End If
            
            SetRect RecNote, Rec.Left, lTop, Rec.Right, lTop + RecNote.Bottom

            
        Case DT_CENTER

            lLeft = ((UserControl.ScaleWidth - Rec.Right) / 2)
            DrawIconEx UserControl.hdc, lLeft - ((m_IconWidth + MarginRight) / 2), lTop, m_Icon, m_IconWidth, m_IconHeight, 0, 0, DI_NORMAL
            lTop = ((UserControl.ScaleHeight - Rec.Bottom - RecNote.Bottom) / 2)
            lLeft = lLeft + (m_IconWidth + MarginRight) / 2
            SetRect Rec, lLeft, lTop, lLeft + Rec.Right, lTop + Rec.Bottom
            lTop = lHeight + ((UserControl.ScaleHeight - lHeight - RecNote.Bottom) / 2) + MarginTop
            SetRect RecNote, 4, lTop, UserControl.ScaleWidth - 4, lTop + RecNote.Bottom
        
        
        Case DT_RIGHT
            
            
            DrawIconEx UserControl.hdc, UserControl.ScaleWidth - Rec.Right - MarginLeft, lTop, m_Icon, m_IconWidth, m_IconHeight, 0, 0, DI_NORMAL
            
            lTop = lTop + ((lHeight - Rec.Bottom) / 2)
            SetRect Rec, UserControl.ScaleWidth - Rec.Right - MarginRight, lTop, UserControl.ScaleWidth - MarginRight, lTop + Rec.Bottom
            
            lTop = MarginTop + lHeight + MarginTop
            SetRect RecNote, MarginLeft, lTop, UserControl.ScaleWidth - 4, lTop + RecNote.Bottom
    End Select
    
    UserControl.ForeColor = m_ForeColor
    DrawText UserControl.hdc, m_Caption, -1, Rec, DT_WORDBREAK Or m_Alignment
    
    UserControl.ForeColor = vbGrayText
    DrawText UserControl.hdc, m_Note, -1, RecNote, DT_WORDBREAK Or m_Alignment
       
    If bFocus Then
        SetRect Rec, 2, 2, UserControl.ScaleWidth - 2, UserControl.ScaleHeight - 2
        UserControl.ForeColor = 0
        DrawFocusRect UserControl.hdc, Rec
    End If
    'DrawIcon UserControl.hdc, 4, 4, m_Icon
    
    
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    c_lhWnd = UserControl.ContainerHwnd
    With PropBag
        m_Caption = .ReadProperty("Caption", Ambient.DisplayName)
        m_Note = .ReadProperty("Note", vbNullString)
        m_WordWrap = .ReadProperty("WordWrap", True)
        m_Alignment = .ReadProperty("Alignment", DT_LEFT)
        m_AutoSize = .ReadProperty("AutoSize", True)
        m_LinkStyle = .ReadProperty("LinkStyle", False)
        m_ForeColor = .ReadProperty("ForeColor", vbHighlight)
        Set UserControl.Font = .ReadProperty("Font", Ambient.Font)
    End With
    
    hHandCursor = LoadCursor(ByVal 0&, IDC_HAND)
 
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        .WriteProperty "Caption", m_Caption, Ambient.DisplayName
        .WriteProperty "Note", m_Note, vbNullString
        .WriteProperty "WordWrap", m_WordWrap, True
        .WriteProperty "Alignment", m_Alignment, DT_LEFT
        .WriteProperty "AutoSize", m_AutoSize, True
        .WriteProperty "LinkStyle", m_LinkStyle, False
        .WriteProperty "ForeColor", m_ForeColor, vbHighlight
        .WriteProperty "Font", UserControl.Font, Ambient.Font
    End With
End Sub

Private Sub UserControl_Resize()

    On Error GoTo ErrHandler
    
    Dim Rec As RECT
    Dim RecNote As RECT
    
    Dim lWidth As Long
    Dim lHeight As Long
    Dim MarginLeft As Long
    Dim MarginRight As Long
    Dim MarginTop As Long
    Dim MarginBottom As Long
    
    
    If m_WordWrap = False And m_AutoSize = False Then
        UserControl.Refresh
        Exit Sub
    End If

    MarginLeft = IIf(m_Icon, 8 + m_IconWidth, 4)
    MarginRight = IIf(Len(Caption), 4, 0)
    MarginTop = 4
    MarginBottom = IIf(Len(m_Note), 8, 4)

    If m_WordWrap Then
        Rec.Right = UserControl.ScaleWidth - MarginLeft - MarginRight
        RecNote.Right = Rec.Right
    End If
    
    DrawText UserControl.hdc, m_Caption, -1, Rec, DT_CALCRECT Or DT_WORDBREAK
    DrawText UserControl.hdc, m_Note, -1, RecNote, DT_CALCRECT Or DT_WORDBREAK
    
    If Len(m_Caption) + Len(Note) = 0 Then
        lWidth = MarginLeft + MarginRight
    Else
        lWidth = IIf(Rec.Right > RecNote.Right, Rec.Right, RecNote.Right) + MarginLeft + MarginRight
    End If
    
    If m_CenterIcon Then
        lHeight = IIf(m_IconHeight > Rec.Bottom + RecNote.Bottom, m_IconHeight - RecNote.Bottom, Rec.Bottom)
    Else
        lHeight = IIf(m_IconHeight > Rec.Bottom, m_IconHeight, Rec.Bottom)
    End If
    lHeight = lHeight + MarginTop + RecNote.Bottom + MarginBottom
       
    If m_AutoSize Then
        UserControl.Size lWidth * Screen.TwipsPerPixelX, lHeight * Screen.TwipsPerPixelY
    End If
    
    If m_WordWrap Then
        UserControl.Height = lHeight * Screen.TwipsPerPixelY
    End If
    
    UserControl.Refresh
    Exit Sub
    
ErrHandler:
    UserControl.Refresh
    'Debug.Print Err.Number, Err.Description
End Sub

Private Sub UserControl_Show()
    Me.Refresh
End Sub

Private Sub UserControl_Terminate()
    If m_Icon Then Call DestroyIcon(m_Icon)
    If hHandCursor Then Call DestroyCursor(hHandCursor)
End Sub

Private Function DrawPopUpItemTheme(ByVal DC As Long, ByVal x As Long, ByVal y As Long, ByVal Width As Long, ByVal Height As Long) As Boolean

    Dim hTheme  As Long
    Dim rtRect As RECT

    hTheme = OpenThemeData(0&, StrPtr("MENU"))

    If (hTheme) Then
        SetRect rtRect, x, y, x + Width, y + Height
        DrawPopUpItemTheme = DrawThemeBackground(hTheme, DC, MENU_POPUPITEM, MPI_HOT, rtRect, ByVal 0&) = 0
        
        Call CloseThemeData(hTheme)
    End If
    
End Function
Public Sub Refresh()
    UserControl_Resize
End Sub

Public Property Get Icon() As Long
    m_Icon = Icon
End Property

Public Property Let Icon(NewIcon As Long)
    Dim BMP As BITMAP
    If m_Icon Then Call DestroyIcon(m_Icon)
    GetIconInfo NewIcon, II
    
    If GetObject(II.hbmColor, Len(BMP), BMP) Then
        m_IconWidth = BMP.bmWidth
        m_IconHeight = BMP.bmHeight
    Else
        m_IconWidth = II.xHotspot
        m_IconHeight = II.yHotspot
    End If
    
    If II.hbmColor Then DeleteObject II.hbmColor
    If II.hbmMask Then DeleteObject II.hbmMask
    
    m_Icon = CopyImage(NewIcon, IMAGE_ICON, m_IconWidth, m_IconHeight, LR_COPYFROMRESOURCE)
    
    Me.Refresh
End Property

Property Get Caption() As String
    Caption = m_Caption
End Property
Property Let Caption(ByVal NewCaption As String)
    m_Caption = NewCaption
    Call PropertyChanged("Caption")
    Me.Refresh
End Property

Property Get Note() As String
    Note = m_Note
End Property
Property Let Note(ByVal NewNote As String)
    m_Note = NewNote
    Call PropertyChanged("Note")
    Me.Refresh
    
End Property

Property Get AutoSize() As Boolean
    AutoSize = m_AutoSize
End Property

Property Let AutoSize(ByVal NewValue As Boolean)
    m_AutoSize = NewValue
    Call PropertyChanged("AutoSize")
    Me.Refresh
End Property

Property Get Alignment() As EnuAlignment
    Alignment = m_Alignment
End Property

Property Let Alignment(ByVal NewValue As EnuAlignment)
    m_Alignment = NewValue
    Call PropertyChanged("Alignment")
    Me.Refresh
End Property

Property Get LinkStyle() As Boolean
    LinkStyle = m_LinkStyle
End Property

Property Let LinkStyle(ByVal NewValue As Boolean)
    m_LinkStyle = NewValue
    Call PropertyChanged("LinkStyle")
    Me.Refresh
End Property

Property Get CenterIcon() As Boolean
    CenterIcon = m_CenterIcon
End Property

Property Let CenterIcon(ByVal NewValue As Boolean)
    m_CenterIcon = NewValue
    Call PropertyChanged("CenterIcon")
    Me.Refresh
End Property

Property Get ForeColor() As OLE_COLOR
    ForeColor = m_ForeColor
End Property

Property Let ForeColor(ByVal NewColor As OLE_COLOR)
    m_ForeColor = NewColor
    Call PropertyChanged("ForeColor")
    Me.Refresh
End Property

Property Get Font() As StdFont
    Set Font = UserControl.Font
End Property

Property Set Font(ByVal NewFont As StdFont)
    Set UserControl.Font = NewFont
    Call PropertyChanged("Font")
    Me.Refresh
End Property

Property Get WordWrap() As Boolean
    WordWrap = m_WordWrap
End Property

Property Let WordWrap(ByVal NewValue As Boolean)
    m_WordWrap = NewValue
    Call PropertyChanged("WordWrap")
    Me.Refresh
End Property


