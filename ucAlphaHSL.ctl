VERSION 5.00
Begin VB.UserControl ucAlphaHSL 
   BackStyle       =   0  'Transparent
   ClientHeight    =   2436
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3480
   EditAtDesignTime=   -1  'True
   ScaleHeight     =   2436
   ScaleWidth      =   3480
   Begin VB.Shape shpDummy 
      FillColor       =   &H00993300&
      FillStyle       =   7  'Diagonal Cross
      Height          =   1812
      Left            =   0
      Shape           =   2  'Oval
      Top             =   0
      Visible         =   0   'False
      Width           =   2052
   End
End
Attribute VB_Name = "ucAlphaHSL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum eTargetType
    [eTargetImageBox] = 0
    [eTargetPictureBox] = 1
    [eTargetParentDirect] = 2
End Enum

Private u_eTarget                               As eTargetType

Private u_cRenderer                             As New stdPicEx2
Private WithEvents u_iImageInParent             As Image
Attribute u_iImageInParent.VB_VarHelpID = -1
Private WithEvents u_pImageInParent             As PictureBox
Attribute u_pImageInParent.VB_VarHelpID = -1

Private WithEvents u_fParent                    As Form
Attribute u_fParent.VB_VarHelpID = -1
Private WithEvents u_pParent                    As PictureBox
Attribute u_pParent.VB_VarHelpID = -1

Private Declare Function GetPixel Lib "gdi32" _
    (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long

Public Event ColorSelected(ByVal Color As Long)

Private Sub u_fParent_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lX          As Long
    Dim lY          As Long
    
    With UserControl.Parent
        If (X >= .Controls(Ambient.DisplayName).Left) And (X <= (.Controls(Ambient.DisplayName).Left + .Controls(Ambient.DisplayName).Width)) And _
           (Y >= .Controls(Ambient.DisplayName).Top) And (Y <= (.Controls(Ambient.DisplayName).Top + .Controls(Ambient.DisplayName).Height)) Then
            lX = ScaleX(X, .ScaleMode, vbPixels)
            lY = ScaleY(Y, .ScaleMode, vbPixels)
    
            RaiseEvent ColorSelected(GetPixel(.hdc, lX, lY))
        End If
    End With
End Sub

Private Sub u_iImageInParent_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lX          As Long
    Dim lY          As Long

    lX = ScaleX(X + u_iImageInParent.Left, UserControl.Parent.ScaleMode, vbPixels)
    lY = ScaleY(Y + u_iImageInParent.Top, UserControl.Parent.ScaleMode, vbPixels)

    RaiseEvent ColorSelected(GetPixel(UserControl.Parent.hdc, lX, lY))
End Sub

Private Sub u_pImageInParent_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lX          As Long
    Dim lY          As Long

    With u_pImageInParent
        lX = ScaleX(X, .ScaleMode, vbPixels)
        lY = ScaleY(Y, .ScaleMode, vbPixels)

        RaiseEvent ColorSelected(GetPixel(.hdc, lX, lY))
    End With
End Sub

Private Sub UserControl_Initialize()
    u_eTarget = eTargetPictureBox 'eTargetImageBox  'eTargetParentDirect
End Sub

Private Sub UserControl_Resize()
    'Debug.Print "Resize! "; Now
    On Error Resume Next
    
    With shpDummy
        .Width = UserControl.Width
        .Height = UserControl.Height
    End With
    
    If (UserControl.Parent Is Nothing) Then
        Exit Sub
    End If

    Select Case u_eTarget
        Case eTargetPictureBox
            If (u_pImageInParent Is Nothing) Then
                Exit Sub
            End If
        
            With u_pImageInParent
                .Left = Parent.Controls(Ambient.DisplayName).Left
                .Top = Parent.Controls(Ambient.DisplayName).Top
                .Width = Parent.Controls(Ambient.DisplayName).Width
                .Height = Parent.Controls(Ambient.DisplayName).Height
                u_pImageInParent.Cls
                u_cRenderer.PaintPictureEx .hdc, _
                                           u_cRenderer.LoadPictureEx("colorwheel.png"), _
                                           0, _
                                           0, _
                                           ScaleX(.Width, .ScaleMode, vbPixels), _
                                           ScaleY(.Height, .ScaleMode, vbPixels)
            End With

        Case eTargetImageBox
            If (u_iImageInParent Is Nothing) Then
                Exit Sub
            End If
        
            With u_iImageInParent
                .Left = Parent.Controls(Ambient.DisplayName).Left
                .Top = Parent.Controls(Ambient.DisplayName).Top
                .Width = Parent.Controls(Ambient.DisplayName).Width
                .Height = Parent.Controls(Ambient.DisplayName).Height
            End With
            
        Case eTargetParentDirect
            With UserControl.Parent
                UserControl.Parent.Cls
                u_cRenderer.PaintPictureEx .hdc, _
                                           u_cRenderer.LoadPictureEx("colorwheel.png"), _
                                           ScaleX(.Controls(Ambient.DisplayName).Left, .ScaleMode, vbPixels), _
                                           ScaleY(.Controls(Ambient.DisplayName).Top, .ScaleMode, vbPixels), _
                                           ScaleX(.Controls(Ambient.DisplayName).Width, .ScaleMode, vbPixels), _
                                           ScaleY(.Controls(Ambient.DisplayName).Height, .ScaleMode, vbPixels)
            End With
    End Select
End Sub

Private Sub UserControl_Show()
    'Debug.Print "Show! "; Now
    On Error Resume Next
    If (UserControl.Parent Is Nothing) Then
        Exit Sub
    End If

    If Not (UserControl.Ambient.UserMode = 0) Then
        UserControl.BackStyle = 0
        shpDummy.Visible = False
    Else
        UserControl.BackStyle = 1
        shpDummy.Visible = True
    End If

    Select Case u_eTarget
        Case eTargetPictureBox
            If (u_pImageInParent Is Nothing) Then
                Set u_pImageInParent = UserControl.Parent.Controls.Add("VB.PictureBox", Ambient.DisplayName & "Img")
                With u_pImageInParent
                    .BorderStyle = 0
                    .BackColor = UserControl.Parent.BackColor
                    .AutoRedraw = True
                    .Visible = True
                    .ZOrder 0
                End With
                
                Call UserControl_Resize
            End If
        Case eTargetImageBox
            If (u_iImageInParent Is Nothing) Then
                Set u_iImageInParent = UserControl.Parent.Controls.Add("VB.Image", Ambient.DisplayName & "Img")
                With u_iImageInParent
                    .Visible = True
                    .Stretch = True
                    Set .Picture = u_cRenderer.LoadPictureEx("colorwheel.png")
                    .ZOrder 0
                End With
                
                Call UserControl_Resize
            End If
        Case eTargetParentDirect
            If TypeOf UserControl.Parent Is Form Then
                Set u_fParent = UserControl.Parent
            ElseIf TypeOf UserControl.Parent Is PictureBox Then
                Set u_pParent = UserControl.Parent
            End If
            
            UserControl.Parent.AutoRedraw = True
            Call UserControl_Resize
    End Select
End Sub

Private Sub UserControl_Terminate()
    'Debug.Print "Terminate! "; Now
    Select Case u_eTarget
        Case eTargetPictureBox
            If Not (u_pImageInParent Is Nothing) Then
                Set u_pImageInParent = Nothing
            End If
        Case eTargetImageBox
            If Not (u_iImageInParent Is Nothing) Then
                Set u_iImageInParent = Nothing
            End If
    End Select
End Sub
