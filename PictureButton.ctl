VERSION 5.00
Begin VB.UserControl PictureButton 
   AutoRedraw      =   -1  'True
   ClientHeight    =   2430
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2715
   EditAtDesignTime=   -1  'True
   ScaleHeight     =   2430
   ScaleWidth      =   2715
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   0
      Top             =   0
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   0
      ScaleHeight     =   855
      ScaleWidth      =   1635
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1635
   End
End
Attribute VB_Name = "PictureButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

Private Type POINTAPI
        x As Long
        Y As Long
End Type

Dim m_Picture As StdPicture
Dim m_PictureHover As StdPicture
Dim m_PictureDown As StdPicture

Dim isOver As Boolean
Dim m_State As Integer
Event Click()
Event DblClick()
Event MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Event MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
Event MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
Event MouseEnter()
Event MouseLeave()

Public Property Get Picture() As StdPicture
    Set Picture = m_Picture
End Property

Public Property Set Picture(m_New_Picture As StdPicture)
    Set m_Picture = m_New_Picture
    DrawState
    PropertyChanged "Picture"
End Property

Public Property Get PictureHover() As StdPicture
    Set PictureHover = m_PictureHover
End Property
Sub loadpics(str As String, kw As String)
On Error Resume Next
Set m_Picture = LoadPicture(str & "\" & kw & "1.bmp")
Set m_PictureHover = LoadPicture(str & "\" & kw & "2.bmp")
Set m_PictureDown = LoadPicture(str & "\" & kw & "3.bmp")
Call DrawState
End Sub
Public Property Set PictureHover(m_New_PictureHover As StdPicture)
    Set m_PictureHover = m_New_PictureHover
    DrawState
    PropertyChanged "PictureHover"
End Property

Public Property Get PictureDown() As StdPicture
    Set PictureDown = m_PictureDown
End Property

Public Property Set PictureDown(m_New_PictureDown As StdPicture)
    Set m_PictureDown = m_New_PictureDown
    DrawState
    PropertyChanged "PictureDown"
End Property

Private Sub DrawState()
    On Error Resume Next
    If m_State = 1 Then 'mouse hover
        UserControl.Cls
        UserControl.PaintPicture m_PictureHover, 0, 0
        Picture1.Picture = m_PictureHover
        UserControl.Width = Picture1.Width
        UserControl.Height = Picture1.Height
    ElseIf m_State = 2 Then 'mouse down
        UserControl.Cls
        UserControl.PaintPicture m_PictureDown, 0, 0
        Picture1.Picture = m_PictureDown
        UserControl.Width = Picture1.Width
        UserControl.Height = Picture1.Height
    Else
        UserControl.Cls
        UserControl.PaintPicture m_Picture, 0, 0
        Picture1.Picture = m_Picture
        UserControl.Width = Picture1.Width
        UserControl.Height = Picture1.Height
    End If
End Sub

Private Sub Timer1_Timer()
    If Not CheckMouseOver Then
        Timer1.Enabled = False
        isOver = False
        RaiseEvent MouseLeave
        m_State = 0
        Call DrawState
    End If
End Sub

Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub UserControl_InitProperties()
    Set m_Picture = Nothing
    Set m_PictureHover = Nothing
    Set m_PictureDown = Nothing
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, x, Y)
    If Button = 1 Then
        m_State = 2
        Call DrawState
    End If
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, x, Y)
    If Button < 2 Then
        If Not CheckMouseOver Then
            isOver = False
            m_State = 0
            Call DrawState
        Else
            If Button = 0 And Not isOver Then
                Timer1.Enabled = True
                isOver = True
                RaiseEvent MouseEnter
                m_State = 1
                Call DrawState
            ElseIf Button = 1 Then
                isOver = True
                m_State = 2
                Call DrawState
                isOver = False
            End If
        End If
    End If
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, x, Y)
    If CheckMouseOver Then
        m_State = 1
    Else
        m_State = 0
    End If
    Call DrawState
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    With PropBag
        Set m_Picture = .ReadProperty("Picture", Nothing)
        Set m_PictureHover = .ReadProperty("PictureHover", Nothing)
        Set m_PictureDown = .ReadProperty("PictureDown", Nothing)
    End With
End Sub

Private Sub UserControl_Resize()
    DrawState
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        Call .WriteProperty("Picture", m_Picture, Nothing)
        Call .WriteProperty("PictureHover", m_PictureHover, Nothing)
        Call .WriteProperty("PictureDown", m_PictureDown, Nothing)
    End With
End Sub

Private Function CheckMouseOver() As Boolean
    Dim pt As POINTAPI
    GetCursorPos pt
    CheckMouseOver = (WindowFromPoint(pt.x, pt.Y) = UserControl.hwnd)
End Function
Sub savepix(str As String, kw As String)
SavePicture m_Picture, str & "\" & kw & "1.bmp"
SavePicture m_PictureHover, str & "\" & kw & "2.bmp"
SavePicture m_PictureDown, str & "\" & kw & "3.bmp"
End Sub
