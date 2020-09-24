VERSION 5.00
Begin VB.Form login 
   BackColor       =   &H00A21000&
   BorderStyle     =   0  'None
   Caption         =   "Login"
   ClientHeight    =   7815
   ClientLeft      =   4545
   ClientTop       =   3660
   ClientWidth     =   10440
   LinkTopic       =   "Form7"
   ScaleHeight     =   7815
   ScaleWidth      =   10440
   ShowInTaskbar   =   0   'False
   Begin VB.Timer tmrval 
      Enabled         =   0   'False
      Interval        =   230
      Left            =   2760
      Top             =   4800
   End
   Begin Project1.LabelText Text2 
      Height          =   255
      Left            =   2280
      TabIndex        =   3
      Top             =   2520
      Visible         =   0   'False
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   450
      Caption         =   "Code"
      Text            =   "pass"
   End
   Begin Project1.LabelText Text1 
      Height          =   255
      Left            =   2280
      TabIndex        =   2
      Top             =   1920
      Visible         =   0   'False
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   450
      Caption         =   "User-ID"
      Text            =   "user"
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   720
      Picture         =   "login.frx":0000
      ScaleHeight     =   465
      ScaleWidth      =   825
      TabIndex        =   1
      Top             =   120
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Timer tmrphr 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   4320
      Top             =   840
   End
   Begin VB.Timer tmrmove 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   4320
      Top             =   840
   End
   Begin Project1.PictureButton Accessor 
      Height          =   450
      Left            =   7200
      TabIndex        =   0
      ToolTipText     =   "Access to Dextop"
      Top             =   4080
      Visible         =   0   'False
      Width           =   1605
      _ExtentX        =   2831
      _ExtentY        =   794
      Picture         =   "login.frx":3A2B9
      PictureHover    =   "login.frx":3C905
      PictureDown     =   "login.frx":3EF51
   End
   Begin Project1.aicAlphaImage Logo 
      Height          =   4875
      Left            =   600
      Top             =   0
      Visible         =   0   'False
      Width           =   9300
      _ExtentX        =   18521
      _ExtentY        =   18521
      Image           =   "login.frx":4159D
      Scaler          =   4
      Props           =   5
      ScaleCx         =   620
      ScaleCy         =   325
   End
   Begin Project1.aicAlphaImage Phrase 
      Height          =   3285
      Index           =   4
      Left            =   7080
      Top             =   3480
      Visible         =   0   'False
      Width           =   3285
      _ExtentX        =   5794
      _ExtentY        =   5794
      Image           =   "login.frx":56E64
      Props           =   5
   End
   Begin Project1.aicAlphaImage Phrase 
      Height          =   2340
      Index           =   3
      Left            =   5400
      Top             =   5880
      Visible         =   0   'False
      Width           =   2340
      _ExtentX        =   4128
      _ExtentY        =   4128
      Image           =   "login.frx":5A360
      Props           =   5
   End
   Begin Project1.aicAlphaImage Phrase 
      Height          =   1890
      Index           =   2
      Left            =   3720
      Top             =   3840
      Visible         =   0   'False
      Width           =   1890
      _ExtentX        =   3334
      _ExtentY        =   3334
      Image           =   "login.frx":5D7A1
      Props           =   5
   End
   Begin Project1.aicAlphaImage Phrase 
      Height          =   2235
      Index           =   1
      Left            =   2160
      Top             =   5880
      Visible         =   0   'False
      Width           =   2235
      _ExtentX        =   3942
      _ExtentY        =   3942
      Image           =   "login.frx":602A9
      Props           =   5
   End
   Begin Project1.aicAlphaImage Phrase 
      Height          =   2415
      Index           =   0
      Left            =   120
      Top             =   4440
      Visible         =   0   'False
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   4260
      Image           =   "login.frx":62AC6
      Props           =   5
   End
   Begin VB.Image back 
      Height          =   11520
      Left            =   0
      Picture         =   "login.frx":67746
      Top             =   0
      Width           =   15360
   End
End
Attribute VB_Name = "login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim xdone  As Boolean
Dim r As Integer
Dim coorX(4) As Long
Dim coorY(4) As Long
Dim coor2X(4) As Long
Dim coor2Y(4) As Long
Private Sub Accessor_Click()
On Error Resume Next
DuCr.I2S App.path & "\User.bmp", App.path & "\User.ini"
If Text1.text = GetFromIni("Main", "UserName", App.path & "\User.ini") Then
If text2.text = GetFromIni("Main", "Password", App.path & "\User.ini") Then
Dim shell As New shell
shell.MinimizeAll
Form1.Show
Set Me.back = Nothing
Set Me.Picture1 = Nothing
list1.Clear
Unload Me
For X = 0 To Phrase.UBound
Phrase(X).ClearImage
Set Phrase(X) = Nothing
Next
Else
GoTo X
End If
Kill App.path & "User.ini"
Else
X:
Dim tpc As StdPicture
Set tpc = back
Set back = Picture1
BackColor = &HA8&
MsgBox "You have Entered Wrong Password", vbCritical, "Access Denied"
BackColor = &HD76539
Set back = tpc
End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
Accessor_Click
End If

End Sub

Private Sub Form_Load()
On Error Resume Next
Dim ang As Long, avg As Long
If App.PrevInstance = True Then GoTo X
text2.pword "!"
Me.Left = 0
Top = 0
Width = Screen.Width
Height = Screen.Height
back.Left = Screen.Width / 2 - (back.Width / 2)
back.Top = 0
xdone = True
tmrmove = True
For X = 1 To 5
avg = (Screen.Width + Screen.Height) / 3
ang = 270 + 22 * (4 - X)
coorX(X) = Cos(3.1416 * ang / 180) * (avg - 1000)
coorY(X) = -(Sin(3.1416 * ang / 180) * (avg - 1000))
ang = 180 + 22 * (4 - X)
coor2X(X) = Cos(3.1416 * ang / 180) * (avg - 1000)
coor2Y(X) = -(Sin(3.1416 * ang / 180) * (avg - 1000))
'MsgBox coorX(x) & "\\" & coorY(x)
Next
r = 0
Exit Sub
X:
MsgBox "Your Workstation is not having the capabilities to create multiple environments", vbDefaultButton1 + vbSystemModal, "Fatal Error"
End
End Sub

Sub DragLogin()
On Error Resume Next
On Error Resume Next
Logo.Left = (Width / 2) - (Logo.Width / 2)
With Logo
.Top = -.Height
.Visible = True
Dim X As Integer
Dim i As Integer
i = -.Height
For X = -.Height To 0 Step 20
.Top = X
.Opacity = i / 32
i = i - 1
Next
.FadeInOut 100
End With
tmrphr_Timer
merge
xdone = False
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
Accessor_Click
End If
End Sub
Private Sub Text2_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
Accessor_Click
End If
End Sub

Private Sub tmrmove_Timer()
On Error Resume Next
If xdone = False Then
tmrmove.Interval = 1
Dim i  As Integer
 If Text1.Left > Logo.Left + 2200 Then GoTo G
text2.Left = text2.Left + 120
Text1.Left = Text1.Left + 120
Else
DragLogin
End If
Exit Sub
G:
xdone = True
tmrmove = False
moveacc
End Sub
Sub merge()
On Error Resume Next
Text1.Left = -Text1.Width
text2.Left = -text2.Width
Text1.Visible = True
text2.Visible = True
End Sub
Sub moveacc()
On Error Resume Next
Accessor.ZOrder 0
Accessor.Top = 3480
Accessor.Visible = True
Dim X As Integer
Dim i As Integer
For X = 3480 To 4320
Accessor.Top = X
Accessor.Left = Text1.Left + Text1.Width + 300
Next
X = 4320
For i = 3480 To 4320
Accessor.Top = X
X = X - 1
Next
tmrphr = True
End Sub

Private Sub tmrphr_Timer()
Dim X As Integer, lft As Integer, tp As Integer, vis As Integer, rd As Integer
On Error Resume Next
Randomize 50
vis = Rnd(50) * 25
For X = 0 To Phrase.UBound
rd = Sgn(Rnd - Rnd)
If rd = 1 Then
Phrase(X).Left = coorX(X + 1)
Else
Phrase(X).Left = coor2X(X + 1)
End If
rd = Sgn(Rnd - Rnd)
If rd = 1 Then
Phrase(X).Top = coorY(X + 1)
Else
Phrase(X).Top = coor2Y(X + 1)
End If
If rd = 1 Then
Phrase(X).Opacity = 100
Phrase(X).Visible = True
Else
Phrase(X).Visible = False
End If
Next
tmrval = True
End Sub

Private Sub tmrval_Timer()
On Error Resume Next
For X = 0 To 4
Phrase(X).FadeInOut 0, 20, 50
Next
tmrval = False
End Sub
