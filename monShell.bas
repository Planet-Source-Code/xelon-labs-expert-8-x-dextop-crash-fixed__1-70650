Attribute VB_Name = "modShell"
Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public Sub ShellFile(path As String)
On Error Resume Next
Dim lng As Long
lng = GethWndByWinTitle("Start Menu")
Call ShellExecute(lng, "open", path, "", "", 1)
End Sub

Sub exp(str As String)
On Error Resume Next
Dim lng As Long
lng = GethWndByWinTitle("Start Menu")
ShellExecute lng, "open", str, "", "", 1
End Sub

Sub Write_INIList(ini As String, lst As ListBox)
On Error Resume Next
WriteIni "Main", lst.name & " Count", lst.ListCount - 1, ini
Dim X As Integer
For X = 0 To lst.ListCount - 1
WriteIni "Main", lst.name & X, lst.List(X), ini
Next
End Sub

Sub Get_INIList(ini As String, lst As ListBox)
On Error Resume Next
Dim cnt As String
cnt = GetFromIni("Main", lst.name & " Count", ini)
Dim X As Integer
For X = 0 To cnt - 1
lst.Additem GetFromIni("Main", lst.name & X, ini)
Next
End Sub

Function InputFrm(Prompt As String, Label As String, text As String, pword As String) As String
On Error Resume Next
Frminput.Label1 = Prompt
Frminput.LabelText1.text = text
Frminput.LabelText1.Caption = Label
Frminput.LabelText1.pword pword
Frminput.Show vbModal
InputFrm = Frminput.LabelText1.text
End Function
