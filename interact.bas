Attribute VB_Name = "vbInteraction"
Option Explicit
Public InputResult As String, ConfirmResult As String

Public Function GetInput(Prompt As String, Optional Title As String, Optional Default As String, Optional X As Single, Optional Y As Single, Optional MaskInput As Boolean = False) As String
InputResult = ""
Load frmPrompt
If MaskInput = True Then
frmPrompt.Check1.Visible = True
frmPrompt.Check1.Value = 1
frmPrompt.Text1.PasswordChar = "*"
Else
frmPrompt.Text1.PasswordChar = ""
End If
frmPrompt.Label1.Caption = Prompt
If Title <> "" Then
frmPrompt.Label2.Caption = Title
End If
frmPrompt.Text1.Text = Default
frmPrompt.Show vbModal
If IsNumeric(X) = True Then frmPrompt.Left = X
If IsNumeric(Y) = True Then frmPrompt.Top = Y
GetInput = InputResult
End Function

Sub Alert(Message As String, Optional WindowTitle As String = "Message")
Load frmMsg
frmMsg.Label2.Caption = WindowTitle
frmMsg.Label1.Caption = Message
frmMsg.Show vbModal
End Sub
Function Confirm(Message As String, Optional Title As String = "Confirm", Optional Cancel As Boolean = True) As String
Load frmConf
frmConf.lTitle.Caption = Title
frmConf.Label1.Caption = Message
frmConf.Command3.Visible = Cancel
frmConf.Show vbModal
Confirm = ConfirmResult
End Function
