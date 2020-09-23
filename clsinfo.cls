VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Prompts"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'----------------------------------------------
'Visual Basic Runtime Procedures Extension
'Sushant Pandurangi <sushant@phreaker.net>
'-----------------------------------------------
Option Explicit

Sub AboutBox()
frmAbout.Show vbModal
End Sub

Function Message(sMessage As String, Optional sTitle As String)
Attribute Message.VB_Description = "Displays a message box."
    Alert sMessage, sTitle
End Function

Function Confirm(sMessage As String, Optional sTitle As String, Optional sCancel As Boolean = True) As String
Attribute Confirm.VB_Description = "Confirms a yes or no from the user. Visibility of cancel button can be specified."
    Confirm = vbInteraction.Confirm(sMessage, sTitle, sCancel)
End Function

Function Prompt(sMessage As String, Optional sDefault As String, Optional sTitle As String, Optional sMaskInput As Boolean = False) As String
Attribute Prompt.VB_Description = "Gets a value for a certain prompt, from the user."
    Prompt = GetInput(sMessage, sTitle, sDefault, , , sMaskInput)
End Function
