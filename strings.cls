VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Strings"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'----------------------------------------------
'Visual Basic Runtime Procedures Extension
'Sushant Pandurangi <sushant@phreaker.net>
'-----------------------------------------------
Option Explicit
Dim pStep As Integer, i As Integer, pCount As Integer

Public Function StrCount(sString As String, sChar As String) As Integer
Attribute StrCount.VB_Description = "Count how many times the specified character appears in the string."
'Somehow I havent seen such a function in the object
'browser till today. VB needed this one badly.
    pStep = InStr(1, sString, Left(sChar, 1))
'pStep is the first occurence
    If pStep = 0 Then Exit Function 'Char dosent exist
looper:
    Do
        pStep = InStr(pStep + 1, sString, Left(sChar, 1))
        pCount = pCount + 1
    Loop Until pStep = 0
        StrCount = pCount: pCount = 0: pStep = 0
End Function

Public Function Repeat(sCharacter As String, Length As Long) As String
Attribute Repeat.VB_Description = "Repeats a string character."
'Maybe the String() function I see in VB6 is present in
'VB5 as well. Maybe not, so here is a confirmation.
Dim i As Integer, TEMP As String
For i = 1 To Length
TEMP = TEMP & sCharacter
Next i
Repeat = TEMP
TEMP = "": i = 0
End Function

Public Function Reverse(sString As String) As String
Attribute Reverse.VB_Description = "Reverse a given string."
'VB6 has this as an in-built function called
'StrReverse(String) but I am not sure of VB5.
Dim i As Integer, s As String
For i = 1 To Len(sString)
s = s & Mid(sString, Len(sString) + 1 - i, 1)
Next i
Reverse = s
End Function

Public Function StrReplace(pTextBox As Object, pFind As String, pReplace As String) As String
Attribute StrReplace.VB_Description = "Replaces a string within a textbox."
'Replace the text to be found. What this does is
'sort of restricted; in that only the text in a textbox
'can be modified.
pTextBox.SelStart = 1
Do
pTextBox.SelStart = InStr(pTextBox.SelStart, pTextBox.Text, pFind) - 1
pTextBox.SelLength = Len(pFind)
pTextBox.SelText = pReplace
If InStr(1, pTextBox.Text, pFind) = 0 Then Exit Do: Exit Function
Loop Until InStr(1, pTextBox.Text, pFind) = 0
End Function

Public Function ParseInt(Expression As Variant) As Integer
Attribute ParseInt.VB_Description = "Return the integer part of an expression."
ParseInt = Int(Expression)
End Function

Public Function CBinary(Expression As Boolean) As Integer
Attribute CBinary.VB_Description = "Convert the given boolean to 0 or 1."
'Useful for converting BOOLs to 0 or 1. CByte() would
'return 255 for true, which wont be useful for setting the
'values of, for instance, a checkbox; as it uses 0 and 1.
If Expression = False Then CBinary = 0 Else CBinary = 1
End Function

Sub AboutBox()
Attribute AboutBox.VB_UserMemId = -552
frmAbout.Show vbModal
End Sub

Function WdCount(pString As String) As Long
Attribute WdCount.VB_Description = "Count the number of words in the string."
'Number of words; decided using number of spaces and other characters
WdCount = StrCount(pString, " ") + 1 + StrCount(pString, "=") + StrCount(pString, "-") + StrCount(pString, "+") + StrCount(pString, "\") + StrCount(pString, "/") + StrCount(pString, ".")
End Function

Function LnCount(pTextBox As Object) As Integer
Attribute LnCount.VB_Description = "Get the number of lines in a Textbox."
'Number of lines
LnCount = SendMessage(pTextBox.hWnd, &HBA, 0, 0&)
End Function

Function SnCount(pText As String) As Integer
Attribute SnCount.VB_Description = "Get number of sentences."
'Number of sentences
SnCount = StrCount(pText, ".")
End Function

Function Up1Level(sPath As String) As String
Attribute Up1Level.VB_Description = "Return the folder that is up one level from the given one."
Dim pos As Long, i As Integer, Dummy As String
If Right(sPath, 1) = "\" Then sPath = Left(sPath, Len(sPath) - 1)
Dummy = Reverse(sPath)
pos = InStr(1, Dummy, "\")
Up1Level = Right$(Dummy, Len(Dummy) - pos)
Up1Level = Reverse(Up1Level)
If Right(Up1Level, 1) = ":" Then Up1Level = Up1Level & "\"
End Function

Function GetFile(sPath As String) As String
Attribute GetFile.VB_Description = "Get the filename portion from the string"
Dim i, j As Integer
i = InStr(1, Reverse(sPath), "\")
GetFile = Right(sPath, i - 1)
End Function

Function GetPath(sPath As String) As String
Attribute GetPath.VB_Description = "Get the pathname portion from the string"
GetPath = Up1Level(sPath) & "\"
End Function

Function InitCap(sString As String) As String
Attribute InitCap.VB_Description = "Returns string with initial capitals."
    InitCap = UCase(Left(sString, 1)) & LCase(Right(sString, Len(sString) - 1))
End Function

Public Function ReadValue(Section As String, Key As String, FileName As String, Optional Default As String)
Attribute ReadValue.VB_Description = "Reads the value of a variable from an INI File."
    Dim sReturn As String
    sReturn = String(255, Chr(0))
    ReadValue = Left(sReturn, GetPrivateProfileString(Section, Key, Default, sReturn, Len(sReturn), FileName))
End Function

Public Sub SaveValue(Section As String, Key As String, Value As String, FileName As String)
Attribute SaveValue.VB_Description = "Saves a value with a variable in an INI file."
    ' Write to INI file
    WritePrivateProfileString Section, Key, Value, FileName
End Sub
