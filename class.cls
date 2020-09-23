VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "General"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'----------------------------------------------
'Visual Basic Runtime Procedures Extension
'Sushant Pandurangi <sushant@phreaker.net>
'-----------------------------------------------
Option Explicit
'-----------------------------------------------
Public Enum EnumErrorCodes
 TYPE_MISMATCH = 13
 ERR_INVALID_CALL = 5
 FILE_DOES_EXIST = 58
 FILE_NOT_FOUND = 53
 UNKNOWN_ERROR = 1
 SUCCESS_DONE = 0
 STACK_OVERFLOW = 6
 OUT_OF_MEMORY = 7
End Enum
'-----------------------------------------------------------

Sub AboutBox()
frmAbout.Show vbModal
End Sub

Function GetError(Optional ErrNumber As EnumErrorCodes) As String
Attribute GetError.VB_Description = "Get the error message corresponding to the given error number or type."
GetError = Error$(ErrNumber)
End Function

Function Decrypt(sText As String, sKey As String) As String
Attribute Decrypt.VB_Description = "Decrypts the given string."
    Dim CipherTest As New Cipher
    CipherTest.KeyString = sKey
    CipherTest.Text = sText
    CipherTest.Shrink
    CipherTest.DoXor
    Decrypt = CipherTest.Text
End Function

Function Encrypt(sText As String, sKey As String) As String
Attribute Encrypt.VB_Description = "Encrypts the given string."
Dim CipherTest As New Cipher
    CipherTest.KeyString = sKey
    CipherTest.Text = sText
    CipherTest.DoXor
    CipherTest.Stretch
    Encrypt = CipherTest.Text
End Function

Function FileOpen(sFileName As String) As String
Attribute FileOpen.VB_Description = "Opens a file."
'Its always better to open it as binary
Open sFileName For Binary As #1
FileOpen = Input(LOF(1), 1)
Close #1
End Function

Sub FileSave(sFileName As String, sContents As String)
Attribute FileSave.VB_Description = "Saves a file. (Text)"
Open sFileName For Binary As #1
Print #1, sContents
Close #1
End Sub

Sub Shortcut(sPath As String, sName As String, sTarget As String, Optional sArguments As String)
Attribute Shortcut.VB_Description = "Create a shortcut."
Dim lReturN
lReturN = fCreateShellLink(sPath, sName, sTarget, sArguments)
End Sub

Sub ListFonts(lComboBox As Object)
Attribute ListFonts.VB_Description = "Lists users' fonts."
'List all user's fonts
Dim i As Integer
For i = 0 To Screen.FontCount - 1
lComboBox.AddItem Screen.Fonts(i)
Next i
End Sub

Function ValidFont(vFontName As String) As Boolean
'Determine if a font exists
Dim i As Integer
For i = 0 To Screen.FontCount - 1
If vFontName = Screen.Fonts(i) Then
ValidFont = True
Exit Function
End If
Next i
End Function

Function FixedFont(vFontName As String) As Boolean
'Determine if a font is fixed-width
Dim i As Integer, imE As frmAbout
Load imE
imE.Hide
imE.Font.Name = vFontName
If imE.TextWidth(".") = imE.TextWidth("W") Then
FixedFont = True
Exit Function
End If
Me.Font.Name = "Tahoma"
FixedFont = False
Unload imE
End Function

Function BMPFont(vFontName As String) As Boolean
'Determine if a font is not a TTF
Load frmAbout
frmAbout.Hide
frmAbout.Font.Name = vFontName
frmAbout.Font.Size = 72.75
If frmAbout.Font.Size <> 72.75 Then BMPFont = False Else BMPFont = True
Unload frmAbout
End Function
