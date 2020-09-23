VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Cipher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private mStrKey As String
Private mStrText As String
Public EncText As String

'~~~.KeyString
'A string (key) used in encryption and decryption
Public Property Let KeyString(strKey As String)
    mStrKey = strKey
    Initialize
End Property

'Write text to be encrypted or decrypted
Public Property Let Text(strText As String)
    mStrText = strText
End Property

'Read text that was encrypted or decrypted
Public Property Get Text() As String
    Text = mStrText
End Property

'Exclusive-or method to encrypt or decrypt
Public Sub DoXor()
On Error Resume Next
    Dim lngC As Long
    Dim intB As Long
    Dim lngN As Long
    For lngN = 1 To Len(mStrText)
        lngC = Asc(Mid(mStrText, lngN, 1))
        intB = Int(Rnd * 256)
        Mid(mStrText, lngN, 1) = Chr(lngC Xor intB)
    Next lngN

End Sub

'Convert any string to a printable, displayable string
Public Sub Stretch()
On Error Resume Next
    Dim lngC As Long
    Dim lngN As Long
    Dim lngJ As Long
    Dim lngK As Long
    Dim lngA As Long
    Dim strB As String
    lngA = Len(mStrText)
    strB = Space(lngA + (lngA + 2) \ 3)
    For lngN = 1 To lngA
        lngC = Asc(Mid(mStrText, lngN, 1))
        lngJ = lngJ + 1
        Mid(strB, lngJ, 1) = Chr((lngC And 63) + 59)
        Select Case lngN Mod 3
        Case 1
            lngK = lngK Or ((lngC \ 64) * 16)
        Case 2
            lngK = lngK Or ((lngC \ 64) * 4)
        Case 0
            lngK = lngK Or (lngC \ 64)
            lngJ = lngJ + 1
            Mid(strB, lngJ, 1) = Chr(lngK + 59)
            lngK = 0
        End Select
    Next lngN
    If lngA Mod 3 Then
        lngJ = lngJ + 1
        Mid(strB, lngJ, 1) = Chr(lngK + 59)
    End If
    mStrText = strB
End Sub

'~~~.Shrink
'Inverse of the Stretch method;
'result can contain any of the 256-byte values
Public Sub Shrink()
On Error Resume Next
    Dim lngC As Long
    Dim lngD As Long
    Dim lngE As Long
    Dim lngA As Long
    Dim lngB As Long
    Dim lngN As Long
    Dim lngJ As Long
    Dim lngK As Long
    Dim strB As String
    lngA = Len(mStrText)
    lngB = lngA - 1 - (lngA - 1) \ 4
    strB = Space(lngB)
    For lngN = 1 To lngB
        lngJ = lngJ + 1
        lngC = Asc(Mid(mStrText, lngJ, 1)) - 59
        Select Case lngN Mod 3
        Case 1
            lngK = lngK + 4
            If lngK > lngA Then lngK = lngA
            lngE = Asc(Mid(mStrText, lngK, 1)) - 59
            lngD = ((lngE \ 16) And 3) * 64
        Case 2
            lngD = ((lngE \ 4) And 3) * 64
        Case 0
            lngD = (lngE And 3) * 64
            lngJ = lngJ + 1
        End Select
        Mid(strB, lngN, 1) = Chr(lngC Or lngD)
    Next lngN
    mStrText = strB
End Sub

'Initializes random numbers using the key string
Private Sub Initialize()
    Dim lngN As Long
    Randomize Rnd(-1)
    For lngN = 1 To Len(mStrKey)
        Randomize Rnd(-Rnd * Asc(Mid(mStrKey, lngN, 1)))
    Next lngN
End Sub
