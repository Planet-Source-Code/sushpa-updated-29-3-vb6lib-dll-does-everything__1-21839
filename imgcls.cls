VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Graphics"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'----------------------------------------------
'Visual Basic Runtime Procedures Extension
'Sushant Pandurangi <sushant@phreaker.net>
'-----------------------------------------------
Option Explicit
Public FileName As String, FileSize As Integer, Width As Integer, Height As Integer
Attribute FileName.VB_VarMemberFlags = "400"
Attribute FileName.VB_VarDescription = "Returns the filename of the loaded image."
Attribute FileSize.VB_VarDescription = "Returns FileSize of the image."
Attribute Width.VB_VarDescription = "Returns the width of the image."
Attribute Height.VB_VarDescription = "Returns the height of the image."

Sub GradeForm(pObject As Object, Optional Colour As Integer = 2, Optional Orientation As Integer = 0, Optional Range As Integer = 700)
Attribute GradeForm.VB_Description = "Create a gradient with the specified colour on the specified object."
pObject.AutoRedraw = True
Dim intY As Integer, sColour
pObject.Scale (0, 0)-(Range, Range)
For intY = 0 To Range
Select Case Colour
Case 0
sColour = RGB(CInt((intY / Range) * 255), 0, 0)
Case 1
sColour = RGB(0, CInt((intY / Range) * 255), 0)
Case 2
sColour = RGB(0, 0, CInt((intY / Range) * 255))
Case 3
sColour = RGB(0, 128, CInt((intY / Range) * 255))
Case 4
sColour = RGB(128, 0, CInt((intY / Range) * 255))
Case 5
sColour = RGB(CInt((intY / Range) * 255), 0, 128)
Case 6
sColour = RGB(CInt((intY / Range) * 255), 128, 0)
Case 7
sColour = RGB(128, CInt((intY / Range) * 255), 0)
Case 8
sColour = RGB(0, CInt((intY / Range) * 255), 128)
Case 9
sColour = RGB(0, CInt((intY / Range) * 255), CInt((intY / Range) * 255))
Case 10
sColour = RGB(CInt((intY / Range) * 255), CInt((intY / Range) * 255), 0)
Case 11
sColour = RGB(CInt((intY / Range) * 255), 0, CInt((intY / Range) * 255))
Case 12
sColour = RGB(255 - CInt((intY / Range) * 255), 0, 0)
Case 13
sColour = RGB(0, 255 - CInt((intY / Range) * 255), 0)
Case 14
sColour = RGB(0, 0, 255 - CInt((intY / Range) * 255))
End Select
If Orientation = 0 Then
pObject.Line (0, intY)-(Range, intY), sColour
Else
pObject.Line (intY, 0)-(intY, Range), sColour
End If
Next intY
End Sub

Sub AboutBox()
frmAbout.Show vbModal
End Sub

Function LoadImage(Optional szFileName As String)
Attribute LoadImage.VB_Description = "Loads an Image File."
FileName = szFileName
Dim colS, Size, wX, wY, LoAdEd
LoAdEd = LoadPicture(szFileName, Size, colS, wX, wY)
FileSize = Size
Width = wX
Height = wY
LoadImage = LoAdEd
End Function

Function SaveImage(szFileName As String, szPicture)
Attribute SaveImage.VB_Description = "Saves an Image File."
SavePicture szPicture, szFileName
End Function

Public Sub Draw3DLine(pForm As Object, ByVal intX1 As Single, ByVal intY1 As Single, ByVal intLength As Long)
Attribute Draw3DLine.VB_Description = "Draws a 3-Dimensional Line on the specified object (Form)."
    pForm.Line (intX1, intY1)-(intX1 + intLength, intY1), lGRAY
    pForm.Line (pForm.CurrentX + 5, intY1 + 20)-(intX1 - 5, intY1 + 20), lWHITE
End Sub

Sub Draw3DText(pObject As Object, sText As String, TextColor As Long)
Attribute Draw3DText.VB_Description = "Draws a 3-Dimensional Appearing replica of the given string."
Dim i As Integer, X As Single, Y As Single
pObject.ForeColor = 0: X = pObject.CurrentX: Y = pObject.CurrentY
For i = 1 To 100
        pObject.Print sText ' Text Here
        X = X + 1: Y = Y + 1: pObject.CurrentX = X: pObject.CurrentY = Y
Next
pObject.ForeColor = TextColor 'Change Color Here
pObject.Print sText ' Text Goes Here
End Sub

Sub Progressbar(PictureBox As Object, ByVal sngPercent As Integer, Optional ByVal fBorderCase As Boolean = False)
If sngPercent > 100 Then Err.Raise 345356, , "Percent Value cannot be over 100."
On Error Resume Next
    Dim strPercent As String
    Dim intX As Integer
    Dim intY As Integer
    Dim intWidth As Integer
    Dim intHeight As Integer
    'For this to work well, we need a white background and any color foreground (blue)
    Const colBackground = &HFFFFFF ' white
    Const colForeground = 0 '&H800000 ' dark blue
    PictureBox.ForeColor = colForeground
    PictureBox.BackColor = colBackground
    'Format percentage and get attributes of text
    Dim intPercent
    intPercent = sngPercent ' Int(100 * sngPercent + 0.5)
    'Never allow the percentage to be 0 or 100 unless it is exactly that value.  This
    'prevents, for instance, the status bar from reaching 100% until we are entirely done.
    If intPercent = 0 Then
        If Not fBorderCase Then
            intPercent = 1
        End If
    ElseIf intPercent = 100 Then
        If Not fBorderCase Then
            intPercent = 99
        End If
    End If
    strPercent = Format$(intPercent) & "%"
    intWidth = PictureBox.TextWidth(strPercent)
    intHeight = PictureBox.TextHeight(strPercent)
    'Now set intX and intY to the starting location for printing the percentage
    intX = PictureBox.Width / 2 - intWidth / 2
    intY = (PictureBox.Height / 2 - intHeight / 2) - 20
    'Need to draw a filled box with the pics background color to wipe out previous
    'percentage display (if any)
    PictureBox.DrawMode = 13 ' Copy Pen
    PictureBox.Line (intX, intY)-Step(intWidth, intHeight), PictureBox.BackColor, BF
    'Back to the center print position and print the text
    PictureBox.CurrentX = intX
    PictureBox.CurrentY = intY
    PictureBox.Print strPercent
    'Now fill in the box with the ribbon color to the desired percentage
    'If percentage is 0, fill the whole box with the background color to clear it
    'Use the "Not XOR" pen so that we change the color of the text to white
    'wherever we touch it, and change the color of the background to blue
    'wherever we touch it.
    PictureBox.DrawMode = 10 ' Not XOR Pen
    If sngPercent > 0 Then
        PictureBox.Line (0, 0)-(PictureBox.Width * sngPercent, PictureBox.Height), PictureBox.ForeColor, BF
    Else
        PictureBox.Line (0, 0)-(PictureBox.Width, PictureBox.Height), PictureBox.BackColor, BF
    End If
    PictureBox.Refresh
End Sub

Function CursorPos(ReturnedX As Long, ReturnedY As Long, Optional ScaleMode = vbPixels) As String
    Dim DivisorX, DivisorY As Integer
GetCursorPos LPOINT
    If ScaleMode = vbPixels Then DivisorX = 1 Else DivisorX = Screen.TwipsPerPixelX
    If ScaleMode = vbPixels Then DivisorY = 1 Else DivisorY = Screen.TwipsPerPixelY
ReturnedX = LPOINT.X * DivisorX
ReturnedY = LPOINT.Y * DivisorY
    CursorPos = ReturnedX & ", " & ReturnedY
End Function


