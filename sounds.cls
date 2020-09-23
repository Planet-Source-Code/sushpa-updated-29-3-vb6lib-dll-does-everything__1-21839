VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Sounds"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'----------------------------------------------
'Visual Basic Runtime Procedures Extension
'Sushant Pandurangi <sushant@phreaker.net>
'-----------------------------------------------
Option Explicit
Private Declare Function mciGetErrorString Lib "winmm.dll" Alias "mciGetErrorStringA" (ByVal dwError As Long, ByVal lpstrBuffer As String, ByVal uLength As Long) As Long
Private Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Private mmOpen As String, sec As Integer, mins As Integer
Enum EnumSoundFlags
 SND_APPLICATION = &H80
 SND_ALIAS = &H10000
 SND_ALIAS_ID = &H110000
 SND_ASYNC = &H1
 SND_FILENAME = &H20000
 SND_LOOP = &H8
 SND_MEMORY = &H4
 SND_NODEFAULT = &H2
 SND_NOSTOP = &H10
 SND_NOWAIT = &H2000
 SND_PURGE = &H40
 SND_RESOURCE = &H40004
 SND_SYNC = &H0
End Enum

Public Sub AboutBox()
frmAbout.Show vbModal
End Sub

Public Function IsPlaying() As Boolean
Attribute IsPlaying.VB_Description = "Returns boolean specifying if file is being played."
Static s As String * 30
    mciSendString "status MP3Play mode", s, Len(s), 0
    IsPlaying = (Mid$(s, 1, 7) = "playing")
End Function

Public Function MP3Play(wndHandle As Long, sFileName As String)
Attribute MP3Play.VB_Description = "Play"
Dim cmdToDo As String * 255
Dim dwReturn As Long
Dim ret As String * 128

Dim tmp As String * 255
Dim lenShort As Long
Dim ShortPathAndFie As String, glo_HWND As Long
    
    If Dir(sFileName) = "" Then
        mmOpen = "Error with input file"
        Exit Function
    End If
    lenShort = GetShortPathName(sFileName, tmp, 255)
    ShortPathAndFie = Left$(tmp, lenShort)
    glo_HWND = wndHandle
    cmdToDo = "open " & ShortPathAndFie & " type MPEGVideo Alias MP3Play"
    dwReturn = mciSendString(cmdToDo, 0&, 0&, 0&)

    If dwReturn <> 0 Then  'not success
        mciGetErrorString dwReturn, ret, 128
        mmOpen = ret
        MsgBox ret, vbCritical
        Exit Function
    End If
    mmOpen = "Success"
    mciSendString "play MP3Play", 0, 0, 0
End Function

Function PlayWave(File As String)
Attribute PlayWave.VB_Description = "Play WAV file"
PlaySound File, ByVal 0&, SND_FILENAME Or SND_ASYNC
End Function

Public Function MP3Pause()
Attribute MP3Pause.VB_Description = "Pause"
    mciSendString "pause MP3Play", 0, 0, 0
End Function

Public Function MP3Stop() As String
Attribute MP3Stop.VB_Description = "Stop"
    mciSendString "stop MP3Play", 0, 0, 0
    mciSendString "close MP3Play", 0, 0, 0
End Function

Public Function PosInSec() As Integer
Attribute PosInSec.VB_Description = "Position in seconds"
Static s As String * 30
    mciSendString "set MP3Play time format milliseconds", 0, 0, 0
    mciSendString "status MP3Play position", s, Len(s), 0
    PosInSec = Round(Mid$(s, 1, Len(s)) / 1000)
End Function

Public Function Position() As String
Attribute Position.VB_Description = "Position"
Static s As String * 30
Dim sec As Integer
Dim mins As Integer
    mciSendString "set MP3Play time format milliseconds", 0, 0, 0
    mciSendString "status MP3Play position", s, Len(s), 0
    sec = Round(Mid$(s, 1, Len(s)) / 1000)
    If sec < 60 Then Position = "0:" & Format(sec, "00")
    If sec > 59 Then
        mins = Int(sec / 60)
        sec = sec - (mins * 60)
        Position = Format(mins, "00") & ":" & Format(sec, "00")
    End If
End Function

Public Function LenInSec()
Attribute LenInSec.VB_Description = "Length in seconds."
Static s As String * 30
    mciSendString "set MP3Play time format milliseconds", 0, 0, 0
    mciSendString "status MP3Play length", s, Len(s), 0
    LenInSec = Round(Val(Mid$(s, 1, Len(s))) / 1000) 'Round(CInt(Mid$(s, 1, Len(s))) / 1000)
End Function

Public Function FileLength()
Attribute FileLength.VB_Description = "Returns the file length (play time)."
Static s As String * 30
    mciSendString "set MP3Play time format milliseconds", 0, 0, 0
    mciSendString "status MP3Play length", s, Len(s), 0
    sec = Round(Val(Mid$(s, 1, Len(s))) / 1000) 'Round(CInt(Mid$(s, 1, Len(s))) / 1000)
    If sec < 60 Then FileLength = "0:" & Format(sec, "00")
    If sec > 59 Then
        mins = Int(sec / 60)
        sec = sec - (mins * 60)
        FileLength = Format(mins, "00") & ":" & Format(sec, "00")
    End If
End Function

Public Function MP3Seek(MilliSecond As Long)
Attribute MP3Seek.VB_Description = "Seek to second/time"
    mciSendString "set MP3Play time format milliseconds", 0, 0, 0
    If IsPlaying = True Then mciSendString "play MP3Play from " & MilliSecond, 0, 0, 0
    If IsPlaying = False Then mciSendString "seek MP3Play to " & MilliSecond, 0, 0, 0
End Function
