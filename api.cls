VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WinAPI"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'--------------------------------------------------------
'Visual Basic Runtime Procedures Extension
'Sushant Pandurangi <sushant@phreaker.net>
'--------------------------------------------------------
Option Explicit
Private Const PROCESSOR_ALPHA_21064 = 21064
Private Const PROCESSOR_INTEL_386 = 386
Private Const PROCESSOR_INTEL_486 = 486
Private Const PROCESSOR_INTEL_PENTIUM = 586
Private Const PROCESSOR_MIPS_R4000 = 4000
Private Declare Sub GetSystemInfo Lib "kernel32" (lpSystemInfo As SYSTEM_INFO)
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Private Const VER_PLATFORM_WIN32_NT = 2
Private Const VER_PLATFORM_WIN32_WINDOWS = 1
Private Const VER_PLATFORM_WIN32s = 0
Private Type SYSTEM_INFO
        dwOemID As Long
        dwPageSize As Long
        lpMinimumApplicationAddress As Long
        lpMaximumApplicationAddress As Long
        dwActiveProcessorMask As Long
        dwNumberOrfProcessors As Long
        dwProcessorType As Long
        dwAllocationGranularity As Long
        dwReserved As Long
End Type
Private Type OSVERSIONINFO
  dwOSVersionInfoSize As Long
  dwMajorVersion As Long
  dwMinorVersion As Long
  dwBuildNumber As Long
  dwPlatformId As Long
  szCSDVersion As String * 128
End Type
Public Enum EnumPropertytype
 SYSTEM_SETTINGS = 1
 INTERNET_SETTINGS = 2
 MODEM_SETTINGS = 3
 ADD_REMOVE_APPS = 4
 ADD_NEW_HARDWARE = 5
 SOUNDS_SETTINGS = 6
 NETWORK_SETTINGS = 7
 MOUSE_SETTINGS = 8
 KEYBOARD_SETTINGS = 9
 TIME_DATE_SETTINGS = 10
 REGIONAL_SETTINGS = 11
 PASSWORD_SETTINGS = 12
 DISPLAY_SETTINGS = 13
End Enum
Private Type MEMORYSTATUS
        dwLength As Long
        dwMemoryLoad As Long
        dwTotalPhys As Long
        dwAvailPhys As Long
        dwTotalPageFile As Long
        dwAvailPageFile As Long
        dwTotalVirtual As Long
        dwAvailVirtual As Long
End Type
Public Enum EnumPositions
 TOPMOST_TRUE = -1
 TOPMOST_FALSE = -2
End Enum
Public Enum EnumExitType
 LOGOFF = 0
 BYEBYE = 1
 REBOOT = 2
End Enum
Dim systemInfo As SYSTEM_INFO
Private Declare Sub GlobalMemoryStatus Lib "kernel32" (lpBuffer As MEMORYSTATUS)


Sub AboutBox()
frmAbout.Show vbModal
End Sub

Public Sub Progress(ByVal bShowProgressBar As Boolean, pForm As Object, Progressbar As Object, Statusbar As Object, Panel As Integer)
Attribute Progress.VB_Description = "Show a progress bar in the status bar, integrated."
    Dim tRC As RECT
    If bShowProgressBar Then
        SendMessageAny Statusbar.hWnd, SB_GETRECT, Panel, tRC
        With tRC
            .Top = (.Top * Screen.TwipsPerPixelY)
            .Left = (.Left * Screen.TwipsPerPixelX)
            .Bottom = (.Bottom * Screen.TwipsPerPixelY) - .Top
            .Right = (.Right * Screen.TwipsPerPixelX) - .Left
        End With
        With Progressbar
            SetParent .hWnd, Statusbar.hWnd
            .Move tRC.Left, tRC.Top, tRC.Right, tRC.Bottom
            .Visible = True
            .Value = 0
        End With
    Else
        SetParent Progressbar.hWnd, pForm.hWnd
        Progressbar.Visible = False
    End If
End Sub

Public Sub DeleteFile(szFileName As String)
Attribute DeleteFile.VB_Description = "Deletes a file from the given path."
Dim SHop As SHFILEOPSTRUCT
Dim strFile As String
strFile = szFileName
With SHop
    .wFunc = FO_DELETE
    .pFrom = strFile
    .fFlags = FOF_ALLOWUNDO
End With
SHFileOperation SHop
End Sub

Public Sub CopyFile(szFileSource As String, szFileDestination As String)
Attribute CopyFile.VB_Description = "Copies a file from the source to the destination"
Dim SHop As SHFILEOPSTRUCT
With SHop
    .wFunc = FO_COPY
    .pTo = szFileDestination
    .pFrom = szFileSource
    .fFlags = FOF_ALLOWUNDO
End With
SHFileOperation SHop
End Sub

Public Sub MoveFile(szFileSource As String, szFileDestination As String)
Attribute MoveFile.VB_Description = "Moves a file from the source to the destination"
Dim SHop As SHFILEOPSTRUCT
With SHop
    .wFunc = FO_MOVE
    .pFrom = szFileSource
    .pTo = szFileDestination
    .fFlags = FOF_ALLOWUNDO
End With
SHFileOperation SHop
End Sub

Public Function PathExists(szPath As String) As Boolean
Attribute PathExists.VB_Description = "Check if a specified path exists and return true or false."
PathExists = (Dir(szPath) <> "")
End Function

Function ShutDown(ExitType As EnumExitType)
Attribute ShutDown.VB_Description = "Shutdown,Restart or log off windows."
Dim lngResult As Long
lngResult = ExitWindowsEx(ExitType, 0&)
End Function

Function EasyMove(pForm As Object)
Attribute EasyMove.VB_Description = "Enables moving the form by dragging any control."
ReleaseCapture
SendMessage pForm.hWnd, &HA1, 2, 0&
End Function

Sub StandBy(MilliSeconds As Long)
Attribute StandBy.VB_Description = "Standby for the specified time."
Sleep MilliSeconds
End Sub

Sub Properties(sType As EnumPropertytype)
Attribute Properties.VB_Description = "System properties."
Dim dblReturn
If sType = ADD_REMOVE_APPS Then
dblReturn = Shell("rundll32.exe shell32.dll,Control_RunDLL appwiz.cpl,,1", 5)
ElseIf sType = ADD_NEW_HARDWARE Then
dblReturn = Shell("rundll32.exe shell32.dll,Control_RunDLL sysdm.cpl @1", 5)
ElseIf sType = INTERNET_SETTINGS Then
dblReturn = Shell("rundll32.exe shell32.dll,Control_RunDLL inetcpl.cpl,,0", 5)
ElseIf sType = KEYBOARD_SETTINGS Then
dblReturn = Shell("rundll32.exe shell32.dll,Control_RunDLL main.cpl @1", 5)
ElseIf sType = MODEM_SETTINGS Then
dblReturn = Shell("rundll32.exe shell32.dll,Control_RunDLL modem.cpl", 5)
ElseIf sType = MOUSE_SETTINGS Then
dblReturn = Shell("rundll32.exe shell32.dll,Control_RunDLL main.cpl @0", 5)
ElseIf sType = NETWORK_SETTINGS Then
dblReturn = Shell("rundll32.exe shell32.dll,Control_RunDLL netcpl.cpl", 5)
ElseIf sType = SOUNDS_SETTINGS Then
dblReturn = Shell("rundll32.exe shell32.dll,Control_RunDLL mmsys.cpl @1", 5)
ElseIf sType = SYSTEM_SETTINGS Then
dblReturn = Shell("rundll32.exe shell32.dll,Control_RunDLL sysdm.cpl,,0", 5)
ElseIf sType = DISPLAY_SETTINGS Then
dblReturn = Shell("rundll32.exe shell32.dll,Control_RunDLL desk.cpl,,0", 5)
ElseIf sType = PASSWORD_SETTINGS Then
dblReturn = Shell("rundll32.exe shell32.dll,Control_RunDLL password.cpl", 5)
ElseIf sType = REGIONAL_SETTINGS Then
dblReturn = Shell("rundll32.exe shell32.dll,Control_RunDLL intl.cpl,,0", 5)
ElseIf sType = TIME_DATE_SETTINGS Then
dblReturn = Shell("rundll32.exe shell32.dll,Control_RunDLL timedate.cpl", 5)
End If
End Sub

Public Sub FindDialog(Optional InitialDirectory As String)
Attribute FindDialog.VB_Description = "Shows the Win32 Find dialog."
ShellExecute 0, "find", IIf(InitialDirectory = "", "", InitialDirectory), vbNullString, vbNullString, 10
End Sub

Public Sub RunDialog(ByRef wndHandle As Long, Optional Title As String = "Run", Optional Description As String = "Type the name of a program to open, then click OK when finished.")
Attribute RunDialog.VB_Description = "Shows Win32 Run Dialog."
SHRunDialog wndHandle, 0, 0, Title, Description, 0
End Sub

Sub CopyText(sTextBox As Object)
Attribute CopyText.VB_Description = "Copies text from the given TextBox."
SendMessage sTextBox.hWnd, WM_COPY, 0, 0&
End Sub

Sub PasteText(sTextBox As Object)
Attribute PasteText.VB_Description = "Pastes text from the clipboard."
SendMessage sTextBox.hWnd, WM_PASTE, 0, 0&
End Sub

Sub UndoEdit(sTextBox As Object)
Attribute UndoEdit.VB_Description = "Undo the edit in a textbox."
SendMessage sTextBox.hWnd, EM_UNDO, 0, 0&
End Sub

Function UsedTime() As Long
Dim lngTickCount As Long
lngTickCount = GetTickCount
UsedTime = CStr(lngTickCount / 1000)
End Function

Function GetMemory(ReturnedTotal As Long, ReturnedAvailable As Long)
Dim memoryInfo As MEMORYSTATUS
GlobalMemoryStatus memoryInfo
ReturnedTotal = memoryInfo.dwTotalPhys
ReturnedAvailable = memoryInfo.dwAvailPhys
End Function

Function Processor() As String
GetSystemInfo systemInfo
Select Case systemInfo.dwProcessorType
Case PROCESSOR_ALPHA_21064 = 21064
Processor = "Alpha"
Case PROCESSOR_INTEL_386
Processor = "Intel 80386"
Case PROCESSOR_INTEL_486
Processor = "Intel 80486"
Case PROCESSOR_INTEL_PENTIUM
Processor = "Intel Pentium"
Case PROCESSOR_MIPS_R4000
Processor = "MIPS"
End Select
End Function

Public Function WindowsVer()
Dim infoStruct As OSVERSIONINFO
infoStruct.dwOSVersionInfoSize = Len(infoStruct)
GetVersionEx infoStruct
If infoStruct.dwPlatformId = VER_PLATFORM_WIN32_WINDOWS Then
    WindowsVer = "Windows 95/98"
Else
    WindowsVer = "Windows NT"
End If
End Function

Function SystemDir() As String
    Dim strBuf As String
    strBuf = Space$(255)
    'Get the system directory and then trim the buffer to the exact length
    'returned and add a dir sep (backslash) if the API didn't return one
    If GetSystemDirectory(strBuf, 255) > 0 Then
        'strBuf = StripTerminator(strBuf)
        'AddDirSep strBuf
       SystemDir = Left(strBuf, GetSystemDirectory(strBuf, 255))
    Else
        SystemDir = vbNullString
    End If
End Function

Function WindowsDir() As String
    Dim strBuf As String
    strBuf = Space$(255)
    'Get the windows directory and then trim the buffer to the exact length
    'returned and add a dir sep (backslash) if the API didn't return one
    If GetWindowsDirectory(strBuf, 255) > 0 Then
        WindowsDir = Left(strBuf, GetWindowsDirectory(strBuf, 255))
    Else
        WindowsDir = vbNullString
    End If
End Function

Sub InstallFont(vFontFileName As String)
'Install a new font
AddFontResource vFontFileName
End Sub

Sub FormTop(pForm As Object)
'Always on top
BringWindowToTop pForm.hWnd
End Sub

Sub CloseWnd(pForm As Object)
CloseWindow pForm.hWnd
End Sub
