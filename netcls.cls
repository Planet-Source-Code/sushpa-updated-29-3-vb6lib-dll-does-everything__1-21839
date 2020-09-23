VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Network"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'----------------------------------------------
'Visual Basic Runtime Procedures Extension
'Sushant Pandurangi <sushant@phreaker.net>
'-----------------------------------------------
Option Explicit
'--------------------------------------------------
Public Enum EnumWindowState
 SHOW_MAXIMIZED = 3
 SHOW_MINIMIZED = 2
 SHOW_DEFAULT = 10
 SHOW_MIN_ACTIVE = 7
 SHOW_NORMAL = 1
End Enum


Function SendPing(szHostname As String, szTripTime As String, bDataMatch As Boolean, sDataSize As Long, sTimeOut As Long)
Attribute SendPing.VB_Description = "Send a ping echo to find out whether the given host is active."
Ping szHostname, szTripTime, bDataMatch, sDataSize, sTimeOut
End Function

Function GetAddress(Optional szHostname As String)
Attribute GetAddress.VB_Description = "Get IP address."
GetAddress = GetIPAddress(szHostname)
End Function

Function GetName(szIPAddress As String)
Attribute GetName.VB_Description = "Get Host from given IP Address."
GetName = GetHostFromIP(szIPAddress)
End Function

Function SendEMail(Optional szEmailAddress As String)
Attribute SendEMail.VB_Description = "Send Email to the address."
Dim lhWnd As Long
Dim Display As EnumWindowState
Display = SHOW_DEFAULT
ShellExecute lhWnd, "open", "mailto:" & szEmailAddress, vbNullString, vbNullString, Display
End Function

Public Function OpenPage(Optional szURL As String, Optional WindowState As EnumWindowState = SHOW_DEFAULT) As Long
Attribute OpenPage.VB_Description = "Open an Internet web page."
    Dim lhWnd As Long, lAns As Long
    lAns = ShellExecute(lhWnd, "open", szURL, vbNullString, vbNullString, WindowState)
    OpenPage = lAns
End Function

Public Function GetLongIP(szIPAddress As String)
Attribute GetLongIP.VB_Description = "Get IP Address converted to long form."
Dim IP As String
IP = GetIPAddress(szIPAddress)
GetLongIP = AddressStringToLong(IP)
End Function

Function NetConnect()
Attribute NetConnect.VB_Description = "Connect to the net."
Dim lResult As Long
lResult = InternetAutodial(Internet_Autodial_Force_Unattended, 0&)
End Function

Function Disconnect()
Attribute Disconnect.VB_Description = "Disconnect from the net if connected."
Dim lResult As Long
lResult = InternetAutodialHangup(0&)
End Function

Sub AboutBox()
frmAbout.Show vbModal
End Sub

