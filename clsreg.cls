VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Registry"
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
Public DefaultPath As String
Attribute DefaultPath.VB_VarProcData = ";Data"
Attribute DefaultPath.VB_VarDescription = "Default path for saving settings."
Public ListValues As Collection
Attribute ListValues.VB_VarDescription = "Collection of Registry values of a specified HKEY."
Public ListKeys As Collection
Attribute ListKeys.VB_VarDescription = "Collection of Registry keys of a specified HKEY."
'-----------------------------------------------
Enum EnumRegConstants
 KEY_QUERY_VALUE = &H1
 KEY_SET_VALUE = &H2
 KEY_CREATE_SUB_KEY = &H4
 KEY_ENUMERATE_SUB_KEYS = &H8
 KEY_NOTIFY = &H10
 KEY_CREATE_LINK = &H20
 KEY_ALL_ACCESS = KEY_QUERY_VALUE And KEY_ENUMERATE_SUB_KEYS And KEY_NOTIFY And KEY_CREATE_SUB_KEY And KEY_CREATE_LINK And KEY_SET_VALUE
 REG_OPTION_NON_VOLATILE = 0
 REG_OPTION_VOLATILE = 1
 REG_SZ = 1
 ERROR_SUCCESS = 0&
End Enum
'--------------------------------------------------
Public Enum EnumRegKeys
 HKEY_CLASSES_ROOT = &H80000000
 HKEY_CURRENT_CONFIG = &H80000005
 HKEY_CURRENT_USER = &H80000001
 HKEY_DYN_DATA = &H80000006
 HKEY_LOCAL_MACHINE = &H80000002
 HKEY_PERFORMANCE_DATA = &H80000004
 HKEY_USERS = &H80000003
End Enum
'-----------------------------------------------

Public Function CreateKey(hKey As EnumRegKeys, SubKey As String)
Attribute CreateKey.VB_Description = "Create a new Registry Key"
    Dim lngRet As Long
    Dim lngResult As Long
    Dim lngDis As Long
    lngRet = RegCreateKeyEx(hKey, SubKey, 0&, 0&, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, 0&, lngResult, lngDis)
    lngRet = RegCloseKey(lngResult) 'Close key
End Function

Public Function DeleteKey(hKey As EnumRegKeys, SubKey As String)
Attribute DeleteKey.VB_Description = "Delete a given key."
    RegDeleteKey hKey, SubKey 'Delete key
End Function

Public Function SetValue(hKey As EnumRegKeys, SubKey As String, ValueName As String, sValue As String)
Attribute SetValue.VB_Description = "Set the value to a given setting."
    Dim lngRet As Long
    Dim lngResult As Long
    lngRet = RegOpenKeyEx(hKey, SubKey, 0, KEY_ALL_ACCESS, lngResult)
    If lngRet = ERROR_SUCCESS Then
    RegSetValueEx lngResult, ValueName, 0, REG_SZ, ByVal sValue, Len(sValue)
    RegFlushKey lngResult
    RegCloseKey lngResult
    End If
End Function

Public Function GetValue(hKey As EnumRegKeys, SubKey As String, ValueName As String, Optional Default As String = "")
Attribute GetValue.VB_Description = "Get the value of a specified setting."
    Dim lngRet As Long
    Dim lngResult As Long
    Dim sData As String
    lngRet = RegOpenKeyEx(hKey, SubKey, 0, KEY_ALL_ACCESS, lngResult)
    If lngRet = ERROR_SUCCESS Then 'If key exist
        sData = String(128, vbNullChar) 'Fill buffer with null chars
        lngRet = RegQueryValueEx(lngResult, ValueName, 0, REG_SZ, ByVal sData, Len(sData))
        If Not lngRet = ERROR_SUCCESS Then GetValue = Default: Exit Function
        GetValue = Left(sData, InStr(1, sData, vbNullChar) - 1)
        RegCloseKey lngResult
    Else
        GetValue = Default
    End If
End Function

Public Function DeleteValue(hKey As EnumRegKeys, SubKey As String, ValueName As String)
Attribute DeleteValue.VB_Description = "Delete the specified value from the registry."
    Dim lngRet As Long
    Dim lngResult As Long
    lngRet = RegOpenKeyEx(hKey, SubKey, 0, KEY_ALL_ACCESS, lngResult)
    RegDeleteValue lngResult, ValueName
End Function

Public Function EnumValues(hKey As EnumRegKeys, SubKey As String)
Attribute EnumValues.VB_Description = "Enum all the values of a HKEY and add them to cEnumValues."
    Dim lngRet As Long
    Dim lngResult As Long
    Dim sData As String
    Dim intIndex As Integer
    lngRet = RegOpenKeyEx(hKey, SubKey, 0, KEY_ALL_ACCESS, lngResult)
    If lngRet = ERROR_SUCCESS Then
        Set ListValues = New Collection
        Do
            sData = String(128, vbNullChar)
            lngRet = RegEnumValue(lngResult, intIndex, sData, Len(sData), 0, ByVal 0&, ByVal 0&, ByVal 0&)
            If lngRet <> 0 Then Exit Do 'If there are no more values exit do
            ListValues.Add Left(sData, InStr(1, sData, vbNullChar) - 1) 'Add values
            intIndex = intIndex + 1 'Increase counter by 1
        Loop
        RegCloseKey lngResult 'Close key
    End If
End Function

Public Function EnumKeys(hKey As EnumRegKeys, SubKey As String)
Attribute EnumKeys.VB_Description = "Enum all the keys of a HKEY and add them to cEnumKeys."
    Dim lngRet As Long
    Dim lngResult As Long
    Dim sData As String
    Dim intIndex As Integer
    lngRet = RegOpenKeyEx(hKey, SubKey, 0, KEY_ALL_ACCESS, lngResult)
    If lngRet = ERROR_SUCCESS Then 'If key exist
        Set ListKeys = New Collection 'Make new collection
        Do
            sData = String(128, vbNullChar) 'Fill buffer with null chars
            lngRet = RegEnumKey(lngResult, intIndex, sData, Len(sData))
            If lngRet <> 0 Then Exit Do 'If there are no more keys exit do
            ListKeys.Add Left(sData, InStr(1, sData, vbNullChar) - 1)  'Add keys
            intIndex = intIndex + 1 'Increase counter by 1
        Loop
        RegCloseKey lngResult 'Close key
    End If
End Function

Sub AboutBox()
frmAbout.Show vbModal
End Sub

