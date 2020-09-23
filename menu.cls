VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Menus"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'----------------------------------------------
'Visual Basic Runtime Procedures Extension
'Sushant Pandurangi <sushant@phreaker.net>
'-----------------------------------------------
'Most of this code is taken from Oliver Martin
'Work was based on Paul Dilascia's from the
'MS System Journal (see www.dilascia.com)
'-----------------------------------------------
Option Explicit

Sub AboutBox()
frmAbout.Show vbModal
End Sub

Function Activate(hWnd As Long, Optional ImageList As Object, Optional HelpObject As HelpObj, Optional MainMenu As Boolean) As Boolean
Attribute Activate.VB_Description = "Activate menus."
Activate = mCoolMenu.Install(hWnd, HelpObject, ImageList, False, MainMenu)
End Function

Function Check3D(pForm As Object, Optional Value As Boolean) As Boolean
Attribute Check3D.VB_Description = "Specifies if checks should appear in 3D or not."
Check3D = ComplexChecks(pForm.hWnd, Value)
End Function

Function Emboss(pForm As Object, Optional Value As Boolean) As Boolean
Attribute Emboss.VB_Description = "Sets/returns property to emboss disabled items in colour."
Emboss = ColorEmbossed(pForm.hWnd, Value)
End Function

Function SelectFull(pForm As Object, Optional Value As Boolean) As Boolean
Attribute SelectFull.VB_Description = "Returns or sets boolean specifying if or not full reigon should be selected."
SelectFull = FullSelect(pForm.hWnd, Value)
End Function

Function MenuFont(pForm As Object, Optional FontName As String) As String
Attribute MenuFont.VB_Description = "Returns or sets menu font."
MenuFont = mCoolMenu.FontName(pForm.hWnd, FontName)
End Function

Function MenuSize(pForm As Object, Optional Size As Integer) As Long
Attribute MenuSize.VB_Description = "Returns or sets menu fontsize."
MenuSize = FontSize(pForm.hWnd, Size)
End Function

Function Deactivate(pForm As Object) As Boolean
Attribute Deactivate.VB_Description = "Remove Menus."
Deactivate = Uninstall(pForm.hWnd)
End Function

Function MDIChild(pForm As Object)
Attribute MDIChild.VB_Description = "Put this in Form_Load of MDIChildren."
MDIChildMenu (pForm.hWnd)
End Function

Function MenuColor(pForm As Object, Optional Value As Long) As Long
Attribute MenuColor.VB_Description = "Sets/returns menu colour."
MenuColor = mCoolMenu.ForeColor(pForm.hWnd, Value)
End Function

Function SelColor(pForm As Object, Optional Value As Long) As Long
Attribute SelColor.VB_Description = "Returns or sets selection colour."
SelColor = mCoolMenu.SelectColor(pForm.hWnd, Value)
End Function
