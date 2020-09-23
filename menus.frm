VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   Caption         =   "Menu Demo"
   ClientHeight    =   1920
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   1920
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      Height          =   285
      Left            =   0
      ScaleHeight     =   225
      ScaleWidth      =   4620
      TabIndex        =   1
      Top             =   1635
      Width           =   4680
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   45
         TabIndex        =   2
         Top             =   0
         Width           =   45
      End
   End
   Begin MSComctlLib.ImageList imlMain 
      Left            =   4140
      Top             =   315
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "menus.frx":0000
            Key             =   ""
            Object.Tag             =   "&Item"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "menus.frx":015C
            Key             =   ""
            Object.Tag             =   "E&xit"
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      Caption         =   $"menus.frx":02B8
      Height          =   645
      Left            =   45
      TabIndex        =   0
      Top             =   45
      Width           =   4605
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuItem 
         Caption         =   "|Item Description|&Item"
         Shortcut        =   ^D
      End
      Begin VB.Menu mnuRadio 
         Caption         =   "*|Option|Radio"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuCheck 
         Caption         =   "#|CheckBox|Check"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuSep 
         Caption         =   "-QUIT"
      End
      Begin VB.Menu muBye 
         Caption         =   "|Exit the application|E&xit"
         Shortcut        =   ^B
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public WithEvents hObj As HelpObj
Attribute hObj.VB_VarHelpID = -1
Public Menu As New Menus

Private Sub Form_Load()
Set hObj = New HelpObj
Activate Me.hWnd, imlMain, hObj, True
End Sub

Private Sub Form_Unload(Cancel As Integer)
Menu.Deactivate Me
Set hObj = Nothing
End Sub

Private Sub hObj_MenuHelp(ByVal MenuText As String, ByVal MenuHelp As String, ByVal Enabled As Boolean)
Label2.Caption = MenuHelp
End Sub

Private Sub muBye_Click()
Unload Me 'DONT PUT END HERE
End Sub
