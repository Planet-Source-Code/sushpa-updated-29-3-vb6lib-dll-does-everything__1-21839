VERSION 5.00
Begin VB.Form frmAbout 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3315
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   4890
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3315
   ScaleWidth      =   4890
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmOK 
      Caption         =   "Clo&se"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   3780
      MaskColor       =   &H00FFFFFF&
      Picture         =   "FORM.frx":0000
      TabIndex        =   5
      Top             =   2655
      UseMaskColor    =   -1  'True
      Width           =   915
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BackColor       =   &H80000002&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   0
      ScaleHeight     =   225
      ScaleWidth      =   4830
      TabIndex        =   0
      Top             =   0
      Width           =   4890
      Begin VB.Image Image2 
         Height          =   240
         Left            =   0
         Picture         =   "FORM.frx":0102
         Top             =   -15
         Width           =   240
      End
      Begin VB.Image Image1 
         Height          =   195
         Left            =   4590
         Picture         =   "FORM.frx":024C
         ToolTipText     =   "Close"
         Top             =   0
         Width           =   225
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "About vb6lib.dll - by Sushant Pandurangi"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   270
         TabIndex        =   3
         Top             =   0
         Width           =   2865
      End
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Credits and thanks:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   180
      TabIndex        =   7
      Top             =   2520
      Width           =   1410
   End
   Begin VB.Image Image3 
      Height          =   480
      Left            =   3825
      Picture         =   "FORM.frx":04FE
      Top             =   1620
      Width           =   480
   End
   Begin VB.Image Image4 
      Height          =   480
      Left            =   3960
      Picture         =   "FORM.frx":0808
      Top             =   1800
      Width           =   480
   End
   Begin VB.Image Image7 
      Height          =   480
      Left            =   3645
      Picture         =   "FORM.frx":0B12
      Top             =   1800
      Width           =   480
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Oliver Martin - Menus       Andrea Batina - Registry"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   420
      Left            =   180
      TabIndex        =   6
      Top             =   2745
      Width           =   2130
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Enhanced Functions & Subs for Visual Basic 5/6 and upwards."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   510
      Left            =   630
      TabIndex        =   4
      Top             =   1440
      UseMnemonic     =   0   'False
      Width           =   2280
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "(c) Sushant Pandurangi, 2000-2001."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   180
      TabIndex        =   2
      Top             =   2295
      Width           =   2625
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "By Sushant Pandurangi"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   195
      Left            =   180
      MouseIcon       =   "FORM.frx":0E1C
      MousePointer    =   99  'Custom
      TabIndex        =   1
      ToolTipText     =   "http://sushantshome.tripod.com"
      Top             =   2070
      Width           =   1665
   End
   Begin VB.Image Image5 
      Height          =   3645
      Left            =   -450
      Picture         =   "FORM.frx":1126
      Top             =   -270
      Width           =   6000
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmOK_Click()
Unload Me
End Sub

Private Sub Image1_Click()
Unload Me
End Sub

Private Sub Image2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
ReleaseCapture
SendMessage hWnd, &HA1, 2, 0&
End Sub

Private Sub Image6_Click(Index As Integer)

End Sub

Private Sub Image3_Click()

End Sub

Private Sub Label2_Click()
ShellExecute 0, "open", "http://sushantshome.tripod.com", "", "", 10
End Sub

Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
ReleaseCapture
SendMessage hWnd, &HA1, 2, 0&
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
ReleaseCapture
SendMessage hWnd, &HA1, 2, 0&
End Sub
