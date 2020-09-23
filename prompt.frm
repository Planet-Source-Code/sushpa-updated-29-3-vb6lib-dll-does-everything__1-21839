VERSION 5.00
Begin VB.Form frmPrompt 
   BackColor       =   &H80000004&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1560
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   4485
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H00000000&
   Icon            =   "prompt.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1560
   ScaleWidth      =   4485
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Text1 
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   45
      TabIndex        =   4
      Top             =   1175
      Width           =   2940
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H80000004&
      Caption         =   "&Hide"
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   45
      TabIndex        =   5
      Top             =   945
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000004&
      ForeColor       =   &H00000000&
      Height          =   120
      Left            =   765
      TabIndex        =   2
      Top             =   945
      Width           =   3660
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H80000004&
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3690
      Picture         =   "prompt.frx":0BC2
      TabIndex        =   1
      Top             =   1125
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H80000004&
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3060
      Picture         =   "prompt.frx":0D0C
      TabIndex        =   0
      Top             =   1125
      Width           =   600
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BackColor       =   &H80000002&
      Height          =   285
      Left            =   0
      ScaleHeight     =   225
      ScaleWidth      =   4425
      TabIndex        =   6
      Top             =   0
      Width           =   4485
      Begin VB.Image Image2 
         Height          =   240
         Left            =   0
         Picture         =   "prompt.frx":0E56
         Top             =   -15
         Width           =   240
      End
      Begin VB.Image Image3 
         Height          =   195
         Left            =   4200
         Picture         =   "prompt.frx":0FA0
         Top             =   3
         Width           =   225
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000009&
         Height          =   195
         Left            =   270
         TabIndex        =   7
         Top             =   0
         Width           =   45
      End
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      ForeColor       =   &H00000000&
      Height          =   600
      Left            =   675
      TabIndex        =   3
      Top             =   360
      UseMnemonic     =   0   'False
      Width           =   3705
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   90
      Picture         =   "prompt.frx":1252
      Top             =   360
      Width           =   480
   End
End
Attribute VB_Name = "frmPrompt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub EasyMove(pForm As Form)
ReleaseCapture
SendMessage pForm.hWnd, &HA1, 2, 0&
End Sub

Private Sub Check1_Click()
If Check1.Value = 1 Then Text1.PasswordChar = "*" Else Text1.PasswordChar = ""
End Sub

Private Sub Command1_Click()
InputResult = Text1.Text
Unload Me
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
On Error Resume Next
Text1.SelStart = 0
Text1.SelLength = Len(Text1.Text)
Text1.SetFocus

End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
EasyMove Me
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
EasyMove Me
End Sub
