VERSION 5.00
Begin VB.Form frmMsg 
   BackColor       =   &H80000004&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1380
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   4575
   ControlBox      =   0   'False
   Icon            =   "msgbox.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1380
   ScaleWidth      =   4575
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FF8080&
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3420
      TabIndex        =   1
      Top             =   945
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000004&
      ForeColor       =   &H00000000&
      Height          =   125
      Left            =   90
      TabIndex        =   0
      Top             =   1080
      Width           =   3255
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   135
      Picture         =   "msgbox.frx":0BC2
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   3
      Top             =   405
      Width           =   480
   End
   Begin VB.PictureBox Picture2 
      Align           =   1  'Align Top
      BackColor       =   &H80000002&
      ForeColor       =   &H80000009&
      Height          =   285
      Left            =   0
      ScaleHeight     =   225
      ScaleWidth      =   4515
      TabIndex        =   4
      Top             =   0
      Width           =   4575
      Begin VB.Image Image2 
         Height          =   240
         Left            =   0
         Picture         =   "msgbox.frx":0ECC
         Top             =   -15
         Width           =   240
      End
      Begin VB.Image Image3 
         Height          =   195
         Left            =   4275
         Picture         =   "msgbox.frx":1016
         Top             =   3
         Width           =   225
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000009&
         Height          =   195
         Left            =   270
         TabIndex        =   5
         Top             =   0
         Width           =   45
      End
   End
   Begin VB.Label Label1 
      ForeColor       =   &H00000000&
      Height          =   600
      Left            =   810
      TabIndex        =   2
      Top             =   360
      Width           =   3660
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmMsg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub EasyMove(pForm As Form)
ReleaseCapture
SendMessage pForm.hWnd, &HA1, 2, 0&
End Sub


Private Sub Command1_Click()
Unload Me
End Sub


Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
EasyMove Me
End Sub

Private Sub Picture2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
EasyMove Me
End Sub
