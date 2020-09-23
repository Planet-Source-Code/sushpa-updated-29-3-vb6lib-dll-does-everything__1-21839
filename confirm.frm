VERSION 5.00
Begin VB.Form frmConf 
   BackColor       =   &H80000004&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1350
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   4545
   ControlBox      =   0   'False
   Icon            =   "confirm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1350
   ScaleWidth      =   4545
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Cancel          =   -1  'True
      Caption         =   "C&ancel"
      Height          =   375
      Left            =   45
      TabIndex        =   4
      Top             =   945
      Width           =   915
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000004&
      ForeColor       =   &H00000000&
      Height          =   125
      Left            =   990
      TabIndex        =   3
      Top             =   1035
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H80000004&
      Caption         =   "&No"
      Height          =   375
      Left            =   3780
      TabIndex        =   2
      Top             =   945
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H80000004&
      Caption         =   "&Yes"
      Default         =   -1  'True
      Height          =   375
      Left            =   2835
      TabIndex        =   1
      Top             =   945
      Width           =   915
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BackColor       =   &H80000002&
      ForeColor       =   &H80000009&
      Height          =   285
      Left            =   0
      ScaleHeight     =   225
      ScaleWidth      =   4485
      TabIndex        =   5
      Top             =   0
      Width           =   4545
      Begin VB.Image Image3 
         Height          =   195
         Left            =   4260
         Picture         =   "confirm.frx":0BC2
         Top             =   15
         Width           =   225
      End
      Begin VB.Image Image2 
         Height          =   240
         Left            =   0
         Picture         =   "confirm.frx":0E74
         Top             =   -15
         Width           =   240
      End
      Begin VB.Label lTitle 
         AutoSize        =   -1  'True
         BackColor       =   &H80000002&
         ForeColor       =   &H80000009&
         Height          =   195
         Left            =   270
         TabIndex        =   6
         Top             =   0
         Width           =   45
      End
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   135
      Picture         =   "confirm.frx":0FBE
      Top             =   405
      Width           =   480
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000004&
      ForeColor       =   &H00000000&
      Height          =   600
      Left            =   675
      TabIndex        =   0
      Top             =   315
      Width           =   3750
   End
End
Attribute VB_Name = "frmConf"
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
ConfirmResult = "Yes"
Unload Me
End Sub

Private Sub Command2_Click()
ConfirmResult = "No"
Unload Me
End Sub

Private Sub Command3_Click()
ConfirmResult = "Cancel"
Unload Me
End Sub


Private Sub lTitle_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
EasyMove Me
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
EasyMove Me
End Sub
