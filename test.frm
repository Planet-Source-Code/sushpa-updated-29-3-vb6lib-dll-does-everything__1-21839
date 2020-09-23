VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form fMain 
   BackColor       =   &H80000004&
   Caption         =   "Visual Basic Library Demo (vb6lib.dll)"
   ClientHeight    =   3930
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   7020
   Icon            =   "test.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3930
   ScaleWidth      =   7020
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.StatusBar sBar 
      Align           =   2  'Align Bottom
      Height          =   285
      Left            =   0
      TabIndex        =   67
      Top             =   3645
      Width           =   7020
      _ExtentX        =   12383
      _ExtentY        =   503
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Picture         =   "test.frx":030A
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   2011
            MinWidth        =   2011
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imMenu 
      Left            =   6300
      Top             =   2070
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "test.frx":0466
            Key             =   ""
            Object.Tag             =   "Exit"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "test.frx":05C2
            Key             =   ""
            Object.Tag             =   "Caption3"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "test.frx":071E
            Key             =   ""
            Object.Tag             =   "Caption2"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "test.frx":087A
            Key             =   ""
            Object.Tag             =   "Caption1"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "test.frx":09D6
            Key             =   ""
            Object.Tag             =   "A&bout"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "test.frx":0B32
            Key             =   ""
            Object.Tag             =   "&Refresh"
         EndProperty
      EndProperty
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6990
      _ExtentX        =   12330
      _ExtentY        =   6376
      _Version        =   393216
      Style           =   1
      Tabs            =   9
      TabsPerRow      =   9
      TabHeight       =   520
      ShowFocusRect   =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Graphics"
      TabPicture(0)   =   "test.frx":0C92
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "pDraw"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "pGrade(0)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "pGrade(1)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "pGrade(2)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "pGrade(3)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "pGrade(4)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "pGrade(5)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "pGrade(6)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "pGrade(7)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "pGrade(8)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "pGrade(9)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "pGrade(10)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "pGrade(11)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).ControlCount=   13
      TabCaption(1)   =   "API "
      TabPicture(1)   =   "test.frx":0CAE
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "imlProps"
      Tab(1).Control(1)=   "cmStandby"
      Tab(1).Control(2)=   "cmLogoff"
      Tab(1).Control(3)=   "cmRestart"
      Tab(1).Control(4)=   "cmShut"
      Tab(1).Control(5)=   "cmNoTop"
      Tab(1).Control(6)=   "cmOnTop"
      Tab(1).Control(7)=   "cmRun"
      Tab(1).Control(8)=   "cmFind"
      Tab(1).Control(9)=   "cmUndo"
      Tab(1).Control(10)=   "cmPaste"
      Tab(1).Control(11)=   "txCopy"
      Tab(1).Control(12)=   "cmCopyT"
      Tab(1).Control(13)=   "cmProgress"
      Tab(1).Control(14)=   "pBar"
      Tab(1).Control(15)=   "cmDel"
      Tab(1).Control(16)=   "Dt2"
      Tab(1).Control(17)=   "Dt1"
      Tab(1).Control(18)=   "cmMove"
      Tab(1).Control(19)=   "cmCopy"
      Tab(1).Control(20)=   "Fl1"
      Tab(1).Control(21)=   "Dr2"
      Tab(1).Control(22)=   "Dr1"
      Tab(1).Control(23)=   "imProps"
      Tab(1).Control(24)=   "lbDrag"
      Tab(1).Control(25)=   "Label1"
      Tab(1).ControlCount=   26
      TabCaption(2)   =   "Sounds"
      TabPicture(2)   =   "test.frx":0CCA
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "lbLen"
      Tab(2).Control(1)=   "Label5"
      Tab(2).Control(2)=   "D1"
      Tab(2).Control(3)=   "D2"
      Tab(2).Control(4)=   "F1"
      Tab(2).Control(5)=   "opWAV"
      Tab(2).Control(6)=   "opMP3"
      Tab(2).Control(7)=   "cmPlay"
      Tab(2).Control(8)=   "cmStop"
      Tab(2).Control(9)=   "cmPause"
      Tab(2).Control(10)=   "slSeek"
      Tab(2).Control(11)=   "tmMP3"
      Tab(2).ControlCount=   12
      TabCaption(3)   =   "Internet"
      TabPicture(3)   =   "test.frx":0CE6
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Label6"
      Tab(3).Control(1)=   "Label7"
      Tab(3).Control(2)=   "Label8"
      Tab(3).Control(3)=   "Label9"
      Tab(3).Control(4)=   "cmPing"
      Tab(3).Control(5)=   "txHost"
      Tab(3).Control(6)=   "txData"
      Tab(3).Control(7)=   "cmHost"
      Tab(3).Control(8)=   "cmIP"
      Tab(3).Control(9)=   "cmIPLong"
      Tab(3).Control(10)=   "txURL"
      Tab(3).Control(11)=   "cmOpenURL"
      Tab(3).Control(12)=   "cmMail"
      Tab(3).Control(13)=   "cmDiscon"
      Tab(3).Control(14)=   "cmConn"
      Tab(3).ControlCount=   15
      TabCaption(4)   =   "Menus"
      TabPicture(4)   =   "test.frx":0D02
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "tCursor"
      Tab(4).Control(1)=   "Frame1"
      Tab(4).Control(2)=   "Check2"
      Tab(4).Control(3)=   "Check1"
      Tab(4).Control(4)=   "Frame2"
      Tab(4).Control(5)=   "Label3"
      Tab(4).ControlCount=   6
      TabCaption(5)   =   "General"
      TabPicture(5)   =   "test.frx":0D1E
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "cmSave"
      Tab(5).Control(1)=   "cmOpen"
      Tab(5).Control(2)=   "txOpen"
      Tab(5).Control(3)=   "txKey"
      Tab(5).Control(4)=   "txEnc"
      Tab(5).Control(5)=   "cmEnc"
      Tab(5).Control(6)=   "cmDec"
      Tab(5).Control(7)=   "cmFonts"
      Tab(5).Control(8)=   "cbFonts(1)"
      Tab(5).Control(9)=   "Label4"
      Tab(5).Control(10)=   "lbFont"
      Tab(5).ControlCount=   11
      TabCaption(6)   =   "Registry"
      TabPicture(6)   =   "test.frx":0D3A
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "lvKeys"
      Tab(6).Control(1)=   "cmChReg"
      Tab(6).Control(2)=   "lvVal"
      Tab(6).Control(3)=   "txValues"
      Tab(6).ControlCount=   4
      TabCaption(7)   =   "Strings"
      TabPicture(7)   =   "test.frx":0D56
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "cmINT"
      Tab(7).Control(1)=   "cmRep"
      Tab(7).Control(2)=   "cmRev"
      Tab(7).Control(3)=   "cmInitCap"
      Tab(7).Control(4)=   "cmBin"
      Tab(7).Control(5)=   "cmCount"
      Tab(7).Control(6)=   "cmLines"
      Tab(7).Control(7)=   "cmWords"
      Tab(7).Control(8)=   "cmPathname"
      Tab(7).Control(9)=   "cmFilename"
      Tab(7).Control(10)=   "cmUp1Level"
      Tab(7).Control(11)=   "txOUT"
      Tab(7).Control(12)=   "txIN"
      Tab(7).Control(13)=   "Label2(1)"
      Tab(7).Control(14)=   "Label2(0)"
      Tab(7).ControlCount=   15
      TabCaption(8)   =   "Messages"
      TabPicture(8)   =   "test.frx":0D72
      Tab(8).ControlEnabled=   0   'False
      Tab(8).Control(0)=   "cmCNF"
      Tab(8).Control(1)=   "cmPT"
      Tab(8).Control(2)=   "cmMsg"
      Tab(8).ControlCount=   3
      Begin VB.PictureBox pGrade 
         AutoRedraw      =   -1  'True
         Height          =   465
         Index           =   11
         Left            =   5940
         ScaleHeight     =   405
         ScaleWidth      =   900
         TabIndex        =   109
         Top             =   3015
         Width           =   960
      End
      Begin VB.PictureBox pGrade 
         AutoRedraw      =   -1  'True
         Height          =   465
         Index           =   10
         Left            =   2475
         ScaleHeight     =   405
         ScaleWidth      =   1080
         TabIndex        =   108
         Top             =   3015
         Width           =   1140
      End
      Begin VB.PictureBox pGrade 
         AutoRedraw      =   -1  'True
         Height          =   465
         Index           =   9
         Left            =   3645
         ScaleHeight     =   405
         ScaleWidth      =   1125
         TabIndex        =   107
         Top             =   3015
         Width           =   1185
      End
      Begin VB.PictureBox pGrade 
         AutoRedraw      =   -1  'True
         Height          =   465
         Index           =   8
         Left            =   4860
         ScaleHeight     =   405
         ScaleWidth      =   990
         TabIndex        =   106
         Top             =   3015
         Width           =   1050
      End
      Begin VB.PictureBox pGrade 
         AutoRedraw      =   -1  'True
         Height          =   465
         Index           =   7
         Left            =   5940
         ScaleHeight     =   405
         ScaleWidth      =   900
         TabIndex        =   105
         Top             =   2520
         Width           =   960
      End
      Begin VB.PictureBox pGrade 
         AutoRedraw      =   -1  'True
         Height          =   465
         Index           =   6
         Left            =   1260
         ScaleHeight     =   405
         ScaleWidth      =   1125
         TabIndex        =   104
         Top             =   2520
         Width           =   1185
      End
      Begin VB.PictureBox pGrade 
         AutoRedraw      =   -1  'True
         Height          =   465
         Index           =   5
         Left            =   2475
         ScaleHeight     =   405
         ScaleWidth      =   1080
         TabIndex        =   103
         Top             =   2520
         Width           =   1140
      End
      Begin VB.PictureBox pGrade 
         AutoRedraw      =   -1  'True
         Height          =   465
         Index           =   4
         Left            =   3645
         ScaleHeight     =   405
         ScaleWidth      =   1125
         TabIndex        =   102
         Top             =   2520
         Width           =   1185
      End
      Begin VB.PictureBox pGrade 
         AutoRedraw      =   -1  'True
         Height          =   465
         Index           =   3
         Left            =   4860
         ScaleHeight     =   405
         ScaleWidth      =   990
         TabIndex        =   101
         Top             =   2520
         Width           =   1050
      End
      Begin VB.PictureBox pGrade 
         AutoRedraw      =   -1  'True
         Height          =   465
         Index           =   2
         Left            =   90
         ScaleHeight     =   405
         ScaleWidth      =   1080
         TabIndex        =   100
         Top             =   3015
         Width           =   1140
      End
      Begin VB.PictureBox pGrade 
         AutoRedraw      =   -1  'True
         Height          =   465
         Index           =   1
         Left            =   1260
         ScaleHeight     =   405
         ScaleWidth      =   1125
         TabIndex        =   99
         Top             =   3015
         Width           =   1185
      End
      Begin VB.PictureBox pGrade 
         AutoRedraw      =   -1  'True
         Height          =   465
         Index           =   0
         Left            =   90
         ScaleHeight     =   405
         ScaleWidth      =   1080
         TabIndex        =   98
         Top             =   2520
         Width           =   1140
      End
      Begin MSComctlLib.ListView lvKeys 
         Height          =   2760
         Left            =   -74910
         TabIndex        =   33
         Top             =   405
         Width           =   3540
         _ExtentX        =   6244
         _ExtentY        =   4868
         View            =   3
         Arrange         =   2
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         Icons           =   "imMenu"
         SmallIcons      =   "imlProps"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Keys in \Software\Microsoft\Windows"
            Object.Width           =   4410
         EndProperty
      End
      Begin MSComctlLib.ImageList imlProps 
         Left            =   -68430
         Top             =   3060
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   3
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "test.frx":0D8E
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "test.frx":10AA
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "test.frx":1206
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.Timer tCursor 
         Interval        =   1000
         Left            =   -69240
         Top             =   3195
      End
      Begin VB.CommandButton cmConn 
         Caption         =   "&Connect"
         Height          =   555
         Left            =   -70320
         Picture         =   "test.frx":195A
         Style           =   1  'Graphical
         TabIndex        =   97
         Top             =   2880
         Width           =   1005
      End
      Begin VB.CommandButton cmDiscon 
         Caption         =   "Disc&onnect"
         Height          =   555
         Left            =   -69285
         Picture         =   "test.frx":1AA4
         Style           =   1  'Graphical
         TabIndex        =   96
         Top             =   2880
         Width           =   1140
      End
      Begin VB.CommandButton cmMail 
         Caption         =   "Send mail"
         Height          =   555
         Left            =   -74010
         Picture         =   "test.frx":1BEE
         Style           =   1  'Graphical
         TabIndex        =   95
         Top             =   2880
         Width           =   870
      End
      Begin VB.CommandButton cmOpenURL 
         Caption         =   "&Open WWW"
         Height          =   555
         Left            =   -73110
         Picture         =   "test.frx":1D38
         Style           =   1  'Graphical
         TabIndex        =   94
         Top             =   2880
         Width           =   1095
      End
      Begin VB.TextBox txURL 
         Height          =   315
         Left            =   -74415
         TabIndex        =   92
         Text            =   "http://sushantshome.tripod.com"
         Top             =   2520
         Width           =   2400
      End
      Begin VB.CommandButton cmIPLong 
         Caption         =   "       .....       Get &Long IP"
         Height          =   600
         Left            =   -73560
         Style           =   1  'Graphical
         TabIndex        =   91
         Top             =   1170
         Width           =   1095
      End
      Begin VB.CommandButton cmIP 
         Caption         =   "Get &IP"
         Height          =   600
         Left            =   -73425
         Picture         =   "test.frx":1E82
         Style           =   1  'Graphical
         TabIndex        =   90
         Top             =   1800
         Width           =   960
      End
      Begin VB.CommandButton cmHost 
         Caption         =   "Get &Host"
         Height          =   600
         Left            =   -74415
         Picture         =   "test.frx":1FCC
         Style           =   1  'Graphical
         TabIndex        =   89
         Top             =   1800
         Width           =   960
      End
      Begin VB.TextBox txData 
         Height          =   315
         Left            =   -74415
         TabIndex        =   87
         Text            =   "32"
         Top             =   810
         Width           =   555
      End
      Begin VB.TextBox txHost 
         Height          =   315
         Left            =   -74415
         TabIndex        =   84
         Text            =   "ping.symantec.com"
         Top             =   450
         Width           =   1950
      End
      Begin VB.CommandButton cmPing 
         Caption         =   "P&ing"
         Height          =   600
         Left            =   -74415
         Picture         =   "test.frx":2116
         Style           =   1  'Graphical
         TabIndex        =   83
         Top             =   1170
         Width           =   825
      End
      Begin VB.CommandButton cmStandby 
         Caption         =   "            ...           S&tandby 5 secs"
         Height          =   510
         Left            =   -69465
         TabIndex        =   82
         Top             =   1665
         Width           =   1320
      End
      Begin VB.CommandButton cmLogoff 
         Caption         =   "Log&off"
         Height          =   600
         Left            =   -68745
         Picture         =   "test.frx":2260
         Style           =   1  'Graphical
         TabIndex        =   81
         Top             =   2205
         Width           =   600
      End
      Begin VB.CommandButton cmRestart 
         Caption         =   "Restar&t"
         Height          =   600
         Left            =   -69465
         Picture         =   "test.frx":23AA
         Style           =   1  'Graphical
         TabIndex        =   80
         Top             =   2205
         Width           =   690
      End
      Begin VB.CommandButton cmShut 
         Caption         =   "Sh&ut down"
         Height          =   600
         Left            =   -69465
         Picture         =   "test.frx":24F4
         Style           =   1  'Graphical
         TabIndex        =   79
         Top             =   2835
         Width           =   1320
      End
      Begin VB.CommandButton cmNoTop 
         Caption         =   "NoT&op"
         Enabled         =   0   'False
         Height          =   600
         Left            =   -68790
         Picture         =   "test.frx":263E
         Style           =   1  'Graphical
         TabIndex        =   78
         Top             =   1035
         Width           =   645
      End
      Begin VB.CommandButton cmOnTop 
         Caption         =   "&Top"
         Enabled         =   0   'False
         Height          =   600
         Left            =   -69465
         Picture         =   "test.frx":2788
         Style           =   1  'Graphical
         TabIndex        =   77
         Top             =   1035
         Width           =   645
      End
      Begin VB.CommandButton cmRun 
         Caption         =   "&Run..."
         Height          =   600
         Left            =   -68790
         Picture         =   "test.frx":28D2
         Style           =   1  'Graphical
         TabIndex        =   76
         Top             =   405
         Width           =   645
      End
      Begin VB.CommandButton cmFind 
         Caption         =   "F&ind"
         Height          =   600
         Left            =   -69465
         Picture         =   "test.frx":2A1C
         Style           =   1  'Graphical
         TabIndex        =   75
         Top             =   405
         Width           =   645
      End
      Begin VB.CommandButton cmUndo 
         Caption         =   "&Undo"
         Height          =   555
         Left            =   -70095
         Picture         =   "test.frx":2B66
         Style           =   1  'Graphical
         TabIndex        =   73
         Top             =   1395
         Width           =   555
      End
      Begin VB.CommandButton cmPaste 
         Caption         =   "P&aste"
         Height          =   555
         Left            =   -70680
         Picture         =   "test.frx":2CB0
         Style           =   1  'Graphical
         TabIndex        =   72
         Top             =   1395
         Width           =   555
      End
      Begin VB.TextBox txCopy 
         Height          =   960
         Left            =   -71310
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   71
         Text            =   "test.frx":2DFA
         Top             =   405
         Width           =   1770
      End
      Begin VB.CommandButton cmCopyT 
         Caption         =   "Co&py"
         Height          =   555
         Left            =   -71265
         Picture         =   "test.frx":2E35
         Style           =   1  'Graphical
         TabIndex        =   70
         Top             =   1395
         Width           =   555
      End
      Begin VB.CommandButton cmProgress 
         Caption         =   "Prog&ress bar"
         Height          =   555
         Left            =   -72585
         Picture         =   "test.frx":2F7F
         Style           =   1  'Graphical
         TabIndex        =   69
         Top             =   2925
         Width           =   1275
      End
      Begin MSComctlLib.ProgressBar pBar 
         Height          =   195
         Left            =   -74070
         TabIndex        =   68
         Top             =   2925
         Visible         =   0   'False
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   344
         _Version        =   393216
         Appearance      =   1
         Scrolling       =   1
      End
      Begin VB.CommandButton cmDel 
         Caption         =   "D&el"
         Height          =   600
         Left            =   -71865
         Picture         =   "test.frx":30C9
         Style           =   1  'Graphical
         TabIndex        =   66
         ToolTipText     =   "Delete the file"
         Top             =   2250
         Width           =   555
      End
      Begin VB.DirListBox Dt2 
         Height          =   1440
         Left            =   -73035
         TabIndex        =   65
         ToolTipText     =   "Destination folder"
         Top             =   765
         Width           =   1725
      End
      Begin VB.DriveListBox Dt1 
         Height          =   315
         Left            =   -73035
         TabIndex        =   64
         ToolTipText     =   "Destination Drive"
         Top             =   405
         Width           =   1725
      End
      Begin VB.CommandButton cmMove 
         Caption         =   "M&ove"
         Height          =   600
         Left            =   -72450
         Picture         =   "test.frx":3213
         Style           =   1  'Graphical
         TabIndex        =   63
         ToolTipText     =   "Move the file"
         Top             =   2250
         Width           =   555
      End
      Begin VB.CommandButton cmCopy 
         Caption         =   "C&opy"
         Height          =   600
         Left            =   -73035
         Picture         =   "test.frx":335D
         Style           =   1  'Graphical
         TabIndex        =   62
         ToolTipText     =   "Copy the file"
         Top             =   2250
         Width           =   555
      End
      Begin VB.FileListBox Fl1 
         Height          =   870
         Left            =   -74925
         TabIndex        =   61
         ToolTipText     =   "Source file"
         Top             =   2025
         Width           =   1860
      End
      Begin VB.DirListBox Dr2 
         Height          =   1215
         Left            =   -74925
         TabIndex        =   60
         ToolTipText     =   "Source path"
         Top             =   765
         Width           =   1860
      End
      Begin VB.DriveListBox Dr1 
         Height          =   315
         Left            =   -74925
         TabIndex        =   59
         ToolTipText     =   "Source drive"
         Top             =   405
         Width           =   1860
      End
      Begin VB.Timer tmMP3 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   -68655
         Top             =   2250
      End
      Begin MSComctlLib.Slider slSeek 
         Height          =   240
         Left            =   -74820
         TabIndex        =   56
         Top             =   2925
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   423
         _Version        =   393216
         MousePointer    =   9
         Max             =   500
         TickStyle       =   3
      End
      Begin VB.CommandButton cmPause 
         Caption         =   "Pa&use"
         Height          =   510
         Left            =   -69780
         Picture         =   "test.frx":34A7
         Style           =   1  'Graphical
         TabIndex        =   55
         Top             =   1800
         Width           =   960
      End
      Begin VB.CommandButton cmStop 
         Caption         =   "S&top"
         Height          =   510
         Left            =   -69780
         OLEDropMode     =   1  'Manual
         Picture         =   "test.frx":35F1
         Style           =   1  'Graphical
         TabIndex        =   54
         Top             =   1260
         Width           =   960
      End
      Begin VB.CommandButton cmPlay 
         Caption         =   "Pl&ay"
         Height          =   510
         Left            =   -69780
         Picture         =   "test.frx":373B
         Style           =   1  'Graphical
         TabIndex        =   53
         Top             =   720
         Width           =   960
      End
      Begin VB.OptionButton opMP3 
         Caption         =   "MPEG-3 files (*.mp3)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   -71085
         TabIndex        =   52
         Top             =   2610
         Value           =   -1  'True
         Width           =   1815
      End
      Begin VB.OptionButton opWAV 
         Caption         =   "WaveForm Files (*.wav)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   -73155
         TabIndex        =   51
         Top             =   2610
         Width           =   2085
      End
      Begin VB.FileListBox F1 
         Height          =   2040
         Left            =   -71377
         Pattern         =   "*.mp3"
         TabIndex        =   50
         Top             =   540
         Width           =   1545
      End
      Begin VB.DirListBox D2 
         Height          =   1665
         Left            =   -73132
         TabIndex        =   49
         Top             =   900
         Width           =   1725
      End
      Begin VB.DriveListBox D1 
         Height          =   315
         Left            =   -73132
         TabIndex        =   48
         Top             =   540
         Width           =   1725
      End
      Begin VB.CommandButton cmSave 
         Caption         =   "S&ave"
         Height          =   555
         Left            =   -69015
         Picture         =   "test.frx":3885
         Style           =   1  'Graphical
         TabIndex        =   47
         Top             =   2970
         Width           =   870
      End
      Begin VB.CommandButton cmOpen 
         Caption         =   "O&pen"
         Height          =   555
         Left            =   -69960
         Picture         =   "test.frx":39CF
         Style           =   1  'Graphical
         TabIndex        =   46
         Top             =   2970
         Width           =   930
      End
      Begin VB.TextBox txOpen 
         Height          =   2490
         Left            =   -72435
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   45
         Text            =   "test.frx":3B19
         Top             =   450
         Width           =   4290
      End
      Begin VB.TextBox txKey 
         Height          =   315
         Left            =   -74910
         TabIndex        =   43
         Top             =   3150
         Width           =   465
      End
      Begin VB.TextBox txEnc 
         Height          =   1545
         Left            =   -74910
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   42
         Text            =   "test.frx":3B54
         Top             =   1395
         Width           =   2445
      End
      Begin VB.CommandButton cmEnc 
         Caption         =   "E&ncrypt"
         Height          =   510
         Left            =   -74415
         Picture         =   "test.frx":3BC5
         Style           =   1  'Graphical
         TabIndex        =   41
         Top             =   2970
         Width           =   1005
      End
      Begin VB.CommandButton cmDec 
         Caption         =   "D&ecrypt"
         Height          =   510
         Left            =   -73380
         Picture         =   "test.frx":3D0F
         Style           =   1  'Graphical
         TabIndex        =   40
         Top             =   2970
         Width           =   915
      End
      Begin VB.CommandButton cmFonts 
         Caption         =   "Add &Fonts"
         Height          =   555
         Left            =   -74910
         Picture         =   "test.frx":3E59
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   810
         Width           =   1005
      End
      Begin VB.ComboBox cbFonts 
         Height          =   315
         Index           =   1
         Left            =   -74910
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   37
         Top             =   450
         Width           =   2445
      End
      Begin VB.CommandButton cmChReg 
         Caption         =   "C&hange"
         Height          =   330
         Left            =   -72210
         TabIndex        =   36
         Top             =   3195
         Width           =   825
      End
      Begin MSComctlLib.ListView lvVal 
         Height          =   3120
         Left            =   -71355
         TabIndex        =   35
         Top             =   405
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   5503
         View            =   3
         Arrange         =   2
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         Icons           =   "imMenu"
         SmallIcons      =   "imlProps"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Values"
            Object.Width           =   4410
         EndProperty
      End
      Begin VB.TextBox txValues 
         Height          =   330
         Left            =   -74910
         TabIndex        =   34
         Top             =   3195
         Width           =   2670
      End
      Begin VB.CommandButton cmINT 
         Caption         =   "ParseInt"
         Height          =   565
         Left            =   -69660
         OLEDropMode     =   1  'Manual
         Picture         =   "test.frx":3FA3
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   2340
         Width           =   1140
      End
      Begin VB.CommandButton cmRep 
         Caption         =   "Repeat first"
         Height          =   565
         Left            =   -70830
         Picture         =   "test.frx":40ED
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   2340
         Width           =   1140
      End
      Begin VB.CommandButton cmRev 
         Caption         =   "Reverse"
         Height          =   565
         Left            =   -72000
         Picture         =   "test.frx":4237
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   2340
         UseMaskColor    =   -1  'True
         Width           =   1140
      End
      Begin VB.CommandButton cmInitCap 
         Caption         =   "Initial Caps"
         Height          =   565
         Left            =   -73125
         Picture         =   "test.frx":4381
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   2340
         Width           =   1095
      End
      Begin VB.CommandButton cmBin 
         Caption         =   "Convert Binary"
         Height          =   565
         Left            =   -74430
         Picture         =   "test.frx":44CB
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   2340
         Width           =   1275
      End
      Begin VB.CommandButton cmCount 
         Caption         =   "Letter count"
         Height          =   565
         Left            =   -69240
         Picture         =   "test.frx":4615
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   2940
         Width           =   1050
      End
      Begin VB.CommandButton cmLines 
         Caption         =   "Line count"
         Height          =   565
         Left            =   -73605
         Picture         =   "test.frx":475F
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   2940
         Width           =   1095
      End
      Begin VB.CommandButton cmWords 
         Caption         =   "Word count"
         Height          =   565
         Left            =   -74820
         Picture         =   "test.frx":48A9
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   2940
         Width           =   1185
      End
      Begin VB.CommandButton cmPathname 
         Caption         =   "Path name"
         Height          =   565
         Left            =   -71310
         Picture         =   "test.frx":49F3
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   2940
         Width           =   1050
      End
      Begin VB.CommandButton cmFilename 
         Caption         =   "File name"
         Height          =   565
         Left            =   -70230
         Picture         =   "test.frx":4B3D
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   2940
         Width           =   960
      End
      Begin VB.CommandButton cmUp1Level 
         Caption         =   "Up one level"
         Height          =   565
         Left            =   -72480
         Picture         =   "test.frx":4C87
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   2940
         Width           =   1140
      End
      Begin VB.TextBox txOUT 
         Height          =   1755
         Left            =   -71445
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   19
         Top             =   540
         Width           =   3255
      End
      Begin VB.TextBox txIN 
         Height          =   1755
         Left            =   -74775
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   18
         Text            =   "test.frx":4DD1
         Top             =   540
         Width           =   3300
      End
      Begin VB.CommandButton cmCNF 
         Caption         =   "Show Co&nfirm Dialog"
         Height          =   500
         Left            =   -72600
         TabIndex        =   17
         Top             =   2265
         Width           =   2225
      End
      Begin VB.CommandButton cmPT 
         Caption         =   "Show Pro&mpt Dialog"
         Height          =   500
         Left            =   -72600
         TabIndex        =   16
         Top             =   1725
         Width           =   2225
      End
      Begin VB.CommandButton cmMsg 
         Caption         =   "Show M&essage Box"
         Height          =   500
         Left            =   -72600
         TabIndex        =   15
         Top             =   1185
         Width           =   2225
      End
      Begin VB.Frame Frame1 
         Caption         =   "Colours"
         Height          =   1095
         Left            =   -74910
         TabIndex        =   6
         Top             =   990
         Width           =   3570
         Begin VB.CommandButton cmMenuColor 
            Caption         =   "Apply MenuColor"
            Height          =   330
            Left            =   1845
            TabIndex        =   11
            Top             =   630
            Width           =   1635
         End
         Begin VB.CommandButton cmSelColor 
            Caption         =   "Apply SelectColor"
            Height          =   330
            Left            =   1845
            TabIndex        =   10
            Top             =   270
            Width           =   1635
         End
         Begin VB.OptionButton opRGB 
            Caption         =   "RGB"
            Height          =   240
            Left            =   900
            TabIndex        =   9
            Top             =   675
            Value           =   -1  'True
            Width           =   645
         End
         Begin VB.OptionButton opHex 
            Caption         =   "Hex"
            Height          =   240
            Left            =   270
            TabIndex        =   8
            Top             =   675
            Width           =   645
         End
         Begin VB.TextBox txColor 
            Alignment       =   2  'Center
            Height          =   315
            Left            =   135
            TabIndex        =   7
            Text            =   "000000000"
            Top             =   270
            Width           =   1635
         End
      End
      Begin MSComctlLib.ImageCombo imProps 
         Height          =   330
         Left            =   -74880
         TabIndex        =   4
         Top             =   3150
         Width           =   2265
         _ExtentX        =   3995
         _ExtentY        =   582
         _Version        =   393216
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         MousePointer    =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "test.frx":4E14
         Locked          =   -1  'True
         Text            =   "System Properties"
         ImageList       =   "imlProps"
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Checks in 3D"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   -73605
         TabIndex        =   3
         Top             =   720
         Value           =   1  'Checked
         Width           =   1230
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Full Selection"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   -74865
         TabIndex        =   2
         Top             =   720
         Value           =   1  'Checked
         Width           =   1230
      End
      Begin VB.PictureBox pDraw 
         AutoRedraw      =   -1  'True
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2130
         Left            =   90
         ScaleHeight     =   2070
         ScaleWidth      =   6705
         TabIndex        =   1
         Top             =   360
         Width           =   6765
      End
      Begin VB.Frame Frame2 
         Caption         =   "Fonts"
         Height          =   1095
         Left            =   -71265
         TabIndex        =   12
         Top             =   990
         Width           =   2940
         Begin VB.ComboBox cbFonts 
            Height          =   315
            Index           =   0
            Left            =   135
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   13
            Top             =   450
            Width           =   2670
         End
      End
      Begin VB.Label Label9 
         Caption         =   "URL:"
         Height          =   195
         Left            =   -74820
         TabIndex        =   93
         Top             =   2565
         Width           =   555
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "(Bytes)"
         Height          =   195
         Left            =   -73830
         TabIndex        =   88
         Top             =   855
         Width           =   480
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Data:"
         Height          =   240
         Left            =   -74820
         TabIndex        =   86
         Top             =   855
         Width           =   390
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Host:"
         Height          =   195
         Left            =   -74820
         TabIndex        =   85
         Top             =   495
         Width           =   375
      End
      Begin VB.Label lbDrag 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Click the left mouse button here; hold and drag. Alternatively the right mouse button can be used for hold-free drag operations."
         Height          =   1290
         Left            =   -71265
         MousePointer    =   15  'Size All
         TabIndex        =   74
         Top             =   2115
         Width           =   1740
      End
      Begin VB.Label Label5 
         Caption         =   "0"
         Height          =   195
         Left            =   -74730
         TabIndex        =   57
         Top             =   3150
         Width           =   330
      End
      Begin VB.Label lbLen 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Height          =   240
         Left            =   -68250
         TabIndex        =   58
         Top             =   3150
         Width           =   45
      End
      Begin VB.Label Label4 
         Caption         =   "Key:"
         Height          =   195
         Left            =   -74910
         TabIndex        =   44
         Top             =   2940
         Width           =   465
      End
      Begin VB.Label lbFont 
         Alignment       =   2  'Center
         Caption         =   "ABCD"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   -73875
         TabIndex        =   39
         Top             =   810
         Width           =   1410
      End
      Begin VB.Label Label2 
         Caption         =   "Output:"
         Height          =   195
         Index           =   1
         Left            =   -71400
         TabIndex        =   21
         Top             =   360
         Width           =   960
      End
      Begin VB.Label Label2 
         Caption         =   "Input:"
         Height          =   195
         Index           =   0
         Left            =   -74775
         TabIndex        =   20
         Top             =   345
         Width           =   960
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   $"test.frx":512E
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1005
         Left            =   -74910
         TabIndex        =   14
         Top             =   2160
         Width           =   6765
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Properties:"
         Height          =   195
         Left            =   -74835
         TabIndex        =   5
         Top             =   2925
         Width           =   750
      End
   End
   Begin VB.Menu mnuCaptions 
      Caption         =   "&CoolMenu"
      Begin VB.Menu mnuItem 
         Caption         =   "|Description for 1 goes here|Caption1"
         Index           =   0
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuItem 
         Caption         =   "|Description for 2 goes here|Caption2"
         Index           =   1
         Shortcut        =   ^B
      End
      Begin VB.Menu mnuItem 
         Caption         =   "|Description for 3 goes here|Caption3"
         Index           =   2
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuCh 
         Caption         =   "-CHECKBOXES"
      End
      Begin VB.Menu mnuCheck 
         Caption         =   "#|CheckBox|CheckBox"
         Checked         =   -1  'True
         Shortcut        =   ^D
      End
      Begin VB.Menu mnuRadio 
         Caption         =   "*|Option button|Option"
         Checked         =   -1  'True
         Shortcut        =   ^E
      End
      Begin VB.Menu mnuRef 
         Caption         =   "-REFRESH VIEW"
      End
      Begin VB.Menu mnuRefresh 
         Caption         =   "|Refresh the view|&Refresh"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuVb 
         Caption         =   "-VB6LIB.DLL"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "|About vb6lib.dll|A&bout..."
         Shortcut        =   ^{F1}
      End
      Begin VB.Menu mnuExit 
         Caption         =   "|Exit the project|Exit"
         Shortcut        =   ^{F4}
      End
   End
   Begin VB.Menu mnuWD 
      Caption         =   ""
   End
   Begin VB.Menu mnuSD 
      Caption         =   ""
   End
   Begin VB.Menu mnuUT 
      Caption         =   ""
   End
   Begin VB.Menu mnuMEM 
      Caption         =   ""
   End
   Begin VB.Menu mnuPro 
      Caption         =   ""
   End
End
Attribute VB_Name = "fMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-------------------------------------------------
'     This project demonstrates the uses of the
'     vb6lib.dll file. This form demo, and the DLL
'     are copyright Sushant Pandurangi, 2000.
'     http://sushantshome.tripod.com/vb
'     sushant@phreaker.net, sushant@bigwig.net
'     For more software visit the website at
'     http://smartappz.tripod.com/download.htm
'-------------------------------------------------
Option Explicit
'These are the object variables
Public Strings As New VB6LIB.Strings
Public Sounds As New VB6LIB.Sounds
Public Internet  As New VB6LIB.Network
Public Graphics As New VB6LIB.Graphics
Public General As New VB6LIB.General
Public Menus As New VB6LIB.Menus
Public Interact As New VB6LIB.Prompts
Public Registry As New VB6LIB.Registry
Public Win32API As New VB6LIB.WinAPI
'Now, this is for the menu help events
'Later we need to set it as a new HelpObj
'Place code in the MenuHelper_MenuHelp event
Public WithEvents MenuHelper As VB6LIB.HelpObj
Attribute MenuHelper.VB_VarHelpID = -1
'This is the variable that denotes if or not the about
'box has been shown. If yes, there is no need to show
'it again.
Public bShown As Boolean


Private Sub cbFonts_Click(Index As Integer)
'Change the menu font to the one selected now
If Index = 0 Then Menus.MenuFont Me, cbFonts(0).List(cbFonts(0).ListIndex)
If Index = 1 Then lbFont.Font.Bold = False: lbFont.Font.Name = cbFonts(1).List(cbFonts(1).ListIndex)
End Sub

Private Sub Check1_Click()
'Specify whether or not the full menu
'item's region should be selected
Menus.SelectFull Me, CBool(Check1.Value)
End Sub

Private Sub Check2_Click()
'Specify if or not check symbols should
'appear in 3D like checkbox controls
Menus.Check3D Me, CBool(Check2.Value)
End Sub

Private Sub cmChReg_Click()
If Interact.Confirm("Are you sure you want to change the value for this key?", "Change", False) = "Yes" Then
Registry.SetValue HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\" & lvKeys.SelectedItem.Text, lvVal.SelectedItem.Text, txValues.Text
End If
End Sub

Private Sub cmCNF_Click()
Interact.Message "You said " & Interact.Confirm("Say yes or no.", "Title", True), "Title"
End Sub

Private Sub cmConn_Click()
Internet.NetConnect
End Sub

Private Sub cmCopy_Click()
Win32API.CopyFile Replace(Fl1.Path & Fl1.FileName, "\\", "\"), Dt2.Path
End Sub

Private Sub cmCount_Click()
On Error Resume Next
Dim pC As String
pC = Interact.Prompt("Enter character to search for.", "x", "Char")
If pC = "" Then Exit Sub
txOUT.Text = Strings.StrCount(txIN.Text, Left(pC, 1))
End Sub

Private Sub cmDec_Click()
txEnc.Text = General.Decrypt(txEnc.Text, txKey.Text)
End Sub

Private Sub cmDel_Click()
Win32API.DeleteFile Replace(Fl1.Path & Fl1.FileName, "\\", "\")
End Sub

Private Sub cmDiscon_Click()
Internet.Disconnect
End Sub

Private Sub cmEnc_Click()
txEnc.Text = General.Encrypt(txEnc.Text, txKey.Text)
End Sub

Private Sub cmFilename_Click()
On Error Resume Next
txOUT.Text = Strings.GetFile(txIN.Text)
End Sub

Private Sub cmFonts_Click()
cbFonts(1).Clear
General.ListFonts cbFonts(1)
cbFonts(1).ListIndex = 0
cmFonts.Caption = Screen.FontCount & " &Font(s)"
End Sub

Private Sub cmHost_Click()
Dim hOst As String
hOst = Internet.GetName(txHost.Text)
Interact.Message hOst
End Sub

Private Sub cmInitCap_Click()
On Error Resume Next
txOUT.Text = Strings.InitCap(txIN.Text)
End Sub

Private Sub cmINT_Click()
On Error Resume Next
txOUT.Text = Strings.ParseInt(txIN.Text)
End Sub

Private Sub cmIP_Click()
Dim ret As String
ret = Internet.GetAddress(txHost.Text)
Interact.Message ret
End Sub

Private Sub cmIPLong_Click()
Dim ret As String
ret = Internet.GetAddress(txHost.Text)
Interact.Message Internet.GetLongIP(ret)
End Sub

Private Sub cmLines_Click()
On Error Resume Next
txOUT.Text = Strings.LnCount(txIN)
End Sub

Private Sub cmLogoff_Click()
Win32API.ShutDown LOGOFF
End Sub

Private Sub cmMail_Click()
Internet.SendEMail txURL.Text

End Sub

Private Sub cmMenuColor_Click()
On Error GoTo hell
If Len(txColor.Text) <> 9 Then Err.Raise 13
Dim sColour As Long
If opHex.Value = True Then
sColour = CLng(txColor.Text)
Else
sColour = RGB(Left(txColor.Text, 3), Mid(txColor.Text, 4, 3), Right(txColor.Text, 3))
End If
Menus.MenuColor Me, sColour
Exit Sub
hell:
MsgBox Error
End Sub

Private Sub cmMove_Click()
Win32API.MoveFile Replace(Fl1.Path & Fl1.FileName, "\\", "\"), Dt2.Path
End Sub

Private Sub cmNoTop_Click()
Win32API.FormTop TOPMOST_FALSE, Me, Me.Left, Me.Top, Me.Width, Me.Height
End Sub

Private Sub cmOpen_Click()
On Error Resume Next
txOpen.Text = General.FileOpen(Interact.Prompt("Please enter the filename to open.", "", "Open"))
End Sub

Private Sub cmOpenURL_Click()
Internet.OpenPage txURL.Text, SHOW_DEFAULT
End Sub

Private Sub cmPaste_Click()
Win32API.PasteText txCopy
End Sub

Private Sub cmPathname_Click()
On Error Resume Next
txOUT.Text = Strings.GetPath(txIN.Text)
End Sub

Private Sub cmPause_Click()
Sounds.MP3Pause
tmMP3.Enabled = False
End Sub

Private Sub cmPing_Click()
Dim strTime As String, bData As Boolean
Internet.SendPing txHost.Text, strTime, bData, txData.Text, 500
If bData = True Then
Interact.Message "The host '" & txHost.Text & "' was pinged in " & strTime, "Ping"
Else
Interact.Message "The host '" & txHost.Text & "' could not be pinged."
End If
End Sub

Private Sub cmPlay_Click()
On Error Resume Next
If opWAV.Value = False Then
Sounds.MP3Play Me.hWnd, Replace(F1.Path & "\" & F1.FileName, "\\", "\")
slSeek.Max = Sounds.LenInSec
lbLen.Caption = slSeek.Max
tmMP3.Enabled = True
pBar.Max = Sounds.LenInSec
Win32API.Progress True, Me, pBar, sBar, 1
pBar.Visible = True
Else
Sounds.PlayWave Replace(F1.Path & "\" & F1.FileName, "\\", "\")
End If
End Sub

Private Sub cmProgress_Click()
Win32API.Progress True, Me, pBar, sBar, 1
pBar.Value = 0
While pBar.Value < pBar.Max
pBar.Value = pBar.Value + 0.5
Wend
End Sub

Private Sub cmPT_Click()
Interact.Message "You said " & Interact.Prompt("Say Something.", "Something", "Title", False), "Title"
End Sub

Private Sub cmRep_Click()
On Error Resume Next
txOUT.Text = Strings.Repeat(Left(txIN.Text, 1), CLng(Len(txIN.Text) + 25))
End Sub

Private Sub cmRestart_Click()
Win32API.ShutDown REBOOT
End Sub

Private Sub cmRev_Click()
On Error Resume Next
txOUT.Text = Strings.Reverse(txIN.Text)
End Sub

Private Sub cmRun_Click()
Win32API.RunDialog Me.hWnd, "RunDialog() - vb6lib.dll", "Put in a message here." & vbNewLine & "(C) Sushant Pandurangi, 2000-2001"
End Sub

Private Sub cmSave_Click()
On Error Resume Next
General.FileSave Interact.Prompt("Please enter the filename to save.", "", "Save"), txOpen.Text
End Sub

Private Sub cmSelColor_Click()
On Error GoTo hell
If Len(txColor.Text) <> 9 Then Err.Raise 13
Dim sColour As Long
If opHex.Value = True Then
sColour = CLng(txColor.Text)
Else
sColour = RGB(Left(txColor.Text, 3), Mid(txColor.Text, 4, 3), Right(txColor.Text, 3))
End If
Menus.SelColor Me, sColour
Exit Sub
hell:
MsgBox Error
End Sub

Private Sub cmMsg_Click()
Interact.Message "This is a message", "This is the title"
End Sub

Private Sub cmShut_Click()
Win32API.ShutDown BYEBYE
End Sub

Private Sub cmStop_Click()
Sounds.MP3Stop
tmMP3.Enabled = False
pBar.Visible = False
End Sub

Private Sub cmUndo_Click()
Win32API.UndoEdit txCopy
End Sub

Private Sub cmUp1Level_Click()
On Error Resume Next
txOUT.Text = Strings.Up1Level(txIN.Text)
End Sub

Private Sub cmWords_Click()
On Error Resume Next
txOUT.Text = Strings.WdCount(txIN.Text)
End Sub

Private Sub CmBin_Click()
On Error Resume Next
txOUT.Text = Strings.CBinary(txIN.Text)
End Sub

Private Sub CmCopyT_Click()
Win32API.CopyText txCopy
End Sub

Private Sub CmFind_Click()
Win32API.FindDialog
End Sub

Private Sub CmOnTop_Click()
Win32API.FormTop TOPMOST_TRUE, Me, Me.Left, Me.Top, Me.Width, Me.Height
End Sub

Private Sub CmStandby_Click()
Win32API.StandBy 5000
End Sub


Private Sub D1_Change()
On Error Resume Next
D2.Path = D1.Drive
End Sub

Private Sub D2_Change()
F1.Path = D2.Path
End Sub

Private Sub Dr1_Change()
On Error Resume Next
Dr2.Path = Dr1.Drive
End Sub

Private Sub Dr2_Change()
On Error Resume Next
Fl1.Path = Dr2.Path
End Sub

Private Sub Dt1_Change()
On Error Resume Next
Dt2.Path = Dt1.Drive
End Sub

Private Sub F1_Click()
cmPlay.Default = True
End Sub

Private Sub Form_Load()
On Error Resume Next
Dim i As Integer, pControl As Control
'Set MenuHelper to a new HelpObj Class
'Now we can use it to display help
    Set MenuHelper = New VB6LIB.HelpObj
'Now add the fonts to the combo box
'The "General" Module contains this Sub
    cbFonts(0).Clear
    General.ListFonts cbFonts(0)
    cbFonts(0).ListIndex = 0
'Initialise Menus. imMenu is the imagelist
'containing the images. MenuHelper is
'the help object. the imagelist should contain
'the images that you want to add to your menus
'in MDI environment both MDI mother and child
'forms use the same imagelist. In the Tag
'property of each image, type the caption
'of the menu item you want to put the image
'on. in MDIChilds, put this line in the load event:
'Menus.MDIChild (Me)
    Menus.Activate Me.hWnd, imMenu, MenuHelper, True
'Apply custom fonts and sizes
    Menus.MenuFont Me, "Tahoma"
    Menus.MenuSize Me, 8
    Menus.SelColor Me, vbBlack
    Menus.MenuColor Me, vbBlack
    Menus.SelectFull Me, True
'Done with applying menus
'Now add properties items in the Win32API section
    AddProperties
'Registry stuff: add keys in some HKEY to ListView
Registry.EnumKeys HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion"
'Clear if anything exists
lvKeys.ListItems.Clear
'go in loop till all are added. ListKeys is in the registry module
'it contains the keys when the EnumKeys() sub is called.
For i = 0 To Registry.ListKeys.Count - 1
lvKeys.ListItems.Add i, "Item" & i, Registry.ListKeys.Item(i), 1, 2
'Remove it from ListKeys; we need to clear it
Registry.ListKeys.Remove i
Next i
'Adjust width
lvKeys.ColumnHeaders(1).Width = lvKeys.Width - 85
'Do the same with values in the key
'Here, its ListValues
Registry.EnumValues HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\" & lvKeys.SelectedItem.Text
lvVal.ListItems.Clear
For i = 0 To Registry.ListValues.Count - 1
lvVal.ListItems.Add i, , Registry.ListValues.Item(i), 1, 2
Registry.ListValues.Remove i
Next i
lvVal.ColumnHeaders(1).Width = lvVal.Width - 310
'Set fonts for all except the SSTab
For Each pControl In Me.Controls
pControl.Font.Name = "Tahoma"
Next pControl
SSTab1.Font.Name = "MS Sans Serif"
'Set progressbar value to max
pBar.Value = pBar.Max
'Aboutbox has not been shown
bShown = False
'System stuff
Dim rT As Long, rA As Long
'Memory; from API module
GetMemory rT, rA
mnuMEM.Caption = (rA / 100) & "/" & (rT / 100)
mnuPro.Caption = Processor
mnuWD.Caption = LCase(WindowsDir)
mnuSD.Caption = LCase(SystemDir)
mnuUT.Caption = UsedTime ' Time  that system is in use from
'Now draw the cool gradients. This feature is only
'in the updated DLL source code as of 25/3.
For i = 0 To pGrade.Count - 1
GradeForm pGrade(i), i, , 700
Next i
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Sounds.IsPlaying = True Then Sounds.MP3Stop
End Sub

Private Sub imProps_Click()
On Error Resume Next
Win32API.Properties imProps.SelectedItem.Index
D1.SetFocus
End Sub

Private Sub lbDrag_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Win32API.EasyMove Me
End Sub

Private Sub lvKeys_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error Resume Next
Dim i As Integer
sBar.Panels(1).Text = Item.Text
txValues.Text = Registry.GetValue(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\" & Item.Text, lvVal.SelectedItem.Text, "")
For i = 0 To Registry.ListValues.Count - 1
Registry.ListValues.Remove i
Next i
    Registry.EnumValues HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion" & "\" & lvKeys.SelectedItem.Text
    lvVal.ListItems.Clear
    For i = 0 To Registry.ListValues.Count - 1
    lvVal.ListItems.Add i, , Registry.ListValues.Item(i), 1, 2
    Registry.ListValues.Remove i
    Next i
End Sub

Private Sub lvVal_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error Resume Next
sBar.Panels(1).Text = Item.Text
txValues.Text = Registry.GetValue(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\" & lvKeys.SelectedItem.Text, Item.Text, "(Default)")
End Sub

Private Sub mnuAbout_Click()
bShown = True
General.AboutBox
End Sub

Private Sub mnuCheck_Click()
    mnuCheck.Checked = Not mnuCheck.Checked
End Sub

Private Sub mnuExit_Click()
    If bShown = False Then General.AboutBox
    Unload Me
    Set MenuHelper = Nothing
'Dont try putting in the END statement here
'using it will cause a nice crash on your system
End Sub

Private Sub mnuItem_Click(Index As Integer)
General.AboutBox
bShown = True
End Sub

Private Sub mnuRadio_Click()
    mnuRadio.Checked = Not mnuRadio.Checked
End Sub

Private Sub mnuRefresh_Click()
Me.Refresh
Form_Load
End Sub

Private Sub opMP3_Click()
F1.Pattern = "*.mp3"
cmStop.Visible = True
cmPause.Visible = True
slSeek.Visible = True
Label5.Visible = True
lbLen.Visible = True
End Sub

Private Sub opWAV_Click()
F1.Pattern = "*.wav"
cmPause.Visible = False
cmStop.Visible = False
slSeek.Visible = False
Label5.Visible = False
lbLen.Visible = False
End Sub

Private Sub slSeek_Click()
If Sounds.IsPlaying = True Then Sounds.MP3Seek slSeek.Value / 1000
End Sub

Private Sub Form_Resize()
On Error Resume Next
    SSTab1.Width = ScaleWidth
    SSTab1.Height = sBar.Top - 45
End Sub

Private Sub MenuHelper_MenuHelp(ByVal MenuText As String, ByVal MenuHelp As String, ByVal Enabled As Boolean)
'----------------------------------------------
'The Menu-related Procedures were
'taken from Oliver Martin's work.
'martin.oliver@bigfoot.com
'His project 'CoolMenu' can be found on the
'several websites on the internet which provide
'free VB source code.
'----------------------------------------------
'Now this is when the menu item is hovered on
'We need to show a description of the item so
'The description should be set in its caption.
'See the menus without running the project.
    If Enabled = True Then
    sBar.Panels(1).Text = MenuHelp
    Else
    sBar.Panels(1).Text = "Press Esc to close menu."
    End If
'Thats all we needed here
End Sub

Private Sub pDraw_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'You have to see the graphics stuff for yourself to believe it
    If Button = 1 Then Graphics.Draw3DLine pDraw, X, Y, pDraw.ScaleWidth - X
    If Button = 2 Then Graphics.Draw3DText pDraw, "3D Text", RGB(100, 100, 100)
    If Shift = 1 Then pDraw.Cls
    If Shift = 4 Then pDraw.Cls: Graphics.ProgressBar pDraw, 23, False
    If Shift = 2 Then Graphics.GradeForm pDraw, 2
End Sub

Private Sub pDraw_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    sBar.Panels(1).Text = "Click or right click; Shift, Alt and Ctrl for more."
End Sub

Sub AddProperties()
With imProps.ComboItems
.Clear
.Add 1, , "System Properties", 1
.Add 2, , "Internet Options", 1
.Add 3, , "Modems settings", 1
.Add 4, , "Add/Remove Programs", 1
.Add 5, , "Add New Hardware", 1
.Add 6, , "Sounds Settings", 1
.Add 7, , "Network Options", 1
.Add 8, , "Mouse settings", 1
.Add 9, , "Keyboard settings", 1
.Add 10, , "Time/Date settings", 1
.Add 11, , "Regional Settings", 1
.Add 12, , "Password settings", 1
.Add 13, , "Display Settings", 1
End With
End Sub

Private Sub SSTab1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
sBar.Panels(1).Text = SSTab1.TabCaption(SSTab1.Tab)
End Sub

Private Sub tCursor_Timer()
sBar.Panels(5).Text = "" & CursorPos(0, 0, vbPixels)
sBar.Panels(4).Text = "" & CursorPos(0, 0, vbTwips)
mnuUT.Caption = UsedTime
End Sub

Private Sub tmMP3_Timer()
On Error Resume Next
lbLen.Caption = Sounds.PosInSec & "/" & Sounds.LenInSec & " (seconds)"
sBar.Panels(3).Text = Sounds.Position & "/" & Sounds.FileLength
slSeek.Value = Sounds.PosInSec
pBar.Value = Sounds.PosInSec
End Sub
