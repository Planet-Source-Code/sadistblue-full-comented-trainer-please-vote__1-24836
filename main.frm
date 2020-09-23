VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form main 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Tenchi Muyo RPG Editor 2.0"
   ClientHeight    =   4350
   ClientLeft      =   3300
   ClientTop       =   2955
   ClientWidth     =   4410
   Icon            =   "main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4350
   ScaleWidth      =   4410
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check1 
      Caption         =   "BU"
      Height          =   255
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   114
      ToolTipText     =   "Create a backup copy of this file."
      Top             =   4080
      Value           =   1  'Checked
      Width           =   375
   End
   Begin VB.CommandButton cmdeasykill 
      Caption         =   "&Easy Kills"
      Height          =   255
      Left            =   3120
      TabIndex        =   113
      ToolTipText     =   "Easy level up, as well as kills"
      Top             =   4080
      Width           =   855
   End
   Begin VB.CommandButton cmdmaxstats 
      Caption         =   "&Max Stats"
      Height          =   255
      Left            =   2280
      TabIndex        =   112
      ToolTipText     =   "Max all the stats"
      Top             =   4080
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Refresh"
      Height          =   255
      Left            =   1560
      TabIndex        =   87
      ToolTipText     =   "Refresh the current file."
      Top             =   4080
      Width           =   735
   End
   Begin VB.CommandButton cmdsave 
      Caption         =   "&Save"
      Height          =   255
      Left            =   840
      TabIndex        =   111
      ToolTipText     =   "Save the current file."
      Top             =   4080
      Width           =   735
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   375
      Left            =   1800
      TabIndex        =   74
      Top             =   4440
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      _Version        =   393217
      TextRTF         =   $"main.frx":1026
   End
   Begin MSComDlg.CommonDialog dlg 
      Left            =   4320
      Top             =   3480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DialogTitle     =   "Open..."
      Filter          =   "ZSNES Saved Stat (*.zs*) | *.zs*"
   End
   Begin VB.CommandButton cmdopen 
      Caption         =   "&Open"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      ToolTipText     =   "Open a ZSNES File for Editing"
      Top             =   4080
      Width           =   735
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4065
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4425
      _ExtentX        =   7805
      _ExtentY        =   7170
      _Version        =   393216
      Tabs            =   7
      TabsPerRow      =   7
      TabHeight       =   520
      TabCaption(0)   =   "P1"
      TabPicture(0)   =   "main.frx":1107
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2(0)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label3(0)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label4(0)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label5(0)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label6(0)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label7(0)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label8(0)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label9(0)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Image1(0)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label9(4)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "p1mov"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "p1preshp"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "p1kia"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "p1presatk"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "p1maxatk"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "p1kill"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "p1maxhp"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "p1presdef"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "p1maxdef"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "p1lv"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).ControlCount=   21
      TabCaption(1)   =   "P2"
      TabPicture(1)   =   "main.frx":1123
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "p2lv"
      Tab(1).Control(1)=   "p2maxdef"
      Tab(1).Control(2)=   "p2presdef"
      Tab(1).Control(3)=   "p2maxatk"
      Tab(1).Control(4)=   "p2presatk"
      Tab(1).Control(5)=   "p2kia"
      Tab(1).Control(6)=   "p2maxhp"
      Tab(1).Control(7)=   "p2preshp"
      Tab(1).Control(8)=   "p2kill"
      Tab(1).Control(9)=   "p2mov"
      Tab(1).Control(10)=   "Label9(5)"
      Tab(1).Control(11)=   "Image1(1)"
      Tab(1).Control(12)=   "Label1(1)"
      Tab(1).Control(13)=   "Label2(1)"
      Tab(1).Control(14)=   "Label3(1)"
      Tab(1).Control(15)=   "Label4(1)"
      Tab(1).Control(16)=   "Label5(1)"
      Tab(1).Control(17)=   "Label6(1)"
      Tab(1).Control(18)=   "Label7(1)"
      Tab(1).Control(19)=   "Label8(1)"
      Tab(1).Control(20)=   "Label9(1)"
      Tab(1).ControlCount=   21
      TabCaption(2)   =   "P3"
      TabPicture(2)   =   "main.frx":113F
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label1(2)"
      Tab(2).Control(1)=   "Label2(2)"
      Tab(2).Control(2)=   "Label3(2)"
      Tab(2).Control(3)=   "Label4(2)"
      Tab(2).Control(4)=   "Label5(2)"
      Tab(2).Control(5)=   "Label6(2)"
      Tab(2).Control(6)=   "Label7(2)"
      Tab(2).Control(7)=   "Label8(2)"
      Tab(2).Control(8)=   "Label9(2)"
      Tab(2).Control(9)=   "Image1(2)"
      Tab(2).Control(10)=   "Label9(6)"
      Tab(2).Control(11)=   "p3maxdef"
      Tab(2).Control(12)=   "p3presdef"
      Tab(2).Control(13)=   "p3maxatk"
      Tab(2).Control(14)=   "p3presatk"
      Tab(2).Control(15)=   "p3kia"
      Tab(2).Control(16)=   "p3maxhp"
      Tab(2).Control(17)=   "p3preshp"
      Tab(2).Control(18)=   "p3kill"
      Tab(2).Control(19)=   "p3mov"
      Tab(2).Control(20)=   "p3lv"
      Tab(2).ControlCount=   21
      TabCaption(3)   =   "P4"
      TabPicture(3)   =   "main.frx":115B
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Label9(3)"
      Tab(3).Control(1)=   "Label8(3)"
      Tab(3).Control(2)=   "Label7(3)"
      Tab(3).Control(3)=   "Label6(3)"
      Tab(3).Control(4)=   "Label5(3)"
      Tab(3).Control(5)=   "Label4(3)"
      Tab(3).Control(6)=   "Label3(3)"
      Tab(3).Control(7)=   "Label2(3)"
      Tab(3).Control(8)=   "Label1(3)"
      Tab(3).Control(9)=   "Image1(3)"
      Tab(3).Control(10)=   "Label9(7)"
      Tab(3).Control(11)=   "p4maxdef"
      Tab(3).Control(12)=   "p4presdef"
      Tab(3).Control(13)=   "p4maxatk"
      Tab(3).Control(14)=   "p4presatk"
      Tab(3).Control(15)=   "p4kia"
      Tab(3).Control(16)=   "p4maxhp"
      Tab(3).Control(17)=   "p4preshp"
      Tab(3).Control(18)=   "p4kill"
      Tab(3).Control(19)=   "p4mov"
      Tab(3).Control(20)=   "p4lv"
      Tab(3).ControlCount=   21
      TabCaption(4)   =   "Enemy"
      TabPicture(4)   =   "main.frx":1177
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Image1(4)"
      Tab(4).Control(1)=   "Label3(4)"
      Tab(4).Control(2)=   "Label3(5)"
      Tab(4).Control(3)=   "Label3(6)"
      Tab(4).Control(4)=   "Label3(7)"
      Tab(4).Control(5)=   "enm1alive"
      Tab(4).Control(6)=   "enm2alive"
      Tab(4).Control(7)=   "enm3alive"
      Tab(4).Control(8)=   "enm4alive"
      Tab(4).Control(9)=   "enm1preshp"
      Tab(4).Control(10)=   "enm2preshp"
      Tab(4).Control(11)=   "enm3preshp"
      Tab(4).Control(12)=   "enm4preshp"
      Tab(4).ControlCount=   13
      TabCaption(5)   =   "Team"
      TabPicture(5)   =   "main.frx":1193
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Image1(5)"
      Tab(5).Control(1)=   "havep1"
      Tab(5).Control(2)=   "havep3"
      Tab(5).Control(3)=   "havep4"
      Tab(5).Control(4)=   "havep5"
      Tab(5).Control(5)=   "havep7"
      Tab(5).Control(6)=   "havep8"
      Tab(5).Control(7)=   "havep6"
      Tab(5).Control(8)=   "havep2"
      Tab(5).Control(9)=   "havep9"
      Tab(5).Control(10)=   "havep11"
      Tab(5).Control(11)=   "havep13"
      Tab(5).Control(12)=   "havep10"
      Tab(5).Control(13)=   "havep12"
      Tab(5).ControlCount=   14
      TabCaption(6)   =   "About"
      TabPicture(6)   =   "main.frx":11AF
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "Image2"
      Tab(6).Control(1)=   "Label11"
      Tab(6).Control(2)=   "Label12"
      Tab(6).ControlCount=   3
      Begin VB.TextBox p4lv 
         Height          =   285
         Left            =   -71400
         MaxLength       =   2
         TabIndex        =   108
         Top             =   3720
         Width           =   615
      End
      Begin VB.TextBox p3lv 
         Height          =   285
         Left            =   -71400
         MaxLength       =   2
         TabIndex        =   107
         Top             =   3720
         Width           =   615
      End
      Begin VB.TextBox p2lv 
         Height          =   285
         Left            =   -71400
         MaxLength       =   2
         TabIndex        =   106
         Top             =   3720
         Width           =   615
      End
      Begin VB.TextBox p1lv 
         Height          =   285
         Left            =   3600
         MaxLength       =   2
         TabIndex        =   105
         Top             =   3720
         Width           =   615
      End
      Begin VB.CheckBox havep12 
         Caption         =   "Have Mizuki"
         Height          =   255
         Left            =   -72480
         TabIndex        =   100
         Top             =   2280
         Width           =   1455
      End
      Begin VB.CheckBox havep10 
         Caption         =   "Have Kamidake"
         Height          =   255
         Left            =   -72480
         TabIndex        =   99
         Top             =   1920
         Width           =   1455
      End
      Begin VB.CheckBox havep13 
         Caption         =   "Have Kusumi"
         Height          =   255
         Left            =   -74400
         TabIndex        =   98
         Top             =   2640
         Width           =   1455
      End
      Begin VB.CheckBox havep11 
         Caption         =   "Have Yokuinojou"
         Height          =   255
         Left            =   -74400
         TabIndex        =   97
         Top             =   2280
         Width           =   1575
      End
      Begin VB.CheckBox havep9 
         Caption         =   "Have Azaka"
         Height          =   255
         Left            =   -74400
         TabIndex        =   96
         Top             =   1920
         Width           =   1455
      End
      Begin VB.CheckBox havep2 
         Caption         =   "Have Ayeka"
         Height          =   255
         Left            =   -72480
         TabIndex        =   95
         Top             =   480
         Width           =   1455
      End
      Begin VB.CheckBox havep6 
         Caption         =   "Have Sasami"
         Height          =   255
         Left            =   -72480
         TabIndex        =   94
         Top             =   1200
         Width           =   1455
      End
      Begin VB.CheckBox havep8 
         Caption         =   "Have Washu"
         Height          =   255
         Left            =   -72480
         TabIndex        =   93
         Top             =   1560
         Width           =   1455
      End
      Begin VB.CheckBox havep7 
         Caption         =   "Have Katsuhito"
         Height          =   255
         Left            =   -74400
         TabIndex        =   92
         Top             =   1560
         Width           =   1455
      End
      Begin VB.CheckBox havep5 
         Caption         =   "Have Ryo-ohki"
         Height          =   255
         Left            =   -74400
         TabIndex        =   91
         Top             =   1200
         Width           =   1455
      End
      Begin VB.CheckBox havep4 
         Caption         =   "Have Mihoshi"
         Height          =   255
         Left            =   -72480
         TabIndex        =   90
         Top             =   840
         Width           =   1455
      End
      Begin VB.CheckBox havep3 
         Caption         =   "Have Ryoko"
         Height          =   255
         Left            =   -74400
         TabIndex        =   89
         Top             =   840
         Width           =   1455
      End
      Begin VB.CheckBox havep1 
         Caption         =   "Have Tenchi"
         Height          =   255
         Left            =   -74400
         TabIndex        =   88
         Top             =   480
         Width           =   1455
      End
      Begin VB.TextBox enm4preshp 
         Height          =   285
         Left            =   -71400
         MaxLength       =   2
         TabIndex        =   86
         Top             =   2280
         Width           =   615
      End
      Begin VB.TextBox enm3preshp 
         Height          =   285
         Left            =   -71400
         MaxLength       =   2
         TabIndex        =   85
         Top             =   1920
         Width           =   615
      End
      Begin VB.TextBox enm2preshp 
         Height          =   285
         Left            =   -71400
         MaxLength       =   2
         TabIndex        =   84
         Top             =   1560
         Width           =   615
      End
      Begin VB.TextBox enm1preshp 
         Height          =   285
         Left            =   -71400
         MaxLength       =   2
         TabIndex        =   83
         Top             =   1200
         Width           =   615
      End
      Begin VB.CheckBox enm4alive 
         Caption         =   "Enemy 4 exists"
         Height          =   255
         Left            =   -72480
         TabIndex        =   78
         Top             =   840
         Width           =   1455
      End
      Begin VB.CheckBox enm3alive 
         Caption         =   "Enemy 3 exists"
         Height          =   255
         Left            =   -74400
         TabIndex        =   77
         Top             =   840
         Width           =   1455
      End
      Begin VB.CheckBox enm2alive 
         Caption         =   "Enemy 2 exists"
         Height          =   255
         Left            =   -72480
         TabIndex        =   76
         Top             =   480
         Width           =   1455
      End
      Begin VB.CheckBox enm1alive 
         Caption         =   "Enemy 1 exists"
         Height          =   255
         Left            =   -74400
         TabIndex        =   75
         Top             =   480
         Width           =   1455
      End
      Begin VB.TextBox p4mov 
         Height          =   285
         Left            =   -71400
         MaxLength       =   2
         TabIndex        =   73
         Top             =   480
         Width           =   615
      End
      Begin VB.TextBox p4kill 
         Height          =   285
         Left            =   -71400
         MaxLength       =   2
         TabIndex        =   72
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox p4preshp 
         Height          =   285
         Left            =   -71400
         MaxLength       =   2
         TabIndex        =   71
         Top             =   1200
         Width           =   615
      End
      Begin VB.TextBox p4maxhp 
         Height          =   285
         Left            =   -71400
         MaxLength       =   2
         TabIndex        =   70
         Top             =   1560
         Width           =   615
      End
      Begin VB.TextBox p4kia 
         Height          =   285
         Left            =   -71400
         MaxLength       =   2
         TabIndex        =   69
         Top             =   1920
         Width           =   615
      End
      Begin VB.TextBox p4presatk 
         Height          =   285
         Left            =   -71400
         MaxLength       =   2
         TabIndex        =   68
         Top             =   2280
         Width           =   615
      End
      Begin VB.TextBox p4maxatk 
         Height          =   285
         Left            =   -71400
         MaxLength       =   2
         TabIndex        =   67
         Top             =   2640
         Width           =   615
      End
      Begin VB.TextBox p4presdef 
         Height          =   285
         Left            =   -71400
         MaxLength       =   2
         TabIndex        =   66
         Top             =   3000
         Width           =   615
      End
      Begin VB.TextBox p4maxdef 
         Height          =   285
         Left            =   -71400
         MaxLength       =   2
         TabIndex        =   65
         Top             =   3360
         Width           =   615
      End
      Begin VB.TextBox p3mov 
         Height          =   285
         Left            =   -71400
         MaxLength       =   2
         TabIndex        =   64
         Top             =   480
         Width           =   615
      End
      Begin VB.TextBox p3kill 
         Height          =   285
         Left            =   -71400
         MaxLength       =   2
         TabIndex        =   63
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox p3preshp 
         Height          =   285
         Left            =   -71400
         MaxLength       =   2
         TabIndex        =   62
         Top             =   1200
         Width           =   615
      End
      Begin VB.TextBox p3maxhp 
         Height          =   285
         Left            =   -71400
         MaxLength       =   2
         TabIndex        =   61
         Top             =   1560
         Width           =   615
      End
      Begin VB.TextBox p3kia 
         Height          =   285
         Left            =   -71400
         MaxLength       =   2
         TabIndex        =   60
         Top             =   1920
         Width           =   615
      End
      Begin VB.TextBox p3presatk 
         Height          =   285
         Left            =   -71400
         MaxLength       =   2
         TabIndex        =   59
         Top             =   2280
         Width           =   615
      End
      Begin VB.TextBox p3maxatk 
         Height          =   285
         Left            =   -71400
         MaxLength       =   2
         TabIndex        =   58
         Top             =   2640
         Width           =   615
      End
      Begin VB.TextBox p3presdef 
         Height          =   285
         Left            =   -71400
         MaxLength       =   2
         TabIndex        =   57
         Top             =   3000
         Width           =   615
      End
      Begin VB.TextBox p3maxdef 
         Height          =   285
         Left            =   -71400
         MaxLength       =   2
         TabIndex        =   56
         Top             =   3360
         Width           =   615
      End
      Begin VB.TextBox p2maxdef 
         Height          =   285
         Left            =   -71400
         MaxLength       =   2
         TabIndex        =   46
         Top             =   3360
         Width           =   615
      End
      Begin VB.TextBox p2presdef 
         Height          =   285
         Left            =   -71400
         MaxLength       =   2
         TabIndex        =   45
         Top             =   3000
         Width           =   615
      End
      Begin VB.TextBox p1maxdef 
         Height          =   285
         Left            =   3600
         MaxLength       =   2
         TabIndex        =   44
         Top             =   3360
         Width           =   615
      End
      Begin VB.TextBox p1presdef 
         Height          =   285
         Left            =   3600
         MaxLength       =   2
         TabIndex        =   43
         Top             =   3000
         Width           =   615
      End
      Begin VB.TextBox p1maxhp 
         Height          =   285
         Left            =   3600
         MaxLength       =   2
         TabIndex        =   42
         Top             =   1560
         Width           =   615
      End
      Begin VB.TextBox p1kill 
         Height          =   285
         Left            =   3600
         MaxLength       =   2
         TabIndex        =   41
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox p2maxatk 
         Height          =   285
         Left            =   -71400
         MaxLength       =   2
         TabIndex        =   40
         Top             =   2640
         Width           =   615
      End
      Begin VB.TextBox p2presatk 
         Height          =   285
         Left            =   -71400
         MaxLength       =   2
         TabIndex        =   39
         Top             =   2280
         Width           =   615
      End
      Begin VB.TextBox p2kia 
         Height          =   285
         Left            =   -71400
         MaxLength       =   2
         TabIndex        =   38
         Top             =   1920
         Width           =   615
      End
      Begin VB.TextBox p2maxhp 
         Height          =   285
         Left            =   -71400
         MaxLength       =   2
         TabIndex        =   37
         Top             =   1560
         Width           =   615
      End
      Begin VB.TextBox p2preshp 
         Height          =   285
         Left            =   -71400
         MaxLength       =   2
         TabIndex        =   36
         Top             =   1200
         Width           =   615
      End
      Begin VB.TextBox p2kill 
         Height          =   285
         Left            =   -71400
         MaxLength       =   2
         TabIndex        =   35
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox p2mov 
         Height          =   285
         Left            =   -71400
         MaxLength       =   2
         TabIndex        =   34
         Top             =   480
         Width           =   615
      End
      Begin VB.TextBox p1maxatk 
         Height          =   285
         Left            =   3600
         MaxLength       =   2
         TabIndex        =   33
         Top             =   2640
         Width           =   615
      End
      Begin VB.TextBox p1presatk 
         Height          =   285
         Left            =   3600
         MaxLength       =   2
         TabIndex        =   32
         Top             =   2280
         Width           =   615
      End
      Begin VB.TextBox p1kia 
         Height          =   285
         Left            =   3600
         MaxLength       =   2
         TabIndex        =   31
         Top             =   1920
         Width           =   615
      End
      Begin VB.TextBox p1preshp 
         Height          =   285
         Left            =   3600
         MaxLength       =   2
         TabIndex        =   30
         Top             =   1200
         Width           =   615
      End
      Begin VB.TextBox p1mov 
         Height          =   285
         Left            =   3600
         MaxLength       =   2
         TabIndex        =   11
         Top             =   480
         Width           =   615
      End
      Begin VB.Label Label12 
         Caption         =   $"main.frx":11CB
         Height          =   960
         Left            =   -74880
         TabIndex        =   110
         Top             =   3000
         Width           =   4215
      End
      Begin VB.Label Label11 
         Caption         =   $"main.frx":12EE
         Height          =   2535
         Left            =   -72480
         TabIndex        =   109
         Top             =   480
         Width           =   1815
      End
      Begin VB.Image Image2 
         BorderStyle     =   1  'Fixed Single
         Height          =   2475
         Left            =   -74880
         Picture         =   "main.frx":141F
         Stretch         =   -1  'True
         Top             =   480
         Width           =   2415
      End
      Begin VB.Label Label9 
         Caption         =   "Level: ------------------------------------------------------"
         Height          =   255
         Index           =   7
         Left            =   -74400
         TabIndex        =   104
         Top             =   3720
         Width           =   3015
      End
      Begin VB.Label Label9 
         Caption         =   "Level: ------------------------------------------------------"
         Height          =   255
         Index           =   6
         Left            =   -74400
         TabIndex        =   103
         Top             =   3720
         Width           =   3015
      End
      Begin VB.Label Label9 
         Caption         =   "Level: ------------------------------------------------------"
         Height          =   255
         Index           =   5
         Left            =   -74400
         TabIndex        =   102
         Top             =   3720
         Width           =   3015
      End
      Begin VB.Label Label9 
         Caption         =   "Level: -----------------------------------------------------"
         Height          =   255
         Index           =   4
         Left            =   600
         TabIndex        =   101
         Top             =   3720
         Width           =   3015
      End
      Begin VB.Image Image1 
         Height          =   3135
         Index           =   5
         Left            =   -74880
         Picture         =   "main.frx":34211
         Stretch         =   -1  'True
         Top             =   480
         Width           =   255
      End
      Begin VB.Label Label3 
         Caption         =   "Enemy 4 HP: --------------------------------------------"
         Height          =   255
         Index           =   7
         Left            =   -74400
         TabIndex        =   82
         Top             =   2280
         Width           =   3015
      End
      Begin VB.Label Label3 
         Caption         =   "Enemy 3 HP: --------------------------------------------"
         Height          =   255
         Index           =   6
         Left            =   -74400
         TabIndex        =   81
         Top             =   1920
         Width           =   3015
      End
      Begin VB.Label Label3 
         Caption         =   "Enemy 2 HP: --------------------------------------------"
         Height          =   255
         Index           =   5
         Left            =   -74400
         TabIndex        =   80
         Top             =   1560
         Width           =   3015
      End
      Begin VB.Label Label3 
         Caption         =   "Enemy 1 HP: --------------------------------------------"
         Height          =   255
         Index           =   4
         Left            =   -74400
         TabIndex        =   79
         Top             =   1200
         Width           =   3015
      End
      Begin VB.Image Image1 
         Height          =   3135
         Index           =   4
         Left            =   -74880
         Picture         =   "main.frx":34381
         Stretch         =   -1  'True
         Top             =   480
         Width           =   255
      End
      Begin VB.Image Image1 
         Height          =   3135
         Index           =   3
         Left            =   -74880
         Picture         =   "main.frx":344F1
         Stretch         =   -1  'True
         Top             =   600
         Width           =   255
      End
      Begin VB.Image Image1 
         Height          =   3135
         Index           =   2
         Left            =   -74880
         Picture         =   "main.frx":34661
         Stretch         =   -1  'True
         Top             =   600
         Width           =   255
      End
      Begin VB.Image Image1 
         Height          =   3135
         Index           =   1
         Left            =   -74880
         Picture         =   "main.frx":347D1
         Stretch         =   -1  'True
         Top             =   600
         Width           =   255
      End
      Begin VB.Image Image1 
         Height          =   3135
         Index           =   0
         Left            =   120
         Picture         =   "main.frx":34941
         Stretch         =   -1  'True
         Top             =   600
         Width           =   255
      End
      Begin VB.Label Label9 
         Caption         =   "Maximum Defense: ----------------------------------"
         Height          =   255
         Index           =   2
         Left            =   -74400
         TabIndex        =   55
         Top             =   3360
         Width           =   3015
      End
      Begin VB.Label Label8 
         Caption         =   "Present Defense: -------------------------------------"
         Height          =   255
         Index           =   2
         Left            =   -74400
         TabIndex        =   54
         Top             =   3000
         Width           =   3015
      End
      Begin VB.Label Label7 
         Caption         =   "Maximum Attack: -------------------------------------"
         Height          =   255
         Index           =   2
         Left            =   -74400
         TabIndex        =   53
         Top             =   2640
         Width           =   3015
      End
      Begin VB.Label Label6 
         Caption         =   "Present Attack: ---------------------------------------"
         Height          =   255
         Index           =   2
         Left            =   -74400
         TabIndex        =   52
         Top             =   2280
         Width           =   3015
      End
      Begin VB.Label Label5 
         Caption         =   "Kiai: ---------------------------------------------------------"
         Height          =   255
         Index           =   2
         Left            =   -74400
         TabIndex        =   51
         Top             =   1920
         Width           =   3015
      End
      Begin VB.Label Label4 
         Caption         =   "Maximum HP: ------------------------------------------"
         Height          =   255
         Index           =   2
         Left            =   -74400
         TabIndex        =   50
         Top             =   1560
         Width           =   3015
      End
      Begin VB.Label Label3 
         Caption         =   "Present HP: ---------------------------------------------"
         Height          =   255
         Index           =   2
         Left            =   -74400
         TabIndex        =   49
         Top             =   1200
         Width           =   3015
      End
      Begin VB.Label Label2 
         Caption         =   "Number of kills to gain a level: -----------------"
         Height          =   255
         Index           =   2
         Left            =   -74400
         TabIndex        =   48
         Top             =   840
         Width           =   3015
      End
      Begin VB.Label Label1 
         Caption         =   "Movement Rate: --------------------------------------"
         Height          =   255
         Index           =   2
         Left            =   -74400
         TabIndex        =   47
         Top             =   480
         Width           =   3015
      End
      Begin VB.Label Label1 
         Caption         =   "Movement Rate: --------------------------------------"
         Height          =   255
         Index           =   3
         Left            =   -74400
         TabIndex        =   29
         Top             =   480
         Width           =   3015
      End
      Begin VB.Label Label2 
         Caption         =   "Number of kills to gain a level: -----------------"
         Height          =   255
         Index           =   3
         Left            =   -74400
         TabIndex        =   28
         Top             =   840
         Width           =   3015
      End
      Begin VB.Label Label3 
         Caption         =   "Present HP: ---------------------------------------------"
         Height          =   255
         Index           =   3
         Left            =   -74400
         TabIndex        =   27
         Top             =   1200
         Width           =   3015
      End
      Begin VB.Label Label4 
         Caption         =   "Maximum HP: ------------------------------------------"
         Height          =   255
         Index           =   3
         Left            =   -74400
         TabIndex        =   26
         Top             =   1560
         Width           =   3015
      End
      Begin VB.Label Label5 
         Caption         =   "Kiai: ---------------------------------------------------------"
         Height          =   255
         Index           =   3
         Left            =   -74400
         TabIndex        =   25
         Top             =   1920
         Width           =   3015
      End
      Begin VB.Label Label6 
         Caption         =   "Present Attack: ---------------------------------------"
         Height          =   255
         Index           =   3
         Left            =   -74400
         TabIndex        =   24
         Top             =   2280
         Width           =   3015
      End
      Begin VB.Label Label7 
         Caption         =   "Maximum Attack: -------------------------------------"
         Height          =   255
         Index           =   3
         Left            =   -74400
         TabIndex        =   23
         Top             =   2640
         Width           =   3015
      End
      Begin VB.Label Label8 
         Caption         =   "Present Defense: -------------------------------------"
         Height          =   255
         Index           =   3
         Left            =   -74400
         TabIndex        =   22
         Top             =   3000
         Width           =   3015
      End
      Begin VB.Label Label9 
         Caption         =   "Maximum Defense: ----------------------------------"
         Height          =   255
         Index           =   3
         Left            =   -74400
         TabIndex        =   21
         Top             =   3360
         Width           =   3015
      End
      Begin VB.Label Label1 
         Caption         =   "Movement Rate: --------------------------------------"
         Height          =   255
         Index           =   1
         Left            =   -74400
         TabIndex        =   20
         Top             =   480
         Width           =   3015
      End
      Begin VB.Label Label2 
         Caption         =   "Number of kills to gain a level: -----------------"
         Height          =   255
         Index           =   1
         Left            =   -74400
         TabIndex        =   19
         Top             =   840
         Width           =   3015
      End
      Begin VB.Label Label3 
         Caption         =   "Present HP: ---------------------------------------------"
         Height          =   255
         Index           =   1
         Left            =   -74400
         TabIndex        =   18
         Top             =   1200
         Width           =   3015
      End
      Begin VB.Label Label4 
         Caption         =   "Maximum HP: ------------------------------------------"
         Height          =   255
         Index           =   1
         Left            =   -74400
         TabIndex        =   17
         Top             =   1560
         Width           =   3015
      End
      Begin VB.Label Label5 
         Caption         =   "Kiai: ---------------------------------------------------------"
         Height          =   255
         Index           =   1
         Left            =   -74400
         TabIndex        =   16
         Top             =   1920
         Width           =   3015
      End
      Begin VB.Label Label6 
         Caption         =   "Present Attack: ---------------------------------------"
         Height          =   255
         Index           =   1
         Left            =   -74400
         TabIndex        =   15
         Top             =   2280
         Width           =   3015
      End
      Begin VB.Label Label7 
         Caption         =   "Maximum Attack: -------------------------------------"
         Height          =   255
         Index           =   1
         Left            =   -74400
         TabIndex        =   14
         Top             =   2640
         Width           =   3015
      End
      Begin VB.Label Label8 
         Caption         =   "Present Defense: -------------------------------------"
         Height          =   255
         Index           =   1
         Left            =   -74400
         TabIndex        =   13
         Top             =   3000
         Width           =   3015
      End
      Begin VB.Label Label9 
         Caption         =   "Maximum Defense: ----------------------------------"
         Height          =   255
         Index           =   1
         Left            =   -74400
         TabIndex        =   12
         Top             =   3360
         Width           =   3015
      End
      Begin VB.Label Label9 
         Caption         =   "Maximum Defense: ---------------------------------"
         Height          =   255
         Index           =   0
         Left            =   600
         TabIndex        =   10
         Top             =   3360
         Width           =   3015
      End
      Begin VB.Label Label8 
         Caption         =   "Present Defense: ------------------------------------"
         Height          =   255
         Index           =   0
         Left            =   600
         TabIndex        =   9
         Top             =   3000
         Width           =   3015
      End
      Begin VB.Label Label7 
         Caption         =   "Maximum Attack: ------------------------------------"
         Height          =   255
         Index           =   0
         Left            =   600
         TabIndex        =   8
         Top             =   2640
         Width           =   3015
      End
      Begin VB.Label Label6 
         Caption         =   "Present Attack: ---------------------------------------"
         Height          =   255
         Index           =   0
         Left            =   600
         TabIndex        =   7
         Top             =   2280
         Width           =   3015
      End
      Begin VB.Label Label5 
         Caption         =   "Kiai: ---------------------------------------------------------"
         Height          =   255
         Index           =   0
         Left            =   600
         TabIndex        =   6
         Top             =   1920
         Width           =   3015
      End
      Begin VB.Label Label4 
         Caption         =   "Maximum HP: ------------------------------------------"
         Height          =   255
         Index           =   0
         Left            =   600
         TabIndex        =   5
         Top             =   1560
         Width           =   3015
      End
      Begin VB.Label Label3 
         Caption         =   "Present HP: ---------------------------------------------"
         Height          =   255
         Index           =   0
         Left            =   600
         TabIndex        =   4
         Top             =   1200
         Width           =   3015
      End
      Begin VB.Label Label2 
         Caption         =   "Number of kills to gain a level: -----------------"
         Height          =   255
         Index           =   0
         Left            =   600
         TabIndex        =   3
         Top             =   840
         Width           =   3015
      End
      Begin VB.Label Label1 
         Caption         =   "Movement Rate: --------------------------------------"
         Height          =   255
         Index           =   0
         Left            =   600
         TabIndex        =   2
         Top             =   480
         Width           =   3015
      End
   End
End
Attribute VB_Name = "main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub SaveState(File As String)
'Ok this says that an error occurs to bring up a dialog saying
'an error occured instead of a messageboax saying saved...
On Error GoTo error
'This will Dim Curval as a Byte. A Byte can access 1 OFFSET.
'If you Dim Curlng as a Long then you can access 4 OFFSETS.
'You ussualy want to use Integer though. Use by if the Offsets
'are 12341, 12340 then you put the following
'Put #1, 12340 + 1, Curint
'Curint = Val(TEXTBOX)
 Dim CurVal As Byte
 'This opens the File for access. You ussualy want to change #1
 'to #FF, and put FF = Freefile before opening it.
 Open File For Binary As #1
 'This says Curval equals the value of the textbox.
 CurVal = Val(p1mov)
 'This says to put that info in that OFFSET
 Put #1, 68649 + 1, CurVal
 CurVal = Val(p1kill)
 Put #1, 68691 + 1, CurVal
   CurVal = Val(p1preshp)
 Put #1, 68692 + 1, CurVal
   CurVal = Val(p1maxhp)
 Put #1, 68692 + 1, CurVal
   CurVal = Val(p1kia)
 Put #1, 68693 + 1, CurVal
  CurVal = Val(p1presat)
 Put #1, 68696 + 1, CurVal
  CurVal = Val(p1maxatk)
 Put #1, 68697 + 1, CurVal
   CurVal = Val(p1presdef)
 Put #1, 68699 + 1, CurVal
   CurVal = Val(p1maxdef)
 Put #1, 68700 + 1, CurVal
   CurVal = Val(p2mov)
 Put #1, 68729 + 1, CurVal
   CurVal = Val(p2kill)
 Put #1, 68768 + 1, CurVal
   CurVal = Val(p2preshp)
 Put #1, 68771 + 1, CurVal
   CurVal = Val(p2maxhp)
 Put #1, 68772 + 1, CurVal
   CurVal = Val(p2kia)
 Put #1, 68773 + 1, CurVal
   CurVal = Val(p2presatk)
 Put #1, 68776 + 1, CurVal
   CurVal = Val(p2maxatk)
 Put #1, 68777 + 1, CurVal
   CurVal = Val(p2presdef)
 Put #1, 68779 + 1, CurVal
   CurVal = Val(p2maxdef)
 Put #1, 68780 + 1, CurVal
   CurVal = Val(p3mov)
 Put #1, 68809 + 1, CurVal
   CurVal = Val(p3kill)
  Put #1, 68848 + 1, CurVal
   CurVal = Val(p3preshp)
  Put #1, 68851 + 1, CurVal
   CurVal = Val(p3maxhp)
  Put #1, 68852 + 1, CurVal
   CurVal = Val(p3kia)
  Put #1, 68853 + 1, CurVal
   CurVal = Val(p3presatk)
  Put #1, 68856 + 1, CurVal
   CurVal = Val(p3maxatk)
  Put #1, 68857 + 1, CurVal
   CurVal = Val(p3presdef)
  Put #1, 68859 + 1, CurVal
   CurVal = Val(p3maxdef)
  Put #1, 68860 + 1, CurVal
   CurVal = Val(p4mov)
  Put #1, 68889 + 1, CurVal
   CurVal = Val(p4kill)
  Put #1, 68928 + 1, CurVal
   CurVal = Val(p4preshp)
  Put #1, 68931 + 1, CurVal
   CurVal = Val(p4maxhp)
  Put #1, 68932 + 1, CurVal
   CurVal = Val(p4kia)
  Put #1, 68933 + 1, CurVal
   CurVal = Val(p4presatk)
  Put #1, 68936 + 1, CurVal
   CurVal = Val(p4maxatk)
  Put #1, 68937 + 1, CurVal
   CurVal = Val(p4presdef)
  Put #1, 68939 + 1, CurVal
   CurVal = Val(p4maxdef)
  Put #1, 68940 + 1, CurVal
  
    CurVal = Val(enm1alive.Value)
  Put #1, 68960 + 1, CurVal
    CurVal = Val(enm2alive.Value)
  Put #1, 69040 + 1, CurVal
    CurVal = Val(enm3alive.Value)
  Put #1, 69120 + 1, CurVal
    CurVal = Val(enm4alive.Value)
  Put #1, 69200 + 1, CurVal
    CurVal = Val(havep1.Value)
  Put #1, 68627 + 1, CurVal
    CurVal = Val(havep2.Value)
  Put #1, 68628 + 1, CurVal
    CurVal = Val(havep3.Value)
  Put #1, 68629 + 1, CurVal
    CurVal = Val(havep4.Value)
  Put #1, 68630 + 1, CurVal
    CurVal = Val(havep5.Value)
  Put #1, 68631 + 1, CurVal
    CurVal = Val(havep6.Value)
  Put #1, 68632 + 1, CurVal
    CurVal = Val(havep7.Value)
  Put #1, 68633 + 1, CurVal
    CurVal = Val(havep8.Value)
  Put #1, 68634 + 1, CurVal
    CurVal = Val(havep9.Value)
  Put #1, 68635 + 1, CurVal
    CurVal = Val(havep10.Value)
  Put #1, 68636 + 1, CurVal
    CurVal = Val(havep11.Value)
  Put #1, 68637 + 1, CurVal
    CurVal = Val(havep12.Value)
  Put #1, 68638 + 1, CurVal
    CurVal = Val(havep13.Value)
  Put #1, 68639 + 1, CurVal
'--------------------------------
'Remember to close the file!
 Close #1
'Suuccessfull Save
 MsgBox "Saved!"
 Exit Sub
'Unsuuccessfull Save
error:
 MsgBox "Error!"
End Sub
Sub OpenState(File As String)
'Same concept as the save function.
On Error Resume Next
 Dim CurVal As Byte
 Open File For Binary As #1
  Get #1, 68649 + 1, CurVal
 p1mov = CurVal
 
  Get #1, 68691 + 1, CurVal
 p1kill = CurVal
 
  Get #1, 68692 + 1, CurVal
 p1preshp = CurVal
 
  Get #1, 68692 + 1, CurVal
 p1maxhp = CurVal
 
  Get #1, 68693 + 1, CurVal
 p1kia = CurVal

  Get #1, 68696 + 1, CurVal
 p1presatk = CurVal

  Get #1, 68697 + 1, CurVal
 p1maxatk = CurVal
 
  Get #1, 68699 + 1, CurVal
 p1presdef = CurVal
 
  Get #1, 68700 + 1, CurVal
 p1maxdef = CurVal
 
  Get #1, 68729 + 1, CurVal
 p2mov = CurVal
 
  Get #1, 68768 + 1, CurVal
 p2kill = CurVal
 
  Get #1, 68771 + 1, CurVal
 p2preshp = CurVal
 
  Get #1, 68772 + 1, CurVal
 p2maxhp = CurVal
 
  Get #1, 68773 + 1, CurVal
 p2kia = CurVal

  Get #1, 68776 + 1, CurVal
 p2presatk = CurVal

  Get #1, 68777 + 1, CurVal
 p2maxatk = CurVal
 
  Get #1, 68779 + 1, CurVal
 p2presdef = CurVal
 
  Get #1, 68780 + 1, CurVal
 p2maxdef = CurVal
 
  Get #1, 68809 + 1, CurVal
 p3mov = CurVal
 
  Get #1, 68848 + 1, CurVal
 p3kill = CurVal
 
  Get #1, 68851 + 1, CurVal
 p3preshp = CurVal
 
  Get #1, 68852 + 1, CurVal
 p3maxhp = CurVal
 
  Get #1, 68853 + 1, CurVal
 p3kia = CurVal

  Get #1, 68856 + 1, CurVal
 p3presatk = CurVal

  Get #1, 68857 + 1, CurVal
 p3maxatk = CurVal
 
  Get #1, 68859 + 1, CurVal
 p3presdef = CurVal
 
  Get #1, 68860 + 1, CurVal
 p3maxdef = CurVal
 
   Get #1, 68889 + 1, CurVal
 p4mov = CurVal
 
  Get #1, 68928 + 1, CurVal
 p4kill = CurVal
 
  Get #1, 68931 + 1, CurVal
 p4preshp = CurVal
 
  Get #1, 68932 + 1, CurVal
 p4maxhp = CurVal
 
  Get #1, 68933 + 1, CurVal
 p4kia = CurVal

  Get #1, 68936 + 1, CurVal
 p4presatk = CurVal

  Get #1, 68937 + 1, CurVal
 p4maxatk = CurVal
 
  Get #1, 68939 + 1, CurVal
 p4presdef = CurVal
 
  Get #1, 68940 + 1, CurVal
 p4maxdef = CurVal
 

  Get #1, 68960 + 1, CurVal
 enm1preshp = CurVal
   Get #1, 69040 + 1, CurVal
 enm2preshp = CurVal
   Get #1, 69120 + 1, CurVal
 enm3preshp = CurVal
   Get #1, 69200 + 1, CurVal
 enm4preshp = CurVal
 If enm1preshp.Text = "0" Then
 enm1alive.Value = 0
 Else
 enm1alive.Value = 1
 End If
 If enm2preshp.Text = "0" Then
 enm2alive.Value = 0
 Else
 enm2alive.Value = 1
 End If
 If enm3preshp.Text = "0" Then
 enm3alive.Value = 0
 Else
 enm3alive.Value = 1
 End If
 If enm4preshp.Text = "0" Then
 enm4alive.Value = 0
 Else
 enm4alive.Value = 1
 End If
  Get #1, 69011 + 1, CurVal
 enm1preshp = CurVal
  Get #1, 69091 + 1, CurVal
 enm2preshp = CurVal
  Get #1, 69171 + 1, CurVal
 enm3preshp = CurVal
  Get #1, 69251 + 1, CurVal
 enm4preshp = CurVal
 
 Get #1, 68627 + 1, CurVal
 havep1.Value = CurVal
 Get #1, 68628 + 1, CurVal
 havep2.Value = CurVal
 Get #1, 68629 + 1, CurVal
 havep3.Value = CurVal
 Get #1, 68630 + 1, CurVal
 havep4.Value = CurVal
 Get #1, 68631 + 1, CurVal
 havep5.Value = CurVal
 Get #1, 68632 + 1, CurVal
 havep6.Value = CurVal
 Get #1, 68633 + 1, CurVal
 havep7.Value = CurVal
 Get #1, 68634 + 1, CurVal
 havep8.Value = CurVal
 Get #1, 68635 + 1, CurVal
 havep9.Value = CurVal
 Get #1, 68636 + 1, CurVal
 havep10.Value = CurVal
 Get #1, 68637 + 1, CurVal
 havep11.Value = CurVal
 Get #1, 68638 + 1, CurVal
 havep12.Value = CurVal
 Get #1, 68639 + 1, CurVal
 havep13.Value = CurVal
 
 Get #1, 68685 + 1, CurVal
 p1lv = CurVal
 Get #1, 68765 + 1, CurVal
 p2lv = CurVal
 Get #1, 68845 + 1, CurVal
 p3lv = CurVal
 Get #1, 68925 + 1, CurVal
 p4lv = CurVal
'----------------------------
 Close #1
End Sub
Private Sub Check2_Click()

End Sub

Private Sub Check8_Click()

End Sub

Private Sub cmdeasykill_Click()
'Extra options
p1kill.Text = "0"
p2kill.Text = "0"
p3kill.Text = "0"
p4kill.Text = "0"
enm1preshp.Text = "1"
enm2preshp.Text = "1"
enm3preshp.Text = "1"
enm4preshp.Text = "1"
enm1alive.Value = 1
enm2alive.Value = 1
enm3alive.Value = 1
enm4alive.Value = 1
End Sub

Private Sub cmdmaxstats_Click()
'Extra options
p1mov.Text = "99"
p2mov.Text = "99"
p3mov.Text = "99"
p4mov.Text = "99"
p1preshp.Text = "99"
p2preshp.Text = "99"
p3preshp.Text = "99"
p4preshp.Text = "99"
p1maxhp.Text = "99"
p2maxhp.Text = "99"
p3maxhp.Text = "99"
p4maxhp.Text = "99"
p1kia.Text = "99"
p2kia.Text = "99"
p3kia.Text = "99"
p4kia.Text = "99"
p1presdef.Text = "99"
p2presdef.Text = "99"
p3presdef.Text = "99"
p4presdef.Text = "99"
p1maxdef.Text = "99"
p2maxdef.Text = "99"
p3maxdef.Text = "99"
p4maxdef.Text = "99"
p1presatk.Text = "99"
p2presatk.Text = "99"
p3presatk.Text = "99"
p4presatk.Text = "99"
p1maxatk.Text = "99"
p2maxatk.Text = "99"
p3maxatk.Text = "99"
p4maxatk.Text = "99"
End Sub

Private Sub cmdOpen_Click()
'Open the file
On Error Resume Next
 dlg.ShowOpen
 If Err.Number > 0 Then Exit Sub
 File = dlg.FileName
 OpenState (File)
 On Error Resume Next
 RichTextBox1.Text = dlg.FileName
 RichTextBox1.SaveFile (App.Path + "\01.dat")
End Sub

Private Sub cmdSave_Click()
'Save the file
'This program has the backup feature which makes a back up in the backup
'folder. Also stores the last opened file.
If Check1.Value = 1 Then
MkDir (App.Path + "\Back")
dlg.FileName = RichTextBox1.Text
FileCopy RichTextBox1.Text, App.Path + "\Back\" + "backup01.BAK"
End If
 File = dlg.FileName
 If File = "" Then
File = RichTextBox1.Text
End If
 SaveState (File)
End Sub

Private Sub Text11_Change()

End Sub

Private Sub Text17_Change()

End Sub

Private Sub Command1_Click()
'Refresh the stats, and prompt for save before opening a new one.
On Error Resume Next
response = MsgBox("Do you wish to save?", vbYesNo)
If response = vbYes Then
File = dlg.FileName
If File = "" Then
File = RichTextBox1.Text
End If
SaveState (File)
End If
If response = vbNo Then
End If
RichTextBox1.LoadFile (App.Path + "\01.dat")
OpenState RichTextBox1.Text
End Sub

Private Sub Command2_Click()

End Sub

Private Sub Form_Initialize()
'Open last opened file.
On Error Resume Next
Open (App.Path + "\01.dat") For Append As #1
Close #1
RichTextBox1.LoadFile (App.Path + "\01.dat")
OpenState RichTextBox1.Text
End Sub

Private Sub Form_Unload(Cancel As Integer)
'Prompt for saving before opening a new one.
response = MsgBox("Do you wish to save?", vbYesNo)
If response = vbYes Then
 File = dlg.FileName
 If File = "" Then
File = RichTextBox1.Text
End If
 SaveState (File)
End If
If response = vbNo Then
End If
End Sub

Private Sub Label10_Click()

End Sub

