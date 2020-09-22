VERSION 5.00
Object = "{22D6F304-B0F6-11D0-94AB-0080C74C7E95}#1.0#0"; "msdxm.ocx"
Begin VB.Form Form1 
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "reVAMP"
   ClientHeight    =   1680
   ClientLeft      =   3210
   ClientTop       =   2685
   ClientWidth     =   4110
   BeginProperty Font 
      Name            =   "MS Serif"
      Size            =   13.5
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "Form1.frx":08CA
   ScaleHeight     =   1680
   ScaleWidth      =   4110
   Begin VB.PictureBox pclick 
      BackColor       =   &H80000007&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4440
      Picture         =   "Form1.frx":0A98
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   34
      Top             =   1440
      Width           =   615
   End
   Begin VB.PictureBox pmove 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4440
      Picture         =   "Form1.frx":1A5A
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   33
      Top             =   120
      Width           =   615
   End
   Begin VB.PictureBox pmouse 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4440
      Picture         =   "Form1.frx":1D64
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   32
      Top             =   720
      Width           =   615
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   300
      Left            =   3480
      Top             =   3360
   End
   Begin VB.PictureBox lstop 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5160
      Picture         =   "Form1.frx":206E
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   24
      Top             =   1320
      Width           =   255
   End
   Begin VB.PictureBox lplay 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5160
      Picture         =   "Form1.frx":223C
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   23
      Top             =   720
      Width           =   255
   End
   Begin VB.PictureBox lpause 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5160
      Picture         =   "Form1.frx":240A
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   22
      Top             =   120
      Width           =   255
   End
   Begin VB.PictureBox main 
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   0
      Picture         =   "Form1.frx":25D8
      ScaleHeight     =   1695
      ScaleWidth      =   4215
      TabIndex        =   1
      Top             =   0
      Width           =   4215
      Begin VB.PictureBox Picture6 
         BorderStyle     =   0  'None
         Height          =   110
         Left            =   2600
         Picture         =   "Form1.frx":19D4A
         ScaleHeight     =   105
         ScaleWidth      =   90
         TabIndex        =   40
         Top             =   900
         Width           =   90
      End
      Begin VB.PictureBox beef 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   179
         Left            =   1670
         ScaleHeight     =   180
         ScaleWidth      =   2325
         TabIndex        =   37
         Top             =   350
         Width           =   2330
         Begin VB.TextBox movenm 
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000C000&
            Height          =   225
            Left            =   20
            Locked          =   -1  'True
            TabIndex        =   38
            Text            =   "*** reVAMP v.1"
            Top             =   0
            Width           =   2295
         End
         Begin VB.Label lblName 
            BackColor       =   &H80000007&
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000C000&
            Height          =   225
            Left            =   0
            TabIndex        =   41
            Top             =   0
            Visible         =   0   'False
            Width           =   2295
         End
      End
      Begin VB.PictureBox Picture5 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3550
         Picture         =   "Form1.frx":19E18
         ScaleHeight     =   375
         ScaleWidth      =   495
         TabIndex        =   36
         Top             =   1370
         Width           =   495
      End
      Begin VB.PictureBox Picture4 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   160
         Left            =   1660
         Picture         =   "Form1.frx":1A7A2
         ScaleHeight     =   165
         ScaleWidth      =   720
         TabIndex        =   35
         Top             =   40
         Width           =   715
      End
      Begin VB.PictureBox balbar 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   230
         Left            =   2610
         Picture         =   "Form1.frx":1AE14
         ScaleHeight     =   225
         ScaleWidth      =   645
         TabIndex        =   29
         Top             =   840
         Width           =   640
         Begin VB.PictureBox balance2 
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   160
            Left            =   240
            Picture         =   "Form1.frx":1B68A
            ScaleHeight     =   165
            ScaleWidth      =   210
            TabIndex        =   31
            Top             =   25
            Visible         =   0   'False
            Width           =   208
         End
         Begin VB.PictureBox balance1 
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   160
            Left            =   240
            Picture         =   "Form1.frx":1B8B0
            ScaleHeight     =   165
            ScaleWidth      =   210
            TabIndex        =   30
            Top             =   25
            Width           =   208
         End
      End
      Begin VB.PictureBox poslong 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   217
         Left            =   240
         Picture         =   "Form1.frx":1BAD6
         ScaleHeight     =   210
         ScaleWidth      =   3780
         TabIndex        =   27
         Top             =   1080
         Width           =   3780
         Begin VB.PictureBox position 
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   155
            Left            =   0
            Picture         =   "Form1.frx":1E470
            ScaleHeight     =   150
            ScaleWidth      =   435
            TabIndex        =   28
            Top             =   30
            Visible         =   0   'False
            Width           =   440
         End
      End
      Begin VB.PictureBox lilstuff 
         BackColor       =   &H80000007&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   21
         Top             =   360
         Width           =   255
      End
      Begin VB.PictureBox Picture3 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   253
         Left            =   1570
         Picture         =   "Form1.frx":1E822
         ScaleHeight     =   255
         ScaleWidth      =   1005
         TabIndex        =   16
         Top             =   820
         Width           =   1000
         Begin VB.PictureBox volknob2 
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   160
            Left            =   480
            Picture         =   "Form1.frx":1F74C
            ScaleHeight     =   165
            ScaleWidth      =   210
            TabIndex        =   18
            Top             =   50
            Visible         =   0   'False
            Width           =   208
         End
         Begin VB.PictureBox volknob 
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   160
            Left            =   480
            Picture         =   "Form1.frx":1F972
            ScaleHeight     =   165
            ScaleWidth      =   210
            TabIndex        =   17
            Top             =   50
            Width           =   208
         End
      End
      Begin VB.PictureBox eject2 
         BackColor       =   &H80000007&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   250
         Left            =   2040
         Picture         =   "Form1.frx":1FB98
         ScaleHeight     =   255
         ScaleWidth      =   330
         TabIndex        =   15
         Top             =   1320
         Visible         =   0   'False
         Width           =   335
      End
      Begin VB.PictureBox eject 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   245
         Left            =   2040
         Picture         =   "Form1.frx":2001A
         ScaleHeight     =   240
         ScaleWidth      =   330
         TabIndex        =   14
         Top             =   1320
         Width           =   335
      End
      Begin VB.PictureBox right2 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   265
         Left            =   1600
         Picture         =   "Form1.frx":2049C
         ScaleHeight     =   270
         ScaleWidth      =   360
         TabIndex        =   13
         Top             =   1320
         Visible         =   0   'False
         Width           =   365
      End
      Begin VB.PictureBox right1 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1620
         Picture         =   "Form1.frx":209EE
         ScaleHeight     =   255
         ScaleWidth      =   345
         TabIndex        =   12
         Top             =   1320
         Width           =   345
      End
      Begin VB.PictureBox stop2 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   1250
         Picture         =   "Form1.frx":20F40
         ScaleHeight     =   270
         ScaleWidth      =   360
         TabIndex        =   11
         Top             =   1320
         Visible         =   0   'False
         Width           =   355
      End
      Begin VB.PictureBox stop0 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1270
         Picture         =   "Form1.frx":21492
         ScaleHeight     =   255
         ScaleWidth      =   345
         TabIndex        =   10
         Top             =   1320
         Width           =   345
      End
      Begin VB.PictureBox pause2 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   275
         Left            =   900
         Picture         =   "Form1.frx":219E4
         ScaleHeight     =   270
         ScaleWidth      =   360
         TabIndex        =   6
         Top             =   1320
         Visible         =   0   'False
         Width           =   365
         Begin VB.PictureBox Picture2 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   360
            ScaleHeight     =   255
            ScaleWidth      =   15
            TabIndex        =   9
            Top             =   0
            Width           =   15
         End
      End
      Begin VB.PictureBox pause 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   930
         Picture         =   "Form1.frx":21F36
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   8
         Top             =   1320
         Width           =   255
      End
      Begin VB.PictureBox play2 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         FillColor       =   &H00C000C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   265
         Left            =   550
         Picture         =   "Form1.frx":22488
         ScaleHeight     =   270
         ScaleWidth      =   345
         TabIndex        =   5
         Top             =   1320
         Visible         =   0   'False
         Width           =   345
         Begin VB.PictureBox Picture1 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   360
            ScaleHeight     =   255
            ScaleWidth      =   15
            TabIndex        =   7
            Top             =   0
            Width           =   15
         End
      End
      Begin VB.PictureBox left2 
         BackColor       =   &H80000003&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   265
         Left            =   240
         Picture         =   "Form1.frx":229DA
         ScaleHeight     =   270
         ScaleWidth      =   345
         TabIndex        =   4
         Top             =   1320
         Visible         =   0   'False
         Width           =   345
      End
      Begin VB.PictureBox play 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   600
         Picture         =   "Form1.frx":22F2C
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   3
         Top             =   1320
         Width           =   255
      End
      Begin VB.PictureBox left1 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         Picture         =   "Form1.frx":2347E
         ScaleHeight     =   255
         ScaleWidth      =   375
         TabIndex        =   2
         Top             =   1320
         Width           =   375
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000007&
         Caption         =   "Label1"
         Height          =   570
         Left            =   200
         TabIndex        =   39
         Top             =   360
         Width           =   120
      End
      Begin VB.Label mp3time 
         BackColor       =   &H80000007&
         Caption         =   "  :"
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   375
         Left            =   600
         TabIndex        =   25
         Top             =   360
         Width           =   855
      End
      Begin VB.Label khz 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   135
         Left            =   2298
         TabIndex        =   20
         Top             =   610
         Width           =   200
      End
      Begin VB.Label kbps 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   130
         Left            =   1650
         TabIndex        =   19
         Top             =   620
         Width           =   245
      End
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000009&
      Caption         =   "Percent Done ="
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   42
      Top             =   2400
      Width           =   2055
   End
   Begin VB.Label mpercent 
      BackColor       =   &H80000007&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   2520
      TabIndex        =   26
      Top             =   2400
      Width           =   1455
   End
   Begin MediaPlayerCtl.MediaPlayer MediaPlayer1 
      Height          =   375
      Left            =   840
      TabIndex        =   0
      Top             =   2880
      Width           =   3135
      AudioStream     =   -1
      AutoSize        =   0   'False
      AutoStart       =   -1  'True
      AnimationAtStart=   -1  'True
      AllowScan       =   -1  'True
      AllowChangeDisplaySize=   -1  'True
      AutoRewind      =   0   'False
      Balance         =   0
      BaseURL         =   ""
      BufferingTime   =   5
      CaptioningID    =   ""
      ClickToPlay     =   -1  'True
      CursorType      =   0
      CurrentPosition =   -1
      CurrentMarker   =   0
      DefaultFrame    =   ""
      DisplayBackColor=   0
      DisplayForeColor=   16777215
      DisplayMode     =   0
      DisplaySize     =   4
      Enabled         =   -1  'True
      EnableContextMenu=   -1  'True
      EnablePositionControls=   -1  'True
      EnableFullScreenControls=   0   'False
      EnableTracker   =   -1  'True
      Filename        =   ""
      InvokeURLs      =   -1  'True
      Language        =   -1
      Mute            =   0   'False
      PlayCount       =   1
      PreviewMode     =   0   'False
      Rate            =   1
      SAMILang        =   ""
      SAMIStyle       =   ""
      SAMIFileName    =   ""
      SelectionStart  =   -1
      SelectionEnd    =   -1
      SendOpenStateChangeEvents=   -1  'True
      SendWarningEvents=   -1  'True
      SendErrorEvents =   -1  'True
      SendKeyboardEvents=   0   'False
      SendMouseClickEvents=   0   'False
      SendMouseMoveEvents=   0   'False
      SendPlayStateChangeEvents=   -1  'True
      ShowCaptioning  =   0   'False
      ShowControls    =   -1  'True
      ShowAudioControls=   -1  'True
      ShowDisplay     =   0   'False
      ShowGotoBar     =   0   'False
      ShowPositionControls=   -1  'True
      ShowStatusBar   =   0   'False
      ShowTracker     =   -1  'True
      TransparentAtStart=   0   'False
      VideoBorderWidth=   0
      VideoBorderColor=   0
      VideoBorder3D   =   0   'False
      Volume          =   0
      WindowlessVideo =   0   'False
   End
   Begin VB.Menu mnu 
      Caption         =   "mnu"
      Visible         =   0   'False
      Begin VB.Menu mnu5 
         Caption         =   "RydeSoft re&VAMP..."
         Index           =   1
      End
      Begin VB.Menu fd 
         Caption         =   "-"
      End
      Begin VB.Menu fdsfds 
         Caption         =   "&Controls"
         Begin VB.Menu mnuopen 
            Caption         =   "Open &File..."
         End
         Begin VB.Menu mnuplay 
            Caption         =   "&Play     "
         End
         Begin VB.Menu mnustop 
            Caption         =   "&Stop     "
         End
         Begin VB.Menu mnupause 
            Caption         =   "P&ause"
         End
      End
      Begin VB.Menu fileinfo 
         Caption         =   "View &file info..."
         Shortcut        =   ^{F3}
      End
      Begin VB.Menu kjh 
         Caption         =   "-"
      End
      Begin VB.Menu mnuexit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bitrate_lookup(7, 15) As Integer
Public actual_bitrate As Long


Public Function SongMove()
'Makes the song title move across the label

HaveTime = True
    lblName.Caption = "*** 1. " & Dir(Filep)
    lblName.Caption = UCase(lblName.Caption)
    
    If Right(lblName.Caption, 4) = ".MP3" Or Right(lblName.Caption, 4) = ".mp3" Then
    
        lblName.Caption = Left(lblName.Caption, Len(lblName.Caption) - 4)
        movenm.Text = lblName.Caption
 
            If Len(lblName.Caption) > 5 Then
                movenm.Width = (Len(lblName.Caption) * 80) '100)
            End If

            If Len(lblName.Caption) > 10 Then
                movenm.Width = (Len(lblName.Caption) * 80) '93)
            End If

            If Len(lblName.Caption) > 30 Then
                movenm.Width = (Len(lblName.Caption) * 80) '110)
            End If

         Timer1.Enabled = True
     End If
End Function



   Public Function Getmp3data(MP3File As String)
   'Getting id3 info..
   'Some parts of this were borrowed,
   'But I mostly edited myself..
   
   On Error Resume Next
   'still some bugs hence the non error alerting
   
     Dim dIN As String
     cr = Chr(10)
     Open MP3File For Binary As #1
     '' read in 1st 4k of .mp3 file to find a frame header
     dIN = Input(4096, #1)
     filesize = LOF(1) '' needed to calculate track duration
     Close #1
     
     '' frame header starts with 12 set bits [sync]
     '' NB this ignores MPEG-2.5 which is 11 set bits, 1 zero bit.
     
     '' my search for the sync bits only works on nibble boundaries,
     '' I'm not sure if it is necessary to search on bit boundaries -
     '' if so then this search will be 4* slower and require a rewrite
     '' of this search section and shift_those_bits.
     
     Do Until i = 4095
       i = i + 1
       d1 = Asc(Mid(dIN, i, 1))
       d2 = Asc(Mid(dIN, i + 1, 1))
       If d1 = &HFF And (d2 And &HF0) = &HF0 Then
         '' get 20 hdr bits - they are last 20 bits of next 3 bytes
         temp_string = Mid(dIN, i + 1, 3)
         mp3bits_string = shift_those_bits(Mid(dIN, i + 1, 3))
         Exit Do
       End If
       '' if we haven't found the sync yet then shift left by 4 bits
       dSHIFT = shift_those_bits(Mid(dIN, i, 3))
       dd1 = Asc(Left(dSHIFT, 1))
       dd2 = Asc(Right(dSHIFT, 1))
       If dd1 = &HFF And (dd2 And &HF0) = &HF0 Then
        '' get 20 hdr bits - they are first 20 bits of next 3 bytes
         mp3bits_string = Mid(dIN, i + 2, 3)
         Exit Do
       End If
     Loop
     
      
     '' 1st 20 bits of mp3bits_string are hdr info for this frame
     '' 1st bit is ID - 0=MPG-2, 1=MPG-1
     mp3_id = (&H80 And Asc(Left(mp3bits_string, 1))) / 128
     ''next 2 bits are Layer
     mp3_layer = (&H60 And Asc(Left(mp3bits_string, 1))) / 32
     ''next bit is Protection
     mp3_prot = &H10 And Asc(Left(mp3bits_string, 1))
     ''next 4 bits are bitrate
     mp3_bitrate = &HF And Asc(Left(mp3bits_string, 1))
     ''next 2 bits are frequency
     mp3_freq = &HC0 And Asc(Mid(mp3bits_string, 2, 1))
     '' next bit is Padding
     mp3_pad = (&H20 And Asc(Mid(mp3bits_string, 2, 1))) / 2
     actual_bitrate = 1000 * CLng((bitrate_lookup((mp3_id * 4) Or mp3_layer, mp3_bitrate)))
     
     'Working out ID
     dat = "ID: "
     If mp3_id = 0 Then
       dat = dat + "MPEG-2"
       mpeg1 = "MPEG 2.0"
     Else
       dat = dat + "MPEG-1"
       mpeg1 = "MPEG 1.0"
     End If
     
      'Working out layer (1,2,3)
      Select Case mp3_layer
        Case 1
          dat = dat + "layer 3"
          layer = "layer 3"
        
        Case 2
          dat = dat + "layer 2"
          layer = "layer 2"
        Case 3
          dat = dat + "layer 1"
          layer = "layer 1"
      End Select
      dat = dat + cr + "Bitrate: " + Str(actual_bitrate)
      
      'Working out freq..
      Select Case (mp3_id * 4) Or mp3_freq
        Case 0
          sample_rate = 22050
        Case 1
          sample_rate = 24000
        Case 2
          sample_rate = 16000
        Case 4
          sample_rate = 44100
        Case 5
          sample_rate = 48000
        Case 6
          sample_rate = 32000
      End Select
      dat = dat + cr + "Sample rate: " + Str(sample_rate)
      
      ' calculate track time
      framesize = ((144 * actual_bitrate) / sample_rate) + mp3_pad
      total_frames = filesize / framesize
      track_length = total_frames / 38.5 '38.5 frames per sec.
      
      'Set the vars..
      mFrames = Str(Int(total_frames))
      mLength = Format(MediaPlayer1.Duration, "#")
      mHz = Str(sample_rate)
      mMpeg = mpeg1 & " " & layer
      mBit = Left(actual_bitrate, 3)
      'Set the 2 captions..
      kbps.Caption = Left(actual_bitrate, 3)
      khz.Caption = Left(sample_rate, 2)
      
      
   End Function
   Public Function shift_those_bits(dIN As String) As String
         '' need to left shift 4 bits losing most significant 4 bits
         Dim sd1, sd2, sd3, do1, do2 As Integer
         duff = Left(dIN, 1)
         duff2 = Asc(duff)
         sd1 = Asc(Left(dIN, 1))
         sd2 = Asc(Mid(dIN, 2, 1))
         sd3 = Asc(Right(dIN, 1))
     
        do1 = ((sd1 And &HF) * 16) Or ((sd2 And &HF0) / 16)
        do2 = ((sd2 And &HF) * 16) Or ((sd3 And &HF0) / 16)
        shift_those_bits = Chr(do1) + Chr(do2)
   End Function


Private Sub balance1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Mouse down on the Balance icon
balance2.Visible = True
End Sub

Private Sub balance1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Mouse move on the Balance icon

If Button > 0 Then
    'Change the icon
    balance1.MouseIcon = pmove.Picture
    balance1.MousePointer = vbCustom
End If

If Button = 0 Then
    'Change the icon
    balance1.MouseIcon = pmouse.Picture
    balance1.MousePointer = vbCustom
End If



Oom = ""
'lbl.Caption = ""
balance2.Top = balance1.Top
balance2.Left = balance1.Left

'Wacky annyoying shit to change the balance.
Moo = X / Form1.Width * 100
Oom = Moo / 100 * (balbar.Width - 100) / 1
End Sub

Private Sub balance1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Mouse up on the Balance icon

movenm.Visible = False
lblName.Visible = True

balance2.Visible = False

Oom = Oom - 12

'Explanation of the Var Oom...
'Oom comes from:
'The pos your cursor is when you click on Balance
'Or X, Divided by the width of the Form,
'Times 100. Then.. That all Divided by 100
'and Times (width of balance_bar -100)
'And the Divided by 1. crazy shit.
'I couldn't think of any other way and it
'took me ages, but it basically gives you
'a representation of where your cursor is once
'a button has been pushed, and it's in motion,
'and then it has stopped.
'it can range from like -808 --> 765 roughly
'and can then be used to calculate the balance.

'You can 'Msgbox Oom, to check it out
'for yourself.

'When the cursor is pushed and mooved on the
'Balance_bar, and then mouse button_up, it
'gives a reading between around 0 --> 111

If Oom > 0 Then
    balance1.Left = balance1.Left + (Oom * 2) '4.6)
End If

If Oom > 55 Then
    balance1.Left = balance1.Left + (Oom * 2) '1.6)
End If

If Oom > 111 Then
    balance1.Left = balance1.Left + (Oom * 2.5) '1.9)
End If

If Oom < 0 Then
    balance1.Left = balance1.Left + (Oom * 3.5)
End If

'Basically used trial and error to calculate
'these bits.

If balance1.Left > 420 Then
    balance1.Left = 420
End If

If balance1.Left < 61 Then
    balance1.Left = 70
End If

If Oom > 1 Then
    balance1.Left = balance1.Left
End If

'Working out the Percentage of Balance,
'disregard 'volpercent' it's not volume
'i just used the name..

volpercent = balance1.Left / 360 * 100 - 19

If volpercent <= 0 Then
    volpercent = "0"
End If

volpercent = Format(volpercent, "#")

If volpercent = "" Then
    volpercent = "0"
End If

If volpercent >= "97" Then
    volpercent = "100"
End If

'balance from -4000 ---> +4000
' thats 8000 (duh!)

If volpercent = 50 Then
    MediaPlayer1.Balance = 0
    lblName.Caption = "BALANCE: CENTER"
End If

If volpercent < 50 Then
    pog = (100 - (volpercent / 50 * 100))
    lblName.Caption = "BALANCE: " & pog & "% LEFT"
    sinbad = (pog / 100 * 4000)
    MediaPlayer1.Balance = ("-" & sinbad)
End If

If volpercent > 50 Then
    mog = 100 - (volpercent / 50 * 100)
    mog = Right(mog, (Len(mog) - 1))
    lblName.Caption = "BALANCE: " & mog & "% RIGHT"
    sinbad = (mog / 100 * 4000)
    MediaPlayer1.Balance = (sinbad)
End If

End Sub

Private Sub balbar_Click()
'Click on the balance_bar..
movenm.Visible = False
lblName.Visible = True
End Sub

Private Sub balbar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Mouse moves on the balance_bar
bla = X
bla2 = X
End Sub

Private Sub balbar_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Mouse up from when you click on the balance_bar.
'To move balance not only when you drag the knob,
'But also when you click on the bar..

'bla = X position
bla = bla / 655 * 100

If bla > 50 Then
    balance1.Left = bla2 - 208
End If

If bla < 50 Then
    balance1.Left = bla2
End If

volpercent = bla

'If 50 then it's in the CENTER
If volpercent = 50 Then
    MediaPlayer1.Balance = 0
    lblName.Caption = "BALANCE: CENTER"
End If

'if < 50 then it's more LEFT speaker
If volpercent < 50 Then
    pog = (100 - (volpercent / 50 * 100))
    pog = Format(pog, "#")
    pog = pog + 25

    If pog > 100 Then
        pog = 100
    End If

    lblName.Caption = "BALANCE: " & pog & "% LEFT"
    sinbad = (pog / 100 * 4000)
    MediaPlayer1.Balance = ("-" & sinbad)
End If

'if >50 then more RIGHT speaker
If volpercent > 50 Then
    mog = 100 - (volpercent / 50 * 100)
    mog = Right(mog, (Len(mog) - 1))
    mog = Format(mog, "#")

    If mog > 100 Then
        mog = 100
    End If

    lblName.Caption = "BALANCE: " & mog & "% RIGHT"
    sinbad = (mog / 100 * 4000)
    MediaPlayer1.Balance = (sinbad)
End If

If balance1.Left < 100 Then
    balance1.Left = 80
End If

End Sub

Private Sub eject_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Mouse move on Eject Button.
'X is from 360 --> 0
'Y is from 30 --> 285

Meject = True

If Button > 0 Then

    If Y > 285 Then
        eject2.Visible = False
        Meject = False
    End If

    If Y < 30 Then
        eject2.Visible = False
        Meject = False
    End If

    If X < 0 Then
        eject2.Visible = False
        Meject = False
    End If

    If X > 360 Then
        eject2.Visible = False
        Meject = False
    End If

End If

End Sub

Private Sub fileinfo_Click()
'Click on the Id3info...
Form2.Show
End Sub

Private Sub Form_Load()
'Blah.. Form load...
'' Playlist.Show??


'System tray icon load..
    With nid
        .cbSize = Len(nid)
        .hwnd = Me.hwnd
        .uId = vbNull
        .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
        .uCallBackMessage = WM_MOUSEMOVE
        .hIcon = Me.Icon
        .szTip = " rydesoft reVAMP " & vbNullChar
    End With
    Shell_NotifyIcon NIM_ADD, nid

'Change to nice silver cursor
main.MouseIcon = pmouse.Picture
main.MousePointer = vbCustom

'Set volume
MediaPlayer1.Volume = "-700"

'Put STOP pic in song screen
lilstuff.Picture = lstop.Picture

'Well as the name says...
SetUpBitrate

'My nice lil easter egg :P
If Right(Command(), 8) = "ryde.mp3" Then
MsgBox "Christopher David Lance Rickard Loves Brooke Emily Street", vbCritical, "Tell it like it is!"
End If

'If it's a mp3 file..
'This is only if you set it as the default
'Player, so if you click on an mp3 it goes here..
If Right(Command(), 4) = ".mp3" Or Right(Command(), 4) = ".MP3" Then
  
  Filep = Command()
    MediaPlayer1.Open (Filep)
  
  lilstuff.Picture = lplay.Picture
  position.Visible = True
  opened = True
    
 'Call Songmove to move the song :P
 SongMove
 
    fname = Filep
  If UCase(Right(fname, 4)) = ".MP3" Then
    Getmp3data (fname)
  Else
   
  End If
  
 'Start Timer1
  Timer1.Enabled = True
  
End If
End Sub


Private Sub SetUpBitrate()

' Setup array for mpeg bitrate info...
  bitrate_data = "032,032,032,032,008,008,"
  bitrate_data = bitrate_data + "064,048,040,048,016,016,"
  bitrate_data = bitrate_data + "096,056,048,056,024,024,"
  bitrate_data = bitrate_data + "128,064,056,064,032,032,"
  bitrate_data = bitrate_data + "160,080,064,080,040,040,"
  bitrate_data = bitrate_data + "192,096,080,096,048,048,"
  bitrate_data = bitrate_data + "224,112,096,112,056,056,"
  bitrate_data = bitrate_data + "256,128,112,128,064,064,"
  bitrate_data = bitrate_data + "288,160,128,144,080,080,"
  bitrate_data = bitrate_data + "320,192,160,160,096,096,"
  bitrate_data = bitrate_data + "352,224,192,176,112,112,"
  bitrate_data = bitrate_data + "384,256,224,192,128,128,"
  bitrate_data = bitrate_data + "416,320,256,224,144,144,"
  bitrate_data = bitrate_data + "448,384,320,256,160,160,"
    
  For Y = 1 To 14
    For X = 7 To 5 Step -1
      bitrate_lookup(X, Y) = Left(bitrate_data, 3)
      bitrate_data = Right(bitrate_data, Len(bitrate_data) - 4)
    Next
    For X = 3 To 1 Step -1
      bitrate_lookup(X, Y) = Left(bitrate_data, 3)
      bitrate_data = Right(bitrate_data, Len(bitrate_data) - 4)
    Next
  Next
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    'System tray icon...
    Dim Result As Long
    Dim msg As Long
    'The value of X will vary depending upon
    'the scalemode setting

    If Me.ScaleMode = vbPixels Then
        msg = X
    Else
        msg = X / Screen.TwipsPerPixelX
    End If

    Select Case msg
        Case WM_LBUTTONUP '514 restore form window
        Me.WindowState = vbNormal
        Result = SetForegroundWindow(Me.hwnd)
        Me.Show
        Case WM_LBUTTONDBLCLK '515 restore form window
        Me.WindowState = vbNormal
        Result = SetForegroundWindow(Me.hwnd)
        Me.Show
        Case WM_RBUTTONUP '517 display popup menu
        Result = SetForegroundWindow(Me.hwnd)
        'Display menu when clicked on system
        'tray icon...
        Me.PopupMenu Me.mnu
    End Select

End Sub

Private Sub Form_Resize()
    'Hide minimized window...
    If Me.WindowState = vbMinimized Then Me.Hide
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Remove icon from system tray...
    Shell_NotifyIcon NIM_DELETE, nid
End Sub

Private Sub eject_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
eject2.Visible = True
End Sub

Private Sub eject_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Mouse up (clicked) on eject button

If Meject = False Then Exit Sub

ballie = 0

'Show the Open File Custom Dialog..
 Dim FileDialog As CFileDialog
  Set FileDialog = New CFileDialog
  
  With FileDialog
    .DefaultExt = "mp3"
    .DialogTitle = "Open file(s)"
    .Filter = "All supported files |*.MP1;*.MP2;*.MP3;*.PLS;*.M3U|MPEG Audio files (*.MP1,2,3)|*.MP1;*.MP2;*.MP3"
    .FilterIndex = 0
    .Flags = FleFileMustExist + FleHideReadOnly + FleCreatePrompt
    .hWndParent = Me.hwnd
    .MaxFileSize = 255
    If .Show(True) Then
        Filen = .FileTitle
        Filep = .FileName
    Else
     ' MsgBox "User cancelled"
    End If
    
  End With


    lblName.Caption = "*** 1. " & Dir(Filep)
    lblName.Caption = UCase(lblName.Caption)
  
  If Right(lblName.Caption, 4) = ".MP3" Then
    lblName.Caption = Left(lblName.Caption, Len(lblName.Caption) - 4)
  End If

    fname = Filep
    
  If UCase(Right(fname, 4)) = ".MP3" Then
    Getmp3data (fname)
  Else
   ' MsgBox "not a .mp3 file"
  End If
  
  HaveTime = True
  
  MediaPlayer1.Open (Filep)
  
  lilstuff.Picture = lplay.Picture
  position.Visible = True
  opened = True
  
  Stopped = False
  
  'Start moving the song in label
  SongMove

  eject2.Visible = False
    
  Timer1.Enabled = True
    
End Sub

Private Sub left1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
left2.Visible = True
End Sub

Private Sub left1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'If you click on LEFT_button

Mleft = True

If Button > 0 Then

    If Y > 285 Then
        left2.Visible = False
        Mleft = False
    End If

    If Y < 30 Then
        left2.Visible = False
        Mleft = False
    End If

    If X < 0 Then
        left2.Visible = False
        Mleft = False
    End If

    If X > 360 Then
        left2.Visible = False
        Mleft = False
    End If

End If

End Sub

Private Sub left1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
left2.Visible = False
If Mleft = False Then Exit Sub
End Sub


Private Sub main_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Show the popum_menu
If Button = 2 Then
    Form1.PopupMenu mnu
End If
End Sub

Private Sub main_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'If mouse moves on main form..

movenm.Visible = True
lblName.Visible = False

If Len(Filen) > 2 Then
    lblName.Caption = UCase(Filen)
    If Right(lblName.Caption, 4) = ".MP3" Then
        lblName.Caption = Left(lblName.Caption, Len(lblName.Caption) - 4)
    End If
End If

End Sub

Private Sub main_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
clicked = False
End Sub

Private Sub mnu5_Click(Index As Integer)
Form3.Show
End Sub

Private Sub mnuexit_Click()
    'Removes icon from system tray...
    Shell_NotifyIcon NIM_DELETE, nid
End
End Sub

Private Sub mnuopen_Click()
'If you click Open on the popup_menu

eject2.Visible = False
 Dim FileDialog As CFileDialog
  Set FileDialog = New CFileDialog
  
  With FileDialog
    .DefaultExt = "mp3"
    .DialogTitle = "Open file(s)"
    .Filter = "All supported files |*.MP1;*.MP2;*.MP3;*.PLS;*.M3U|MPEG Audio files (*.MP1,2,3)|*.MP1;*.MP2;*.MP3"
    .FilterIndex = 0
    .Flags = FleFileMustExist + FleHideReadOnly + FleCreatePrompt
    .hWndParent = Me.hwnd
    .MaxFileSize = 255
    If .Show(True) Then
        Filen = .FileTitle
        Filep = .FileName
        
    Else
     ' MsgBox "User cancelled"
    End If
    
  End With
  
  lblName.Caption = UCase(Filen)
  
  If Right(lblName.Caption, 4) = ".MP3" Then
    lblName.Caption = Left(lblName.Caption, Len(lblName.Caption) - 4)
  End If

    fname = Filep
  If UCase(Right(fname, 4)) = ".MP3" Then
    Getmp3data (fname)
  Else
   ' MsgBox "not a .mp3 file"
  End If
    
  MediaPlayer1.Open (Filep)
  
  lilstuff.Picture = lplay.Picture
  position.Visible = True
    opened = True
  
  'Start moving the song
  SongMove
    
  Timer1.Enabled = True
End Sub

Private Sub mnupause_Click()
'If you click Pause on the popup_Menu

pause2.Visible = False
Timer1.Enabled = False

If Stopped = True Then
    Exit Sub
End If

If MediaPlayer1.OpenState = 0 Then
    Exit Sub
End If

Stopped = False

If Paused = True Then
    MediaPlayer1.play
    lilstuff.Picture = lplay.Picture
    Paused = False
    Timer1.Enabled = True
    popy = 23
    GoTo cowboy
End If

If Paused = False Then
    Paused = True
    MediaPlayer1.pause
    GoTo cowboy
End If

cowboy:

If popy = "23" Then
    GoTo bnaba
End If

lilstuff.Picture = lpause.Picture
bnaba:
End Sub

Private Sub mnuplay_Click()
'If you click play on the popup_Menu

play.Visible = False
Timer1.Enabled = True
Stopped = False

If Paused = False Then
    MediaPlayer1.Open (Filep)
End If

If Paused = True Then
    MediaPlayer1.play
End If

lilstuff.Picture = lplay.Picture
position.Visible = True
End Sub

Private Sub mnustop_Click()
'If you clicked Stop on the popup_menu

Timer1.Enabled = False
Stopped = True
stop2.Visible = False
Paused = False
MediaPlayer1.Stop
lilstuff.Picture = lstop.Picture
position.Visible = False
position.Left = 0
movenm.Left = 0


End Sub

Private Sub mp3time_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
Form1.PopupMenu mnu
End If
End Sub

Private Sub pause_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
pause2.Visible = True
End Sub

Private Sub pause_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Mouse down on Pause Button

Mpause = True

If Button > 0 Then

    If Y > 285 Then
        pause2.Visible = False
        Mpause = False
    End If

    If Y < 30 Then
        pause2.Visible = False
        Mpause = False
    End If

    If X < 0 Then
        pause2.Visible = False
        Mpause = False
    End If

    If X > 360 Then
        pause2.Visible = False
        Mpause = False
    End If

End If
End Sub

Private Sub pause_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Mouse up on Pause_Button

If Mpause = False Then Exit Sub

    pause2.Visible = False
    Timer1.Enabled = False

    If Stopped = True Then
        Exit Sub
    End If

    If MediaPlayer1.OpenState = 0 Then
        Exit Sub
    End If

    Stopped = False

    If Paused = True Then
        MediaPlayer1.play
        lilstuff.Picture = lplay.Picture
        Paused = False
        Timer1.Enabled = True
        popy = 23
        GoTo cowboy
    End If

    If Paused = False Then
        Paused = True
        MediaPlayer1.pause
        GoTo cowboy
    End If
    
cowboy:

    If popy = "23" Then
        GoTo bnaba
    End If
    
lilstuff.Picture = lpause.Picture
bnaba:
End Sub

Private Sub Picture3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
volie = X
End Sub

Private Sub Picture3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Mouse up on the VOLUME_Box...
'Changes volume depending on where the
'Cursor clicked_up from...

If X > volknob.Left Then
    volknob.Left = volie - 208
    GoTo hell
End If

volknob.Left = volie
hell:

If volie > 930 Then
    volknob.Left = 780
End If

If volie < 61 Then
    volknob.Left = 60
End If

If volie < 50 Then
    volie = 5
End If
'Change the volume..
MediaPlayer1.Volume = "-" & (1000 - volie)

'Make the Volume_Percentage...
volpercent = volknob.Left / 730 * 100 - 10

If volpercent <= 0 Then
    volpercent = "0"
End If

volpercent = Format(volpercent, "#")

If volpercent = "" Then
    volpercent = "0"
End If

'Say the % Volume..
lblName.Caption = "VOLUME: " & volpercent & "%"

End Sub

Private Sub Picture4_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
Form1.PopupMenu mnu
End If
End Sub

Private Sub Picture5_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
Form1.PopupMenu mnu
End If
End Sub

Private Sub play_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Mouse move on Play_Button...

Mplay = True

If Button > 0 Then

    If Y > 285 Then
        play2.Visible = False
        Mplay = False
    End If

    If Y < 30 Then
        play2.Visible = False
        Mplay = False
    End If

    If X < 0 Then
        play2.Visible = False
        Mplay = False
    End If

    If X > 360 Then
        play2.Visible = False
        Mplay = False
    End If

End If
End Sub

Private Sub position_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
clicked = True
End Sub

Private Sub position_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Mouse move on the Song position_knob..

'Change the Icon the Position_drag pic..
If clicked = True Then
    position.MouseIcon = pclick.Picture
    position.MousePointer = vbCustom
End If

If clicked = False Then
    position.MouseIcon = pmove.Picture
    position.MousePointer = vbCustom
End If


Moo2 = X / Form1.Width * 100
Oom2 = Moo2 / 100 * Picture3.Width / 1

If Button > 0 Then
    '
End If

End Sub

Private Sub poslong_Click()
'CLick on the Position Box..

' Position thing ==  0  -->  3780

blub = posX / 3780 * 100 ' = blub
'      10   %  (5)         = 2
'      2    *  (5)         = 10
'      blub *  3780 * 100  = posX

pick = InStr(blub, ".")
If pick = 2 Then
    blub = Left(blub, 1)
End If
If pick = 3 Then
    blub = Left(blub, 2)
End If
If pick = 4 Then
    blub = Left(blub, 3)
End If

If position.Left < 0 Then
    position.Left = 0
End If
If position.Left > (3780 - 440) Then
    position.Left = (3780 - 440)
End If

MediaPlayer1.CurrentPosition = Format((blub * MediaPlayer1.Duration) / 100, "##0")

If MediaPlayer1.CurrentPosition > 0.5 Then
    mpercent.Caption = Format((MediaPlayer1.CurrentPosition / MediaPlayer1.Duration) * 100, "##0") & "%"
    dod = Format(((MediaPlayer1.CurrentPosition) / (MediaPlayer1.Duration)) * 100, "##0")
    position.Left = (dod / 100 * 3780) / 1.133  '- 220 '3475 - 125
End If


'position.Left = blub * 3780 * 100
'...............................
'END CLICKING THE POSITION THING
'...............................
End Sub

Private Sub poslong_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Mouse move on Position_bar...
movenm.Visible = True
lblName.Visible = False
'Change icon..
poslong.MouseIcon = pmove.Picture
poslong.MousePointer = vbCustom
posX = X
End Sub

Private Sub right1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Mouse move on Right_Button..

Mright = True

If Button > 0 Then

    If Y > 285 Then
        right2.Visible = False
        Mright = False
    End If

    If Y < 30 Then
        right2.Visible = False
        Mright = False
    End If

    If X < 0 Then
        right2.Visible = False
        Mright = False
    End If

    If X > 360 Then
        right2.Visible = False
        Mright = False
    End If

End If
End Sub



Private Sub stop0_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Mouse move on Stop_Button...

Mstop = True

If Button > 0 Then

    If Y > 285 Then
        stop2.Visible = False
        Mstop = False
    End If

    If Y < 30 Then
        stop2.Visible = False
        Mstop = False
    End If

    If X < 0 Then
        stop2.Visible = False
        Mstop = False
    End If

    If X > 360 Then
        stop2.Visible = False
        Mstop = False
    End If

End If
End Sub

Private Sub Timer1_Timer()
'Timer one...
'*Makes Position_Knob Move..
'*Calculates Time Left
'*Calculates the % Left..

'A Increasing Variable..
ballie = ballie + 1


If opened = False Then
    GoTo adios
End If

If movenm.Left < (0 - movenm.Width) Then
    movenm.Left = 2280
    GoTo pr4nk
End If

movenm.Left = movenm.Left - 200
pr4nk:


If InStr(MediaPlayer1.CurrentPosition, ".") = 2 Then
    First = Left(MediaPlayer1.CurrentPosition, 1)
End If

If InStr(MediaPlayer1.CurrentPosition, ".") = 3 Then
    First = Left(MediaPlayer1.CurrentPosition, 2)
End If

If InStr(MediaPlayer1.CurrentPosition, ".") = 4 Then
    First = Left(MediaPlayer1.CurrentPosition, 3)
End If

If First = "" Then GoTo adios

If Len(First) = 2 Then
    mp3time.Caption = "00:" & First
End If

If Len(First) = 1 Then
    mp3time.Caption = "00:0" & First
End If

If MediaPlayer1.CurrentPosition > 60 Then

    slut = First / 60
    slut = Left(slut, 1)
    bang = slut * 60
    stump = First - bang

    If Len(stump) = 1 Then
        mp3time.Caption = "0" & slut & ":0" & stump
    End If

    If Len(stump) = 2 Then
        mp3time.Caption = "0" & slut & ":" & stump
    End If
    
End If

If MediaPlayer1.CurrentPosition > 0.5 Then
    mpercent.Caption = Format((MediaPlayer1.CurrentPosition / MediaPlayer1.Duration) * 100, "##0") & "%"
    dod = Format(((MediaPlayer1.CurrentPosition) / (MediaPlayer1.Duration)) * 100, "##0")
    position.Left = (dod / 100 * 3780) / 1.133  '- 220 '3475 - 125
End If

' Position thing ==  0  -->  3780

Call Stop_Timer
adios:
End Sub

Private Sub Stop_Timer()
'Stopping the timer..

Timer1.Enabled = False
Timer1.Enabled = True
'Retire all Vars..
slut = ""
dos = ""
bang = ""
stump = ""
End Sub

Private Sub volknob_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'The Volume knob Mouse up thing..
'Sheesh this took me ages, basically the same
'as the balance_one, but i done this first.

If Oom > 0 Then
'MsgBox Oom
End If

volknob.Visible = True
volknob2.Visible = False

If Oom > 0 Then
    volknob.Left = volknob.Left + (Oom * 1) '4.6)
End If

If Oom > 55 Then
    volknob.Left = volknob.Left + (Oom * 1) '1.6)
End If

If Oom > 111 Then
    volknob.Left = volknob.Left + (Oom * 1.5) '1.9)
End If

If Oom < 0 Then
    volknob.Left = volknob.Left + (Oom * 5)
End If

If volknob.Left > 800 Then
    volknob.Left = 800
End If
If volknob.Left < 61 Then
    volknob.Left = 70
End If

If Oom > 1 Then
    volknob.Left = volknob.Left
End If

volpercent = volknob.Left / 730 * 100 - 10

If volpercent <= 0 Then
    volpercent = "0"
End If
    volpercent = Format(volpercent, "#")
If volpercent = "" Then
    volpercent = "0"
End If
If volpercent >= "97" Then
    volpercent = "100"
End If

movenm.Visible = False
lblName.Visible = True

lblName.Caption = "VOLUME: " & volpercent & "%"
MediaPlayer1.Volume = "-" & (1000 - volknob.Left)

End Sub


Private Sub volknob_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
volknob.Visible = False
volknob2.Visible = True
End Sub

Private Sub volknob_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Mouse move on Volume_Knob

'Change the Icon..
If Button > 0 Then
    volknob.MouseIcon = pmove.Picture
    volknob.MousePointer = vbCustom
End If

If Button = 0 Then
    volknob.MouseIcon = pmouse.Picture
    volknob.MousePointer = vbCustom
End If

'Annoying stuff, much same as the balance,
'Better notation in balance code, so refer to
'that if you don't get it..

Oom = ""
volknob2.Top = volknob.Top
volknob2.Left = volknob.Left
Moo = X / Form1.Width * 100
Oom = Moo / 100 * Picture3.Width / 1

End Sub

Private Sub play_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
play2.Visible = True
End Sub

Private Sub play_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'mouse up on play_button

If Mplay = False Then Exit Sub

Timer1.Enabled = True
play2.Visible = False
If Paused = False Then
    MediaPlayer1.Open (Filep)
End If

If Paused = True Then
    MediaPlayer1.play
End If

lilstuff.Picture = lplay.Picture
position.Visible = True
End Sub

Private Sub right1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
right2.Visible = True
End Sub

Private Sub right1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
right2.Visible = False
If Mright = False Then Exit Sub
End Sub

Private Sub stop0_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
stop2.Visible = True
End Sub

Private Sub stop0_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Mouse up on stop_button
If Mstop = False Then
    Exit Sub
End If
Stopped = True
Timer1.Enabled = False
stop2.Visible = False
Paused = False
MediaPlayer1.Stop
lilstuff.Picture = lstop.Picture
position.Visible = False
position.Left = 0
movenm.Left = 0
End Sub

Private Sub position_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'mouse up on the Position_Knob..

'Change Icon..
If Y > 110 Or Y < 40 Then
    grr = "POP"
    position.MouseIcon = pmove.Picture
    position.MousePointer = vbCustom
End If


clicked = False

position.MouseIcon = pmove.Picture
position.MousePointer = vbCustom

If grr = "POP" Then
    grr = "dfds"
    GoTo Hell2
End If

If Oom2 < 0 Then
    pick2 = InStr(Oom2, ".")
        If pick2 = 2 Then
            blub2 = Left(Oom2, 1)
        End If
        If pick2 = 3 Then
            blub2 = Left(Oom2, 2)
        End If
        If pick2 = 4 Then
            blub2 = Left(Oom2, 3)
        End If
        If pick2 = 5 Then
            blub2 = Left(Oom2, 4)
        End If

        Oom2 = blub2

'This is all my st00pid trial and error
'stuff, i had to manually calc it all
'In order to move the knob on mouse_up

        If Left(Oom2, 1) = "-" Then
            Oom2 = Right(Oom2, (Len(Oom2) - 1))
        End If

        If Oom2 > 1 Then
            pod = 6.5
        End If

        If Oom2 > 100 Then
            pod = 5.4
        End If

        If Oom2 > 200 Then
         pod = 5
        End If

        If Oom2 > 300 Then
            pod = 5
        End If

        If Oom2 > 400 Then
             pod = 4.55
        End If

        If Oom2 > 450 Then
            pod = 4.50001
        End If

        If Oom2 > 500 Then
            pod = 4.6
        End If

        If Oom2 > 600 Then
            pod = 4.88888888888889
        End If

        If Oom2 > 600 Then
            pod = 5
        End If

        Oom2 = "-" & Oom2
    
        If Oom2 < "0" Then
            position.Left = position.Left + (Oom2 * pod)
        End If

    End If

'''''''''''''''''''''''''''''''

If Oom2 > 0 And Oom2 < 300 Then
    position.Left = position.Left + Oom2 * 3
End If

If Oom2 > 299 And Oom2 < 400 Then
    position.Left = position.Left + Oom2 * 3.6
End If

If Oom2 > 399 And Oom2 < 500 Then
    position.Left = position.Left + Oom2 * 3.9
End If

If Oom2 > 499 And Oom2 < 900 Then
    position.Left = position.Left + Oom2 * 3.9
End If

If position.Left < 0 Then
    position.Left = 0
End If

If position.Left > (3780 - 440) Then
    position.Left = (3780 - 440)
End If
'0 --> 3340

dumb = position.Left / 3340 * 100
blub = dumb
pick = InStr(blub, ".")
If pick = 2 Then
    blub = Left(blub, 1)
End If
If pick = 3 Then
    blub = Left(blub, 2)
End If
If pick = 4 Then
    blub = Left(blub, 3)
End If

dumb = blub

MediaPlayer1.CurrentPosition = Format((dumb * MediaPlayer1.Duration) / 100, "##0")

If MediaPlayer1.CurrentPosition > 0.5 Then
    mpercent.Caption = Format((MediaPlayer1.CurrentPosition / MediaPlayer1.Duration) * 100, "##0") & "%"
    dod = Format(((MediaPlayer1.CurrentPosition) / (MediaPlayer1.Duration)) * 100, "##0")
    position.Left = (dod / 100 * 3780) / 1.133  '- 220 '3475 - 125
End If


Hell2:
End Sub
