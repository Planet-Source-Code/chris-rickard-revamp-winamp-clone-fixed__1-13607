VERSION 5.00
Begin VB.Form Form3 
   BackColor       =   &H8000000B&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "About reVAMP"
   ClientHeight    =   5850
   ClientLeft      =   2340
   ClientTop       =   1770
   ClientWidth     =   7035
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5850
   ScaleWidth      =   7035
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Other 
      Height          =   5055
      Left            =   120
      ScaleHeight     =   4995
      ScaleWidth      =   6795
      TabIndex        =   20
      Top             =   720
      Visible         =   0   'False
      Width           =   6855
      Begin VB.PictureBox Picture5 
         BorderStyle     =   0  'None
         Height          =   4455
         Left            =   360
         Picture         =   "Form3.frx":0000
         ScaleHeight     =   4455
         ScaleWidth      =   6135
         TabIndex        =   21
         Top             =   360
         Width           =   6135
      End
   End
   Begin VB.PictureBox History 
      BackColor       =   &H00C0C0C0&
      Height          =   5055
      Left            =   120
      ScaleHeight     =   4995
      ScaleWidth      =   6795
      TabIndex        =   17
      Top             =   720
      Visible         =   0   'False
      Width           =   6855
      Begin VB.TextBox Text1 
         BackColor       =   &H8000000A&
         BorderStyle     =   0  'None
         Height          =   3735
         Left            =   240
         MultiLine       =   -1  'True
         TabIndex        =   19
         Text            =   "Form3.frx":4D82E
         Top             =   1320
         Width           =   5415
      End
      Begin VB.PictureBox Picture4 
         BorderStyle     =   0  'None
         Height          =   1215
         Left            =   120
         Picture         =   "Form3.frx":4DC1B
         ScaleHeight     =   1215
         ScaleWidth      =   3975
         TabIndex        =   18
         Top             =   120
         Width           =   3975
      End
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00000000&
      Height          =   5055
      Left            =   120
      ScaleHeight     =   4995
      ScaleWidth      =   6795
      TabIndex        =   6
      Top             =   720
      Visible         =   0   'False
      Width           =   6855
      Begin VB.PictureBox Picture3 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   6855
         Left            =   1080
         ScaleHeight     =   6855
         ScaleWidth      =   5775
         TabIndex        =   7
         Top             =   3720
         Width           =   5775
         Begin VB.Label Label13 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ARTWORK     Chris Rickard"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   1800
            MousePointer    =   10  'Up Arrow
            TabIndex        =   16
            Top             =   3120
            Width           =   2775
         End
         Begin VB.Label Label11 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "PRODUCT TESTING      Brooke Street"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   840
            MousePointer    =   10  'Up Arrow
            TabIndex        =   15
            Top             =   3960
            Width           =   3795
         End
         Begin VB.Label Label10 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "WHY DID I DO THIS?       Good Question"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   660
            MousePointer    =   10  'Up Arrow
            TabIndex        =   14
            Top             =   4800
            Width           =   4065
         End
         Begin VB.Label Label9 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sorry but reVAMP was the best title i could think of"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   360
            MousePointer    =   10  'Up Arrow
            TabIndex        =   13
            Top             =   5640
            Width           =   4755
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "THIS WILL BE NCFDSA"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   975
            Left            =   960
            TabIndex        =   12
            Top             =   0
            Width           =   3255
         End
         Begin VB.Label Label6 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "PRODUCTION AND DESIGN    Chris Rickard"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   0
            MousePointer    =   10  'Up Arrow
            TabIndex        =   11
            Top             =   1320
            Width           =   4545
         End
         Begin VB.Label Label7 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "CREW     Chris Rickard"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   2280
            MousePointer    =   10  'Up Arrow
            TabIndex        =   10
            Top             =   2160
            Width           =   2265
         End
         Begin VB.Label Label8 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   " Brooke Street"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   3240
            MousePointer    =   10  'Up Arrow
            TabIndex        =   9
            Top             =   2445
            Width           =   1365
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackColor       =   &H00FF8080&
            BackStyle       =   0  'Transparent
            Caption         =   "Over and out... Game Over"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   1440
            TabIndex        =   8
            Top             =   6360
            Width           =   2505
         End
      End
      Begin VB.Timer Timer1 
         Interval        =   30
         Left            =   4080
         Top             =   600
      End
   End
   Begin VB.PictureBox revamp 
      BackColor       =   &H80000007&
      Height          =   5055
      Left            =   120
      Picture         =   "Form3.frx":5B80D
      ScaleHeight     =   4995
      ScaleWidth      =   6795
      TabIndex        =   5
      Top             =   720
      Width           =   6855
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   975
      Left            =   120
      ScaleHeight     =   975
      ScaleWidth      =   6855
      TabIndex        =   0
      Top             =   0
      Width           =   6855
      Begin VB.Line Line3 
         X1              =   5160
         X2              =   5160
         Y1              =   240
         Y2              =   720
      End
      Begin VB.Line Line2 
         X1              =   3240
         X2              =   3240
         Y1              =   240
         Y2              =   720
      End
      Begin VB.Line Line1 
         X1              =   1440
         X2              =   1440
         Y1              =   240
         Y2              =   720
      End
      Begin VB.Label Label4 
         Caption         =   "       Other"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5280
         TabIndex        =   4
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "       History"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3480
         TabIndex        =   3
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "    Credits"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1800
         TabIndex        =   2
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "   reVamp"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1215
      End
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Timer1_Timer()
If Picture3.Top < Picture2.Height - Picture2.Height - Picture3.Height Then
    Picture3.Top = Picture3.Height - 1
    
    Picture3.Top = Label12.Top - 10
    
Else
    Picture3.Top = Picture3.Top - 10
    
End If

End Sub










Private Sub Form_Load()

Label5.Caption = "              reVAMP [tm]     " & vbCrLf & "_______________________" & vbCrLf & "  Copyright © 2000 - rydesoft"


Label1.FontBold = True
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.ForeColor = vbBlue
End Sub

Private Sub Label1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.ForeColor = vbBlack
Label1.FontBold = True
Label2.FontBold = False
Label3.FontBold = False
Label4.FontBold = False


revamp.Visible = True
Picture2.Visible = False
History.Visible = False
Other.Visible = False

End Sub
Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.ForeColor = vbBlue

End Sub

Private Sub Label2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.ForeColor = vbBlack
Label2.FontBold = True
Label1.FontBold = False
Label3.FontBold = False
Label4.FontBold = False
Picture3.Top = 3720
History.Visible = False
Other.Visible = False
revamp.Visible = False
Picture2.Visible = True
End Sub
Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label3.ForeColor = vbBlue
End Sub

Private Sub Label3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label3.ForeColor = vbBlack
Label3.FontBold = True
Label1.FontBold = False
Label2.FontBold = False
Label4.FontBold = False

revamp.Visible = False
Picture2.Visible = False
History.Visible = True
Other.Visible = False
End Sub
Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label4.ForeColor = vbBlue
End Sub

Private Sub Label4_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label4.ForeColor = vbBlack
Label4.FontBold = True
Label1.FontBold = False
Label3.FontBold = False
Label2.FontBold = False

revamp.Visible = False
Picture2.Visible = False
History.Visible = False
Other.Visible = True
End Sub





