VERSION 5.00
Begin VB.Form Playlist 
   BorderStyle     =   0  'None
   Caption         =   "Form4"
   ClientHeight    =   3120
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4200
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3120
   ScaleWidth      =   4200
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   3135
      Left            =   0
      Picture         =   "Playlist.frx":0000
      ScaleHeight     =   3135
      ScaleWidth      =   4215
      TabIndex        =   0
      Top             =   0
      Width           =   4215
      Begin VB.Label Label1 
         BackColor       =   &H80000007&
         Caption         =   "Not Functional (yet) Feel free to do it yaself!"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   495
         Left            =   840
         TabIndex        =   1
         Top             =   960
         Width           =   2655
      End
   End
End
Attribute VB_Name = "Playlist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
