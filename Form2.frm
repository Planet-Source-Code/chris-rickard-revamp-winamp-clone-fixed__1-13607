VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MPEG file info box & ID3 tag editor"
   ClientHeight    =   2895
   ClientLeft      =   2970
   ClientTop       =   2685
   ClientWidth     =   5985
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2895
   ScaleWidth      =   5985
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command3 
      Caption         =   "Remove ID3"
      Height          =   375
      Left            =   2520
      TabIndex        =   16
      Top             =   2400
      Width           =   1095
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1320
      TabIndex        =   15
      Top             =   2400
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Save"
      Height          =   375
      Left            =   120
      TabIndex        =   14
      Top             =   2400
      Width           =   1095
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   960
      TabIndex        =   13
      Top             =   2040
      Width           =   2655
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   2160
      TabIndex        =   12
      Text            =   "Combo1"
      Top             =   1680
      Width           =   1455
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   960
      TabIndex        =   11
      Top             =   1680
      Width           =   495
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   960
      TabIndex        =   10
      Top             =   1320
      Width           =   2655
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   960
      TabIndex        =   9
      Top             =   960
      Width           =   2655
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   960
      TabIndex        =   8
      Top             =   600
      Width           =   2655
   End
   Begin VB.Frame Frame1 
      Caption         =   "MPEG info"
      Height          =   2295
      Left            =   3720
      TabIndex        =   1
      Top             =   480
      Width           =   2175
      Begin VB.Label lblinfo 
         Height          =   1935
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.TextBox Text0 
      BackColor       =   &H8000000A&
      Height          =   285
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   5775
   End
   Begin VB.Label Label6 
      Caption         =   "Genre"
      Height          =   255
      Left            =   1560
      TabIndex        =   7
      Top             =   1695
      Width           =   495
   End
   Begin VB.Label Label5 
      Caption         =   "Comment"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   2040
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "Year"
      Height          =   255
      Left            =   435
      TabIndex        =   5
      Top             =   1680
      Width           =   615
   End
   Begin VB.Label Label3 
      Caption         =   "Album"
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   1320
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   " Artist"
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   960
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "  Title"
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   600
      Width           =   615
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command3_Click()
'Removing the ID3 Tag...

response = MsgBox("Removing the ID3 tag can only be done whilst the song is unloaded." & vbCrLf & "Do you want to unload to song and save the changes?", vbYesNoCancel, "ID3 tag saving")

If response = 7 Then 'no
    GoTo pop
    songie = "nup"
End If

If response = 2 Then 'cancel
    GoTo pop
    songie = "nup"
End If

If response = 6 Then 'yes
    Form1.MediaPlayer1.Stop
    songie = "uhhu"
    Unload Form1
    
    GoTo bangerang
End If
     

' just filling in the information into the type
bangerang:
id3Info.Title = ""
id3Info.Artist = ""
id3Info.Album = ""
id3Info.sYear = ""
id3Info.Comments = ""
id3Info.Genre = "0"

' Calling the Saveid3 function
SaveId3 Filep, id3Info

ErrHandle:
'If Err.Number = 75 Then
'MsgBox "File is Write Protected"
'Else
'MsgBox Err.Description
'End If
pop:

End Sub

Private Sub Command4_Click()
'Cancel click

If songie = "uhhu" Then
    Load Form1
    Form1.Show

    Form1.Timer1.Enabled = True
    Form1.MediaPlayer1.Open (Filep)
    Form1.lilstuff.Picture = Form1.lplay.Picture
End If

Unload Form2

End Sub

Private Sub Form_Load()
'Loading the form..

'If the filename is Null
If Filep = "" Then
    MsgBox "No music file has been loaded", vbOKOnly, "Error reading file"
    Text0.Text = "No music file has been loaded"
    Combo1.Text = ""
    Text1.Enabled = False
    Text1.BackColor = vbScrollBars
    Text2.BackColor = vbScrollBars
    Text3.BackColor = vbScrollBars
    Text4.BackColor = vbScrollBars
    Text5.BackColor = vbScrollBars
    Text2.Enabled = False
    Text3.Enabled = False
    Text4.Enabled = False
    Text5.Enabled = False
    Combo1.BackColor = vbScrollBars
    Combo1.Enabled = False
    Command2.Enabled = False
    Command3.Enabled = False

    Exit Sub

End If
    
    mLength = Format(Form1.MediaPlayer1.Duration, "#")


Filename2 = Filep
Text0.Text = Filep
GenreArray = Split(sGenreMatrix, "|")   ' we fill the array with the Genre's
For i = LBound(GenreArray) To UBound(GenreArray)
Combo1.AddItem GenreArray(i)        ' now fill the Combobox with the array, and voila, the code you
                                    ' you recieve form the Genre part of the Type, represents the combobox Listindex =)
Next

Combo1.ListIndex = 1
GetId3 Filep ' Get the filename
Text1 = RTrim(id3Info.Title)            ' since the fields in the type are
Text2 = RTrim(id3Info.Artist)                  ' fixed lenght, we use Rtrim to cut the
Text3 = RTrim(id3Info.Album)                   ' trailing bytes
Text4 = RTrim(id3Info.sYear)
Text5 = RTrim(id3Info.Comments)
If id3Info.Genre = 255 Then
GoTo bang
End If
Combo1.ListIndex = id3Info.Genre

bang:
'Command2.Enabled = True
      
     ' mFrames = Str(Int(total_frames))
     ' mLength = MediaPlayer1.Duration
     ' mHz = Str(sample_rate)
     ' mMpeg = mpeg1 & " " & layer
     ' mBit = Left(actual_bitrate, 3)

lblinfo.Caption = "Size: " & FileLen(Filep) & vbCrLf
lblinfo.Caption = lblinfo.Caption & "Length: " & mLength & " seconds" & vbCrLf
lblinfo.Caption = lblinfo.Caption & mMpeg & vbCrLf
lblinfo.Caption = lblinfo.Caption & mBit & "bit, " & mFrames & " frames" & vbCrLf
lblinfo.Caption = lblinfo.Caption & mHz & "hz" & " Joint Stereo" & vbCrLf
lblinfo.Caption = lblinfo.Caption & "Private: No" & vbCrLf
lblinfo.Caption = lblinfo.Caption & "CRCs: Yes" & vbCrLf
lblinfo.Caption = lblinfo.Caption & "Copyrighted: No" & vbCrLf
lblinfo.Caption = lblinfo.Caption & "Original: Yes" & vbCrLf
lblinfo.Caption = lblinfo.Caption & "Emphasis: None" & vbCrLf
End Sub

Private Sub Command2_Click()

response = MsgBox("Saving over the ID3 tag can only be done whilst the song is unloaded." & vbCrLf & "Do you want to unload to song and save the changes?", vbYesNoCancel, "ID3 tag saving")
'MsgBox response

If response = 7 Then 'no
GoTo pop
songie = "nup"
End If

If response = 2 Then 'cancel
GoTo pop
songie = "nup"
End If

If response = 6 Then 'yes
Form1.MediaPlayer1.Stop
songie = "uhhu"
Unload Form1
GoTo bangerang
End If
        

        
bangerang:
id3Info.Title = Text1.Text    ' just filling in the information into the type
id3Info.Artist = Text2.Text
id3Info.Album = Text3.Text
id3Info.sYear = Text4.Text
id3Info.Comments = Text5.Text
id3Info.Genre = Combo1.ListIndex

'''MsgBox Filep
'On Error GoTo ErrHandle             ' If the file is writeprotected
SaveId3 Filep, id3Info     ' Calling the Saveid3 function
'GoTo pop


ErrHandle:
'If Err.Number = 75 Then
'MsgBox "File is Write Protected"
'Else
'MsgBox Err.Description
'End If
pop:
End Sub


Private Sub Form_Unload(Cancel As Integer)
'MsgBox songie
If songie = "uhhu" Then
Load Form1
Form1.Show
'opened = True

Form1.Timer1.Enabled = True
Form1.MediaPlayer1.Open (Filep)
Form1.lilstuff.Picture = Form1.lplay.Picture


End If
'Unload Form2

End Sub
