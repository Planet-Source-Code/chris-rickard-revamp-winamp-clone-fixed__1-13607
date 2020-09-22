VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CFileDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Class       : CFileDialog
' Description : Class for displaying the File Open/Save Common Dialog.
' Source      : Total VB SourceBook 6

Private Type OPENFILENAME
  lStructSize As Long
  hwndOwner As Long
  hInstance As Long
  lpstrFilter As String
  lpstrCustomFilter As String
  nMaxCustFilter As Long
  nFilterIndex As Long
  lpstrFile As String
  nMaxFile As Long
  lpstrFileTitle As String
  nMaxFileTitle As Long
  lpstrInitialDir As String
  lpstrTitle As String
  Flags As Long
  nFileOffset As Integer
  nFileExtension As Integer
  lpstrDefExt As String
  lCustData As Long
  lpfnHook As Long
  lpTemplateName As String
End Type

Private Declare Function GetOpenFileName _
  Lib "comdlg32.dll" _
  Alias "GetOpenFileNameA" _
  (pOpenfilename As OPENFILENAME) _
  As Long

Private Declare Function GetSaveFileName _
  Lib "comdlg32.dll" _
  Alias "GetSaveFileNameA" _
  (pOpenfilename As OPENFILENAME) _
  As Long

Private m_strDefaultExt As String
Private m_strDialogTitle As String
Private m_strFileName As String
Private m_strFileTitle As String
Private m_strInitialDir As String
Private m_strFilter As String
Private m_intFilterIndex As Integer
Private m_eFlags As EnumFilFlags
Private m_intMaxFileSize As Integer
Private m_lnghWndParent As Long

Private Const cintMaxFileLength As Integer = 260

Public Enum EnumFilFlags
  FleReadOnly = &H1
  FleOverWritePrompt = &H2
  FleHideReadOnly = &H4
  FleNoChangeDir = &H8
  FleShowHelp = &H10
  FleEnableHook = &H20
  FleEnableTemplate = &H40
  FleEnableTemplateHandle = &H80
  FleNoValidate = &H100
  FleAllowMultiSelect = &H200
  FleExtensionDifferent = &H400
  FlePathMustExist = &H800
  FleFileMustExist = &H1000
  FleCreatePrompt = &H2000
  FleShareAware = &H4000
  FleNoReadOnlyReturn = &H8000
  FleNoTestFileCreate = &H10000
  FleNoNetworkButton = &H20000
  FleExplorer = &H80000
  FleLongnames = &H200000
End Enum

Private Sub Class_Initialize()
  ' Source: Total VB SourceBook 6
  
  m_intMaxFileSize = cintMaxFileLength
  
End Sub

Public Property Get DefaultExt() As String
  ' Returns: The default extension
  ' Source: Total VB SourceBook 6

  DefaultExt = m_strDefaultExt

End Property

Public Property Let DefaultExt(ByVal strValue As String)
  ' strValue: Set the default extension used for a filename
  ' Source: Total VB SourceBook 6

  m_strDefaultExt = strValue

End Property

Public Property Get DialogTitle() As String
  ' Returns: The title displayed in the dialog
  ' Source: Total VB SourceBook 6

  DialogTitle = m_strDialogTitle

End Property

Public Property Let DialogTitle(ByVal strValue As String)
  ' strValue: Set the title displayed in the dialog
  ' Source: Total VB SourceBook 6

  m_strDialogTitle = strValue

End Property

Public Property Get FileName() As String
  ' Returns: The path and filename
  ' Source: Total VB SourceBook 6

  FileName = m_strFileName

End Property

Public Property Let FileName(ByVal strValue As String)
  ' strValue: Set the filename
  ' Source: Total VB SourceBook 6

  m_strFileName = strValue

End Property

Public Property Get FileTitle() As String
  ' Returns: The filename without the path
  ' Source: Total VB SourceBook 6

  FileTitle = m_strFileTitle

End Property

Public Property Let FileTitle(ByVal strValue As String)
  ' strValue: Set the file title
  ' Source: Total VB SourceBook 6

  m_strFileTitle = strValue

End Property

Public Property Get Filter() As String
  ' Returns: The filter string
  ' Source: Total VB SourceBook 6

  Filter = m_strFilter

End Property

Public Property Let Filter(ByVal strValue As String)
  ' strValue: Set the filter string
  ' Source: Total VB SourceBook 6

  m_strFilter = strValue

End Property

Public Property Get FilterIndex() As Integer
  ' Returns: The index of the filter to display
  ' Source: Total VB SourceBook 6

  FilterIndex = m_intFilterIndex

End Property

Public Property Let FilterIndex(ByVal intValue As Integer)
  ' Set the index of the filter to display
  ' Source: Total VB SourceBook 6

  m_intFilterIndex = intValue

End Property

Public Property Get Flags() As EnumFilFlags
  ' Returns: The flags
  ' Source: Total VB SourceBook 6

  Flags = m_eFlags

End Property

Public Property Let Flags(ByVal eValue As EnumFilFlags)
  ' eValue: Set the flags
  ' Source: Total VB SourceBook 6

  m_eFlags = eValue

End Property

Public Property Get hWndParent() As Long
  ' Returns: The parent hwnd
  ' Source: Total VB SourceBook 6

  hWndParent = m_lnghWndParent

End Property

Public Property Let hWndParent(ByVal lngValue As Long)
  ' lngValue: Set the parent hwnd
  ' Source: Total VB SourceBook 6

  m_lnghWndParent = lngValue

End Property

Public Property Get InitialDir() As String
  ' Returns: The current value of InitialDir
  ' Source : Total VB SourceBook 6

  InitialDir = m_strInitialDir

End Property

Public Property Let InitialDir(ByVal strValue As String)
  ' strValue: Set to the path where the dialog should open
  ' Source  : Total VB SourceBook 6

  m_strInitialDir = strValue

End Property

Public Property Get MaxFileSize() As Integer
  ' Returns: The maximum length of the filename
  ' Source: Total VB SourceBook 6

  MaxFileSize = m_intMaxFileSize

End Property

Public Property Let MaxFileSize(ByVal intValue As Integer)
  ' intValue: Set the maximum length of the filename
  ' Source: Total VB SourceBook 6

  m_intMaxFileSize = intValue

End Property

Public Function Show(fOpen As Boolean) As Boolean
  ' Comments  : This procedure displays the file Open/Save common dialog
  ' Parameters: fOpen - Determines if the Open or Save dialog is displayed.
  ' Returns   : False if cancel selected, true otherwise.
  ' Source    : Total VB SourceBook 6
  ' Update    : Code Service Pack 3
  '

  Dim of As OPENFILENAME
  Dim strChar As String * 1
  Dim intCounter As Integer
  Dim strTemp As String
  
  On Error GoTo PROC_ERR
  
  ' Initialize the OPENFILENAME type
  of.lpstrTitle = m_strDialogTitle & ""
  of.Flags = m_eFlags
  of.lpstrDefExt = m_strDefaultExt & ""
  of.lStructSize = LenB(of)
  of.lpstrFilter = m_strFilter & "||"
  of.nFilterIndex = m_intFilterIndex
  
  ' To make Windows-style filter, replace pipes with nulls
  For intCounter = 1 To Len(m_strFilter)
    strChar = Mid$(m_strFilter, intCounter, 1)
    If strChar = "|" Then
      strTemp = strTemp & vbNullChar
    Else
      strTemp = strTemp & strChar
    End If
  Next
  
  ' Put double null at end
  strTemp = strTemp & vbNullChar & vbNullChar
  of.lpstrFilter = strTemp
  
  ' Pad file and file title buffers to maximum path length
  strTemp = m_strFileName & String$(m_intMaxFileSize - Len(m_strFileName), 0)
  of.lpstrFile = strTemp
  of.nMaxFile = m_intMaxFileSize
  
  strTemp = m_strFileTitle & String$(m_intMaxFileSize - Len(m_strFileTitle), 0)
  of.lpstrFileTitle = strTemp
  of.lpstrInitialDir = m_strInitialDir
  of.nMaxFileTitle = m_intMaxFileSize
  of.hwndOwner = m_lnghWndParent
  
  ' If fOpen is true, show the Open file dialog, otherwise show the Save dialog
  If fOpen Then
    If GetOpenFileName(of) Then
      Show = True
      ' Assign property variables to appropriate values
      m_strFileName = TrimNulls(of.lpstrFile)
      m_strFileTitle = TrimNulls(of.lpstrFileTitle)
    Else
      Show = False
    End If
  Else
    If GetSaveFileName(of) Then
      Show = True
      ' Assign property variables to appropriate values
      m_strFileName = TrimNulls(of.lpstrFile)
      m_strFileTitle = TrimNulls(of.lpstrFileTitle)
    Else
      Show = False
    End If
  End If
  
PROC_EXIT:
  Exit Function
  
PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
  "Show"
  Resume PROC_EXIT

End Function

Private Function TrimNulls(ByVal strIn As String) As String
  ' Comments  : Returns the passed string terminated at the first double null character
  ' Parameters: strIn - Value to parse
  ' Returns   : Parsed string
  ' Source    : Total VB SourceBook 6
  ' Update    : Code Service Pack 3
  '
  Dim intPos As Integer
  Dim strTemp As String
  Dim strTrimNullsRet As String
  Dim strChar As String
  
  Dim intCounter As Integer
  
  On Error GoTo PROC_ERR
  
  ' Find the first occurrence
  intPos = InStr(strIn, vbNullChar)
  
  Do While intPos > 0
    ' Check to see if this is a double null termination
    ' If so, trim the rest of the string
    If (intPos + 1 < Len(strIn)) Then
      If (Mid$(strIn, intPos + 1, 1) = vbNullChar) Then
        Exit Do
      End If
    End If
    ' Find last occurrence
    intPos = InStr(intPos + 1, strIn, vbNullChar)

  Loop
  
  If intPos = 0 Then
    ' No nulls in the string, just return it as is
    TrimNulls = strIn
  Else
    If intPos = 1 Then
      ' If the null character is at the first position, the
      ' entire string is a null string, so return a zero-length string
      strTrimNullsRet = ""
    Else
      ' Not at the first position, so return the contents up
      ' to the occurrence of the null character
      strTrimNullsRet = Left$(strIn, intPos - 1)
    End If
  End If
  
  ' Replace any null characters with a space
  For intCounter = 1 To Len(strTrimNullsRet)
    strChar = Mid$(strTrimNullsRet, intCounter, 1)
    If strChar = vbNullChar Then
      strTemp = strTemp & " "
    Else
      strTemp = strTemp & strChar
    End If
  Next
  
  TrimNulls = strTemp
    
PROC_EXIT:
  Exit Function
  
PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "TrimNulls"
  Resume PROC_EXIT
    
End Function


