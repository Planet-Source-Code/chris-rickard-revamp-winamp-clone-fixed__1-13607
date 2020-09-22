Attribute VB_Name = "Module1"

Public Type NOTIFYICONDATA
    cbSize As Long
    hwnd As Long
    uId As Long
    uFlags As Long
    uCallBackMessage As Long
    hIcon As Long
    szTip As String * 64
    End Type
    'constants required by Shell_NotifyIcon
    '     API call:
    Public Const NIM_ADD = &H0
    Public Const NIM_MODIFY = &H1
    Public Const NIM_DELETE = &H2
    Public Const NIF_MESSAGE = &H1
    Public Const NIF_ICON = &H2
    Public Const NIF_TIP = &H4
    Public Const WM_MOUSEMOVE = &H200
    Public Const WM_LBUTTONDOWN = &H201 'Button down
    Public Const WM_LBUTTONUP = &H202 'Button up
    Public Const WM_LBUTTONDBLCLK = &H203 'Double-click
    Public Const WM_RBUTTONDOWN = &H204 'Button down
    Public Const WM_RBUTTONUP = &H205 'Button up
    Public Const WM_RBUTTONDBLCLK = &H206 'Double-click


Public Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long


Public Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
    Public nid As NOTIFYICONDATA
    
Global Mpause As Boolean
Global Mleft As Boolean
Global Mstop As Boolean
Global Mplay As Boolean
Global Mright As Boolean
Global Meject As Boolean
Global Stopped As Boolean
Global ballie As Integer
Global HaveTime As Boolean
Global ispop As Boolean
Global Frmup As Boolean
Global bla As String
Global bla2 As String
Global Filename2 As String
Global clicked As Boolean
Global volie As Integer
Global Filen As String
Global Filep As String
Global Paused As Boolean
Global Moo As String
Global Oom As String
Global gun As String
Global opened As Boolean
Global slut As String
Global posX As String
Global Moo2 As String
Global Oom2 As String
Global pod As Integer
Global grr As String

Global mMpeg As String
Global mSize As String
Global mLength As String
Global mBit As String
Global mFrames As String
Global mHz As String
Global songie As String
