Attribute VB_Name = "modsnd"
Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Global Const SND_ASYNC = &H1
Global Const SND_NODEFAULT = &H2
Global Const SND_FLAG = SND_ASYNC Or SND_NODEFAULT
Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean

Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Nick As String
Public Declare Function SetForegroundWindow _
Lib "user32" (ByVal hwnd As Long) As Long
' Constants used to detect clicking on the icon
Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_RBUTTONUP = &H205

' Constants used to control the icon
Public Const NIM_ADD = &H0
Public Const NIM_MODIFY = &H1
Public Const NIF_MESSAGE = &H1
Public Const NIM_DELETE = &H2
Public Const NIF_ICON = &H2
Public Const NIF_TIP = &H4

' Used as the ID of the call back message
Public Const WM_MOUSEMOVE = &H200

' Used by Shell_NotifyIcon
Public Type NOTIFYICONDATA
    cbSize As Long
    hwnd As Long
    uId As Long
    uFlags As Long
    ucallbackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type

'create variable of type NOTIFYICONDATA
Public TrayIcon As NOTIFYICONDATA
Dim FileOPen As Boolean

