VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmmain 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Net"
   ClientHeight    =   4710
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6165
   ClipControls    =   0   'False
   DrawStyle       =   5  'Transparent
   Icon            =   "frmmain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MouseIcon       =   "frmmain.frx":0E42
   MousePointer    =   99  'Custom
   Picture         =   "frmmain.frx":114C
   ScaleHeight     =   4710
   ScaleWidth      =   6165
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   WindowState     =   1  'Minimized
   Begin VB.Timer Timer5 
      Enabled         =   0   'False
      Interval        =   300
      Left            =   2475
      Top             =   2115
   End
   Begin Net.chameleonButton Command3 
      Height          =   285
      Left            =   5175
      TabIndex        =   16
      Top             =   4050
      Width           =   825
      _ExtentX        =   1455
      _ExtentY        =   503
      BTYPE           =   3
      TX              =   "Save"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmmain.frx":5285
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Net.chameleonButton Command5 
      Height          =   285
      Left            =   4410
      TabIndex        =   15
      Top             =   4050
      Width           =   690
      _ExtentX        =   1217
      _ExtentY        =   503
      BTYPE           =   3
      TX              =   "Clear"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmmain.frx":52A1
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Net.chameleonButton cmdsend 
      Default         =   -1  'True
      Height          =   285
      Left            =   3510
      TabIndex        =   14
      Top             =   4050
      Width           =   825
      _ExtentX        =   1455
      _ExtentY        =   503
      BTYPE           =   3
      TX              =   "&Send"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmmain.frx":52BD
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Net.chameleonButton Command4 
      Height          =   330
      Left            =   3510
      TabIndex        =   13
      Top             =   405
      Width           =   2625
      _ExtentX        =   4630
      _ExtentY        =   582
      BTYPE           =   3
      TX              =   "Check if selected user is online"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmmain.frx":52D9
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Net.chameleonButton Command2 
      Height          =   285
      Left            =   2790
      TabIndex        =   12
      Top             =   45
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   503
      BTYPE           =   3
      TX              =   "Delete Entry"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmmain.frx":52F5
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Net.chameleonButton Command1 
      Height          =   285
      Left            =   1080
      TabIndex        =   11
      Top             =   45
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   503
      BTYPE           =   3
      TX              =   "Add New Recipient"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmmain.frx":5311
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   0
      Top             =   4950
   End
   Begin VB.Timer Timer3 
      Interval        =   100
      Left            =   2565
      Top             =   4860
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   675
      Top             =   4905
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock wsfile 
      Left            =   585
      Top             =   5625
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemotePort      =   11221
      LocalPort       =   11221
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   300
      Left            =   315
      Picture         =   "frmmain.frx":532D
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   9
      Top             =   4905
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   2475
      Top             =   2475
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   0
      Top             =   5625
   End
   Begin MSWinsockLib.Winsock ws 
      Left            =   5715
      Top             =   4770
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
      LocalPort       =   788
   End
   Begin VB.ListBox List1 
      Height          =   645
      ItemData        =   "frmmain.frx":5368
      Left            =   135
      List            =   "frmmain.frx":536A
      TabIndex        =   6
      Top             =   405
      Width           =   3270
   End
   Begin VB.TextBox txtmain 
      Height          =   2850
      Left            =   135
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   1125
      Width           =   5925
   End
   Begin VB.TextBox txtself 
      Height          =   495
      Left            =   960
      TabIndex        =   4
      Top             =   1890
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.TextBox txtname 
      Height          =   495
      Left            =   840
      TabIndex        =   3
      Top             =   1770
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox txtimg 
      Height          =   495
      Left            =   720
      TabIndex        =   2
      Top             =   1650
      Visible         =   0   'False
      Width           =   1695
   End
   Begin MSWinsockLib.Winsock wskudp 
      Left            =   90
      Top             =   2370
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
      LocalPort       =   787
   End
   Begin VB.TextBox txtsend 
      Height          =   285
      Left            =   135
      TabIndex        =   0
      Top             =   4050
      Width           =   3225
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "View Log"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   195
      Left            =   5265
      MouseIcon       =   "frmmain.frx":536C
      MousePointer    =   99  'Custom
      TabIndex        =   17
      Top             =   4410
      Width           =   825
   End
   Begin VB.Image imgoff 
      Height          =   330
      Left            =   5580
      Picture         =   "frmmain.frx":5676
      Top             =   45
      Width           =   360
   End
   Begin VB.Image imgon 
      Height          =   330
      Left            =   5580
      Picture         =   "frmmain.frx":5CEA
      Top             =   45
      Width           =   360
   End
   Begin VB.Label Label4 
      Caption         =   "Label4"
      Height          =   330
      Left            =   1890
      TabIndex        =   10
      Top             =   4905
      Width           =   465
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3600
      TabIndex        =   8
      Top             =   810
      Width           =   2400
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Send To:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   180
      TabIndex        =   7
      Top             =   45
      Width           =   825
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   225
      Left            =   4080
      TabIndex        =   5
      Top             =   1440
      Width           =   210
   End
   Begin VB.Menu mnuMenu 
      Caption         =   "Menu"
      Visible         =   0   'False
      Begin VB.Menu mnuShow 
         Caption         =   "Show"
      End
      Begin VB.Menu mnusave 
         Caption         =   "Save"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ifsend As Boolean
Dim X, Y As Long
Dim sound As Boolean
'Dim msg As Integer
Dim Saveit As String
Private Reset As Boolean
Dim sending As String
Dim file As String




Private Sub cmdsend_Click()
Label3.Caption = ""
wskudp.RemoteHost = Left(List1.Text, 15)

Dim cmd As String
Dim fFile As Integer
fFile = FreeFile


On Error Resume Next
cmd = txtsend.Text
Select Case cmd
Case ""
    wskudp.SendData wskudp.LocalHostName & " have no word to send but ENTER key!:)" & vbCrLf
    
    
Case Else
    wskudp.SendData Nick & ": " & txtsend.Text
    txtmain.Text = txtmain.Text & Time & ": " & Nick & ": " & txtsend.Text & vbCrLf
    txtmain.SelStart = Len(txtmain.Text)
End Select
txtsend.Text = ""
txtsend.SetFocus
'msg = msg + 1
'StatusBar1.Panels(1).Text = "messengs sent!" & " = " & Str(msg)
Timer4.Enabled = True
End Sub




Private Sub Command1_Click()

Dim newthing As String
Dim newname As String

newthing = InputBox("Enter Users IP Address")
newname = InputBox("Enter Users name")

List1.AddItem newthing & "                " & newname
Dim i
Open App.Path & "\text\users.txt" For Output As #1
List1.ListIndex = -1

For i = 0 To List1.ListCount - 1
List1.ListIndex = List1.ListIndex + 1
Print #1, List1.Text
Next i
Close #1

End Sub

Private Sub Command2_Click()
On Error Resume Next
List1.RemoveItem List1.ListIndex
Dim i
Open App.Path & "\text\users.txt" For Output As #1
List1.ListIndex = -1

For i = 0 To List1.ListCount - 1
List1.ListIndex = List1.ListIndex + 1
Print #1, List1.Text
Next i
Close #1
End Sub
Private Sub Command3_Click()
Dim fFile As Integer
    Open App.Path & "\text\" & "Log" & ".txt" For Append As #1
    Print #1, Date
    Print #1, txtmain.Text
    Close 1
   ' StatusBar1.Panels(1).Text = "File has saved!"
End Sub

Private Sub Command4_Click()
On Error Resume Next
Label3.Caption = "Checking....."
ws.RemoteHost = Left(List1.Text, 12)
ws.SendData "Ping"
Timer1.Enabled = True
End Sub

Private Sub Command5_Click()
txtmain = ""
End Sub



Private Sub Form_Load()
imgoff.Visible = False
sending = "no"
Dim a As String
Open App.Path & "\text\users.txt" For Input As #1
Do While Not EOF(1)
Line Input #1, a
List1.AddItem a
Loop
Close #1


On Error Resume Next
wskudp.RemotePort = 787
ws.RemotePort = 788
wskudp.Bind
ws.Bind

ifsend = False
sound = True
'txtmain.Text = "Welcome --------------------------------->>>>>" & vbCrLf
'wskudp.SendData wskudp.LocalHostName & " enter the Room >>>>>" & vbCrLf
X = frmmain.Width
Y = frmmain.Height
End Sub

Private Sub Form_Resize()
On Error Resume Next
If frmmain.Width <> X <> frmmain.Height <> Y Then
frmmain.Width = X
frmmain.Height = Y
End If
Select Case Me.WindowState

        Case vbMinimized
       App.TaskVisible = False
        Me.Hide
        
        TrayIcon.cbSize = Len(TrayIcon)
    
    ' Handle of the window used to handle messages - which is the this form
    TrayIcon.hWnd = Me.hWnd
    
    ' ID code of the icon
    TrayIcon.uId = vbNull
    
    ' Flags
    TrayIcon.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
    
    ' ID of the call back message
    TrayIcon.ucallbackMessage = WM_MOUSEMOVE
    
    ' The icon - sets the icon that should be used
    TrayIcon.hIcon = Me.Icon
    
    ' The Tooltip for the icon - sets the Tooltip that will be displayed
    TrayIcon.szTip = "Menu" & Chr$(0)
    
    ' Add icon to the tray by calling the Shell_NotifyIcon API
    'NIM_ADD is a Constant - add icon to tray
     Call Shell_NotifyIcon(NIM_ADD, TrayIcon)
     
    Case vbMaximized
     Timer1.Enabled = False
    Call Shell_NotifyIcon(NIM_DELETE, TrayIcon)
    App.TaskVisible = True
            Me.Show
        Case vbNormal
        Timer1.Enabled = False
        Timer5.Enabled = True
        Call Shell_NotifyIcon(NIM_DELETE, TrayIcon)
            Me.Show
            App.TaskVisible = True
    End Select
End Sub

Private Sub Form_Terminate()
Call Shell_NotifyIcon(NIM_DELETE, TrayIcon)
End Sub

Private Sub Form_Unload(Cancel As Integer)
Call Shell_NotifyIcon(NIM_DELETE, TrayIcon)
wskudp.Close
ws.Close
End
End Sub

Private Sub imgoff_Click()
If sound = True Then
sound = False
imgon.Visible = False
imgoff.Visible = True

Else
sound = True
imgon.Visible = True
imgoff.Visible = False

End If
End Sub

Private Sub imgon_Click()
If sound = True Then
sound = False
imgon.Visible = False
imgoff.Visible = True

Else
sound = True
imgon.Visible = True
imgoff.Visible = False

End If
End Sub

Private Sub lblsound_Click()

End Sub

Private Sub Label5_Click()
Form1.Show
End Sub

'Private Sub lblsound_Click()
'If sound = True Then
'sound = False
'lblsound.BackColor = RGB(255, 0, 0)
'lblsound.Caption = "soundOff"
'Else
'sound = True
'lblsound.BackColor = RGB(0, 255, 0)
'lblsound.Caption = "sound on"
'End If
'End Sub




Private Sub List1_Click()
Label3.Caption = ""
Timer1.Enabled = False
End Sub

Private Sub mnuexit_Click()
Call Shell_NotifyIcon(NIM_DELETE, TrayIcon)
ws.Close
wskudp.Close

Unload frmlogin
Unload frmmain

End Sub

Private Sub Timer1_Timer()
If Label3.Caption = "Checking....." Then
Label3.Caption = "User not online"
Else

Label3.Caption = ""
Timer1.Enabled = False
End If
End Sub

Private Sub Timer2_Timer()

If TrayIcon.hIcon = Me.Icon Then
TrayIcon.hIcon = frmlogin.Icon

Shell_NotifyIcon NIM_MODIFY, TrayIcon
Else
TrayIcon.hIcon = Me.Icon
Shell_NotifyIcon NIM_MODIFY, TrayIcon
End If

End Sub

Private Sub Timer3_Timer()
Label4.Caption = wsfile.State
End Sub

Private Sub Timer4_Timer()
Label3.Caption = "Message not received"
Timer4.Enabled = False
End Sub

Private Sub Timer5_Timer()
txtmain.SelStart = Len(txtmain.Text)
Timer5.Enabled = False
End Sub

Private Sub txtsend_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
KeyAscii = 0
cmdsend_Click
End If
End Sub

Private Sub ws_DataArrival(ByVal bytesTotal As Long)
Dim msg As String
On Error Resume Next
ws.GetData msg
If msg = "Ping" Then
ws.RemoteHost = ws.RemoteHostIP
ws.SendData "Pong"
End If
If msg = "Pong" Then
Label3.Caption = "User online"
Timer1.Enabled = True
End If
If msg = "!@" Then
Timer4.Enabled = False
Label3.Caption = "Message received"
End If
End Sub



Private Sub wskudp_DataArrival(ByVal bytesTotal As Long)
On Error Resume Next

Dim incoming As String
Dim cmd As String
wskudp.GetData incoming
If incoming <> "" Then
ws.RemoteHost = wskudp.RemoteHostIP
ws.SendData "!@"
cmd = Mid(incoming, 1, 2)


Select Case cmd

Case Else
'wskudp.GetData incoming

txtmain.Text = txtmain.Text & Time & ": " & incoming & vbCrLf
txtmain.SelStart = Len(txtmain.Text)
If Me.WindowState = vbMinimized Then Timer2.Enabled = True
If sound = True And incoming <> "" Then
Playsound (App.Path & "\sound\msg.wav")
End If

End Select
DoneI:
End If
End Sub
Public Sub Playsound(WavFile As String)
On Error Resume Next
        Call sndPlaySound(WavFile$, SND_FLAG)
End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim result
Static Message As Long
Static RR As Boolean
    
    'x is the current mouse location along the x-axis
    Message = X / Screen.TwipsPerPixelX
    
    If RR = False Then
        RR = True
        Select Case Message
            ' Left double click (This should bring up a dialog box)
            Case WM_LBUTTONDBLCLK
            Me.WindowState = vbNormal
            Timer2.Enabled = False
                Me.Show
            ' Right button up (This should bring up a menu)
            Case WM_RBUTTONUP
            result = SetForegroundWindow(Me.hWnd)

                Me.PopupMenu mnuMenu
        End Select
        RR = False
    End If
End Sub


Private Sub mnushow_click()
Me.WindowState = vbNormal
Me.Show
Timer2.Enabled = False
End Sub

Private Sub mnusave_click()
Dim fFile As Integer
    Open App.Path & "\text\" & "Log" & ".txt" For Append As #1
    Print #1, "Year:" & Date & "|| Time: " & Time
    Print #1, txtmain.Text
    Close 1
    'StatusBar1.Panels(1).Text = "File has saved!"
End Sub

