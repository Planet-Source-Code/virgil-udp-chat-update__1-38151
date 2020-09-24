VERSION 5.00
Begin VB.Form frmlogin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Net Login..."
   ClientHeight    =   1140
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3030
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmlogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1140
   ScaleWidth      =   3030
   StartUpPosition =   2  'CenterScreen
   WindowState     =   1  'Minimized
   Begin VB.Timer Timer1 
      Interval        =   5000
      Left            =   900
      Top             =   315
   End
   Begin VB.CommandButton cmdexit 
      Cancel          =   -1  'True
      Caption         =   "&Exit"
      Height          =   360
      Left            =   1980
      TabIndex        =   3
      Top             =   630
      Width           =   855
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "&Ok"
      Default         =   -1  'True
      Height          =   360
      Left            =   90
      TabIndex        =   2
      Top             =   630
      Width           =   855
   End
   Begin VB.TextBox txtname 
      Height          =   270
      Left            =   105
      MaxLength       =   10
      TabIndex        =   0
      Top             =   315
      Width           =   2775
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Nick Name:"
      Height          =   255
      Left            =   105
      TabIndex        =   1
      Top             =   75
      Width           =   975
   End
End
Attribute VB_Name = "frmlogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim userid As Integer

Private Sub cmdexit_Click()
End
End Sub

Private Sub cmdok_Click()
Open App.Path & "\text\username.txt" For Output As #1
Print #1, txtname

Close #1
If txtname = "" Then
MsgBox "Please Input a Name first!", vbOKOnly, "Warnning"
Exit Sub
End If
Nick = txtname
Me.Hide
frmmain.Show
End Sub



Private Sub Form_Load()

On Error GoTo onerror
Dim a As String
Open App.Path & "\text\username.txt" For Input As #1
Line Input #1, a
txtname = a
Close #1

GoTo done
onerror:
txtname = ""
If txtname = "" Then Timer1.Enabled = False
Me.WindowState = vbNormal
Me.Visible = True
done:
Close #1
End Sub

Private Sub Form_Resize()
If Me.WindowState = vbNormal Then Timer1.Enabled = False
End Sub

Private Sub Timer1_Timer()

cmdok_Click
Timer1.Enabled = False

End Sub
