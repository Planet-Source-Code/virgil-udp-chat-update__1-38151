VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "Net Save"
   ClientHeight    =   3570
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6150
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3570
   ScaleWidth      =   6150
   StartUpPosition =   3  'Windows Default
   Begin Net.chameleonButton chameleonButton1 
      Height          =   240
      Left            =   990
      TabIndex        =   2
      Top             =   45
      Width           =   915
      _ExtentX        =   1614
      _ExtentY        =   423
      BTYPE           =   3
      TX              =   "Save As"
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
      MICON           =   "Form1.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   2835
      Top             =   1530
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DefaultExt      =   "Log"
      Filter          =   ".Log"
   End
   Begin VB.Timer Timer1 
      Interval        =   2000
      Left            =   2475
      Top             =   1530
   End
   Begin Net.chameleonButton cmdsave 
      Height          =   240
      Left            =   90
      TabIndex        =   1
      Top             =   45
      Width           =   780
      _ExtentX        =   1376
      _ExtentY        =   423
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
      MICON           =   "Form1.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.TextBox Text1 
      Height          =   3255
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   315
      Width           =   6135
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub chameleonButton1_Click()
On Error GoTo waserror
Dim filename
cd.ShowSave
filename = cd.filename

Open filename For Output As #3
Print #3, Text1
Close #3
waserror:
End Sub

Private Sub cmdsave_Click()
Open App.Path & "\text\" & "Log" & ".txt" For Output As #2
Print #2, Text1
Close #2
End Sub

Private Sub Form_Load()
Me.Icon = frmmain.Icon

Me.Top = frmmain.Top

Open App.Path & "\text\" & "Log" & ".txt" For Input As #1
Dim a
Do While Not EOF(1)
Line Input #1, a
Text1 = Text1 & a & vbCrLf
Loop
Close #1

End Sub

Private Sub Form_Resize()
If Me.Width > 6270 Then Me.Width = 6270
If Me.Height > 3975 Then Me.Height = 3975
End Sub

Private Sub Timer1_Timer()
Text1.SelStart = Len(Text1.Text)
Timer1.Enabled = False
End Sub
