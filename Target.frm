VERSION 5.00
Begin VB.Form Target 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   Caption         =   "Shooting"
   ClientHeight    =   6915
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   6855
   Icon            =   "Target.frx":0000
   LinkTopic       =   "Form1"
   MousePointer    =   2  'Cross
   ScaleHeight     =   -3012.929
   ScaleLeft       =   -1500
   ScaleMode       =   0  'User
   ScaleTop        =   1513
   ScaleWidth      =   2986.928
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Draw_load 
      Interval        =   200
      Left            =   10320
      Top             =   2520
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Default         =   -1  'True
      Height          =   105
      Left            =   15120
      TabIndex        =   19
      Top             =   10560
      Width           =   90
   End
   Begin VB.CommandButton Cd 
      BackColor       =   &H00808080&
      Height          =   180
      Left            =   15120
      TabIndex        =   18
      Top             =   10635
      Width           =   135
   End
   Begin VB.TextBox qual 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   465
      Index           =   1
      Left            =   7440
      TabIndex        =   12
      Top             =   5760
      Width           =   688
   End
   Begin VB.TextBox qual 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   465
      Index           =   2
      Left            =   8280
      TabIndex        =   11
      Top             =   5760
      Width           =   688
   End
   Begin VB.TextBox qual 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   465
      Index           =   3
      Left            =   9120
      TabIndex        =   10
      Top             =   5760
      Width           =   688
   End
   Begin VB.TextBox qual 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   465
      Index           =   4
      Left            =   9960
      TabIndex        =   9
      Top             =   5760
      Width           =   688
   End
   Begin VB.TextBox qual 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   465
      Index           =   5
      Left            =   10800
      TabIndex        =   8
      Top             =   5760
      Width           =   688
   End
   Begin VB.TextBox qual 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   465
      Index           =   6
      Left            =   7440
      TabIndex        =   7
      Top             =   6360
      Width           =   688
   End
   Begin VB.TextBox qual 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   465
      Index           =   7
      Left            =   8280
      TabIndex        =   6
      Top             =   6360
      Width           =   688
   End
   Begin VB.TextBox qual 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   465
      Index           =   8
      Left            =   9120
      TabIndex        =   5
      Top             =   6360
      Width           =   688
   End
   Begin VB.TextBox qual 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   465
      Index           =   9
      Left            =   9960
      TabIndex        =   4
      Top             =   6360
      Width           =   688
   End
   Begin VB.TextBox qual 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   465
      Index           =   10
      Left            =   10800
      TabIndex        =   3
      Top             =   6360
      Width           =   688
   End
   Begin VB.CommandButton Begin 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      Caption         =   "Start"
      Height          =   462
      Left            =   4131
      TabIndex        =   0
      TabStop         =   0   'False
      ToolTipText     =   "开始"
      Top             =   0
      Width           =   744
   End
   Begin VB.TextBox Text1 
      Height          =   320
      Left            =   9360
      TabIndex        =   1
      Top             =   1200
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Timer mov 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   9720
      Top             =   2520
   End
   Begin VB.Label allt 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   615
      Left            =   10440
      TabIndex        =   17
      Top             =   4920
      Width           =   735
   End
   Begin VB.Label allm 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      Caption         =   "000.0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9120
      TabIndex        =   16
      Top             =   4800
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Total"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   375
      Left            =   7920
      TabIndex        =   15
      Top             =   4800
      Width           =   855
   End
   Begin VB.Label tot 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      Caption         =   "000.0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9120
      TabIndex        =   14
      Top             =   5280
      Width           =   855
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Final-Score"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   375
      Left            =   7680
      TabIndex        =   13
      Top             =   5280
      Width           =   1455
   End
   Begin VB.Label Mark 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4800
      TabIndex        =   2
      Top             =   6000
      Width           =   1575
   End
   Begin VB.Shape Tar 
      BorderColor     =   &H00808080&
      Height          =   462
      Left            =   3213
      Shape           =   3  'Circle
      Top             =   3080
      Visible         =   0   'False
      Width           =   459
   End
   Begin VB.Shape position 
      BackColor       =   &H00000000&
      FillColor       =   &H00404040&
      FillStyle       =   0  'Solid
      Height          =   465
      Left            =   2520
      Shape           =   3  'Circle
      Top             =   3120
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Menu total 
      Caption         =   "成绩|Score(&S)"
   End
   Begin VB.Menu hard 
      Caption         =   "速度|Speed(&D)"
      Begin VB.Menu fastest 
         Caption         =   "最快|Fastest"
      End
      Begin VB.Menu highest 
         Caption         =   "很快|Faster"
      End
      Begin VB.Menu high 
         Caption         =   "较快|Fast"
      End
      Begin VB.Menu mid 
         Caption         =   "中等|Normal"
      End
      Begin VB.Menu low 
         Caption         =   "较慢|Slow"
      End
      Begin VB.Menu lowest 
         Caption         =   "很慢|Slower"
      End
      Begin VB.Menu slowest 
         Caption         =   "最慢|Slowest"
      End
   End
   Begin VB.Menu Option5 
      Caption         =   "选项| Options(&O)"
      Begin VB.Menu Sound 
         Caption         =   "声音| Sound on(&M)"
      End
      Begin VB.Menu jg 
         Caption         =   "-"
      End
      Begin VB.Menu ws 
         Caption         =   "50发 | 50 Shots"
         Checked         =   -1  'True
      End
      Begin VB.Menu qs 
         Caption         =   "70发 | 70 Shots"
      End
   End
   Begin VB.Menu help 
      Caption         =   "帮助|Help(&H)"
   End
End
Attribute VB_Name = "Target"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private shots As Integer
'Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Private Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" (ByVal lpszName As String, ByVal hModule As Long, ByVal dwflags As Long) As Long

Private tm, after As Long
Private level, go As Single
Private tl
Const SND_ASYNC = 0
Public m$
Private Sub Begin_MouseMove(button As Integer, shift As Integer, x As Single, y As Single)
position.Visible = False
'Tar.Visible = True
Tar.Left = -100
Tar.Top = 100
'Mark.Caption = ""
End Sub

Private Sub Begin_Click()
AutoRedraw = True
Mark.Caption = ""
Tar.Visible = True
position.Visible = False
For c = 1 To 1000 Step 200
Target.Circle (0, 0), c
Next c
mov.Enabled = True
after = after + 1
allt.Caption = Val(allt.Caption) + 1

If after = 11 Then
  
  tl = 0
  after = 1
  For i = 1 To 10
  qual(i).Text = ""
  Next i
  tot.Caption = "Ready"
End If
 AutoRedraw = False

End Sub



Private Sub Draw_load_Timer()
For c = 1 To 1000 Step 200
Target.Circle (0, 0), c
Next c
Draw_load.Enabled = False
End Sub

Private Sub fastest_Click()
level = 3
End Sub

Private Sub help_Click()
msg = MsgBox("1. Set speed and acceleration of your mouse." & Chr(13) & "    设置鼠标移动速度和精确度。" & Chr(13) & "2. Click [Start]." & Chr(13) & "    单击 [Start]。" & Chr(13) & "3. Move the cursor onto the target immediately." & Chr(13) & "    迅速把鼠标移到靶子上。" & Chr(13) & Chr(13) & "名言:" & Chr(13) & "失之毫厘,谬以千里。" & Chr(13) & "善始者实繁,克终者盖寡。" & Chr(13) & "合抱之木,生于毫末; 九层之台,起于累土; ······", , "How to play")
End Sub

Private Sub high_Click()
level = 10
End Sub

Private Sub highest_Click()
level = 5
End Sub

Private Sub low_Click()
level = 20
End Sub

Private Sub lowest_Click()
level = 25
End Sub

Private Sub qs_Click()
If qs.Checked = False Then
  qs.Checked = True
  ws.Checked = False
  shots = 70
Else
  ws.Checked = False
  qs.Checked = True
  shots = 50
End If
End Sub

Private Sub slowest_Click()
level = 30
End Sub

Private Sub mid_Click()
level = 15
End Sub

Private Sub mov_Timer()
wFlags = SND_ASYNC Or snd_nodefault

tm = tm + 1
If tm >= level Then
Tar.Visible = False
position.Visible = True
If Sound.Checked = True Then PlaySound App.Path & "\wav\shoot.wav", 0, SND_ASYNC
If go = 0 Then Mark.Caption = "Miss": qual(after) = "0.0"
tm = 0
'qual(after).SetFocus
Cd.SetFocus
tl = tl + Val(qual(after).Text)
tota$ = Str(tl)
If InStr(tota$, ".") = 0 Then tota$ = tota$ + ".0"
If Val(allt.Caption) > (shots - 10) Then
  allm.Caption = Val(allm.Caption) + Val(qual(after).Text) '*****************
  If InStr(allm.Caption, ".") = 0 Then allm.Caption = allm.Caption + ".0"
Else
  allm.Caption = Val(allm.Caption) + Fix(Val(qual(after).Text)) '*****************
End If
tot.Caption = tota$
go = 0

If Sound.Checked = True Then PlaySound App.Path & "\wav\no.wav", 0, SND_ASYNC
If Sound.Checked = True Then PlaySound App.Path & "\wav\" & Trim(Str(after)) & "b.wav", 0, SND_ASYNC
If Sound.Checked = True Then PlaySound App.Path & "\wav\space.wav", 0, SND_ASYNC

If qual(after).Text <> "" Then
If Sound.Checked = True Then Begin.Enabled = False
If Len(Trim(m$)) = 3 Then
   If Sound.Checked = True Then PlaySound App.Path & "\wav\" & Left(Trim(m$), 1) & "a.wav", 0, SND_ASYNC
Else
If Left(Trim(m$), 2) = "10" Then If Sound.Checked = True Then PlaySound App.Path & "\wav\10a.wav", 0, SND_ASYNC

End If
If Sound.Checked = True Then PlaySound App.Path & "\wav\point.wav", 0, SND_ASYNC
If Sound.Checked = True Then PlaySound App.Path & "\wav\" & Right(Trim(m$), 1) & "b.wav", 0, SND_ASYNC
'DoEvents
If Left(Trim(m$), 2) = "10" Then If Sound.Checked = True Then PlaySound App.Path & "\wav\win.wav", 0, SND_ASYNC
'DoEvents
Mark.Caption = qual(after)
End If
If Trim(qual(after).Text) = "" Then
If Sound.Checked = True Then Begin.Enabled = False: PlaySound App.Path & "\wav\unshot.wav", 0, SND_ASYNC
End If
mov.Enabled = False
DoEvents
If Sound.Checked = True Then: Begin.Enabled = False: PlaySound App.Path & "\wav\space.wav", 0, SND_ASYNC
If Sound.Checked = True Then PlaySound App.Path & "\wav\space.wav", 0, SND_ASYNC
If Sound.Checked = True Then PlaySound App.Path & "\wav\space.wav", 0, SND_ASYNC

If Sound.Checked = True Then PlaySound App.Path & "\wav\start.wav", 0, SND_ASYNC: Begin.Enabled = True
DoEvents
End If

End Sub

Private Sub form_MouseMove(button As Integer, shift As Integer, x As Single, y As Single)
Command1.SetFocus
wFlags = SND_ASYNC Or snd_nodefault
Text1.Text = x & "," & y
Target.PSet (0, 0)
If mov.Enabled = True Then
go = 1
Tar.Visible = True
position.Visible = False
Tar.Left = x - 100
Tar.Top = y + 100



mar = (110 - Fix(Sqr(x ^ 2 + y ^ 2) / 10)) / 10
If mar < 0 Then mar = 0
'If mar = 11 Then mar = 10.9
m$ = Str(mar - 0.1)
If m$ = "-.1" Then m$ = "0.0"
If InStr(m$, ".") = 0 Then m$ = m$ + ".0"
If Left(Trim(m$), 1) = "." Then m$ = "0" + Trim(m$)
'Mark.Caption = m$

qual(after) = m$


position.Left = Tar.Left
position.Top = Tar.Top
End If
If Val(allt.Caption) = shots And mov.Enabled = 0 Then
  msg = MsgBox("    " + allm.Caption, vbOKOnly, "Final score")
  position.Visible = False
  allt.Caption = "0"
  allm.Caption = "0"
End If


End Sub

Private Sub Form_Load()
If ws.Checked = True Then shots = 50 Else shots = 70
level = 17
tl = 0
For c = 0 To 1000 Step 200
Target.Circle (0, 0), c
Next c
'Mark.Caption = "Ready"
'wflags = snd_async Or snd_nodefault
'if sound.checked=true then PlaySound App.Path & "\start.wav", 0, SND_ASYNC
End Sub
Private Sub Form_unLoad(Cancel As Integer)
'main.WindowState = 0
after = 0
End Sub


Private Sub Sound_Click()
If Sound.Checked = True Then Sound.Checked = False Else Sound.Checked = True
End Sub

Private Sub total_Click()
Target.WindowState = 2
End Sub

Private Sub ws_Click()
If ws.Checked = False Then
  ws.Checked = True
  qs.Checked = False
  shots = 50
Else
  qs.Checked = False
  ws.Checked = True
  shots = 70
End If
End Sub
