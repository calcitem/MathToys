VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Speed 
   Caption         =   "Speed"
   ClientHeight    =   1935
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7155
   Icon            =   "Speed.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1935
   ScaleWidth      =   7155
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   480
      Top             =   1320
   End
   Begin VB.CommandButton Start 
      Caption         =   "Start"
      Default         =   -1  'True
      Height          =   450
      Left            =   4338
      TabIndex        =   1
      ToolTipText     =   "Press Z and M to start. "
      Top             =   1320
      Width           =   800
   End
   Begin MSComctlLib.ProgressBar B 
      Height          =   495
      Left            =   0
      TabIndex        =   2
      Top             =   480
      Width           =   6000
      _ExtentX        =   10583
      _ExtentY        =   873
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin MSComctlLib.ProgressBar A 
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6000
      _ExtentX        =   10583
      _ExtentY        =   873
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Label ta 
      Alignment       =   1  'Right Justify
      Caption         =   "00.00"
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6000
      TabIndex        =   5
      Top             =   120
      Width           =   975
   End
   Begin VB.Label tb 
      Alignment       =   1  'Right Justify
      Caption         =   "00.00"
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6000
      TabIndex        =   4
      Top             =   600
      Width           =   975
   End
   Begin VB.Label Timework 
      Alignment       =   1  'Right Justify
      Caption         =   "00.00"
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   375
      Left            =   5760
      TabIndex        =   3
      Top             =   1320
      Width           =   1215
   End
End
Attribute VB_Name = "Speed"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private tm, begina, beginb, f, l, m, n, wait, stp

'Private Sub form_Unload(cancel As Integer)
'main.WindowState = 0
'End Sub

Private Sub start_click()
Timer1.Enabled = False
Timework.Caption = "Ready"
Beep: wait = 1: stp = 0
ta.Caption = "": tb.Caption = ""
'Timework = "Ready"
a.Value = 0: B.Value = 0
ready = Timer
1:  'Timework.Caption = Timer - ready
DoEvents
If Timer - ready < 2 Then GoTo 1 Else GoTo 2:
'2: Timework.Caption = "Go"
'If Timer - ready < 4 Then GoTo 2 Else GoTo 3:
'3:
2: Beep: wait = 0

If stp = 0 Then
begina = 1: beginb = 1
tm = Timer
Timer1.Enabled = True
End If
End Sub

Private Sub start_KeyDown(keycode As Integer, shift As Integer)

If keycode = vbKeyZ Then
If wait = 1 Then msg = MsgBox("Player 1 ", , "Transgressor"): ta.Caption = "Loser": stp = 1
If begina = 1 Then f = Timer
End If

If keycode = vbKeyM Then
If wait = 1 Then msg = MsgBox("Player 2 ", , "Transgressor"): tb.Caption = "Loser": stp = 1
If beginb = 1 Then m = Timer
End If
End Sub

Private Sub start_KeyUp(keycode As Integer, shift As Integer)
If keycode = vbKeyZ Then
  l = Timer - f
  If a.Value < 100 And l < 0.16 Then a.Value = a.Value + 1
  If a.Value = 100 And begina = 1 Then
    ta.Caption = Fix((Timer - tm) * 100) / 100
    begina = 0
    'Beep
  End If
End If

If keycode = vbKeyM Then
n = Timer - m
If B.Value < 100 And n < 0.16 Then B.Value = B.Value + 1
  If B.Value = 100 And beginb = 1 Then
    tb.Caption = Fix((Timer - tm) * 100) / 100
    beginb = 0
    'Beep
  End If
End If
If a.Value = 100 And B.Value = 100 Then
  Timer1.Enabled = False
  TC! = Val(tb.Caption) - Val(ta.Caption)
  Timework.Caption = TC!
  If InStr(Timework.Caption, ".") = 1 Then Timework.Caption = "0" + Timework.Caption
  If InStr(Timework.Caption, "-.") = 1 Then Timework.Caption = "-0" + Right(Timework.Caption, Len(Timework.Caption) - 1)
  If InStr(ta.Caption, ".") = 0 Then ta.Caption = ta.Caption + ".00"
  If InStr(tb.Caption, ".") = 0 Then tb.Caption = tb.Caption + ".00"
  If Len(ta.Caption) = InStr(ta.Caption, ".") + 1 Then ta.Caption = ta.Caption + "0"
  If Len(tb.Caption) = InStr(tb.Caption, ".") + 1 Then tb.Caption = tb.Caption + "0"
End If
End Sub

Private Sub Timer1_Timer()
Timework.Caption = Fix((Timer - tm) * 100) / 100
End Sub
