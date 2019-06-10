VERSION 5.00
Begin VB.Form series 
   Caption         =   "总和与连乘"
   ClientHeight    =   6015
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6870
   Icon            =   "SERIES.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6015
   ScaleWidth      =   6870
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox n2 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   345
      Left            =   3240
      TabIndex        =   9
      Text            =   "1"
      Top             =   4630
      Width           =   615
   End
   Begin VB.TextBox n1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   345
      Left            =   3240
      TabIndex        =   4
      Text            =   "1"
      Top             =   2470
      Width           =   615
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   1920
      TabIndex        =   0
      ToolTipText     =   "通项公式"
      Top             =   240
      Width           =   4215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "连乘"
      Height          =   495
      Left            =   4680
      TabIndex        =   8
      ToolTipText     =   "单击这里 启动/停止 计算"
      Top             =   4560
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "总和"
      Height          =   495
      Left            =   4680
      TabIndex        =   3
      ToolTipText     =   "单击这里 启动/停止 计算"
      Top             =   2400
      Width           =   1215
   End
   Begin VB.TextBox pi2 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   345
      Left            =   1200
      TabIndex        =   7
      ToolTipText     =   "初值"
      Top             =   4490
      Width           =   615
   End
   Begin VB.TextBox pi1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   345
      Left            =   960
      TabIndex        =   6
      ToolTipText     =   "终值(不填则表示正无穷大)"
      Top             =   3240
      Width           =   615
   End
   Begin VB.TextBox sgm2 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   345
      Left            =   1200
      TabIndex        =   2
      ToolTipText     =   "初值"
      Top             =   2450
      Width           =   615
   End
   Begin VB.TextBox sgm1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   345
      Left            =   960
      TabIndex        =   1
      ToolTipText     =   "终值(不填则表示正无穷大)"
      Top             =   1080
      Width           =   615
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   3240
      TabIndex        =   10
      Top             =   3840
      Width           =   2895
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   3240
      TabIndex        =   5
      Top             =   1680
      Width           =   2895
   End
   Begin VB.Label Lab 
      BackStyle       =   0  'Transparent
      Caption         =   "步长:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   12
      Left            =   2640
      TabIndex        =   26
      Top             =   4680
      Width           =   495
   End
   Begin VB.Label Lab 
      BackStyle       =   0  'Transparent
      Caption         =   "步长:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   11
      Left            =   2640
      TabIndex        =   25
      Top             =   2520
      Width           =   495
   End
   Begin VB.Label Lab 
      BackStyle       =   0  'Transparent
      Caption         =   "="
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   10
      Left            =   1440
      TabIndex        =   24
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Lab 
      BackStyle       =   0  'Transparent
      Caption         =   "x"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   9
      Left            =   1080
      TabIndex        =   23
      Top             =   360
      Width           =   255
   End
   Begin VB.Label Lab 
      BackStyle       =   0  'Transparent
      Caption         =   "a"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   8
      Left            =   720
      TabIndex        =   22
      Top             =   120
      Width           =   375
   End
   Begin VB.Label lb 
      Height          =   255
      Left            =   240
      TabIndex        =   21
      Top             =   5640
      Width           =   6495
   End
   Begin VB.Label Lab 
      BackStyle       =   0  'Transparent
      Caption         =   "x"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   7
      Left            =   2280
      TabIndex        =   20
      Top             =   3960
      Width           =   255
   End
   Begin VB.Label Lab 
      BackStyle       =   0  'Transparent
      Caption         =   "x"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   6
      Left            =   2280
      TabIndex        =   19
      Top             =   1800
      Width           =   255
   End
   Begin VB.Label Lab 
      BackStyle       =   0  'Transparent
      Caption         =   "a"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   5
      Left            =   1920
      TabIndex        =   18
      Top             =   3720
      Width           =   375
   End
   Begin VB.Label Lab 
      BackStyle       =   0  'Transparent
      Caption         =   "a"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   4
      Left            =   1920
      TabIndex        =   17
      Top             =   1560
      Width           =   375
   End
   Begin VB.Label Lab 
      BackStyle       =   0  'Transparent
      Caption         =   "x="
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   720
      TabIndex        =   16
      Top             =   4440
      Width           =   375
   End
   Begin VB.Label Lab 
      BackStyle       =   0  'Transparent
      Caption         =   "x="
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   720
      TabIndex        =   15
      Top             =   2400
      Width           =   375
   End
   Begin VB.Label Lab 
      BackStyle       =   0  'Transparent
      Caption         =   "="
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   0
      Left            =   2760
      TabIndex        =   14
      Top             =   3720
      Width           =   615
   End
   Begin VB.Label Lab 
      BackStyle       =   0  'Transparent
      Caption         =   "="
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   3
      Left            =   2640
      TabIndex        =   13
      Top             =   1560
      Width           =   615
   End
   Begin VB.Label lbpi 
      Caption         =   "∏"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   42
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   840
      TabIndex        =   12
      Top             =   3600
      Width           =   975
   End
   Begin VB.Label Lbsgm 
      Caption         =   "∑"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   42
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   840
      TabIndex        =   11
      Top             =   1440
      Width           =   975
   End
End
Attribute VB_Name = "series"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Running, running2 As Boolean
Private Sub Command1_Click()
On Error GoTo l1:


If Running = False Then Running = True Else Running = False
Lb.Caption = ""
Text1.Text = ""
j = (Val(sgm2.Text))
k = (Val(sgm1.Text))
n = (Val(n1.Text))
If sgm1.Text = "" Then k = 1.79769 * 10 ^ 308
If j > k And n >= 0 Then
  temp = k: k = j: j = temp
End If
aop$ = Text3.Text

errmsg = ExpChk(aop$)
If errmsg <> "" Then
  msg$ = MsgBox(errmsg & Chr(13) & Chr(13) & "您要中止计算吗?", vbYesNo + vbQuestion + vbDefaultButton1, "问题")
  If msg = 6 Then Exit Sub
End If

Do Until InStr(aop$, "x") = 0
   aop$ = Left(aop$, InStr(aop$, "x") - 1) + "(V)" + Right(aop$, Len(aop$) - InStr(aop$, "x"))
Loop
aoo$ = Bracket(translate(aop$))

For i = j To k Step n
  If Running = False Then Exit For
  aop$ = aoo$
  x = i
  s = s + Fc(aop$, x, 0)
DoEvents
  Text1.Text = s
Next i
Text1.Text = s
l1:
If err <> 0 Then
Text1.Text = "#" + Text1.Text
If err = 6 Then Lb.Caption = "当x=" & Str(i) & "时溢出。": Resume l2 Else Lb.Caption = "当x=" & Str(i) & "时表达式无意义。": Resume Next
End If
l2: Running = False: s = 0
End Sub

Private Sub Command2_Click()
On Error GoTo l1:

If running2 = False Then running2 = True Else running2 = False
Lb.Caption = ""
Text2.Text = ""
s = 1
j = (Val(pi2.Text))
k = (Val(pi1.Text))
n = (Val(n2.Text))

If pi1.Text = "" Then k = 1.79769 * 10 ^ 308

If j > k And n >= 0 Then
  temp = k: k = j: j = temp
End If
aop$ = Text3.Text

errmsg = ExpChk(aop$)
If errmsg <> "" Then
  msg$ = MsgBox(errmsg & Chr(13) & Chr(13) & "您要中止计算吗?", vbYesNo + vbQuestion + vbDefaultButton1, "问题")
  If msg = 6 Then Exit Sub
End If

Do Until InStr(aop$, "x") = 0
   aop$ = Left(aop$, InStr(aop$, "x") - 1) + "(V)" + Right(aop$, Len(aop$) - InStr(aop$, "x"))
Loop
aoo$ = Bracket(translate(aop$))

For i = j To k Step n
If running2 = False Then Exit For
  DoEvents

  aop$ = aoo$
  x = i
  s = s * Fc(aop$, x, 0)
DoEvents
  Text2.Text = s
Next i
Text2.Text = s
l1:
If err <> 0 Then
Text2.Text = "#" + Text2.Text
If err = 6 Then Lb.Caption = "当 x =" & Str(i) & " 时溢出。": Resume l2 Else Lb.Caption = "当 x = " & Str(i) & " 时表达式无意义。": Resume Next
End If
l2: running2 = False: s = 1
End Sub

Private Sub text3_KeyDown(keycode As Integer, Shift As Integer)
 ShiftDown = (Shift And vbShiftMask) > 0
 altdown = (Shift And vbAltMask) > 0
 CtrlDown = (Shift And vbCtrlMask) > 0

 Select Case keycode
   Case 83
   If ShiftDown Then SendKeys "{BACKSPACE}": Text3.Text = Text3.Text + "sin": T1sf
   Case 79
   If ShiftDown Then SendKeys "{BACKSPACE}": Text3.Text = Text3.Text + "cot": T1sf
   Case 88
   If ShiftDown Then SendKeys "{BACKSPACE}": Text3.Text = Text3.Text + "exp": T1sf
   Case 84
   If ShiftDown Then SendKeys "{BACKSPACE}": Text3.Text = Text3.Text + "tan": T1sf
   Case 76
   If ShiftDown Then SendKeys "{BACKSPACE}": Text3.Text = Text3.Text + "log": T1sf
   Case 67
   If ShiftDown Then SendKeys "{BACKSPACE}": Text3.Text = Text3.Text + "cos": T1sf
   Case 65
   If ShiftDown Then SendKeys "{BACKSPACE}": Text3.Text = Text3.Text + "arc": T1sf
   Case 69
   If ShiftDown Then SendKeys "{BACKSPACE}": Text3.Text = Text3.Text + "[e]": T1sf
End Select
End Sub

Private Sub T1sf()
Text3.SelStart = Len(Text3.Text): Text3.SetFocus
End Sub
