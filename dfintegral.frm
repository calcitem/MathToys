VERSION 5.00
Begin VB.Form dfintegral 
   Caption         =   "数值积分"
   ClientHeight    =   4725
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8400
   Icon            =   "dfintegral.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   4725
   ScaleWidth      =   8400
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox Txtfx 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   2040
      TabIndex        =   0
      ToolTipText     =   "被积函数,自变量只能是x."
      Top             =   170
      Width           =   5775
   End
   Begin VB.TextBox txteps 
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
      Left            =   2880
      TabIndex        =   4
      Text            =   "15"
      ToolTipText     =   "例如,""15"" 表示精确到10^(-15)"
      Top             =   2950
      Width           =   615
   End
   Begin VB.TextBox txtn 
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
      Left            =   5160
      TabIndex        =   5
      Text            =   "8"
      ToolTipText     =   "分割子区间的步数.取值越大,计算越精确,用时越长"
      Top             =   3015
      Width           =   615
   End
   Begin VB.CommandButton di 
      Caption         =   "计 算"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6360
      TabIndex        =   3
      ToolTipText     =   "单击此处启动或终止计算"
      Top             =   3720
      Width           =   1215
   End
   Begin VB.TextBox Txt 
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
      Left            =   4920
      TabIndex        =   6
      ToolTipText     =   "积分值,若计算尚未完成,则以#开头"
      Top             =   1500
      Width           =   2895
   End
   Begin VB.TextBox txtb 
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
      Left            =   1920
      TabIndex        =   2
      ToolTipText     =   "积分上限.(允许用数学表达式表示)"
      Top             =   1080
      Width           =   615
   End
   Begin VB.TextBox txta 
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
      Left            =   1920
      TabIndex        =   1
      ToolTipText     =   "积分下限,(允许用数学表达式表示)"
      Top             =   1920
      Width           =   615
   End
   Begin VB.Label Lab 
      BackStyle       =   0  'Transparent
      Caption         =   "f(x)="
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   2
      Left            =   720
      TabIndex        =   11
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Lab 
      BackStyle       =   0  'Transparent
      Caption         =   "步数:"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   1
      Left            =   4320
      TabIndex        =   10
      Top             =   3015
      Width           =   615
   End
   Begin VB.Label Lab 
      BackStyle       =   0  'Transparent
      Caption         =   "精度:"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   0
      Left            =   2040
      TabIndex        =   9
      Top             =   3000
      Width           =   735
   End
   Begin VB.Label Lab 
      BackStyle       =   0  'Transparent
      Caption         =   "f(x) dx ="
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   10
      Left            =   2640
      TabIndex        =   8
      Top             =   1440
      Width           =   2295
   End
   Begin VB.Label Lb 
      Caption         =   "∫"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   72
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Index           =   0
      Left            =   480
      TabIndex        =   7
      ToolTipText     =   "积分号"
      Top             =   960
      Width           =   1335
   End
End
Attribute VB_Name = "dfintegral"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Running As Boolean
Private Sub di_Click()
On Error GoTo lr3:

If Running = False Then
  Running = True
  di.Caption = "停 止"
Else
  Running = False
  di.Caption = "计 算"
  'GoTo lr4:
End If

'Dim e, n As long
'Dim a, b, eps, integral As Double
'Dim t(0 to 1024)
'Dim d, d1, f, l, u, u1, y As Double
Dim a, b, ar, t, fa, fm, fb, da, sx, t1, sum, f1, eps As Double
Dim i As Long
Dim dx(0 To 1024)
Dim epsp(0 To 1024) As Double
Dim x2(0 To 1024) As Double
Dim x3(0 To 1024) As Double
Dim f2(0 To 1024) As Double
Dim f3(0 To 1024) As Double
Dim f4(0 To 1024) As Double
Dim fmp(0 To 1024) As Double
Dim fbp(0 To 1024) As Double
Dim t2(0 To 1024) As Double
Dim t3(0 To 1024) As Double
Dim rt(1 To 100) As Double
Dim v(1 To 100, 1 To 3) As Double
Dim n As Long



'n = Val(txtn.Text)
eps = 10 ^ (-Abs(Val(txteps.Text)))

a = Fc(Bracket(translate(txta.Text)), 0, 0)
b = Fc(Bracket(translate(txtb.Text)), 0, 0)
aop$ = Txtfx.Text
n = Abs(Val(txtn.Text))
If n > 100 Then n = 8


errmsg = ExpChk(aop$)
If errmsg <> "" Then
  msg$ = MsgBox(errmsg & Chr(13) & Chr(13) & "您要中止计算吗?", vbYesNo + vbQuestion + vbDefaultButton1, "错误")
  If msg = 6 Then Exit Sub
End If

Do Until InStr(aop$, "x") = 0
   aop$ = Left(aop$, InStr(aop$, "x") - 1) + "(V)" + Right(aop$, Len(aop$) - InStr(aop$, "x"))
Loop
aoo$ = Bracket(translate(aop$))
aop$ = aoo$


  
i = 0
t = 1
ar = 1
da = b - a
fa = Fc(aop$, a, 0)
fm = Fc(aop$, (a + b) / 2, 0) * 4
fb = Fc(aop$, b, 0)
r: i = i + 1
dx(i) = da / 3
sx = dx(i) / 6
f1 = Fc(aop$, a + dx(i) / 2, 0) * 4
x2(i) = a + dx(i)
f2(i) = Fc(aop$, x2(i), 0)
x3(i) = x2(i) + dx(i)
f3(i) = Fc(aop$, x3(i), 0)
epsp(i) = eps
f4(i) = Fc(aop$, x3(i) + dx(i) / 2, 0) * 4
fmp(i) = fm
t1 = (fa + f1 + f2(i)) * sx
fbp(i) = fb
t2(i) = (f2(i) + f3(i) + fm) * sx
t3(i) = (f3(i) + f4(i) + fb) * sx
sum = t1 + t2(i) + t3(i)
'DoEvents
Txt.Text = "# " & sum
If Running = False Then GoTo lr4:
az = ar - Abs(t) + Abs(t1) + Abs(t2(i)) + Abs(t3(i))
If Abs(t - sum) <= epsp(i) * ar And t <> 1 Or i >= n Then
u:  i = i - 1
    v(i, rt(i)) = sum
    Select Case rt(i)
      Case 1
      GoTo l1
      Case 2
      GoTo l2
      Case 3
      GoTo l3
      Case 4
      GoTo r:
      Case 5
      GoTo u:
    End Select
    
End If
rt(i) = 1
da = dx(i)
fm = f1
fb = f2(i)
eps = epsp(i) / 1.7
t = t1
GoTo r:
l1: rt(i) = 2
da = dx(i)
fa = f2(i)
fm = fmp(i)
fb = f3(i)
eps = epsp(i) / 1.7
t = t2(i)
a = x2(i)
GoTo r:
l2: rt(i) = 3
da = dx(i)
fa = f3(i)
fm = f4(i)
fb = fbp(i)
eps = epsp(i) / 1.7
t = t3(i)
a = x3(i)
GoTo r:
l3: sum = v(i, 1) + v(i, 2) + v(i, 3)
If i > 1 Then GoTo u:
Txt.Text = sum





'l = b - a
'e = 1
'm = 1
'd1 = 10 ^ 15

'aoo$ = aop$
'x = a
'u1 = Fc(aoo$, x, 0)

'aoo$ = aop$
'x = b
'u1 = Fc(aoo$, x, 0) + u1
't(1) = u1 / 2

'For i = 1 To n
'  If Running = False Then Exit For
'  y = t(1)
'  u = 0
'  For j = 1 To m + m - 1 Step 2
'    aoo$ = aop$
'    x = a + j * l / m / 2
'    u = Fc(aoo$, x, 0) + u
'    DoEvents
'    If Running = False Then Exit For
'  Next j
'  t(i + 1) = (u / m + t(i)) / 2
'  f = 1
'  For j = i To 1 Step -1
'    f = 4 * f
'    t(j) = t(j + 1) + (t(j + 1) - t(j)) / (f - 1)
'    If Running = False Then Exit For
'  Next j
'  d = Abs(t(1) - y) 'd = Abs(t(i) - y)
'  If d < eps Then
'    e = 0
'    GoTo lr1:
'  End If
'  If d < d1 Then
'    d1 = d
'  Else
'    e = -1
'    integral = l * y
'    GoTo lr2:
'  End If
'  m = m + m
'  DoEvents: Txt.Text = "# " & t(i) * l
'Next i
'lr1:
'integral = l * t(1)
'lr2:

'Txt.Text = integral
'If Running = False Then Txt.Text = "# " & Txt.Text

lr3:
If err <> 0 Then
  If err = 6 Then msg = MsgBox("溢出。", vbExclamation, "错误")
  If err = 11 Then msg = MsgBox("除数为零。", vbExclamation, "错误")
  If err <> 6 And err <> 11 Then msg = MsgBox("未知错误" & " #" & err.Number, vbExclamation, "错误")
  If msg = 6 Then Resume lr4
  If err = 11 Then Resume Next
  
End If
lr4:

Running = False: di.Caption = "计 算"
End Sub


Private Sub txtfx_KeyDown(keycode As Integer, Shift As Integer)
 ShiftDown = (Shift And vbShiftMask) > 0
 altdown = (Shift And vbAltMask) > 0
 CtrlDown = (Shift And vbCtrlMask) > 0

 Select Case keycode
   Case 83
   If ShiftDown Then SendKeys "{BACKSPACE}": Txtfx.Text = Txtfx.Text + "sin": T1sf
   Case 79
   If ShiftDown Then SendKeys "{BACKSPACE}": Txtfx.Text = Txtfx.Text + "cot": T1sf
   Case 88
   If ShiftDown Then SendKeys "{BACKSPACE}": Txtfx.Text = Txtfx.Text + "exp": T1sf
   Case 84
   If ShiftDown Then SendKeys "{BACKSPACE}": Txtfx.Text = Txtfx.Text + "tan": T1sf
   Case 76
   If ShiftDown Then SendKeys "{BACKSPACE}": Txtfx.Text = Txtfx.Text + "log": T1sf
   Case 67
   If ShiftDown Then SendKeys "{BACKSPACE}": Txtfx.Text = Txtfx.Text + "cos": T1sf
   Case 65
   If ShiftDown Then SendKeys "{BACKSPACE}": Txtfx.Text = Txtfx.Text + "arc": T1sf
   Case 69
   If ShiftDown Then SendKeys "{BACKSPACE}": Txtfx.Text = Txtfx.Text + "[e]": T1sf
End Select
End Sub

Private Sub T1sf()
Txtfx.SelStart = Len(Txtfx.Text): Txtfx.SetFocus
End Sub
