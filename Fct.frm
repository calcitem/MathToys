VERSION 5.00
Begin VB.Form Fct 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "一元方程求解器"
   ClientHeight    =   5760
   ClientLeft      =   45
   ClientTop       =   225
   ClientWidth     =   8835
   Icon            =   "Fct.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "Fct.frx":08CA
   ScaleHeight     =   15574.87
   ScaleMode       =   0  'User
   ScaleWidth      =   13092.57
   StartUpPosition =   1  '所有者中心
   Begin VB.PictureBox Guage 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   283
      Picture         =   "Fct.frx":3ABF
      ScaleHeight     =   255
      ScaleMode       =   0  'User
      ScaleWidth      =   130
      TabIndex        =   21
      ToolTipText     =   "求解进度"
      Top             =   3960
      Width           =   1320
      Begin VB.PictureBox glider 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   120
         Picture         =   "Fct.frx":3E57
         ScaleHeight     =   225
         ScaleWidth      =   120
         TabIndex        =   22
         Top             =   15
         Width           =   120
      End
   End
   Begin VB.PictureBox Picture6 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   1417
      Picture         =   "Fct.frx":4001
      ScaleHeight     =   315
      ScaleWidth      =   285
      TabIndex        =   20
      ToolTipText     =   "单击此处(或单击""=0"")清除方程式"
      Top             =   296
      Width           =   285
   End
   Begin VB.PictureBox change 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3855
      Left            =   5640
      Picture         =   "Fct.frx":43CD
      ScaleHeight     =   3855
      ScaleWidth      =   2745
      TabIndex        =   16
      Top             =   1516
      Visible         =   0   'False
      Width           =   2745
      Begin VB.Timer showdog 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   1680
         Top             =   960
      End
      Begin VB.Image dog 
         Height          =   375
         Index           =   1
         Left            =   690
         Picture         =   "Fct.frx":65AF
         Top             =   2835
         Width           =   855
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFFFFF&
         Caption         =   "方程式"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   600
         TabIndex        =   19
         Top             =   1680
         Width           =   735
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFFFFF&
         Caption         =   "步长"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   600
         TabIndex        =   18
         Top             =   1360
         Width           =   735
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "求解精度"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   600
         TabIndex        =   17
         Top             =   1050
         Width           =   735
      End
   End
   Begin VB.PictureBox Picture5 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   345
      Left            =   480
      Picture         =   "Fct.frx":6ABD
      ScaleHeight     =   345
      ScaleWidth      =   345
      TabIndex        =   15
      ToolTipText     =   "查看函数图象以了解实零点的分布情况(请先在右边输入方程)  F5"
      Top             =   518
      Width           =   345
   End
   Begin VB.PictureBox Picture4 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   645
      Left            =   924
      Picture         =   "Fct.frx":7177
      ScaleHeight     =   645
      ScaleWidth      =   645
      TabIndex        =   14
      ToolTipText     =   "停止"
      Top             =   4512
      Width           =   645
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   645
      Left            =   924
      Picture         =   "Fct.frx":7A3F
      ScaleHeight     =   645
      ScaleWidth      =   645
      TabIndex        =   13
      ToolTipText     =   "停止解方程"
      Top             =   4512
      Width           =   645
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C000C0&
      Default         =   -1  'True
      Height          =   180
      Left            =   8760
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   5760
      Width           =   90
   End
   Begin VB.TextBox stp 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000001&
      ForeColor       =   &H80000005&
      Height          =   270
      Left            =   240
      TabIndex        =   2
      Text            =   "0.1"
      ToolTipText     =   "步长 [根(断点)的最小间距]。若不能确定根距, 请把此值设为充分小。如果显示的根的个数比实际的少, 请降低此值。"
      Top             =   1680
      Width           =   1335
   End
   Begin VB.ComboBox precision 
      Appearance      =   0  'Flat
      BackColor       =   &H80000001&
      ForeColor       =   &H80000005&
      Height          =   300
      ItemData        =   "Fct.frx":82B4
      Left            =   240
      List            =   "Fct.frx":82C7
      TabIndex        =   5
      Text            =   "-15"
      ToolTipText     =   "求解精度 [以10的n次幂为单位, 例如, -15 表示精确到 10^(-15)]。如果显示的根的个数比实际的少, 请增加此值, 反之则降低此值。 "
      Top             =   1080
      Width           =   1335
   End
   Begin VB.TextBox Rig 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000001&
      ForeColor       =   &H80000005&
      Height          =   270
      Left            =   240
      TabIndex        =   4
      Text            =   "10"
      ToolTipText     =   "根的区间的端点值"
      Top             =   2760
      Width           =   1335
   End
   Begin VB.TextBox Lef 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000001&
      ForeColor       =   &H80000005&
      Height          =   270
      Left            =   240
      TabIndex        =   3
      Text            =   "-10"
      ToolTipText     =   "根的区间的端点值"
      Top             =   2280
      Width           =   1335
   End
   Begin VB.TextBox Value 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   3375
      Left            =   1795
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      ToolTipText     =   "方程的解"
      Top             =   2040
      Width           =   6924
   End
   Begin VB.TextBox Inf 
      Height          =   975
      Left            =   6000
      MultiLine       =   -1  'True
      TabIndex        =   8
      Top             =   3120
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   645
      Left            =   283
      Picture         =   "Fct.frx":82E4
      ScaleHeight     =   645
      ScaleWidth      =   645
      TabIndex        =   10
      ToolTipText     =   "解方程"
      Top             =   4512
      Width           =   645
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   36
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   758
      Left            =   1856
      TabIndex        =   0
      ToolTipText     =   "在此处输入方程式并回车 (非数字间的乘号 * 可省略, 自变量只能是 x)"
      Top             =   337
      Width           =   5895
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   645
      Left            =   283
      Picture         =   "Fct.frx":8BD0
      ScaleHeight     =   645
      ScaleWidth      =   645
      TabIndex        =   12
      ToolTipText     =   "解方程"
      Top             =   4512
      Width           =   645
   End
   Begin VB.Label ifm 
      BackColor       =   &H00400040&
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   1920
      TabIndex        =   9
      ToolTipText     =   "在设定的求解区间内可能的解的个数"
      Top             =   1800
      Width           =   495
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   36
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   758
      Left            =   8280
      TabIndex        =   7
      ToolTipText     =   "单击此处清除方程式"
      Top             =   337
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "="
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   36
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   758
      Left            =   7800
      TabIndex        =   6
      ToolTipText     =   "单击此处清除方程式"
      Top             =   337
      Width           =   495
   End
End
Attribute VB_Name = "Fct"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private xv, m$, ms#, mr#, mo#, DR, alfa, un, ml#, wf, much, digit, ab$, al$, aoo$, erro, web, timen, pw, askpw, genshu, sstop, Fctexp$
Dim x(0 To 1024)
Dim Root(0 To 1024) As Double
Dim u#(1 To 1024)
Private Sub change_Click()
change.Visible = False: showdog.Enabled = False
End Sub
Private Sub CheckRoot(n)
Dim crt As Single
On Error GoTo rroot:

resBak$ = Value.text

For i = 1 To n
 If Abs(Root(i)) > 1.4E-45 And Abs(Root(i)) < 3.4E+38 Then
   crt = Root(i)
   aoo$ = Fctexp$
   Do Until InStr(aoo$, "(V)") = 0
    aoo$ = Left(aoo$, InStr(aoo$, "(V)") - 1) + "(" + Str(crt) + ")" + Right(aoo$, Len(aoo$) - InStr(aoo$, "(V)") - 2)
   Loop
   Do Until InStr(aoo$, "(v)") = 0
    aoo$ = Left(aoo$, InStr(aoo$, "(v)") - 1) + "(" + Str(crt) + ")" + Right(aoo$, Len(aoo$) - InStr(aoo$, "(v)") - 2)
   Loop
   
   Call Htl
   ny = ms#
   If ny = 0 Then Root(i) = Val(Str(crt))
 End If
Next i

no0 = 0
For i = 1 To n
 If InStr(Str(Root(i)), "E-1") > 0 Then
   no0 = no0 + 1
   j = i
 End If
Next i
If no0 = 1 Then
   xv = 0
   aoo$ = Fctexp$
   Call Htl
   ny = ms#
   If ny = 0 Then Root(j) = 0
End If

Value.text = ""

For i = 1 To n
 r$ = Str(Root(i))
 If Left(r$, 2) = " ." Then r$ = " 0" & Trim(r$)
 If Left(r$, 2) = "-." Then r$ = "-0" & Right(r$, Len(r$) - 1)
 Value.text = Value.text & r$ & Chr(13) & Chr(10)
Next i
Exit Sub
rroot:


Value.text = resBak$
msg = MsgBox("计算器优化根失败, 请将此问题报告给作者。", vbExclamation + vbOKOnly, "错误")
Resume Next
End Sub
Private Sub Command1_Click()
On Error GoTo nxt

If Text1.text = "" Then Exit Sub
If InStr(Text1.text, "y") > 0 Then
    msg$ = MsgBox("不支持变量y.", , "错误")
    Exit Sub
End If
errmsg = ExpChk(Text1.text)
If errmsg <> "" Then
  msg$ = MsgBox(errmsg & Chr(13) & Chr(13) & "您要中止计算吗?", vbYesNo + vbQuestion + vbDefaultButton1, "错误")
  If msg = 6 Then Exit Sub
End If
translated$ = translate(Text1.text)
genshu = 0
timen = Timer: gn = 0
Erase x
Value.text = "": ifm.Caption = ""
k = 0
l = Val(Lef.text)
r = Val(Rig.text)
st = Val(stp.text)
If st <= 0 Or st > Abs(l - r) Then st = Abs(l - r)
Inf.text = Inf.text + "第一阶段 确定根的数目" + Chr(13)


aoo$ = translated$
If InStr(aoo$, "=") > 0 Then
  aoo$ = Left(aoo$, InStr(aoo$, "=") - 1) + "-(" + _
  Right(aoo$, Len(aoo$) - InStr(aoo$, "=")) + ")"
End If

aop$ = aoo$
    Do Until InStr(aop$, "x") = 0
      aop$ = Left(aop$, InStr(aop$, "x") - 1) + "(V)" + Right(aop$, Len(aop$) - InStr(aop$, "x"))
    Loop
aop$ = Bracket(translate(aop$))

Fctexp$ = aop$

For nx = l To r Step st
    ifm.Caption = Str(Fix(Abs(nx - l) / Abs(r - l) * 100)) + "%"
    glider.Left = 10 + Fix(Abs(nx - l) / Abs(r - l) * 100)
    xv = nx
  
  aoo$ = aop$
  Call Htl
  ny = ms#
  If ny = 0 Then
    Inf.text = Inf.text + "确定" & nx & "是其中一个根" + Chr(13)
    x(k) = nx
    i = k + 1
    x(i) = Fix(r) + 10
    k = k + 2
    GoTo l1
  End If
  If pass = 1 And ly <> 0 And ny <> 0 And Sgn(ny) <> Sgn(ly) Then
    If Sgn(ny) = -1 Then x(k) = nx: i = k + 1: x(i) = lx Else x(k) = lx: i = k + 1: x(i) = nx
    k = k + 2
  End If
l1: lx = nx
   ly = ny
   pass = 1
nxt: If err <> 0 Then Resume l5

 If Timer - timen > 20 Then
  msg$ = "计算超时" & Chr(13) & "请降低求解精度和步长,或是缩小根的求解范围." & Chr(13) & "如果仍有问题, 请从软件作者处获得帮助。" & Chr(13) & "" & Chr(13) & "您要继续计算吗?"
  Style = vbYesNo + vbQuestion + vbDefaultButton1
  goon = MsgBox(msg$, Style, "计算超时")
  If goon = 7 Then timen = Timer: sstop = 1: GoTo l6 Else timen = Timer + 200
End If
DoEvents
If sstop = 1 Then GoTo l6
l5: Next nx
 Inf.text = Inf.text + "至多有" & k / 2 - 1 & "个根" + Chr(13)

pass = 0

On Error GoTo l2
Inf.text = Inf.text + "第二阶段 解方程 " + Chr(13)
For j = 0 To k - 2 Step 2
  a = x(j)
  b = x(j + 1)
  If b = Fix(r) + 10 Then
    Value.text = Value.text + Str(a) + Chr(13) + Chr(10)
    genshu = genshu + 1
    Root(genshu) = a
    
    GoTo l2
  End If
l3:
If Timer - timen > 15 Then
  msg$ = "计算超时" & Chr(13) & "请降低求解精度和步长,或是缩小根的求解范围。" & Chr(13) & "如果仍有问题, 请从软件作者处获得帮助." & Chr(13) & "" & Chr(13) & "您要继续计算吗?"
  Style = vbYesNo + vbQuestion + vbDefaultButton1
  goon = MsgBox(msg$, Style, "计算超时")
  If goon = 7 Then timen = Timer: sstop = 1: GoTo l6 Else timen = Timer + 150: gn = 1
End If
  x0 = (a + b) / 2
  If x0 = 0 Then GoTo l2
  If Val(a) = Val(b) Then GoTo l4
  
  aoo$ = translate(aop$)
   xv = x0: Call Htl: y0 = ms#
  
  
  aoo$ = translate(aop$)
   xv = a: Call Htl: ya = ms#
  
  
  aoo$ = translate(aop$)
   xv = b: Call Htl: yb = ms#
  
' Stop
 If y0 = yb Then GoTo l4
   
  If y0 > 0 And y0 < yb Then
    b = x0
    Else
    If y0 < 0 And y0 > ya Then a = x0 Else GoTo l2 'l2
  End If
  'If (y0 > 0 And y0 > yb) Or (y0 < 0 And y0 < ya) Then
  If Abs(y0) > 10 ^ Val(precision.text) Then GoTo l3
l4:  Value.text = Value.text + Str(x0) + Chr(13) + Chr(10)
genshu = genshu + 1
Root(genshu) = x0
l2: If err <> 0 Then Resume Next
DoEvents
Next j
l6: k = 0
j = 0
Text1.SetFocus
CheckRoot (genshu)
If sstop = 0 And Value.text = "" Then Value.text = "求解完毕,没有结果可显示。" _
& Chr(13): change.Visible = True: showdog.Enabled = True
If sstop = 1 Then Value.text = "操作已取消。": sstop = 0
ifm.Caption = genshu
End Sub

Public Sub Htl()

DR = 1

change.Visible = False: showdog.Enabled = False
10:
al$ = ab$
Erase u#
un = 0:  alfa = 0



11:
ab$ = aoo$

12:
aoo$ = UCase$(aoo$)



'此处可以跳过
'aoo$ = translate(aoo$)



e$ = ""
m$ = ""

If Len(aoo$) = 1 Then
    ml# = ms#
    ms# = Val(aoo$)
    GoTo 400
End If

20:  If InStr(aoo$, "(") = 0 Then GoTo 70

30
If alfa = 0 Then GoTo o:
If Left(aoo$, 3) = "(UN" And Right(aoo$, 1) = ")" And Len(Str$(Val(Right(aoo$, Len(aoo$) - 3)))) + 4 = Len(aoo$) Then
   GoTo 80:
End If
o:
a = Len(aoo$)
bo$ = Right(aoo$, a - InStr(aoo$, "(") + 1)
c$ = Left(aoo$, a - Len(bo$))


p:

mb = InStr(Right(bo$, Len(bo$) - 1), "(")
nb = InStr(Right(bo$, Len(bo$) - 1), ")")

If mb < nb And mb <> 0 Then
    c$ = c$ + Left(bo$, mb)
    bo$ = Right(aoo$, a - Len(c$))
    GoTo p:
Else
    no$ = Left(bo$, InStr(bo$, ")"))
    d$ = Right(aoo$, a - Len(c$) - Len(no$))
    no$ = mid$(no$, 2, Len(no$) - 2)

End If

If InStr(no$, "+") = 0 And InStr(no$, "-") = 0 Then m$ = no$: Call Beta: GoTo 65
If InStr(no$, "*") = 0 And InStr(no$, "/") = 0 And InStr(no$, "\") = 0 And InStr(no$, "|") = 0 And InStr(no$, "^") = 0 And InStr(no$, "@") = 0 Then GoTo 60


40:
a = Len(no$)
b = 32767
If InStr(no$, "+") > 0 Then b = InStr(no$, "+")
If InStr(no$, "-") > 0 And InStr(no$, "-") < b Then b = InStr(no$, "-")
If b = 32767 Then GoTo 60


50:
m$ = Left(no$, b - 1)
If InStr(m$, "^") > 0 Or InStr(m$, "@") > 0 Then Call Beta Else Call alpha
e$ = e$ + m$ + mid$(no$, b, 1)
no$ = Right(no$, a - b)
If InStr(no$, "*") > 0 Or InStr(no$, "/") > 0 Or InStr(no$, "\") > 0 Or InStr(no$, "|") > 0 Or InStr(no$, "^") > 0 Or InStr(no$, "@") > 0 Or InStr(no$, "+") > 0 Or InStr(no$, "-") > 0 Then GoTo 40


60:
m$ = no$
If InStr(m$, "^") > 0 Or InStr(m$, "@") > 0 Then Call Beta Else Call alpha
m$ = e$ + m$
e$ = ""
If InStr(m$, "^") > 0 Or InStr(m$, "@") > 0 Then Call Beta Else Call alpha


65:
no$ = m$
aoo$ = c$ + no$ + d$
GoTo 20


70:
If Len(c$) = 0 And Len(d$) = 0 And alfa > 0 Then GoTo 80
aoo$ = "(" + aoo$ + ")": GoTo 30


80:

ml# = ms#
ms# = mo#
GoTo 400




400:

End Sub
Public Sub alpha()
alfa = alfa + 1
mo# = 0
PI# = 4 * Atn(1)
If Left(m$, 1) <> "-" Then m$ = "+" + m$

Do Until InStr(m$, "+@") = 0
   Mid$(m$, InStr(m$, "+@"), 2) = "2@"
Loop
If Left(m$, 1) <> "+" Then m$ = "+" + m$

a:
a = Len(m$)
ao$ = Right(m$, a - 1)
b = 32767
If InStr(ao$, "+") > 0 Then b = InStr(ao$, "+")
If InStr(ao$, "-") > 0 And InStr(ao$, "-") < b Then b = InStr(ao$, "-")
If InStr(ao$, "*") > 0 And InStr(ao$, "*") < b Then b = InStr(ao$, "*")
If InStr(ao$, "/") > 0 And InStr(ao$, "/") < b Then b = InStr(ao$, "/")
If InStr(ao$, "\") > 0 And InStr(ao$, "\") < b Then b = InStr(ao$, "\")
If InStr(ao$, "|") > 0 And InStr(ao$, "|") < b Then b = InStr(ao$, "|")
If InStr(ao$, "^") > 0 And InStr(ao$, "^") < b Then b = InStr(ao$, "^")
If InStr(ao$, "@") > 0 And InStr(ao$, "@") < b Then b = InStr(ao$, "@")
If b = 32767 Then c$ = Left(m$, 1): Last = 1: no$ = ao$: GoTo b
bo$ = Left(m$, b)
c$ = Left(bo$, 1)
no$ = Right(bo$, b - 1)

b:
If no$ = "V" Then n# = xv: GoTo f:
If no$ = "M" Then n# = ms#: GoTo f:
If no$ = "MR" Then n# = mr#: GoTo f:
If no$ = "ML" Then n# = ml#: GoTo f:

p$ = InsFun.Funname(no$)

If p$ <> "" And p$ <> "!" Then no$ = Right(no$, Len(no$) - Len(p$))
If p$ = "LOG" Then GoTo d:

g:
If InStr(no$, "UN") = 1 Then n# = u#(Val(Right(no$, Len(no$) - 2))) Else n# = Val(no$)
If loga = 1 Then GoTo h:
If loga = 2 Then GoTo i:

d:
If p$ = "UN" Then n# = u#(Val(no$))
If Right(no$, 1) = "!" And p$ <> "!" Then
   s = 1
   For i = 1 To n#
   s = s * i
   Next i
   n# = s
End If

nd# = n# * PI# / 180

If p$ = "" Then GoTo f:

Select Case p$
    Case "ABS"
    n# = Abs(n#)
    Case "SQR"
    n# = Sqr(n#)
    Case "INT"
    n# = Int(n#)
    Case "TRUNC"
    n# = Fix(n#)
    Case "LN"
    n# = Log(n#)
    Case "LNA"
    n# = Log(Abs(n#))
    Case "SIN"

    If DR = 0 Then n# = Sin(nd#) Else n# = Sin(n#)
    Case "COS"
    If DR = 0 Then n# = Sin((90 - n#) * PI# / 180) Else n# = Cos(n#)
    Case "TAN", "TG"
    If DR = 0 Then n# = Tan(nd#) Else n# = Tan(n#)
    Case "ARCSIN", "ASIN"
    If DR = 0 Then n# = (Atn(n# / Sqr(1 - n# ^ 2))) * 180 / PI# Else n# = Atn(n# / Sqr(1 - n# ^ 2))
    Case "ARCCOS", "ACOS"
    If DR = 0 Then n# = (PI# / 2 - Atn(n# / Sqr(1 - n# ^ 2))) * 180 / PI# Else n# = PI# / 2 - Atn(n# / Sqr(1 - n# ^ 2))
    Case "ARCTG", "ATN", "ARCTAN", "ATAN"
    If DR = 0 Then n# = (Atn(n#)) * 180 / PI# Else n# = Atn(n#)
    Case "ARCCTG", "ACOT", "ARCCOT"
    If DR = 0 Then n# = (PI# / 2 - Atn(n#)) * 180 / PI# Else n# = PI# / 2 - Atn(n#)
    Case "ARCSEC", "ASEC"
    n# = 1 / n#
    If DR = 0 Then n# = (PI# / 2 - Atn(n# / Sqr(1 - n# ^ 2))) * 180 / PI# Else n# = PI# / 2 - Atn(n# / Sqr(1 - n# ^ 2))
    Case "ARCCSC", "ACSC"
    n# = 1 / n#
    If DR = 0 Then n# = (Atn(n# / Sqr(1 - n# ^ 2))) * 180 / PI# Else n# = Atn(n# / Sqr(1 - n# ^ 2))
    Case "EP"
    n# = Exp(n#)
    Case "SGN", "SIGN"
    n# = Sgn(n#)
    Case "COT"
    If DR = 0 Then n# = 1 / (Tan(nd#)) Else n# = 1 / (Tan(n#))
    Case "SEC"
    If DR = 0 Then n# = 1 / (Cos(nd#)) Else n# = 1 / (Cos(n#))
    Case "CSC"
    If DR = 0 Then n# = 1 / (Sin(nd#)) Else n# = 1 / (Sin(n#))
    Case "LOG"
    If InStr(no$, "`") > 0 Then
      nao$ = Left(no$, InStr(no$, "`"))
      nno$ = Right(no$, Len(no$) - InStr(no$, "`"))
      no$ = nao$: loga = 1: GoTo g:
    Else: p$ = "LN": GoTo g:
    End If
h:      na# = n#
    no$ = nno$: loga = 2: GoTo g:
i:      nn# = n#
    n# = Log(nn#) / Log(na#)
    loga = 0: GoTo f:
    Case "LG"
    n# = Log(n#) / Log(10)
    Case "SH", "SINH"
    n# = (Exp(n#) - Exp(-n#)) / 2
    Case "CH", "COSH"
    n# = (Exp(n#) + Exp(-n#)) / 2
    Case "TH", "TANH"
    n# = (Exp(n#) - Exp(-n#)) / (Exp(n#) + Exp(-n#))
    Case "CTH", "COTH"
    n# = (Exp(n#) + Exp(-n#)) / (Exp(n#) - Exp(-n#))
    Case "SECH"
    n# = 2 / (Exp(n#) + Exp(-n#))
    Case "CSCH"
    n# = 2 / (Exp(n#) - Exp(-n#))
    Case "ARSH", "ASINH"
    n# = Log(n# + Sqr(n# ^ 2 + 1))
    Case "ARCH", "ACOSH"
    n# = Log(n# + Sqr(n# ^ 2 - 1)): Beep '+
    Case "ARTH", "ATANH"
    n# = (Log((n# + 1) / (1 - n#))) / 2
    Case "ARCTH", "ACOTH"
    n# = (Log((n# + 1) / (n# - 1))) / 2
    Case "ARSECH", "ASECH"
    n# = Arsech(n#) '+
    Case "ARCSCH", "ACSCH" '?
    n# = Log((Sgn(n#) * Sqr(n# ^ 2 + 1) + 1) / n#)
    Case "DMS"
    n# = Fix(n#) + (Fix((n# - Fix(n#)) * 100)) / 60 + (n# * 100 - Fix(n# * 100)) / 36
    Case "DEG"
    n# = Deg(n#)
    Case "ROUND"
    n# = Round(n#)
    Case "!"
        s = 1
        For i = 1 To n#
        s = s * i
        Next i
        n# = s
End Select



f:
nd# = 0
p$ = ""
If c$ = "+" Then mo# = mo# + n#
If c$ = "-" Then mo# = mo# - n#
If c$ = "*" Then mo# = mo# * n#
If c$ = "/" Then mo# = mo# / n#
If c$ = "\" Then mo# = mo# \ n#
If c$ = "|" Then mo# = mo# Mod n#

'If c$ = "^" And mo# < 0 And Abs(n#) > 0 And Abs(n#) < 1 Then
If c$ = "^" And mo# < 0 And Fix(Abs(n#)) <> Abs(n#) Then
If pw = 0 And askpw = 0 Then
  askpw = 1
  msg$ = "当x小于零且n不为整数时, 计算器不能判断 x^n 的正负。" & Chr(13) & "我们对此引起的不便表示抱歉。" & Chr(13) & Chr(13) & "要继续计算, 计算器需要知道 x^n 的符号。" & Chr(13) & Chr(13) & "如果其值为正, 请选择“是”;" & Chr(13) & "如果为负, 请选择“否”;" & Chr(13) & "如果 x^n 没有意义, 请选择“取消”。"
  Sty = vbYesNoCancel + vbQuestion + vbDefaultButton1
  gon = MsgBox(msg$, Sty, "x^n 大于零吗?")
If gon = 6 Then pw = 6
If gon = 7 Then pw = 7
End If
If pw = 6 Then mo# = (-mo#) ^ n#: GoTo j:
If pw = 7 Then mo# = -(-mo#) ^ n#: GoTo j:
End If

'If c$ = "@" And n# < 0 And mo# <> 0 Then
If c$ = "@" And n# < 0 And Parity(mo#) = 0 Then
If Fix(Abs(1 / mo#)) <> Abs(1 / mo#) Then
If pw = 0 And askpw = 0 Then
  askpw = 1
  msg$ = "当x小于零且n的倒数不为整数时, 计算器不能判断 n√x 的正负。" & Chr(13) & "我们对此引起的不便表示抱歉。" & Chr(13) & Chr(13) & "要继续计算, 计算器需要知道 n√x 的符号。" & Chr(13) & Chr(13) & "如果其值为正, 请选择“是”;" & Chr(13) & "如果为负, 请选择“否”;" & Chr(13) & "如果 n√x 没有意义, 请选择“取消”。"
  Sty = vbYesNoCancel + vbQuestion + vbDefaultButton1
  gon = MsgBox(msg$, Sty, "n√x 大于零吗?")
If gon = 6 Then pw = 6
If gon = 7 Then pw = 7
End If
If pw = 6 Then mo# = -(-n#) ^ (1 / mo#): GoTo j:
If pw = 7 Then mo# = -(-n#) ^ (1 / mo#): GoTo j:
End If
End If

If c$ = "@" And n# < 0 And Parity(mo#) <> 0 Then
     If Parity(mo#) = 1 Then mo# = -(-n#) ^ (1 / mo#) Else mo# = Log(-1)
    
     GoTo j:
End If

If c$ = "^" Then mo# = mo# ^ n#
If c$ = "@" Then mo# = n# ^ (1 / mo#)
j:
If Last = 1 Then GoTo e:
m$ = Right(m$, a - b)
GoTo a

e:
Last = 0
un = un + 1
u#(un) = mo#
m$ = "UN" + Str$(un)
End Sub

Public Sub Beta()
f$ = ""
f$ = m$

15:
a = Len(f$)
b = 32767
If InStr(f$, "*") > 0 Then b = InStr(f$, "*")
If InStr(f$, "/") > 0 And InStr(f$, "/") < b Then b = InStr(f$, "/")
If InStr(f$, "\") > 0 And InStr(f$, "\") < b Then b = InStr(f$, "\")
If InStr(f$, "|") > 0 And InStr(f$, "|") < b Then b = InStr(f$, "|")
If b = 32767 Then GoTo 35

25:
m$ = Left(f$, b - 1)
If Len(m$) <> 0 Then Call alpha
e$ = e$ + m$ + mid$(f$, b, 1)
f$ = Right(f$, a - b)
If InStr(f$, "^") > 0 Or InStr(f$, "@") > 0 Then GoTo 15 Else m$ = f$: GoTo 37

35:
m$ = f$
If Len(m$) <> 0 Then Call alpha

37:
m$ = e$ + m$
e$ = ""
Call alpha
End Sub





Private Sub form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Picture2.Visible = True
Picture4.Visible = True
End Sub


Private Sub Form_unLoad(Cancel As Integer)
'main.WindowState = 0
End Sub
Private Sub Label1_Click()
Text1.text = ""
Text1.SetFocus
End Sub

Private Sub Label2_Click()
Text1.text = ""
Text1.SetFocus
End Sub


Private Sub Label3_Click()
precision.text = Val(precision.text) + 1
change.Visible = False: showdog.Enabled = False
Sendkeys "{enter}"

End Sub

Private Sub Label4_Click()
stp.text = Val(stp.text) / 10
change.Visible = False: showdog.Enabled = False
Sendkeys "{enter}"

End Sub

Private Sub Label5_Click()
chg$ = MsgBox("您可以给方程的左边加上或减去一个充分小的数。这个数可适当调整其符号和大小。" & Chr(13) _
& "例如：原方程为 sinx-1=0, 更改后为 sinx-1+0.001=0。原方程为 sinx+1=0, 更改后为 sinx-1-0.001=0。", , "请按下列说明执行")
change.Visible = False: showdog.Enabled = False
End Sub

Private Sub Picture1_Click()
Sendkeys "{enter}"
End Sub

Private Sub Picture4_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Picture4.Visible = False
End Sub

Private Sub Picture2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Picture2.Visible = False
End Sub
Private Sub picture1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Picture4.Visible = True
End Sub
Private Sub Picture3_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Picture2.Visible = True
End Sub

Private Sub Picture3_Click()
sstop = 1
End Sub

Private Sub Picture5_Click()
Pic.Show
Pic.showweb1.Checked = False
Pic.Showscale.Checked = False
Pic.Picture1.Cls
Pic.Text1.text = Text1.text
'Pic.Xinc.Text = stp.Text
'Pic.Yinc.Text = stp.Text
Pic.xmin.text = Lef.text
Pic.xmax.text = Rig.text
Pic.ymin.text = -Abs(Val(Lef.text) - Val(Rig.text)) / 2
Pic.ymax.text = Abs(Val(Lef.text) - Val(Rig.text)) / 2

Sendkeys "{enter}"
End Sub

Private Sub Picture6_Click()
Text1.text = ""
Text1.SetFocus
End Sub

Private Sub showdog_Timer()
If Val(Right(Str(Timer), 1)) Mod 2 = 1 Then
If dog(1).Visible = False Then
    dog(1).Visible = True
    showdog.Interval = 400
Else
    dog(1).Visible = False
    showdog.Interval = 3000
End If
End If
End Sub

Private Sub Text1_Change()
Text1.ToolTipText = Text1.text
End Sub
Private Sub value_KeyDown(keycode As Integer, Shift As Integer)
 Select Case keycode
   Case vbKeyF5
   Call Picture5_Click
End Select
End Sub

Private Sub text1_KeyDown(keycode As Integer, Shift As Integer)
 ShiftDown = (Shift And vbShiftMask) > 0
 altdown = (Shift And vbAltMask) > 0
 CtrlDown = (Shift And vbCtrlMask) > 0

 Select Case keycode
   Case 83
   If ShiftDown Then Sendkeys "{BACKSPACE}": Text1.text = Text1.text + "sin": T1sf
   Case 79
   If ShiftDown Then Sendkeys "{BACKSPACE}": Text1.text = Text1.text + "cot": T1sf
   Case 88
   If ShiftDown Then Sendkeys "{BACKSPACE}": Text1.text = Text1.text + "exp": T1sf
   Case 84
   If ShiftDown Then Sendkeys "{BACKSPACE}": Text1.text = Text1.text + "tan": T1sf
   Case 76
   If ShiftDown Then Sendkeys "{BACKSPACE}": Text1.text = Text1.text + "log": T1sf
   Case 67
   If ShiftDown Then Sendkeys "{BACKSPACE}": Text1.text = Text1.text + "cos": T1sf
   Case 65
   If ShiftDown Then Sendkeys "{BACKSPACE}": Text1.text = Text1.text + "arc": T1sf
   Case 69
   If ShiftDown Then Sendkeys "{BACKSPACE}": Text1.text = Text1.text + "[e]": T1sf
   Case vbKeyF5
   Call Picture5_Click
End Select
End Sub

Private Sub T1sf()
Text1.SelStart = Len(Text1.text): Text1.SetFocus
End Sub
