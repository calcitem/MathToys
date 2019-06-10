VERSION 5.00
Begin VB.Form tri 
   BackColor       =   &H80000000&
   Caption         =   "解三角形"
   ClientHeight    =   7800
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5445
   Icon            =   "TRIANGLE.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   Picture         =   "TRIANGLE.frx":08CA
   ScaleHeight     =   7800
   ScaleWidth      =   5445
   Begin VB.TextBox superficial 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1200
      TabIndex        =   25
      ToolTipText     =   "三角形的面积"
      Top             =   5040
      Width           =   3735
   End
   Begin VB.TextBox radius 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1200
      TabIndex        =   24
      ToolTipText     =   "三角形外接圆半径"
      Top             =   5520
      Width           =   3735
   End
   Begin VB.TextBox rad 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1200
      TabIndex        =   23
      ToolTipText     =   "三角形内接圆半径"
      Top             =   6000
      Width           =   3735
   End
   Begin VB.TextBox ha 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1200
      TabIndex        =   19
      ToolTipText     =   "a边上的高"
      Top             =   3480
      Width           =   3735
   End
   Begin VB.TextBox ma 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1200
      TabIndex        =   18
      ToolTipText     =   "a边上的中线"
      Top             =   3960
      Width           =   3735
   End
   Begin VB.TextBox ta 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1200
      TabIndex        =   17
      ToolTipText     =   "a边上的角分线"
      Top             =   4440
      Width           =   3735
   End
   Begin VB.OptionButton radian 
      BackColor       =   &H80000000&
      Caption         =   "弧度"
      Height          =   255
      Left            =   4080
      TabIndex        =   16
      Top             =   6720
      Width           =   735
   End
   Begin VB.OptionButton dfm 
      BackColor       =   &H80000000&
      Caption         =   "度-分-秒"
      Height          =   255
      Left            =   2400
      TabIndex        =   15
      Top             =   6720
      Width           =   1095
   End
   Begin VB.OptionButton degree 
      BackColor       =   &H80000000&
      Caption         =   "度"
      Height          =   255
      Left            =   1320
      TabIndex        =   14
      Top             =   6720
      Value           =   -1  'True
      Width           =   615
   End
   Begin VB.CommandButton Command2 
      Caption         =   "重置"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      TabIndex        =   7
      Top             =   7200
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "求解"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3720
      TabIndex        =   6
      Top             =   7200
      Width           =   1095
   End
   Begin VB.TextBox pointc 
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1200
      TabIndex        =   5
      Top             =   2760
      Width           =   3735
   End
   Begin VB.TextBox pointb 
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1200
      TabIndex        =   4
      Top             =   2280
      Width           =   3735
   End
   Begin VB.TextBox pointa 
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1200
      TabIndex        =   3
      Top             =   1800
      Width           =   3735
   End
   Begin VB.TextBox linec 
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1200
      TabIndex        =   2
      ToolTipText     =   "边长"
      Top             =   1200
      Width           =   3735
   End
   Begin VB.TextBox lineb 
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1200
      TabIndex        =   1
      ToolTipText     =   "边长"
      Top             =   720
      Width           =   3735
   End
   Begin VB.TextBox linea 
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   1200
      TabIndex        =   0
      ToolTipText     =   "边长"
      Top             =   210
      Width           =   3735
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "S="
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   28
      Top             =   5040
      Width           =   495
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "R="
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   27
      Top             =   5520
      Width           =   495
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "r="
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   26
      Top             =   6000
      Width           =   495
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C0C0C0&
      Caption         =   "m="
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   22
      Top             =   3960
      Width           =   495
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0C0C0&
      Caption         =   "h="
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   21
      Top             =   3480
      Width           =   495
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0C0C0&
      Caption         =   "t="
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   20
      Top             =   4440
      Width           =   495
   End
   Begin VB.Label pc 
      BackColor       =   &H00C0C0C0&
      Caption         =   "C="
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   13
      Top             =   2760
      Width           =   495
   End
   Begin VB.Label pb 
      BackColor       =   &H00C0C0C0&
      Caption         =   "B="
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   12
      Top             =   2280
      Width           =   495
   End
   Begin VB.Label pa 
      BackColor       =   &H00C0C0C0&
      Caption         =   "A="
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   11
      Top             =   1800
      Width           =   495
   End
   Begin VB.Label lc 
      BackColor       =   &H00C0C0C0&
      Caption         =   "c="
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      TabIndex        =   10
      Top             =   1200
      Width           =   375
   End
   Begin VB.Label lb 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "b="
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      TabIndex        =   9
      Top             =   720
      Width           =   375
   End
   Begin VB.Label la 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "a="
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      TabIndex        =   8
      Top             =   240
      Width           =   375
   End
End
Attribute VB_Name = "tri"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub triangle()
On Error GoTo err
Dim a, b, c, pa, pb, pc, s, r, min As Double

PI = 4 * Atn(1)

a = Val(linea.text)
b = Val(lineb.text)
c = Val(linec.text)

If radian.Value = False Then
  pa = Val(pointa.text) * PI / 180
  pb = Val(pointb.text) * PI / 180
  pc = Val(pointc.text) * PI / 180
Else
  pa = Val(pointa.text)
  pb = Val(pointb.text)
  pc = Val(pointc.text)
End If

If dfm.Value = True Then
  pa = Deg(pa)
  pb = Deg(pb)
  pc = Deg(pc)
End If


tl:
If a <> 0 And b <> 0 And c <> 0 Then
  pa = acos((b ^ 2 + c ^ 2 - a ^ 2) / (2 * b * c))
  pb = acos((c ^ 2 + a ^ 2 - b ^ 2) / (2 * c * a))
  pc = acos((a ^ 2 + b ^ 2 - c ^ 2) / (2 * a * b))
  GoTo pri:
End If

If b <> 0 And c <> 0 And pa <> 0 And a = 0 And pb = 0 And pc = 0 Then
  a = Sqr(b ^ 2 + c ^ 2 - 2 * b * c * Cos(pa))
  GoTo tl:
End If

If c <> 0 And a <> 0 And pb <> 0 And b = 0 And pa = 0 And pc = 0 Then
  b = Sqr(c ^ 2 + a ^ 2 - 2 * c * a * Cos(pb))
  GoTo tl:
End If

If a <> 0 And b <> 0 And pc <> 0 And c = 0 And pa = 0 And pb = 0 Then
  c = Sqr(a ^ 2 + b ^ 2 - 2 * a * b * Cos(pc))
  GoTo tl:
End If


If pa <> 0 And pb <> 0 And pc = 0 Then pc = PI - pa - pb
If pa <> 0 And pc <> 0 And pb = 0 Then pb = PI - pa - pc
If pb <> 0 And pc <> 0 And pa = 0 Then pa = PI - pb - pc

If a <> 0 And b = 0 And c = 0 Then
  b = Sin(pb) * a / Sin(pa)
  c = Sin(pc) * a / Sin(pa)
  GoTo pri:
End If

If b <> 0 And a = 0 And c = 0 Then
  a = Sin(pa) * b / Sin(pb)
  c = Sin(pc) * b / Sin(pb)
  GoTo pri:
End If

If c <> 0 And a = 0 And b = 0 Then
  a = Sin(pa) * c / Sin(pc)
  b = Sin(pb) * c / Sin(pc)
  GoTo pri:
End If


ang = MsgBox("要求解的三角形是锐角三角形吗？", 4, "三角形有两解")

If a <> 0 And b <> 0 And pa <> 0 And c = 0 And pb = 0 And pc = 0 Then
  If ang = 6 Then pb = asin(Sin(pa) * b / a) Else pb = PI - asin(Sin(pa) * b / a)
  pc = PI - pa - pb
  c = Sin(pc) * a / Sin(pa)
End If

If a <> 0 And b <> 0 And pb <> 0 And c = 0 And pc = 0 And pa = 0 Then
  If ang = 6 Then pa = asin(Sin(pb) * a / b) Else pa = PI - asin(Sin(pb) * a / b)
  pc = PI - pa - pb
  c = Sin(pc) * b / Sin(pb)
End If

If a <> 0 And c <> 0 And pa <> 0 And b = 0 And pb = 0 And pc = 0 Then
  If ang = 6 Then pc = asin(Sin(pa) * c / a) Else pc = PI - asin(Sin(pa) * c / a)
  pb = PI - pa - pc
  b = Sin(pb) * a / Sin(pa)
End If

If a <> 0 And c <> 0 And pc <> 0 And b = 0 And pa = 0 And pb = 0 Then
  If ang = 6 Then pa = asin(Sin(pc) * a / c) Else pa = PI - asin(Sin(pc) * a / c)
  pb = PI - pa - pc
  b = Sin(pb) * c / Sin(pc)
End If

If b <> 0 And c <> 0 And pb <> 0 And a = 0 And pa = 0 And pc = 0 Then
  If ang = 6 Then pc = asin(Sin(pb) * c / b) Else pc = PI - asin(Sin(pb) * c / b)
  pa = PI - pb - pc
  a = Sin(pa) * b / Sin(pb)
End If

If b <> 0 And c <> 0 And pc <> 0 And a = 0 And pa = 0 And pb = 0 Then
  If ang = 6 Then pb = asin(Sin(pc) * b / c) Else pb = PI - asin(Sin(pc) * b / c)
  pa = PI - pb - pc
  a = Sin(pa) * c / Sin(pc)
End If

pri:
linea.text = a
lineb.text = b
linec.text = c


If degree.Value = True Then
  pointa.text = pa * 180 / PI
  pointb.text = pb * 180 / PI
  pointc.text = pc * 180 / PI
Else

If dfm.Value = True Then
  pointa.text = Dms(pa * 180 / PI)
  pointb.text = Dms(pb * 180 / PI)
  pointc.text = Dms(pc * 180 / PI)
Else
  pointa.text = pa
  pointb.text = pb
  pointc.text = pc
End If
End If
 
If pb < pa Then min = pb Else min = pa
If pc < min Then min = pc

If pa = min Then s = b * c * Sin(pa) * 0.5
If pb = min Then s = c * a * Sin(pb) * 0.5
If pc = min Then s = a * b * Sin(pc) * 0.5

superficial.text = s
radius.text = (a / Sin(pa) + b / Sin(pb) + c / Sin(pc)) / 6
rad.text = s / (0.5 * (a + b + c))

ha.text = b * Sin(pc)
ma.text = 0.5 * Sqr(2 * (b ^ 2 + c ^ 2) - a ^ 2)
ta.text = Sqr(b * c * ((b + c) ^ 2 - a ^ 2)) / (b + c)
GoTo endsub

err:
ok = MsgBox("此三角形无解．", 64, "计算器")
Resume endsub
endsub: End Sub

Private Function asin(no As Double) As Double
PI = 4 * Atn(1)
asin = Atn(no / Sqr(1 - no ^ 2))
End Function

Private Function acos(no As Double) As Double
PI = 4 * Atn(1)
acos = PI / 2 - Atn(no / Sqr(1 - no ^ 2))
End Function

Private Sub Command1_Click()
Call triangle
End Sub

Private Sub Command2_Click()
linea.text = ""
lineb.text = ""
linec.text = ""
pointa.text = ""
pointb.text = ""
pointc.text = ""
superficial.text = ""
radius.text = ""
rad.text = ""
ha.text = ""
ma.text = ""
ta.text = ""
End Sub


Private Sub Form_unLoad(Cancel As Integer)
'main.WindowState = 0
End Sub

