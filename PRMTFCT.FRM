VERSION 5.00
Begin VB.Form prmtfct 
   Caption         =   "参数方程"
   ClientHeight    =   2715
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5670
   Icon            =   "PRMTFCT.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   2715
   ScaleWidth      =   5670
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox tr 
      BackColor       =   &H80000004&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      TabIndex        =   5
      Text            =   "10"
      Top             =   1890
      Width           =   615
   End
   Begin VB.TextBox tl 
      BackColor       =   &H80000004&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      TabIndex        =   4
      Text            =   "-10"
      Top             =   1890
      Width           =   735
   End
   Begin VB.CommandButton Command5 
      DisabledPicture =   "PRMTFCT.frx":08CA
      Height          =   450
      Left            =   3960
      Picture         =   "PRMTFCT.frx":0A14
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "清解析式"
      Top             =   1920
      Width           =   525
   End
   Begin VB.CommandButton Draw 
      BackColor       =   &H8000000A&
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   4680
      MousePointer    =   99  'Custom
      Picture         =   "PRMTFCT.frx":0F9E
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "点击这里 开始绘制图象"
      Top             =   1920
      Width           =   525
   End
   Begin VB.TextBox xpa 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   405
      Left            =   1440
      MousePointer    =   3  'I-Beam
      MultiLine       =   -1  'True
      TabIndex        =   0
      ToolTipText     =   "参数方程输入区 (自变量只能是 t)"
      Top             =   240
      Width           =   3855
   End
   Begin VB.TextBox ypa 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   405
      Left            =   1440
      MousePointer    =   3  'I-Beam
      MultiLine       =   -1  'True
      TabIndex        =   1
      ToolTipText     =   "参数方程输入区 (自变量只能是 t)"
      Top             =   1080
      Width           =   3855
   End
   Begin VB.Label Lab 
      BackStyle       =   0  'Transparent
      Caption         =   "     t ∈(        ,             )"
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
      Index           =   2
      Left            =   600
      TabIndex        =   9
      ToolTipText     =   "定义域可以是表达式"
      Top             =   1920
      Width           =   3015
   End
   Begin VB.Label Lab 
      BackStyle       =   0  'Transparent
      Caption         =   "y ="
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   720
      TabIndex        =   8
      Top             =   1080
      Width           =   615
   End
   Begin VB.Label Lab 
      BackStyle       =   0  'Transparent
      Caption         =   "x ="
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   720
      TabIndex        =   7
      Top             =   240
      Width           =   615
   End
   Begin VB.Label Lab 
      BackStyle       =   0  'Transparent
      Caption         =   "{"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Index           =   0
      Left            =   120
      TabIndex        =   6
      Top             =   360
      Width           =   495
   End
End
Attribute VB_Name = "prmtfct"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Private Const HWND_TOPMOST = -1

Private Const SWP_SHOWWINDOWS = &H40


Private Sub Command5_Click()

xpa.text = ""
ypa.text = ""
xpa.SetFocus

End Sub

Public Sub Draw_Click()
prmt = True
Pic.Show
If InStr(xpa.text, "t") = 0 And InStr(ypa.text, "t") = 0 Then msg = MsgBox("缺少参数 t。", vbExclamation, "参数方程")
Pic.Text1.SetFocus
If Pic.Text1.text = "" Then Pic.Text1.text = "x"
If Pic.ExplicitFun.Checked = False Then
  Pic.ExplicitFun.Checked = True
  Pic.ImplicitFun.Checked = False
End If
Sendkeys "{Enter}"
'prmtfct.Show


End Sub

Private Sub form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If tl.text <> "" Then tl.text = Fc(Bracket(translate(tl.text)), 0, 0)
If tr.text <> "" Then tr.text = Fc(Bracket(translate(tr.text)), 0, 0)
End Sub
Private Sub Form_unLoad(Cancel As Integer)
If Pic.cshsh.Checked = True Then
  Pic.ImplicitFun.Checked = False
  Pic.ExplicitFun.Checked = True
  Pic.cshsh.Checked = False
End If
End Sub

Private Sub Form_Load()
Dim retValue As Long

retValue = SetWindowPos(Me.hwnd, HWND_TOPMOST, Me.CurrentX + 200, Me.CurrentY + 200, 390, 210, SWP_SHOWWINDOWS)
End Sub



Private Sub xpa_KeyDown(keycode As Integer, Shift As Integer)
 ShiftDown = (Shift And vbShiftMask) > 0
 altdown = (Shift And vbAltMask) > 0
 CtrlDown = (Shift And vbCtrlMask) > 0

 Select Case keycode
   Case 83
   If ShiftDown Then Sendkeys "{BACKSPACE}": xpa.text = xpa.text + "sin": T1sf
   Case 79
   If ShiftDown Then Sendkeys "{BACKSPACE}": xpa.text = xpa.text + "cot": T1sf
   Case 88
   If ShiftDown Then Sendkeys "{BACKSPACE}": xpa.text = xpa.text + "exp": T1sf
   Case 84
   If ShiftDown Then Sendkeys "{BACKSPACE}": xpa.text = xpa.text + "tan": T1sf
   Case 76
   If ShiftDown Then Sendkeys "{BACKSPACE}": xpa.text = xpa.text + "log": T1sf
   Case 67
   If ShiftDown Then Sendkeys "{BACKSPACE}": xpa.text = xpa.text + "cos": T1sf
   Case 65
   If ShiftDown Then Sendkeys "{BACKSPACE}": xpa.text = xpa.text + "arc": T1sf
   Case 69
   If ShiftDown Then Sendkeys "{BACKSPACE}": xpa.text = xpa.text + "[e]": T1sf
End Select
End Sub

Private Sub T1sf()
xpa.SelStart = Len(xpa.text): xpa.SetFocus
End Sub
Private Sub ypa_KeyDown(keycode As Integer, Shift As Integer)
 ShiftDown = (Shift And vbShiftMask) > 0
 altdown = (Shift And vbAltMask) > 0
 CtrlDown = (Shift And vbCtrlMask) > 0

 Select Case keycode
   Case 83
   If ShiftDown Then Sendkeys "{BACKSPACE}": ypa.text = ypa.text + "sin": T2sf
   Case 79
   If ShiftDown Then Sendkeys "{BACKSPACE}": ypa.text = ypa.text + "cot": T2sf
   Case 88
   If ShiftDown Then Sendkeys "{BACKSPACE}": ypa.text = ypa.text + "exp": T2sf
   Case 84
   If ShiftDown Then Sendkeys "{BACKSPACE}": ypa.text = ypa.text + "tan": T2sf
   Case 76
   If ShiftDown Then Sendkeys "{BACKSPACE}": ypa.text = ypa.text + "log": T2sf
   Case 67
   If ShiftDown Then Sendkeys "{BACKSPACE}": ypa.text = ypa.text + "cos": T2sf
   Case 65
   If ShiftDown Then Sendkeys "{BACKSPACE}": ypa.text = ypa.text + "arc": T2sf
   Case 69
   If ShiftDown Then Sendkeys "{BACKSPACE}": ypa.text = ypa.text + "[e]": T2sf
End Select
End Sub

Private Sub T2sf()
ypa.SelStart = Len(ypa.text): ypa.SetFocus
End Sub
