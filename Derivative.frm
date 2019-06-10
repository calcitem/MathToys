VERSION 5.00
Begin VB.Form Der 
   BackColor       =   &H00C0C0C0&
   Caption         =   "显函数求导器"
   ClientHeight    =   4215
   ClientLeft      =   60
   ClientTop       =   480
   ClientWidth     =   4935
   FillStyle       =   0  'Solid
   Icon            =   "Derivative.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   Picture         =   "Derivative.frx":08CA
   ScaleHeight     =   4215
   ScaleWidth      =   4935
   StartUpPosition =   1  '所有者中心
   Begin VB.Timer Timer1 
      Interval        =   200
      Left            =   7080
      Top             =   5280
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Test"
      Height          =   495
      Left            =   5880
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   4320
      Width           =   855
   End
   Begin VB.TextBox fx 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   96
      TabIndex        =   0
      Top             =   105
      Width           =   4795
   End
   Begin VB.CommandButton der 
      Caption         =   "d/dx"
      Default         =   -1  'True
      Height          =   495
      Left            =   5760
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   3600
      Width           =   1335
   End
   Begin VB.TextBox Dfx 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   3270
      Left            =   96
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   620
      Width           =   4795
   End
End
Attribute VB_Name = "Der"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Dfx.Text = DelStr(fx.Text, "/x")
'l = Derivative_PM(fx.Text)
End Sub

Private Sub der_Click()
Select Case fx.Text
  Case "dummy"
    Dfx.Text = fx.Text + "不可导": Exit Sub
End Select

b$ = LCase(Replace(fx.Text, "exp", "ep"))
b$ = Replace(b$, "x", "(v)")
b$ = Replace(b$, "e^", "exp")
b$ = translate(b$)
b$ = Replace(b$, "(v)", "x")
a$ = ExpChk_d(b$)
If a$ = "" Then a$ = CleanUpExrp(d_fx(b$))
a$ = Replace(a$, "ep", "exp")
Dfx.Text = a$
Dfx.ToolTipText = a$
fx.ToolTipText = fx.Text
End Sub



Private Sub Form_Resize()
On Error Resume Next
Me.Height = 4650
Me.Width = 5085
End Sub

Private Sub Timer1_Timer()
fx.SetFocus
Timer1.Enabled = False
End Sub
Private Sub fx_KeyDown(keycode As Integer, Shift As Integer)
Select Case keycode
   Case vbKeyEscape
   fx.Text = ""
End Select
End Sub
Private Sub dfx_KeyDown(keycode As Integer, Shift As Integer)
Select Case keycode
   Case vbKeyEscape
   fx.Text = ""
   fx.SetFocus
End Select
End Sub

