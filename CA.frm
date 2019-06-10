VERSION 5.00
Begin VB.Form CA 
   Caption         =   "排列与组合"
   ClientHeight    =   3840
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7230
   Icon            =   "CA.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   3840
   ScaleWidth      =   7230
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer Timer1 
      Interval        =   6000
      Left            =   3000
      Top             =   1680
   End
   Begin VB.CommandButton calcC 
      Caption         =   "="
      Height          =   375
      Left            =   2760
      TabIndex        =   9
      ToolTipText     =   "计算组合"
      Top             =   2520
      Width           =   615
   End
   Begin VB.CommandButton calcA 
      Caption         =   "="
      Height          =   375
      Left            =   2760
      TabIndex        =   8
      ToolTipText     =   "计算排列"
      Top             =   720
      Width           =   615
   End
   Begin VB.TextBox cc 
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
      Left            =   3600
      TabIndex        =   7
      ToolTipText     =   "m!/[n!*(m-n)!]"
      Top             =   2520
      Width           =   2895
   End
   Begin VB.TextBox aa 
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
      Left            =   3600
      TabIndex        =   6
      ToolTipText     =   "m!/(m-n)!"
      Top             =   720
      Width           =   2895
   End
   Begin VB.TextBox cm 
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
      Left            =   1680
      TabIndex        =   5
      ToolTipText     =   "m"
      Top             =   2880
      Width           =   735
   End
   Begin VB.TextBox cn 
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
      Left            =   1680
      TabIndex        =   4
      ToolTipText     =   "n"
      Top             =   2160
      Width           =   735
   End
   Begin VB.TextBox am 
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
      Left            =   1680
      TabIndex        =   3
      ToolTipText     =   "m"
      Top             =   1080
      Width           =   735
   End
   Begin VB.TextBox an 
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
      Left            =   1680
      TabIndex        =   2
      ToolTipText     =   "n"
      Top             =   360
      Width           =   735
   End
   Begin VB.Label Lb 
      Caption         =   "A"
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
      Index           =   1
      Left            =   720
      TabIndex        =   1
      ToolTipText     =   "排列"
      Top             =   240
      Width           =   975
   End
   Begin VB.Label Lb 
      Caption         =   "C"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   72
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Index           =   0
      Left            =   720
      TabIndex        =   0
      ToolTipText     =   "组合"
      Top             =   2040
      Width           =   855
   End
End
Attribute VB_Name = "CA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub calcA_Click()
On Error GoTo 10:
aa.Text = ""
m = Fix(Abs(Val(am.Text)))
n = Fix(Abs(Val(an.Text)))

If m < n Then
  t = m
  m = n
  n = t
End If

am.Text = m: an.Text = n
a = Arrange(m, n)
aa.Text = a
10: If err <> 0 Then msg = MsgBox("计算器无法完成计算。", vbOKOnly, "计算器")
End Sub

Private Sub calcC_Click()
On Error GoTo 10:
cc.Text = ""
m = Fix(Abs(Val(cm.Text)))
n = Fix(Abs(Val(cn.Text)))

If m < n Then
  t = m
  m = n
  n = t
End If

cm.Text = m: cn.Text = n
c = Combination(m, n)
cc.Text = c
10: If err <> 0 Then msg = MsgBox("计算器无法完成计算。", vbOKOnly, "计算器")
End Sub

Private Sub Timer1_Timer()
Me.Caption = "排列组合"
Timer1.Enabled = False
End Sub
