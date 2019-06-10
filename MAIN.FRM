VERSION 5.00
Begin VB.Form main 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FF8080&
   BorderStyle     =   0  'None
   Caption         =   "方解石"
   ClientHeight    =   5460
   ClientLeft      =   1410
   ClientTop       =   975
   ClientWidth     =   1485
   ClipControls    =   0   'False
   FillColor       =   &H00FF8080&
   ForeColor       =   &H00FF8080&
   Icon            =   "MAIN.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   OLEDropMode     =   1  'Manual
   Picture         =   "MAIN.frx":08CA
   ScaleHeight     =   5460
   ScaleWidth      =   1485
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Expand 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   960
      Top             =   4680
   End
   Begin VB.Label mainmenu 
      BackStyle       =   0  'Transparent
      Caption         =   "求定积分"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   7
      Left            =   410
      TabIndex        =   8
      ToolTipText     =   "求数值积分"
      Top             =   4440
      Width           =   735
   End
   Begin VB.Label mainmenu 
      BackStyle       =   0  'Transparent
      Caption         =   "高精度计算"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   330
      TabIndex        =   7
      ToolTipText     =   "执行高精度加、减、乘、除、开平方运算"
      Top             =   840
      Width           =   975
   End
   Begin VB.Label mainmenu 
      BackStyle       =   0  'Transparent
      Caption         =   "总和与连乘"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   5
      Left            =   330
      TabIndex        =   6
      ToolTipText     =   "数列求和、求积"
      Top             =   3240
      Width           =   1095
   End
   Begin VB.Label mainmenu 
      BackStyle       =   0  'Transparent
      Caption         =   "帮助"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   8
      Left            =   570
      TabIndex        =   5
      ToolTipText     =   "查看联机文档"
      Top             =   5040
      Width           =   375
   End
   Begin VB.Label mainmenu 
      BackStyle       =   0  'Transparent
      Caption         =   "显函数求导"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   6
      Left            =   360
      TabIndex        =   4
      ToolTipText     =   "显函数求导"
      Top             =   3840
      Width           =   975
   End
   Begin VB.Label mainmenu 
      BackStyle       =   0  'Transparent
      Caption         =   "方程曲线"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   410
      TabIndex        =   3
      ToolTipText     =   "绘制 y=f(x)、F(x,y)=0 的图像。"
      Top             =   1440
      Width           =   735
   End
   Begin VB.Label mainmenu 
      BackStyle       =   0  'Transparent
      Caption         =   "线性方程组"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   4
      Left            =   330
      TabIndex        =   2
      ToolTipText     =   "解含任意多个未知数的线性方程组"
      Top             =   2640
      Width           =   975
   End
   Begin VB.Label mainmenu 
      BackStyle       =   0  'Transparent
      Caption         =   "解一元方程"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   330
      TabIndex        =   1
      ToolTipText     =   "解形如 f(x)=0 的方程"
      Top             =   2040
      Width           =   975
   End
   Begin VB.Label mainmenu 
      BackStyle       =   0  'Transparent
      Caption         =   "科学计算器"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   330
      TabIndex        =   0
      ToolTipText     =   "输入表达式，一次计算得出结果"
      Top             =   240
      Width           =   975
   End
End
Attribute VB_Name = "main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub Expand_Timer()
If Me.Height >= 5260 Then Me.Height = 5260: Expand.Enabled = False
  Me.Height = Me.Height + 250

End Sub

Private Sub form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
For Index = 0 To 8
mainmenu(Index).Font.Underline = False
Next Index
End Sub
Private Sub Form_Load()
On Error Resume Next

Me.Height = 5460
Me.Height = 1
Expand.Enabled = True

End Sub
Private Sub Form_unLoad(Cancel As Integer)
'MainWin.Expr.SetFocus
End Sub
Private Sub Form_click()
Unload main
End Sub

Private Sub mainmenu_mousemove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
mainmenu(Index).Font.Underline = True
End Sub

Private Sub mainmenu_click(Index As Integer)

Select Case Index
  Case 0
  Calc.Show
  Case 1
  hp.Show
  Case 2
  Pic.Show
  Case 3
  Fct.Show
  Case 4
  GJ.Show
  Case 5
  series.Show
  Case 6
  der.Show
  Case 7
  dfintegral.Show
  Case 8
  On Error GoTo rd
  readme = Shell("explorer " & App.Path & "\README.html", 1)
rd:  If err = 53 Then
        msg = MsgBox("联机文档没有找到", vbOKOnly, "帮助")
        Resume Next
     End If
 
End Select
Unload main

End Sub



