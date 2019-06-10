VERSION 5.00
Begin VB.Form MainWin 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "方解石"
   ClientHeight    =   3390
   ClientLeft      =   150
   ClientTop       =   540
   ClientWidth     =   2250
   Icon            =   "MainWin.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3390
   ScaleWidth      =   2250
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton DrawGraph 
      Default         =   -1  'True
      Height          =   375
      Left            =   2400
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   2760
      Width           =   615
   End
   Begin VB.TextBox Expr 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   1500
      Left            =   70
      MultiLine       =   -1  'True
      TabIndex        =   0
      ToolTipText     =   "请在此处输入数学表达式, 回车绘图, 单击下方的按钮调用程序计算"
      Top             =   930
      Width           =   2100
   End
   Begin VB.PictureBox Close1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   1730
      Picture         =   "MainWin.frx":08CA
      ScaleHeight     =   375
      ScaleWidth      =   240
      TabIndex        =   5
      TabStop         =   0   'False
      ToolTipText     =   "彻底关闭所有子窗口"
      Top             =   0
      Width           =   240
      Begin VB.PictureBox Close0 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   0
         MouseIcon       =   "MainWin.frx":0BD6
         MousePointer    =   99  'Custom
         Picture         =   "MainWin.frx":0D28
         ScaleHeight     =   375
         ScaleWidth      =   240
         TabIndex        =   7
         TabStop         =   0   'False
         ToolTipText     =   "关闭所有子窗口并退出程序"
         Top             =   0
         Visible         =   0   'False
         Width           =   240
      End
   End
   Begin VB.PictureBox Menu 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   510
      Left            =   0
      MouseIcon       =   "MainWin.frx":1030
      MousePointer    =   99  'Custom
      Picture         =   "MainWin.frx":1182
      ScaleHeight     =   510
      ScaleWidth      =   2250
      TabIndex        =   3
      TabStop         =   0   'False
      ToolTipText     =   "点击此处打开工具列表"
      Top             =   368
      Width           =   2250
   End
   Begin VB.PictureBox Min1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   1440
      Picture         =   "MainWin.frx":19D0
      ScaleHeight     =   375
      ScaleWidth      =   240
      TabIndex        =   4
      Top             =   0
      Width           =   240
      Begin VB.PictureBox Min0 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   0
         MouseIcon       =   "MainWin.frx":1C94
         MousePointer    =   99  'Custom
         Picture         =   "MainWin.frx":1DE6
         ScaleHeight     =   375
         ScaleWidth      =   240
         TabIndex        =   6
         TabStop         =   0   'False
         ToolTipText     =   "最小化到桌面下方"
         Top             =   0
         Visible         =   0   'False
         Width           =   240
      End
   End
   Begin VB.PictureBox Title 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   0
      Picture         =   "MainWin.frx":20AA
      ScaleHeight     =   375
      ScaleWidth      =   2250
      TabIndex        =   2
      Top             =   0
      Width           =   2250
   End
   Begin VB.Image Tool 
      Height          =   480
      Index           =   2
      Left            =   1440
      Picture         =   "MainWin.frx":2806
      ToolTipText     =   "求解一元方程"
      Top             =   2760
      Width           =   480
   End
   Begin VB.Image Tool 
      Height          =   480
      Index           =   1
      Left            =   840
      Picture         =   "MainWin.frx":30D0
      ToolTipText     =   "查看方程曲线"
      Top             =   2760
      Width           =   480
   End
   Begin VB.Image Tool 
      Height          =   480
      Index           =   0
      Left            =   240
      Picture         =   "MainWin.frx":399A
      ToolTipText     =   "计算表达式"
      Top             =   2760
      Width           =   480
   End
   Begin VB.Image ToolBar 
      Height          =   2505
      Left            =   0
      Picture         =   "MainWin.frx":4264
      Top             =   885
      Width           =   2250
   End
End
Attribute VB_Name = "MainWin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function AnimateWindow Lib "user32" (ByVal hwnd As Long, ByVal dwtime As Long, ByVal dwflags As Long) As Long



Private Sub DrawGraph_Click()
Tool_click (1)
End Sub

Private Sub Expr_click()
Unload main
End Sub

Private Sub Form_Load()
Expr.text = "正在加载方解石 请稍候..."
On Error Resume Next
Const aw_blend = &H80000
Const aw_active = &H20000
Const aw_hide = &H10000
Const aw_ver_positive = &H4

'AnimateWindow hwnd, (Fix(Timer) Mod 10) * 100, aw_active + aw_blend

Title.Cls
Menu.Cls
Me.Cls
Min1.Cls
Close1.Cls
Expr.text = ""

End Sub

Private Sub Menu_Click()
main.Show
main.Left = Me.Left + 400
main.Top = Me.Top + 930
'Me.PopupMenu acc
End Sub
Private Sub Tool_click(Index As Integer)
Select Case Index
    Case 0
    Calc.Show
    If Expr.text <> "" Then
      Calc.Text8.text = Expr.text
      Call Calc.result_Click
    End If
    Case 1
    Pic.Show
    If Expr.text <> "" Then
      Call Pic.dnc_Click
      Pic.Text1.text = Expr.text
      If InStr(Expr.text, "y") > 0 Then
      Call Pic.ImplicitFun_Click
      Else
      Call Pic.ExplicitFun_Click
      End If
     Sendkeys "{enter}"
    End If
    Case 2
    Fct.Show
    If Expr.text <> "" Then
      Fct.Text1.text = Expr.text
      Sendkeys "{enter}"
    End If
End Select
Expr.text = ""
End Sub
Private Sub Tool_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
  Tool(Index).Top = 2700
End Sub
Private Sub ToolBar_click()
Unload main
End Sub
Private Sub ToolBar_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
 
 For i = 0 To 2
   If Tool(i).Top <> 2760 Then Tool(i).Top = 2760
 Next i
End Sub

Private Sub Close0_Click()
  End
End Sub

Private Sub Min0_Click()
 Me.WindowState = 1
End Sub

Private Sub Min1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
 Min0.Visible = True
End Sub
Private Sub Close1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
 Close0.Visible = True
End Sub
Private Sub title_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
 Expr.SetFocus
End Sub

Private Sub title_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  Unload main
  
  Min0.Visible = False
  Close0.Visible = False
  'Expr.Text = Button & "  " & x & ",  " & y
  If Button = 1 Then
    Me.Left = Me.Left + x - 1100
    Me.Top = Me.Top + y - 180
  End If
End Sub

