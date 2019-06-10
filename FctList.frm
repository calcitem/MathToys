VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Begin VB.Form FctList 
   Caption         =   "我的方程列表"
   ClientHeight    =   4365
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5565
   ForeColor       =   &H00000000&
   Icon            =   "FctList.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   291
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   371
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton DrawList 
      BackColor       =   &H00E0E0E0&
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
      Height          =   420
      Left            =   4680
      MousePointer    =   99  'Custom
      Picture         =   "FctList.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "开始/停止绘制列表中的方程曲线"
      Top             =   3360
      Width           =   525
   End
   Begin VB.CommandButton ClearPic 
      Height          =   420
      Left            =   3960
      Picture         =   "FctList.frx":0E54
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "清除图象"
      Top             =   3360
      Width           =   525
   End
   Begin VB.CommandButton AddList 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   4920
      TabIndex        =   14
      ToolTipText     =   "把方程输入区中的方程添加到方程列表"
      Top             =   2040
      Width           =   375
   End
   Begin VB.CommandButton AddList 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   4920
      TabIndex        =   13
      ToolTipText     =   "把方程输入区中的方程添加到方程列表"
      Top             =   1440
      Width           =   375
   End
   Begin VB.CommandButton AddList 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   4920
      TabIndex        =   12
      ToolTipText     =   "把方程输入区中的方程添加到方程列表"
      Top             =   840
      Width           =   375
   End
   Begin VB.CommandButton AddList 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   4920
      TabIndex        =   11
      ToolTipText     =   "把方程输入区中的方程添加到方程列表"
      Top             =   240
      Width           =   375
   End
   Begin VB.CommandButton AddList 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   4920
      TabIndex        =   15
      ToolTipText     =   "把方程输入区中的方程添加到方程列表"
      Top             =   2640
      Width           =   375
   End
   Begin VB.TextBox FList 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   720
      TabIndex        =   4
      ToolTipText     =   "双击此处指定方程曲线的颜色"
      Top             =   2640
      Width           =   3975
   End
   Begin VB.TextBox FList 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   720
      TabIndex        =   3
      ToolTipText     =   "双击此处指定方程曲线的颜色"
      Top             =   2040
      Width           =   3975
   End
   Begin VB.TextBox FList 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   720
      TabIndex        =   2
      ToolTipText     =   "双击此处指定方程曲线的颜色"
      Top             =   1440
      Width           =   3975
   End
   Begin VB.TextBox FList 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   720
      TabIndex        =   1
      ToolTipText     =   "双击此处指定方程曲线的颜色"
      Top             =   840
      Width           =   3975
   End
   Begin VB.TextBox FList 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   720
      TabIndex        =   0
      ToolTipText     =   "双击此处指定方程曲线的颜色"
      Top             =   240
      Width           =   3975
   End
   Begin VB.CheckBox ListCk 
      Height          =   495
      Index           =   4
      Left            =   240
      TabIndex        =   16
      Top             =   1980
      Value           =   1  'Checked
      Width           =   375
   End
   Begin VB.CheckBox ListCk 
      Height          =   495
      Index           =   3
      Left            =   240
      TabIndex        =   9
      Top             =   1380
      Value           =   1  'Checked
      Width           =   375
   End
   Begin VB.CheckBox ListCk 
      Height          =   495
      Index           =   5
      Left            =   240
      TabIndex        =   10
      Top             =   2580
      Value           =   1  'Checked
      Width           =   375
   End
   Begin VB.CheckBox ListCk 
      Height          =   495
      Index           =   1
      Left            =   240
      TabIndex        =   7
      Top             =   180
      Value           =   1  'Checked
      Width           =   375
   End
   Begin VB.CheckBox ListCk 
      Height          =   495
      Index           =   2
      Left            =   240
      TabIndex        =   8
      Top             =   780
      Value           =   1  'Checked
      Width           =   375
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   120
      Top             =   3720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "FctList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Running As Boolean
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Const HWND_TOPMOST = -1
Private Const SWP_SHOWWINDOWS = &H40
Private Sub AddList_Click(Index As Integer)
If Pic.ImplicitFun.Checked = True Then
  msg = MsgBox("不允许将隐函数添加到列表中。", vbOKOnly + vbExclamation, "方程列表")
  Exit Sub
Else
  FList(Index).ForeColor = Pic.Picture1.ForeColor
  FList(Index).text = Pic.Text1.text
  DrawList.SetFocus
End If
End Sub

Private Sub ClearPic_Click()
Call Pic.Command1_Click
DrawList.SetFocus
End Sub

Public Sub DrawList_Click()
Dim FctBak As String
Dim FctColor As String
If Running = True Then
  Running = False
  Call Pic.Command3_Click
Else
  Running = True
End If
FctBak = Pic.Text1.text
FctColor = Pic.Picture1.ForeColor
Pic.ImplicitFun.Checked = False
Pic.ExplicitFun.Checked = True
For i = 1 To 5
 If ListCk(i).Value = 1 And Trim(FList(i).text) <> "" Then
    Call Rcflist
    FList(i).BackColor = &HE0E0E0
    Pic.Pfc = FList(i).ForeColor
    Pic.Text1.text = FList(i)
    If Running = False Then Exit For
    Call Pic.Command3_Click
 End If
 DoEvents
 If Running = False Then Exit For
Next i
Call Rcflist
Pic.Text1.text = FctBak
Pic.Pfc = FctColor
Running = False
End Sub

Private Sub FList_DblClick(Index As Integer)
CommonDialog1.ShowColor
FList(Index).ForeColor = CommonDialog1.color
End Sub
Private Sub Rcflist()
For i = 1 To 5
FList(i).BackColor = &HFFFFFF
Next i
End Sub
Private Sub Form_Load()
Dim retValue As Long
Running = False
retValue = SetWindowPos(Me.hwnd, HWND_TOPMOST, Me.CurrentX + 520, Me.CurrentY + 48, 380, 310, SWP_SHOWWINDOWS)
End Sub
Private Sub Form_unLoad(Cancel As Integer)
Running = False
Pic.fclb.Checked = False
End Sub

Private Sub ListCk_Click(Index As Integer)
Dim FctBak As String
Dim FctColor As String


FctBak = Pic.Text1.text
FctColor = Pic.Picture1.ForeColor
Pic.ImplicitFun.Checked = False
Pic.ExplicitFun.Checked = True

 If Trim(FList(Index).text) <> "" Then

    If ListCk(Index).Value = 0 Then
      Pic.Pfc = Pic.Picture1.BackColor
    Else
      Pic.Pfc = FList(Index).ForeColor
    End If
    
    Pic.Text1.text = FList(Index).text
    
    Call Pic.Command3_Click
 End If
Pic.Text1.text = FctBak
Pic.Pfc = FctColor


End Sub
