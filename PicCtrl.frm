VERSION 5.00
Begin VB.Form PicCtrl 
   BackColor       =   &H80000000&
   Caption         =   "图像控制面板"
   ClientHeight    =   4350
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7035
   FillColor       =   &H00C000C0&
   ForeColor       =   &H00FF0000&
   Icon            =   "PicCtrl.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4350
   ScaleWidth      =   7035
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox l 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   495
      Index           =   2
      Left            =   4920
      TabIndex        =   2
      Text            =   "2"
      ToolTipText     =   "指定缩小的单位长度"
      Top             =   2160
      Width           =   1455
   End
   Begin VB.TextBox l 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   495
      Index           =   1
      Left            =   3000
      TabIndex        =   1
      Text            =   "2"
      ToolTipText     =   "指定缩小的单位长度"
      Top             =   2160
      Width           =   1455
   End
   Begin VB.TextBox l 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   495
      Index           =   0
      Left            =   720
      TabIndex        =   0
      Text            =   "2"
      ToolTipText     =   "指定移动的单位长度"
      Top             =   2160
      Width           =   1455
   End
   Begin VB.Image dx 
      Height          =   1335
      Index           =   1
      Left            =   4920
      Picture         =   "PicCtrl.frx":08CA
      Stretch         =   -1  'True
      ToolTipText     =   "缩小显示范围, 另可用Alt+鼠标左键放大图像"
      Top             =   360
      Width           =   1335
   End
   Begin VB.Image back 
      Height          =   480
      Left            =   1200
      Picture         =   "PicCtrl.frx":0F20
      Stretch         =   -1  'True
      ToolTipText     =   "恢复默认值"
      Top             =   840
      Width           =   480
   End
   Begin VB.Image dx 
      Height          =   1335
      Index           =   0
      Left            =   3000
      Picture         =   "PicCtrl.frx":14AA
      Stretch         =   -1  'True
      ToolTipText     =   "扩大显示范围"
      Top             =   360
      Width           =   1335
   End
   Begin VB.Image Mv 
      Height          =   480
      Index           =   2
      Left            =   1200
      Picture         =   "PicCtrl.frx":1AFE
      Stretch         =   -1  'True
      ToolTipText     =   "向上移动"
      Top             =   240
      Width           =   480
   End
   Begin VB.Image Mv 
      Height          =   480
      Index           =   1
      Left            =   600
      Picture         =   "PicCtrl.frx":2088
      Stretch         =   -1  'True
      ToolTipText     =   "向左移动"
      Top             =   840
      Width           =   480
   End
   Begin VB.Image Mv 
      Height          =   480
      Index           =   3
      Left            =   1800
      Picture         =   "PicCtrl.frx":2612
      Stretch         =   -1  'True
      ToolTipText     =   "向右移动"
      Top             =   840
      Width           =   480
   End
   Begin VB.Image Mv 
      Height          =   480
      Index           =   4
      Left            =   1200
      Picture         =   "PicCtrl.frx":2B9C
      Stretch         =   -1  'True
      ToolTipText     =   "向下移动"
      Top             =   1440
      Width           =   480
   End
End
Attribute VB_Name = "PicCtrl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Private Const HWND_TOPMOST = -1

Private Const SWP_SHOWWINDOWS = &H40
Private Sub Zoom_change()
Pic.xmax.text = Val(Pic.xmax.text) * 2 ^ Zoom.Value
Pic.xmin.text = Val(Pic.xmin.text) * 2 ^ Zoom.Value
Pic.ymax.text = Val(Pic.ymax.text) * 2 ^ Zoom.Value
Pic.ymin.text = Val(Pic.ymin.text) * 2 ^ Zoom.Value
Pic.Xinc.text = Val(Pic.Xinc.text) * 2 ^ Zoom.Value
Pic.Yinc.text = Val(Pic.Yinc.text) * 2 ^ Zoom.Value
Call DrawGrp
End Sub

Private Sub back_Click()
Call Pic.Frame5_dblclick
Call DrawGrp
End Sub

Public Sub dx_Click(Index As Integer)
'n = l(1).Text
n = Fc(Bracket(translate(l(1).text)), 0, 0)
'm = l(2).Text
m = Fc(Bracket(translate(l(2).text)), 0, 0)
Select Case Index
Case 0
Pic.xmax.text = Val(Pic.xmax.text) + n
Pic.xmin.text = Val(Pic.xmin.text) - n
Pic.ymax.text = Val(Pic.ymax.text) + n
Pic.ymin.text = Val(Pic.ymin.text) - n
Case 1
Pic.xmax.text = Val(Pic.xmax.text) - m
Pic.xmin.text = Val(Pic.xmin.text) + m
Pic.ymax.text = Val(Pic.ymax.text) - m
Pic.ymin.text = Val(Pic.ymin.text) + m
End Select
Call DrawGrp
End Sub

Private Sub Form_Load()
Dim retValue As Long

retValue = SetWindowPos(Me.hwnd, HWND_TOPMOST, Me.CurrentX + 25, Me.CurrentY + 480, 480, 230, SWP_SHOWWINDOWS)
End Sub

Private Sub Mv_Click(Index As Integer)
'n = Val(l(0).Text)
n = Fc(Bracket(translate(l(0).text)), 0, 0)
Select Case Index
Case 1
Pic.xmax.text = Val(Pic.xmax.text) + n
Pic.xmin.text = Val(Pic.xmin.text) + n
Case 2
Pic.ymax.text = Val(Pic.ymax.text) + n
Pic.ymin.text = Val(Pic.ymin.text) + n
Case 3
Pic.xmax.text = Val(Pic.xmax.text) - n
Pic.xmin.text = Val(Pic.xmin.text) - n
Case 4
Pic.ymax.text = Val(Pic.ymax.text) - n
Pic.ymin.text = Val(Pic.ymin.text) - n
End Select

Call DrawGrp

End Sub
