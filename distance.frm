VERSION 5.00
Begin VB.Form distance 
   Caption         =   "平面距离公式计算"
   ClientHeight    =   3540
   ClientLeft      =   2760
   ClientTop       =   2340
   ClientWidth     =   6345
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   9
      Charset         =   134
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "distance.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3540
   ScaleWidth      =   6345
   WhatsThisHelp   =   -1  'True
   Begin VB.TextBox Text10 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   330
      Left            =   840
      TabIndex        =   28
      Text            =   "0."
      Top             =   240
      Width           =   5055
   End
   Begin VB.CommandButton Command2 
      Caption         =   "计算"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2040
      TabIndex        =   26
      Top             =   2880
      Width           =   3735
   End
   Begin VB.TextBox Text9 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   5160
      TabIndex        =   22
      Top             =   2505
      Width           =   615
   End
   Begin VB.TextBox Text8 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   4200
      TabIndex        =   21
      Top             =   2505
      Width           =   615
   End
   Begin VB.TextBox Text7 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   3240
      TabIndex        =   20
      Top             =   2505
      Width           =   615
   End
   Begin VB.Frame Frame1 
      Caption         =   "，"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   4800
      TabIndex        =   17
      Top             =   1920
      Width           =   255
   End
   Begin VB.TextBox Text6 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   5040
      TabIndex        =   15
      Top             =   1905
      Width           =   615
   End
   Begin VB.TextBox Text5 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   4200
      TabIndex        =   14
      Top             =   1905
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "计算"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1920
      TabIndex        =   11
      Top             =   1200
      Width           =   3855
   End
   Begin VB.Frame Frame1 
      Caption         =   "，"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   4680
      TabIndex        =   9
      Top             =   840
      Width           =   255
   End
   Begin VB.Frame Frame1 
      Caption         =   "，"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   2520
      TabIndex        =   6
      Top             =   840
      Width           =   255
   End
   Begin VB.TextBox Text4 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   5040
      TabIndex        =   4
      Top             =   825
      Width           =   615
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   4080
      TabIndex        =   3
      Top             =   825
      Width           =   615
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   2760
      TabIndex        =   2
      Top             =   825
      Width           =   615
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   1920
      TabIndex        =   1
      Top             =   825
      Width           =   615
   End
   Begin VB.Label Label7 
      Caption         =   "d ="
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   27
      Top             =   240
      Width           =   495
   End
   Begin VB.Label Label2 
      Caption         =   "=0"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   5880
      TabIndex        =   25
      Top             =   2475
      Width           =   375
   End
   Begin VB.Label Label2 
      Caption         =   "Y+"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   4800
      TabIndex        =   24
      Top             =   2520
      Width           =   375
   End
   Begin VB.Label Label2 
      Caption         =   "X+"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   3840
      TabIndex        =   23
      Top             =   2520
      Width           =   375
   End
   Begin VB.Label Label6 
      Caption         =   "直线方程："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2160
      TabIndex        =   19
      Top             =   2550
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "）"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   5640
      TabIndex        =   18
      Top             =   1920
      Width           =   255
   End
   Begin VB.Label Label2 
      Caption         =   "（"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   4080
      TabIndex        =   16
      Top             =   1920
      Width           =   255
   End
   Begin VB.Label Label5 
      Caption         =   "输入已知点的坐标："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2160
      TabIndex        =   13
      Top             =   1920
      Width           =   1695
   End
   Begin VB.Label Label4 
      Caption         =   "点到直线的距离计算"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   375
      Left            =   120
      TabIndex        =   12
      Top             =   1920
      Width           =   2055
   End
   Begin VB.Label Label3 
      Caption         =   "）"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   5640
      TabIndex        =   10
      Top             =   840
      Width           =   255
   End
   Begin VB.Label Label2 
      Caption         =   "（"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   3960
      TabIndex        =   8
      Top             =   840
      Width           =   255
   End
   Begin VB.Label Label3 
      Caption         =   "）"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   3360
      TabIndex        =   7
      Top             =   840
      Width           =   255
   End
   Begin VB.Label Label2 
      Caption         =   "（"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   1800
      TabIndex        =   5
      Top             =   840
      Width           =   255
   End
   Begin VB.Label Label1 
      Caption         =   "两点间的距离计算"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   1695
   End
End
Attribute VB_Name = "distance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()
On Error Resume Next
x1 = Val(Text1.text)
y1 = Val(Text2.text)
x2 = Val(Text3.text)
y2 = Val(Text4.text)
d = Sqr((x2 - x1) ^ 2 + (y2 - y1) ^ 2)
Cls
Text10.text = d
End Sub

Private Sub Command2_Click()
On Error Resume Next
x0 = Val(Text5.text)
y0 = Val(Text6.text)
a = Val(Text7.text)
b = Val(Text8.text)
If a ^ 2 + b ^ 2 = 0 Then MsgBox "A、B不能同时为0": a = 1: b = 1
c = Val(Text9.text)
d2 = Abs(a * x0 + b * y0 + c) / Sqr(a ^ 2 + b ^ 2)
Cls
Text10.text = d2
End Sub

Private Sub mnuHelpAbout_Click()
End Sub

Private Sub mnuFileExit_Click()
    '卸载窗体
    Unload Me

End Sub

