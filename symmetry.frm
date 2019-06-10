VERSION 5.00
Begin VB.Form symmetry 
   Caption         =   "关于直线的对称点坐标"
   ClientHeight    =   2775
   ClientLeft      =   3855
   ClientTop       =   2340
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   12
      Charset         =   134
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "symmetry.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2775
   ScaleWidth      =   4680
   Begin VB.CommandButton mnuFileNew 
      Caption         =   "重置"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   480
      TabIndex        =   20
      Top             =   1380
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "计算"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2640
      TabIndex        =   19
      Top             =   1380
      Width           =   1575
   End
   Begin VB.TextBox Text7 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   3000
      TabIndex        =   15
      Text            =   "0"
      Top             =   2040
      Width           =   1335
   End
   Begin VB.TextBox Text6 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   1320
      TabIndex        =   14
      Text            =   "0"
      Top             =   2040
      Width           =   1335
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H00C0C0C0&
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
      Left            =   3120
      TabIndex        =   9
      Top             =   840
      Width           =   615
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00C0C0C0&
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
      Left            =   2160
      TabIndex        =   8
      Top             =   840
      Width           =   615
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00C0C0C0&
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
      Left            =   1200
      TabIndex        =   7
      Top             =   840
      Width           =   615
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00C0C0C0&
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
      Left            =   2280
      TabIndex        =   2
      Top             =   240
      Width           =   615
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0C0C0&
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
      Left            =   1440
      TabIndex        =   1
      Top             =   240
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "（"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   6
      Left            =   960
      TabIndex        =   18
      Top             =   2040
      Width           =   255
   End
   Begin VB.Label Label2 
      Caption         =   "）"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   4320
      TabIndex        =   17
      Top             =   2040
      Width           =   255
   End
   Begin VB.Label Label2 
      Caption         =   "，"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   2760
      TabIndex        =   16
      Top             =   1920
      Width           =   255
   End
   Begin VB.Label Label6 
      Caption         =   "对称点坐标为："
      BeginProperty Font 
         Name            =   "华文中宋"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   13
      Top             =   1970
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "＝0"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   3720
      TabIndex        =   12
      Top             =   840
      Width           =   495
   End
   Begin VB.Label Label2 
      Caption         =   "X+"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   1800
      TabIndex        =   11
      Top             =   840
      Width           =   375
   End
   Begin VB.Label Label2 
      Caption         =   "Y+"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   2760
      TabIndex        =   10
      Top             =   840
      Width           =   495
   End
   Begin VB.Label Label5 
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
      Left            =   240
      TabIndex        =   6
      Top             =   870
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "）"
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
      Left            =   2880
      TabIndex        =   5
      Top             =   240
      Width           =   255
   End
   Begin VB.Label Label3 
      Caption         =   "，"
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
      Left            =   2040
      TabIndex        =   4
      Top             =   240
      Width           =   255
   End
   Begin VB.Label Label2 
      Caption         =   "（"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   1200
      TabIndex        =   3
      Top             =   240
      Width           =   255
   End
   Begin VB.Label Label1 
      Caption         =   "输入点P的坐标："
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
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   975
   End
End
Attribute VB_Name = "symmetry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()
On Error Resume Next
x0 = Val(Text1.text)
y0 = Val(Text2.text)
a = Val(Text3.text)
b = Val(Text4.text)
c = Val(Text5.text)
If a ^ 2 + b ^ 2 = 0 Then MsgBox "这不是直线方程。": a = 1: b = 1: Text3 = a: Text4 = b: Exit Sub
x1 = ((b ^ 2 - a ^ 2) * x0 - 2 * a * b * y0 - 2 * a * c) / (a ^ 2 + b ^ 2)
y1 = ((a ^ 2 - b ^ 2) * y0 - 2 * a * b * x0 - 2 * b * c) / (a ^ 2 + b ^ 2)
Text6 = x1
Text7 = y1
If a * x0 + b * y0 + c = 0 Then MsgBox "点在直线上。"
End Sub



Private Sub mnuFileNew_Click()
    Text1 = "": Text2 = "": Text3 = "": Text4 = "": Text5 = "":
    Text6 = 0: Text7 = 0
End Sub

