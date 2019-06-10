VERSION 5.00
Begin VB.Form Stri 
   Caption         =   "三角形面积计算"
   ClientHeight    =   3270
   ClientLeft      =   3465
   ClientTop       =   2340
   ClientWidth     =   5535
   ForeColor       =   &H00000000&
   Icon            =   "Stri.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3270
   ScaleWidth      =   5535
   Begin VB.CommandButton mnuFileNew 
      Caption         =   "重置"
      Height          =   375
      Left            =   3240
      TabIndex        =   20
      Top             =   600
      Width           =   855
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF80&
      Height          =   375
      Index           =   1
      Left            =   480
      TabIndex        =   19
      Text            =   "0."
      Top             =   120
      Width           =   4695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "计算"
      Height          =   375
      Left            =   4320
      TabIndex        =   18
      Top             =   600
      Width           =   855
   End
   Begin VB.TextBox Text9 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4200
      TabIndex        =   17
      Top             =   2640
      Width           =   855
   End
   Begin VB.TextBox Text8 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2640
      TabIndex        =   16
      Top             =   2640
      Width           =   855
   End
   Begin VB.TextBox Text7 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      TabIndex        =   15
      Top             =   2640
      Width           =   855
   End
   Begin VB.TextBox Text6 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4680
      TabIndex        =   11
      Top             =   1920
      Width           =   735
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   2640
      TabIndex        =   9
      Top             =   1920
      Width           =   855
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   1080
      TabIndex        =   7
      Top             =   1920
      Width           =   855
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   3360
      TabIndex        =   5
      Top             =   1200
      Width           =   855
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   1080
      TabIndex        =   3
      Top             =   1200
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "③ 边Ⅰ："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   255
      Index           =   8
      Left            =   120
      TabIndex        =   14
      Top             =   2760
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "边Ⅱ："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   255
      Index           =   7
      Left            =   2040
      TabIndex        =   13
      Top             =   2760
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "边Ⅲ："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   255
      Index           =   6
      Left            =   3600
      TabIndex        =   12
      Top             =   2760
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "夹角(度)："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   255
      Index           =   5
      Left            =   3600
      TabIndex        =   10
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "边Ⅱ："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   255
      Index           =   4
      Left            =   2040
      TabIndex        =   8
      Top             =   2040
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "② 边Ⅰ："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   6
      Top             =   2040
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "底边上的高："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   2
      Left            =   2040
      TabIndex        =   4
      Top             =   1320
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "① 底："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "请选择输入已知条件:"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   11.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   2535
   End
   Begin VB.Label Label1 
      Caption         =   "S="
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   375
   End
End
Attribute VB_Name = "Stri"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error Resume Next
s1 = 0: s2 = 0: s3 = 0
a1 = Val(Text2)
h1 = Val(Text3)
a2 = Val(Text4)
b2 = Val(Text5)
c0 = Val(Text6)
c = c0 * 3.14159265358979 / 180
a3 = Val(Text7)
b3 = Val(Text8)
c3 = Val(Text9)
s1 = a1 * h1 / 2
s2 = a2 * b2 * Sin(c) / 2
If a3 * b3 * c3 <> 0 And (a3 + b3 + c3) * (a3 + b3 - c3) * (a3 - b3 + c3) * (b3 + c3 - a3) > 0 Then GoTo 5
If a3 * b3 * c3 <> 0 And (a3 + b3 + c3) * (a3 + b3 - c3) * (a3 - b3 + c3) * (b3 + c3 - a3) <= 0 Then GoTo 30
If a3 * b3 * c3 = 0 Then s3 = 0: GoTo 6
5 s3 = Sqr((a3 + b3 + c3) * (a3 + b3 - c3) * (a3 - b3 + c3) * (b3 + c3 - a3)) / 4
6 If Abs(s2) + Abs(s3) = 0 Then Text1(1) = s1: GoTo 40
If Abs(s1) + Abs(s3) = 0 Then Text1(1) = s2: GoTo 40
If Abs(s1) + Abs(s2) = 0 Then Text1(1) = s3: GoTo 40
Text1(1) = "只须且只能选择一组已知条件。": Text2 = "": Text3 = "": Text4 = "": Text5 = "": Text6 = "": Text7 = "": Text8 = "": Text9 = "": GoTo 40
30  Text1(1) = "这三边不能构成三角形。": Text2 = "": Text3 = "": Text4 = "": Text5 = "": Text6 = "": Text7 = "": Text8 = "": Text9 = ""
40 End Sub



Private Sub mnuFileNew_Click()
    Text1(1) = "0.": Text2 = "": Text3 = "": Text4 = "": Text5 = "": Text6 = "": Text7 = "": Text8 = "": Text9 = ""
End Sub


