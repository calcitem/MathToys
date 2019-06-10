VERSION 5.00
Begin VB.Form fun 
   Caption         =   "解三阶线性方程组"
   ClientHeight    =   5520
   ClientLeft      =   60
   ClientTop       =   240
   ClientWidth     =   8475
   Icon            =   "LINEFUN.frx":0000
   LinkTopic       =   "linefun"
   MaxButton       =   0   'False
   ScaleHeight     =   5520
   ScaleWidth      =   8475
   StartUpPosition =   3  '窗口缺省
   Begin VB.Frame cover 
      BorderStyle     =   0  'None
      Height          =   1815
      Index           =   2
      Left            =   4200
      TabIndex        =   33
      Top             =   240
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Frame cover 
      BorderStyle     =   0  'None
      Height          =   855
      Index           =   0
      Left            =   480
      TabIndex        =   31
      Top             =   1920
      Visible         =   0   'False
      Width           =   7455
   End
   Begin VB.Frame cover 
      BorderStyle     =   0  'None
      Height          =   855
      Index           =   1
      Left            =   360
      TabIndex        =   32
      Top             =   4560
      Visible         =   0   'False
      Width           =   6495
   End
   Begin VB.CommandButton change 
      Caption         =   "二阶"
      Height          =   615
      Left            =   7080
      TabIndex        =   13
      Top             =   3120
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "初始化"
      Height          =   615
      Left            =   7080
      TabIndex        =   14
      Top             =   3840
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "解"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7080
      TabIndex        =   0
      Top             =   4560
      Width           =   1095
   End
   Begin VB.TextBox z 
      BackColor       =   &H80000004&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   495
      Left            =   1320
      MaxLength       =   256
      MousePointer    =   3  'I-Beam
      TabIndex        =   17
      Top             =   4680
      Width           =   5415
   End
   Begin VB.TextBox y 
      BackColor       =   &H80000004&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   495
      Left            =   1320
      MaxLength       =   256
      MousePointer    =   3  'I-Beam
      TabIndex        =   16
      Top             =   3960
      Width           =   5415
   End
   Begin VB.TextBox x 
      BackColor       =   &H80000004&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   495
      Left            =   1320
      MaxLength       =   256
      MousePointer    =   3  'I-Beam
      TabIndex        =   15
      Top             =   3120
      Width           =   5415
   End
   Begin VB.TextBox d3 
      BackColor       =   &H80000004&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   495
      Left            =   6720
      MaxLength       =   256
      MousePointer    =   3  'I-Beam
      TabIndex        =   12
      Top             =   2040
      Width           =   1095
   End
   Begin VB.TextBox d2 
      BackColor       =   &H80000004&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   495
      Left            =   6720
      MaxLength       =   256
      MousePointer    =   3  'I-Beam
      TabIndex        =   6
      Top             =   1200
      Width           =   1095
   End
   Begin VB.TextBox d1 
      BackColor       =   &H80000004&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   495
      Left            =   6720
      MaxLength       =   256
      MousePointer    =   3  'I-Beam
      TabIndex        =   3
      Top             =   360
      Width           =   1095
   End
   Begin VB.TextBox c3 
      BackColor       =   &H80000004&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   495
      Left            =   4680
      MaxLength       =   256
      MousePointer    =   3  'I-Beam
      TabIndex        =   11
      Top             =   2040
      Width           =   1095
   End
   Begin VB.TextBox c2 
      BackColor       =   &H80000004&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   495
      Left            =   4680
      MaxLength       =   256
      MousePointer    =   3  'I-Beam
      TabIndex        =   8
      Top             =   1200
      Width           =   1095
   End
   Begin VB.TextBox c1 
      BackColor       =   &H80000004&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   495
      Left            =   4680
      MaxLength       =   256
      MousePointer    =   3  'I-Beam
      TabIndex        =   7
      Top             =   360
      Width           =   1095
   End
   Begin VB.TextBox b3 
      BackColor       =   &H80000004&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   495
      Left            =   2640
      MaxLength       =   256
      MousePointer    =   3  'I-Beam
      TabIndex        =   10
      Top             =   2040
      Width           =   1095
   End
   Begin VB.TextBox b2 
      BackColor       =   &H80000004&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   495
      Left            =   2640
      MaxLength       =   256
      MousePointer    =   3  'I-Beam
      TabIndex        =   5
      Top             =   1200
      Width           =   1095
   End
   Begin VB.TextBox a3 
      BackColor       =   &H80000004&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   495
      Left            =   600
      MaxLength       =   256
      MousePointer    =   3  'I-Beam
      TabIndex        =   9
      Top             =   2040
      Width           =   1095
   End
   Begin VB.TextBox a2 
      BackColor       =   &H80000004&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   495
      Left            =   600
      MaxLength       =   256
      MousePointer    =   3  'I-Beam
      TabIndex        =   4
      Top             =   1200
      Width           =   1095
   End
   Begin VB.TextBox a1 
      BackColor       =   &H80000004&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   495
      Left            =   600
      MaxLength       =   256
      MousePointer    =   3  'I-Beam
      TabIndex        =   1
      Top             =   360
      Width           =   1095
   End
   Begin VB.TextBox b1 
      BackColor       =   &H80000004&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   495
      Left            =   2640
      MaxLength       =   256
      MousePointer    =   3  'I-Beam
      TabIndex        =   2
      Top             =   360
      Width           =   1095
   End
   Begin VB.Label Label13 
      Caption         =   "z="
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   480
      TabIndex        =   30
      Top             =   4560
      Width           =   735
   End
   Begin VB.Label Label12 
      Caption         =   "y ="
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   480
      TabIndex        =   29
      Top             =   3720
      Width           =   735
   End
   Begin VB.Label Label11 
      Caption         =   "x ="
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   480
      TabIndex        =   28
      Top             =   2880
      Width           =   735
   End
   Begin VB.Label Label10 
      Caption         =   "z ="
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5880
      TabIndex        =   27
      Top             =   1920
      Width           =   735
   End
   Begin VB.Label Label9 
      Caption         =   "z ="
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5880
      TabIndex        =   26
      Top             =   1080
      Width           =   735
   End
   Begin VB.Label Label8 
      Caption         =   "z ="
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5880
      TabIndex        =   25
      Top             =   240
      Width           =   735
   End
   Begin VB.Label Label7 
      Caption         =   "y +"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3840
      TabIndex        =   24
      Top             =   1920
      Width           =   735
   End
   Begin VB.Label Label6 
      Caption         =   "y +"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3840
      TabIndex        =   23
      Top             =   1080
      Width           =   735
   End
   Begin VB.Label Label5 
      Caption         =   "y +"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3840
      TabIndex        =   22
      Top             =   240
      Width           =   735
   End
   Begin VB.Label Label4 
      Caption         =   "x +"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1800
      TabIndex        =   21
      Top             =   1920
      Width           =   735
   End
   Begin VB.Label Label3 
      Caption         =   "x +"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1800
      TabIndex        =   20
      Top             =   1080
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "x +"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1800
      TabIndex        =   19
      Top             =   240
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "{"
      BeginProperty Font 
         Name            =   "新宋体"
         Size            =   72
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   -120
      TabIndex        =   18
      ToolTipText     =   ""
      Top             =   720
      Width           =   735
   End
End
Attribute VB_Name = "fun"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub change_Click()
If change.Caption = "三阶" Then
  cover(1).Visible = False
  cover(2).Visible = False
  cover(0).Visible = False

End If

If change.Caption = "二阶" Then
  cover(1).Visible = True
  cover(2).Visible = True
  cover(0).Visible = True
  z.text = ""
End If
If change.Caption = "二阶" Then change.Caption = "三阶" Else change.Caption = "二阶"
End Sub

Private Sub Command1_Click()
x.text = ""
y.text = ""
z.text = ""

On Error GoTo er

a1 = Val(a1.text)
a2 = Val(a2.text)
a3 = Val(a3.text)

b1 = Val(b1.text)
b2 = Val(b2.text)
b3 = Val(b3.text)

c1 = Val(c1.text)
c2 = Val(c2.text)
c3 = Val(c3.text)

d1 = Val(d1.text)
d2 = Val(d2.text)
d3 = Val(d3.text)

If change.Caption = "二阶" Then
delta = a1 * (b2 * c3 - b3 * c2) - a2 * (b1 * c3 - b3 * c1) + a3 * (b1 * c2 - b2 * c1)
If delta = 0 Then msg = MsgBox("此方程组无解或有无穷多个解。", , "解线性方程组"): GoTo endsub
deltax = d1 * (b2 * c3 - b3 * c2) - d2 * (b1 * c3 - b3 * c1) + d3 * (b1 * c2 - b2 * c1)
deltay = a1 * (d2 * c3 - d3 * c2) - a2 * (d1 * c3 - d3 * c1) + a3 * (d1 * c2 - d2 * c1)
deltaz = a1 * (b2 * d3 - b3 * d2) - a2 * (b1 * d3 - b3 * d1) + a3 * (b1 * d2 - b2 * d1)
End If

If change.Caption = "三阶" Then
delta = a1 * b2 - a2 * b1
If delta = 0 Then msg = MsgBox("此方程组无解或有无穷多个解。", , "解线性方程组"): GoTo endsub
deltax = d1 * b2 - d2 * b1
deltay = a1 * d2 - a2 * d1
End If

x = deltax / delta
y = deltay / delta
z = deltaz / delta

x.text = x
y.text = y
z.text = z
er: If err <> 0 Then msg = MsgBox("溢出", , "解线性方程组"): Resume endsub
endsub:
End Sub


Private Sub Command2_Click()
a1.text = ""
a2.text = ""
a3.text = ""

b1.text = ""
b2.text = ""
b3.text = ""

c1.text = ""
c2.text = ""
c3.text = ""

d1.text = ""
d2.text = ""
d3.text = ""


End Sub

Private Sub Form_unLoad(Cancel As Integer)
'main.WindowState = 0
End Sub
