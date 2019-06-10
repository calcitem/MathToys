VERSION 5.00
Begin VB.Form hp 
   Caption         =   "高精度计算器"
   ClientHeight    =   5010
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   9465
   Icon            =   "HPCS.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   5010
   ScaleWidth      =   9465
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Height          =   735
      Left            =   8520
      Picture         =   "HPCS.frx":0CCA
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "运行大数计算器"
      Top             =   600
      Width           =   615
   End
   Begin VB.ComboBox lendvs 
      BackColor       =   &H80000000&
      Height          =   300
      ItemData        =   "HPCS.frx":1594
      Left            =   8520
      List            =   "HPCS.frx":15AA
      TabIndex        =   17
      Text            =   "128"
      ToolTipText     =   "指定当商的长度为何值时终止计算 (注意: 不一定精确到那一位)"
      Top             =   3120
      Width           =   615
   End
   Begin VB.ComboBox clearzero 
      BackColor       =   &H80000000&
      Height          =   300
      ItemData        =   "HPCS.frx":15CD
      Left            =   8520
      List            =   "HPCS.frx":15E6
      TabIndex        =   16
      Text            =   "16"
      ToolTipText     =   "对商而言,几个零后的数字将被剔除"
      Top             =   2760
      Width           =   615
   End
   Begin VB.ComboBox sqrtdeg 
      BackColor       =   &H80000000&
      Height          =   300
      ItemData        =   "HPCS.frx":160D
      Left            =   8520
      List            =   "HPCS.frx":1629
      TabIndex        =   18
      Text            =   "16"
      ToolTipText     =   "指定开方结果的有效数字 (最后一位不精确)"
      Top             =   4200
      Width           =   615
   End
   Begin VB.CommandButton cs 
      Caption         =   "="
      Height          =   255
      Left            =   240
      TabIndex        =   13
      Top             =   4440
      Width           =   735
   End
   Begin VB.CommandButton ca 
      Caption         =   "="
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   840
      Width           =   735
   End
   Begin VB.CommandButton cd 
      Caption         =   "="
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   3240
      Width           =   735
   End
   Begin VB.TextBox sa 
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
      Height          =   645
      Left            =   1080
      MousePointer    =   3  'I-Beam
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   14
      Top             =   4200
      Width           =   7095
   End
   Begin VB.TextBox s1 
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
      Height          =   285
      Left            =   1080
      MousePointer    =   3  'I-Beam
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   12
      Top             =   3840
      Width           =   7095
   End
   Begin VB.TextBox da 
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
      Height          =   285
      Left            =   1080
      MousePointer    =   3  'I-Beam
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   11
      Top             =   3240
      Width           =   7095
   End
   Begin VB.TextBox d2 
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
      Height          =   285
      Left            =   1080
      MousePointer    =   3  'I-Beam
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   9
      Top             =   2880
      Width           =   7095
   End
   Begin VB.TextBox d1 
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
      Height          =   285
      Left            =   1080
      MousePointer    =   3  'I-Beam
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   8
      Top             =   2520
      Width           =   7095
   End
   Begin VB.CommandButton cm 
      Caption         =   "="
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   2040
      Width           =   735
   End
   Begin VB.TextBox ma 
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
      Height          =   285
      Left            =   1080
      MousePointer    =   3  'I-Beam
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   7
      Top             =   2040
      Width           =   7095
   End
   Begin VB.TextBox m2 
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
      Height          =   285
      Left            =   1080
      MousePointer    =   3  'I-Beam
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Top             =   1680
      Width           =   7095
   End
   Begin VB.TextBox m1 
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
      Height          =   285
      Left            =   1080
      MousePointer    =   3  'I-Beam
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   1320
      Width           =   7095
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
      Height          =   285
      Left            =   1080
      MousePointer    =   3  'I-Beam
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   840
      Width           =   7095
   End
   Begin VB.TextBox a2 
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
      Height          =   285
      Left            =   1080
      MousePointer    =   3  'I-Beam
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   480
      Width           =   7095
   End
   Begin VB.TextBox a1 
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
      Height          =   285
      Left            =   1080
      MousePointer    =   3  'I-Beam
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   120
      Width           =   7095
   End
   Begin VB.Label Label4 
      Caption         =   "√￣"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   720
      TabIndex        =   22
      Top             =   3720
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "÷"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   21
      Top             =   2760
      Width           =   375
   End
   Begin VB.Label Label2 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   20
      Top             =   360
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "×"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   19
      Top             =   1560
      Width           =   375
   End
End
Attribute VB_Name = "hp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub ca_Click()
aa.Text = ""
a1.ToolTipText = a1.Text
a2.ToolTipText = a2.Text
aa.ToolTipText = aa.Text
aa.Text = Add(a1.Text, a2.Text)
End Sub

Private Sub cd_Click()
da.Text = ""
d1.ToolTipText = d1.Text
d2.ToolTipText = d2.Text
da.ToolTipText = da.Text
If Val(d2.Text) = 0 Then Exit Sub
If Val(d1.Text) > Val(d2.Text) Then
  msg$ = "被除数比除数大。" & Chr(13) & "因此您所要求的操作可能要花很长时间才能完成。" & Chr(13) & "您要继续高精度计算吗?"
  Style = vbYesNo + vbQuestion + vbDefaultButton1
  msg = MsgBox(msg$, Style, "计算器")
  If msg = 6 Then
    hp.Caption = "高精度计算器    - [正在计算...]"
    da.Text = Dvs(d1.Text, d2.Text, Val(clearzero.Text), Val(lendvs.Text))
    hp.Caption = "高精度计算器"
    Exit Sub
  Else
    da.Text = Str(Val(d1.Text) / Val(d2.Text))
    Exit Sub
  End If
End If

  hp.Caption = "高精度计算器    - [正在计算...]"
  da.Text = Dvs(d1.Text, d2.Text, Val(clearzero.Text), Val(lendvs.Text))
  hp.Caption = "高精度计算器"

End Sub

Private Sub cm_Click()
ma.Text = ""
m1.ToolTipText = m1.Text
m2.ToolTipText = m2.Text
ma.ToolTipText = ma.Text
ma.Text = Mpc(m1.Text, m2.Text)
End Sub

Private Sub Command1_Click()
Calculator.Show
End Sub

Private Sub cs_Click()
sa.Text = ""
hp.Caption = "高精度计算器    - [正在计算...]"
s1.ToolTipText = s1.Text
sa.ToolTipText = sa.Text
sa.Text = Sqrt(s1, Val(sqrtdeg.Text))
hp.Caption = "高精度计算器"
End Sub

