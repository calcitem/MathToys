VERSION 5.00
Begin VB.Form polygon 
   BackColor       =   &H80000000&
   Caption         =   "��n���ε��йؼ���  [n Ϊ��Ҫ, a��r��R ��ѡ����һ��ֵ]"
   ClientHeight    =   3735
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8070
   Icon            =   "polygon.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3735
   ScaleWidth      =   8070
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton Command1 
      Caption         =   "���"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6000
      TabIndex        =   4
      Top             =   3120
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4680
      TabIndex        =   5
      Top             =   3120
      Width           =   1095
   End
   Begin VB.TextBox n 
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   3360
      TabIndex        =   0
      ToolTipText     =   "����(��������)"
      Top             =   240
      Width           =   3735
   End
   Begin VB.TextBox a 
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   3360
      TabIndex        =   1
      ToolTipText     =   "�߳�(ѡ������)"
      Top             =   840
      Width           =   3735
   End
   Begin VB.TextBox r1 
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3360
      TabIndex        =   2
      ToolTipText     =   "���ľ�(ѡ������)"
      Top             =   1320
      Width           =   3735
   End
   Begin VB.TextBox r2 
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3360
      TabIndex        =   3
      ToolTipText     =   "���Բ�뾶(ѡ������)"
      Top             =   1830
      Width           =   3735
   End
   Begin VB.TextBox s 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3360
      TabIndex        =   6
      ToolTipText     =   "���(��������)"
      Top             =   2400
      Width           =   3735
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "         ���� n ="
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   840
      TabIndex        =   11
      Top             =   240
      Width           =   2415
   End
   Begin VB.Label la 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "         �߳� a ="
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   840
      TabIndex        =   10
      Top             =   840
      Width           =   2415
   End
   Begin VB.Label lb 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "����Բ�뾶 r ="
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   910
      TabIndex        =   9
      ToolTipText     =   "����Բ�뾶(ѡ������)"
      Top             =   1320
      Width           =   2415
   End
   Begin VB.Label lc 
      BackColor       =   &H80000000&
      Caption         =   "���Բ�뾶 R ="
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   840
      TabIndex        =   8
      ToolTipText     =   "���Բ�뾶"
      Top             =   1800
      Width           =   2415
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000000&
      Caption         =   "         ��� S ="
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   860
      TabIndex        =   7
      ToolTipText     =   "(��������)"
      Top             =   2400
      Width           =   2415
   End
End
Attribute VB_Name = "polygon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()
On Error GoTo 20
PI = 3.14159265358979
polygon.Caption = "��n���ε��йؼ��� "
n.text = Fix(Abs(Val(n.text)))
If Val(n.text) < 3 Then n.text = "3"
th = PI / Val(n.text)
n = Val(n.text)
a = Val(a.text)
r1 = Val(r1.text)
r2 = Val(r2.text)
s = Val(s.text)

If a <> 0 Then
  s = n * a ^ 2 * (1 / Tan(th)) / 4
  r1 = a * (1 / Tan(th)) / 2
  r2 = a * (1 / Sin(th)) / 2
  GoTo 10
End If


If r1 <> 0 Then
  s = n * r1 ^ 2 * Tan(th)
  a = 2 * r1 * Tan(th)
  r2 = r1 * (1 / Cos(th))
  GoTo 10
End If

If r2 <> 0 Then
  s = n * r2 ^ 2 * Sin(2 * th) / 2
  a = 2 * r2 * Sin(th)
  r1 = r2 * Cos(th)
  GoTo 10
End If


10:
a.text = Str(a)
r1.text = Str(r1)
r2.text = Str(r2)
s.text = Str(s)

If Left(a.text, 1) = " " Then a.text = Right(a.text, Len(a.text) - 1)
If Left(r1.text, 1) = " " Then r1.text = Right(r1.text, Len(r1.text) - 1)
If Left(r2.text, 1) = " " Then r2.text = Right(r2.text, Len(r2.text) - 1)
If Left(s.text, 1) = " " Then s.text = Right(s.text, Len(s.text) - 1)

If Left(a.text, 1) = "." Then a.text = "0" + a.text
If Left(r1.text, 1) = "." Then r1.text = "0" + r1.text
If Left(r2.text, 1) = "." Then r2.text = "0" + r2.text
If Left(s.text, 1) = "." Then s.text = "0" + s.text

Exit Sub

20:
  
 msg = MsgBox("�������޷����һ��������", vbOKOnly, "����")
Resume Next
End Sub

Private Sub Command2_Click()
n.text = ""
r1.text = ""
r2.text = ""
a.text = ""
s.text = ""
End Sub
