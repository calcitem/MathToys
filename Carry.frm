VERSION 5.00
Begin VB.Form Carry 
   Caption         =   "��λת����"
   ClientHeight    =   3165
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4650
   Icon            =   "Carry.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3165
   ScaleWidth      =   4650
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton Command4 
      Caption         =   "q��10 ת����"
      Height          =   495
      Left            =   2040
      TabIndex        =   6
      ToolTipText     =   "�������q������ת��Ϊ10��������ʾ������"
      Top             =   1920
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "10��q ת����"
      Height          =   495
      Left            =   2040
      TabIndex        =   3
      ToolTipText     =   "�������10������ת��Ϊq��������ʾ������"
      Top             =   720
      Width           =   1335
   End
   Begin VB.ComboBox q1 
      BackColor       =   &H00E0E0E0&
      Height          =   300
      ItemData        =   "Carry.frx":08CA
      Left            =   3600
      List            =   "Carry.frx":0937
      TabIndex        =   8
      Text            =   "16"
      ToolTipText     =   "����������q��ֵ,����ʾ������ֵĽ�λ�ƵĻ�"
      Top             =   2640
      Width           =   615
   End
   Begin VB.CommandButton Command2 
      Caption         =   "10��q ת����"
      Height          =   495
      Left            =   360
      TabIndex        =   5
      ToolTipText     =   "�������10������ת��Ϊq��������ʾ������"
      Top             =   1920
      Width           =   1335
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00E0E0E0&
      Height          =   495
      Left            =   360
      MultiLine       =   -1  'True
      TabIndex        =   7
      ToolTipText     =   "q������"
      Top             =   2520
      Width           =   3015
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00E0E0E0&
      Height          =   495
      Left            =   360
      MultiLine       =   -1  'True
      TabIndex        =   4
      ToolTipText     =   "ʮ������"
      Top             =   1320
      Width           =   3015
   End
   Begin VB.ComboBox q 
      BackColor       =   &H00E0E0E0&
      Height          =   300
      ItemData        =   "Carry.frx":09C7
      Left            =   3600
      List            =   "Carry.frx":0A34
      TabIndex        =   1
      Text            =   "16"
      ToolTipText     =   "����������q��ֵ,����ʾ������ֵĽ�λ�ƵĻ�"
      Top             =   240
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00B97764&
      Caption         =   "q��10 ת����"
      Height          =   495
      Left            =   360
      TabIndex        =   2
      ToolTipText     =   "�������q������ת��Ϊ10��������ʾ������"
      Top             =   720
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00E0E0E0&
      Height          =   495
      Left            =   360
      MultiLine       =   -1  'True
      TabIndex        =   0
      ToolTipText     =   "q������"
      Top             =   120
      Width           =   3015
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "( 10 )"
      Height          =   255
      Left            =   3600
      TabIndex        =   9
      ToolTipText     =   "������ֵĽ�λ�ƵĻ�Ϊ10"
      Top             =   1560
      Width           =   735
   End
End
Attribute VB_Name = "Carry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Text2.text = qto10(Text1.text, Val(q.text))
End Sub

Private Sub Command2_Click()
On Error Resume Next
Text3.text = dtoq(Val(Text2.text), Val(q1.text))

End Sub

Private Sub Command3_Click()
On Error Resume Next
Text1.text = dtoq(Val(Text2.text), Val(q.text))
End Sub

Private Sub Command4_Click()
Text2.text = qto10(Text3.text, Val(q1.text))
End Sub
