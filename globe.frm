VERSION 5.00
Begin VB.Form Globe 
   Caption         =   "��������������"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6195
   ForeColor       =   &H00000000&
   Icon            =   "globe.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   6195
   StartUpPosition =   3  '����ȱʡ
   Begin VB.OptionButton degree 
      Caption         =   "��"
      Height          =   255
      Left            =   1080
      TabIndex        =   13
      Top             =   2640
      Value           =   -1  'True
      Width           =   615
   End
   Begin VB.OptionButton dfm 
      Caption         =   "��-��-��"
      Height          =   255
      Left            =   2280
      TabIndex        =   12
      Top             =   2640
      Width           =   1095
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   375
      Index           =   1
      Left            =   4440
      TabIndex        =   3
      ToolTipText     =   "����2"
      Top             =   960
      Width           =   1300
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   375
      Index           =   0
      Left            =   4440
      TabIndex        =   2
      ToolTipText     =   "γ��2"
      Top             =   360
      Width           =   1300
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   5
      ToolTipText     =   "�ԡ��ס�Ϊ��λ"
      Top             =   1560
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Index           =   1
      Left            =   1560
      TabIndex        =   1
      ToolTipText     =   "����1"
      Top             =   960
      Width           =   1300
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Index           =   0
      Left            =   1560
      TabIndex        =   0
      ToolTipText     =   "γ��1"
      Top             =   360
      Width           =   1300
   End
   Begin VB.Label Label7 
      Caption         =   "����Ϊ��, ����Ϊ��; ��γΪ��, ��γΪ��"
      ForeColor       =   &H008080FF&
      Height          =   255
      Left            =   480
      TabIndex        =   11
      Top             =   2160
      Width           =   3615
   End
   Begin VB.Label Label4 
      Caption         =   "γ�ȣ�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Index           =   1
      Left            =   3360
      TabIndex        =   10
      Top             =   360
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "���ȣ�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Index           =   0
      Left            =   3360
      TabIndex        =   9
      Top             =   960
      Width           =   975
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   4680
      TabIndex        =   4
      ToolTipText     =   "��ʼ����, ����������Ϊ��λ"
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "������룺"
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   375
      Index           =   0
      Left            =   480
      TabIndex        =   8
      ToolTipText     =   "��"
      Top             =   1560
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "γ�ȣ�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   375
      Index           =   1
      Left            =   480
      TabIndex        =   7
      Top             =   360
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "���ȣ�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   375
      Index           =   0
      Left            =   480
      TabIndex        =   6
      Top             =   960
      Width           =   975
   End
End
Attribute VB_Name = "Globe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit

Const PI As Double = 3.14159265358979
Const r As Double = 6371004 '����뾶

Private Sub Label3_Click()
    ReDim P1(1) As Double, P2(1) As Double
    
    P1(0) = Val(Text1(0).text): P1(1) = Val(Text1(1).text)
    P2(0) = Val(Text3(0).text): P2(1) = Val(Text3(1).text)
    
    If dfm.Value = True Then
      P1(0) = Deg(P1(0))
      P2(0) = Deg(P2(0))
      P1(1) = Deg(P1(1))
      P2(1) = Deg(P2(1))
    End If
    
    
    
    If Abs(P1(1)) > 180 Or Abs(P2(1)) > 180 Or Abs(P1(0)) > 90 Or Abs(P2(0)) > 90 Then
      msg = MsgBox("�Ƿ��ľ�γ�ȡ�", , "������")
      'msg = MsgBox("����FTP����" & Chr(13) & "����������ϵ����˴���", vbCritical, "����")
      Exit Sub
    End If
    
    Text2.text = Fix(PTP(P1(), P2()))
    
    l = Len(Text2.text)
    If l > 6 Then
      Text2.text = Left(Text2.text, l - 6) + "," + Left(Right(Text2.text, 6), 3) + "," + Right(Text2.text, 3)
    Else
      If l > 3 Then
        Text2.text = Left(Text2.text, l - 3) + "," + Right(Text2.text, 3)
      End If
    End If
    
End Sub



'
'#####################################
'----------������֮���������---------
'-------------v1.00-------------------
'-------------vb6.0����ͨ��-----------
'-------------Written by DRD----------
'#####################################
'



'Option Explicit

Public Function PTP(P1() As Double, P2() As Double) As Double
    '
    '-------------------------------------
    '-------------��������----------------
    '--------P1() ��һ���γ�ȡ����ȶ�----
    '--------P2() �ڶ����γ�ȡ����ȶ�----
    '-------------------------------------
    On Error GoTo ftperr
    Dim Angle1 As Double, Angle2 As Double
    Dim LLength As Double
    ReDim Edge1(1) As Double, Edge2(1) As Double, Edge3(1) As Double
    
    '���ȼн�
    Angle1 = P2(1) - P1(1)
    If Angle1 > 180 Then Angle1 = 360 - Angle1
    '��Ϊ����
    P1(0) = PI * P1(0) / 180
    P2(0) = -PI * P2(0) / 180
    Angle1 = PI * Angle1 / 180
    
    '�������α߳�
    Edge1(0) = Abs(r * Cos(P1(0)))
    Edge1(1) = r * Sin(P1(0))
    
    Edge2(0) = Abs(r * Cos(P2(0)))
    Edge2(1) = r * Sin(P2(0))
    
    Edge3(0) = r ^ 2 * (Cos(P1(0)) ^ 2 + Cos(P2(0)) ^ 2)
    Edge3(0) = Edge3(0) - 2 * Edge1(0) * Edge2(0) * Cos(Angle1)
    Edge3(1) = (Edge1(1) + Edge2(1)) ^ 2
    
    '��ֱ�߾���
    LLength = Edge3(1) + Edge3(0)
    '�������Բ�н�
    Angle2 = FunArccos(1 - LLength / (2 * r ^ 2))
    '�������Բ����
    PTP = r * Angle2
    Exit Function
ftperr:

msg = MsgBox("����FTP����" & Chr(13) _
& "����������ϵ����˴���", vbCritical, "����")
Resume Next
End Function

Private Function FunArccos(ByVal CosValue As Double) As Double
    '
    '----------------------------------
    '-----------������---------------
    '----------------------------------
    '
     On Error GoTo acoserr
    If CosValue = 1 Then
      FunArccos = 0
      Exit Function
    Else
       If CosValue = -1 Then
         FunArccos = PI
         Exit Function
       End If
    End If
    FunArccos = -Atn(CosValue / Sqr(1 - CosValue ^ 2)) + PI / 2
    ' n# = pi# / 2 - Atn(n# / Sqr(1 - n# ^ 2))
Exit Function
acoserr:

msg = MsgBox("����Acos����" & Chr(13) _
& "����������ϵ����˴���", vbCritical, "����")
Resume Next
End Function


