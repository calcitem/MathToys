VERSION 5.00
Begin VB.Form Prime 
   Caption         =   "�ֽ�������"
   ClientHeight    =   4230
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5430
   Icon            =   "Prime.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4230
   ScaleWidth      =   5430
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton Command2 
      Caption         =   "ֹͣ"
      Height          =   375
      Left            =   2520
      Picture         =   "Prime.frx":08CA
      TabIndex        =   2
      ToolTipText     =   "ֹͣ�ֽ�������"
      Top             =   3480
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "��ʼ"
      Default         =   -1  'True
      Height          =   375
      Left            =   3840
      Picture         =   "Prime.frx":1194
      TabIndex        =   1
      ToolTipText     =   "��ʼ�ֽ�������"
      Top             =   3480
      Width           =   855
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   480
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   1200
      Width           =   4455
   End
   Begin VB.TextBox num 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   840
      TabIndex        =   0
      Top             =   360
      Width           =   3735
   End
End
Attribute VB_Name = "Prime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Running As Boolean

'Private Sub Command1_Click()
'On Error GoTo 10
'running = True
'Text2.Text = ""
'p$ = ""
'n = Val(num.Text)
'm = n
'Prime.Caption = "�ֽ�������    - [���ڷֽ�...]"
'For i = 2 To m
'  DoEvents
'  Do Until n Mod i <> 0
'    n = n / i
'    k = k + 1
'  Loop
  
'  If k <> 0 Then
'    If k > 1 Then p$ = p$ + Str(i) + " ^" + Str(k) + " *" Else p$ = p$ + Str(i) + " *"
'    Text2.Text = p$
'  End If
'  k = 0
'  If n = 1 Then Exit For
  
'  If running = False Then Exit For
'Next i
'Text2.Text = Left(p$, Len(p$) - 1)
'If running = False Then Text2.Text = Text2.Text + "* ..."
'running = False
'Prime.Caption = "�ֽ�������"

'Exit Sub
'10:
'msg = MsgBox("�������޷���ɼ��㡣", vbOKOnly, "����")
'Resume Next
'End Sub

Private Sub Command2_Click()

Running = False

End Sub
'void main()
'{ long double n,i;
'  int k;
'  cout.precision(15);
'  cin>>n;
'  i=2;
'  While (i <= Sqrt(n))
'  { k=0;
  
'    while (floor(n/i)*i==n)
'    { n/=i;
'      k++;
'          };
'    if (k>0) cout<<i<<'^'<<k<<"   ";
'    i++;
'  };
'  if (n>1) cout<<n<<'^'<<1<<endl;
  
'}
Private Sub Command1_Click()


On Error GoTo 10
Dim n, i, k As Double
If Len(num.Text) >= 16 Then msg = MsgBox("    �����", vbOKOnly, "����"): Exit Sub
Prime.Caption = "�ֽ�������    - [���ڷֽ�...]"
Running = True
Text2.Text = ""
p$ = ""
n = Val(num.Text)

i = 2
Do Until i > Sqr(n)
  k = 0
  While Fix(n / i) * i = n
    DoEvents
    n = n / i
    k = k + 1
  Wend
  If k > 0 Then Text2.Text = Text2.Text & i & "^" & k & "*"
  i = i + 1
  DoEvents
  If Running = False Then Exit Do
Loop
If n > 1 Then Text2.Text = Text2.Text & n & "^1*"
Text2.Text = Left(Text2.Text, Len(Text2.Text) - 1)
If Running = False Then Text2.Text = Text2.Text + "* ..."
Running = False
Prime.Caption = "�ֽ�������"
Exit Sub
10:
msg = MsgBox("�������޷���ɼ��㡣", vbOKOnly, "����")
Resume 20
20: Prime.Caption = "�ֽ�������    - [�жϷֽ�]"
End Sub
