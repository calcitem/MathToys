VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Begin VB.Form Calc 
   BackColor       =   &H00699885&
   Caption         =   "��ѧ������"
   ClientHeight    =   5925
   ClientLeft      =   60
   ClientTop       =   510
   ClientWidth     =   9675
   BeginProperty Font 
      Name            =   "����"
      Size            =   14.25
      Charset         =   134
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   Icon            =   "Calc.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   5925
   ScaleWidth      =   9675
   StartUpPosition =   1  '����������
   Begin VB.PictureBox wrong 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00699885&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   9160
      Picture         =   "Calc.frx":0CCA
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   82
      ToolTipText     =   "���ʽ�д���, �����˴��޸ı��ʽ��"
      Top             =   840
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H00739D8C&
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
      ForeColor       =   &H00184A00&
      Height          =   330
      Left            =   480
      MousePointer    =   4  'Icon
      ScrollBars      =   2  'Vertical
      TabIndex        =   18
      Text            =   "                                   Ubiquitous Computing    �������ڵļ���"
      ToolTipText     =   "���ʽ��ʾ��"
      Top             =   1440
      Width           =   8655
   End
   Begin VB.TextBox fsh 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00739D8C&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H002D4238&
      Height          =   270
      Left            =   6720
      MaxLength       =   1024
      MousePointer    =   3  'I-Beam
      MultiLine       =   -1  'True
      TabIndex        =   17
      TabStop         =   0   'False
      ToolTipText     =   "С�����ֻ�Ϊ����(����)"
      Top             =   240
      Width           =   2415
   End
   Begin VB.PictureBox Advance 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   9480
      MouseIcon       =   "Calc.frx":1594
      Picture         =   "Calc.frx":189E
      ScaleHeight     =   480
      ScaleWidth      =   150
      TabIndex        =   16
      ToolTipText     =   "�л����߼�����׼ģʽ"
      Top             =   2760
      Width           =   150
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   11400
      Top             =   7800
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton memory 
      BackColor       =   &H0097B5A7&
      Caption         =   "MR"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   3
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "��-�洢���е���"
      Top             =   2520
      Width           =   495
   End
   Begin VB.CommandButton memory 
      BackColor       =   &H0097B5A7&
      Caption         =   "ML"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   2
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "��-�ϴμ�����"
      Top             =   2520
      Width           =   495
   End
   Begin VB.CommandButton memory 
      BackColor       =   &H0097B5A7&
      Caption         =   "M"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   1
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "��-��ǰ������"
      Top             =   2520
      Width           =   495
   End
   Begin VB.CommandButton Msave 
      BackColor       =   &H0097B5A7&
      Caption         =   "MS"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "����ʾ���ִ���洢��, Ȼ�����Ϳ����ڱ��ʽ���� (mr) ��ʾ�����"
      Top             =   2520
      Width           =   615
   End
   Begin MSComctlLib.Slider Slider1 
      Height          =   255
      Left            =   9720
      TabIndex        =   8
      ToolTipText     =   "���㾫��(λ)"
      Top             =   720
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   450
      _Version        =   393216
      LargeChange     =   2
      Min             =   -1
      Max             =   7
      SelStart        =   -1
      TickStyle       =   1
      Value           =   -1
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H00699885&
      Caption         =   "��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   300
      Left            =   3960
      TabIndex        =   7
      ToolTipText     =   "�����Ǻ�����������Ϊ����"
      Top             =   2520
      Width           =   615
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00699885&
      Caption         =   "��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3240
      TabIndex        =   6
      ToolTipText     =   "�����Ǻ�����������Ϊ�Ƕ�"
      Top             =   2520
      Value           =   -1  'True
      Width           =   615
   End
   Begin VB.TextBox Text6 
      Appearance      =   0  'Flat
      BackColor       =   &H00739D8C&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00597B68&
      Height          =   330
      Left            =   360
      MaxLength       =   1024
      MousePointer    =   4  'Icon
      MultiLine       =   -1  'True
      TabIndex        =   4
      ToolTipText     =   "��һ�����ʽ"
      Top             =   6000
      Width           =   8895
   End
   Begin VB.TextBox Text5 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00739D8C&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H002D4238&
      Height          =   270
      Left            =   1080
      MaxLength       =   1024
      MousePointer    =   3  'I-Beam
      MultiLine       =   -1  'True
      TabIndex        =   3
      TabStop         =   0   'False
      ToolTipText     =   "�洢��[MR]"
      Top             =   240
      Width           =   2415
   End
   Begin VB.TextBox Text4 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00739D8C&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H002D4238&
      Height          =   270
      Left            =   3840
      MaxLength       =   1024
      MousePointer    =   3  'I-Beam
      MultiLine       =   -1  'True
      TabIndex        =   2
      TabStop         =   0   'False
      ToolTipText     =   "�ϴμ�����[ML]"
      Top             =   240
      Width           =   2415
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      BackColor       =   &H00739D8C&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "����"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00597B68&
      Height          =   420
      Left            =   360
      MaxLength       =   1024
      MousePointer    =   4  'Icon
      MultiLine       =   -1  'True
      TabIndex        =   5
      ToolTipText     =   "������ʱ  ���ڡ�ʱ��"
      Top             =   6480
      Width           =   8895
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00B5CCC2&
      BeginProperty Font 
         Name            =   "����"
         Size            =   32.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H002D4238&
      Height          =   735
      Left            =   480
      MaxLength       =   1024
      MousePointer    =   3  'I-Beam
      TabIndex        =   1
      TabStop         =   0   'False
      Text            =   "��ӭʹ��  "
      ToolTipText     =   "��������ʾ��  ��ʾ��ǰ������[M]"
      Top             =   600
      Width           =   8655
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00699885&
      ForeColor       =   &H00C0C0C0&
      Height          =   2775
      Left            =   4320
      TabIndex        =   14
      Top             =   2760
      Width           =   4935
      Begin VB.CommandButton Command 
         BackColor       =   &H0092B1A3&
         Caption         =   "dms"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   25
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   80
         ToolTipText     =   "���öȱ�ʾ������ת��Ϊ��""��-��-��""��ʽ��ʾ�ĸ�ʽ"
         Top             =   360
         Width           =   495
      End
      Begin VB.CommandButton Command 
         Appearance      =   0  'Flat
         BackColor       =   &H0092B1A3&
         Caption         =   "`"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   31
         Left            =   960
         MaskColor       =   &H00000000&
         TabIndex        =   79
         ToolTipText     =   "�ָ�����������"
         Top             =   840
         Width           =   255
      End
      Begin VB.CommandButton Command 
         BackColor       =   &H0092B1A3&
         Caption         =   "log"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   30
         Left            =   720
         Style           =   1  'Graphical
         TabIndex        =   78
         ToolTipText     =   "����"
         Top             =   840
         Width           =   495
      End
      Begin VB.CommandButton Command 
         BackColor       =   &H0092B1A3&
         Caption         =   "deg"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   34
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   77
         ToolTipText     =   "����""��-��-��""��ʽ��ʾ������ת��Ϊ�öȱ�ʾ�ĸ�ʽ"
         Top             =   840
         Width           =   495
      End
      Begin VB.CommandButton result 
         BackColor       =   &H0086AA99&
         Caption         =   "="
         Default         =   -1  'True
         BeginProperty Font 
            Name            =   "����"
            Size            =   15.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   4320
         MouseIcon       =   "Calc.frx":1CE0
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   76
         ToolTipText     =   "��ʼ����"
         Top             =   1800
         Width           =   495
      End
      Begin VB.CommandButton Command 
         BackColor       =   &H00A5C0B4&
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   50
         Left            =   3120
         Style           =   1  'Graphical
         TabIndex        =   75
         ToolTipText     =   "������5���ڱ��ʽ������"
         Top             =   1320
         Width           =   495
      End
      Begin VB.CommandButton Command 
         BackColor       =   &H00A5C0B4&
         Caption         =   "6"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   51
         Left            =   3720
         Style           =   1  'Graphical
         TabIndex        =   74
         ToolTipText     =   "������6���ڱ��ʽ������"
         Top             =   1320
         Width           =   495
      End
      Begin VB.CommandButton Command 
         BackColor       =   &H00A5C0B4&
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   49
         Left            =   2520
         Style           =   1  'Graphical
         TabIndex        =   73
         ToolTipText     =   "������4���ڱ��ʽ������"
         Top             =   1320
         Width           =   495
      End
      Begin VB.CommandButton Command 
         BackColor       =   &H00A5C0B4&
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   46
         Left            =   2520
         Style           =   1  'Graphical
         TabIndex        =   72
         ToolTipText     =   "������1���ڱ��ʽ������"
         Top             =   1800
         Width           =   495
      End
      Begin VB.CommandButton Command 
         BackColor       =   &H00A5C0B4&
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   47
         Left            =   3120
         Style           =   1  'Graphical
         TabIndex        =   71
         ToolTipText     =   "������2���ڱ��ʽ������"
         Top             =   1800
         Width           =   495
      End
      Begin VB.CommandButton Command 
         BackColor       =   &H00A5C0B4&
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   48
         Left            =   3720
         Style           =   1  'Graphical
         TabIndex        =   70
         ToolTipText     =   "������3���ڱ��ʽ������"
         Top             =   1800
         Width           =   495
      End
      Begin VB.CommandButton Command 
         BackColor       =   &H00A5C0B4&
         Caption         =   "7"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   52
         Left            =   2520
         Style           =   1  'Graphical
         TabIndex        =   69
         ToolTipText     =   "������7���ڱ��ʽ������"
         Top             =   840
         Width           =   495
      End
      Begin VB.CommandButton Command 
         BackColor       =   &H00A5C0B4&
         Caption         =   "8"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   53
         Left            =   3120
         Style           =   1  'Graphical
         TabIndex        =   68
         ToolTipText     =   "������8���ڱ��ʽ������"
         Top             =   840
         Width           =   495
      End
      Begin VB.CommandButton Command 
         BackColor       =   &H00A5C0B4&
         Caption         =   "9"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   54
         Left            =   3720
         Style           =   1  'Graphical
         TabIndex        =   67
         ToolTipText     =   "������9���ڱ��ʽ������"
         Top             =   840
         Width           =   495
      End
      Begin VB.CommandButton Cmod 
         BackColor       =   &H009AB8AA&
         Caption         =   "mod"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1920
         Style           =   1  'Graphical
         TabIndex        =   66
         ToolTipText     =   "ȡģ  ������"
         Top             =   2280
         Width           =   495
      End
      Begin VB.CommandButton Zero 
         BackColor       =   &H00A5C0B4&
         Caption         =   "0        "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2520
         Style           =   1  'Graphical
         TabIndex        =   65
         ToolTipText     =   "������0���ڱ��ʽ������"
         Top             =   2280
         Width           =   1095
      End
      Begin VB.CommandButton Command 
         BackColor       =   &H00A5C0B4&
         Caption         =   "."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   45
         Left            =   3720
         Style           =   1  'Graphical
         TabIndex        =   64
         ToolTipText     =   "����С����"
         Top             =   2280
         Width           =   495
      End
      Begin VB.CommandButton Command 
         BackColor       =   &H009AB8AA&
         Caption         =   "\"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   39
         Left            =   1920
         Style           =   1  'Graphical
         TabIndex        =   63
         ToolTipText     =   "����"
         Top             =   1800
         Width           =   495
      End
      Begin VB.CommandButton Command 
         BackColor       =   &H009AB8AA&
         Caption         =   "/"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   41
         Left            =   3120
         Style           =   1  'Graphical
         TabIndex        =   62
         ToolTipText     =   "����  ��"
         Top             =   360
         Width           =   495
      End
      Begin VB.CommandButton Command 
         BackColor       =   &H009AB8AA&
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   42
         Left            =   3720
         Style           =   1  'Graphical
         TabIndex        =   61
         ToolTipText     =   "�˺�"
         Top             =   360
         Width           =   495
      End
      Begin VB.CommandButton Command 
         BackColor       =   &H009AB8AA&
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   43
         Left            =   4320
         Style           =   1  'Graphical
         TabIndex        =   60
         ToolTipText     =   "��������"
         Top             =   360
         Width           =   495
      End
      Begin VB.CommandButton Command 
         BackColor       =   &H009AB8AA&
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   44
         Left            =   4320
         Style           =   1  'Graphical
         TabIndex        =   59
         ToolTipText     =   "�ӡ�����"
         Top             =   840
         Width           =   495
      End
      Begin VB.CommandButton Command 
         BackColor       =   &H009AB8AA&
         Caption         =   "^"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   38
         Left            =   1920
         Style           =   1  'Graphical
         TabIndex        =   58
         ToolTipText     =   "�˷�"
         Top             =   840
         Width           =   495
      End
      Begin VB.CommandButton Sqrt 
         BackColor       =   &H009AB8AA&
         Caption         =   "�̣� "
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
         Left            =   1920
         Style           =   1  'Graphical
         TabIndex        =   57
         ToolTipText     =   "����  ���� (�ڱ��ʽ����@��ʾ)"
         Top             =   1320
         Width           =   495
      End
      Begin VB.CommandButton Command 
         BackColor       =   &H0092B1A3&
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   37
         Left            =   1320
         Style           =   1  'Graphical
         TabIndex        =   56
         ToolTipText     =   "�ٷֺ�"
         Top             =   840
         Width           =   495
      End
      Begin VB.CommandButton circle 
         BackColor       =   &H0092B1A3&
         Caption         =   "��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1320
         Style           =   1  'Graphical
         TabIndex        =   55
         ToolTipText     =   "��-Բ���� �ڱ��ʽ����pi��ʾ)"
         Top             =   2280
         Width           =   495
      End
      Begin VB.CommandButton Command 
         BackColor       =   &H0092B1A3&
         Caption         =   "abs"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   24
         Left            =   1320
         Style           =   1  'Graphical
         TabIndex        =   54
         ToolTipText     =   "����ֵ"
         Top             =   1320
         Width           =   495
      End
      Begin VB.CommandButton Command 
         BackColor       =   &H0092B1A3&
         Caption         =   "exp"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   29
         Left            =   720
         Style           =   1  'Graphical
         TabIndex        =   53
         ToolTipText     =   "ָ������   e��ĳ�η�"
         Top             =   360
         Width           =   495
      End
      Begin VB.CommandButton Command 
         BackColor       =   &H0092B1A3&
         Caption         =   "lg"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   32
         Left            =   720
         Style           =   1  'Graphical
         TabIndex        =   52
         ToolTipText     =   "���ö�������10Ϊ�׵Ķ�����"
         Top             =   1320
         Width           =   495
      End
      Begin VB.CommandButton Command 
         BackColor       =   &H0092B1A3&
         Caption         =   "ln"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   33
         Left            =   720
         Style           =   1  'Graphical
         TabIndex        =   51
         ToolTipText     =   "��Ȼ��������eΪ�׵Ķ�����"
         Top             =   1800
         Width           =   495
      End
      Begin VB.CommandButton Leftparenthesis 
         BackColor       =   &H00A5C0B4&
         Caption         =   "("
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1920
         Style           =   1  'Graphical
         TabIndex        =   50
         ToolTipText     =   "������"
         Top             =   360
         Width           =   495
      End
      Begin VB.CommandButton Rightparenthesis 
         BackColor       =   &H00A5C0B4&
         Caption         =   ")"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2520
         Style           =   1  'Graphical
         TabIndex        =   49
         ToolTipText     =   "������"
         Top             =   360
         Width           =   495
      End
      Begin VB.CommandButton Command 
         BackColor       =   &H0092B1A3&
         Caption         =   "fix"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   27
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   48
         ToolTipText     =   "��������"
         Top             =   1320
         Width           =   495
      End
      Begin VB.CommandButton Command 
         BackColor       =   &H0092B1A3&
         Caption         =   "int"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   28
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   47
         ToolTipText     =   "������x���������"
         Top             =   1800
         Width           =   495
      End
      Begin VB.CommandButton Command 
         BackColor       =   &H0092B1A3&
         Caption         =   "sgn"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   26
         Left            =   1320
         Style           =   1  'Graphical
         TabIndex        =   46
         ToolTipText     =   "���ź���"
         Top             =   1800
         Width           =   495
      End
      Begin VB.CommandButton factorial 
         BackColor       =   &H0092B1A3&
         Caption         =   "n !"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   720
         Style           =   1  'Graphical
         TabIndex        =   45
         ToolTipText     =   "�׳�"
         Top             =   2280
         Width           =   495
      End
      Begin VB.CommandButton Command 
         BackColor       =   &H0092B1A3&
         Caption         =   "e"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   40
         Left            =   1320
         Style           =   1  'Graphical
         TabIndex        =   44
         ToolTipText     =   "�����־,�ȼ��� *10^"
         Top             =   360
         Width           =   495
      End
      Begin VB.CommandButton Command 
         BackColor       =   &H0092B1A3&
         Caption         =   "round"
         BeginProperty Font 
            Name            =   "����"
            Size            =   7.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   35
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   43
         ToolTipText     =   "�������뵽��λ"
         Top             =   2280
         Width           =   495
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00699885&
      Height          =   2775
      Left            =   360
      TabIndex        =   13
      Top             =   2760
      Width           =   3975
      Begin VB.CommandButton Command 
         BackColor       =   &H008BAF9F&
         Caption         =   "arch"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Index           =   1
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   42
         ToolTipText     =   "��˫������(��ֵ)"
         Top             =   645
         Width           =   855
      End
      Begin VB.CommandButton Command 
         BackColor       =   &H008BAF9F&
         Caption         =   "arth"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Index           =   2
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   41
         ToolTipText     =   "��˫������"
         Top             =   1080
         Width           =   855
      End
      Begin VB.CommandButton Command 
         BackColor       =   &H008BAF9F&
         Caption         =   "arsh"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Index           =   0
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   40
         ToolTipText     =   "��˫������"
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton Command 
         BackColor       =   &H008BAF9F&
         Caption         =   "arcth"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Index           =   3
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   39
         ToolTipText     =   "��˫������"
         Top             =   1485
         Width           =   855
      End
      Begin VB.CommandButton Command 
         BackColor       =   &H008BAF9F&
         Caption         =   "arsech"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Index           =   4
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   38
         ToolTipText     =   "��˫������(��ֵ)"
         Top             =   1920
         Width           =   855
      End
      Begin VB.CommandButton Command 
         BackColor       =   &H008BAF9F&
         Caption         =   "arcsch"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Index           =   5
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   37
         ToolTipText     =   "��˫�����"
         Top             =   2325
         Width           =   855
      End
      Begin VB.CommandButton Command 
         BackColor       =   &H008BAF9F&
         Caption         =   "sh"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Index           =   6
         Left            =   1080
         Style           =   1  'Graphical
         TabIndex        =   36
         ToolTipText     =   "˫������"
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton Command 
         BackColor       =   &H008BAF9F&
         Caption         =   "ch"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Index           =   7
         Left            =   1080
         Style           =   1  'Graphical
         TabIndex        =   35
         ToolTipText     =   "˫������"
         Top             =   645
         Width           =   855
      End
      Begin VB.CommandButton Command 
         BackColor       =   &H008BAF9F&
         Caption         =   "th"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Index           =   8
         Left            =   1080
         Style           =   1  'Graphical
         TabIndex        =   34
         ToolTipText     =   "˫������"
         Top             =   1080
         Width           =   855
      End
      Begin VB.CommandButton Command 
         BackColor       =   &H008BAF9F&
         Caption         =   "cth"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Index           =   9
         Left            =   1080
         Style           =   1  'Graphical
         TabIndex        =   33
         ToolTipText     =   "˫������"
         Top             =   1485
         Width           =   855
      End
      Begin VB.CommandButton Command 
         BackColor       =   &H008BAF9F&
         Caption         =   "sech"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Index           =   10
         Left            =   1080
         Style           =   1  'Graphical
         TabIndex        =   32
         ToolTipText     =   "˫������"
         Top             =   1920
         Width           =   855
      End
      Begin VB.CommandButton Command 
         BackColor       =   &H008BAF9F&
         Caption         =   "csch"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Index           =   11
         Left            =   1080
         Style           =   1  'Graphical
         TabIndex        =   31
         ToolTipText     =   "˫�����"
         Top             =   2325
         Width           =   855
      End
      Begin VB.CommandButton Command 
         BackColor       =   &H008BAF9F&
         Caption         =   "arcsin"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Index           =   12
         Left            =   2040
         Style           =   1  'Graphical
         TabIndex        =   30
         ToolTipText     =   "������(��ֵ)"
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton Command 
         BackColor       =   &H008BAF9F&
         Caption         =   "arccos"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Index           =   13
         Left            =   2040
         Style           =   1  'Graphical
         TabIndex        =   29
         ToolTipText     =   "������(��ֵ)"
         Top             =   645
         Width           =   855
      End
      Begin VB.CommandButton Command 
         BackColor       =   &H008BAF9F&
         Caption         =   "arctan"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Index           =   14
         Left            =   2040
         Style           =   1  'Graphical
         TabIndex        =   28
         ToolTipText     =   "������(��ֵ)"
         Top             =   1080
         Width           =   855
      End
      Begin VB.CommandButton Command 
         BackColor       =   &H008BAF9F&
         Caption         =   "arccot"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Index           =   15
         Left            =   2040
         Style           =   1  'Graphical
         TabIndex        =   27
         ToolTipText     =   "������(��ֵ)"
         Top             =   1485
         Width           =   855
      End
      Begin VB.CommandButton Command 
         BackColor       =   &H008BAF9F&
         Caption         =   "arcsec"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Index           =   16
         Left            =   2040
         Style           =   1  'Graphical
         TabIndex        =   26
         ToolTipText     =   "������(��ֵ)"
         Top             =   1920
         Width           =   855
      End
      Begin VB.CommandButton Command 
         BackColor       =   &H008BAF9F&
         Caption         =   "arccsc"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Index           =   17
         Left            =   2040
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "�����(��ֵ)"
         Top             =   2325
         Width           =   855
      End
      Begin VB.CommandButton Command 
         BackColor       =   &H008BAF9F&
         Caption         =   "sin"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Index           =   18
         Left            =   3000
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "����"
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton Command 
         BackColor       =   &H008BAF9F&
         Caption         =   "cos"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Index           =   19
         Left            =   3000
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "����"
         Top             =   645
         Width           =   855
      End
      Begin VB.CommandButton Command 
         BackColor       =   &H008BAF9F&
         Caption         =   "tan"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Index           =   20
         Left            =   3000
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "����"
         Top             =   1080
         UseMaskColor    =   -1  'True
         Width           =   855
      End
      Begin VB.CommandButton Command 
         BackColor       =   &H008BAF9F&
         Caption         =   "cot"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Index           =   21
         Left            =   3000
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "����"
         Top             =   1485
         Width           =   855
      End
      Begin VB.CommandButton Command 
         BackColor       =   &H008BAF9F&
         Caption         =   "sec"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Index           =   22
         Left            =   3000
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "����"
         Top             =   1920
         Width           =   855
      End
      Begin VB.CommandButton Command 
         BackColor       =   &H008BAF9F&
         Caption         =   "csc"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Index           =   23
         Left            =   3000
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "���"
         Top             =   2325
         Width           =   855
      End
   End
   Begin VB.TextBox Text9 
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00685758&
      Height          =   6855
      Left            =   9720
      MultiLine       =   -1  'True
      TabIndex        =   15
      ToolTipText     =   "�ֲ�������"
      Top             =   960
      Width           =   2055
   End
   Begin VB.TextBox Text8 
      BackColor       =   &H0088A897&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H003F584D&
      Height          =   405
      Left            =   480
      MousePointer    =   3  'I-Beam
      TabIndex        =   0
      ToolTipText     =   "���ʽ������  �ڴ˴�������ʽ"
      Top             =   1920
      Width           =   8655
   End
   Begin VB.Label cue 
      BackStyle       =   0  'Transparent
      Caption         =   "���ڴ˴�������ѧ���ʽ��"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00298000&
      Height          =   255
      Left            =   6840
      TabIndex        =   81
      Top             =   2400
      Width           =   2295
   End
   Begin VB.Image backspace 
      Height          =   405
      Left            =   9120
      Picture         =   "Calc.frx":1E32
      Stretch         =   -1  'True
      ToolTipText     =   "������������ı����˸�,���򷵻ص���һ�����ʽ��"
      Top             =   1920
      Width           =   240
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   120
      Left            =   10320
      Picture         =   "Calc.frx":2024
      Top             =   0
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.Menu acc 
      Caption         =   "����(&P)"
      Begin VB.Menu Wincalc 
         Caption         =   "Windows ������"
         Shortcut        =   {F2}
      End
      Begin VB.Menu hcalc 
         Caption         =   "�߾��ȼ���"
         Shortcut        =   {F3}
      End
      Begin VB.Menu step16 
         Caption         =   "-"
      End
      Begin VB.Menu Imgfun 
         Caption         =   "�������߲鿴��"
         Shortcut        =   {F4}
      End
      Begin VB.Menu func 
         Caption         =   "��һԪ����"
         Shortcut        =   {F5}
      End
      Begin VB.Menu lfun 
         Caption         =   "�����Է�����"
         Shortcut        =   {F6}
      End
      Begin VB.Menu qdjf 
         Caption         =   "�󶨻���"
         Shortcut        =   {F7}
      End
      Begin VB.Menu step17 
         Caption         =   "-"
      End
      Begin VB.Menu triangle 
         Caption         =   "��������"
      End
      Begin VB.Menu Striangle 
         Caption         =   "���������"
      End
      Begin VB.Menu znbx 
         Caption         =   "��n���μ���"
      End
      Begin VB.Menu step18 
         Caption         =   "-"
      End
      Begin VB.Menu jshls 
         Caption         =   "��������ʽ"
      End
      Begin VB.Menu qnjz 
         Caption         =   "�������"
      End
      Begin VB.Menu szzh 
         Caption         =   "����ת��"
      End
      Begin VB.Menu plzh 
         Caption         =   "�������"
      End
      Begin VB.Menu fjzys 
         Caption         =   "�ֽ�������"
      End
      Begin VB.Menu dqqmjl 
         Caption         =   "�����������"
      End
   End
   Begin VB.Menu files 
      Caption         =   "�ļ�(&F)"
      Begin VB.Menu batch 
         Caption         =   "��������"
      End
      Begin VB.Menu vw 
         Caption         =   "�򿪼�¼�ļ�"
      End
   End
   Begin VB.Menu options 
      Caption         =   "ѡ��(&O)"
      Begin VB.Menu jiaodu 
         Caption         =   "�Ƕ�"
         Checked         =   -1  'True
      End
      Begin VB.Menu hudu 
         Caption         =   "����"
      End
      Begin VB.Menu step20 
         Caption         =   "-"
      End
      Begin VB.Menu color 
         Caption         =   "��ɫ"
      End
      Begin VB.Menu step21 
         Caption         =   "-"
      End
      Begin VB.Menu wjlwj 
         Caption         =   "д��¼�ļ�"
      End
      Begin VB.Menu realtime 
         Caption         =   "��ʱ����"
      End
   End
   Begin VB.Menu History 
      Caption         =   "��ʷ��¼(&H)"
      Begin VB.Menu hlg 
         Caption         =   ""
         Index           =   0
      End
   End
   Begin VB.Menu wl 
      Caption         =   "������(&C)"
      Begin VB.Menu zkzgs 
         Caption         =   "����й���(��/��)"
      End
      Begin VB.Menu zkzcdl 
         Caption         =   "��մŵ���(��/��)"
      End
      Begin VB.Menu jxjgcs 
         Caption         =   "��ϸ�ṹ����"
         Begin VB.Menu jxjgcsa 
            Caption         =   "��"
         End
         Begin VB.Menu jxjgcsb 
            Caption         =   "1/��"
         End
      End
      Begin VB.Menu zkdrl 
         Caption         =   "��յ�����(��/��)"
      End
      Begin VB.Menu jbdh 
         Caption         =   "�������(��)"
      End
      Begin VB.Menu plkcl 
         Caption         =   "���ʿ˳���"
         Begin VB.Menu plkclh 
            Caption         =   "h"
         End
         Begin VB.Menu plkclhpi 
            Caption         =   "h/2��"
         End
      End
      Begin VB.Menu afgdlcl 
         Caption         =   "����٤���޳���(Ħ^-1)"
      End
      Begin VB.Menu yzzldw 
         Caption         =   "ԭ��������λ(ǧ��)"
      End
      Begin VB.Menu step1 
         Caption         =   "-"
      End
      Begin VB.Menu dzjzzl 
         Caption         =   "���Ӿ�ֹ����"
         Begin VB.Menu dkg 
            Caption         =   "(ǧ��)"
         End
         Begin VB.Menu dyz 
            Caption         =   "(ԭ��������λ)"
         End
      End
      Begin VB.Menu zhizdzzl 
         Caption         =   "���Ӿ�ֹ����"
         Begin VB.Menu zhikg 
            Caption         =   "(ǧ��)"
         End
         Begin VB.Menu zhiyz 
            Caption         =   "(ԭ��������λ)"
         End
      End
      Begin VB.Menu zhongzdzzl 
         Caption         =   "���Ӿ�ֹ����"
         Begin VB.Menu zhongkg 
            Caption         =   "(ǧ��)"
         End
         Begin VB.Menu zhongyz 
            Caption         =   "(ԭ��������λ)"
         End
      End
      Begin VB.Menu step2 
         Caption         =   "-"
      End
      Begin VB.Menu fldcs 
         Caption         =   "�����ڳ���(��/Ħ)"
      End
      Begin VB.Menu lbdcl 
         Caption         =   "��²�����(��^-1)"
      End
      Begin VB.Menu step8 
         Caption         =   "-"
      End
      Begin VB.Menu bebj 
         Caption         =   "�����뾶(��)"
      End
      Begin VB.Menu jddzbj 
         Caption         =   "������Ӱ뾶(��)"
      End
      Begin VB.Menu step3 
         Caption         =   "-"
      End
      Begin VB.Menu becz 
         Caption         =   "��������(������^2)"
      End
      Begin VB.Menu dzcj 
         Caption         =   "���Ӵž�(������^2)"
      End
      Begin VB.Menu zhizcj 
         Caption         =   "���Ӵž�(������^2)"
      End
      Begin VB.Menu step4 
         Caption         =   "-"
      End
      Begin VB.Menu mzjzzl 
         Caption         =   "���Ӿ�ֹ����"
         Begin VB.Menu mkg 
            Caption         =   "(ǧ��)"
         End
         Begin VB.Menu myz 
            Caption         =   "(ԭ��������λ)"
         End
      End
      Begin VB.Menu step5 
         Caption         =   "-"
      End
      Begin VB.Menu dzdkpdbc 
         Caption         =   "���ӵĿ��նٲ���(��)"
      End
      Begin VB.Menu zhizdkpdbc 
         Caption         =   "���ӵĿ��նٲ���(��)"
      End
      Begin VB.Menu zhongzdkpdbc 
         Caption         =   "���ӵĿ��նٲ���(��)"
      End
      Begin VB.Menu step6 
         Caption         =   "-"
      End
      Begin VB.Menu lxqtdmetj 
         Caption         =   "���������Ħ�����(��^3/Ħ)"
      End
      Begin VB.Menu meqtcl 
         Caption         =   "Ħ�����峣��(����Ħ^-1����^-1)"
      End
      Begin VB.Menu bezmcl 
         Caption         =   "������������(��/��)"
      End
      Begin VB.Menu ylcl 
         Caption         =   "��������(ţ����^2��ǧ��^-2)"
      End
      Begin VB.Menu lsmtcl 
         Caption         =   "��ʩ���س���(��׼״̬)��(��^-3)"
      End
      Begin VB.Menu bzdqyp 
         Caption         =   "��׼����ѹ(��)"
      End
      Begin VB.Menu step7 
         Caption         =   "-"
      End
      Begin VB.Menu sdsxdwd 
         Caption         =   "ˮ��������¶�"
         Begin VB.Menu kew 
            Caption         =   "(������)"
         End
         Begin VB.Menu ssd 
            Caption         =   "(���϶�)"
         End
      End
      Begin VB.Menu jdld 
         Caption         =   "�������(���϶�)"
      End
   End
   Begin VB.Menu tw 
      Caption         =   "���ĳ���(&A)"
      Begin VB.Menu twdw 
         Caption         =   "1���ĵ�λ(��)"
      End
      Begin VB.Menu gn 
         Caption         =   "1����"
         Begin VB.Menu gnm 
            Caption         =   "(��)"
         End
         Begin VB.Menu gntwdw 
            Caption         =   "(���ĵ�λ)"
         End
      End
      Begin VB.Menu mcj 
         Caption         =   "1����"
         Begin VB.Menu mcjm 
            Caption         =   "(��)"
         End
         Begin VB.Menu mcjtwdw 
            Caption         =   "(���ĵ�λ)"
         End
         Begin VB.Menu mcjgn 
            Caption         =   "(����)"
         End
      End
      Begin VB.Menu hcjj 
         Caption         =   "�Ƴཻ��(��)"
      End
      Begin VB.Menu step9 
         Caption         =   "-"
      End
      Begin VB.Menu yhxr 
         Caption         =   "1������(ƽ̫����)"
      End
      Begin VB.Menu ptyr 
         Caption         =   "1ƽ̫����(������)"
      End
      Begin VB.Menu step10 
         Caption         =   "-"
      End
      Begin VB.Menu swy 
         Caption         =   "1˷����(ƽ̫����)"
      End
      Begin VB.Menu hxr 
         Caption         =   "1������(ƽ̫����)"
      End
      Begin VB.Menu step11 
         Caption         =   "-"
      End
      Begin VB.Menu hgn 
         Caption         =   "1�ع���(ƽ̫����)"
      End
      Begin VB.Menu hxn 
         Caption         =   "1������(ƽ̫����)"
      End
      Begin VB.Menu rln 
         Caption         =   "1������"
         Begin VB.Menu rlnptyr 
            Caption         =   "(ƽ̫����)"
         End
         Begin VB.Menu rlnptys 
            Caption         =   "(ƽ̫��ʱ)"
         End
         Begin VB.Menu rlnptyf 
            Caption         =   "(ƽ̫����)"
         End
         Begin VB.Menu rlnptym 
            Caption         =   "(ƽ̫����)"
         End
      End
      Begin VB.Menu gln 
         Caption         =   "1������(ƽ̫����)"
      End
      Begin VB.Menu tyn 
         Caption         =   "1̫����(ƽ̫����)"
      End
   End
   Begin VB.Menu dw 
      Caption         =   "��λ����(&U)"
      Begin VB.Menu cd 
         Caption         =   "����"
         Begin VB.Menu hl 
            Caption         =   "����"
         End
         Begin VB.Menu yl 
            Caption         =   "Ӣ��"
         End
         Begin VB.Menu chi 
            Caption         =   "��"
         End
         Begin VB.Menu cun 
            Caption         =   "��"
         End
         Begin VB.Menu cdma 
            Caption         =   "��"
         End
         Begin VB.Menu yx 
            Caption         =   "ӢѰ"
         End
      End
      Begin VB.Menu tjrj 
         Caption         =   "������ݻ�"
         Begin VB.Menu mjl 
            Caption         =   "������"
         End
         Begin VB.Menu yjl 
            Caption         =   "Ӣ����"
         End
      End
      Begin VB.Menu sd 
         Caption         =   "�ٶ�"
         Begin VB.Menu jie 
            Caption         =   "��"
         End
      End
      Begin VB.Menu zl 
         Caption         =   "����"
         Begin VB.Menu bang 
            Caption         =   "��"
         End
         Begin VB.Menu aich 
            Caption         =   "��˾(����)"
         End
      End
      Begin VB.Menu wd 
         Caption         =   "�¶�"
         Begin VB.Menu hsd 
            Caption         =   "���϶�"
         End
         Begin VB.Menu lsd 
            Caption         =   "���϶�"
         End
      End
      Begin VB.Menu step12 
         Caption         =   "-"
      End
      Begin VB.Menu li 
         Caption         =   "��"
         Begin VB.Menu qkl 
            Caption         =   "ǧ����"
         End
         Begin VB.Menu bl 
            Caption         =   "����"
         End
      End
      Begin VB.Menu ylyqyl 
         Caption         =   "ѹ����ѹǿ��Ӧ��"
         Begin VB.Menu bzdqy 
            Caption         =   "��׼����ѹ"
         End
         Begin VB.Menu gcdqy 
            Caption         =   "���̴���ѹ"
         End
         Begin VB.Menu hmsz 
            Caption         =   "����ˮ��"
         End
         Begin VB.Menu hmgz 
            Caption         =   "���׹���"
         End
      End
      Begin VB.Menu ngrl 
         Caption         =   "�ܣ���������"
         Begin VB.Menu gjje 
            Caption         =   "���ʽ���"
         End
         Begin VB.Menu gjzqbk 
            Caption         =   "����������"
         End
         Begin VB.Menu rhxk 
            Caption         =   "�Ȼ�ѧ��"
         End
         Begin VB.Menu ssdk 
            Caption         =   "15���϶ȿ�"
         End
         Begin VB.Menu sdqy 
            Caption         =   "������ѹ"
         End
         Begin VB.Menu sgcdqy 
            Caption         =   "�����̴���ѹ"
         End
         Begin VB.Menu qklm 
            Caption         =   "ǧ������"
         End
         Begin VB.Menu mlxs 
            Caption         =   "����Сʱ"
         End
         Begin VB.Menu dgmlxs 
            Caption         =   "�繤����Сʱ"
         End
      End
      Begin VB.Menu gl 
         Caption         =   "����"
         Begin VB.Menu gjwt 
            Caption         =   "��������"
         End
         Begin VB.Menu km 
            Caption         =   "��/��"
         End
         Begin VB.Menu qks 
            Caption         =   "ǧ��/ʱ"
         End
         Begin VB.Menu qklmf 
            Caption         =   "ǧ������/��"
         End
         Begin VB.Menu dgml 
            Caption         =   "�繤����"
         End
         Begin VB.Menu mzml 
            Caption         =   "��������"
         End
         Begin VB.Menu yzml 
            Caption         =   "Ӣ������"
         End
      End
      Begin VB.Menu step13 
         Caption         =   "-"
      End
      Begin VB.Menu dl 
         Caption         =   "����"
         Begin VB.Menu gjap 
            Caption         =   "���ʰ���"
         End
         Begin VB.Menu ja 
            Caption         =   "���簲��"
         End
      End
      Begin VB.Menu dy 
         Caption         =   "��ѹ"
         Begin VB.Menu gjft 
            Caption         =   "���ʷ���"
         End
         Begin VB.Menu jdft 
            Caption         =   "�������"
         End
      End
      Begin VB.Menu dz 
         Caption         =   "����"
         Begin VB.Menu gjom 
            Caption         =   "����ŷķ"
         End
         Begin VB.Menu jo 
            Caption         =   "����ŷķ"
         End
      End
      Begin VB.Menu dh 
         Caption         =   "���"
         Begin VB.Menu gjkl 
            Caption         =   "���ʿ���"
         End
         Begin VB.Menu jk 
            Caption         =   "�������"
         End
      End
      Begin VB.Menu dianrong 
         Caption         =   "����"
         Begin VB.Menu gjfl 
            Caption         =   "���ʷ���"
         End
         Begin VB.Menu jdfl 
            Caption         =   "���編��"
         End
      End
      Begin VB.Menu ddao 
         Caption         =   "�絼"
         Begin VB.Menu jdxmz 
            Caption         =   "����������"
         End
      End
      Begin VB.Menu dg 
         Caption         =   "���"
         Begin VB.Menu gjhl 
            Caption         =   "���ʺ���"
         End
         Begin VB.Menu jh 
            Caption         =   "�������"
         End
      End
      Begin VB.Menu step14 
         Caption         =   "-"
      End
      Begin VB.Menu ctl 
         Caption         =   "��ͨ��"
         Begin VB.Menu gjwb 
            Caption         =   "����Τ��"
         End
      End
      Begin VB.Menu ccqd 
         Caption         =   "�ų�ǿ��"
         Begin VB.Menu ast 
            Caption         =   "��˹��"
         End
      End
      Begin VB.Menu step15 
         Caption         =   "-"
      End
      Begin VB.Menu fgqd 
         Caption         =   "����ǿ��"
         Begin VB.Menu gjzg 
            Caption         =   "�������"
         End
         Begin VB.Menu hfhzg 
            Caption         =   "���������"
         End
      End
      Begin VB.Menu gzd 
         Caption         =   "���ն�"
         Begin VB.Menu yczg 
            Caption         =   "Ӣ�����"
         End
      End
      Begin VB.Menu zsl 
         Caption         =   "������"
         Begin VB.Menu lq 
            Caption         =   "����"
         End
      End
   End
   Begin VB.Menu Helper 
      Caption         =   "����(&H)"
      NegotiatePosition=   3  'Right
      Begin VB.Menu content 
         Caption         =   "����"
         Shortcut        =   {F1}
      End
      Begin VB.Menu about 
         Caption         =   "����"
      End
   End
End
Attribute VB_Name = "Calc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private m$, ms#, mr#, mo#, DR, alfa, un, ml#, wf, much, digit, ab$, al$, hlog
Private Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Private Declare Function SetMenuItemBitmaps Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal hBitmapUnchecked As Long, ByVal hBitmapChecked As Long) As Long
Dim u#(1 To 512)



Public Sub Ht()
On Error GoTo 90



cue.Visible = False
starttime = Timer
wrong.Visible = False
digit = Slider1.Value

If much = 1 Then GoTo 200
10:
al$ = ab$
Erase u#
un = 0: ao$ = "": alfa = 0



11:

ao$ = translate(Text8.text)
ao$ = Trim(UCase$(ao$))
If InStr(ao$, "=") > 0 Then ao$ = Left(ao$, InStr(ao$, "="))
ab$ = Text8.text


If ao$ <> "Q" And ao$ <> "q" And ao$ <> "MS" And ao$ <> "ms" And ao$ <> "DMS" And ao$ <> "dms" Then Text4.text = mo#: Text6.text = LCase(al$)

12:


Select Case ao$
 Case "Q"
 Close #1: End
 Case ""
 Text1.text = ms#: GoTo 400
 

 
 Case "MS"
 mr# = mo#: Text5.text = mr#: GoTo 400
End Select

If wf = 1 Then Print #1, "": Print #1, ao$

'ao$ = translate(ao$)

alb = InStr(ao$, "("): arb = InStr(ao$, ")")
If alb >= 0 And arb > 0 And arb < alb Then
Text3.text = "��© '( '"
'Text1.ForeColor = &HC00000
ao$ = "(" + ao$: Text1.text = ao$
If realtime.Checked = False Then wrong.Visible = True
End If


l:
Lb = 0: Rb = 0
cbo$ = ao$
For cb = 1 To Len(cbo$)
  If Left(cbo$, 1) = "(" Then Lb = Lb + 1
  If Left(cbo$, 1) = ")" Then Rb = Rb + 1
  cbo$ = Right(cbo$, Len(cbo$) - 1)
Next cb
'If lr = 0 Then Text7.Text = Str(DR) + "       " + Str(digit) + "        " + Str(lb) + " " + Str(rb)
If Lb > Rb Then
  Text3.text = "��© ')'"
  'Text1.ForeColor = &HC00000
  ao$ = ao$ + ")"
  Text1.text = ao$
  If realtime.Checked = False Then wrong.Visible = True: Beep
  GoTo l
End If

If Lb < Rb Then
  Text3.text = "��© '( '"
  'Text1.ForeColor = &HC00000
  ao$ = "(" + ao$
  Text1.text = ao$
  If realtime.Checked = False Then wrong.Visible = True: Beep
  GoTo l
End If

k:
Lb = 0: Rb = 0
e$ = ""
m$ = ""

If Len(ao$) = 1 Then
    ml# = ms#
    ms# = Val(ao$)
    Text1.text = ms#
    GoTo 400
End If

20:  If InStr(ao$, "(") = 0 Then GoTo 70

30
If alfa = 0 Then GoTo o:
If Left(ao$, 3) = "(UN" And Right(ao$, 1) = ")" And Len(Str(Val(Right(ao$, Len(ao$) - 3)))) + 4 = Len(ao$) Then
   GoTo 80
End If
o:
a = Len(ao$)
bo$ = Right(ao$, a - InStr(ao$, "(") + 1)
c$ = Left(ao$, a - Len(bo$))


p:
mb = InStr(Right(bo$, Len(bo$) - 1), "(")
nb = InStr(Right(bo$, Len(bo$) - 1), ")")

If mb < nb And mb <> 0 Then
    c$ = c$ + Left(bo$, mb)
    bo$ = Right(ao$, a - Len(c$))
    GoTo p:
Else
    no$ = Left(bo$, InStr(bo$, ")"))
    d$ = Right(ao$, a - Len(c$) - Len(no$))
    no$ = mid$(no$, 2, Len(no$) - 2)

End If


If InStr(no$, "+") = 0 And InStr(no$, "-") = 0 Then m$ = no$: Call Beta: GoTo 65
If InStr(no$, "*") = 0 And InStr(no$, "/") = 0 And InStr(no$, "\") = 0 And InStr(no$, "|") = 0 And InStr(no$, "^") = 0 And InStr(no$, "@") = 0 Then GoTo 60


40:
a = Len(no$)
b = 32767
If InStr(no$, "+") > 0 Then b = InStr(no$, "+")
If InStr(no$, "-") > 0 And InStr(no$, "-") < b Then b = InStr(no$, "-")
If b = 32767 Then GoTo 60


50:
m$ = Left(no$, b - 1)
If InStr(m$, "^") > 0 Or InStr(m$, "@") > 0 Then Call Beta Else Call alpha
e$ = e$ + m$ + mid$(no$, b, 1)
no$ = Right(no$, a - b)
If InStr(no$, "*") > 0 Or InStr(no$, "/") > 0 Or InStr(no$, "\") > 0 Or InStr(no$, "|") > 0 Or InStr(no$, "^") > 0 Or InStr(no$, "@") > 0 Or InStr(no$, "+") > 0 Or InStr(no$, "-") > 0 Then GoTo 40


60:
m$ = no$
If InStr(m$, "^") > 0 Or InStr(m$, "@") > 0 Then Call Beta Else Call alpha
m$ = e$ + m$
e$ = ""
If InStr(m$, "^") > 0 Or InStr(m$, "@") > 0 Then Call Beta Else Call alpha


65:
no$ = m$
ao$ = c$ + no$ + d$
GoTo 20


70:
If Len(c$) = 0 And Len(d$) = 0 And alfa > 0 Then GoTo 80
ao$ = "(" + ao$ + ")": GoTo 30


80:
If much = 1 Then GoTo 220

If digit = -1 Then
  Text1.text = mo#
Else
  If digit = 7 Then
    If Abs(mo#) > 1.4E-45 And Abs(mo#) < 3.4E+38 Then
      pri! = mo#
      Text1.text = pri!
    End If
  Else
    Text1.text = (Round(mo# * 10 ^ digit)) / 10 ^ digit
  End If
End If

If wf = 1 Then Print #1, mo#
ml# = ms#
ms# = mo#
hlg(0).Visible = True
If realtime.Checked = False Then
hlog = hlog + 1
Load hlg(hlog)
hlg(hlog).Caption = LCase(ab$) & "=" & mo#
If hlog > 30 Then Unload hlg(hlog - 30)
End If
GoTo 400

200:
CommonDialog1.ShowOpen
If CommonDialog1.FileName = "" Then much = 0: GoTo 400
Open CommonDialog1.FileName For Input As #2
Open "Output.txt" For Output As #3
Text1.text = "Computing..."
210:
un = 0: ao$ = "": alfa = 0
If EOF(2) Then GoTo 230
Input #2, ao$
Print #3, "  "; ao$
mucherr = ExpChk(ao$)
ao$ = Trim(UCase$(ao$))
If realtime.Checked = False Then Text9.text = Text9.text & "         " & ao$ & "="
If Len(ao$) = 0 Then GoTo 230
If Left(ao$, 1) = "#" Then Print #3, "": GoTo 210
much = 1
GoTo 12
220:
If mucherr = "" Then
  Print #3, mo#
Else
  Print #3, mucherr
  serr = serr + 1
End If
If realtime.Checked = False Then Text9.text = "  " & Text9.text & mo# & Chr(13) & Chr(10)
GoTo 210
230:

Close #2, #3
much = 0
Text3.text = serr & " ������"
Text1.text = serr & " Error(s)"
If serr = 0 Then Text1.text = "0 Error": Text3.text = "û�д���"
msg = MsgBox("�Ƿ������ļ�?", vbYesNo + vbQuestion, "��ѧ������")
If msg = 6 Then RetVal = Shell("notepad.exe output.txt", 1)
GoTo 400

90: Select Case err
  Case 5
  er$ = "����������Ч": If realtime.Checked = False Then wrong.Visible = True
  Case 6
  er$ = "���": If realtime.Checked = False Then wrong.Visible = True
  Case 11
  er$ = "����Ϊ��": If realtime.Checked = False Then wrong.Visible = True
  Case 16
  er$ = "���ʽ̫����": If realtime.Checked = False Then wrong.Visible = True
  Case 51
  er$ = "�ڲ�����"
  Case 52
  er$ = "�Ҳ����ļ�"
  Case 53
  er$ = "�Ҳ����ļ�": much = 0: GoTo 95
  Case 55
  er$ = "�ļ��Ѵ�"
  Case 57
  er$ = "�豸 I/O ����"
  Case 61
  er$ = "��������"
  Case 68
  er$ = "�豸������"
  Case 70
  er$ = "����д����"
  Case 71
  er$ = "����δ׼����"
  Case 75
  er$ = "·�����ļ����ʴ���"
  Case 2446
  er$ = "����ʱ�ڴ治��"
  Case 31036
  er$ = "�ļ����ش���"
  Case 31037
  er$ = "�ļ��������"
  Case Else
  er$ = "δ֪����"
End Select
If much = 1 And er$ <> "·�����ļ����ʴ���" And er$ <> "�Ҳ����ļ�" And er$ <> "�ļ��Ѵ�" Then Print #3, er$: If realtime.Checked = False Then Text9.text = Text9.text + "Error" + Str$(err): serr = serr + 1: Resume 210
If much = 1 And (er$ = "·�����ļ����ʴ���" Or er$ = "�Ҳ����ļ�" Or er$ = "�ļ��Ѵ�") Then much = 0: GoTo 95
95:
Text6.text = er$
'If much = 0 and er$ <> "·�����ļ����ʴ���" Then wrong.Visible = True
If realtime.Checked = False Then Text9.text = Text9.text & Chr(13) & er$ & Chr(13) & Chr(10)
Text1.text = er$
Resume 400

400:
If realtime.Checked = False Then Text8.text = ""
Text8.SetFocus
finishtime = "  " & hlog & "       " & Abs(Timer * (100000000) - starttime * (100000000)) & "      " & Date & "   " & Time
Text3.text = finishtime
'Text1.ForeColor = &H2D4238
End Sub
Public Sub alpha()
alfa = alfa + 1
mo# = 0
PI# = 4 * Atn(1)
If Left(m$, 1) <> "-" Then m$ = "+" + m$

Do Until InStr(m$, "+@") = 0
   Mid$(m$, InStr(m$, "+@"), 2) = "2@"
Loop
If Left(m$, 1) <> "+" Then m$ = "+" + m$

a:
a = Len(m$)
ao$ = Right(m$, a - 1)
b = 32767
If InStr(ao$, "+") > 0 Then b = InStr(ao$, "+")
If InStr(ao$, "-") > 0 And InStr(ao$, "-") < b Then b = InStr(ao$, "-")
If InStr(ao$, "*") > 0 And InStr(ao$, "*") < b Then b = InStr(ao$, "*")
If InStr(ao$, "/") > 0 And InStr(ao$, "/") < b Then b = InStr(ao$, "/")
If InStr(ao$, "\") > 0 And InStr(ao$, "\") < b Then b = InStr(ao$, "\")
If InStr(ao$, "|") > 0 And InStr(ao$, "|") < b Then b = InStr(ao$, "|")
If InStr(ao$, "^") > 0 And InStr(ao$, "^") < b Then b = InStr(ao$, "^")
If InStr(ao$, "@") > 0 And InStr(ao$, "@") < b Then b = InStr(ao$, "@")
If b = 32767 Then c$ = Left(m$, 1): Last = 1: no$ = ao$: GoTo b
bo$ = Left(m$, b)
c$ = Left(bo$, 1)
no$ = Right(bo$, b - 1)

b:
If no$ = "G" Then n# = 9.80665: GoTo f:
If no$ = "M" Then n# = ms#: GoTo f:
If no$ = "MR" Then n# = mr#: GoTo f:
If no$ = "ML" Then n# = ml#: GoTo f:

p$ = InsFun.Funname(no$)

If p$ <> "" And p$ <> "!" Then no$ = Right(no$, Len(no$) - Len(p$))
If p$ = "LOG" Then GoTo d:

g:
If InStr(no$, "UN") = 1 Then n# = u#(Val(Right(no$, Len(no$) - 2))) Else n# = Val(no$)
If loga = 1 Then GoTo h:
If loga = 2 Then GoTo i:

d:
If p$ = "UN" Then n# = u#(Val(no$))
If Right(no$, 1) = "!" And p$ <> "!" Then
   s = 1
   For i = 1 To n#
   s = s * i
   Next i
   n# = s
End If
npr# = n#
nd# = n# * PI# / 180

If p$ = "" Then GoTo f:

Select Case p$
    Case "ABS"
    n# = Abs(n#)
    Case "SQR"
    n# = Sqr(n#)
    Case "INT"
    n# = Int(n#)
    Case "FIX", "TRUNC"
    n# = Fix(n#)
    Case "LN"
    n# = Log(n#)
    Case "LNA"
    n# = Log(Abs(n#))
    Case "SIN", "SIGN"
    If DR = 0 And (n# / 180 = Fix(n# / 180)) Then nd# = 0
    If DR = 0 Then n# = Sin(nd#) Else n# = Sin(n#)
    Case "COS"
    If DR = 0 And ((n# + 90) / 180 = Fix((n# + 90) / 180)) Then nd# = 0
    If DR = 0 Then n# = Cos(nd#) Else n# = Cos(n#)
    Case "TAN", "TG"
    If DR = 0 And (n# / 180 = Fix(n# / 180)) Then nd# = 0
    If DR = 0 And ((n# + 90) / 180 = Fix((n# + 90) / 180)) Then nd# = Log(-1)
    If DR = 0 Then n# = Tan(nd#) Else n# = Tan(n#)
    Case "ARCSIN", "ASIN"
    If DR = 0 Then n# = (Atn(n# / Sqr(1 - n# ^ 2))) * 180 / PI# Else n# = Atn(n# / Sqr(1 - n# ^ 2))
    Case "ARCCOS", "ACOS"
    If DR = 0 Then n# = (PI# / 2 - Atn(n# / Sqr(1 - n# ^ 2))) * 180 / PI# Else n# = PI# / 2 - Atn(n# / Sqr(1 - n# ^ 2))
    Case "ARCTG", "ATN", "ARCTAN", "ATAN"
    If DR = 0 Then n# = (Atn(n#)) * 180 / PI# Else n# = Atn(n#)
    Case "ARCCTG", "ACOT", "ARCCOT"
    If DR = 0 Then n# = (PI# / 2 - Atn(n#)) * 180 / PI# Else n# = PI# / 2 - Atn(n#)
    Case "ARCSEC", "ASEC"
    n# = 1 / n#
    If DR = 0 Then n# = (PI# / 2 - Atn(n# / Sqr(1 - n# ^ 2))) * 180 / PI# Else n# = PI# / 2 - Atn(n# / Sqr(1 - n# ^ 2))
    Case "ARCCSC", "ACSC"
    n# = 1 / n#
    If DR = 0 Then n# = (Atn(n# / Sqr(1 - n# ^ 2))) * 180 / PI# Else n# = Atn(n# / Sqr(1 - n# ^ 2))
    Case "EXP", "EP"
    n# = Exp(n#)
    Case "SGN"
    n# = Sgn(n#)
    Case "COT"
    If DR = 0 And (n# / 180 = Fix(n# / 180)) Then n = Log(-1)
    If DR = 0 And ((n# + 90) / 180 = Fix((n# + 90) / 180)) Then n# = 0: GoTo f:
    If DR = 0 Then n# = 1 / (Tan(nd#)) Else n# = 1 / (Tan(n#))
    Case "SEC"
    If DR = 0 And ((n# + 90) / 180 = Fix((n# + 90) / 180)) Then nd# = Log(-1)
    If DR = 0 Then n# = 1 / (Cos(nd#)) Else n# = 1 / (Cos(n#))
    Case "CSC"
    If DR = 0 And (n# / 180 = Fix(n# / 180)) Then nd# = Log(-1)
    If DR = 0 Then n# = 1 / (Sin(nd#)) Else n# = 1 / (Sin(n#))
    Case "LOG"
    If InStr(no$, "`") > 0 Then
      nao$ = Left(no$, InStr(no$, "`"))
      nno$ = Right(no$, Len(no$) - InStr(no$, "`"))
      no$ = nao$: loga = 1: GoTo g:
    Else: p$ = "LN": GoTo g:
    End If
h:      na# = n#
    no$ = nno$: loga = 2: GoTo g:
i:      nn# = n#
    n# = Log(nn#) / Log(na#)
    loga = 0: GoTo f:
    Case "LG"
    n# = Log(n#) / Log(10)
    Case "SH", "SINH"
    'If DR = 0 Then n# = Sin(nd#) Else n# = Sin(n#)
    If DR = 0 Then n# = (Exp(nd#) - Exp(-nd#)) / 2 Else n# = (Exp(n#) - Exp(-n#)) / 2
    Case "CH", "COSH"
    If DR = 0 Then n# = (Exp(nd#) + Exp(-nd#)) / 2 Else n# = (Exp(n#) + Exp(-n#)) / 2
    Case "TH", "TANH"
    If DR = 0 Then n# = (Exp(nd#) - Exp(-nd#)) / (Exp(nd#) + Exp(-nd#)) Else n# = (Exp(n#) - Exp(-n#)) / (Exp(n#) + Exp(-n#))
    Case "CTH", "COTH"
    If DR = 0 Then n# = (Exp(nd#) + Exp(-nd#)) / (Exp(nd#) - Exp(-nd#)) Else n# = (Exp(n#) + Exp(-n#)) / (Exp(n#) - Exp(-n#))
    Case "SECH"
    If DR = 0 Then n# = 2 / (Exp(nd#) + Exp(-nd#)) Else n# = 2 / (Exp(n#) + Exp(-n#))
    Case "CSCH"
    If DR = 0 Then n# = 2 / (Exp(nd#) - Exp(-nd#)) Else n# = 2 / (Exp(n#) - Exp(-n#))
    Case "ARSH", "ASINH"
    'If DR = 0 Then n# = (Atn(n# / Sqr(1 - n# ^ 2))) * 180 / PI# Else n# = Atn(n# / Sqr(1 - n# ^ 2))
    If DR = 0 Then n# = (Log(n# + Sqr(n# ^ 2 + 1))) * 180 / PI# Else n# = Log(n# + Sqr(n# ^ 2 + 1))
    Case "ARCH", "ACOSH"
    If DR = 0 Then n# = (Log(n# + Sqr(n# ^ 2 - 1))) * 180 / PI# Else n# = Log(n# + Sqr(n# ^ 2 - 1)): Beep '+
    Case "ARTH", "ATANH"
    If DR = 0 Then n# = ((Log((n# + 1) / (1 - n#))) / 2) * 180 / PI# Else n# = (Log((n# + 1) / (1 - n#))) / 2
    Case "ARCTH", "ACOTH"
    If DR = 0 Then n# = ((Log((n# + 1) / (n# - 1))) / 2) * 180 / PI# Else n# = (Log((n# + 1) / (n# - 1))) / 2
    Case "ARSECH", "ASECH"
    If DR = 0 Then n# = Arsech(n#) * 180 / PI# Else n# = Log((1 + Sqr(1 - n# ^ 2)) / (1 - Sqr(1 - n# ^ 2))) / 2: Beep '+
    Case "ARCSCH", "ACSCH" '?
    If DR = 0 Then n# = (Log((Sgn(n#) * Sqr(n# ^ 2 + 1) + 1) / n#)) * 180 / PI# Else n# = Log((Sgn(n#) * Sqr(n# ^ 2 + 1) + 1) / n#)
    Case "DMS"
    n# = Dms(n#)
    Case "DEG"
    n# = Deg(n#)
    Case "ROUND"
    n# = Round(n#)
    Case "!"
        s = 1
        For i = 1 To n#
        s = s * i
        Next i
        n# = s
End Select

f:
If p$ <> "UN" And much = 0 Then
  If realtime.Checked = False And Trim(p$) <> "" Then
    If p$ <> "!" Then Text9.text = Text9.text & Chr(10) & LCase(p$) & npr# & "= " & n# & "   "
    
    If p$ = "!" Then Text9.text = Text9.text & Chr(10) & npr# & LCase(p$) & "= " & n# & "   "
  End If
End If
nd# = 0     '?
p$ = ""
mopr# = mo#
If c$ = "+" Then mo# = mo# + n#
If c$ = "-" Then mo# = mo# - n#
If c$ = "*" Then mo# = mo# * n#
If c$ = "/" Then mo# = mo# / n#
If c$ = "\" Then mo# = mo# \ n#
If c$ = "|" Then mo# = mo# Mod n#
If c$ = "^" And mo# = 0 And n# = 0 Then Text3.text = "Division by zero": Beep: wrong.Visible = True
If realtime.Checked = False And c$ = "^" And mo# < 0 And ((Abs(n#) < 1 And Abs(n#) > 0) Or Fix(n#) <> n#) Then
msg$ = "(" & mo# & ") ^ (" & n# & ") ��ֵ��������?"
  Sty = vbYesNo + vbQuestion + vbDefaultButton1
  gon = MsgBox(msg$, Sty, "�޷�ȷ������")
If gon = 7 Then mo# = -(-mo#) ^ n#
If gon = 6 Then mo# = (-mo#) ^ n#
GoTo j:
End If

If c$ = "@" And n# < 0 And Parity(mo#) = 0 And realtime.Checked = False Then
If (Abs(mo#) > 1 Or (1 / mo#) <> Fix(1 / mo#)) Then
msg$ = "(" & n# & ") �� (" & mo# & ") ��ֵ��������?"
  Sty = vbYesNo + vbQuestion + vbDefaultButton1
  gon = MsgBox(msg$, Sty, "�޷�ȷ������")
If gon = 7 Then mo# = -(-n#) ^ (1 / mo#)
If gon = 6 Then mo# = (-n#) ^ (1 / mo#)
GoTo j:
End If
End If

If c$ = "@" And n# < 0 And Parity(mo#) <> 0 Then
     If Parity(mo#) = 1 Then mo# = -(-n#) ^ (1 / mo#) Else mo# = Log(-1)
    
     GoTo j:
End If


If c$ = "^" Then mo# = mo# ^ n#
If c$ = "@" Then mo# = n# ^ (1 / mo#)
If Len(c$) > 0 And mopr# <> 0 And Len(c$) > 0 And much = 0 Then
If realtime.Checked = False And Trim(c$) <> "" Then Text9.text = Text9.text & Chr(10) & mopr# & LCase(c$) & n# & "= " & mo# & "    "
End If
j:
If Last = 1 Then GoTo e:
m$ = Right(m$, a - b)
GoTo a

e:
Last = 0
un = un + 1
u#(un) = mo#
m$ = "UN" + Str$(un)
End Sub

Public Sub Beta()
f$ = ""
f$ = m$

15:
a = Len(f$)
b = 32767
If InStr(f$, "*") > 0 Then b = InStr(f$, "*")
If InStr(f$, "/") > 0 And InStr(f$, "/") < b Then b = InStr(f$, "/")
If InStr(f$, "\") > 0 And InStr(f$, "\") < b Then b = InStr(f$, "\")
If InStr(f$, "|") > 0 And InStr(f$, "|") < b Then b = InStr(f$, "|")
If b = 32767 Then GoTo 35

25:
m$ = Left(f$, b - 1)
If Len(m$) <> 0 Then Call alpha
e$ = e$ + m$ + mid$(f$, b, 1)
f$ = Right(f$, a - b)
If InStr(f$, "^") > 0 Or InStr(f$, "@") > 0 Then GoTo 15 Else m$ = f$: GoTo 37

35:
m$ = f$
If Len(m$) <> 0 Then Call alpha

37:
m$ = e$ + m$
e$ = ""
Call alpha
End Sub

Private Sub about_Click()
frmAbout.Show
End Sub

Private Sub Advance_Click()
If Calc.WindowState = 2 Then Calc.WindowState = 0 Else Calc.WindowState = 2
Advance.Visible = False
'Advance.Left = 780
End Sub

Private Sub advertisement_Click()
msg = MsgBox("�����Ҫ�Ķ������?", vbYesNo + vbInformation, "���")
If msg = 6 Then Advertise.Show
End Sub

Private Sub afgdlcl_Click()
Text8.text = Text8.text + "(6.0221367E+23)"
End Sub
Private Sub aich_Click()
Text8.text = Text8.text + "28.3459g"
End Sub
Private Sub ast_Click()
Text8.text = Text8.text + "(1000/(4*pi))"
Text3.text = "��λ ������"
End Sub
Private Sub bang_Click()
Text8.text = Text8.text + "453.592g"
End Sub
Private Sub bebj_Click()
Text8.text = Text8.text + "(5.29177249E-11)"
End Sub
Private Sub becz_Click()
Text8.text = Text8.text + "(9.2740154E-24)"
End Sub
Private Sub bezmcl_Click()
Text8.text = Text8.text + "(1.380658E-23)"
End Sub
Private Sub bl_Click()
Text8.text = Text8.text + "4.44822N"
End Sub
Private Sub bzdqy_Click()
Text8.text = Text8.text + "101325Pa"
End Sub

Private Sub bzdqyp_Click()
Text8.text = Text8.text + "(101325)"
End Sub

Private Sub color_Click()
CommonDialog1.ShowColor
Calc.BackColor = CommonDialog1.color
'Check1.BackColor = CommonDialog1.color
Option1.BackColor = CommonDialog1.color
Option2.BackColor = CommonDialog1.color
'Label2.BackColor = CommonDialog1.color
Frame1.BackColor = CommonDialog1.color
Frame2.BackColor = CommonDialog1.color
End Sub

Private Sub content_Click()
On Error GoTo rd
 readme = Shell("explorer " & App.Path & "\README.html", 1)
rd:  If err = 53 Then
        msg = MsgBox("�����ĵ�û���ҵ�", vbOKOnly, "����")
        Resume Next
     End If
End Sub

Private Sub cue_Click()
cue.Visible = False
End Sub

Private Sub dqqmjl_Click()
Globe.Show
End Sub

Private Sub fjzys_Click()
Prime.Show
End Sub

Private Sub func_Click()
Fct.Show
End Sub



Private Sub hcalc_Click()
hp.Show
End Sub

Private Sub hudu_Click()
If hudu.Checked = False Then
  hudu.Checked = True
  jiaodu.Checked = False
  Option2.Value = True
  Option1.Value = False
End If
End Sub

Private Sub ImgFun_Click()
Pic.Show
End Sub



Private Sub jiaodu_Click()
If jiaodu.Checked = False Then
  jiaodu.Checked = True
  hudu.Checked = False
  Option1.Value = True
  Option2.Value = False
End If
End Sub

Private Sub jshls_Click()
DeterForm.Show
End Sub

Private Sub jxjgcsb_Click()
Text8.text = Text8.text + "(137.0359895)"
End Sub

Private Sub lfun_Click()
fun.Show 0
End Sub

Private Sub memory_Click(Index As Integer)
Text8.text = Text8.text + "(" + memory(Index).Caption + ")"
Text8.SelStart = Len(Text8.text)
Text8.SetFocus
End Sub

Private Sub Msave_Click()
mr# = mo#: Text5.text = mr#
Text8.SetFocus
End Sub



Private Sub plzh_Click()
ca.Show
End Sub

Private Sub qdjf_Click()
dfintegral.Show
End Sub

Private Sub qnjz_Click()
GJM.Show
End Sub

Private Sub Race_Click()
Speed.Show
End Sub

Private Sub realtime_Click()
If realtime.Checked = False Then
  realtime.Checked = True
  'realtm.Enabled = True
  Text2.text = ""
 ' wrong.Enabled = False
Else
  realtime.Checked = False
 ' realtm.Enabled = False
 ' wrong.Enabled = True
End If
Text8.SetFocus
End Sub



Public Sub result_Click()
Text8.ToolTipText = Text8.text
Text1.ForeColor = &H2D4238

errmsg = ExpChk(Text8.text)
If errmsg <> "" Then
  Text1.ForeColor = &HC00000
  msg = MsgBox(errmsg & Chr(13) & Chr(13) & "��Ҫ��ֹ������?", vbYesNo + vbQuestion + vbDefaultButton1, "���������󱨸�����")
  If msg = 6 Then
    Text1.ForeColor = &H2D4238
    Text8.SetFocus
    Exit Sub
  End If
End If

tmp$ = LCase(Text8.text)
Do Until InStr(tmp$, "exp") = 0
    tmp$ = Left(tmp$, InStr(tmp$, "exp") - 1) + "eop" + Right(tmp$, Len(tmp$) - InStr(tmp$, "exp") - 2)
Loop

Do Until InStr(tmp$, "fix") = 0
    tmp$ = Left(tmp$, InStr(tmp$, "fix") - 1) + "fio" + Right(tmp$, Len(tmp$) - InStr(tmp$, "fix") - 2)
Loop

If InStr(tmp$, "x") > 0 Or InStr(tmp$, "y") > 0 Then


If InStr(Text8.text, "=") = 0 Then Pic.Show: Pic.Text1.text = Text8.text

If InStr(Text8.text, "=") > 0 And InStr(Text8.text, "y") = 0 Then
  Fct.Show
  Fct.Text1.text = Left(Text8.text, InStr(Text8.text, "=") - 1) + "-(" + Right(Text8.text, Len(Text8.text) - InStr(Text8.text, "=")) + ")"
  If Right(Fct.Text1.text, 3) = "-()" Then Fct.Text1.text = Left(Fct.Text1.text, Len(Fct.Text1.text) - 3)
End If

If InStr(Text8.text, "=") > 0 And InStr(Text8.text, "y") > 0 Then
  Pic.Show
  Pic.Text1.text = Left(Text8.text, InStr(Text8.text, "=") - 1) + "-(" + Right(Text8.text, Len(Text8.text) - InStr(Text8.text, "=")) + ")"
  If Right(Pic.Text1.text, 3) = "-()" Then Pic.Text1.text = Left(Pic.Text1.text, Len(Pic.Text1.text) - 3)
End If

If InStr(Text8, "y") > 0 Then
  If Pic.ImplicitFun.Checked = False Then
    Pic.ImplicitFun.Checked = True
    Pic.ExplicitFun.Checked = False
  End If
End If

Sendkeys "{Enter}"
Exit Sub


End If
  
  
  If realtime.Checked = False Then Text2.text = LCase(Text8.text)
  Text9.text = ""
  fsh.text = ""
  Call Ht
DoEvents
On Error Resume Next


'___________________________
'mo# = mo# / 10 ^ 17 * (10 ^ 17)
Text1Text = Str(mo#)
If InStr(Text1Text, "E-") > 0 And (InStr(Text1Text, "99999999") > 0 Or InStr(Text1Text, "00000000") > 0) Then
 mos! = mo#
 Text1Text = Str(mos!)
End If

If Slider1.Value = -1 Then Text1.text = Text1Text

If Left(Text1.text, 1) = "." Then Text1.text = "0" + Text1.text
If Left(Text1.text, 2) = "-." Then
  Text1.text = "-0." + Right(Text1.text, Len(Text1.text) - 2)
End If

If InStr(Text1Text, ".") > 0 Then
  xxbf$ = Right(Text1Text, Len(Text1Text) - InStr(Text1Text, "."))
  If InStr(xxbf$, "E") > 0 Then
    xxbf$ = Left(xxbf$, InStr(xxbf$, "E") - 1)
  End If
  gtyc = 0
  xxbfbak$ = xxbf$
  l = Len(xxbf$)
  If l <= 4 Then
  fm = Val("1" + String(l, "0"))
  fz = Val(xxbf$)
  'gys = Gcd(fm, fz)
  'If fsh.Text = "" Then
  
    gys = Gcd(fm, fz)
    If gys <> 0 Then
     fz = fz / gys
     fm = fm / gys
     fsh.text = Str(fz) + " /" + Str(fm)
    End If
  'End If
 Exit Sub
 End If
  
  
  l = Len(xxbf$)
  fm = Val("1" + String(l, "0"))
  fz = Val(xxbf$)
yc:
  
  For hf = 1 To 500
    For fh = hf + 1 To 500
      DoEvents
      yxbf = hf / fh
      yxb$ = Right(Str(yxbf), Len(Str(yxbf)) - 2)
      
      If InStr(yxb$, xxbf$) = 1 Then
        fsh.text = Str(hf) + " /" + Str(fh)
        GoTo yb:
      End If
    Next fh
  Next hf
If fsh.text = "" Then
  xxbf$ = Left(xxbf$, Len(xxbf$) - 1)
  If gtyc = 0 Then gtyc = 1: GoTo yc
End If

yb:
If fsh.text = "" Then
xxbf$ = xxbfbak$
l = Len(xxbf$)
  
  fm = Val("1" + String(l, "0"))
  fz = Val(xxbf$)
  gys = Gcd(fm, fz)
  
  
    
    
     fz = fz / gys
     fm = fm / gys
     fsh.text = Str(fz) + " /" + Str(fm)
    
End If

End If
 

  



End Sub
Private Sub Cmod_Click()
'Text8.Text = Text8.Text + "|"
Text8.SetFocus

Sendkeys ("|")
Text3.text = "�ڽ��� Mod (|) ����ʱ����������� number1 �� number2 ������������������������������������������ء�"
End Sub

Private Sub Shoot_Click()
Target.Show
End Sub

Private Sub Striangle_Click()
Stri.Show
End Sub

Private Sub szzh_Click()
Carry.Show
End Sub
Private Sub text8_KeyDown(keycode As Integer, Shift As Integer)
 ShiftDown = (Shift And vbShiftMask) > 0
 altdown = (Shift And vbAltMask) > 0
 CtrlDown = (Shift And vbCtrlMask) > 0

 Select Case keycode
   Case 83
   If ShiftDown Then Sendkeys "{BACKSPACE}": Text8.text = Text8.text + "sin": T1sf
   Case 79
   If ShiftDown Then Sendkeys "{BACKSPACE}": Text8.text = Text8.text + "cot": T1sf
   Case 88
   If ShiftDown Then Sendkeys "{BACKSPACE}": Text8.text = Text8.text + "exp": T1sf
   Case 84
   If ShiftDown Then Sendkeys "{BACKSPACE}": Text8.text = Text8.text + "tan": T1sf
   Case 76
   If ShiftDown Then Sendkeys "{BACKSPACE}": Text8.text = Text8.text + "log": T1sf
   Case 67
   If ShiftDown Then Sendkeys "{BACKSPACE}": Text8.text = Text8.text + "cos": T1sf
   Case 65
   If ShiftDown Then Sendkeys "{BACKSPACE}": Text8.text = Text8.text + "arc": T1sf
   Case 69
   If ShiftDown Then Sendkeys "{BACKSPACE}": Text8.text = Text8.text + "[e]": T1sf
   
   Case vbKeyEscape
   Text8.text = "": Sendkeys "{BACKSPACE}"

End Select
End Sub

Private Sub T1sf()
Text8.SelStart = Len(Text8.text): Text8.SetFocus
End Sub
Private Sub Text8_click()
cue.Visible = False
End Sub
Private Sub text8_change()
If realtime.Checked = True Then
  Text1.ForeColor = &H2D4238
  Call Ht
 errmsg = ExpChk(Text8.text)
 If errmsg <> "" Then Text1.ForeColor = &HC00000

 'Text1.ForeColor = &H2D4238
End If
End Sub
Private Sub vw_Click()
'shell notepad c:\temp\calc.Text1
On Error GoTo rv
'RetVal = Shell("C:\Documents and Settings\Windows\My Documents\calc.txt", 1)
RetVal = Shell("notepad.exe calc.txt", 1)
Exit Sub
rv: msg = MsgBox("�ļ�û���ҵ���", vbInformation, "��ѧ������")
Resume Next
End Sub

Private Sub Wincalc_Click()
On Error GoTo rd
  'readme = Shell("c:\Program Files\Windows NT\Accessories\wordpad.exe readme.doc", 1)

wc = Shell("calc.exe", 1)
rd:  If err = 53 Then
        msg = MsgBox("�ļ�û���ҵ�", vbOKOnly, "������֧������")
        Resume Next
     End If
End Sub

Private Sub wjlwj_Click()
On Error Resume Next
If wjlwj.Checked = False Then
wjlwj.Checked = True
wf = 1
Open "Calc.txt" For Output Shared As #1
Else
wjlwj.Checked = False
wf = 0
Close #1
End If
Text8.SetFocus
End Sub

Private Sub wrong_Click()
Text8.text = Text2.text
Text8.SelStart = Len(Text8.text)
Text8.SetFocus
wrong.Visible = False
End Sub

Private Sub Zero_Click()
Text8.SetFocus
Sendkeys "0"
'Text8.Text = Text8.Text + "0"
End Sub
Private Sub batch_Click()
Text9.text = ""
much = 1
Call Ht
End Sub
Private Sub backspace_Click()
'If Len(Text8.Text) > 0 Then Text8.Text = Left(Text8.Text, Len(Text8.Text) - 1) Else Text8.Text = Text2.Text
Text8.SetFocus
If Len(Text8.text) > 0 Then Sendkeys ("{BACKSPACE}") Else Text8.text = Text2.text: Text8.SelStart = Len(Text8.text)

End Sub
Private Sub sqrt_Click()
Text8.text = Text8.text + "@"
Text8.SelStart = Len(Text8.text)
Text8.SetFocus
End Sub
Private Sub Command_Click(Index As Integer)
'Text8.Text = Text8.Text + Command(Index).Caption
Text8.SetFocus
Select Case Command(Index).Caption
Case "+"
  Sendkeys "{+}"
Case "%"
  Sendkeys "{%}"
Case "^"
  Sendkeys "{^}"
Case "("
  Sendkeys "{(}"
Case ")"
  Sendkeys "{)}"
'Case "["
  'SendKeys "{[}"
'Case "]"
  'SendKeys "{]}"
Case Else
  Sendkeys Command(Index).Caption
End Select
'Text8.SelStart = Len(Text8.Text)
End Sub
Private Sub Command_mousemove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
If Index = 31 Or Index = 30 Then Command(31).Visible = True Else Command(31).Visible = False
'Command(Index).BackColor = &HB5CCC2
End Sub

Private Sub circle_Click()
Text8.text = Text8.text + "pi"
Text8.SelStart = Len(Text8.text)
Text8.SetFocus
End Sub

Private Sub Leftparenthesis_Click()
Text8.text = Text8.text + "("
Text8.SelStart = Len(Text8.text)
Text8.SetFocus
End Sub
Private Sub Rightparenthesis_Click()
Text8.text = Text8.text + ")"
Text8.SelStart = Len(Text8.text)
Text8.SetFocus
End Sub
Private Sub factorial_Click()
Text8.SetFocus
Sendkeys ("!")
'Text8.Text = Text8.Text + "!"

End Sub
Private Sub text9_dblClick()
Text9.Font.Size = Text9.Font.Size + 1
End Sub
Private Sub dgml_Click()
Text8.text = Text8.text + "764W"
End Sub
Private Sub dgmlxs_Click()
Text8.text = Text8.text + "2.68560MJ"
End Sub
Private Sub dkg_Click()
Text8.text = Text8.text + "(9.1093897E-31)"
End Sub
Private Sub dyz_Click()
Text8.text = Text8.text + "(5.48579903E-4)"
End Sub
Private Sub dzcj_Click()
Text8.text = Text8.text + "(9.2847701E-24)"
End Sub
Private Sub dzdkpdbc_Click()
Text8.text = Text8.text + "(2.42631058E-12)"
End Sub
Private Sub fldcs_Click()
Text8.text = Text8.text + "96485.309"
End Sub
Private Sub gcdqy_Click()
Text8.text = Text8.text + "98066.5Pa"
End Sub
Private Sub gjap_Click()
Text8.text = Text8.text + "0.99985A"
End Sub
Private Sub gjfl_Click()
Text8.text = Text8.text + "0.99951F"
End Sub
Private Sub gjft_Click()
Text8.text = Text8.text + "1.00034V"
End Sub
Private Sub gjhl_Click()
Text8.text = Text8.text + "1.00049H"
End Sub
Private Sub gjje_Click()
Text8.text = Text8.text + "1.00019J"
End Sub
Private Sub gjkl_Click()
Text8.text = Text8.text + "0.99985C"
End Sub
Private Sub gjom_Click()
Text8.text = Text8.text + "1.00049��"
End Sub
Private Sub gjwb_Click()
Text8.text = Text8.text + "1.00034Wb"
End Sub
Private Sub gjwt_Click()
Text8.text = Text8.text + "1.00019W"
End Sub
Private Sub gjzg_Click()
Text8.text = Text8.text + "1.019cd"
End Sub
Private Sub gjzqbk_Click()
Text8.text = Text8.text + "4.1868J"
End Sub
Private Sub gln_Click()
Text8.text = Text8.text + "365.2425"
End Sub
Private Sub gnm_Click()
Text8.text = Text8.text + "(9.4605E+15)"
End Sub
Private Sub gntwdw_Click()
Text8.text = Text8.text + "63240"
End Sub
Private Sub hcjj_Click()
Text8.text = Text8.text + "23.4392911111"
Text3.text = "2000��1��1.5�յĻƳཻ��"
End Sub
Private Sub hfhzg_Click()
Text8.text = Text8.text + "0.903cd"
End Sub
Private Sub hgn_Click()
Text8.text = Text8.text + "365.24220"
End Sub
Private Sub hl_Click()
Text8.text = Text8.text + "1852m"
End Sub
Private Sub hlg_Click(Index As Integer)
Text8.text = hlg(Index).Caption
End Sub
Private Sub hmgz_Click()
Text8.text = Text8.text + "133.322Pa"
End Sub
Private Sub hmsz_Click()
Text8.text = Text8.text + "9.80665Pa"
End Sub
Private Sub hsd_Click()
Text8.text = Text8.text + "0.555555555555556K"
End Sub
Private Sub hxn_Click()
Text8.text = Text8.text + "365.25636"
End Sub
Private Sub hxr_Click()
Text8.text = Text8.text + "27.321662"
End Sub
Private Sub ja_Click()
Text8.text = Text8.text + "(3.33564E-10)"
End Sub
Private Sub jbdh_Click()
Text8.text = Text8.text + "(1.60217733E-19)"
End Sub
Private Sub jddzbj_Click()
Text8.text = Text8.text + "(2.81794092E-15)"
End Sub
Private Sub jdfl_Click()
Text8.text = Text8.text + "(1.11265E-12)" 'F
End Sub
Private Sub jdft_Click()
Text8.text = Text8.text + "299.7925V"
End Sub
Private Sub jdld_Click()
Text8.text = Text8.text + "(-273.15)"
End Sub
Private Sub jdxmz_Click()
Text8.text = Text8.text + "(1.11265E-12)��"
End Sub
Private Sub jf_Click()

End Sub
Private Sub jh_Click()
Text8.text = Text8.text + "(8.98755E+11)H"
End Sub
Private Sub jie_Click()
Text8.text = Text8.text + "(1852/3600)"
Text3.text = "��λ���ף���"
End Sub
Private Sub jk_Click()
Text8.text = Text8.text + "(3.33564E-10)"
End Sub
Private Sub jo_Click()
Text8.text = Text8.text + "(8.98755E+11)��"
End Sub
Private Sub jxjgcsa_Click()
Text8.text = Text8.text + "(7.29735308E-3)"
End Sub
Private Sub kew_Click()
Text8.text = Text8.text + "273.16"
End Sub
Private Sub km_Click()
Text8.text = Text8.text + "4.1868W"
End Sub
Private Sub lbdcl_Click()
Text8.text = Text8.text + "10973731.534"
End Sub
Private Sub lq_Click()
Text8.text = Text8.text + "0.000258"
Text3.text = "��λ �⣯ǧ��"
End Sub
Private Sub lsd_Click()
Text8.text = Text8.text + "1.25��"
End Sub
Private Sub lsmtcl_Click()
Text8.text = Text8.text + "(2.686763E+25)"
End Sub
Private Sub lxqtdmetj_Click()
Text8.text = Text8.text + "(2.241410E-2)"
Text3.text = "���������ڱ�׼�¶ȡ���ѹ�µ�Ħ�����(��^3/Ħ)"
End Sub
Private Sub mcjgn_Click()
Text8.text = Text8.text + "3.262"
End Sub
Private Sub mcjm_Click()
Text8.text = Text8.text + "(3.0857E+16)"
End Sub
Private Sub mcjtwdw_Click()
Text8.text = Text8.text + "206265"
End Sub
Private Sub meqtcl_Click()
Text8.text = Text8.text + "8.314510"
End Sub
Private Sub mjl_Click()
Text8.text = Text8.text + "3.785L"
End Sub
Private Sub mkg_Click()
Text8.text = Text8.text + "(1.8835327E-28)"
End Sub
Private Sub mlxs_Click()
Text8.text = Text8.text + "2.64779MJ"
End Sub
Private Sub myz_Click()
Text8.text = Text8.text + "0.113428913"
End Sub
Private Sub mzml_Click()
Text8.text = Text8.text + "735.499W"
End Sub
Private Sub Option1_Click()
Let DR = 0
ms# = ms# * 180 / (4 * Atn(1))
Text1.text = ms#
jiaodu.Checked = True
hudu.Checked = False
End Sub
Private Sub Option2_Click()
Let DR = 1
ms# = ms# * 4 * Atn(1) / 180
Text1.text = ms#
hudu.Checked = True
jiaodu.Checked = False
End Sub


Private Sub Form_unLoad(Cancel As Integer)
'main.WindowState = 0
If wf = 1 Then Close #1: wf = 0
End Sub

Private Sub Form_Load()
'Text2.Text = " ��ѧ���ţ�  + �ӡ�����  - ��������  * ��  / �����  \ ����  | ������  ^ �˷� @ ���� " _
'& " ������ abs x ����ֵ exp x ָ����������eΪ��)  fix x ��������  int x ������x���������  log a`x ��aΪ�׵Ķ���   ln x  log x ��eΪ�׵Ķ�������Ȼ������   lg x��10Ϊ�׵Ķ��������ö�����  sgn x ���ź���  sqr x ��ƽ�� " _
'& "sin x ����  cos x ���� tan x ����  cot x ����  sec x ����  csc x ���  " _
'& "arcsin x ������  arccos x ������  arctan x ������  arccot x ������  arcsec x ������  arccsc x �����  " _
'& "sh x ˫������  ch x ˫������  th x ˫������   cth x ˫������  sech x ˫������  csch x ˫�����  " _
'& "arsh x ��˫������  arch x ��˫������  arth x ��˫������  arcth x ��˫������  arsech x ��˫������  arcsch x ��˫�����  " _
'& "n! n�Ľ׳�  dms x ��xת��Ϊ�öȱ�ʾ�ĸ�ʽ   ����:pi Բ����  m �˴μ�����  ml�ϴμ����� mr �洢����   " _
'& "����: dms ����ʾ����ת��Ϊ���ȣ��֣��롱��ʽ   ms �洢������  " _
'& "ע��:1 �˺Ų���ʡ�� 2.���ڵ���ѧ���ű���������'( )'����,'[ ]'��'( )'��Ч 3. ����˳��:���š��׳�(!)��������ָ������(^��@)���˷�������������������\������ģ���������ӷ�(+)�ͼ���(-)"
'Const myflag = &H400&
'Dim mHandle As Long, lRet As Long, sHandle As Long
'mHandle = GetMenu(hwnd)
'sHandle = GetSubMenu(mHandle, 0)
'lRet = SetMenuItemBitmaps(sHandle, 1, myflag, Image1.Picture, Image1.Picture)
'lRet = SetMenuItemBitmaps(sHandle, 2, myflag, Image1.Picture, Image1.Picture)
'lRet = SetMenuItemBitmaps(sHandle, 3, myflag, Image1.Picture, Image1.Picture)
'lRet = SetMenuItemBitmaps(sHandle, 4, myflag, Image1.Picture, Image1.Picture)
'lRet = SetMenuItemBitmaps(sHandle, 0, myflag, Image1.Picture, Image1.Picture)

'Text2.Text = " ��ӭʹ��"
Text3.text = "      Ubiquitous Computing    �������ڵļ���"

End Sub
Private Sub pi_Click()
Text8.text = Text8.text + "3.14159265358979323846264338327950288419716939931148196659300057"
End Sub
Private Sub Picture1_Click()
  Picture1.Visible = False
End Sub
Private Sub ����й���_Click()
End Sub
Private Sub plkclh_Click()
Text8.text = Text8.text + "(6.6260755E-34)"
End Sub
Private Sub plkclhpi_Click()
Text8.text = Text8.text + "(1.05457266E-34)"
End Sub
Private Sub ptyr_Click()
Text8.text = Text8.text + "1.00273791"
End Sub
Private Sub qkl_Click()
Text8.text = Text8.text + "9.80665N"
End Sub
Private Sub qklm_Click()
Text8.text = Text8.text + "9.80665J"
End Sub
Private Sub qklmf_Click()
Text8.text = Text8.text + "0.163444W"
End Sub
Private Sub qks_Click()
Text8.text = Text8.text + "1.163W"
End Sub
Private Sub rhxk_Click()
Text8.text = Text8.text + "4.184J"
End Sub
Private Sub rlnptyf_Click()
Text8.text = Text8.text + "525960"
End Sub
Private Sub rlnptym_Click()
Text8.text = Text8.text + "31557600"
End Sub
Private Sub rlnptyr_Click()
Text8.text = Text8.text + "365.25"
End Sub
Private Sub rlnptys_Click()
Text8.text = Text8.text + "8766"
End Sub
Private Sub sdqy_Click()
Text8.text = Text8.text + "101.325J"
End Sub
Private Sub sgcdqy_Click()
Text8.text = Text8.text + "98.0665J"
End Sub
Private Sub ssd_Click()
Text8.text = Text8.text + "0.01"
End Sub
Private Sub ssdk_Click()
Text8.text = Text8.text + "4.1855J"
End Sub
Private Sub swy_Click()
Text8.text = Text8.text + "29.530589"
End Sub
Private Sub triangle_Click()
tri.Show 0
End Sub
Private Sub twdw_Click()
Text8.text = Text8.text + "(1.4959787E+11)"
End Sub
Private Sub tyn_Click()
Text8.text = Text8.text + "354.3671"
End Sub
Private Sub yczg_Click()
Text8.text = Text8.text + "10.76lx"
End Sub
Private Sub yhxr_Click()
Text8.text = Text8.text + "0.99726957"
End Sub
Private Sub yjl_Click()
Text8.text = Text8.text + "4.545L"
End Sub
Private Sub yl_Click()
Text8.text = Text8.text + "1.609344km"
End Sub
Private Sub ylcl_Click()
Text8.text = Text8.text + "(6.67259E-11)"
End Sub
Private Sub yzl_Click()
Text8.text = Text8.text + "3.14159265358979323846264338327950288419716939931148196659300057"
End Sub
Private Sub yzml_Click()
Text8.text = Text8.text + "754.7W"
End Sub
Private Sub yzzldw_Click()
Text8.text = Text8.text + "(1.6605402E-27)"
End Sub
Private Sub zhikg_Click()
Text8.text = Text8.text + "(1.6726231E-27)"
End Sub
Private Sub zhiyz_Click()
Text8.text = Text8.text + "1.007276470"
End Sub
Private Sub zhizcj_Click()
Text8.text = Text8.text + "(1.41060761E-26)"
End Sub
Private Sub zhizdkpdbc_Click()
Text8.text = Text8.text + "(1.32141002E-15)"
End Sub
Private Sub zhongkg_Click()
Text8.text = Text8.text + "(1.6749286E-27)"
End Sub
Private Sub zhongyz_Click()
Text8.text = Text8.text + "1.008664904"
End Sub
Private Sub zhongzdkpdbc_Click()
Text8.text = Text8.text + "(1.31959110E-15)"
End Sub
Private Sub zkdrl_Click()
Text8.text = Text8.text + "(8.854187817E-12)"
End Sub
Private Sub zkzcdl_Click()
Text8.text = Text8.text + "(1.2566370614E-6)"
End Sub
Private Sub zkzgs_Click()
Text8.text = Text8.text + "(2.99792458E+8)"
End Sub
Private Sub zrsdsdd_Click()
Text8.text = Text8.text + "2.718281828459045235360287471353"
End Sub

Private Sub znbx_Click()
polygon.Show
End Sub
