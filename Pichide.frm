VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Picdraw 
   Caption         =   "附件"
   ClientHeight    =   8505
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11085
   Icon            =   "Pichide.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8505
   ScaleMode       =   0  'User
   ScaleWidth      =   12000
   StartUpPosition =   3  '窗口缺省
   Begin VB.Frame Frame3 
      Height          =   1095
      Left            =   4320
      TabIndex        =   39
      Top             =   120
      Width           =   2655
      Begin VB.CommandButton Command9 
         Caption         =   "关于直线的对称点坐标"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   240
         TabIndex        =   41
         ToolTipText     =   "按左边的基本元素作椭圆"
         Top             =   600
         Width           =   2295
      End
      Begin VB.CommandButton Command8 
         Caption         =   "求两点距离与点线距离"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   240
         TabIndex        =   40
         ToolTipText     =   "按左边的基本元素作椭圆"
         Top             =   240
         Width           =   2295
      End
   End
   Begin VB.CommandButton Command7 
      Caption         =   "清除"
      Height          =   855
      Left            =   3120
      Picture         =   "Pichide.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   37
      ToolTipText     =   "清空图像"
      Top             =   240
      Width           =   735
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   120
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command6 
      Caption         =   "打开"
      Height          =   855
      Left            =   2400
      Picture         =   "Pichide.frx":1194
      Style           =   1  'Graphical
      TabIndex        =   36
      ToolTipText     =   "打开图像"
      Top             =   240
      Width           =   735
   End
   Begin VB.CommandButton Command5 
      Caption         =   "保存"
      Height          =   855
      Left            =   1680
      Picture         =   "Pichide.frx":171E
      Style           =   1  'Graphical
      TabIndex        =   35
      ToolTipText     =   "保存图像"
      Top             =   240
      Width           =   735
   End
   Begin VB.CommandButton Command3 
      Caption         =   "编辑"
      Height          =   855
      Left            =   960
      Picture         =   "Pichide.frx":1FE8
      Style           =   1  'Graphical
      TabIndex        =   34
      ToolTipText     =   "编辑图画"
      Top             =   240
      Width           =   735
   End
   Begin VB.Frame Frame2 
      Caption         =   "通过平面上两点作直线"
      Height          =   1935
      Left            =   7320
      TabIndex        =   25
      Top             =   240
      Width           =   3495
      Begin VB.CommandButton Command2 
         Caption         =   "连结两点"
         Height          =   375
         Left            =   600
         TabIndex        =   33
         Top             =   1440
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         Caption         =   "过两点作直线"
         Height          =   375
         Left            =   1800
         TabIndex        =   32
         Top             =   1440
         Width           =   1455
      End
      Begin VB.TextBox ly2 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   10.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1920
         MousePointer    =   3  'I-Beam
         TabIndex        =   30
         ToolTipText     =   "y2"
         Top             =   840
         Width           =   735
      End
      Begin VB.TextBox lx2 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   10.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   840
         MousePointer    =   3  'I-Beam
         TabIndex        =   29
         ToolTipText     =   "x2"
         Top             =   840
         Width           =   735
      End
      Begin VB.TextBox ly1 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   10.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1920
         MousePointer    =   3  'I-Beam
         TabIndex        =   28
         ToolTipText     =   "y1"
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox lx1 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   10.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   840
         MousePointer    =   3  'I-Beam
         TabIndex        =   27
         ToolTipText     =   "x1"
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "(           ,           )"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   600
         TabIndex        =   31
         Top             =   960
         Width           =   2295
      End
      Begin VB.Label Label1 
         Caption         =   "(           ,           )"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   600
         TabIndex        =   26
         Top             =   360
         Width           =   2295
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "通过平面上的三点作二次曲线"
      Height          =   3975
      Left            =   7560
      TabIndex        =   14
      Top             =   4320
      Width           =   2655
      Begin VB.CommandButton zyuan 
         Caption         =   "作圆"
         Height          =   735
         Left            =   600
         TabIndex        =   38
         Top             =   2760
         Width           =   735
      End
      Begin VB.CommandButton zpwx 
         Caption         =   "作二次抛物线"
         Height          =   735
         Left            =   1560
         TabIndex        =   21
         Top             =   2760
         Width           =   735
      End
      Begin VB.TextBox ty3 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   10.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1560
         MousePointer    =   3  'I-Beam
         TabIndex        =   20
         ToolTipText     =   "y3"
         Top             =   1800
         Width           =   735
      End
      Begin VB.TextBox tx3 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   10.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         MousePointer    =   3  'I-Beam
         TabIndex        =   19
         ToolTipText     =   "x3"
         Top             =   1800
         Width           =   735
      End
      Begin VB.TextBox ty2 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   10.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1560
         MousePointer    =   3  'I-Beam
         TabIndex        =   18
         ToolTipText     =   "y2"
         Top             =   1200
         Width           =   735
      End
      Begin VB.TextBox tx2 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   10.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         MousePointer    =   3  'I-Beam
         TabIndex        =   17
         ToolTipText     =   "x2"
         Top             =   1200
         Width           =   735
      End
      Begin VB.TextBox ty1 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   10.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1560
         MousePointer    =   3  'I-Beam
         TabIndex        =   16
         ToolTipText     =   "y1"
         Top             =   600
         Width           =   735
      End
      Begin VB.TextBox tx1 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   10.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         MousePointer    =   3  'I-Beam
         TabIndex        =   15
         ToolTipText     =   "x1"
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "(           ,           )"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   24
         Top             =   1920
         Width           =   2295
      End
      Begin VB.Label Label1 
         Caption         =   "(           ,           )"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   23
         Top             =   1320
         Width           =   2295
      End
      Begin VB.Label Label1 
         Caption         =   "(           ,           )"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   22
         Top             =   720
         Width           =   2295
      End
   End
   Begin VB.Frame cir 
      Caption         =   "作椭圆"
      Height          =   1815
      Left            =   7320
      MousePointer    =   1  'Arrow
      TabIndex        =   5
      ToolTipText     =   "圆形工具"
      Top             =   2280
      Width           =   3495
      Begin VB.CommandButton Command4 
         Caption         =   "作圆"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   2520
         TabIndex        =   10
         ToolTipText     =   "按左边的基本元素作椭圆"
         Top             =   1200
         Width           =   615
      End
      Begin VB.TextBox zxzbx 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   10.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1320
         MousePointer    =   3  'I-Beam
         TabIndex        =   9
         ToolTipText     =   "椭圆中心的横坐标 "
         Top             =   240
         Width           =   615
      End
      Begin VB.TextBox zxzby 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2160
         MousePointer    =   3  'I-Beam
         TabIndex        =   8
         ToolTipText     =   "椭圆中心的纵坐标 "
         Top             =   240
         Width           =   615
      End
      Begin VB.TextBox banjing 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   10.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1320
         MousePointer    =   3  'I-Beam
         TabIndex        =   7
         Text            =   "3"
         ToolTipText     =   "这里指椭圆长轴的长度的一半"
         Top             =   840
         Width           =   735
      End
      Begin VB.TextBox ysxx 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   10.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1320
         MousePointer    =   3  'I-Beam
         TabIndex        =   6
         Text            =   "1.0"
         ToolTipText     =   "圆短轴与长轴的尺寸比"
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label ybj 
         Caption         =   "半径"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   480
         TabIndex        =   13
         Top             =   960
         Width           =   615
      End
      Begin VB.Label Label17 
         Caption         =   "中心坐标 (        ,        )"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   12
         Top             =   360
         Width           =   2655
      End
      Begin VB.Label Label18 
         Caption         =   "压缩系数"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   480
         TabIndex        =   11
         Top             =   1440
         Width           =   735
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "作二次曲线"
      Height          =   6855
      Left            =   240
      TabIndex        =   42
      Top             =   1320
      Width           =   6975
      Begin VB.CommandButton zecqx 
         Caption         =   "作椭圆"
         Height          =   615
         Left            =   5280
         TabIndex        =   4
         Top             =   5760
         Width           =   1215
      End
      Begin VB.TextBox b 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   10.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3360
         MousePointer    =   3  'I-Beam
         TabIndex        =   1
         ToolTipText     =   "b"
         Top             =   1400
         Width           =   735
      End
      Begin VB.TextBox a 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   10.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   960
         MousePointer    =   3  'I-Beam
         TabIndex        =   0
         ToolTipText     =   "a"
         Top             =   1400
         Width           =   735
      End
      Begin VB.TextBox y0 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   10.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3600
         MousePointer    =   3  'I-Beam
         TabIndex        =   3
         ToolTipText     =   "y0"
         Top             =   680
         Width           =   735
      End
      Begin VB.TextBox x0 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   10.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1200
         MousePointer    =   3  'I-Beam
         TabIndex        =   2
         ToolTipText     =   "x0"
         Top             =   680
         Width           =   735
      End
      Begin VB.OptionButton p 
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2400
         Style           =   1  'Graphical
         TabIndex        =   44
         ToolTipText     =   "椭圆"
         Top             =   840
         Value           =   -1  'True
         Width           =   495
      End
      Begin VB.OptionButton m 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2400
         Style           =   1  'Graphical
         TabIndex        =   43
         ToolTipText     =   "双曲线"
         Top             =   1200
         Width           =   495
      End
      Begin VB.TextBox x 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   10.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   600
         MousePointer    =   3  'I-Beam
         TabIndex        =   78
         ToolTipText     =   "x0"
         Top             =   4320
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label Rst 
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
         Height          =   375
         Index           =   13
         Left            =   1800
         TabIndex        =   77
         Top             =   6000
         Width           =   2895
      End
      Begin VB.Label Rst 
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
         Height          =   375
         Index           =   12
         Left            =   1800
         TabIndex        =   76
         Top             =   5520
         Width           =   2895
      End
      Begin VB.Label Rst 
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
         Height          =   375
         Index           =   11
         Left            =   1800
         TabIndex        =   75
         Top             =   4920
         Width           =   2415
      End
      Begin VB.Label Rst 
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
         Height          =   375
         Index           =   10
         Left            =   5880
         TabIndex        =   74
         Top             =   4320
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Rst 
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
         Height          =   375
         Index           =   9
         Left            =   3360
         TabIndex        =   73
         Top             =   4320
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Rst 
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
         Height          =   375
         Index           =   8
         Left            =   1800
         TabIndex        =   72
         Top             =   3840
         Width           =   2535
      End
      Begin VB.Label Rst 
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
         Height          =   375
         Index           =   7
         Left            =   5520
         TabIndex        =   71
         Top             =   3360
         Width           =   855
      End
      Begin VB.Label Rst 
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
         Height          =   375
         Index           =   6
         Left            =   4440
         TabIndex        =   70
         Top             =   3360
         Width           =   855
      End
      Begin VB.Label Rst 
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
         Height          =   375
         Index           =   5
         Left            =   2280
         TabIndex        =   69
         Top             =   3360
         Width           =   855
      End
      Begin VB.Label Rst 
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
         Height          =   375
         Index           =   4
         Left            =   1200
         TabIndex        =   68
         Top             =   3360
         Width           =   855
      End
      Begin VB.Label Rst 
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
         Height          =   375
         Index           =   3
         Left            =   4440
         TabIndex        =   67
         Top             =   2760
         Width           =   855
      End
      Begin VB.Label Rst 
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
         Height          =   375
         Index           =   2
         Left            =   1200
         TabIndex        =   66
         Top             =   2760
         Width           =   855
      End
      Begin VB.Label Rst 
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
         Height          =   375
         Index           =   1
         Left            =   2400
         TabIndex        =   65
         Top             =   2160
         Width           =   855
      End
      Begin VB.Label Rst 
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
         Height          =   375
         Index           =   0
         Left            =   1320
         TabIndex        =   64
         Top             =   2160
         Width           =   855
      End
      Begin VB.Label Label6 
         Caption         =   "准线方程 l2:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   2
         Left            =   360
         TabIndex        =   63
         Top             =   6120
         Width           =   1575
      End
      Begin VB.Label Label6 
         Caption         =   "准线方程 l1:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   1
         Left            =   360
         TabIndex        =   62
         Top             =   5640
         Width           =   1575
      End
      Begin VB.Label Label5 
         Caption         =   "x =               则"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   10.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   8
         Left            =   240
         TabIndex        =   61
         Top             =   4320
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label Label5 
         Caption         =   "焦点半径 r2 ="
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   10.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   7
         Left            =   4440
         TabIndex        =   60
         Top             =   4320
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Label Label5 
         Caption         =   "焦点半径 r1 ="
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   10.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   6
         Left            =   1920
         TabIndex        =   59
         Top             =   4320
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Label Label5 
         Caption         =   "焦点 F2(                ,                  )"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   10.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   5
         Left            =   3600
         TabIndex        =   58
         Top             =   3360
         Width           =   3135
      End
      Begin VB.Label Label5 
         Caption         =   "焦点 F1(                ,                 )"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   10.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   4
         Left            =   360
         TabIndex        =   57
         Top             =   3360
         Width           =   2895
      End
      Begin VB.Label Label5 
         Caption         =   "短轴"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   10.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   3
         Left            =   3840
         TabIndex        =   56
         Top             =   2760
         Width           =   1095
      End
      Begin VB.Label Label5 
         Caption         =   "长轴"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   10.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   2
         Left            =   600
         TabIndex        =   55
         Top             =   2760
         Width           =   1095
      End
      Begin VB.Label Label5 
         Caption         =   "中心 （                ，              ）          "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   10.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   600
         TabIndex        =   54
         Top             =   2160
         Width           =   3015
      End
      Begin VB.Label Label6 
         Caption         =   "离心率 e ="
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   10.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   0
         Left            =   480
         TabIndex        =   53
         Top             =   4920
         Width           =   1095
      End
      Begin VB.Label Label5 
         Caption         =   "焦距 c ="
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   10.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   0
         Left            =   720
         TabIndex        =   52
         Top             =   3840
         Width           =   1095
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "=1"
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
         Left            =   4800
         TabIndex        =   51
         Top             =   960
         Width           =   615
      End
      Begin VB.Label power2 
         BackStyle       =   0  'Transparent
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   1680
         TabIndex        =   50
         Top             =   1200
         Width           =   255
      End
      Begin VB.Label power2 
         BackStyle       =   0  'Transparent
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   1920
         TabIndex        =   49
         Top             =   480
         Width           =   255
      End
      Begin VB.Label power2 
         BackStyle       =   0  'Transparent
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   4320
         TabIndex        =   48
         Top             =   480
         Width           =   255
      End
      Begin VB.Label power2 
         BackStyle       =   0  'Transparent
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   4080
         TabIndex        =   47
         Top             =   1200
         Width           =   255
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         Index           =   1
         X1              =   3000
         X2              =   4680
         Y1              =   1200
         Y2              =   1200
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         Index           =   0
         X1              =   600
         X2              =   2280
         Y1              =   1200
         Y2              =   1200
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "(x-         )"
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
         Left            =   600
         TabIndex        =   46
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "(y-         )"
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
         Left            =   3000
         TabIndex        =   45
         Top             =   600
         Width           =   1695
      End
   End
End
Attribute VB_Name = "Picdraw"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()
x1 = Val(lx1.Text)
y1 = Val(ly1.Text)
x2 = Val(lx2.Text)
y2 = Val(ly2.Text)

If x2 <> x1 Then
k = Trim(Str((y2 - y1) / (x2 - x1)))
b = Trim(Str((y1 * x2 - y2 * x1) / (x2 - x1)))
Pic.Text1.Text = "(" + k + ")*x+(" + b + ")"
Pic.ExplicitFun.Checked = True
Pic.ImplicitFun.Checked = False
Pic.Text1.SetFocus
SendKeys "{enter}"
Else
Pic.Picture1.Line (x1, Val(Pic.ymin.Text))-(x1, Val(Pic.ymax.Text))
End If

End Sub

Private Sub Command2_Click()
Pic.Picture1.Line (Val(lx1.Text), Val(ly1.Text))-(Val(lx2.Text), Val(ly2.Text))
End Sub

Private Sub Command3_Click()
On Error GoTo l:
SavePicture Pic.Picture1.Image, App.Path & "\edit.bmp"
s = Shell("mspaint.exe " & App.Path & "\edit.bmp", 1)
l: If err <> 0 Then msg = MsgBox("找不到文件 mspaint.exe 。", vbInformation, "错误"): Resume Next
End Sub

Private Sub Command4_Click()
On Error Resume Next
Pic.Picture1.Circle (Val(zxzbx.Text), Val(zxzby.Text)), Val(banjing.Text), , , , Val(ysxx.Text)
End Sub

Private Sub Command5_Click()
CommonDialog1.ShowSave
If CommonDialog1.FileName <> "" And InStr(CommonDialog1.FileName, ".") = 0 Then
CommonDialog1.FileName = CommonDialog1.FileName & ".bmp"
End If
If CommonDialog1.FileName <> "" Then SavePicture Pic.Picture1.Image, CommonDialog1.FileName

End Sub

Private Sub Command6_Click()
On Error Resume Next
CommonDialog1.ShowOpen
Pic.Picture1.Picture = LoadPicture(CommonDialog1.FileName)
End Sub

Private Sub Command7_Click()
Pic.Picture1.Picture = LoadPicture()
End Sub

Private Sub Command8_Click()
distance.Show
End Sub

Private Sub Command9_Click()
symmetry.Show
End Sub





Private Sub m_Click()
Label5(2).Caption = "实轴"
Label5(3).Caption = "虚轴"
 zecqx.Caption = "作双曲线"
End Sub

Private Sub p_Click()
  Label5(2).Caption = "长轴"
  Label5(3).Caption = "短轴"

  zecqx.Caption = "作椭圆"
End Sub

Private Sub zecqx_Click()
On Error Resume Next
x0 = Val(x0.Text)
y0 = Val(y0.Text)
a = Val(a.Text)
b = Val(b.Text)
x = Val(x.Text)

If a >= b Then a0 = a: b0 = b Else a0 = b: b0 = a
Rst(0).Caption = x0
Rst(1).Caption = y0
Rst(2).Caption = 2 * a0
Rst(3).Caption = 2 * b0




Pic.Show


If p.Value = True Then
  Rst(8) = Sqr(a0 ^ 2 - b0 ^ 2)
  If a >= b Then
    Pic.Picture1.Circle (x0, y0), a, , , , b / a
  Else
    Pic.Picture1.Circle (x0, y0), b, , , , b / a
  End If
   
Else
  Rst(8).Caption = Hypot(a, b)
  
  prmtfct.tl.Text = "0": prmtfct.tr.Text = "6.283185307"
  prmtfct.xpa.Text = "(" & a.Text & ")*sect+(" & x0.Text & ")"
  prmtfct.ypa.Text = "(" & b.Text & ")*tant+(" & y0.Text & ")"
  Call prmtfct.Draw_Click
  prmtfct.WindowState = 1
End If

Rst(11) = Rst(8) / a0
If a >= b Then
  Rst(4).Caption = "-" & Rst(8)
  Rst(5).Caption = "0"
  Rst(6).Caption = Rst(8)
  Rst(7).Caption = "0"
  Rst(12).Caption = "x=" & a0 ^ 2 / Val(Rst(8)) + x0
  Rst(13).Caption = "x=" & -a0 ^ 2 / Val(Rst(8)) + x0
Else
  Rst(5).Caption = "-" & Rst(8)
  Rst(4).Caption = "0"
  Rst(7).Caption = Rst(8)
  Rst(6).Caption = "0"
  Rst(12).Caption = "y=" & a0 ^ 2 / Val(Rst(8)) + y0
  Rst(13).Caption = "y=" & -a0 ^ 2 / Val(Rst(8)) + y0
End If
Rst(4).Caption = Val(Rst(4).Caption) + x0
Rst(5).Caption = Val(Rst(5).Caption) + y0
Rst(6).Caption = Val(Rst(6).Caption) + x0
Rst(7).Caption = Val(Rst(7).Caption) + y0
End Sub

Private Sub zpwx_Click()
On Error GoTo l1:
x1 = Val(tx1.Text)
y1 = Val(ty1.Text)
x2 = Val(tx2.Text)
y2 = Val(ty2.Text)
x3 = Val(tx3.Text)
y3 = Val(ty3.Text)

a = "(" + Trim(Str(y1 / ((x1 - x2) * (x1 - x3)))) + ")*"
b = "(" + Trim(Str(y2 / ((x2 - x1) * (x2 - x3)))) + ")*"
c = "(" + Trim(Str(y3 / ((x3 - x1) * (x3 - x2)))) + ")*"

Pic.Text1.Text = _
a + "(x-" + "(" + Trim(Str(x2)) + "))*(x-" + "(" + Trim(Str(x3)) + "))" _
+ "+" + b + "(x-" + "(" + Trim(Str(x1)) + "))*(x-" + "(" + Trim(Str(x3)) + "))" _
+ "+" + c + "(x-" + "(" + Trim(Str(x1)) + "))*(x-" + "(" + Trim(Str(x2)) + "))"
Pic.Text1.SetFocus
If Pic.ExplicitFun.Checked = False Then Pic.ExplicitFun.Checked = True
SendKeys "{enter}"
l1: If err Then msg = MsgBox("计算器无法完成一个绘图操作。", vbOKOnly, "错误")
End Sub

Private Sub zyuan_Click()
On Error GoTo l1:
x1 = Val(tx1.Text)
y1 = Val(ty1.Text)
x2 = Val(tx2.Text)
y2 = Val(ty2.Text)
x3 = Val(tx3.Text)
y3 = Val(ty3.Text)

a = (y1 - y2) * (x3 - x2) * (x3 + x2)
b = (y3 - y2) * (x1 - x2) * (x1 + x2)
c = (y3 - y1) * (y3 - y2) * (y1 - y2)
d = 2 * ((x3 - x2) * (y1 - y2) - (y3 - y2) * (x1 - x2))
x = (a - b + c) / d
Y = (x2 - x1) * (x - (x1 + x2) / 2) / (y1 - y2) + (y1 + y2) / 2
r = Sqr((x - x1) ^ 2 + (Y - y1) ^ 2)
Pic.Picture1.Circle (x, Y), r
l1: If err Then msg = MsgBox("计算器无法完成一个绘图操作。", vbOKOnly, "错误")
End Sub
