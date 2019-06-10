VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Begin VB.Form Pic 
   BackColor       =   &H00C0C0C0&
   Caption         =   "方程曲线查看器"
   ClientHeight    =   8190
   ClientLeft      =   2205
   ClientTop       =   555
   ClientWidth     =   7965
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Line.frx":0000
   LinkTopic       =   "Form1"
   MouseIcon       =   "Line.frx":08CA
   ScaleHeight     =   546
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   531
   Begin VB.Timer Text1h 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   9960
      Top             =   8040
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   285
      Left            =   120
      TabIndex        =   49
      Text            =   "T"
      Top             =   10200
      Width           =   2295
   End
   Begin VB.Timer TextSetFocus 
      Interval        =   1
      Left            =   10560
      Top             =   8040
   End
   Begin VB.PictureBox mainmenu 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   270
      MouseIcon       =   "Line.frx":0A1C
      MousePointer    =   99  'Custom
      Picture         =   "Line.frx":0B6E
      ScaleHeight     =   210
      ScaleWidth      =   180
      TabIndex        =   39
      ToolTipText     =   "主菜单 F10"
      Top             =   135
      Width           =   180
   End
   Begin VB.TextBox mousey 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   210
      Left            =   7740
      TabIndex        =   23
      Text            =   "y"
      ToolTipText     =   "y  或  ρ "
      Top             =   4200
      Width           =   1875
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   405
      Left            =   1560
      MousePointer    =   3  'I-Beam
      MultiLine       =   -1  'True
      TabIndex        =   0
      ToolTipText     =   "曲线方程输入区 (变量只能是 x、y, 极坐标下也请使用 x 作为变量。若在方程后添加分号, 则在绘图前先清空原图)"
      Top             =   30
      Width           =   3975
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   60
      Left            =   210
      TabIndex        =   21
      Top             =   510
      Width           =   7545
      _ExtentX        =   13309
      _ExtentY        =   106
      _Version        =   393216
      Appearance      =   0
      Max             =   1000
      Scrolling       =   1
   End
   Begin VB.Frame Frame5 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5535
      Left            =   225
      MousePointer    =   1  'Arrow
      TabIndex        =   40
      ToolTipText     =   "双击此处恢复默认值"
      Top             =   525
      Visible         =   0   'False
      Width           =   2055
      Begin VB.TextBox xmin 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         MousePointer    =   3  'I-Beam
         TabIndex        =   13
         Text            =   "-10"
         ToolTipText     =   "横坐标最小值"
         Top             =   720
         Width           =   1755
      End
      Begin VB.TextBox xmax 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         MousePointer    =   3  'I-Beam
         TabIndex        =   14
         Text            =   "10"
         ToolTipText     =   "横坐标最大值"
         Top             =   1440
         Width           =   1755
      End
      Begin VB.TextBox ymin 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         MousePointer    =   3  'I-Beam
         TabIndex        =   15
         Text            =   "-10"
         ToolTipText     =   "纵坐标最小值"
         Top             =   2400
         Width           =   1755
      End
      Begin VB.TextBox ymax 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         MousePointer    =   3  'I-Beam
         TabIndex        =   16
         Text            =   "10"
         ToolTipText     =   "纵坐标最大值"
         Top             =   3120
         Width           =   1755
      End
      Begin VB.TextBox Xinc 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         MousePointer    =   3  'I-Beam
         TabIndex        =   17
         Text            =   "1"
         ToolTipText     =   "水平网格间距"
         Top             =   4200
         Width           =   1755
      End
      Begin VB.TextBox Yinc 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         MousePointer    =   3  'I-Beam
         TabIndex        =   18
         Text            =   "1"
         ToolTipText     =   "垂直网格间距"
         Top             =   4920
         Width           =   1755
      End
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   1440
         Top             =   4440
      End
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         Caption         =   "网格间距"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   255
         Left            =   120
         TabIndex        =   52
         Top             =   3720
         Width           =   1335
      End
      Begin VB.Label Label19 
         BackStyle       =   0  'Transparent
         Caption         =   "垂直显示范围"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   255
         Left            =   120
         TabIndex        =   51
         Top             =   1920
         Width           =   1335
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   "水平显示范围"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   255
         Left            =   120
         TabIndex        =   50
         Top             =   240
         Width           =   1335
      End
      Begin VB.Image xyalarm 
         Height          =   240
         Left            =   1440
         Picture         =   "Line.frx":0E5B
         Stretch         =   -1  'True
         ToolTipText     =   "注意: abs(xmax-xmin) <> abs(ymax-ymin)"
         Top             =   360
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "X Min"
         BeginProperty Font 
            Name            =   "@宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   46
         ToolTipText     =   "单击此处使 Xmin = -Xmax"
         Top             =   480
         Width           =   615
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "X Max"
         BeginProperty Font 
            Name            =   "@宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   45
         ToolTipText     =   "单击此处使 Xmax = -Xmin"
         Top             =   1200
         Width           =   615
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Y Min"
         BeginProperty Font 
            Name            =   "@宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   44
         ToolTipText     =   "单击此处使 Ymin = Xmin"
         Top             =   2160
         Width           =   615
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Y Max"
         BeginProperty Font 
            Name            =   "@宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   43
         ToolTipText     =   "单击此处使 Ymax = Xmax"
         Top             =   2880
         Width           =   615
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "X Increment"
         BeginProperty Font 
            Name            =   "@宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   42
         ToolTipText     =   "单击此处使 X Increment = Y Increment"
         Top             =   3960
         Width           =   1215
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Y Increment"
         BeginProperty Font 
            Name            =   "@宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   41
         ToolTipText     =   "单击此处使 Y Increment = X Increment"
         Top             =   4680
         Width           =   1335
      End
   End
   Begin VB.TextBox linewide 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1.50000e5
      TabIndex        =   38
      Text            =   "1"
      Top             =   120
      Width           =   495
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "精度"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1875
      Left            =   8400
      TabIndex        =   34
      Top             =   600
      Width           =   3135
      Begin VB.ComboBox precision3 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         ItemData        =   "Line.frx":0FA5
         Left            =   1560
         List            =   "Line.frx":0FBB
         TabIndex        =   6
         Text            =   "10"
         ToolTipText     =   "若设置值过低, 则无法绘制曲线; 设置值越大, 绘图速度越慢。该选项仅在绘隐函数 F(x,y)=0 的图像时有效"
         Top             =   1320
         Width           =   855
      End
      Begin VB.ComboBox precision 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400040&
         Height          =   360
         ItemData        =   "Line.frx":0FD3
         Left            =   1560
         List            =   "Line.frx":1004
         TabIndex        =   5
         Text            =   "3"
         ToolTipText     =   $"Line.frx":1044
         Top             =   840
         Width           =   855
      End
      Begin VB.ComboBox Combo2 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         ItemData        =   "Line.frx":10CC
         Left            =   1560
         List            =   "Line.frx":10EE
         TabIndex        =   4
         Text            =   "100"
         ToolTipText     =   "自变量 x 的绘图精度 (数值越大，精度越高，绘图越慢)"
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label16 
         BackColor       =   &H00C0C0C0&
         Caption         =   "因变量精度"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   37
         ToolTipText     =   "1/dy"
         Top             =   1395
         Width           =   1095
      End
      Begin VB.Label Label14 
         BackColor       =   &H00C0C0C0&
         Caption         =   "对比精度"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   530
         TabIndex        =   36
         ToolTipText     =   "epsilon"
         Top             =   880
         Width           =   735
      End
      Begin VB.Label Label8 
         BackColor       =   &H00C0C0C0&
         Caption         =   "自变量精度"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   35
         ToolTipText     =   "k/dx"
         Top             =   315
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "最值"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   8400
      TabIndex        =   30
      ToolTipText     =   "此处显示函数在横轴显示范围内可能的最值(以上精度设置越高, 最值越准)"
      Top             =   5880
      Width           =   3135
      Begin VB.TextBox Text5 
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
         ForeColor       =   &H00800000&
         Height          =   345
         Left            =   120
         TabIndex        =   11
         ToolTipText     =   "函数在显示范围内可能的最大值"
         Top             =   600
         Width           =   2895
      End
      Begin VB.TextBox Text6 
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
         ForeColor       =   &H00800000&
         Height          =   345
         Left            =   120
         TabIndex        =   12
         ToolTipText     =   "函数在显示范围内可能的最小值"
         Top             =   1320
         Width           =   2895
      End
      Begin VB.Label Label6 
         BackColor       =   &H00C0C0C0&
         Caption         =   "最小值 :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   32
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C0C0C0&
         Caption         =   "最大值 :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00C0C0C0&
      Caption         =   "定义域"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8400
      TabIndex        =   26
      ToolTipText     =   "Domain"
      Top             =   2760
      Width           =   3135
      Begin VB.TextBox Text8 
         BackColor       =   &H8000000A&
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
         Left            =   720
         TabIndex        =   7
         ToolTipText     =   "不填写则默认为负无穷大 (极坐标系下默认为 0)"
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox Text9 
         BackColor       =   &H8000000A&
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
         Left            =   1800
         TabIndex        =   8
         ToolTipText     =   "不填写则默认为正无穷大 (极坐标系下默认为 2π)"
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0C0C0&
         Caption         =   "(                  ,                   )"
         Height          =   345
         Left            =   480
         TabIndex        =   27
         Top             =   240
         Width           =   2565
      End
   End
   Begin VB.TextBox mousex 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   180
      Left            =   3960
      TabIndex        =   24
      Text            =   "x"
      ToolTipText     =   "x 或θ"
      Top             =   8040
      Width           =   2115
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   11400
      Top             =   8040
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00E0E0E0&
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   5760
      MousePointer    =   99  'Custom
      Picture         =   "Line.frx":1120
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "点击这里开始/停止绘制方程曲线, 输入表达式直接回车亦可。图像显示后，可以通过鼠标右键拖动图像改变显示范围。"
      Top             =   30
      Width           =   525
   End
   Begin VB.CommandButton Command2 
      DisabledPicture =   "Line.frx":16AA
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   6480
      Picture         =   "Line.frx":17F4
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "清空表达式输入框"
      Top             =   30
      Width           =   525
   End
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   7080
      Picture         =   "Line.frx":1D7E
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "清除图象  Esc  (要清除包括网格在内的所有背景图像, 请单击右边的""V"")"
      Top             =   30
      Width           =   525
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   7500
      Left            =   225
      MouseIcon       =   "Line.frx":2308
      MousePointer    =   2  'Cross
      ScaleHeight     =   -20
      ScaleLeft       =   -10
      ScaleMode       =   0  'User
      ScaleTop        =   10
      ScaleWidth      =   20
      TabIndex        =   19
      Top             =   525
      Width           =   7500
      Begin VB.Line Line1 
         BorderStyle     =   0  'Transparent
         X1              =   0
         X2              =   0
         Y1              =   10
         Y2              =   -10
      End
      Begin VB.Line Liney 
         BorderColor     =   &H00C0C000&
         BorderStyle     =   3  'Dot
         Visible         =   0   'False
         X1              =   -0.08
         X2              =   -0.08
         Y1              =   10
         Y2              =   -9.84
      End
      Begin VB.Line Linex 
         BorderColor     =   &H00C0C000&
         BorderStyle     =   3  'Dot
         Visible         =   0   'False
         X1              =   -10
         X2              =   9.84
         Y1              =   0.08
         Y2              =   0.08
      End
   End
   Begin VB.Frame mapping 
      BackColor       =   &H00C0C0C0&
      Caption         =   "映射"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   8400
      TabIndex        =   28
      ToolTipText     =   "此功能在显式 y=f(x) 下有效。"
      Top             =   3720
      Width           =   3135
      Begin VB.TextBox ResultD 
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
         ForeColor       =   &H00404040&
         Height          =   345
         Left            =   120
         TabIndex        =   53
         Text            =   "f'(x0)"
         ToolTipText     =   "导数 f'(x0)"
         Top             =   1450
         Width           =   2895
      End
      Begin VB.TextBox Result 
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
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   120
         TabIndex        =   10
         Text            =   "f(x0)"
         ToolTipText     =   "象 f(x0)"
         Top             =   1000
         Width           =   2895
      End
      Begin VB.TextBox ifxtheny 
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
         ForeColor       =   &H00800000&
         Height          =   345
         Left            =   120
         MaxLength       =   255
         MousePointer    =   3  'I-Beam
         MultiLine       =   -1  'True
         TabIndex        =   9
         ToolTipText     =   "在此处输入原象 x0 (可以是数学表达式)"
         Top             =   480
         Width           =   2895
      End
      Begin VB.Label Label7 
         BackColor       =   &H00C0C0C0&
         Caption         =   "f :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   33
         ToolTipText     =   "对应法则就是函数曲线输入框内的表达式。"
         Top             =   240
         Width           =   615
      End
      Begin VB.Label label15 
         BackColor       =   &H00C0C0C0&
         Caption         =   "→"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   29
         ToolTipText     =   "当 x=原象 时, f(x) 的值在下面显示"
         Top             =   720
         Width           =   615
      End
   End
   Begin VB.PictureBox Advance 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   750
      Left            =   7800
      MouseIcon       =   "Line.frx":248E
      Picture         =   "Line.frx":2798
      ScaleHeight     =   750
      ScaleWidth      =   90
      TabIndex        =   25
      ToolTipText     =   "切换到高级/标准视图"
      Top             =   3915
      Width           =   90
   End
   Begin VB.Label Label17 
      BackColor       =   &H00C0C0C0&
      Caption         =   "= 0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4800
      TabIndex        =   48
      ToolTipText     =   "双击此处切换到显函数模式"
      Top             =   0
      Width           =   615
   End
   Begin VB.Label clearimg 
      BackColor       =   &H00C0C0C0&
      Caption         =   "v"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00A56E3A&
      Height          =   255
      Left            =   7680
      TabIndex        =   47
      ToolTipText     =   "清除全部图像"
      Top             =   150
      Width           =   135
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "y ="
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   840
      TabIndex        =   20
      ToolTipText     =   "双击此处切换到隐函数模式"
      Top             =   0
      Width           =   615
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0C0&
      Height          =   5415
      Left            =   0
      TabIndex        =   22
      Top             =   600
      Width           =   105
   End
   Begin VB.Menu popup 
      Caption         =   "菜单"
      Visible         =   0   'False
      Begin VB.Menu files 
         Caption         =   "文件"
         Begin VB.Menu savepic 
            Caption         =   "保存图象"
         End
         Begin VB.Menu loadpic 
            Caption         =   "打开图象"
         End
         Begin VB.Menu bjtx 
            Caption         =   "编辑图像"
         End
      End
      Begin VB.Menu fun 
         Caption         =   "支持的函数/运算符"
         Begin VB.Menu basicOper 
            Caption         =   "基本运算符"
            Begin VB.Menu op 
               Caption         =   " +"
               Index           =   1
            End
            Begin VB.Menu op 
               Caption         =   " -"
               Index           =   2
            End
            Begin VB.Menu op 
               Caption         =   " *"
               Index           =   3
            End
            Begin VB.Menu op 
               Caption         =   " /"
               Index           =   4
            End
            Begin VB.Menu op 
               Caption         =   " ^"
               Index           =   5
            End
            Begin VB.Menu op 
               Caption         =   " @"
               Index           =   6
            End
            Begin VB.Menu op 
               Caption         =   "abs"
               Index           =   7
            End
            Begin VB.Menu op 
               Caption         =   "sqr"
               Index           =   8
            End
         End
         Begin VB.Menu trigfun 
            Caption         =   "三角函数"
            Begin VB.Menu o 
               Caption         =   "sin"
               Index           =   1
            End
            Begin VB.Menu o 
               Caption         =   "Arcsin"
               Index           =   2
            End
            Begin VB.Menu o 
               Caption         =   "sh"
               Index           =   3
            End
            Begin VB.Menu o 
               Caption         =   "arsh"
               Index           =   4
            End
            Begin VB.Menu o 
               Caption         =   "cos"
               Index           =   5
            End
            Begin VB.Menu o 
               Caption         =   "Arccos"
               Index           =   6
            End
            Begin VB.Menu o 
               Caption         =   "ch"
               Index           =   7
            End
            Begin VB.Menu o 
               Caption         =   "Arch"
               Index           =   8
            End
            Begin VB.Menu o 
               Caption         =   "tan"
               Index           =   9
            End
            Begin VB.Menu o 
               Caption         =   "Arctan"
               Index           =   10
            End
            Begin VB.Menu o 
               Caption         =   "th"
               Index           =   11
            End
            Begin VB.Menu o 
               Caption         =   "Arth"
               Index           =   12
            End
            Begin VB.Menu o 
               Caption         =   "cot"
               Index           =   13
            End
            Begin VB.Menu o 
               Caption         =   "Arccot"
               Index           =   14
            End
            Begin VB.Menu o 
               Caption         =   "cth"
               Index           =   15
            End
            Begin VB.Menu o 
               Caption         =   "Arcth"
               Index           =   16
            End
            Begin VB.Menu o 
               Caption         =   "sec"
               Index           =   17
            End
            Begin VB.Menu o 
               Caption         =   "Arcsec"
               Index           =   18
            End
            Begin VB.Menu o 
               Caption         =   "sech"
               Index           =   19
            End
            Begin VB.Menu o 
               Caption         =   "Arsech"
               Index           =   20
            End
            Begin VB.Menu o 
               Caption         =   "csc"
               Index           =   21
            End
            Begin VB.Menu o 
               Caption         =   "Arccsc"
               Index           =   22
            End
            Begin VB.Menu o 
               Caption         =   "csch"
               Index           =   23
            End
            Begin VB.Menu o 
               Caption         =   "Arcsch"
               Index           =   24
            End
         End
         Begin VB.Menu logfun 
            Caption         =   "对数函数"
            Begin VB.Menu ope 
               Caption         =   "exp"
               Index           =   1
            End
            Begin VB.Menu ope 
               Caption         =   "log"
               Index           =   2
            End
            Begin VB.Menu ope 
               Caption         =   "lg"
               Index           =   3
            End
            Begin VB.Menu ope 
               Caption         =   "ln"
               Index           =   4
            End
         End
      End
      Begin VB.Menu hisr 
         Caption         =   "历史记录"
         Begin VB.Menu sytx 
            Caption         =   "上一图象"
         End
         Begin VB.Menu hlg 
            Caption         =   "-"
            Index           =   0
         End
      End
      Begin VB.Menu fgf1 
         Caption         =   "-"
      End
      Begin VB.Menu coordinate 
         Caption         =   "坐标系"
         Begin VB.Menu square 
            Caption         =   "平面直角坐标系"
            Checked         =   -1  'True
         End
         Begin VB.Menu polar1 
            Caption         =   "极坐标系"
         End
      End
      Begin VB.Menu fgf2 
         Caption         =   "-"
      End
      Begin VB.Menu funtape 
         Caption         =   "曲线方程的形式"
         Begin VB.Menu ImplicitFun 
            Caption         =   "隐式 F(x,y)=0"
         End
         Begin VB.Menu ExplicitFun 
            Caption         =   "显式 y=f(x)"
            Checked         =   -1  'True
         End
         Begin VB.Menu cshsh 
            Caption         =   "参数式 x=x(t); y=y(t)"
         End
      End
      Begin VB.Menu unfun1 
         Caption         =   "改为绘制反函数"
      End
      Begin VB.Menu drvfun 
         Caption         =   "同时绘制导函数"
      End
      Begin VB.Menu fhhs 
         Caption         =   "复合函数生成工具"
      End
      Begin VB.Menu fgf11 
         Caption         =   "-"
      End
      Begin VB.Menu auxiliary 
         Caption         =   "辅助计算工具"
         Begin VB.Menu lingdian 
            Caption         =   "求实零点"
         End
         Begin VB.Menu qdsh 
            Caption         =   "求导数"
         End
         Begin VB.Menu qdjf 
            Caption         =   "求定积分"
         End
         Begin VB.Menu zhlc 
            Caption         =   "总和·连乘"
         End
         Begin VB.Menu chzh 
            Caption         =   "拟合与插值"
            Begin VB.Menu zxecf 
               Caption         =   "最小二乘法拟合直线 "
            End
            Begin VB.Menu lang 
               Caption         =   "拉格朗日插值"
            End
         End
      End
      Begin VB.Menu fgf3 
         Caption         =   "-"
      End
      Begin VB.Menu fsdfzscm 
         Caption         =   "负数的非整数次幂"
         Begin VB.Menu ddyl 
            Caption         =   "设置为 都大于零"
         End
         Begin VB.Menu dxyl 
            Caption         =   "设置为 都小于零"
         End
         Begin VB.Menu dwyy 
            Caption         =   "设置为 都无意义"
            Checked         =   -1  'True
         End
      End
      Begin VB.Menu htsc 
         Caption         =   "绘图输出"
         Begin VB.Menu dpm 
            Caption         =   "到屏幕 (边计算边绘图)"
            Checked         =   -1  'True
         End
         Begin VB.Menu dnc 
            Caption         =   "到内存 (计算完成后再绘图)"
         End
         Begin VB.Menu fgf15 
            Caption         =   "-"
         End
         Begin VB.Menu CoverGraph 
            Caption         =   "覆盖原图"
         End
      End
      Begin VB.Menu view 
         Caption         =   "视图"
         Begin VB.Menu fclb 
            Caption         =   "方程列表"
         End
         Begin VB.Menu txkz 
            Caption         =   "图像控制"
         End
         Begin VB.Menu Magnififier 
            Caption         =   "放大镜"
         End
         Begin VB.Menu fgf5 
            Caption         =   "-"
         End
         Begin VB.Menu dline 
            Caption         =   "画线"
            Checked         =   -1  'True
         End
         Begin VB.Menu dpset 
            Caption         =   "描点"
         End
         Begin VB.Menu fgf9 
            Caption         =   "-"
         End
         Begin VB.Menu xoy1 
            Caption         =   "显示坐标轴"
            Checked         =   -1  'True
         End
         Begin VB.Menu Showscale 
            Caption         =   "显示刻度"
            Checked         =   -1  'True
         End
         Begin VB.Menu showweb1 
            Caption         =   "显示网格"
            Checked         =   -1  'True
         End
         Begin VB.Menu mousexoy1 
            Caption         =   "显示坐标"
            Checked         =   -1  'True
         End
         Begin VB.Menu fgf8 
            Caption         =   "-"
         End
         Begin VB.Menu yanse 
            Caption         =   "颜色"
            Begin VB.Menu backboard 
               Caption         =   "背景"
               Begin VB.Menu zdybjs 
                  Caption         =   "自定义"
               End
               Begin VB.Menu white 
                  Caption         =   "白"
               End
               Begin VB.Menu blue 
                  Caption         =   "蓝"
               End
               Begin VB.Menu gray 
                  Caption         =   "灰"
               End
               Begin VB.Menu black 
                  Caption         =   "黑"
               End
            End
            Begin VB.Menu xtys 
               Caption         =   "线条"
            End
            Begin VB.Menu clwg 
               Caption         =   "网格"
            End
            Begin VB.Menu clkd 
               Caption         =   "刻度"
            End
            Begin VB.Menu clzbz 
               Caption         =   "坐标轴"
            End
         End
         Begin VB.Menu fgf 
            Caption         =   "-"
         End
         Begin VB.Menu lw 
            Caption         =   "线宽"
         End
         Begin VB.Menu dashed1 
            Caption         =   "虚线"
         End
      End
      Begin VB.Menu fgf10 
         Caption         =   "-"
      End
      Begin VB.Menu other 
         Caption         =   "其他绘图工具"
      End
      Begin VB.Menu quit 
         Caption         =   "退出"
      End
   End
End
Attribute VB_Name = "Pic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private x As Single
Private xa As Single
Private package As Boolean
Public yh As Single
Dim y As Single
Private hlog As Long
Public Pfc
Private aoo$, erro, Running, drawed, t1e
Public ated As String
Public xoycolor, Webcolor, Scalecolor
Public cleared As Boolean
Private th As Boolean   'th记录Pic.text1是否动态展开过
Private Type PointAPI
x As Long
y As Long
End Type

Private DownX, DownY, MoveX, MoveY

Private Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, lpPoint As PointAPI) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Private Sub Line0(x1, y1, x2, y2, x3, y3, x4, y4)
  Dim Pointapi1 As PointAPI
  If dpm.Checked = True Then
       Picture1.Line (x1, y1)-(x2, y2)
  Else
       mt = MoveToEx(Picture1.hdc, x3, y3, Pointapi1)
       lt = LineTo(Picture1.hdc, x4, y4)
  End If
End Sub
Private Sub Advance_Click()
If Pic.Width <= 8085 Then Pic.Width = 12000 Else Pic.Width = 8085
'If Pic.WindowState = 2 Then Pic.WindowState = 0 Else Pic.WindowState = 2
'Advance.Left = 780
End Sub

Private Sub bjtx_Click()
On Error GoTo l:
SavePicture Pic.Picture1.Image, App.Path & "\未命名.bmp"
s = Shell("mspaint.exe " & App.Path & "\未命名.bmp", 1)
l: If err <> 0 Then msg = MsgBox("找不到文件 mspaint.exe 。", vbInformation, "错误"): Resume Next
End Sub

Private Sub black_Click()
Picture1.BackColor = &H0&
Picture1.ForeColor = &HFFFFFF
End Sub

Private Sub blue_Click()
Picture1.BackColor = &H80000001
Picture1.ForeColor = &HFFFF&
End Sub



Private Sub check1p_Click()
If check1p.Checked = False Then
  check1p.Checked = True
  dline.Checked = True
Else
check1p.Checked = False
  dpset.Checked = True
End If
End Sub



Private Sub clearImg_Click()
On Error Resume Next
If cleared = False And Running = 0 Then SavePicture Picture1.Image, App.Path & "\Backup.bmp"
cleared = True
Pic.Picture1.Picture = LoadPicture()
End Sub

Private Sub clkd_Click()
CommonDialog1.ShowColor
Scalecolor = CommonDialog1.color
End Sub

Private Sub clwg_Click()
CommonDialog1.ShowColor
Webcolor = CommonDialog1.color
End Sub

Private Sub clzbz_Click()
CommonDialog1.ShowColor
xoycolor = CommonDialog1.color
End Sub

Public Sub Command1_Click()
On Error Resume Next
If Text1.text <> "" And cleared = False And Running = 0 Then SavePicture Picture1.Image, App.Path & "\Backup.bmp"
cleared = True
drawed = 0
Call ResetInc


Text1.SetFocus
Picture1.Cls
If Val(linewide.text) <= 0 Then linewide.text = "1"
Picture1.DrawWidth = Val(linewide.text)
xmin = Val(xmin.text)
xmax = Val(xmax.text)
ymin = Val(ymin.text)
ymax = Val(ymax.text)

Picture1.ScaleTop = ymax ' 为网格的顶部设置刻度。
    Picture1.ScaleLeft = xmin ' 为网格的左部设置刻度。
    Picture1.ScaleWidth = Abs(xmin - xmax) ' 设置刻度范围 （-1 到1）。
    Picture1.ScaleHeight = -Abs(ymax - ymin)

If showweb1.Checked = True Then

If polar1.Checked = False Then
If Xinc > 0 Then
For wang = 0 To xmin Step -Xinc
   Picture1.Line (wang, ymin)-(wang, ymax), Webcolor
Next wang
For wang = 0 To xmax Step Xinc
   Picture1.Line (wang, ymin)-(wang, ymax), Webcolor
Next wang
End If
If Yinc > 0 Then
For wang = 0 To ymin Step -Yinc
   Picture1.Line (xmin, wang)-(xmax, wang), Webcolor
Next wang
For wang = 0 To ymax Step Yinc
   Picture1.Line (xmin, wang)-(xmax, wang), Webcolor
Next wang
End If

Else

radius = 0
If Abs(xmin) > radius Then radius = Abs(xmin)
If Abs(xmax) > radius Then radius = Abs(xmax)
If Abs(ymin) > radius Then radius = Abs(ymin)
If Abs(ymax) > radius Then radius = Abs(ymax)
For ras = 0 To radius Step Xinc
  Picture1.Circle (0, 0), ras, Webcolor
Next ras
End If
End If

 If xoy1.Checked = True Then
 Picture1.Line (xmin, 0)-(xmax, 0), xoycolor ' 画水平线。
 Picture1.Line (0, ymin)-(0, ymax), xoycolor
 End If


If Showscale.Checked = True Then
colorbak = Picture1.ForeColor
Picture1.ForeColor = Scalecolor
If polar1.Checked = False Then
If Xinc > 0 Then
For wang = -Xinc To xmin Step -Xinc
   Picture1.CurrentX = wang
   Picture1.CurrentY = 0
   Picture1.Print Trim(Str(wang))
Next wang
For wang = Xinc To xmax Step Xinc
   Picture1.CurrentX = wang
   Picture1.CurrentY = 0
   Picture1.Print Trim(Str(wang))
Next wang
End If

If Yinc > 0 Then
For wang = -Yinc To ymin Step -Yinc
   Picture1.CurrentX = 0
   Picture1.CurrentY = wang
   Picture1.Print Trim(Str(wang))
Next wang
For wang = 0 To ymax Step Yinc
   Picture1.CurrentX = 0
   Picture1.CurrentY = wang
   Picture1.Print Str(wang)
Next wang
End If

Else

radius = 0
If Abs(xmin) > radius Then radius = Abs(xmin)
If Abs(xmax) > radius Then radius = Abs(xmax)
If Abs(ymin) > radius Then radius = Abs(ymin)
If Abs(ymax) > radius Then radius = Abs(ymax)
For ras = 0 To radius Step Xinc
  Picture1.CurrentX = ras
  Picture1.CurrentY = 0
  Picture1.Print Trim(Str(ras))
Next ras
End If
Picture1.ForeColor = colorbak
End If

End Sub

Private Sub Command2_Click()
'If Val(xmin.Text) <= 0 Then xmin.Text = 10
If Val(xmin.text) = Val(xmax.text) Then
  If Val(xmax) < 0 And Val(xmin) < 0 Then xmax.text = Str(-Val(xmax.text))
  If Val(xmax) > 0 And Val(xmin) > 0 Then xmin.text = Str(-Val(xmin.text))
End If
If Val(ymin.text) = Val(ymax.text) Then
  If Val(ymax) < 0 And Val(ymin) < 0 Then ymax.text = Str(-Val(ymax.text))
  If Val(ymax) > 0 And Val(ymin) > 0 Then ymin.text = Str(-Val(ymin.text))
End If

If Val(xmin.text) > Val(xmax.text) Then jh = Val(xmax.text): xmax.text = xmin.text: xmin.text = Str(jh)
If Val(ymin.text) > Val(ymax.text) Then jh = Val(ymax.text): ymax.text = ymin.text: ymin.text = Str(jh)

If Val(Combo2.text) <= 0 Then Combo2.text = "100"
Text1.text = ""
Text1.SetFocus
'Text8.Text = ""
'Text9.Text = ""
End Sub
'Private Sub command3_KeyDown(keycode As Integer, shift As Integer)
'If keycode = vbKeyF1 Then Text1.Text = 4
'End Sub






Private Sub CoverGraph_Click()
If CoverGraph.Checked = False Then CoverGraph.Checked = True Else CoverGraph.Checked = False
End Sub

Private Sub cshsh_Click()
If cshsh.Checked = False Then
  ImplicitFun.Checked = False
  ExplicitFun.Checked = False
  cshsh.Checked = True
End If
prmtfct.Show
prmtfct.WindowState = 0
End Sub

Private Sub dashed1_Click()
If dashed1.Checked = False Then
  dashed1.Checked = True
  Picture1.DrawStyle = 2
Else
dashed1.Checked = False
Picture1.DrawStyle = 0
End If

End Sub

Private Sub ddyl_Click()
If ddyl.Checked = False Then
  ddyl.Checked = True
  dxyl.Checked = False
  dwyy.Checked = False
End If
End Sub



Private Sub dline_Click()
If dline.Checked = False Then
  dline.Checked = True
  dpset.Checked = False
End If
If dline.Checked = True And (Val(Combo2.text) > 40 Or Val(Combo2.text) = 0) Then Combo2.text = 50
If dpset.Checked = True Then Combo2.text = "500"
End Sub

Public Sub dnc_Click()
If dnc.Checked = False Then
  dnc.Checked = True
  dpm.Checked = False
End If
End Sub

Private Sub dpm_Click()
If dpm.Checked = False Then
  dpm.Checked = True
  dnc.Checked = False
End If
End Sub

Private Sub dpset_Click()
If dpset.Checked = False Then
  dpset.Checked = True
  dline.Checked = False
End If
If dline.Checked = True And (Val(Combo2.text) > 40 Or Val(Combo2.text) = 0) Then Combo2.text = 50
If dpset.Checked = True Then Combo2.text = "500"
End Sub



Private Sub drvfun_Click()
If drvfun.Checked = False Then
  If polar1.Checked = False Then drvfun.Checked = True Else msg = MsgBox("抱歉, 暂不支持极坐标系下绘制导函数的图形。", vbInformation, "限制")
  If Val(Combo2.text) < 200 Then Combo2.text = "200"

Else
drvfun.Checked = False
End If
End Sub

Private Sub dwyy_Click()
If dwyy.Checked = False Then
  ddyl.Checked = False
  dxyl.Checked = False
  dwyy.Checked = True
End If
End Sub

Private Sub dxyl_Click()
If dxyl.Checked = False Then
  dxyl.Checked = True
  ddyl.Checked = False
  dwyy.Checked = False
End If
End Sub

Public Sub ExplicitFun_Click()
If ExplicitFun.Checked = False Then
  ExplicitFun.Checked = True
  Text1.Left = 104
  ImplicitFun.Checked = False
  cshsh.Checked = False
  If polar1.Checked = True Then Label1.Caption = "ρ =" Else Label1.Caption = "y ="
  lingdian.Enabled = True
  ifxtheny.Enabled = True
  dline.Enabled = True
  dpset.Enabled = True
  Pic.Caption = "方程曲线查看器  [显式 y=f(x)]"
  precision.Enabled = False
  precision3.Enabled = False
End If
End Sub

Private Sub fclb_Click()
If fclb.Checked = False Then
  FctList.Show
  fclb.Checked = True
Else
  FctList.Hide
  fclb.Checked = False
End If
End Sub

Private Sub fhhs_Click()
Dim Inputfct(0 To 100) As String
Dim i As Integer
i = 1

Do
  Inputfct(i) = LCase(InputBox("    请您按照从内层到外层的顺序输入各层显函数的解析式。自变量只能是 x。按 Esc 终止输入。本向导将把生成的复合函数置入曲线方程输入区。" & _
  Chr(10) & Chr(13) & Chr(10) & Chr(13) & "    现在,请输入第" & i & "层函数的解析式:" & Chr(10) & Chr(13) & "y=", "复合函数生成向导"))
  Do Until InStr(Inputfct(i), "fix") = 0
    Inputfct(i) = Left(Inputfct(i), InStr(Inputfct(i), "fix") - 1) + "fi?" + Right(Inputfct(i), Len(Inputfct(i)) - InStr(Inputfct(i), "fix") - 2)
  Loop
  Do Until InStr(Inputfct(i), "exp") = 0
    Inputfct(i) = Left(Inputfct(i), InStr(Inputfct(i), "exp") - 1) + "e?p" + Right(Inputfct(i), Len(Inputfct(i)) - InStr(Inputfct(i), "exp") - 2)
  Loop
  If i > 1 Then
    Do Until InStr(Inputfct(i - 1), "x") = 0
        Inputfct(i - 1) = Left(Inputfct(i - 1), InStr(Inputfct(i - 1), "x") - 1) + "v" + Right(Inputfct(i - 1), Len(Inputfct(i - 1)) - InStr(Inputfct(i - 1), "x"))
    Loop
    
    Do Until InStr(Inputfct(i), "x") = 0
      
       Inputfct(i) = Left(Inputfct(i), InStr(Inputfct(i), "x") - 1) + "[" + Inputfct(i - 1) + "]" + Right(Inputfct(i), Len(Inputfct(i)) - InStr(Inputfct(i), "x"))

    Loop
  End If
  i = i + 1
Loop Until Trim(Inputfct(i - 1)) = ""

Do Until InStr(Inputfct(i - 2), "v") = 0
        Inputfct(i - 2) = Left(Inputfct(i - 2), InStr(Inputfct(i - 2), "v") - 1) + "x" + Right(Inputfct(i - 2), Len(Inputfct(i - 2)) - InStr(Inputfct(i - 2), "v"))
Loop
Do Until InStr(Inputfct(i - 2), "?") = 0
        Inputfct(i - 2) = Left(Inputfct(i - 2), InStr(Inputfct(i - 2), "?") - 1) + "x" + Right(Inputfct(i - 2), Len(Inputfct(i - 2)) - InStr(Inputfct(i - 2), "?"))
Loop
Text1.text = Inputfct(i - 2)
End Sub

Private Sub Form_unLoad(Cancel As Integer)
If Running = 1 Then
Command3_Click
End If
End Sub
Private Sub Form_Resize()
  If Pic.WindowState = 2 Then
    Pic.WindowState = 0
    Me.Width = 12000
    Me.Height = 8700
  End If
  If Me.Width > 12000 Then Me.Width = 12000
  If Me.Height > 8700 Then Me.Height = 8700
End Sub
Private Sub Form_Load()
'Call Me.
'KeyPreview = True
prmt = False
t1e = 0
Frame5.Height = 0
Frame5.Width = 0
xoycolor = RGB(128, 128, 128)
Webcolor = RGB(128, 255, 255)
Scalecolor = 8388736

Picture1.Cls
 xmin = Val(xmin.text)
 xmax = Val(xmax.text)
 ymin = Val(ymin.text)
 ymax = Val(ymax.text)
 Xinc = Abs(Val(Xinc.text))
 Yinc = Abs(Val(Yinc.text))
 
If Xinc > Abs(xmax - xmin) Then Xinc = Abs(xmax - xmin)
If Yinc > Abs(ymax - ymin) Then Yinc = Abs(ymax - ymin)

If polar1.Checked = False Then
If Xinc > 0 Then
For wang = 0 To xmin Step -Xinc
   Picture1.Line (wang, ymin)-(wang, ymax), Webcolor
Next wang
For wang = 0 To xmax Step Xinc
   Picture1.Line (wang, ymin)-(wang, ymax), Webcolor
Next wang
End If
If Yinc > 0 Then
For wang = 0 To ymin Step -Yinc
   Picture1.Line (xmin, wang)-(xmax, wang), Webcolor
Next wang
For wang = 0 To ymax Step Yinc
   Picture1.Line (xmin, wang)-(xmax, wang), Webcolor
Next wang
End If
End If
If xoy1.Checked = True Then
 Picture1.Line (xmin, 0)-(xmax, 0), xoycolor ' 画水平线。
 Picture1.Line (0, ymin)-(0, ymax), xoycolor
End If
If Showscale.Checked = True Then
colorbak = Picture1.ForeColor
Picture1.ForeColor = Scalecolor
If polar1.Checked = False Then
If Xinc > 0 Then
For wang = -Xinc To xmin Step -Xinc
   Picture1.CurrentX = wang
   Picture1.CurrentY = 0
   Picture1.Print Trim(Str(wang))
Next wang
For wang = Xinc To xmax Step Xinc
   Picture1.CurrentX = wang
   Picture1.CurrentY = 0
   Picture1.Print Trim(Str(wang))
Next wang
End If

If Yinc > 0 Then
For wang = -Yinc To ymin Step -Yinc
   Picture1.CurrentX = 0
   Picture1.CurrentY = wang
   Picture1.Print Trim(Str(wang))
Next wang
For wang = 0 To ymax Step Yinc
   Picture1.CurrentX = 0
   Picture1.CurrentY = wang
   Picture1.Print Str(wang)
Next wang
End If

Else

radius = 0
If Abs(xmin) > radius Then radius = Abs(xmin)
If Abs(xmax) > radius Then radius = Abs(xmax)
If Abs(ymin) > radius Then radius = Abs(ymin)
If Abs(ymax) > radius Then radius = Abs(ymax)
For ras = 0 To radius Step Xinc
  Picture1.CurrentX = ras
  Picture1.CurrentY = 0
  Picture1.Print Trim(Str(ras))
Next ras
End If
Picture1.ForeColor = colorbak
End If
End Sub
Private Sub from_Unload(Cancel As Integer)

End Sub


Public Sub Frame5_dblclick()
xmin.text = "-10"
xmax.text = "10"
ymin.text = "-10"
ymax.text = "10"
Xinc.text = "1"
Yinc.text = "1"
End Sub
Private Sub Frame5_mousemove(Button As Integer, Shift As Integer, x As Single, y As Single)
xmin = Fc(Bracket(translate(xmin.text)), 0, 0)
xmax = Fc(Bracket(translate(xmax.text)), 0, 0)
ymin = Fc(Bracket(translate(ymin.text)), 0, 0)
ymax = Fc(Bracket(translate(ymax.text)), 0, 0)
Xinc = Fc(Bracket(translate(Xinc.text)), 0, 0)
Yinc = Fc(Bracket(translate(Yinc.text)), 0, 0)
If Abs(Abs(xmax - xmin) - Abs(ymax - ymin)) > 0.00001 Then xyalarm.Visible = True Else xyalarm.Visible = False
Call ResetInc
End Sub

Private Sub frame6_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If Text8.text <> "" Then Text8.text = Fc(Bracket(translate(Text8.text)), 0, 0)
If Text9.text <> "" Then Text9.text = Fc(Bracket(translate(Text9.text)), 0, 0)
End Sub

Private Sub ResetInc()
If Val(xmin.text) = Val(xmax.text) Then
  If Val(xmax) < 0 And Val(xmin) < 0 Then xmax.text = Str(-Val(xmax.text))
  If Val(xmax) > 0 And Val(xmin) > 0 Then xmin.text = Str(-Val(xmin.text))
End If
If Val(ymin.text) = Val(ymax.text) Then
  If Val(ymax) < 0 And Val(ymin) < 0 Then ymax.text = Str(-Val(ymax.text))
  If Val(ymax) > 0 And Val(ymin) > 0 Then ymin.text = Str(-Val(ymin.text))
End If

If Val(xmin.text) > Val(xmax.text) Then jh = Val(xmax.text): xmax.text = xmin.text: xmin.text = Str(jh)
If Val(ymin.text) > Val(ymax.text) Then jh = Val(ymax.text): ymax.text = ymin.text: ymin.text = Str(jh)

If Val(Combo2.text) < 0 Or Combo2.text = "0" Then Combo2.text = "100"

If Val(Xinc.text) < Abs(Val(xmax.text) - Val(xmin.text)) / 50 Then
  Xinc.text = Abs(Val(xmax.text) - Val(xmin.text)) / 50
End If
If Val(Yinc.text) < Abs(Val(ymax.text) - Val(ymin.text)) / 50 Then
  Yinc.text = Abs(Val(ymax.text) - Val(ymin.text)) / 50
End If
  
End Sub




Private Sub gray_Click()
Picture1.BackColor = &HD8E9EC
Picture1.ForeColor = &HA56E3A
End Sub



Private Sub hlg_Click(Index As Integer)
Text1.text = hlg(Index).Caption
End Sub

Private Sub ifxtheny_Change()

On Error GoTo com:
cp$ = ifxtheny.text
aoo$ = Bracket(translate(cp$))

cpt# = Fc(aoo$, 0, 0)


aoo$ = Bracket(translate(Text1.text))

b$ = LCase(Replace(aoo$, "exp", "ep"))
b$ = Replace(b$, "x", "(v)")
b$ = Replace(b$, "e^", "exp")
b$ = translate(b$)
b$ = Replace(b$, "(v)", "x")
a$ = ExpChk_d(b$)
If a$ <> "" Then a$ = "" Else a$ = CleanUpExrp(d_fx(b$))
a$ = Replace(a$, "ep", "exp")
aod$ = a$

aod$ = Replace(aod$, "x", "(V)")

Do Until InStr(aoo$, "x") = 0
   aoo$ = Left(aoo$, InStr(aoo$, "x") - 1) + "(V)" + Right(aoo$, Len(aoo$) - InStr(aoo$, "x"))
Loop

aoo$ = Bracket(translate(aoo$))

Result.text = Fc(aoo$, cpt#, 0)


compu = 1

ResultD.text = Fc(aod$, cpt#, 0)
If InStr(aod$, "'") > 0 Or aod$ = "" Then ResultD.text = ""
com:
If compu = 0 Then
  Result.text = "Null"
  ResultD.text = "Null"
  Resume comput
End If
comput: compu = 0

End Sub



Public Sub ImplicitFun_Click()
If ImplicitFun.Checked = False Then
  ImplicitFun.Checked = True
  ExplicitFun.Checked = False
  cshsh.Checked = False
  Text1.Left = 48 'Label1.Caption = "F(x,y)="
  lingdian.Enabled = False
  ifxtheny.Enabled = False
  dline.Enabled = False
  dpset.Enabled = False
  Pic.Caption = "方程曲线查看器  [隐式 F(x,y)=0]"
  drvfun.Enabled = False
  precision.Enabled = True
  precision3.Enabled = True
End If
End Sub

Private Sub inputt_Click()
  Sendkeys "(t)"
End Sub

Private Sub Label1_DblClick()
  If ImplicitFun.Checked = False Then Call ImplicitFun_Click
End Sub
Private Sub Label17_DblClick()
'ExplicitFun
  If ExplicitFun.Checked = False Then Call ExplicitFun_Click

End Sub
Private Sub Label10_Click()
ymin.text = Str(Val(xmin.text))
End Sub

Private Sub Label11_Click()
ymax.text = Str(Val(xmax.text))
End Sub

Private Sub Label12_Click()
Xinc.text = Str(Val(Yinc.text))
End Sub

Private Sub Label13_Click()
Yinc.text = Str(Val(Xinc.text))
End Sub

Private Sub Label2_mousemove(Button As Integer, Shift As Integer, x As Single, y As Single)
Frame5.Visible = True
'Timer2.Enabled = False
If t1e = 0 Then Timer1.Enabled = True
drawed = 0
End Sub

Private Sub Label3_Click()
xmin.text = Str(-Val(xmax.text))
End Sub



Private Sub Label9_Click()
xmax.text = Str(-Val(xmin.text))
End Sub





Private Sub lang_Click()
Dim a(), x(), y() As Double
  Dim m As Double
  Dim k, i, j As Integer
  Dim l, p As String
  Dim iptX, iptY As String
  On Error GoTo l1:
  n = 0
  Picture1.DrawWidth = Val(linewide.text) + 2

 Do
l2:
    iptX = InputBox("x(" & n & ")=", "拉格朗日插值多项式生成向导")
    If iptX = "" Then
      iptY = ""
    Else
      If iptX <> "b" Then iptY = InputBox("y(" & n & ")=", "拉格朗日插值多项式生成向导")
    End If
    
    If iptX = "b" Or iptY = "b" Then
      colorbak = Picture1.ForeColor

      Picture1.ForeColor = Picture1.BackColor
      
      

      n = n - 1
      Picture1.PSet (x(n), y(n))
      Picture1.ForeColor = colorbak
      GoTo l2:
    End If
    ReDim Preserve x(n + 1)
    ReDim Preserve y(n + 1)
    ReDim Preserve a(n + 1)
    If iptX <> "" Then
      Picture1.CurrentX = x(k)
      Picture1.CurrentY = y(k)
   'Picture1.Print "(" & x(k) & "," & y(k) & ")"
      Picture1.PSet (iptX, iptY)
    End If
    
    x(n) = Val(iptX)
    y(n) = Val(iptY)
    n = n + 1
  Loop Until iptX = "" Or iptY = ""
    
    n = n - 2
    If n < 1 Then Exit Sub
    
  For i = 0 To n
      
      
      m = 1
      For k = 0 To n
        If i <> k Then m = m * (x(i) - x(k))
      Next k
     
      
      a(i) = y(i) / m
           
      l = ""
      For j = 0 To n
        If j <> i Then l = l & "(x" & AddNum(-x(j)) & ")"
      Next j
      
      p = p & AddNum(a(i)) & l
         
           
         
  Next i
    
 If Left(p, 1) = "+" Then p = Right(p, Len(p) - 1)
    
    
 Text1.text = p
 
 
  Picture1.DrawWidth = Val(linewide.text)

 Exit Sub
l1: msg = MsgBox("无法生成拉格朗日插值多项式。", vbExclamation, "n阶拉格朗日插值")
End Sub

Private Sub lingdian_Click()
Fct.Show
Fct.Text1.text = Text1.text
Fct.Lef.text = xmin.text
Fct.Rig.text = xmax.text
Sendkeys "{enter}"
'Do Until InStr(Fct.ifm.Caption, "%") <> 0 And Fct.ifm.Caption <> ""
'Realzero.Text = Fct.Value.Text
'Loop
End Sub

Private Sub loadpic_Click()
On Error Resume Next
CommonDialog1.ShowOpen
Picture1.Picture = LoadPicture(CommonDialog1.FileName)
End Sub




Private Sub lw_Click()
linewide.text = Str(Fix(Abs(InputBox("请输入曲线的宽度", "提示"))))
End Sub

Private Sub Magnififier_Click()
frmMagnifier.Show
End Sub

Private Sub mainmenu_click()
Pic.PopupMenu popup
End Sub

Private Sub mousexoy1_Click()
If mousexoy1.Checked = False Then
  mousexoy1.Checked = True
  mousex.Visible = True: mousey.Visible = True
Else
mousexoy1.Checked = False
 mousex.Visible = False: mousey.Visible = False
End If

End Sub



Private Sub o_Click(Index As Integer)
Text1.text = Text1.text + o(Index).Caption
Text1.SetFocus
End Sub

Private Sub op_Click(Index As Integer)
Text1.text = Text1.text + op(Index).Caption
Text1.SetFocus
End Sub

Private Sub ope_Click(Index As Integer)
Text1.text = Text1.text + ope(Index).Caption
Text1.SetFocus
End Sub

Private Sub open_Click()
On Error Resume Next
CommonDialog1.ShowOpen
Picture1.Picture = LoadPicture(CommonDialog1.FileName)
End Sub


Private Sub Option1_Click()
Picture1.BackColor = &HFFFFFF
Picture1.ForeColor = &H0&
End Sub

Private Sub Option2_Click()
Picture1.BackColor = &H0&
Picture1.ForeColor = &HFFFFFF
End Sub

Public Sub Command3_Click()
Dim xtime As Double
Dim Xpix As Long
Dim Ypix As Long
Dim laXpix As Long
Dim laYpix As Long
Dim Pointapi1 As PointAPI
Const PI = 3.14159265258979
On Error Resume Next

If Right(Text1.text, 1) = ";" Then
  Text1.text = Left(Text1.text, Len(Text1.text) - 1)
  Call Command1_Click
End If

xtime = 0
th = True



If ExplicitFun.Checked = True Then
  If InStr(Text1.text, "y") > 0 Then
      msg$ = MsgBox("显函数模式下不支持变量y.", , "错误")
      GoTo unnext:
  End If
End If

If InStr(Text1.text, ")62") > 0 Then
    msg$ = MsgBox("计算器怀疑您把 ^ 误输入为 6 。" & Chr(13) & Chr(13) & "您要把 62 替换成 ^2 吗?", vbYesNo + vbQuestion + vbDefaultButton1, "可能的输入错误")
    If msg = 6 Then
      Do Until InStr(Text1.text, ")62") = 0
           Text1.text = Left(Text1.text, InStr(Text1.text, ")62") - 1) + ")^2" + Right(Text1.text, Len(Text1.text) - InStr(Text1.text, ")62") - 2)
      Loop
    End If
End If

If Val(precision3.text) <= 1 Then precision3.text = "10"
If CoverGraph.Checked = True Then Call Command1_Click


If prmt = False Then
  errmsg = ExpChk(Text1.text)
  If errmsg <> "" Then
    msg$ = MsgBox(errmsg & Chr(13) & Chr(13) & "您要中止计算吗?", vbYesNo + vbQuestion + vbDefaultButton1, "错误")
    If msg = 6 Then GoTo unnext:
  End If
End If

package = True
hlg(0).Visible = True

If LCase(Text1.text) <> hlg(hlog - 1).Caption Then
  hlog = hlog + 1
  Load hlg(hlog)
  hlg(hlog).Caption = LCase(Text1.text)
End If
If hlog > 30 Then Unload hlg(hlog - 30)

 Text1h.Enabled = False: Text1.Height = 27
Text1.SelStart = Len(Text1.text): Text1.SetFocus
If Running = 1 Then Running = 0 Else Running = 1
passline = 0
If Val(Combo2.text) <= 0 Then Combo2.text = "100"
Picture1.MousePointer = 11
gn = 0

If cleared = False Then
  If Running = 1 Then SavePicture Picture1.Image, App.Path & "\Backup.bmp"
End If

cleared = False
On Error GoTo 168

Call ResetInc

ProgressBar1.Value = 0
If Val(linewide.text) <= 0 Then linewide.text = "1"
Picture1.DrawWidth = Val(linewide.text)
xmin = Val(xmin.text)
xmax = Val(xmax.text)
ymin = Val(ymin.text)
ymax = Val(ymax.text)


ste = Abs(xmax - xmin) / (5 * Val(Combo2.text))

aoo$ = Bracket(translate(Text1.text))

Picture1.ScaleTop = ymax ' 为网格的顶部设置刻度。
    Picture1.ScaleLeft = xmin ' 为网格的左部设置刻度。
    Picture1.ScaleWidth = Abs(xmax - xmin) ' 设置刻度范围。
    Picture1.ScaleHeight = -Abs(ymax - ymin)
If xoy1.Checked = True Then
 Picture1.Line (xmin, 0)-(xmax, 0), xoycolor ' 画水平线。
 Picture1.Line (0, ymin)-(0, ymax), xoycolor
End If




If dline.Checked = True Then
 Picture1.CurrentX = -(Abs(xmin - xmax) / 2)
 Picture1.CurrentY = 0
End If
ms# = 0
  lasty = 1E+308

If prmt = False Then
  If Text8.text = "" Then
    If unfun1.Checked = True Then
      If polar1.Checked = False Then fleft = Val(ymin) Else fleft = 0
    Else
      If polar1.Checked = False Then fleft = Val(xmin) Else fleft = 0
    End If
  Else
    fleft = Val(Text8.text)
  End If
  If Text9.text = "" Then
    If unfun1.Checked = True Then
      If polar1.Checked = False Then tright = Val(ymax) Else tright = 2 * PI
    Else
      If polar1.Checked = False Then tright = Val(xmax) Else tright = 2 * PI
    End If
  Else
    tright = Val(Text9.text)
  End If
  If unfun1.Checked = False Then
    If fleft < Val(xmin) Then fleft = Val(xmin)
    If tright > Val(xmax) Then tright = Val(xmax)
  End If
  sw = ste / Abs(fleft - tright)
  
Else
  If prmtfct.tl.text = "" Then fleft = xmin * 2 Else fleft = Val(prmtfct.tl.text)
  If prmtfct.tr.text = "" Then tright = xmax * 2 Else tright = Val(prmtfct.tr.text)
  'sw = 1 / (ste * Abs(xmin - xmax))
  sw = ste / Abs(xmin - xmax)
End If




If showweb1.Checked = True Then

   xmin = Val(xmin.text)
   xmax = Val(xmax.text)
   ymin = Val(ymin.text)
   ymax = Val(ymax.text)
   Xinc = Abs(Val(Xinc.text))
   Yinc = Abs(Val(Yinc.text))
 
  If Xinc > Abs(xmax - xmin) Then Xinc = Abs(xmax - xmin)
  If Yinc > Abs(ymax - ymin) Then Yinc = Abs(ymax - ymin)

  If drawed = 0 Then
    If polar1.Checked = False Then
      If Xinc > 0 Then
        For wang = 0 To xmin Step -Xinc
           Picture1.Line (wang, ymin)-(wang, ymax), Webcolor
        Next wang
        For wang = 0 To xmax Step Xinc
           Picture1.Line (wang, ymin)-(wang, ymax), Webcolor
        Next wang
      End If
      If Yinc > 0 Then
        For wang = 0 To ymin Step -Yinc
          Picture1.Line (xmin, wang)-(xmax, wang), Webcolor
        Next wang
        For wang = 0 To ymax Step Yinc
           Picture1.Line (xmin, wang)-(xmax, wang), Webcolor
        Next wang
      End If
    Else
      radius = 0
      If Abs(xmin) > radius Then radius = Abs(xmin)
      If Abs(xmax) > radius Then radius = Abs(xmax)
      If Abs(ymin) > radius Then radius = Abs(ymin)
      If Abs(ymax) > radius Then radius = Abs(ymax)
      For ras = 0 To radius Step Xinc
        Picture1.Circle (0, 0), ras, Webcolor
      Next ras
    End If
    If xoy1.Checked = True Then
      If drawed = 0 Then
       Picture1.Line (xmin, 0)-(xmax, 0), xoycolor ' 画水平线。
       Picture1.Line (0, ymin)-(0, ymax), xoycolor
      End If
   End If
  End If
End If

If Showscale.Checked = True Then
colorbak = Picture1.ForeColor
Picture1.ForeColor = Scalecolor
If polar1.Checked = False Then
If Xinc > 0 Then
For wang = -Xinc To xmin Step -Xinc
   Picture1.CurrentX = wang
   Picture1.CurrentY = 0
   Picture1.Print Trim(Str(wang))
Next wang
For wang = Xinc To xmax Step Xinc
   Picture1.CurrentX = wang
   Picture1.CurrentY = 0
   Picture1.Print Trim(Str(wang))
Next wang
End If

If Yinc > 0 Then
For wang = -Yinc To ymin Step -Yinc
   Picture1.CurrentX = 0
   Picture1.CurrentY = wang
   Picture1.Print Trim(Str(wang))
Next wang
For wang = 0 To ymax Step Yinc
   Picture1.CurrentX = 0
   Picture1.CurrentY = wang
   Picture1.Print Str(wang)
Next wang
End If

Else

radius = 0
If Abs(xmin) > radius Then radius = Abs(xmin)
If Abs(xmax) > radius Then radius = Abs(xmax)
If Abs(ymin) > radius Then radius = Abs(ymin)
If Abs(ymax) > radius Then radius = Abs(ymax)
For ras = 0 To radius Step Xinc
  Picture1.CurrentX = ras
  Picture1.CurrentY = 0
  Picture1.Print Trim(Str(ras))
Next ras
End If
Picture1.ForeColor = colorbak
End If
If aoo$ = "" Then GoTo 168


If ExplicitFun.Checked = True Then
ymaxa = -1.79769313486231E+308
ymina = 1.79769313486231E+308
Pic.Caption = "方程曲线查看器"

If prmt = False Then

aop$ = aoo$

If InStr(aop$, "=") > 0 Then
  aop$ = Left(aop$, InStr(aop$, "=") - 1) + "-(" + _
  Right(aop$, Len(aop$) - InStr(aop$, "=")) + ")"
End If

Do Until InStr(aop$, "x") = 0
   aop$ = Left(aop$, InStr(aop$, "x") - 1) + "(V)" + Right(aop$, Len(aop$) - InStr(aop$, "x"))
Loop
aoo$ = Bracket(translate(aop$))
aop$ = aoo$

Else

aop1$ = Bracket(translate(bracketT(prmtfct.xpa.text)))
  
errmsg = ExpChk(aop1$)
If errmsg <> "" Then
  msg$ = MsgBox(errmsg & Chr(13) & Chr(13) & "您要中止计算吗?", vbYesNo + vbQuestion + vbDefaultButton1, "错误")
  If msg = 6 Then GoTo unnext
End If
  
  Do Until InStr(aop1$, "(t)") = 0
   aop1$ = Left(aop1$, InStr(aop1$, "(t)") - 1) + "(V)" + Right(aop1$, Len(aop1$) - InStr(aop1$, "(t)") - 2)
  Loop
  aoo$ = Bracket(translate(aop1$))
  aop1$ = aoo$

  aop2$ = Bracket(translate(bracketT(prmtfct.ypa.text)))
  errmsg = ExpChk(aop2$)
  If errmsg <> "" Then
    msg$ = MsgBox(errmsg & Chr(13) & Chr(13) & "您要中止计算吗?", vbYesNo + vbQuestion + vbDefaultButton1, "错误")
    If msg = 6 Then GoTo unnext
  End If
  
  Do Until InStr(aop2$, "(t)") = 0
   aop2$ = Left(aop2$, InStr(aop2$, "(t)") - 1) + "(V)" + Right(aop2$, Len(aop2$) - InStr(aop2$, "(t)") - 2)
  Loop
  aoo$ = Bracket(translate(aop2$))
  aop2$ = aoo$

End If

timerb = Timer
Picture1.ForeColor = Pfc

For xa = fleft To tright Step ste

  If prmt = False Then
    x = xa
    
  If package = False Then y = Fc(aoo$, x, 0) Else Call pack
  End If

  If prmt = True Then
  aoo$ = aop2$
  x = xa
 
  y = Fc(aoo$, x, 0)
  aoo$ = aop1$
  x = xa
 
  x = Fc(aoo$, x, 0)
  
  
  End If
  
'____________________________________________________
If dnc.Checked = True Then
    Xpix = Pixel(500, Val(xmin.text), Val(xmax.text), x, False)
    Ypix = Pixel(500, Val(ymin.text), Val(ymax.text), y, True)
End If
'____________________________________________________
je:

If erro = 0 And unfun1.Checked = False Then
 
 
 If dpset.Checked = True Then
  
  If polar1.Checked = False Then Picture1.PSet (x, y) Else Picture1.PSet (y * Cos(x), y * Sin(x))
  
  If drvfun.Checked = True Then
    If polar1.Checked = False Then
      Picture1.ForeColor = QBColor(12)
      Picture1.PSet (x, (y - lay) / (x - lax))
    End If
  End If
 
  Picture1.ForeColor = Pfc
 
 End If
 
 
 allow = 1
 
 If cshsh.Checked = True Then
   If Abs(lax - x) > Abs(Val(xmin) - Val(xmax)) Then allow = 0
 End If
 
 'If Abs(lay - y) > Abs(Val(ymin.Text) - Val(ymax.Text)) Then allow = 0
 If Abs(lay - y) > 10 * Abs(llay - lay) Then allow = 0
 
 If (lay >= Val(ymax) And y <= Val(ymin)) Or (lay <= Val(ymin) And y >= Val(ymax)) Then allow = 0
 
 
 If dline.Checked = True Then
 If lasterr = 0 Then
 If passline = 1 Then
 If allow = 1 Then
   If polar1.Checked = False Then
     Call Line0(lax, lay, x, y, laXpix, laYpix, Xpix, Ypix)
       
     
     xtime = xtime + 1
   Else
     Picture1.Line (lay * Cos(lax), lay * Sin(lax))-(y * Cos(x), y * Sin(x))
   End If
   
   If drvfun.Checked = True Then
     If polar1.Checked = False Then
       If xtime > 2 Then
        If Abs(lay - y) < Abs(ymax - ymin) Then
          Picture1.ForeColor = QBColor(12)
          Picture1.Line (lax, (lay - llay) / (lax - llax))-(x, (y - lay) / (x - lax))
        End If
       End If
     End If
  End If
 
 
  Picture1.ForeColor = Pfc
 End If
 End If
 End If
 End If
 
 If lasterr = 1 Then
   CurrentX = x: CurrentY = y: mt = MoveToEx(Picture1.hdc, laXpix, laYpix, Pointapi1)
 End If
 
 llax = lax: llay = lay ''''''
 lax = x: lay = y
 laXpix = Xpix: laYpix = Ypix

 passline = 1
 If y > ymaxa Then ymaxa = y
 If y < ymina Then ymina = y

End If

If erro = 0 And unfun1.Checked = True Then
 
 If dpset.Checked = True Then
 If polar1.Checked = False Then Picture1.PSet (y, x) Else Picture1.PSet (x * Cos(y), x * Sin(y)) '(y * Cos(x), y * Sin(x))
  
  If polar1.Checked = False Then
    If drvfun.Checked = True Then
      Picture1.ForeColor = QBColor(12)
      Picture1.PSet (y, (x - lax) / (y - lay))
    End If
  End If
 
 Picture1.ForeColor = Pfc
 End If
 
 allow = 1
 
 If cshsh.Checked = True Then
   If Abs(lax - x) > Abs(Val(xmin) - Val(xmax)) Then allow = 0
 End If
 
 If Abs(lay - y) > 10 * Abs(llay - lay) Then allow = 0
 
 If (lay >= Val(ymax) And y <= Val(ymin)) Or (lay <= Val(ymin) And y >= Val(ymax)) Then allow = 0
 
 If dline.Checked = True Then
 If lasterr = 0 Then
 If passline = 1 Then
 If allow = 1 Then
   If polar1.Checked = False Then
     Picture1.Line (lay, lax)-(y, x)
     'mt = MoveToEx(Picture1, laypix, laxpix, Pointapi1)
     'lt = LineTo(Picture1, ypix, xpix)
     xtime = xtime + 1
   Else
     Picture1.Line (lax * Cos(lay), lax * Sin(lay))-(x * Cos(y), x * Sin(y))
   End If
 
   
   If polar1.Checked = False Then
     If drvfun.Checked = True Then
      If xtime > 2 Then
        If Abs(lay - y) < Abs(ymax - ymin) Then
          Picture1.ForeColor = QBColor(12)
          Picture1.Line ((lay - llay) / (lax - llax), lax)-((y - lay) / (x - lax), x)
        End If
      End If
    End If
   End If
 Picture1.ForeColor = Pfc
 End If
 End If
 End If
 End If
 
 'If dline.Checked = True And lasterr = 1 Then CurrentX = y: CurrentY = x
 If lasterr = 1 Then CurrentX = y: CurrentY = x
 
 llax = lax: llay = lay ''''''
 lax = x: lay = y
 laXpix = Xpix: laYpix = Ypix
 passline = 1
 If y > ymaxa Then ymaxa = y
 If y < ymina Then ymina = y
End If


If erro = 1 Then lasterr = 1 Else lasterr = 0
erro = 0

jump:  aoo$ = aop$
If st <= 1 Then ProgressBar1.Value = st * 1000
st = st + sw

If Fix(ProgressBar1.Value) = 10 Then
  If Text8.text = "" And Text9.text = "" And Timer - timerb >= 0.3 And gn = 0 Then
    msg$ = "您所要求的操作大约要花" & Str(Fix((Timer - timerb) * 100)) & "秒的时间才能完成．" & Chr(13) & Chr(13) & "您要继续计算吗?"
    Style = vbYesNo + vbQuestion + vbDefaultButton1
    goon = MsgBox(msg$, Style, "Graph")
    If goon = 7 Then GoTo 168 Else gn = 1
  End If
End If

DoEvents

If Running = 0 Then
 Pic.Caption = "方程曲线查看器  [已停止]"
 GoTo unnext
End If

Next xa
prmt = False
End If

If ExplicitFun.Checked = False Then
Pic.Caption = "方程曲线查看器  [绘制中...]"
aop$ = Text1.text
If InStr(aop$, "=") > 0 Then
  aop$ = Left(aop$, InStr(aop$, "=") - 1) + "-(" + _
  Right(aop$, Len(aop$) - InStr(aop$, "=")) + ")"
End If

Do Until InStr(aop$, "y") = 0
   aop$ = Left(aop$, InStr(aop$, "y") - 1) + "(V)" + Right(aop$, Len(aop$) - InStr(aop$, "y"))
Loop
Do Until InStr(aop$, "x") = 0
   aop$ = Left(aop$, InStr(aop$, "x") - 1) + "(W)" + Right(aop$, Len(aop$) - InStr(aop$, "x"))
Loop
aop$ = Bracket(translate(aop$))
ated = aop$



timerb = Timer
For yh = fleft To tright Step ste

Call PicPset(1, xmin, xmax)
DoEvents
If Running = 0 Then GoTo unnext
If st <= 1 Then ProgressBar1.Value = st * 1000
st = st + sw
Next yh

Pic.Caption = "方程曲线查看器  [即将绘制完成...]"
st = 0: ProgressBar1.Value = 0
fleft = ymin
tright = ymax


aop$ = Text1.text
If InStr(aop$, "=") > 0 Then
  aop$ = Left(aop$, InStr(aop$, "=") - 1) + "-(" + _
  Right(aop$, Len(aop$) - InStr(aop$, "=")) + ")"
End If
Do Until InStr(aop$, "y") = 0
   aop$ = Left(aop$, InStr(aop$, "y") - 1) + "(W)" + Right(aop$, Len(aop$) - InStr(aop$, "y"))
Loop
Do Until InStr(aop$, "x") = 0
   aop$ = Left(aop$, InStr(aop$, "x") - 1) + "(V)" + Right(aop$, Len(aop$) - InStr(aop$, "x"))
Loop
aop$ = Bracket(translate(aop$))
ated = aop$

timerb = Timer
For yh = fleft To tright Step ste

Call PicPset(0, xmin, xmax)
DoEvents
If Running = 0 Then GoTo unnext
If st <= 1 Then ProgressBar1.Value = st * 1000
st = st + sw
Next yh

Pic.Caption = "方程曲线查看器"

End If





168: If ImplicitFun.Checked = True Then Resume Next
If err <> 0 Then erro = 1: Resume je:
st = 0


Text5.text = ymaxa
Text6.text = ymina
If ymaxa > Val(ymax) Then Text5.text = Text5.text + " ?"
If ymina < Val(ymin) Then Text6.text = Text6.text + " ?"

If unfun1.Checked = True Then
  Text5.text = ""
  Text6.text = ""
End If
Pic.Caption = "方程曲线查看器 "
unnext: unok = 0
If ifxtheny.text <> "" Then ifxtheny_Change

Running = 0
Picture1.MousePointer = 2
If ProgressBar1.Value < 990 Then ProgressBar1.Value = 1000
prmt = False
drawed = 1
Text2.text = Timer - timerb
Picture1.Visible = False
Picture1.Visible = True
End Sub




Private Sub Option5_Click()
Picture1.BackColor = &HD8E9EC
Picture1.ForeColor = &HA56E3A
End Sub

Private Sub Option6_Click()
Picture1.BackColor = &H80000001
Picture1.ForeColor = &HFFFF&
End Sub



Private Sub other_Click()
Picdraw.Show: Picdraw.WindowState = 0
End Sub



Private Sub picture1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next
Timer1.Enabled = False
Text1h.Enabled = False: Text1.Height = 27
'Text1.SetFocus

If mousexoy1.Checked = True Then
  If polar1.Checked = False Then
    mousex.text = x
    mousey.text = y
  Else
    PI = Atn(1) * 4
    mousey.text = Sqr(x ^ 2 + y ^ 2)
    If x > 0 And y > 0 Then mousex.text = Atn(y / x) / PI & "π"
    If x < 0 Then mousex.text = 1 + Atn(y / x) / PI & "π"
    If x > 0 And y < 0 Then mousex.text = 2 + Atn(y / x) / PI & "π"
    If x > 0 And y = 0 Then mousex.text = "0"
    If x < 0 And y = 0 Then mousex.text = "π"
    If x = 0 And y > 0 Then mousex.text = "0.5π"
    If x = 0 And y < 0 Then mousex.text = "1.5π"
  End If

  If (Val(xmin.text) <> 0 Or Val(xmax.text) <> 0) And (Val(ymin.text) <> 0 Or Val(ymax.text) <> 0) And xmin.text <> xmax.text And ymin.text <> ymax.text Then
    mousex.Left = 500 * Abs(x - xmin) / Abs(xmax - xmin) + 15
    mousey.Top = 500 - 500 * (Abs(y - ymin) / Abs(ymax - ymin)) + 30
    mousey.Width = 6 * Len(mousey.text)
  End If

  If Button = 1 Then
    Linex.Visible = True
    Liney.Visible = True
    Linex.x1 = Val(xmin.text)
    Linex.x2 = Val(xmax.text)
    Linex.y1 = y
    Linex.y2 = y
    Liney.y1 = Val(ymin.text)
    Liney.y2 = Val(ymax.text)
    Liney.x1 = x
    Liney.x2 = x
    ifxtheny.text = x
  Else
    Linex.Visible = False
    Liney.Visible = False
  End If
 'If button = 2 And shift <> 4 Then Picture1.PSet (x, y)
 'If button = 2 And shift = 4 Then
 '  Picture1.PSet (x, y), Picture1.BackColor
 'End If
End If
Frame5.Visible = False

End Sub
'Private Sub picture1_click()
' Picture1.PSet (x, y)
'End Sub
Private Sub picture1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 1 Then
If Shift = 1 Then Picture1.Line -(x, y)
If Shift = 2 Then Picture1.CurrentX = x: Picture1.CurrentY = y: PSet (x, y)

If Shift = 3 Then
  Picture1.CurrentX = x: Picture1.CurrentY = y
  If ImplicitFun.Checked = True Then
    If InStr(Text1.text, "=") = 0 Then Picture1.Print Text1.text + "=0" Else Picture1.Print Text1.text
  Else
    If ExplicitFun.Checked = True And cshsh.Checked = False Then
      Picture1.Print "y=" + Text1.text
    Else
      If cshsh.Checked = True Then Picture1.Print "x=" + prmtfct.xpa.text + ", y=" + prmtfct.ypa.text
    End If
  End If
End If

If Shift = 4 Then
  i = Abs(Val(xmax.text) - Val(xmin.text)) / 5
  xmax.text = Str(x + i)
  xmin.text = Str(x - i)
  i = Abs(Val(ymax.text) - Val(ymin.text)) / 5
  ymax.text = Str(x + i)
  ymin.text = Str(x - i)
  
  Call DrawGrp
End If
End If

If Button = 2 Then
  Picture1.MousePointer = 15
  DownX = x
  DownY = y
Else
  If Button = 4 Then
   nlg = (Val(Xinc.text) + Val(Yinc.text)) / 2
   If Shift = 1 Then
     Pic.xmax.text = Val(Pic.xmax.text) + nlg
     Pic.xmin.text = Val(Pic.xmin.text) - nlg
     Pic.ymax.text = Val(Pic.ymax.text) + nlg
     Pic.ymin.text = Val(Pic.ymin.text) - nlg
   Else
      Pic.xmax.text = Val(Pic.xmax.text) - nlg
      Pic.xmin.text = Val(Pic.xmin.text) + nlg
      Pic.ymax.text = Val(Pic.ymax.text) - nlg
      Pic.ymin.text = Val(Pic.ymin.text) + nlg
   End If
   Call DrawGrp
  End If
End If
End Sub
Private Sub picture1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

If Button = 2 Then
  If DownX = x And DownY = y Then
    MoveX = -MoveX
    MoveY = -MoveY
  Else
    MoveX = x - DownX
    MoveY = y - DownY
  End If
  xmin.text = Val(xmin.text) - MoveX
  xmax.text = Val(xmax.text) - MoveX
  ymin.text = Val(ymin.text) - MoveY
  ymax.text = Val(ymax.text) - MoveY
  Picture1.MousePointer = 2
  Call DrawGrp
End If
End Sub



Private Sub polar1_Click()
If polar1.Checked = False Then
  polar1.Checked = True
  square.Checked = False
End If

If polar1.Checked = True And unfun1.Checked = True And ImplicitFun.Checked = False Then Combo2.text = "500"
If polar1.Checked = True Then
If ExplicitFun.Checked = True Then Label1.Caption = "ρ =": Text1.Left = 104
Yinc.Visible = False
Label13.Visible = False
Label12.Caption = "ρ Increment"
Label12.ToolTipText = ""
End If


If polar1.Checked = False Then
'If ExplicitFun.Checked = True Then Label1.Caption = "y =" Else Label1.Caption = "F(x,y)="
If ExplicitFun.Checked = True Then Text1.Left = 104 Else Text1.Left = 48
Yinc.Visible = True
Label13.Visible = True
Label12.Caption = "X Increment"

End If
End Sub




Private Sub savep_Click()
CommonDialog1.ShowSave
If CommonDialog1.FileName <> "" And InStr(CommonDialog1.FileName, ".") = 0 Then
CommonDialog1.FileName = CommonDialog1.FileName & ".bmp"
End If
If CommonDialog1.FileName <> "" Then SavePicture Picture1.Image, CommonDialog1.FileName

End Sub



Public Sub pack()
aoo$ = UCase$(aoo$)
PI# = 4 * Atn(1)
Select Case aoo$
    Case "ABS(V)"
    y = Abs(x)
    
    Case "SQR(V)"
    If x >= 0 Then y = Sqr(x) Else y = ymax + 1
    
    Case "LN(V)"
    y = Log(x)
    
    Case "SIN(V)"
     y = Sin(x)
    
    Case "COS(V)"
      y = Cos(x)
    
    Case "TAN(V)", "TG(V)"
      y = Tan(x)
    
    Case "ARCSIN(V)"
      y = Atn(x / Sqr(1 - x ^ 2))
    
    Case "ARCCOS(V)"
      y = PI# / 2 - Atn(x / Sqr(1 - x ^ 2))
    
    
    Case "(V)^2"
    y = x ^ 2
    Case "(V)^3"
    y = x ^ 3
    
    Case "ARCTG(V)", "ATN(V)"
     y = Atn(x)
    
    Case "ARCCTG(V)"
     y = PI# / 2 - Atn(x)
    
    Case "ARCSEC(V)"
    
     y = PI# / 2 - Atn((1 / x) / Sqr(1 - (1 / x) ^ 2))
    Case "ARCCSC(V)"
    
     y = Atn((1 / x) / Sqr(1 - (1 / x) ^ 2))
    
    Case "EP(V)", "EXP(V)"
    y = Exp(x)
    
    Case "COT(V)"
     y = 1 / (Tan(x))
    
    Case "SEC(V)"
     y = 1 / (Cos(x))
    
    Case "CSC(V)"
     y = 1 / (Sin(x))
    
    Case "LG(V)"
    y = Log(x) / Log(10)
    
    Case "SH(V)"
    y = (Exp(x) - Exp(-x)) / 2
    
    Case "CH(V)"
    y = (Exp(x) + Exp(-x)) / 2
    
    Case "TH(V)"
    y = (Exp(x) - Exp(-x)) / (Exp(x) + Exp(-x))
    
    Case "CTH(V)"
    y = (Exp(x) + Exp(-x)) / (Exp(x) - Exp(-x))
    
    Case "SECH(V)"
    y = 2 / (Exp(x) + Exp(-x))
    
    Case "CSCH(V)"
    y = 2 / (Exp(x) - Exp(-x))
    
    Case "ARSH(V)"
    y = Log(x + Sqr(x ^ 2 + 1))
    
    Case "ARCH(V)"
    y = Log(x + Sqr(x ^ 2 - 1)): Beep '+
    
    Case "ARTH(V)"
    y = (Log((x + 1) / (1 - x))) / 2
    
    Case "ARCTH(V)"
    y = (Log((x + 1) / (x - 1))) / 2
    
    Case "ARSECH(V)"
    y = Arsech(x) '+
    
    Case "ARCSCH(V)"  '?
    y = Log((Sgn(x) * Sqr(x ^ 2 + 1) + 1) / x)
    
    Case "(1+1/SIN(V))^(V)"
    y = (1 + 1 / Sin(x)) ^ x
    Case "SIN(V)/(V)"
    y = Sin(x) / x
    Case "(1+1/(V))^(V)"
    y = (1 + 1 / x) ^ x
    Case "(1+(V))^(1/(V))"
    y = (1 + x) ^ (1 / x)
    Case "LN(1+(V))/(V)"
    y = (Log(1 + x)) / x
    Case "((EP1)^(V)-1)/(V)"
    y = (Exp(x) - 1) / x
    
    
    

Case Else
package = False: y = Fc(aoo$, x, 0) 'Call Calc
End Select
End Sub

Private Sub qdjf_Click()
dfintegral.Show
dfintegral.Txtfx.text = Pic.Text1.text
End Sub

Private Sub qdsh_Click()
 der.Show
 der.fx.text = Text1.text
 Sendkeys "{Enter}"
 
End Sub

Private Sub quit_Click()
If Running = 1 Then
Command3_Click
End If
Pic.Hide
End Sub



Private Sub savepic_Click()
CommonDialog1.ShowSave
If CommonDialog1.FileName <> "" And InStr(CommonDialog1.FileName, ".") = 0 Then
CommonDialog1.FileName = CommonDialog1.FileName & ".bmp"
End If
If CommonDialog1.FileName <> "" Then SavePicture Picture1.Image, CommonDialog1.FileName

End Sub

Private Sub showweb_Click()
If showweb = 0 Then Picture1.Cls
End Sub


Private Sub Showscale_Click()
If Showscale.Checked = False Then
  Showscale.Checked = True
  'showscale.Checked = True
Else
  Showscale.Checked = False
  'showscale.Checked = false
End If
End Sub

Public Sub showweb1_Click()
If showweb1.Checked = False Then
  showweb1.Checked = True
  'showweb1.Checked = True
Else
showweb1.Checked = False
  'showweb1.Checked = false
End If
End Sub



Private Sub square_Click()
If square.Checked = False Then
  square.Checked = True
  polar1.Checked = False
  'polar1.Checked = False
End If

If polar1.Checked = True And unfun1.Checked = True And ImplicitFun.Checked = False Then Combo2.text = "500"
If polar1.Checked = True Then
Label1.Caption = "ρ ="
Yinc.Visible = False
Label13.Visible = False
Label12.Caption = "ρ Increment"
Label12.ToolTipText = ""
End If


If polar1.Checked = False Then
If ExplicitFun.Checked = True Then Label1.Caption = "y =" 'Else Label1.Caption = "F(x,y)="
If ExplicitFun.Checked = True Then Text1.Left = 104 Else Text1.Left = 48
Yinc.Visible = True
Label13.Visible = True
Label12.Caption = "X Increment"

End If
End Sub

Private Sub sytx_Click()
On Error Resume Next
SavePicture Picture1.Image, App.Path & "\Temp.bmp"
Picture1.Picture = LoadPicture(App.Path & "\Backup.bmp")
Kill App.Path & "\Backup.bmp"
Name App.Path & "\temp.bmp" As App.Path & "\Backup.bmp"
End Sub

Private Sub text1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
 
  If th = False Then
    Text1h.Enabled = True
  Else
    Text1.Height = 215
  End If
  
End Sub
Private Sub text1_KeyDown(keycode As Integer, Shift As Integer)
 ShiftDown = (Shift And vbShiftMask) > 0
 altdown = (Shift And vbAltMask) > 0
 CtrlDown = (Shift And vbCtrlMask) > 0

 Select Case keycode
   Case vbKeyF2
   Call sytx_Click
   Case vbKeyF4
   frmMagnifier.Show
   Case vbKeyF6
   If fclb.Checked = False Then
     FctList.Show
     fclb.Checked = True
   Else
    FctList.Hide
    fclb.Checked = False
   End If
   Case vbKeyF3
   PicCtrl.Show
   Case vbKeyF5
   Call lingdian_Click
   Case vbKeyF7
   Call dnc_Click
   Case vbKeyF10
   Pic.PopupMenu popup
   Case vbKeyF12
   xtys_Click
   Case vbKeyF8
   Showscale_Click
   Case vbKeyF9
   showweb1_Click
   Case vbKeyEscape
   Call Command1_Click
   Case vbKeySpace
   If ShiftDown Then Text1.text = "": Sendkeys "{BACKSPACE}"
   Case 83
   If ShiftDown Then Sendkeys "{BACKSPACE}": Text1.text = Text1.text + "sin": T1sf
   Case 79
   If ShiftDown Then Sendkeys "{BACKSPACE}": Text1.text = Text1.text + "cot": T1sf
   Case 88
   If ShiftDown Then Sendkeys "{BACKSPACE}": Text1.text = Text1.text + "exp": T1sf
   Case 84
   If ShiftDown Then Sendkeys "{BACKSPACE}": Text1.text = Text1.text + "tan": T1sf
   Case 76
   If ShiftDown Then Sendkeys "{BACKSPACE}": Text1.text = Text1.text + "log": T1sf
   Case 67
   If ShiftDown Then Sendkeys "{BACKSPACE}": Text1.text = Text1.text + "cos": T1sf
   Case 65
   If ShiftDown Then Sendkeys "{BACKSPACE}": Text1.text = Text1.text + "arc": T1sf
   Case 69
   If ShiftDown Then Sendkeys "{BACKSPACE}": Text1.text = Text1.text + "[e]": T1sf
   'Case vbKeyAdd
   'Case vbKeySubtract
   'Case vbKeyNumpad8
   'Case vbKeyNumpad4
   'Case vbKeyNumpad6
   'Case vbKeyNumpad2

   
End Select
End Sub
Private Sub T1sf()

Text1.SelStart = Len(Text1.text): Text1.SetFocus
End Sub

Private Sub Text1h_Timer()
If Text1.Height >= 215 Then Text1.Height = 215: Text1h.Enabled = False: Exit Sub
  Text1.Height = Text1.Height + 5
End Sub

Private Sub TextSetFocus_Timer()
On Error Resume Next
Text1.SetFocus
TextSetFocus.Enabled = False
End Sub

Private Sub Timer1_Timer()
If Frame5.Height < 369 Then Frame5.Height = Frame5.Height + 13.467
If Frame5.Width < 137 Then Frame5.Width = Frame5.Width + 5
If Frame5.Height >= 369 And Frame5.Width >= 137 Then Timer1.Enabled = False: t1e = 1

End Sub


Private Sub txkz_Click()
PicCtrl.Show
End Sub

Private Sub unfun1_Click()
If unfun1.Checked = False Then
  unfun1.Checked = True
Else
unfun1.Checked = False
End If
If polar1.Checked = True And unfun1.Checked = True And ImplicitFun.Checked = False Then Combo2.text = "500"
End Sub

Private Sub white_Click()
Picture1.BackColor = &HFFFFFF
Picture1.ForeColor = &H0&
End Sub

Private Sub xoy1_Click()
If xoy1.Checked = False Then
  xoy1.Checked = True
  
Else
xoy1.Checked = False
 
End If
End Sub

Private Sub PicPset(zh, xmin, xmax)
Dim xp(0 To 1024)
'Dim u#(1 To 1024)
On Error GoTo nxt
PI = 4 * Atn(1)

Erase xp
k = 0
If zh = 1 Then
  If polar1.Checked = False Then l = Val(Pic.ymin.text) Else l = 0
  If polar1.Checked = False Then r = Val(Pic.ymax.text) Else r = 2 * PI
Else
 If Text8.text = "" Then
   If polar1.Checked = False Then l = xmin Else l = 0
 Else
   l = Val(Text8.text)
 End If
 
 If Text9.text = "" Then
   If polar1.Checked = False Then r = xmax Else l = 2 * PI
 Else
   r = Val(Text9.text)
 End If
 If l < Val(xmin) Then l = xmin
 If r > Val(xmax) Then r = xmax
End If

'st = Val(Pic.Xinc.Text)
'If precision3.Text = "自动" Then
  'st = Abs(xmax - xmin) / 10
'Else
  st = Abs(xmax - xmin) / Abs(Val(Trim(precision3.text)))
'End If

If st <= 0 Or st > Abs(l - r) Then st = Abs(l - r)



For nx = l To r Step st
   'pbv = Fix(Abs(nx - l) / Abs(R - l) * 1000)
   'If pbc <= 100 Then Pic.ProgressBar1.Value = pbv
    
    x = nx
  
  aoo$ = ated
 ' Call Calc
  'ny = ms#
  ny = Fc(aoo$, x, yh)
  
  If ny = 0 Then
   
    xp(k) = nx
    i = k + 1
    xp(i) = Fix(r) + 10
    k = k + 2
    GoTo l1
  End If
  If pass = 1 And ly <> 0 And ny <> 0 And Sgn(ny) <> Sgn(ly) Then
    If Sgn(ny) = -1 Then xp(k) = nx: i = k + 1: xp(i) = lx Else xp(k) = lx: i = k + 1: xp(i) = nx
    k = k + 2
  End If
l1: lx = nx
   ly = ny
   pass = 1

nxt: If err <> 0 Then Resume l5

 
DoEvents
l5: Next nx

pass = 0

On Error GoTo l2
For j = 0 To k - 2 Step 2
  a = xp(j)
  b = xp(j + 1)
  If b = Fix(r) + 10 Then
    'If zh = 1 Then Picture1.PSet (yh, A) Else Pic.Picture1.PSet (A, yh)
If (unfun1.Checked = False Eqv zh = 0) = ture Then
  If polar1.Checked = False Then
     Pic.Picture1.PSet (yh, a)
  Else
     Pic.Picture1.PSet (a * Cos(yh), a * Sin(yh))
  End If
Else
  If polar1.Checked = False Then
     Pic.Picture1.PSet (a, yh)
  Else
  Pic.Picture1.PSet (yh * Cos(a), yh * Sin(a))
  End If
End If
    
    
    GoTo l2
  End If
l3:

  x0 = (a + b) / 2
  If x0 = 0 Then GoTo l2
  If Val(a) = Val(b) Then GoTo l4
  
  aoo$ = ated
   x = x0 ': Call Calc: y0 = ms#
  y0 = Fc(aoo$, x, yh)
  
  aoo$ = ated 'translate(aop$)
   x = a ': Call Calc: ya = ms#
  ya = Fc(aoo$, x, yh)
  
  aoo$ = ated 'translate(aop$)
   x = b ': Call Calc: yb = ms#
  yb = Fc(aoo$, x, yh)

 If y0 = yb Then GoTo l4
   
  If y0 > 0 And y0 < yb Then
    b = x0
    Else
    If y0 < 0 And y0 > ya Then a = x0 Else GoTo l2 'l2
  End If
 
  If Abs(y0) > 10 ^ (-Val(Pic.precision.text)) Then GoTo l3
l4:  'If zh = 1 Then Pic.Picture1.PSet (yh, x0) Else Pic.Picture1.PSet (x0, yh)
If (unfun1.Checked = False Eqv zh = 0) = ture Then
  If polar1.Checked = False Then
     Pic.Picture1.PSet (yh, x0)
  Else
     Pic.Picture1.PSet (x0 * Cos(yh), x0 * Sin(yh))
  End If
Else
  If polar1.Checked = False Then
     Pic.Picture1.PSet (x0, yh)
  Else
  Pic.Picture1.PSet (yh * Cos(x0), yh * Sin(x0))
  End If
End If


l2: If err <> 0 Then Resume Next
DoEvents
Next j
l6: k = 0
j = 0
End Sub






Private Sub xtys_Click()
CommonDialog1.ShowColor
Pic.Picture1.ForeColor = CommonDialog1.color
Pfc = Pic.Picture1.ForeColor
End Sub

Private Sub zdybjs_Click()
CommonDialog1.ShowColor
Pic.Picture1.BackColor = CommonDialog1.color
End Sub

Private Sub zhlc_Click()
series.Show
series.Text3.text = Pic.Text1.text
End Sub


Private Sub zxecf_Click()
Dim a(), x(), y() As Double
  Dim m As Double
  Dim k, i, j As Integer
  Dim l, p As String
  Dim iptX, iptY As String
  On Error GoTo l1:
  n = 0
  Picture1.DrawWidth = Val(linewide.text) + 2

 Do
l2:
    iptX = InputBox("x(" & n & ")=", "最小二乘法拟合直线")
    If iptX = "" Then
      iptY = ""
    Else
      If iptX <> "b" Then iptY = InputBox("y(" & n & ")=", "最小二乘法拟合直线")
    End If
    
    If iptX = "b" Or iptY = "b" Then
      colorbak = Picture1.ForeColor

      Picture1.ForeColor = Picture1.BackColor
      
      

      n = n - 1
      Picture1.PSet (x(n), y(n))
      Picture1.ForeColor = colorbak
      GoTo l2:
    End If
    ReDim Preserve x(n + 1)
    ReDim Preserve y(n + 1)
    ReDim Preserve a(n + 1)
    If iptX <> "" Then
      Picture1.CurrentX = x(k)
      Picture1.CurrentY = y(k)
      Picture1.PSet (iptX, iptY)
    End If
    
    x(n) = Val(iptX)
    y(n) = Val(iptY)
    n = n + 1
  Loop Until iptX = "" Or iptY = ""
    
    n = n - 2
    If n < 1 Then Exit Sub  '/(n+1)
    
  Dim xb, yb, xyb, xpb, ypb As Double
  For i = 0 To n
    xb = xb + x(i)
    yb = yb + y(i)
    xyb = xyb + x(i) * y(i)
    xpb = xpb + x(i) ^ 2
    ypb = ypb + y(i) ^ 2
  Next i
  xb = xb / (n + 1)
  yb = yb / (n + 1)
  xyb = xyb / (n + 1)
  xpb = xpb / (n + 1)
  ypb = ypb / (n + 1)
  
  Dim a0 As Double
  Dim b As Double
  
  a0 = (xyb - xb * yb) / (xpb - xb ^ 2)
  b = yb - a0 * xb
  
  Dim Lxy, Lxx, Lyy, r As Double
  Lxy = xyb - xb * yb
  Lxx = xpb - xb ^ 2
  Lyy = ypb - yb ^ 2
  r = Lxy / Sqr(Lxx * Lyy)
  
  msg = MsgBox(r, vbDefaultButton1, "线性关系符合程度")
  Text1.text = Trim(Str((a0))) & "x" & AddNum(b)
   
 
 
  Picture1.DrawWidth = Val(linewide.text)

 Exit Sub
l1: msg = MsgBox("无法用最小二乘法拟合直线。", vbExclamation, "最小二乘法拟合直线")
End Sub
