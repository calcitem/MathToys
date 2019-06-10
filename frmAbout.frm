VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "关于 CalciteM"
   ClientHeight    =   4140
   ClientLeft      =   2340
   ClientTop       =   1830
   ClientWidth     =   6195
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   -1  'True
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmAbout.frx":08CA
   ScaleHeight     =   2857.502
   ScaleMode       =   0  'User
   ScaleWidth      =   5817.425
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   840
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Text            =   "frmAbout.frx":0ED7
      Top             =   1920
      Width           =   4695
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "确 定"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   4080
      TabIndex        =   0
      ToolTipText     =   "Accept"
      Top             =   3480
      Width           =   1260
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "M   "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   585
      Left            =   3546
      TabIndex        =   8
      Top             =   0
      Width           =   525
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "方解石   "
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   465
      Left            =   2400
      TabIndex        =   7
      Top             =   480
      Width           =   1485
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   240
      Picture         =   "frmAbout.frx":1043
      Top             =   3480
      Width           =   480
   End
   Begin VB.Label lblVersion 
      BackStyle       =   0  'Transparent
      Caption         =   " 版本:  ?.??.????"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   225
      Left            =   840
      TabIndex        =   5
      Top             =   1320
      Width           =   2445
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "    "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   780
      Left            =   600
      TabIndex        =   4
      ToolTipText     =   $"frmAbout.frx":1494
      Top             =   2160
      Width           =   5265
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   180
      Left            =   240
      TabIndex        =   3
      Top             =   3120
      Width           =   3105
   End
   Begin VB.Label lblDisclaimer 
      BackStyle       =   0  'Transparent
      Caption         =   "版权所有 (C) 2000-2022 方解石工作组"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   180
      Left            =   948
      TabIndex        =   2
      ToolTipText     =   " "
      Top             =   1560
      Width           =   3840
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   0
      X2              =   5859.682
      Y1              =   810.316
      Y2              =   810.316
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Calcite  "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   465
      Left            =   2400
      TabIndex        =   1
      Top             =   43
      Width           =   1365
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdOK_Click()
  Unload Me
End Sub

Private Sub Form_Load()
    lblVersion.Caption = " 版本:  " & App.Major & "." & App.Minor & "." & App.Revision
 
End Sub

