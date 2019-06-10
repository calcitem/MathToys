VERSION 5.00
Begin VB.Form H 
   Caption         =   "历史记录"
   ClientHeight    =   3165
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6255
   Icon            =   "History.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3165
   ScaleWidth      =   6255
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox lsjl 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
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
      Height          =   375
      Index           =   5
      Left            =   360
      MouseIcon       =   "History.frx":058A
      MousePointer    =   99  'Custom
      TabIndex        =   5
      Top             =   2520
      Width           =   5655
   End
   Begin VB.TextBox lsjl 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
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
      Height          =   375
      Index           =   4
      Left            =   360
      MouseIcon       =   "History.frx":0894
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   2040
      Width           =   5655
   End
   Begin VB.TextBox lsjl 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
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
      Height          =   375
      Index           =   3
      Left            =   360
      MouseIcon       =   "History.frx":0B9E
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   1560
      Width           =   5655
   End
   Begin VB.TextBox lsjl 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
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
      Height          =   375
      Index           =   2
      Left            =   360
      MouseIcon       =   "History.frx":0EA8
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   1080
      Width           =   5655
   End
   Begin VB.TextBox lsjl 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
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
      Height          =   375
      Index           =   1
      Left            =   360
      MouseIcon       =   "History.frx":11B2
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   600
      Width           =   5655
   End
   Begin VB.TextBox lsjl 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
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
      Height          =   375
      Index           =   0
      Left            =   360
      MouseIcon       =   "History.frx":14BC
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   120
      Width           =   5655
   End
End
Attribute VB_Name = "H"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub lsjl_Click(Index As Integer)
Pic.Text1.Text = lsjl(Index).Text
End Sub
