VERSION 5.00
Begin VB.Form test 
   Caption         =   "Test"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.CommandButton Test 
      Caption         =   "test"
      Height          =   615
      Left            =   2160
      TabIndex        =   2
      Top             =   2040
      Width           =   2055
   End
   Begin VB.TextBox b 
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   960
      Width           =   4095
   End
   Begin VB.TextBox a 
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   4095
   End
End
Attribute VB_Name = "test"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Test_Click()
 b.Text = Multinomial(a.Text)
End Sub
