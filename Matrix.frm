VERSION 5.00
Begin VB.Form Matrix 
   Caption         =   "æÿ’Û‘ÀÀ„"
   ClientHeight    =   4755
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12225
   LinkTopic       =   "Form1"
   ScaleHeight     =   4755
   ScaleWidth      =   12225
   StartUpPosition =   3  '¥∞ø⁄»± °
   Begin VB.CommandButton MatrixMutil 
      Caption         =   "*"
      Height          =   495
      Left            =   4560
      TabIndex        =   3
      Top             =   4200
      Width           =   1095
   End
   Begin VB.TextBox MatrixC 
      BackColor       =   &H00B5CCC2&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3975
      Left            =   8160
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   2
      Top             =   120
      Width           =   3735
   End
   Begin VB.TextBox MatrixB 
      BackColor       =   &H00B5CCC2&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3975
      Left            =   4200
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Top             =   120
      Width           =   3735
   End
   Begin VB.TextBox MatrixA 
      BackColor       =   &H00B5CCC2&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3975
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   120
      Width           =   3735
   End
End
Attribute VB_Name = "Matrix"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub MatrixMutil_Click()


For i = 1 To m
  For j = 1 To n
    For k = 1 To s
      c(i, j) = c(i, j) + a(i, k) * b(k, j)
    Next k
  Next j
Next i
End Sub
