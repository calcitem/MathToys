VERSION 5.00
Begin VB.Form MatrixOper 
   Caption         =   "¾ØÕóÔËËã"
   ClientHeight    =   3015
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9660
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   9660
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.CommandButton Command2 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3000
      TabIndex        =   4
      Top             =   720
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "¡Á"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3000
      TabIndex        =   3
      Top             =   1320
      Width           =   615
   End
   Begin VB.TextBox C 
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
      Height          =   2655
      Left            =   6840
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   2
      Top             =   120
      Width           =   2535
   End
   Begin VB.TextBox B 
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
      Height          =   2655
      Left            =   3840
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Top             =   120
      Width           =   2535
   End
   Begin VB.TextBox A 
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
      Height          =   2655
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   120
      Width           =   2535
   End
End
Attribute VB_Name = "MatrixOper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim n As Integer '¾ØÕó½×Êý
Dim a(1 To 1024, 1 To 1024) As Double
Dim b(1 To 1024, 1 To 1024) As Double
Dim w, p As Double
Dim i, j, k As Integer

a.SetFocus

On Error GoTo l6

Mx = 0: ymax = 0: xmax = 0

nn.Caption = SpaceNum(Left(a.Text, InStr(a.Text, Chr(10)))) + 1
n = Fix(Abs(Val(nn.Caption)))
v.Text = ""
Matrix$ = a.Text
Matrix$ = LTrim(Matrix$)
For x = 1 To n
  If InStr(1, Matrix$, Chr(10)) = 0 Then mat$ = Matrix$ Else mat$ = Left(Matrix$, (InStr(1, Matrix$, Chr(10))))
  mt = Len(mat$)
  For y = 1 To n
    sp = InStr(1, mat$, " ")
    If sp <> 0 Then
      a(x, y) = Val(Left(mat$, sp))
      mat$ = Right(mat$, Len(mat$) - Len(Str(a(x, y))))
      mat$ = LTrim(mat$)
    Else
      a(x, y) = Val(mat$)
    End If
  Next y
  
 
  
  If InStr(1, mat$, " ") = 0 Then a(x, y) = Val(mat$) Else a(x, y) = Val(Left(mat$, (InStr(1, mat$, " "))))
  Matrix$ = Right(Matrix$, Len(Matrix$) - mt)
  Matrix$ = LTrim(Matrix$)

Next x



l6:

End Sub
