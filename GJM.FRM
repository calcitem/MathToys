VERSION 5.00
Begin VB.Form GJM 
   Caption         =   "逆矩阵"
   ClientHeight    =   4920
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11355
   Icon            =   "GJM.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4920
   ScaleWidth      =   11355
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox Outputtext 
      BackColor       =   &H00B5CCC2&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
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
      Text            =   "GJM.frx":08CA
      Top             =   120
      Width           =   6855
   End
   Begin VB.CommandButton sol 
      Appearance      =   0  'Flat
      DownPicture     =   "GJM.frx":092C
      Height          =   495
      Left            =   10080
      Picture         =   "GJM.frx":0A76
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "求逆矩阵"
      Top             =   4320
      Width           =   495
   End
   Begin VB.TextBox inputtext 
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
   Begin VB.CommandButton Clear 
      DownPicture     =   "GJM.frx":1340
      Height          =   495
      Left            =   9360
      Picture         =   "GJM.frx":164A
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "重置"
      Top             =   4320
      Width           =   495
   End
   Begin VB.Label nn 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00298000&
      Height          =   495
      Left            =   8040
      TabIndex        =   4
      ToolTipText     =   "阶数(不需要填写)"
      Top             =   4280
      Width           =   975
   End
End
Attribute VB_Name = "GJM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False






Private Sub Clear_Click()
inputtext.text = ""
Outputtext.text = ""
inputtext.SetFocus
End Sub




Private Sub sol_Click()
Dim n As Integer '矩阵阶数
Dim a(1 To 1024, 1 To 1024), e1(1 To 1024), e2(1 To 1024), e3(1 To 1024)
Dim w As Double
Dim f, i, j, p, r As Integer

inputtext.SetFocus
Outputtext.text = ""
On Error GoTo l6

Mx = 0: ymax = 0: xmax = 0

nn.Caption = SpaceNum(Left(inputtext.text, InStr(inputtext.text, Chr(10)))) + 1
n = Fix(Abs(Val(nn.Caption)))

Matrix$ = inputtext.text
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






For i = 1 To n
  e2(i) = i
  e3(i) = i
Next i

For p = 1 To n - 1
  
  w = 0
  
  For i = p To n
    For j = p To n
      If Abs(a(i, j)) > w Then
        w = Abs(a(i, j))
        r = i: f = j:
      End If
    Next j
  Next i

l1:
  If p <> r Then
    For i = 1 To n
      e1(i) = a(p, i)
    Next i
  End If
  
  w = e2(p)
  e2(p) = e2(r)
  e2(r) = w
  w = e3(p)
  e3(p) = e3(f)
  e3(f) = w
  
  If p <> r Then
    For i = 1 To n
      a(p, i) = a(r, i)
      a(r, i) = e1(i)
    Next i
  End If
  
  If p <> f Then
    
    For i = 1 To n
      e1(i) = a(i, p)
    Next i
    
    For i = 1 To n
      a(i, p) = a(i, f)
      a(i, f) = e1(i)
    Next i
  
  End If
  
  
  For i = p + 1 To n
    
    w = a(i, p) / a(p, p)
    
    For j = 1 To n
      a(i, j) = a(i, j) - a(p, j) * w
    Next j
    
    a(i, p) = -w
  
  Next i


Next p


For p = n To 1 Step -1
  
  e1(p) = 1
  
  For i = p + 1 To n
    e1(i) = 0
  Next i
  
  For i = 1 To p - 1
    e1(i) = a(p, i)
  Next i
  
  For i = 1 To n
    w = 0
    For j = p + 1 To n
      w = w + a(j, i) * a(p, j)
    Next j
    e1(i) = (e1(i) - w) / a(p, p)
  Next i
  
  For i = 1 To n
    a(p, i) = e1(i)
  Next i

Next p


For i = 1 To n
  
  For j = 1 To n
    e1(j) = a(i, j)
  Next j
  
  For j = 1 To n
    w = e2(j)
    a(i, w) = e1(j)
  Next j

Next i


For i = 1 To n
  
  For j = 1 To n
    e1(j) = a(j, i)
  Next j
  
  For j = 1 To n
    w = e3(j)
    a(w, i) = e1(j)
  
  Next j

Next i

For x = 1 To n
  For y = 1 To n
    Outputtext.text = Outputtext.text & a(x, y) & "   "
  Next y
  Outputtext.text = Trim(Outputtext.text) + Chr(13) + Chr(10)
Next x

l6: If err <> 0 Then msg = MsgBox("求不出逆矩阵。", vbOKOnly, "计算器")
l7:
End Sub




