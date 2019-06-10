VERSION 5.00
Begin VB.Form DeterForm 
   Caption         =   "行列式"
   ClientHeight    =   5895
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4050
   Icon            =   "determinant.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5895
   ScaleWidth      =   4050
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox v 
      Alignment       =   1  'Right Justify
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
      ForeColor       =   &H00000000&
      Height          =   345
      Left            =   480
      TabIndex        =   4
      Top             =   4320
      Width           =   2895
   End
   Begin VB.CommandButton Clear 
      DownPicture     =   "determinant.frx":08CA
      Height          =   495
      Left            =   2280
      Picture         =   "determinant.frx":0BD4
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "重置"
      Top             =   5040
      Width           =   495
   End
   Begin VB.TextBox inputtext 
      BackColor       =   &H00B5CCC2&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3975
      Left            =   480
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      ToolTipText     =   "在此处输入行列式,元素按行存放,元素间用空格隔开"
      Top             =   120
      Width           =   3135
   End
   Begin VB.CommandButton sol 
      Appearance      =   0  'Flat
      DownPicture     =   "determinant.frx":149E
      Height          =   495
      Left            =   2880
      Picture         =   "determinant.frx":15E8
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "计算"
      Top             =   5040
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
      Left            =   960
      TabIndex        =   3
      ToolTipText     =   "阶数(不需要填写)"
      Top             =   4995
      Width           =   975
   End
End
Attribute VB_Name = "DeterForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Clear_Click()
inputtext.Text = ""
v.Text = ""
End Sub

Private Sub sol_Click()
Dim n As Integer '矩阵阶数
Dim a(1 To 1024, 1 To 1024) As Double
Dim b(1 To 1024) As Double
Dim w, p As Double
Dim i, j, k As Integer

inputtext.SetFocus
'Outputtext.Text = ""
On Error GoTo l6

Mx = 0: ymax = 0: xmax = 0

nn.Caption = SpaceNum(Left(inputtext.Text, InStr(inputtext.Text, Chr(10)))) + 1
n = Fix(Abs(Val(nn.Caption)))
v.Text = ""
Matrix$ = inputtext.Text
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

If n < 5 Then
If n = 1 Then w = a(1, 1)

If n = 2 Then
  w = Det2(a(1, 1), a(1, 2), a(2, 1), a(2, 2))
End If

If n = 3 Then
  w = Det3(a(1, 1), a(1, 2), a(1, 3), a(2, 1), a(2, 2), a(2, 3), a(3, 1), a(3, 2), a(3, 3))
End If

If n = 4 Then
  w = Det4(a(1, 1), a(1, 2), a(1, 3), a(1, 4), a(2, 1), a(2, 2), a(2, 3), a(2, 4), a(3, 1), a(3, 2), a(3, 3), a(3, 4), a(4, 1), a(4, 2), a(4, 3), a(4, 4))
End If
GoTo l5:
End If
  


For i = 1 To n - 1
       

  For j = i + 1 To n
    If a(i, i) = 0 Then j = 1 / 0 Else b(j) = -a(j, i) / a(i, i)
  Next j
  For j = i + 1 To n
  
    
    
    For k = 1 To n
     
      a(j, k) = a(i, k) * b(j) + a(j, k)
    Next k
    
    
  Next j
  
Next i

w = 1
For i = 1 To n
  w = w * a(i, i)
Next i
l5:
v.Text = w

l6: If err <> 0 Then msg = MsgBox("无法计算该行列式。请检查该行列式是否输入错误,或者对其进行初等变换后再尝试一次。", vbOKOnly, "计算器")
l7:
End Sub
