VERSION 5.00
Begin VB.Form GJ 
   Caption         =   "解实系数线性代数方程组"
   ClientHeight    =   4560
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7350
   Icon            =   "GJ.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4560
   ScaleWidth      =   7350
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      BackColor       =   &H00597B68&
      DownPicture     =   "GJ.frx":058A
      Height          =   495
      Left            =   5400
      Picture         =   "GJ.frx":0894
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "转到解三阶线性方程组"
      Top             =   3720
      Width           =   495
   End
   Begin VB.CommandButton Clear 
      DownPicture     =   "GJ.frx":155E
      Height          =   495
      Left            =   6000
      Picture         =   "GJ.frx":1868
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "重置"
      Top             =   3720
      Width           =   495
   End
   Begin VB.TextBox outputtext 
      BackColor       =   &H00B5CCC2&
      Height          =   3255
      Left            =   4200
      MaxLength       =   1024
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Text            =   "GJ.frx":2132
      Top             =   240
      Width           =   2895
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
      Top             =   240
      Width           =   3735
   End
   Begin VB.CommandButton solution 
      Appearance      =   0  'Flat
      DownPicture     =   "GJ.frx":21A6
      Height          =   495
      Left            =   6480
      Picture         =   "GJ.frx":22F0
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "解出"
      Top             =   3720
      Width           =   495
   End
   Begin VB.Label n 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00B5CCC2&
      Height          =   495
      Left            =   4080
      TabIndex        =   5
      ToolTipText     =   "阶数 (不需要填写)"
      Top             =   3720
      Width           =   975
   End
End
Attribute VB_Name = "GJ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a(1 To 1024, 1 To 1024) As Double
Dim b(1 To 1024) As Double
Dim ik(1 To 1024)
Dim jk(1 To 1024)
'Dim n As Integer
'dim a(),b(),ik(),jk();

Dim i, j, k, l, i0, j0 As Integer
Dim c, d As Double





Private Sub Clear_Click()
inputtext.text = ""
inputtext.SetFocus
End Sub

Private Sub Command1_Click()
fun.Show
End Sub

Private Sub solution_Click()
inputtext.SetFocus
Outputtext.text = ""
On Error GoTo l6

Mx = 0: ymax = 0: xmax = 0

n.Caption = SpaceNum(Left(inputtext.text, InStr(inputtext.text, Chr(10))))
n = Fix(Abs(Val(n.Caption)))

Matrix$ = inputtext.text
Matrix$ = LTrim(Matrix$)
For x = 1 To n
  If InStr(1, Matrix$, Chr(10)) = 0 Then mat$ = Matrix$ Else mat$ = Left(Matrix$, (InStr(1, Matrix$, Chr(10))))
  mt = Len(mat$)
  For y = 1 To n
    
    a(x, y) = Val(Left(mat$, (InStr(1, mat$, " "))))
    If Abs(a(x, y)) > Mx Then Mx = Abs(a(x, y)): ymax = x: xmax = y
    mat$ = Right(mat$, Len(mat$) - Len(Str(a(x, y))))
    mat$ = LTrim(mat$)
  Next y
  
 
  
  If InStr(1, mat$, " ") = 0 Then b(x) = mat$ Else b(x) = Val(Left(mat$, (InStr(1, mat$, " "))))
  Matrix$ = Right(Matrix$, Len(Matrix$) - mt)
  Matrix$ = LTrim(Matrix$)

Next x

'For i = 1 To n
'  temp = a(i, 1)
'  a(i, 1) = a(i, xmax)
'  a(i, xmax) = temp
'Next i

'For i = 1 To n
'  temp = a(ymax, i)
'  a(ymax, i) = a(1, i)
'  a(1, i) = temp
'Next i

'temp = b(1)
'b(1) = b(ymax)
'b(ymax) = temp


'For k = 2 To n
'  For l = k To n
'    If Abs(a(k, l)) > Abs(a(k, k)) Then
'      For i = 1 To n
'        temp = a(i, l)
'        a(i, l) = a(i, k)
'        a(i, k) = temp
'      Next i
'      End If
'    Next l
'Next k

  





For k = 1 To n
  
    c = 0

    For i = 1 To n
      For l = 1 To k - 1
        If i = ik(l) Then GoTo l2:
      Next l
      For j = 1 To n
        If Abs(a(i, j)) <= Abs(c) Then GoTo l1:
        c = a(i, j)
        i0 = i
        j0 = j
l1:   Next j
l2: Next i

    ik(k) = i0
  
    jk(k) = j0
    
    For j = 1 To n
      For l = 1 To k
        If j = jk(l) Then GoTo l4:
      Next l
      If a(i0, j) = 0 Then GoTo l4:
      a(i0, j) = a(i0, j) / c '??
      d = a(i0, j)
      For i = 1 To n
        If i = i0 Then GoTo l3:
        a(i, j) = a(i, j) - a(i, j0) * d
l3:   Next i
l4: Next j
    
    b(i0) = b(i0) / c

    d = b(i0) '??

    For i = 1 To n
      If i = i0 Then GoTo l5:
      b(i) = b(i) - a(i, j0) * d
l5: Next i

    For i = 1 To n
      a(i, j0) = 0
    Next i

Next k

For k = 1 To n
    a(1, jk(k)) = b(ik(k))
Next k


       
For k = 1 To n
  b(k) = a(1, k)
  Outputtext.text = Outputtext.text + "x" & k & "=  " & b(k) & Chr(13) & Chr(10)
Next k


GoTo l7
l6: msg = MsgBox("此方程组无唯一解。", vbOKOnly, "计算器")
l7: End Sub
