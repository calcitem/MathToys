VERSION 5.00
Begin VB.Form PowerPi 
   Caption         =   "Power Pi"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   1200
      TabIndex        =   5
      Top             =   2520
      Width           =   1215
   End
   Begin VB.TextBox yy 
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   2160
      Width           =   975
   End
   Begin VB.TextBox xx 
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   1560
      Width           =   975
   End
   Begin VB.TextBox pi 
      Height          =   1455
      Left            =   1800
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   120
      Width           =   2295
   End
   Begin VB.CommandButton calcpi 
      Height          =   495
      Left            =   2280
      TabIndex        =   1
      Top             =   1920
      Width           =   1215
   End
   Begin VB.TextBox p1 
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Text            =   "10"
      Top             =   240
      Width           =   975
   End
End
Attribute VB_Name = "PowerPi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub calcpi_Click()
Dim q, x, y, n, m As String
Dim p As Long
p = Val(p1.Text)
q = "0." + String(p, "0") + "1"
x = "0.5"
y = "0"
n = "0"

Do Until x = "1"
 y = "0"
  Do Until y = x
    m = Add(Mpc(x, x), Mpc(y, y))
    If Left(m, 2) = "0." Or m = "1" Then
      n = Add(n, 1): pi.Text = n
    End If
    DoEvents
    y = Add(y, q)
  Loop
  x = Add(x, q)
Loop

n = Mpc(n, "8")
n = Left(n, 1) + "." + Right(n, Len(n) - 1)
pi.Text = n


'x = q
'Do Until x = "0.5"
'y = "1"

'Do Until y = Subt("1", x)
'  If y <> x Then
'     m = Trim(Add(Mpc(x, x), Mpc(y, y)))
'    If Left(m, 2) = "0." Or m = "1" Then
'     n = Add(n, 1): ' DoEvents: 'pi.Text = n
'    End If
'  End If
'    y = Subt(y, q)
'Loop
'x = Add(x, q)
'Loop

'x = Add("0.5", q)
'Do Until x = "1"
'y = "1"

'Do Until y = x
'  If y <> x Then
'   m = Add(Mpc(x, x), Mpc(y, y))
'    If Left(m, 2) = "0." Or m = "1" Then
'     n = Add(n, 1): ' DoEvents: 'pi.Text = n
'    End If
'   End If
'     y = Subt(y, q)
'Loop
'x = Add(x, q)
'Loop
'n = Mpc(n, "8")
'n = Left(n, 1) + "." + Right(n, Len(n) - 1)
'pi.Text = n

End Sub

Private Sub Command1_Click()
pi.Text = DegXY(Val(xx.Text), Val(yy.Text))
End Sub
