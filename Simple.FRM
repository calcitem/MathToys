VERSION 5.00
Begin VB.Form Calculator 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "大数计算器"
   ClientHeight    =   4365
   ClientLeft      =   2580
   ClientTop       =   1485
   ClientWidth     =   3540
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "System"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Simple.frx":0000
   LinkMode        =   1  'Source
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4365
   ScaleWidth      =   3540
   WhatsThisHelp   =   -1  'True
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   2655
      Left            =   12000
      TabIndex        =   20
      Top             =   1680
      Width           =   3255
      Begin VB.Label Label1 
         Caption         =   "  请稍候..."
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
         Left            =   0
         TabIndex        =   21
         Top             =   120
         Width           =   855
      End
   End
   Begin VB.TextBox Readout 
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   1500
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   19
      Text            =   "Simple.frx":0CCA
      Top             =   120
      Width           =   3015
   End
   Begin VB.CommandButton Number 
      Caption         =   "7"
      Height          =   480
      Index           =   7
      Left            =   240
      TabIndex        =   7
      Top             =   1800
      Width           =   480
   End
   Begin VB.CommandButton Number 
      Caption         =   "8"
      Height          =   480
      Index           =   8
      Left            =   840
      TabIndex        =   8
      Top             =   1800
      Width           =   480
   End
   Begin VB.CommandButton Number 
      Caption         =   "9"
      Height          =   480
      Index           =   9
      Left            =   1440
      TabIndex        =   9
      Top             =   1800
      Width           =   480
   End
   Begin VB.CommandButton Cancel 
      Caption         =   "C"
      Height          =   480
      Left            =   2160
      TabIndex        =   10
      Top             =   1800
      Width           =   480
   End
   Begin VB.CommandButton CancelEntry 
      Caption         =   "CE"
      Height          =   480
      Left            =   2760
      TabIndex        =   11
      Top             =   1800
      Width           =   480
   End
   Begin VB.CommandButton Number 
      Caption         =   "4"
      Height          =   480
      Index           =   4
      Left            =   240
      TabIndex        =   4
      Top             =   2400
      Width           =   480
   End
   Begin VB.CommandButton Number 
      Caption         =   "5"
      Height          =   480
      Index           =   5
      Left            =   840
      TabIndex        =   5
      Top             =   2400
      Width           =   480
   End
   Begin VB.CommandButton Number 
      Caption         =   "6"
      Height          =   480
      Index           =   6
      Left            =   1440
      TabIndex        =   6
      Top             =   2400
      Width           =   480
   End
   Begin VB.CommandButton Operator 
      Caption         =   "+"
      Height          =   480
      Index           =   1
      Left            =   2160
      TabIndex        =   12
      Top             =   2400
      Width           =   480
   End
   Begin VB.CommandButton Operator 
      Caption         =   "-"
      Height          =   480
      Index           =   3
      Left            =   2760
      TabIndex        =   13
      Top             =   2400
      Width           =   480
   End
   Begin VB.CommandButton Number 
      Caption         =   "1"
      Height          =   480
      Index           =   1
      Left            =   240
      TabIndex        =   1
      Top             =   3000
      Width           =   480
   End
   Begin VB.CommandButton Number 
      Caption         =   "2"
      Height          =   480
      Index           =   2
      Left            =   840
      TabIndex        =   2
      Top             =   3000
      Width           =   480
   End
   Begin VB.CommandButton Number 
      Caption         =   "3"
      Height          =   480
      Index           =   3
      Left            =   1440
      TabIndex        =   3
      Top             =   3000
      Width           =   480
   End
   Begin VB.CommandButton Operator 
      Caption         =   "X"
      Height          =   480
      Index           =   2
      Left            =   2160
      TabIndex        =   14
      Top             =   3000
      Width           =   480
   End
   Begin VB.CommandButton Operator 
      Caption         =   "/"
      Height          =   480
      Index           =   0
      Left            =   2760
      TabIndex        =   15
      Top             =   3000
      Width           =   480
   End
   Begin VB.CommandButton Number 
      Caption         =   "0"
      Height          =   480
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   3600
      Width           =   1080
   End
   Begin VB.CommandButton Decimal 
      Caption         =   "."
      Height          =   480
      Left            =   1440
      TabIndex        =   18
      Top             =   3600
      Width           =   480
   End
   Begin VB.CommandButton Operator 
      Caption         =   "="
      Height          =   480
      Index           =   4
      Left            =   2160
      TabIndex        =   16
      Top             =   3600
      Width           =   480
   End
   Begin VB.CommandButton Percent 
      Caption         =   "√"
      Height          =   480
      Left            =   2760
      TabIndex        =   17
      Top             =   3600
      Width           =   480
   End
End
Attribute VB_Name = "Calculator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim Op1, Op2                ' 前面输入的操作数
Dim DecimalFlag As Integer  ' 小数点仍然存在吗？
Dim NumOps As Integer       ' 操作数个数
Dim LastInput               ' 指示上一次按键事件的类型
Dim OpFlag                  ' 指示未完成的操作
Dim TempReadout

' C (取消) 按钮的 Click 事件过程
' 重新设置显示并初始化变量
Private Sub Cancel_Click()
    Readout = Format(0, "0.")
    Op1 = 0
    Op2 = 0
    Form_Load
End Sub

' CE (取消输入) 按钮的 Click 事件过程
Private Sub CancelEntry_Click()
    Readout = Format(0, "0.")
    DecimalFlag = False
    LastInput = "CE"
End Sub

' 小数点 (.) 按钮的 Click 事件过程
' 如果上一次按键为运算符，初始化 readout 为 "0."；
' 否则显示时追加一个小数点
Private Sub Decimal_Click()
    If LastInput = "NEG" Then
        Readout = Format(0, "-0.")
    ElseIf LastInput <> "NUMS" Then
        Readout = Format(0, "0.")
    End If
    DecimalFlag = True
    LastInput = "NUMS"
End Sub

' 窗体的初始化过程
' 设置所有变量为其初始值
Private Sub Form_Load()
    DecimalFlag = False
    NumOps = 0
    LastInput = "NONE"
    OpFlag = " "
    Readout = Format(0, "0.")
    'Decimal.Caption = Format(0, ".")
End Sub

' 数字键 (0-9) 的 Click 事件过程
' 向显示中的数追加新数
Private Sub Number_Click(Index As Integer)
    If LastInput <> "NUMS" Then
        Readout = Format(0, ".")
        DecimalFlag = False
    End If
    If DecimalFlag Then
        Readout = Readout + Number(Index).Caption
    Else
        Readout = Left(Readout, InStr(Readout, Format(0, ".")) - 1) + Number(Index).Caption + Format(0, ".")
    End If
    If LastInput = "NEG" Then Readout = "-" & Readout
    LastInput = "NUMS"
End Sub

' 运算符 (+, -, x, /, =) 的 Click 事件过程
' 如果接下来的按键是数字键，增加 NumOps。
' 如果有一个操作数，则设置 Op1。
' 如果有两个操作数，则将 Op1 设置为 Op1 与
' 当前输入字符串的运算结果，并显示结果
Private Sub Operator_Click(Index As Integer)
    TempReadout = Readout
    If LastInput = "NUMS" Then
        NumOps = NumOps + 1
    End If
    Select Case NumOps
        Case 0
        If Operator(Index).Caption = "-" And LastInput <> "NEG" Then
            Readout = "-" & Readout
            LastInput = "NEG"
        End If
        Case 1
        Op1 = Readout
        If Operator(Index).Caption = "-" And LastInput <> "NUMS" And OpFlag <> "=" Then
            Readout = "-"
            LastInput = "NEG"
        End If
        Case 2
        Op2 = TempReadout
        Select Case OpFlag
            Case "+"
                Op1 = Add((Op1), (Op2))
            Case "-"
                Op1 = Subt((Op1), (Op2))
            Case "X"
                Op1 = Mpc((Op1), (Op2))
            Case "/"
                If Wnzero(Op2) = 0 Then
                   MsgBox "除数不能为 0", 48, "计算器"
                Else
                 ' Op1 = Dvs((Op1), (Op2), 10, 256)
                  Frame1.Left = 120
                  Op1 = Left(Dvs((Op1), (Op2), 10, 256), 64)
                  Frame1.Left = 10000
                End If
            Case "="
                Op1 = Op2
            Case "√"
                Frame1.Left = 120
                Op1 = CDbl(Op1) * CDbl(Op2)
                Frame1.Left = 10000
            End Select
        Readout = Op1
        NumOps = 1
    End Select
    If LastInput <> "NEG" Then
        LastInput = "OPS"
        OpFlag = Operator(Index).Caption
    End If
End Sub

' 百分比键 (√) 的 Click 事件过程
' 计算并显示第一个操作数的百分数
Private Sub Percent_Click()
    Frame1.Left = 120
    Readout = Sqrt(Readout, 64)
    Frame1.Left = 10000
    LastInput = "Ops"
    OpFlag = "√"
    NumOps = NumOps + 1
    DecimalFlag = True
End Sub
Private Sub number_Keydown(Index As Integer, keycode As Integer, Shift As Integer)
If keycode >= 48 And keycode <= 57 Then
  Number(keycode - 48).SetFocus
End If
'Select Case keycode
'  Case vbKeyMultiply
 ' Operator(2).SetFocus
  'Case vbKeyAdd
  '  Operator(1).SetFocus

 ' Case vbKeySeparator
'    Operator(4).SetFocus

'  Case vbKeySubtract
'   Operator(3).SetFocus
'
 ' Case vbKeyDecimal
'    Decimal_Click
    'Decimal.SetFocus
''  Case vbKeyDivide
'    Operator(0).SetFocus
  
'End Select
  SendKeys "{Enter}"
  
End Sub
Private Sub operator_Keydown(Index As Integer, keycode As Integer, Shift As Integer)
If keycode >= 48 And keycode <= 57 Then
  Number(keycode - 48).SetFocus
End If
'Select Case keycode
'  Case vbKeyMultiply
'  Operator(2).SetFocus
'  Case vbKeyAdd
'    Operator(1).SetFocus

'  Case vbKeySeparator
'    Operator(4).SetFocus

'  Case vbKeySubtract
'    Operator(3).SetFocus

'  Case vbKeyDecimal
'    Decimal_Click
'    'Decimal.SetFocus
'  Case vbKeyDivide
'    Operator(0).SetFocus
  
'End Select
  SendKeys "{Enter}"
  
End Sub


