Attribute VB_Name = "ExprChk"
Public Function ExpChk(Exn As String) As String
If Exn = "" Then Exit Function
Ex = LCase$(Exn)

a = "#"
If InStr(Ex, a) > 0 Then
  ExpChk = "无法识别的字符: " & a & "   (位置:" & InStr(Ex, a) & ")    "
  Exit Function
End If
a = "$"
If InStr(Ex, a) > 0 Then
  ExpChk = "无法识别的字符: " & a & "   (位置:" & InStr(Ex, a) & ")    "
  Exit Function
End If
a = "&"
If InStr(Ex, a) > 0 Then
  ExpChk = "无法识别的字符: " & a & "   (位置:" & InStr(Ex, a) & ")    "
  Exit Function
End If
a = "_"
If InStr(Ex, a) > 0 Then
  ExpChk = "无法识别的字符: " & a & "   (位置:" & InStr(Ex, a) & ")    "
  Exit Function
End If
a = "~"
If InStr(Ex, a) > 0 Then
  ExpChk = "无法识别的字符: " & a & "   (位置:" & InStr(Ex, a) & ")    "
  Exit Function
End If
a = """"
If InStr(Ex, a) > 0 Then
  ExpChk = "无法识别的字符: " & a & "   (位置:" & InStr(Ex, a) & ")    "
  Exit Function
End If
a = ","
If InStr(Ex, a) > 0 Then
  ExpChk = "无法识别的字符: " & a & "   (位置:" & InStr(Ex, a) & ")    "
  Exit Function
End If
a = "?"
If InStr(Ex, a) > 0 Then
  ExpChk = "无法识别的字符: " & a & "   (位置:" & InStr(Ex, a) & ")    "
  Exit Function
End If
a = ":"
If InStr(Ex, a) > 0 Then
  ExpChk = "无法识别的字符: " & a & "   (位置:" & InStr(Ex, a) & ")    "
  Exit Function
End If
a = ";"
If InStr(Ex, a) > 0 Then
  ExpChk = "无法识别的字符: " & a & "   (位置:" & InStr(Ex, a) & ")    "
  Exit Function
End If
a = "<"
If InStr(Ex, a) > 0 Then
  ExpChk = "无法识别的字符: " & a & "   (位置:" & InStr(Ex, a) & ")    "
  Exit Function
End If
a = ">"
If InStr(Ex, a) > 0 Then
  ExpChk = "无法识别的字符: " & a & "   (位置:" & InStr(Ex, a) & ")    "
  Exit Function
End If

a = "^-"
If InStr(Ex, a) > 0 Then
  ExpChk = "^ 与 - 间缺少括号。" & "  (位置:" & InStr(Ex, a) & ")    "
  Exit Function
End If


While InStr(Ex, "[") > 0
    Mid$(Ex, InStr(Ex, "["), 1) = "("
Wend

While InStr(Ex, "]") > 0
    Mid$(Ex, InStr(Ex, "]"), 1) = ")"
Wend

a = InStr(Ex, "("): B = InStr(Ex, ")")
If (a > 0 And B > 0 And a > B) Or (a = 0 And B > 0) Then
  ExpChk = "表达式左端遗漏 '( '"
  Exit Function
End If

Lb = 0: rb = 0
cbo$ = Ex
For cb = 1 To Len(cbo$)
  If Left(cbo$, 1) = "(" Then Lb = Lb + 1
  If Left(cbo$, 1) = ")" Then rb = rb + 1
  cbo$ = Right(cbo$, Len(cbo$) - 1)
Next cb
If Lb > rb Then
  ExpChk = "括号不匹配。遗漏" & Str(Abs(Lb - rb)) & "个 ') '。 "
  Exit Function
End If
If Lb < rb Then
  ExpChk = "括号不匹配。遗漏" & Str(Abs(Lb - rb)) & "个 '( '。 "
  Exit Function
End If

If InStr(Ex, "`") > 0 And InStr(Ex, "log") = 0 Then
  ExpChk = "没有与 底数真数分隔符` 相对应的函数log." & "  (位置:" & InStr(Ex, "`") & ")    "
  Exit Function
End If


a = Left(Ex, 1)
Select Case a
  Case "!", "%", "^", "*", "|", "\", "/"
  ExpChk = "表达式头错误: 左端缺少表达式。 "
  Exit Function
  Case "b", "g", "h", "n", "o", "q", "u"
  ExpChk = "表达式头错误。"
  Exit Function
 End Select
 
a = Right(Ex, 1)



Select Case a
  Case "+", "-", "*", "/", "\", "@", "^", "|"
  ExpChk = "表达式尾错误: 右端缺少表达式。 "
  Exit Function
End Select



Dim kw(70) As String
kw(0) = "mod"
kw(1) = "e"
kw(2) = "abs"
kw(3) = "sqr"
kw(4) = "ep"
kw(5) = "int"
kw(6) = "fi"
kw(7) = "ln"
kw(8) = "sin"
kw(9) = "cos"
kw(10) = "tan"
kw(11) = "tg"
kw(12) = "arcsin"
kw(13) = "arccos"
kw(14) = "atn"
kw(15) = "arctan"
kw(16) = "arctg"
kw(17) = "arcctg"
kw(18) = "arccot"
kw(19) = "arcsec"
kw(20) = "arccsc"
kw(21) = "asin"
kw(22) = "acos"
kw(23) = "atan"
kw(24) = "acot"
kw(25) = "asec"
kw(26) = "acsc"
kw(27) = "ep"
kw(28) = "sgn"
kw(29) = "cot"
kw(30) = "ctg"
kw(31) = "sec"
kw(32) = "csc"
kw(33) = "log"
kw(34) = "lg"
kw(35) = "sh"
kw(36) = "ch"
kw(37) = "th"
kw(38) = "cth"
kw(39) = "sinh"
kw(40) = "cosh"
kw(41) = "tanh"
kw(42) = "coth"
kw(43) = "sech"
kw(44) = "csch"
kw(45) = "arsh"
kw(46) = "arch"
kw(47) = "arth"
kw(48) = "arcth"
kw(49) = "asinh"
kw(50) = "acosh"
kw(51) = "atanh"
kw(52) = "acoth"
kw(53) = "arsech"
kw(54) = "arcsch"
kw(55) = "asech"
kw(56) = "acsch"
kw(57) = "dms"
kw(58) = "deg"
kw(59) = "trunc"
kw(60) = "sign"
kw(61) = "round"
kw(62) = "ml"
kw(63) = "m"
kw(64) = "mr"
kw(65) = "eop"
kw(66) = "lna"

Dim sg(10) As String
sg(1) = "+"
sg(2) = "-"
sg(3) = "*"
sg(4) = "/"
sg(5) = "\"
sg(6) = "|"
sg(7) = "^"
sg(8) = "@"

B = ""


Do Until InStr(Ex, "(t)") = 0
    Ex = Left(Ex, InStr(Ex, "(t)") - 1) + "1" + Right(Ex, Len(Ex) - InStr(Ex, "(t)") - 2)
Loop
Do Until InStr(Ex, "exp") = 0
    Ex = Left(Ex, InStr(Ex, "exp") - 1) + "ep" + Right(Ex, Len(Ex) - InStr(Ex, "exp") - 2)
Loop
Do Until InStr(Ex, "x") = 0
    Ex = Left(Ex, InStr(Ex, "x") - 1) + "1" + Right(Ex, Len(Ex) - InStr(Ex, "x"))
Loop
Do Until InStr(Ex, "y") = 0
    Ex = Left(Ex, InStr(Ex, "y") - 1) + "1" + Right(Ex, Len(Ex) - InStr(Ex, "y"))
Loop
Do Until InStr(Ex, "pi") = 0
    Ex = Left(Ex, InStr(Ex, "pi") - 1) + "1" + Right(Ex, Len(Ex) - InStr(Ex, "pi") - 1)
Loop

a = Ex

l = Left(a, 1)


Do Until Len(a) < 1

Rt = 0

Do Until Asc(l) < 97 Or Asc(l) > 122
  
  
  B = B & l
  
  
  a = Right(a, Len(a) - 1)
  If a = "" Then Exit Do
  l = Left(a, 1)
Loop

If B <> "" Then
  For i = 0 To 66
    If B = kw(i) Then Rt = 1: Exit For
  Next i
  If Rt = 0 Then
    ExpChk = "无法识别的函数: " & B & "  (位置:" & InStr(Ex, B) & ")    "
    Exit Function
  End If
  
  er = 0
  
  If B <> "e" And B <> "m" And B <> "ml" And B <> "mr" Then
  For i = 1 To 8
    c = mid(Ex, InStr(Ex, B) + Len(B), 1)
    If c = sg(i) Then er = 1
  Next i
  If c = ")" Then er = 1
  
  If er = 1 Then
    If B <> "ep" Then
      ExpChk = B & "之后缺少表达式。" & "  (位置:" & InStr(Ex, B) + Len(B) - 1 & ")    "
    Else
      ExpChk = "exp" & "之后缺少表达式。" & "  (位置:" & InStr(Ex, B) + Len(B) & ")"
    End If
    Exit Function
  End If
  End If
  
End If

B = ""

If a = "" Then Exit Do
l = Left(a, 1)

Do Until Asc(l) >= 97 And Asc(l) <= 122

  If Len(a) > 0 Then a = Right(a, Len(a) - 1)
  If a = "" Then Exit Do
  l = Left(a, 1)
  
Loop

Loop





a = Ex


Do Until a = ""
 ls = 0: rs = 0:
 l = Left(a, 1)
 a = Right(a, Len(a) - 1)
 r = Left(a, 1)
 For i = 1 To 7
   If l = sg(i) Then ls = 1
   If r = sg(i) Then rs = 1
 Next i
 If ls = 1 And rs = 1 Then
    ExpChk = l & " 与 " & r & " 间缺少表达式。" & "  (位置:" & InStr(Ex, l & r) & ")    "
    Exit Function
 End If
 
Loop

er = 0
a = InStr(Ex, "@")
If a > 0 Then
  l = Left(Right(Ex, Len(Ex) - InStr(Ex, "@")), 1)
  For i = 1 To 8
   If l = sg(i) Then er = 1
  Next i
  If er = 1 Then
    ExpChk = "@ 与 " & l & " 间缺少表达式。" & "  (位置:" & InStr(Ex, "@" & l) & ")    "
    Exit Function
  End If
End If

a = Right(Ex, 1)
If Asc(a) >= 97 And Asc(a) <= 122 And a <> "m" And a <> "r" And a <> "l" Then
  ExpChk = "表达式尾错误: 右端缺少表达式。 "
  Exit Function
End If




End Function
