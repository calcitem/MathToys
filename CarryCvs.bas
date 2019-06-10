Attribute VB_Name = "Crcvs"
Public Function qto10(a$, q)
On Error GoTo l1
sgna = 0
If Left(a$, 1) = "-" Then
  sgna = 1
  a$ = Right(a$, Len(a$) - 1)
End If
rp = InStr(a$, ".")
If rp <> 0 Then
  ig$ = Left(a$, rp - 1)
  fra$ = Right(a$, Len(a$) - rp)
Else
  ig$ = a$
  fra$ = ""
End If

n = Len(ig)

For i = 1 To n
  a1 = tran(Right(ig$, 1)) * q ^ (i - 1)
  ig$ = Left(ig$, n - i)
  rst = rst + a1
Next i
 
If fra$ <> "" Then
n = Len(fra$)

For i = 1 To n
  a1 = tran(Left(fra$, 1)) * q ^ (-i)
  fra$ = Right(fra$, n - i)
  rst = rst + a1
Next i
End If
qto10 = rst
If sgna = 1 Then qto10 = -qto10
GoTo l2
l1: msg = MsgBox("计算器在尝试将" & q & "进制数" & Left(a$, 9) & "...转换成10进制数的过程中失败。" & Chr(13) _
& "这通常是因为数字太大, 也可能是由于程序设计存在缺陷。" & Chr(13) _
& "请与作者联系报告此错误。", vbCritical, "错误")
l2:
End Function

Public Function dtoq(a10, q)
On Error GoTo l1
a = Abs(a10)

ig = Fix(a)

quotient = 1

Do Until quotient = 0
  quotient = Fix(ig / q)
  remainder = ig - quotient * q
  rst$ = untran(remainder) + rst$
  ig = quotient
Loop

If ig <> a Then
fra = a - Fix(a)
rst$ = rst$ + "."
product = 1.1
Do Until Fix(product) = product
  product = q * fra
  rst$ = rst$ + untran(Fix(product))
  fra = product - Fix(product)
Loop
End If

If a10 < 0 Then dtoq = "-" + rst$ Else dtoq = rst$

GoTo l2
l1: msg = MsgBox("计算器在尝试将10进制数" & a10 & "转换成" & q & "进制数的过程中失败。" & Chr(13) _
& "这通常是因为数字太大, 也可能是由于程序设计存在缺陷。" & Chr(13) _
& "请与作者联系报告此错误。", vbCritical, "错误")
l2:
End Function
Public Function tran(nb$)

If Asc(nb) < 58 Then
  tran = Val(nb)
Else

If Asc(nb) > 64 And Asc(nb) < 91 Then
  tran = Asc(nb) - 55
Else

If Asc(nb) > 96 And Asc(nb) < 123 Then
  tran = Asc(nb) - 87
End If
End If
End If

End Function

Public Function untran(nb) As String

If nb >= 0 And nb <= 9 Then
  untran = Right(Str(nb), 1)
Else

If nb > 9 And nb < 36 Then
  untran = Chr(nb + 55)
End If

If nb > 36 Then
msg = MsgBox("计算器无法完成10进制数" & nb & "的进制转换。" & Chr(13) _
& "这是因为进位制的基数大于36, 也可能是由于程序设计存在缺陷。" & Chr(13), _
vbCritical, "错误")
End If
End If

End Function




