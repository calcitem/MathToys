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
l1: msg = MsgBox("�������ڳ��Խ�" & q & "������" & Left(a$, 9) & "...ת����10�������Ĺ�����ʧ�ܡ�" & Chr(13) _
& "��ͨ������Ϊ����̫��, Ҳ���������ڳ�����ƴ���ȱ�ݡ�" & Chr(13) _
& "����������ϵ����˴���", vbCritical, "����")
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
l1: msg = MsgBox("�������ڳ��Խ�10������" & a10 & "ת����" & q & "�������Ĺ�����ʧ�ܡ�" & Chr(13) _
& "��ͨ������Ϊ����̫��, Ҳ���������ڳ�����ƴ���ȱ�ݡ�" & Chr(13) _
& "����������ϵ����˴���", vbCritical, "����")
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
msg = MsgBox("�������޷����10������" & nb & "�Ľ���ת����" & Chr(13) _
& "������Ϊ��λ�ƵĻ�������36, Ҳ���������ڳ�����ƴ���ȱ�ݡ�" & Chr(13), _
vbCritical, "����")
End If
End If

End Function




