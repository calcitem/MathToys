Attribute VB_Name = "hpcsq"

Public Function Add(s1, s2)
Dim a() As Integer
Dim b() As Integer
Dim c() As Integer
Dim i, l, k1, k2 As Integer
On Error GoTo 40
sm1 = s1: sm2 = s2
If sm1 = "0" Then Add = sm2: GoTo 30
If sm2 = "0" Then Add = sm1: GoTo 30

If Left(sm1, 1) = "+" Then sm1 = Right(sm1, Len(sm1) - 1)
If Left(sm2, 1) = "+" Then sm2 = Right(sm2, Len(sm2) - 1)

If Left(sm1, 1) = "-" And Left(sm2, 1) <> "-" Then
  Add = Subt(sm2, Right(sm1, Len(sm1) - 1))
  GoTo 30
End If

If Left(sm1, 1) = "-" And Left(sm2, 1) = "-" Then
  sm1 = Right(sm1, Len(sm1) - 1)
  sm2 = Right(sm2, Len(sm2) - 1)
  fh = "-"
End If

If Left(sm1, 1) <> "-" And Left(sm2, 1) = "-" Then
  Add = Subt(sm1, Right(sm2, Len(sm2) - 1))
  GoTo 30
End If

If Len(sm1) <= 6 And Len(sm2) <= 6 Then
  Add = Str(Val(sm1) + Val(sm2))
  GoTo 20
End If

5:
P1 = InStr(sm1, ".")
P2 = InStr(sm2, ".")

If P1 = 0 And P2 = 0 Then GoTo 10:
If P1 = 0 Then sm1 = sm1 + ".0": GoTo 5:
If P2 = 0 Then sm2 = sm2 + ".0": GoTo 5

s1r = Right(sm1, Len(sm1) - P1)
s2r = Right(sm2, Len(sm2) - P2)
n = Len(s1r) - Len(s2r)
If n < 0 Then s1r = s1r + String(-n, "0"): p = Len(s2r)
If n > 0 Then s2r = s2r + String(n, "0"): p = Len(s1r)
If n = 0 Then p = Len(s1r)
sm1 = Left(sm1, P1 - 1) + s1r
sm2 = Left(sm2, P2 - 1) + s2r

10:
sm1 = String(4 - (Len(sm1) Mod 4), "0") + sm1
sm2 = String(4 - (Len(sm2) Mod 4), "0") + sm2

If Len(sm1) > Len(sm2) Then lgth = Len(sm1) \ 4 + 6 Else lgth = Len(sm2) \ 4 + 6
ReDim a(1 To lgth) As Integer
ReDim b(1 To lgth) As Integer
ReDim c(1 To lgth) As Integer



     
    l = Len(sm1)
     k1 = lgth
     Do Until l <= 0
          s = mid(sm1, l - 3, 4)
          a(k1) = Val(s)
          k1 = k1 - 1
          sm1 = mid(sm1, 1, l - 4)
          l = l - 4
     Loop
     
   l = Len(sm2)
     k2 = lgth
    Do Until l <= 0
          s = mid(sm2, l - 3, 4)
          b(k2) = Val(s)
          k2 = k2 - 1
          sm2 = mid(sm2, 1, l - 4)
          l = l - 4
     Loop
     

    If k1 < k2 Then k = k1 + 1 Else k = k2 + 1
     y = 0
     For i = lgth To k Step -1
     
         x = a(i) + b(i) + y
         c(i) = x Mod 10000
         y = x \ 10000
    Next i
      If y <> 0 Then k = k - 1: c(k) = y

   For i = k To lgth
    If c(i) < 1000 Then Add = Add + "0"
    If c(i) < 100 Then Add = Add + "0"
    If c(i) < 10 Then Add = Add + "0"
     
     Add = Add + LTrim(Str(c(i)))
  Next i
  
  

If P1 = 0 And P2 = 0 Then GoTo 20
If Len(Add) < p Then
  Add = String(p - Len(Add), "0") + Add
End If
Add = Left(Add, Len(Add) - p) + "." + Right(Add, p)
20:
Do Until Left(Add, 1) <> "0" Or mid(Add, 2, 1) = "."
   Add = Right(Add, Len(Add) - 1)
Loop

If InStr(Add, ".") <> 0 Then
Do Until Right(Add, 1) <> "0"
   Add = Left(Add, Len(Add) - 1)
Loop
End If

Add = Trim(Add)
If Right(Add, 1) = "." Then Add = Left(Add, Len(Add) - 1)
If Left(Add, 1) = "." Then Add = "0" + Add
Add = Trim(fh + Add)
30: If Add = "-0" Then Add = "0"
40: If err <> 0 Then msg = MsgBox("计算器无法完成一个加法操作。" & Chr(13) _
& "这极有可能是由于程序设计缺陷造成的。" & Chr(13) _
& "请与作者联系报告此错误。", vbCritical, "未知错误")
End Function


Public Function Mpc(s1, s2)
On Error GoTo 30
sm1 = s1: sm2 = s2
If sm1 = "0" Or sm2 = "0" Then Mpc = "0": GoTo 30
If Left(sm1, 1) = "+" Then sm1 = Right(sm1, Len(sm1) - 1)
If Left(sm2, 1) = "+" Then sm2 = Right(sm2, Len(sm2) - 1)
If Left(sm1, 1) = "-" And Left(sm2, 1) = "-" Then
  sm1 = Right(sm1, Len(sm1) - 1)
  sm2 = Right(sm2, Len(sm2) - 1)
  fh = ""
End If
If Left(sm1, 1) = "-" Then
  sm1 = Right(sm1, Len(sm1) - 1)
  fh = "-"
End If
If Left(sm2, 1) = "-" Then
  sm2 = Right(sm2, Len(sm2) - 1)
  fh = "-"
End If

If Len(sm1) <= 7 And Len(sm2) <= 7 Then
  Mpc = Str(Val(sm1) * Val(sm2))
  GoTo 20
End If

5:
P1 = InStr(sm1, ".")
P2 = InStr(sm2, ".")

If P1 = 0 And P2 = 0 Then GoTo 10:
If P1 = 0 Then sm1 = sm1 + ".0": GoTo 5:
If P2 = 0 Then sm2 = sm2 + ".0": GoTo 5

s1r = Right(sm1, Len(sm1) - P1)
s2r = Right(sm2, Len(sm2) - P2)
p = Len(s1r) + Len(s2r)
sm1 = Left(sm1, P1 - 1) + s1r
sm2 = Left(sm2, P2 - 1) + s2r

10:
Do Until sm2 = ""
  s = Left(sm2, 1)
  sm2 = Right(sm2, Len(sm2) - 1)
  
  Mpc = Add(Mpc, Mp(sm1, s) + String(Len(sm2), "0"))

Loop
If P1 = 0 And P2 = 0 Then GoTo 20:
If Len(Mpc) < p Then Mpc = String(p - Len(Mpc), "0") + Mpc
Mpc = Left(Mpc, Len(Mpc) - p) + "." + Right(Mpc, p)
20:
Do Until Left(Mpc, 1) <> "0" Or mid(Mpc, 2, 1) = "."
   Mpc = Right(Mpc, Len(Mpc) - 1)
Loop


If InStr(Mpc, ".") <> 0 Then
Do Until Right(Mpc, 1) <> "0"
   Mpc = Left(Mpc, Len(Mpc) - 1)
Loop
End If

Mpc = Trim(Mpc)
If Right(Mpc, 1) = "." Then Mpc = Left(Mpc, Len(Mpc) - 1)
If Left(Mpc, 1) = "." Then Mpc = "0" + Mpc
Mpc = Trim(fh + Mpc)
30: If err <> 0 Then msg = MsgBox("计算器无法完成一个乘法操作。" & Chr(13) _
& "这极有可能是由于程序设计缺陷造成的。" & Chr(13) _
& "请与作者联系报告此错误。", vbCritical, "未知错误")
End Function


Public Function Mp(aa, bb)
Dim a() As Integer
Dim c() As Integer
Dim i, l, k1, k2 As Integer
On Error GoTo 30
If bb = "0" Then Mp = "0000": GoTo 30
lgt = Len(aa) \ 4 + 6
ReDim a(1 To lgt) As Integer
ReDim c(1 To lgt) As Integer
aa1 = aa: bb1 = bb

aa1 = String(4 - (Len(aa1) Mod 4), "0") + aa1
l = Len(aa1)
k1 = lgt
Do Until l <= 0
    s = mid(aa1, l - 3, 4)
    a(k1) = Val(s)
    k1 = k1 - 1
    aa1 = mid(aa1, 1, l - 4)
    l = l - 4
Loop


b = Val(bb1)

k = k1 + 1
y = 0
For i = lgt To k Step -1
         x = a(i) * b + y
         c(i) = x Mod 10000
         y = x \ 10000
Next i
     
If y <> 0 Then k = k - 1: c(k) = y
For i = k To lgt
    If c(i) < 1000 Then Mp = Mp + "0"
    If c(i) < 100 Then Mp = Mp + "0"
    If c(i) < 10 Then Mp = Mp + "0"
     
     Mp = Mp + LTrim(Str(c(i)))
Next i
30: If err <> 0 Then msg = MsgBox("计算器无法完成一个内部乘法操作。" & Chr(13) _
& "这极有可能是由于程序设计缺陷造成的。" & Chr(13) _
& "请与作者联系报告此错误。", vbCritical, "未知错误"): err.Clear
End Function

Public Function Subt(s1, s2)
Dim s As String
Dim a() As Integer
Dim b() As Integer
Dim c() As Integer
Dim i, l, k1 As Integer
On Error GoTo 30
sm1 = s1: sm2 = s2
If Left(sm1, 1) = "+" Then sm1 = Right(sm1, Len(sm1) - 1)
If Left(sm2, 1) = "+" Then sm2 = Right(sm2, Len(sm2) - 1)

If Left(sm1, 1) = "-" And Left(sm2, 1) <> "-" Then
  Subt = "-" + Add(Right(sm1, Len(sm1) - 1), sm2)
  GoTo 30
End If

If Left(sm1, 1) = "-" And Left(sm2, 1) = "-" Then
  s = Right(sm1, Len(sm1) - 1)
  sm1 = Right(sm2, Len(sm2) - 1)
  sm2 = s
End If

If Left(sm1, 1) <> "-" And Left(sm2, 1) = "-" Then
  Subt = Add(sm1, Right(sm2, Len(sm2) - 1))
  GoTo 30
End If

If Len(sm1) <= 8 And Len(sm2) <= 8 Then
  Subt = Str(Val(sm1) - Val(sm2))
  GoTo 20
End If

5:
P1 = InStr(sm1, ".")
P2 = InStr(sm2, ".")

If P1 = 0 And P2 = 0 Then GoTo 10:
If P1 = 0 Then sm1 = sm1 + ".0": GoTo 5:
If P2 = 0 Then sm2 = sm2 + ".0": GoTo 5

s1r = Right(sm1, Len(sm1) - P1)
s2r = Right(sm2, Len(sm2) - P2)
n = Len(s1r) - Len(s2r)
If n < 0 Then s1r = s1r + String(-n, "0"): p = Len(s2r)
If n > 0 Then s2r = s2r + String(n, "0"): p = Len(s1r)
If n = 0 Then p = Len(s1r)
sm1 = Left(sm1, P1 - 1) + s1r
sm2 = Left(sm2, P2 - 1) + s2r

10:


     j = Len(sm1) - Len(sm2)
     
    If j > 0 Then
       sm2 = String(j, "0") + sm2
    Else
       sm1 = String(-j, "0") + sm1
       
    End If
    
    sm1 = String(4 - (Len(sm1) Mod 4), "0") + sm1
    sm2 = String(4 - (Len(sm2) Mod 4), "0") + sm2


     fh = ""
     If sm1 < sm2 Then fh = "-": s = sm1: sm1 = sm2: sm2 = s
        
If Len(sm1) > Len(sm2) Then lgth = Len(sm1) + 2 Else lgth = Len(sm2) + 2
ReDim a(1 To lgth) As Integer
ReDim b(1 To lgth) As Integer
ReDim c(1 To lgth) As Integer


    

 l = Len(sm1)
     k1 = lgth
     Do Until l <= 0
          s = mid(sm1, l - 3, 4)
          a(k1) = Val(s)
          s = mid(sm2, l - 3, 4)
          b(k1) = Val(s)
          
          k1 = k1 - 1
          sm1 = mid(sm1, 1, l - 4)
          sm2 = mid(sm2, 1, l - 4)
          l = l - 4
     Loop



k1 = k1 + 1

 
For i = lgth To k1 Step -1
         x = a(i) - b(i)
         If x < 0 Then x = x + 10000: a(i - 1) = a(i - 1) - 1
         c(i) = x
         
Next i



    For i = k1 To lgth
    If c(i) < 1000 Then Subt = Subt + "0"
    If c(i) < 100 Then Subt = Subt + "0"
    If c(i) < 10 Then Subt = Subt + "0"
    
    Subt = Subt + LTrim(Str(c(i)))
     Next i
   
If P1 = 0 And P2 = 0 Then GoTo 20

If Len(Subt) < p Then
  Subt = String(p - Len(Subt), "0") + Subt
End If

Subt = Left(Subt, Len(Subt) - p) + "." + Right(Subt, p)
20:
Do Until Left(Subt, 1) <> "0" Or mid(Subt, 2, 1) = "."
   Subt = Right(Subt, Len(Subt) - 1)
Loop

If InStr(Subt, ".") <> 0 Then
Do Until Right(Subt, 1) <> "0"
   Subt = Left(Subt, Len(Subt) - 1)
Loop
End If

Subt = Trim(Subt)
If Left(Subt, 1) = "." Then Subt = "0" + Subt
If Right(Subt, 1) = "." Then Subt = Left(Subt, Len(Subt) - 1)
Subt = Trim(fh + Subt)
If Left(Subt, 2) = "-." Then Subt = "-0." + Right(Subt, Len(Subt) - 2)
30: If err <> 0 Then msg = MsgBox("计算器无法完成一个加法操作。" & Chr(13) _
& "这极有可能是由于程序设计缺陷造成的。" & Chr(13) _
& "请与作者联系报告此错误。", vbCritical, "未知错误")
End Function
Public Function Sqrt(s, pcs)
Dim dn, yushu, e1, s1 As String
On Error GoTo 30
sm = Trim(s)
If Left(sm, 1) = "-" Then sm = Right(sm, Len(sm) - 1): cp = 1
If Wnzero(s) = 0 Then Sqrt = "0": GoTo 30
If InStr(sm, ".") = 0 Then sm = sm + "."
p = InStr(sm, ".")
l = p - 1
If Fix(l / 2) <> l / 2 Then sm = "0" + sm
p = (InStr(sm, ".") - 1) / 2
l = InStr(sm, ".")
sm = Left(sm, l - 1) + Right(sm, Len(sm) - l)

dn = Left(sm, 2)

sm = Right(sm, Len(sm) - 2) + "00"

e1 = "0"
yushu = dn
  
25:
    a1 = Mpc(e1, e1)
    b1 = Nocp(a1, yushu)
    c1 = Mpc(Add(e1, "1"), Add(e1, "1"))
    d1 = Nocp(c1, yushu)
    If b1 <> 1 And d1 = 1 Then GoTo 15
    e1 = Add(e1, "1")
    GoTo 25:
15:
  

dn = Left(sm, 2)
sm = Right(sm, Len(sm) - 2) + "00"
Sqrt = Sqrt + Trim(e1)
yushu = Subt(yushu, Mpc(e1, e1)) + dn


Do Until (Wnzero(sm) = 0 And Wnzero(yushu) = 0) Or Len(Sqrt) >= pcs

  s1 = 0
  Do Until Nocp(Mpc((Add(Mpc("20", e1), s1)), s1), yushu) <> 1 _
  And Nocp(Mpc((Add(Mpc("20", e1), Add(s1, "1"))), Add(s1, "1")), yushu) = 1
    s1 = Add(s1, "1")
  DoEvents
  Loop

dn = Left(sm, 2)
sm = Right(sm, Len(sm) - 2) + "00"
Sqrt = Sqrt + Trim(s1)
yushu = Add(Subt(yushu, Mpc(Add(Mpc("20", e1), s1), s1)) + "00", dn)
e1 = Sqrt
Loop

10:
If Len(Sqrt) < p Then Sqrt = Sqrt + String(p - Len(Sqrt), "0")
Sqrt = Left(Sqrt, p) + "." + Right(Sqrt, Len(Sqrt) - p)
If Right(Sqrt, 1) = "." Then Sqrt = Left(Sqrt, Len(Sqrt) - 1)
Do Until Left(Sqrt, 1) <> "0" Or mid(Sqrt, 2, 1) = "."
   Sqrt = Right(Sqrt, Len(Sqrt) - 1)
Loop
If cp = 1 Then Sqrt = Sqrt + "i"
30: If err <> 0 Then msg = MsgBox("计算器无法完成一个开方操作。" & Chr(13) _
& "这极有可能是由于程序设计缺陷造成的。" & Chr(13) _
& "请与作者联系报告此错误。", vbCritical, "未知错误")
End Function


Public Function Nocp(s1, s2) As Integer
Nocp = 0
ss1 = s1: ss2 = s2
cp = Subt(ss1, ss2)
If Left(cp, 1) = "-" Then
Nocp = -1
Else

For i = 1 To Len(cp)
  a = mid(cp, i, 1)
  If Asc(a) >= 49 And Asc(a) <= 57 Then Nocp = 1
Next i
End If
End Function



Public Function Wnzero(s)
Wnzero = 0
For i = 1 To Len(s)
  a = mid(s, i, 1)
  If Asc(a) >= 49 And Asc(a) <= 57 Then Wnzero = 1
Next i
End Function

Public Function Dvs(s1, s2, pcs, lennum)
On Error GoTo 30
sm1 = s1: sm2 = s2
If sm1 = "0" Then Dvs = "0": GoTo 30

If Wnzero(sm2) = 0 Then msg = MsgBox("除数为零。", vbExclamation, "计算器") _
: Dvs = "Null": GoTo 30

If Left(sm1, 1) = "+" Then sm1 = Right(sm1, Len(sm1) - 1)
If Left(sm2, 1) = "+" Then sm2 = Right(sm2, Len(sm2) - 1)
If Left(sm1, 1) = "-" And Left(sm2, 1) = "-" Then
  sm1 = Right(sm1, Len(sm1) - 1)
  sm2 = Right(sm2, Len(sm2) - 1)
  fh = ""
End If
If Left(sm1, 1) = "-" Then
  sm1 = Right(sm1, Len(sm1) - 1)
  fh = "-"
End If
If Left(sm2, 1) = "-" Then
  sm2 = Right(sm2, Len(sm2) - 1)
  fh = "-"
End If


Do Until Nocp(Mpc(sm2, j), "0") = 1
  j = Add(k, "100")
Loop

quit = String(pcs, "0")
m = "0"
Do Until Nocp(k, sm1) = 0 Or InStr(j, quit) > 0 Or Len(j) >= lennum
  k = Mpc(j, sm2)
  If Nocp(k, sm1) = 1 Then
    j1 = j
    j = Mpc(Add(j, m), "0.5")
  End If
  If Nocp(k, sm1) = -1 Then
    m = j
    j = j1
  End If
  DoEvents
Loop
If InStr(j, quit) > 0 Then j = Left(j, InStr(j, quit))
Dvs = j

20:
Do Until Left(Dvs, 1) <> "0" Or mid(Dvs, 2, 1) = "."
   Dvs = Right(Dvs, Len(Dvs) - 1)
Loop


If InStr(Dvs, ".") <> 0 Then
Do Until Right(Dvs, 1) <> "0"
   Dvs = Left(Dvs, Len(Dvs) - 1)
Loop
End If

Dvs = Trim(Dvs)
If Left(Dvs, 1) = "." Then Dvs = "0" + Dvs
Dvs = Trim(fh + Dvs)
30: If err <> 0 Then msg = MsgBox("计算器无法完成一个除法操作。" & Chr(13) _
& "这极有可能是由于程序设计缺陷。" & Chr(13) _
& "请与作者联系报告此错误。", vbCritical, "未知错误")
End Function
