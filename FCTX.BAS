Attribute VB_Name = "fctx"
Private m$, ms#, mo#, alfa, un, aoo$
Private x, yh As Single
Dim u#(1 To 1024)
Public Function Fc(fctx$, fx, fY)
aoo$ = fctx$
x = fx
yh = fY
Call Calc
Fc = ms#
End Function


Public Sub Calc()




10:

Erase u#
un = 0:  alfa = 0



11:


12:
aoo$ = UCase$(aoo$)





e$ = ""
m$ = ""

If Len(aoo$) = 1 Then
    'ml# = ms#
    ms# = Val(aoo$)
    GoTo 400
End If

20:  If InStr(aoo$, "(") = 0 Then GoTo 70

30
If alfa = 0 Then GoTo o:
If Left(aoo$, 3) = "(UN" And Right(aoo$, 1) = ")" And Len(Str$(Val(Right(aoo$, Len(aoo$) - 3)))) + 4 = Len(aoo$) Then
   GoTo 80:
End If
o:
a = Len(aoo$)
bo$ = Right(aoo$, a - InStr(aoo$, "(") + 1)
c$ = Left(aoo$, a - Len(bo$))


p:

mb = InStr(Right(bo$, Len(bo$) - 1), "(")
nb = InStr(Right(bo$, Len(bo$) - 1), ")")

If mb < nb And mb <> 0 Then
    c$ = c$ + Left(bo$, mb)
    bo$ = Right(aoo$, a - Len(c$))
    GoTo p:
Else
    no$ = Left(bo$, InStr(bo$, ")"))
    d$ = Right(aoo$, a - Len(c$) - Len(no$))
    no$ = mid$(no$, 2, Len(no$) - 2)

End If

If InStr(no$, "+") = 0 Then
  If InStr(no$, "-") = 0 Then m$ = no$: Call bet: GoTo 65
End If

If InStr(no$, "*") = 0 Then
  If InStr(no$, "/") = 0 Then
    If InStr(no$, "\") = 0 Then
      If InStr(no$, "|") = 0 Then
        If InStr(no$, "^") = 0 Then
          If InStr(no$, "@") = 0 Then GoTo 60
        End If
      End If
    End If
  End If
End If


40:
a = Len(no$)
B = 32767
If InStr(no$, "+") > 0 Then B = InStr(no$, "+")
If InStr(no$, "-") > 0 And InStr(no$, "-") < B Then B = InStr(no$, "-")
If B = 32767 Then GoTo 60


50:
m$ = Left(no$, B - 1)
If InStr(m$, "^") > 0 Or InStr(m$, "@") > 0 Then Call bet Else Call alp
e$ = e$ + m$ + mid$(no$, B, 1)
no$ = Right(no$, a - B)
If InStr(no$, "*") > 0 Or InStr(no$, "/") > 0 Or InStr(no$, "\") > 0 Or InStr(no$, "|") > 0 Or InStr(no$, "^") > 0 Or InStr(no$, "@") > 0 Or InStr(no$, "+") > 0 Or InStr(no$, "-") > 0 Then GoTo 40


60:
m$ = no$
If InStr(m$, "^") > 0 Or InStr(m$, "@") > 0 Then Call bet Else Call alp
m$ = e$ + m$
e$ = ""
If InStr(m$, "^") > 0 Or InStr(m$, "@") > 0 Then Call bet Else Call alp


65:
no$ = m$
aoo$ = c$ + no$ + d$
GoTo 20


70:
If Len(c$) = 0 And Len(d$) = 0 And alfa > 0 Then GoTo 80
aoo$ = "(" + aoo$ + ")": GoTo 30


80:

'ml# = ms#
ms# = mo#
GoTo 400




400:

End Sub
Public Sub alp()
alfa = alfa + 1
mo# = 0
PI# = 4 * Atn(1)
If Left(m$, 1) <> "-" Then m$ = "+" + m$

Do Until InStr(m$, "+@") = 0
   Mid$(m$, InStr(m$, "+@"), 2) = "2@"
Loop
If Left(m$, 1) <> "+" Then m$ = "+" + m$

a:
a = Len(m$)
ao$ = Right(m$, a - 1)
B = 32767

'有没有必要跳过？
If InStr(ao$, "+") > 0 Then B = InStr(ao$, "+")

If InStr(ao$, "-") > 0 Then
  If InStr(ao$, "-") < B Then B = InStr(ao$, "-")
End If

If InStr(ao$, "*") > 0 Then
  If InStr(ao$, "*") < B Then B = InStr(ao$, "*")
End If

If InStr(ao$, "/") > 0 Then
  If InStr(ao$, "/") < B Then B = InStr(ao$, "/")
End If

If InStr(ao$, "\") > 0 Then
  If InStr(ao$, "\") < B Then B = InStr(ao$, "\")
End If

If InStr(ao$, "|") > 0 Then
  If InStr(ao$, "|") < B Then B = InStr(ao$, "|")
End If

If InStr(ao$, "^") > 0 Then
  If InStr(ao$, "^") < B Then B = InStr(ao$, "^")
End If

If InStr(ao$, "@") > 0 Then
  If InStr(ao$, "@") < B Then B = InStr(ao$, "@")
End If

If B = 32767 Then c$ = Left(m$, 1): Last = 1: no$ = ao$: GoTo B
bo$ = Left(m$, B)
c$ = Left(bo$, 1)
no$ = Right(bo$, B - 1)

B:
If no$ = "V" Then n# = x: GoTo f:
If no$ = "W" Then n# = yh: GoTo f:


p$ = InsFun.Funname(no$)

If p$ <> "" Then
  If p$ <> "!" Then no$ = Right(no$, Len(no$) - Len(p$))
End If

If p$ = "LOG" Then GoTo d:

g:
If InStr(no$, "UN") = 1 Then n# = u#(Val(Right(no$, Len(no$) - 2))) Else n# = Val(no$)
If loga = 1 Then GoTo H:
If loga = 2 Then GoTo i:

d:
If p$ = "UN" Then n# = u#(Val(no$))

If Right(no$, 1) = "!" Then
If p$ <> "!" Then
   s = 1
   For i = 1 To n#
   s = s * i
   Next i
   n# = s
End If
End If

nd# = n# * PI# / 180

If p$ = "" Then GoTo f:
If p$ = "UN" Then GoTo f: '3.8新加的

Select Case p$
    Case "ABS"
    n# = Abs(n#)
    Case "SQR"
    n# = Sqr(n#)
    Case "INT"
    n# = Int(n#)
    Case "TRUNC"
    n# = Fix(n#)
    Case "LN"
    n# = Log(n#)
    Case "LNA"
    n# = Log(Abs(n#))
    Case "SIN"
     n# = Sin(n#)
    Case "COS"
     n# = Cos(n#)
    Case "TAN", "TG"
     n# = Tan(n#)
    Case "ARCSIN", "ASIN"
     n# = Atn(n# / Sqr(1 - n# ^ 2))
    Case "ARCCOS", "ACOS"
     n# = PI# / 2 - Atn(n# / Sqr(1 - n# ^ 2))
    Case "ARCTG", "ARCTAN", "ATN", "ATAN"
     n# = Atn(n#)
    Case "ARCCTG", "ACOT", "ARCCOT"
     n# = PI# / 2 - Atn(n#)
    Case "ARCSEC", "ASEC"
    n# = 1 / n#
     n# = PI# / 2 - Atn(n# / Sqr(1 - n# ^ 2))
    Case "ARCCSC", "ACSC"
    n# = 1 / n#
     n# = Atn(n# / Sqr(1 - n# ^ 2))
    Case "EP"
    n# = Exp(n#)
    Case "SGN", "SIGN"
    n# = Sgn(n#)
    Case "COT"
     n# = 1 / (Tan(n#))
    Case "SEC"
     n# = 1 / (Cos(n#))
    Case "CSC"
     n# = 1 / (Sin(n#))
    Case "LOG"
    If InStr(no$, "`") > 0 Then
      nao$ = Left(no$, InStr(no$, "`"))
      nno$ = Right(no$, Len(no$) - InStr(no$, "`"))
      no$ = nao$: loga = 1: GoTo g:
    Else: p$ = "LN": GoTo g:
    End If
H:      na# = n#
    no$ = nno$: loga = 2: GoTo g:
i:      nn# = n#
    n# = Log(nn#) / Log(na#)
    loga = 0: GoTo f:
    Case "LG"
    n# = Log(n#) / Log(10)
    Case "SH", "SINH"
    n# = (Exp(n#) - Exp(-n#)) / 2
    Case "CH", "COSH"
    n# = (Exp(n#) + Exp(-n#)) / 2
    Case "TH", "TANH"
    n# = (Exp(n#) - Exp(-n#)) / (Exp(n#) + Exp(-n#))
    Case "CTH", "COTH"
    n# = (Exp(n#) + Exp(-n#)) / (Exp(n#) - Exp(-n#))
    Case "SECH"
    n# = 2 / (Exp(n#) + Exp(-n#))
    Case "CSCH"
    n# = 2 / (Exp(n#) - Exp(-n#))
    Case "ARSH", "ASINH"
    n# = Log(n# + Sqr(n# ^ 2 + 1))
    Case "ARCH", "ACOSH"
    n# = Log(n# + Sqr(n# ^ 2 - 1)): Beep '+
    Case "ARTH", "ATANH"
    n# = (Log((n# + 1) / (1 - n#))) / 2
    Case "ARCTH", "ACOTH"
    n# = (Log((n# + 1) / (n# - 1))) / 2
    Case "ARSECH", "ASECH"
    n# = Arsech(n#) '+
    Case "ARCSCH", "ACSCH" '?
    n# = Log((Sgn(n#) * Sqr(n# ^ 2 + 1) + 1) / n#)
    Case "DMS"
    n# = Fix(n#) + (Fix((n# - Fix(n#)) * 100)) / 60 + (n# * 100 - Fix(n# * 100)) / 36
    Case "DEG"
    n# = Deg(n#)
    Case "ROUND"
    n# = Round(n#)
    Case "!"
        s = 1
        For i = 1 To n#
        s = s * i
        Next i
        n# = s
End Select



f:
nd# = 0
p$ = ""

Select Case c$
Case "+"
mo# = mo# + n#
Case "-"
mo# = mo# - n#
Case "*"
mo# = mo# * n#
Case "/"
mo# = mo# / n#
Case "\"
mo# = mo# \ n#
Case "|"
mo# = mo# Mod n#

Case "^"
If Pic.dwyy.Checked = False And mo# < 0 And Fix(Abs(n#)) <> Abs(n#) Then
  If Pic.dxyl.Checked = True Then mo# = -(-mo#) ^ n#: GoTo j:
  If Pic.ddyl.Checked = True Then mo# = (-mo#) ^ n#: GoTo j:
End If
If mo# = 0 And n# = 0 Then mo# = Log(-1) Else mo# = mo# ^ n#

Case "@"

If Pic.dwyy.Checked = False And n# < 0 Then
If Parity(mo#) = 0 Then
If Fix(Abs(1 / mo#)) <> Abs(1 / mo#) Then
  If Pic.dxyl.Checked = True Then mo# = -(-n#) ^ (1 / mo#): GoTo j:
  If Pic.ddyl.Checked = True Then mo# = (-n#) ^ (1 / mo#): GoTo j:
End If
End If
End If

If n# < 0 Then
If Parity(mo#) <> 0 Then
     If Parity(mo#) = 1 Then mo# = -(-n#) ^ (1 / mo#) Else mo# = Log(-1)
    
     GoTo j:
End If
End If

mo# = n# ^ (1 / mo#)

End Select






j:
If Last = 1 Then GoTo e:
m$ = Right(m$, a - B)
GoTo a

e:
Last = 0
un = un + 1
u#(un) = mo#
m$ = "UN" + Str$(un)
End Sub

Public Sub bet()
f$ = ""
f$ = m$

15:
a = Len(f$)
B = 32767
If InStr(f$, "*") > 0 Then B = InStr(f$, "*")
If InStr(f$, "/") > 0 Then
  If InStr(f$, "/") < B Then B = InStr(f$, "/")
End If
If InStr(f$, "\") > 0 Then
  If InStr(f$, "\") < B Then B = InStr(f$, "\")
End If
If InStr(f$, "|") > 0 Then
  If InStr(f$, "|") < B Then B = InStr(f$, "|")
End If
If B = 32767 Then GoTo 35

25:
m$ = Left(f$, B - 1)
If Len(m$) <> 0 Then Call alp
e$ = e$ + m$ + mid$(f$, B, 1)
f$ = Right(f$, a - B)
If InStr(f$, "^") > 0 Or InStr(f$, "@") > 0 Then GoTo 15 Else m$ = f$: GoTo 37

35:
m$ = f$
If Len(m$) <> 0 Then Call alp

37:
m$ = e$ + m$
e$ = ""
Call alp
End Sub
