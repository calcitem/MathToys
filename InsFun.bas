Attribute VB_Name = "InsFun"
Public Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Public Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, lpPoint As PointAPI) As Long
Type PointAPI
x As Long
y As Long
End Type

Public LpPoint1 As PointAPI
Public prmt As Boolean
Public Function Det2(a1 As Double, b1 As Double, a2 As Double, b2 As Double) As Double '2阶行列式计算
  Det2 = a1 * b2 - a2 * b1
End Function
Public Function Det3(a1 As Double, b1 As Double, c1 As Double, a2 As Double, b2 As Double, c2 As Double, a3 As Double, b3 As Double, c3 As Double)
  Det3 = a1 * Det2(b2, c2, b3, c3) - a2 * Det2(b1, c1, b3, c3) + a3 * Det2(b1, c1, b2, c2)
End Function
Public Function Det4(a1 As Double, b1 As Double, c1 As Double, d1 As Double, a2 As Double, b2 As Double, c2 As Double, d2 As Double, a3 As Double, b3 As Double, c3 As Double, d3 As Double, a4 As Double, b4 As Double, c4 As Double, d4 As Double)
  Det4 = a1 * Det3(b2, c2, d2, b3, c3, d3, b4, c4, d4) - a2 * Det3(b1, c1, d1, b3, c3, d3, b4, c4, d4) + a3 * Det3(b1, c1, d1, b2, c2, d2, b4, c4, d4) - a4 * Det3(b1, c1, d1, b2, c2, d2, b3, c3, d3)
End Function

Public Function Dms(n) '可能有问题如dms3.50
    'dms = Fix(n#) + (Fix((n# - Fix(n#)) * 100)) / 60 + (n# * 100 - Fix(n# * 100)) / 36
  sign = Sgn(n)
  n = Abs(n)
  d = Fix(n)
  m = (n - d) * 60
  t = (m - Fix(m)) * 60
  Dms = (d + (Fix(m) / 100) + t / 10000) * sign
End Function
Public Function Hypot(x As Double, y As Double) As Double
  Hypot = Sqr(x ^ 2 + y ^ 2)
End Function
Public Function Deg(n) '可能有问题如
 'd! = ((ms# - Fix(ms#)) * 60 - Fix((ms# - Fix(ms#)) * 60)) * 60
'If d! = 60 Then
 '  deg = Val(Str(Fix(ms#)) + Str(Fix((ms# - Fix(ms#)) * 60) + 1))
' Else
 '  deg = Val(Str(Fix(ms#)) + Str(Fix((ms# - Fix(ms#)) * 60)) + Str(dms!))
' End If
sign = Sgn(n)
n = Abs(n)
a = Fix(n)
b = Fix((n - a) * 100) / 60
Deg = a + b
n = n * 100
a = Fix(n)
b = ((n - a) * 100) / 3600
Deg = sign * (Deg + b)
End Function

Public Sub calcfun(DR, n#, p$)

End Sub

Public Function Funname(no$) As String
If InStr(no$, "UN") = 1 Then Funname = "UN": Exit Function

If Right(no$, 1) = "!" Then
  Funname = "!"
Else
  If Trim(no$) <> "" Then
    If Asc(no$) > 47 And Asc(no$) < 58 Then Funname = "": Exit Function
  End If
End If


If InStr(no$, "SINH") = 1 Then Funname = "SINH": Exit Function
If InStr(no$, "COSH") = 1 Then Funname = "COSH": Exit Function
If InStr(no$, "TANH") = 1 Then Funname = "TANH": Exit Function
If InStr(no$, "COTH") = 1 Then Funname = "COTH": Exit Function
If InStr(no$, "SECH") = 1 Then Funname = "SECH": Exit Function
If InStr(no$, "CSCH") = 1 Then Funname = "CSCH": Exit Function

If InStr(no$, "SIN") = 1 Then Funname = "SIN": Exit Function
If InStr(no$, "COS") = 1 Then Funname = "COS": Exit Function
If InStr(no$, "SQR") = 1 Then Funname = "SQR": Exit Function
If InStr(no$, "EP") = 1 Then Funname = "EP": Exit Function
If InStr(no$, "LNA") = 1 Then Funname = "LNA": Exit Function
If InStr(no$, "LN") = 1 Then Funname = "LN": Exit Function
If InStr(no$, "TAN") = 1 Then Funname = "TAN": Exit Function
If InStr(no$, "COT") = 1 Or InStr(no$, "CTG") = 1 Then Funname = "COT": Exit Function
If InStr(no$, "LG") = 1 Then Funname = "LG": Exit Function
If InStr(no$, "SEC") = 1 Then Funname = "SEC": Exit Function
If InStr(no$, "CSC") = 1 Then Funname = "CSC": Exit Function
If InStr(no$, "LOG") = 1 Then Funname = "LOG": Exit Function
If InStr(no$, "ABS") = 1 Then Funname = "ABS": Exit Function
If InStr(no$, "ASECH") = 1 Then Funname = "ASECH": Exit Function
If InStr(no$, "ACSCH") = 1 Then Funname = "ACSCH": Exit Function
If InStr(no$, "ASINH") = 1 Then Funname = "ASINH": Exit Function
If InStr(no$, "ACOSH") = 1 Then Funname = "ACOSH": Exit Function
If InStr(no$, "ATANH") = 1 Then Funname = "ATANH": Exit Function
If InStr(no$, "ACOTH") = 1 Then Funname = "ACOTH": Exit Function
If InStr(no$, "ASIN") = 1 Then Funname = "ASIN": Exit Function
If InStr(no$, "ACOS") = 1 Then Funname = "ACOS": Exit Function
If InStr(no$, "ATAN") = 1 Then Funname = "ATAN": Exit Function
If InStr(no$, "ACOT") = 1 Then Funname = "ACOT": Exit Function
If InStr(no$, "ASEC") = 1 Then Funname = "ASEC": Exit Function
If InStr(no$, "ACSC") = 1 Then Funname = "ACSC": Exit Function
If InStr(no$, "ARCSIN") = 1 Then Funname = "ARCSIN": Exit Function
If InStr(no$, "ARCCOS") = 1 Then Funname = "ARCCOS": Exit Function
If InStr(no$, "TG") = 1 Then Funname = "TG": Exit Function
If InStr(no$, "INT") = 1 Then Funname = "INT": Exit Function
If InStr(no$, "FIX") = 1 Then Funname = "FIX": Exit Function
If InStr(no$, "ATN") = 1 Then Funname = "ATN": Exit Function
If InStr(no$, "ARCTAN") = 1 Then Funname = "ARCTAN": Exit Function
If InStr(no$, "ARCTG") = 1 Then Funname = "ARCTG": Exit Function
If InStr(no$, "ARCCOT") = 1 Or InStr(no$, "ARCCTG") = 1 Then Funname = "ARCCTG": Exit Function
If InStr(no$, "ARCSEC") = 1 Then Funname = "ARCSEC": Exit Function
If InStr(no$, "ARCCSC") = 1 Then Funname = "ARCCSC": Exit Function
If InStr(no$, "SH") = 1 Then Funname = "SH": Exit Function
If InStr(no$, "CH") = 1 Then Funname = "CH": Exit Function
If InStr(no$, "TH") = 1 Then Funname = "TH": Exit Function
If InStr(no$, "CTH") = 1 Then Funname = "CTH": Exit Function
If InStr(no$, "ARSH") = 1 Then Funname = "ARSH": Exit Function
If InStr(no$, "ARCH") = 1 Then Funname = "ARCH": Exit Function
If InStr(no$, "ARTH") = 1 Then Funname = "ARTH": Exit Function
If InStr(no$, "ARCTH") = 1 Then Funname = "ARCTH": Exit Function

If InStr(no$, "ARSECH") = 1 Then Funname = "ARSECH": Exit Function
If InStr(no$, "ARCSCH") = 1 Then Funname = "ARCSCH": Exit Function

If InStr(no$, "SGN") = 1 Then Funname = "SGN": Exit Function
If InStr(no$, "DMS") = 1 Then Funname = "DMS": Exit Function
If InStr(no$, "DEG") = 1 Then Funname = "DEG": Exit Function
If InStr(no$, "TRUNC") = 1 Then Funname = "TRUNC": Exit Function
If InStr(no$, "SIGN") = 1 Then Funname = "SIGN": Exit Function
If InStr(no$, "ROUND") = 1 Then Funname = "ROUND": Exit Function
'If InStr(no$, "SQRT") = 1 Then Funname = "SQRT": Exit Function
End Function

Public Function translate(ao$) As String

ao$ = UCase(ao$)

ao$ = Multinomial(ao$) '>>>>>>>>>>>>>>>>>>>>>>>

While InStr(ao$, "[") > 0
    Mid$(ao$, InStr(ao$, "["), 1) = "("
Wend

While InStr(ao$, "]") > 0
    Mid$(ao$, InStr(ao$, "]"), 1) = ")"
Wend



Do Until InStr(ao$, "(E)") = 0
    ao$ = Left(ao$, InStr(ao$, "(E)") - 1) + "(EXP1)" + Right(ao$, Len(ao$) - InStr(ao$, "(E)") - 2)
Loop


Do Until InStr(ao$, "E+") = 0
    epl$ = Left(ao$, InStr(ao$, "E+") - 1)
    epr$ = Right(ao$, Len(ao$) - InStr(ao$, "E+") - 1)
    ao$ = epl$ + "*10^" + epr$
Loop

Do Until InStr(ao$, "E-") = 0
    eml$ = Left(ao$, InStr(ao$, "E-") - 1)
    emr$ = Right(ao$, Len(ao$) - InStr(ao$, "E-"))
    If InStr(emr$, "-0") = 1 Then
    emr$ = "-" + Right(emr$, Len(emr$) - 2)
    End If
    emr$ = Left(emr$, Len(Str((Val(emr$))))) + ")" + Right(emr$, Len(emr$) - Len(Left(emr$, Len(Str(Val(emr$))))))
    ao$ = eml$ + "*10^(" + emr$
Loop



Do Until InStr(ao$, "MOD") = 0
    ao$ = Left(ao$, InStr(ao$, "MOD") - 1) + "|" + Right(ao$, Len(ao$) - InStr(ao$, "MOD") - 2)
Loop

Do Until InStr(ao$, "FIX") = 0
    ao$ = Left(ao$, InStr(ao$, "FIX") - 1) + "TRUNC" + Right(ao$, Len(ao$) - InStr(ao$, "FIX") - 2)
Loop
Do Until InStr(ao$, "PI(") = 0
    ao$ = Left(ao$, InStr(ao$, "PI(") - 1) + "PI*(" + Right(ao$, Len(ao$) - InStr(ao$, "PI(") - 2)
Loop

'Do Until InStr(ao$, "PI") = 0
'    ao$ = Left(ao$, InStr(ao$, "PI") - 1) + "(Pi)" + Right(ao$, Len(ao$) - InStr(ao$, "PI") - 1)
'Loop

Do Until InStr(ao$, "%") = 0
    ao$ = Left(ao$, InStr(ao$, "%") - 1) + "/100" + Right(ao$, Len(ao$) - InStr(ao$, "%"))
Loop

Do Until InStr(ao$, ")(") = 0
    ao$ = Left(ao$, InStr(ao$, ")(") - 1) + ")*(" + Right(ao$, Len(ao$) - InStr(ao$, ")(") - 1)
Loop
Do Until InStr(ao$, "0(") = 0
    ao$ = Left(ao$, InStr(ao$, "0(") - 1) + "0*(" + Right(ao$, Len(ao$) - InStr(ao$, "0(") - 1)
Loop
Do Until InStr(ao$, "1(") = 0
    ao$ = Left(ao$, InStr(ao$, "1(") - 1) + "1*(" + Right(ao$, Len(ao$) - InStr(ao$, "1(") - 1)
Loop
Do Until InStr(ao$, "2(") = 0
    ao$ = Left(ao$, InStr(ao$, "2(") - 1) + "2*(" + Right(ao$, Len(ao$) - InStr(ao$, "2(") - 1)
Loop
Do Until InStr(ao$, "3(") = 0
    ao$ = Left(ao$, InStr(ao$, "3(") - 1) + "3*(" + Right(ao$, Len(ao$) - InStr(ao$, "3(") - 1)
Loop
Do Until InStr(ao$, "4(") = 0
    ao$ = Left(ao$, InStr(ao$, "4(") - 1) + "4*(" + Right(ao$, Len(ao$) - InStr(ao$, "4(") - 1)
Loop
Do Until InStr(ao$, "5(") = 0
    ao$ = Left(ao$, InStr(ao$, "5(") - 1) + "5*(" + Right(ao$, Len(ao$) - InStr(ao$, "5(") - 1)
Loop
Do Until InStr(ao$, "6(") = 0
    ao$ = Left(ao$, InStr(ao$, "6(") - 1) + "6*(" + Right(ao$, Len(ao$) - InStr(ao$, "6(") - 1)
Loop
Do Until InStr(ao$, "7(") = 0
    ao$ = Left(ao$, InStr(ao$, "7(") - 1) + "7*(" + Right(ao$, Len(ao$) - InStr(ao$, "7(") - 1)
Loop
Do Until InStr(ao$, "8(") = 0
    ao$ = Left(ao$, InStr(ao$, "8(") - 1) + "8*(" + Right(ao$, Len(ao$) - InStr(ao$, "8(") - 1)
Loop
Do Until InStr(ao$, "9(") = 0
    ao$ = Left(ao$, InStr(ao$, "9(") - 1) + "9*(" + Right(ao$, Len(ao$) - InStr(ao$, "9(") - 1)
Loop
Do Until InStr(ao$, ")0") = 0
    ao$ = Left(ao$, InStr(ao$, ")0") - 1) + ")*0" + Right(ao$, Len(ao$) - InStr(ao$, ")0") - 1)
Loop
Do Until InStr(ao$, ")1") = 0
    ao$ = Left(ao$, InStr(ao$, ")1") - 1) + ")*1" + Right(ao$, Len(ao$) - InStr(ao$, ")1") - 1)
Loop
Do Until InStr(ao$, ")2") = 0
    ao$ = Left(ao$, InStr(ao$, ")2") - 1) + ")*2" + Right(ao$, Len(ao$) - InStr(ao$, ")2") - 1)
Loop
Do Until InStr(ao$, ")3") = 0
    ao$ = Left(ao$, InStr(ao$, ")3") - 1) + ")*3" + Right(ao$, Len(ao$) - InStr(ao$, ")3") - 1)
Loop
Do Until InStr(ao$, ")4") = 0
    ao$ = Left(ao$, InStr(ao$, ")4") - 1) + ")*4" + Right(ao$, Len(ao$) - InStr(ao$, ")4") - 1)
Loop
Do Until InStr(ao$, ")5") = 0
    ao$ = Left(ao$, InStr(ao$, ")5") - 1) + ")*5" + Right(ao$, Len(ao$) - InStr(ao$, ")5") - 1)
Loop
Do Until InStr(ao$, ")6") = 0
    ao$ = Left(ao$, InStr(ao$, ")6") - 1) + ")*6" + Right(ao$, Len(ao$) - InStr(ao$, ")6") - 1)
Loop
Do Until InStr(ao$, ")7") = 0
    ao$ = Left(ao$, InStr(ao$, ")7") - 1) + ")*7" + Right(ao$, Len(ao$) - InStr(ao$, ")7") - 1)
Loop
Do Until InStr(ao$, ")8") = 0
    ao$ = Left(ao$, InStr(ao$, ")8") - 1) + ")*8" + Right(ao$, Len(ao$) - InStr(ao$, ")8") - 1)
Loop
Do Until InStr(ao$, ")9") = 0
    ao$ = Left(ao$, InStr(ao$, ")9") - 1) + ")*9" + Right(ao$, Len(ao$) - InStr(ao$, ")9") - 1)
Loop
Do Until InStr(ao$, ")A") = 0
    ao$ = Left(ao$, InStr(ao$, ")A") - 1) + ")*A" + Right(ao$, Len(ao$) - InStr(ao$, ")A") - 1)
Loop
Do Until InStr(ao$, ")S") = 0
    ao$ = Left(ao$, InStr(ao$, ")S") - 1) + ")*S" + Right(ao$, Len(ao$) - InStr(ao$, ")S") - 1)
Loop
Do Until InStr(ao$, ")C") = 0
    ao$ = Left(ao$, InStr(ao$, ")C") - 1) + ")*C" + Right(ao$, Len(ao$) - InStr(ao$, ")C") - 1)
Loop
Do Until InStr(ao$, ")T") = 0
    ao$ = Left(ao$, InStr(ao$, ")T") - 1) + ")*T" + Right(ao$, Len(ao$) - InStr(ao$, ")T") - 1)
Loop
Do Until InStr(ao$, ")M") = 0
    ao$ = Left(ao$, InStr(ao$, ")M") - 1) + ")*M" + Right(ao$, Len(ao$) - InStr(ao$, ")M") - 1)
Loop
Do Until InStr(ao$, ")P") = 0
    ao$ = Left(ao$, InStr(ao$, ")P") - 1) + ")*P" + Right(ao$, Len(ao$) - InStr(ao$, ")P") - 1)
Loop
Do Until InStr(ao$, ")D") = 0
    ao$ = Left(ao$, InStr(ao$, ")D") - 1) + ")*D" + Right(ao$, Len(ao$) - InStr(ao$, ")D") - 1)
Loop
Do Until InStr(ao$, ")F") = 0
    ao$ = Left(ao$, InStr(ao$, ")F") - 1) + ")*F" + Right(ao$, Len(ao$) - InStr(ao$, ")F") - 1)
Loop
Do Until InStr(ao$, ")L") = 0
    ao$ = Left(ao$, InStr(ao$, ")L") - 1) + ")*L" + Right(ao$, Len(ao$) - InStr(ao$, ")L") - 1)
Loop
Do Until InStr(ao$, ")E") = 0
    ao$ = Left(ao$, InStr(ao$, ")E") - 1) + ")*E" + Right(ao$, Len(ao$) - InStr(ao$, ")E") - 1)
Loop

Do Until InStr(ao$, ")I") = 0
    ao$ = Left(ao$, InStr(ao$, ")I") - 1) + ")*I" + Right(ao$, Len(ao$) - InStr(ao$, ")I") - 1)
Loop


Do Until InStr(ao$, "1A") = 0
    ao$ = Left(ao$, InStr(ao$, "1A") - 1) + "1*A" + Right(ao$, Len(ao$) - InStr(ao$, "1A") - 1)
Loop
Do Until InStr(ao$, "2A") = 0
    ao$ = Left(ao$, InStr(ao$, "2A") - 1) + "2*A" + Right(ao$, Len(ao$) - InStr(ao$, "2A") - 1)
Loop
Do Until InStr(ao$, "3A") = 0
    ao$ = Left(ao$, InStr(ao$, "3A") - 1) + "3*A" + Right(ao$, Len(ao$) - InStr(ao$, "3A") - 1)
Loop
Do Until InStr(ao$, "4A") = 0
    ao$ = Left(ao$, InStr(ao$, "4A") - 1) + "4*A" + Right(ao$, Len(ao$) - InStr(ao$, "4A") - 1)
Loop
Do Until InStr(ao$, "5A") = 0
    ao$ = Left(ao$, InStr(ao$, "5A") - 1) + "5*A" + Right(ao$, Len(ao$) - InStr(ao$, "5A") - 1)
Loop
Do Until InStr(ao$, "6A") = 0
    ao$ = Left(ao$, InStr(ao$, "6A") - 1) + "6*A" + Right(ao$, Len(ao$) - InStr(ao$, "6A") - 1)
Loop
Do Until InStr(ao$, "7A") = 0
    ao$ = Left(ao$, InStr(ao$, "7A") - 1) + "7*A" + Right(ao$, Len(ao$) - InStr(ao$, "7A") - 1)
Loop
Do Until InStr(ao$, "8A") = 0
    ao$ = Left(ao$, InStr(ao$, "8A") - 1) + "8*A" + Right(ao$, Len(ao$) - InStr(ao$, "8A") - 1)
Loop
Do Until InStr(ao$, "9A") = 0
    ao$ = Left(ao$, InStr(ao$, "9A") - 1) + "9*A" + Right(ao$, Len(ao$) - InStr(ao$, "9A") - 1)
Loop
Do Until InStr(ao$, "0A") = 0
    ao$ = Left(ao$, InStr(ao$, "0A") - 1) + "0*A" + Right(ao$, Len(ao$) - InStr(ao$, "0A") - 1)
Loop

Do Until InStr(ao$, "1C") = 0
    ao$ = Left(ao$, InStr(ao$, "1C") - 1) + "1*C" + Right(ao$, Len(ao$) - InStr(ao$, "1C") - 1)
Loop
Do Until InStr(ao$, "2C") = 0
    ao$ = Left(ao$, InStr(ao$, "2C") - 1) + "2*C" + Right(ao$, Len(ao$) - InStr(ao$, "2C") - 1)
Loop
Do Until InStr(ao$, "3C") = 0
    ao$ = Left(ao$, InStr(ao$, "3C") - 1) + "3*C" + Right(ao$, Len(ao$) - InStr(ao$, "3C") - 1)
Loop
Do Until InStr(ao$, "4C") = 0
    ao$ = Left(ao$, InStr(ao$, "4C") - 1) + "4*C" + Right(ao$, Len(ao$) - InStr(ao$, "4C") - 1)
Loop
Do Until InStr(ao$, "5C") = 0
    ao$ = Left(ao$, InStr(ao$, "5C") - 1) + "5*C" + Right(ao$, Len(ao$) - InStr(ao$, "5C") - 1)
Loop
Do Until InStr(ao$, "6C") = 0
    ao$ = Left(ao$, InStr(ao$, "6C") - 1) + "6*C" + Right(ao$, Len(ao$) - InStr(ao$, "6C") - 1)
Loop
Do Until InStr(ao$, "7C") = 0
    ao$ = Left(ao$, InStr(ao$, "7C") - 1) + "7*C" + Right(ao$, Len(ao$) - InStr(ao$, "7C") - 1)
Loop
Do Until InStr(ao$, "8C") = 0
    ao$ = Left(ao$, InStr(ao$, "8C") - 1) + "8*C" + Right(ao$, Len(ao$) - InStr(ao$, "8C") - 1)
Loop
Do Until InStr(ao$, "9C") = 0
    ao$ = Left(ao$, InStr(ao$, "9C") - 1) + "9*C" + Right(ao$, Len(ao$) - InStr(ao$, "9C") - 1)
Loop
Do Until InStr(ao$, "0C") = 0
    ao$ = Left(ao$, InStr(ao$, "0C") - 1) + "0*C" + Right(ao$, Len(ao$) - InStr(ao$, "0C") - 1)
Loop

Do Until InStr(ao$, "1F") = 0
    ao$ = Left(ao$, InStr(ao$, "1F") - 1) + "1*F" + Right(ao$, Len(ao$) - InStr(ao$, "1F") - 1)
Loop
Do Until InStr(ao$, "2F") = 0
    ao$ = Left(ao$, InStr(ao$, "2F") - 1) + "2*F" + Right(ao$, Len(ao$) - InStr(ao$, "2F") - 1)
Loop
Do Until InStr(ao$, "3F") = 0
    ao$ = Left(ao$, InStr(ao$, "3F") - 1) + "3*F" + Right(ao$, Len(ao$) - InStr(ao$, "3F") - 1)
Loop
Do Until InStr(ao$, "4F") = 0
    ao$ = Left(ao$, InStr(ao$, "4F") - 1) + "4*F" + Right(ao$, Len(ao$) - InStr(ao$, "4F") - 1)
Loop
Do Until InStr(ao$, "5F") = 0
    ao$ = Left(ao$, InStr(ao$, "5F") - 1) + "5*F" + Right(ao$, Len(ao$) - InStr(ao$, "5F") - 1)
Loop
Do Until InStr(ao$, "6F") = 0
    ao$ = Left(ao$, InStr(ao$, "6F") - 1) + "6*F" + Right(ao$, Len(ao$) - InStr(ao$, "6F") - 1)
Loop
Do Until InStr(ao$, "7F") = 0
    ao$ = Left(ao$, InStr(ao$, "7F") - 1) + "7*F" + Right(ao$, Len(ao$) - InStr(ao$, "7F") - 1)
Loop
Do Until InStr(ao$, "8F") = 0
    ao$ = Left(ao$, InStr(ao$, "8F") - 1) + "8*F" + Right(ao$, Len(ao$) - InStr(ao$, "8F") - 1)
Loop
Do Until InStr(ao$, "9F") = 0
    ao$ = Left(ao$, InStr(ao$, "9F") - 1) + "9*F" + Right(ao$, Len(ao$) - InStr(ao$, "9F") - 1)
Loop
Do Until InStr(ao$, "0F") = 0
    ao$ = Left(ao$, InStr(ao$, "0F") - 1) + "0*F" + Right(ao$, Len(ao$) - InStr(ao$, "0F") - 1)
Loop

Do Until InStr(ao$, "1I") = 0
    ao$ = Left(ao$, InStr(ao$, "1I") - 1) + "1*I" + Right(ao$, Len(ao$) - InStr(ao$, "1I") - 1)
Loop
Do Until InStr(ao$, "2I") = 0
    ao$ = Left(ao$, InStr(ao$, "2I") - 1) + "2*I" + Right(ao$, Len(ao$) - InStr(ao$, "2I") - 1)
Loop
Do Until InStr(ao$, "3I") = 0
    ao$ = Left(ao$, InStr(ao$, "3I") - 1) + "3*I" + Right(ao$, Len(ao$) - InStr(ao$, "3I") - 1)
Loop
Do Until InStr(ao$, "4I") = 0
    ao$ = Left(ao$, InStr(ao$, "4I") - 1) + "4*I" + Right(ao$, Len(ao$) - InStr(ao$, "4I") - 1)
Loop
Do Until InStr(ao$, "5I") = 0
    ao$ = Left(ao$, InStr(ao$, "5I") - 1) + "5*I" + Right(ao$, Len(ao$) - InStr(ao$, "5I") - 1)
Loop
Do Until InStr(ao$, "6I") = 0
    ao$ = Left(ao$, InStr(ao$, "6I") - 1) + "6*I" + Right(ao$, Len(ao$) - InStr(ao$, "6I") - 1)
Loop
Do Until InStr(ao$, "7I") = 0
    ao$ = Left(ao$, InStr(ao$, "7I") - 1) + "7*I" + Right(ao$, Len(ao$) - InStr(ao$, "7I") - 1)
Loop
Do Until InStr(ao$, "8I") = 0
    ao$ = Left(ao$, InStr(ao$, "8I") - 1) + "8*I" + Right(ao$, Len(ao$) - InStr(ao$, "8I") - 1)
Loop
Do Until InStr(ao$, "9I") = 0
    ao$ = Left(ao$, InStr(ao$, "9I") - 1) + "9*I" + Right(ao$, Len(ao$) - InStr(ao$, "9I") - 1)
Loop
Do Until InStr(ao$, "0I") = 0
    ao$ = Left(ao$, InStr(ao$, "0I") - 1) + "0*I" + Right(ao$, Len(ao$) - InStr(ao$, "0I") - 1)
Loop

Do Until InStr(ao$, "1L") = 0
    ao$ = Left(ao$, InStr(ao$, "1L") - 1) + "1*L" + Right(ao$, Len(ao$) - InStr(ao$, "1L") - 1)
Loop
Do Until InStr(ao$, "2L") = 0
    ao$ = Left(ao$, InStr(ao$, "2L") - 1) + "2*L" + Right(ao$, Len(ao$) - InStr(ao$, "2L") - 1)
Loop
Do Until InStr(ao$, "3L") = 0
    ao$ = Left(ao$, InStr(ao$, "3L") - 1) + "3*L" + Right(ao$, Len(ao$) - InStr(ao$, "3L") - 1)
Loop
Do Until InStr(ao$, "4L") = 0
    ao$ = Left(ao$, InStr(ao$, "4L") - 1) + "4*L" + Right(ao$, Len(ao$) - InStr(ao$, "4L") - 1)
Loop
Do Until InStr(ao$, "5L") = 0
    ao$ = Left(ao$, InStr(ao$, "5L") - 1) + "5*L" + Right(ao$, Len(ao$) - InStr(ao$, "5L") - 1)
Loop
Do Until InStr(ao$, "6L") = 0
    ao$ = Left(ao$, InStr(ao$, "6L") - 1) + "6*L" + Right(ao$, Len(ao$) - InStr(ao$, "6L") - 1)
Loop
Do Until InStr(ao$, "7L") = 0
    ao$ = Left(ao$, InStr(ao$, "7L") - 1) + "7*L" + Right(ao$, Len(ao$) - InStr(ao$, "7L") - 1)
Loop
Do Until InStr(ao$, "8L") = 0
    ao$ = Left(ao$, InStr(ao$, "8L") - 1) + "8*L" + Right(ao$, Len(ao$) - InStr(ao$, "8L") - 1)
Loop
Do Until InStr(ao$, "9L") = 0
    ao$ = Left(ao$, InStr(ao$, "9L") - 1) + "9*L" + Right(ao$, Len(ao$) - InStr(ao$, "9L") - 1)
Loop
Do Until InStr(ao$, "0L") = 0
    ao$ = Left(ao$, InStr(ao$, "0L") - 1) + "0*L" + Right(ao$, Len(ao$) - InStr(ao$, "0L") - 1)
Loop

Do Until InStr(ao$, "1R") = 0
    ao$ = Left(ao$, InStr(ao$, "1R") - 1) + "1*R" + Right(ao$, Len(ao$) - InStr(ao$, "1R") - 1)
Loop
Do Until InStr(ao$, "2R") = 0
    ao$ = Left(ao$, InStr(ao$, "2R") - 1) + "2*R" + Right(ao$, Len(ao$) - InStr(ao$, "2R") - 1)
Loop
Do Until InStr(ao$, "3R") = 0
    ao$ = Left(ao$, InStr(ao$, "3R") - 1) + "3*R" + Right(ao$, Len(ao$) - InStr(ao$, "3R") - 1)
Loop
Do Until InStr(ao$, "4R") = 0
    ao$ = Left(ao$, InStr(ao$, "4R") - 1) + "4*R" + Right(ao$, Len(ao$) - InStr(ao$, "4R") - 1)
Loop
Do Until InStr(ao$, "5R") = 0
    ao$ = Left(ao$, InStr(ao$, "5R") - 1) + "5*R" + Right(ao$, Len(ao$) - InStr(ao$, "5R") - 1)
Loop
Do Until InStr(ao$, "6R") = 0
    ao$ = Left(ao$, InStr(ao$, "6R") - 1) + "6*R" + Right(ao$, Len(ao$) - InStr(ao$, "6R") - 1)
Loop
Do Until InStr(ao$, "7R") = 0
    ao$ = Left(ao$, InStr(ao$, "7R") - 1) + "7*R" + Right(ao$, Len(ao$) - InStr(ao$, "7R") - 1)
Loop
Do Until InStr(ao$, "8R") = 0
    ao$ = Left(ao$, InStr(ao$, "8R") - 1) + "8*R" + Right(ao$, Len(ao$) - InStr(ao$, "8R") - 1)
Loop
Do Until InStr(ao$, "9R") = 0
    ao$ = Left(ao$, InStr(ao$, "9R") - 1) + "9*R" + Right(ao$, Len(ao$) - InStr(ao$, "9R") - 1)
Loop
Do Until InStr(ao$, "0R") = 0
    ao$ = Left(ao$, InStr(ao$, "0R") - 1) + "0*R" + Right(ao$, Len(ao$) - InStr(ao$, "0R") - 1)
Loop

Do Until InStr(ao$, "1S") = 0
    ao$ = Left(ao$, InStr(ao$, "1S") - 1) + "1*S" + Right(ao$, Len(ao$) - InStr(ao$, "1S") - 1)
Loop
Do Until InStr(ao$, "2S") = 0
    ao$ = Left(ao$, InStr(ao$, "2S") - 1) + "2*S" + Right(ao$, Len(ao$) - InStr(ao$, "2S") - 1)
Loop
Do Until InStr(ao$, "3S") = 0
    ao$ = Left(ao$, InStr(ao$, "3S") - 1) + "3*S" + Right(ao$, Len(ao$) - InStr(ao$, "3S") - 1)
Loop
Do Until InStr(ao$, "4S") = 0
    ao$ = Left(ao$, InStr(ao$, "4S") - 1) + "4*S" + Right(ao$, Len(ao$) - InStr(ao$, "4S") - 1)
Loop
Do Until InStr(ao$, "5S") = 0
    ao$ = Left(ao$, InStr(ao$, "5S") - 1) + "5*S" + Right(ao$, Len(ao$) - InStr(ao$, "5S") - 1)
Loop
Do Until InStr(ao$, "6S") = 0
    ao$ = Left(ao$, InStr(ao$, "6S") - 1) + "6*S" + Right(ao$, Len(ao$) - InStr(ao$, "6S") - 1)
Loop
Do Until InStr(ao$, "7S") = 0
    ao$ = Left(ao$, InStr(ao$, "7S") - 1) + "7*S" + Right(ao$, Len(ao$) - InStr(ao$, "7S") - 1)
Loop
Do Until InStr(ao$, "8S") = 0
    ao$ = Left(ao$, InStr(ao$, "8S") - 1) + "8*S" + Right(ao$, Len(ao$) - InStr(ao$, "8S") - 1)
Loop
Do Until InStr(ao$, "9S") = 0
    ao$ = Left(ao$, InStr(ao$, "9S") - 1) + "9*S" + Right(ao$, Len(ao$) - InStr(ao$, "9S") - 1)
Loop
Do Until InStr(ao$, "0S") = 0
    ao$ = Left(ao$, InStr(ao$, "0S") - 1) + "0*S" + Right(ao$, Len(ao$) - InStr(ao$, "0S") - 1)
Loop

Do Until InStr(ao$, "1T") = 0
    ao$ = Left(ao$, InStr(ao$, "1T") - 1) + "1*T" + Right(ao$, Len(ao$) - InStr(ao$, "1T") - 1)
Loop
Do Until InStr(ao$, "2T") = 0
    ao$ = Left(ao$, InStr(ao$, "2T") - 1) + "2*T" + Right(ao$, Len(ao$) - InStr(ao$, "2T") - 1)
Loop
Do Until InStr(ao$, "3T") = 0
    ao$ = Left(ao$, InStr(ao$, "3T") - 1) + "3*T" + Right(ao$, Len(ao$) - InStr(ao$, "3T") - 1)
Loop
Do Until InStr(ao$, "4T") = 0
    ao$ = Left(ao$, InStr(ao$, "4T") - 1) + "4*T" + Right(ao$, Len(ao$) - InStr(ao$, "4T") - 1)
Loop
Do Until InStr(ao$, "5T") = 0
    ao$ = Left(ao$, InStr(ao$, "5T") - 1) + "5*T" + Right(ao$, Len(ao$) - InStr(ao$, "5T") - 1)
Loop
Do Until InStr(ao$, "6T") = 0
    ao$ = Left(ao$, InStr(ao$, "6T") - 1) + "6*T" + Right(ao$, Len(ao$) - InStr(ao$, "6T") - 1)
Loop
Do Until InStr(ao$, "7T") = 0
    ao$ = Left(ao$, InStr(ao$, "7T") - 1) + "7*T" + Right(ao$, Len(ao$) - InStr(ao$, "7T") - 1)
Loop
Do Until InStr(ao$, "8T") = 0
    ao$ = Left(ao$, InStr(ao$, "8T") - 1) + "8*T" + Right(ao$, Len(ao$) - InStr(ao$, "8T") - 1)
Loop
Do Until InStr(ao$, "9T") = 0
    ao$ = Left(ao$, InStr(ao$, "9T") - 1) + "9*T" + Right(ao$, Len(ao$) - InStr(ao$, "9T") - 1)
Loop
Do Until InStr(ao$, "0T") = 0
    ao$ = Left(ao$, InStr(ao$, "0T") - 1) + "0*T" + Right(ao$, Len(ao$) - InStr(ao$, "0T") - 1)
Loop

Do Until InStr(ao$, "1P") = 0
    ao$ = Left(ao$, InStr(ao$, "1P") - 1) + "1*P" + Right(ao$, Len(ao$) - InStr(ao$, "1P") - 1)
Loop
Do Until InStr(ao$, "2P") = 0
    ao$ = Left(ao$, InStr(ao$, "2P") - 1) + "2*P" + Right(ao$, Len(ao$) - InStr(ao$, "2P") - 1)
Loop
Do Until InStr(ao$, "3P") = 0
    ao$ = Left(ao$, InStr(ao$, "3P") - 1) + "3*P" + Right(ao$, Len(ao$) - InStr(ao$, "3P") - 1)
Loop
Do Until InStr(ao$, "4P") = 0
    ao$ = Left(ao$, InStr(ao$, "4P") - 1) + "4*P" + Right(ao$, Len(ao$) - InStr(ao$, "4P") - 1)
Loop
Do Until InStr(ao$, "5P") = 0
    ao$ = Left(ao$, InStr(ao$, "5P") - 1) + "5*P" + Right(ao$, Len(ao$) - InStr(ao$, "5P") - 1)
Loop
Do Until InStr(ao$, "6P") = 0
    ao$ = Left(ao$, InStr(ao$, "6P") - 1) + "6*P" + Right(ao$, Len(ao$) - InStr(ao$, "6P") - 1)
Loop
Do Until InStr(ao$, "7P") = 0
    ao$ = Left(ao$, InStr(ao$, "7P") - 1) + "7*P" + Right(ao$, Len(ao$) - InStr(ao$, "7P") - 1)
Loop
Do Until InStr(ao$, "8P") = 0
    ao$ = Left(ao$, InStr(ao$, "8P") - 1) + "8*P" + Right(ao$, Len(ao$) - InStr(ao$, "8P") - 1)
Loop
Do Until InStr(ao$, "9P") = 0
    ao$ = Left(ao$, InStr(ao$, "9P") - 1) + "9*P" + Right(ao$, Len(ao$) - InStr(ao$, "9P") - 1)
Loop
Do Until InStr(ao$, "0P") = 0
    ao$ = Left(ao$, InStr(ao$, "0P") - 1) + "0*P" + Right(ao$, Len(ao$) - InStr(ao$, "0P") - 1)
Loop

Do Until InStr(ao$, "1D") = 0
    ao$ = Left(ao$, InStr(ao$, "1D") - 1) + "1*D" + Right(ao$, Len(ao$) - InStr(ao$, "1D") - 1)
Loop
Do Until InStr(ao$, "2D") = 0
    ao$ = Left(ao$, InStr(ao$, "2D") - 1) + "2*D" + Right(ao$, Len(ao$) - InStr(ao$, "2D") - 1)
Loop
Do Until InStr(ao$, "3D") = 0
    ao$ = Left(ao$, InStr(ao$, "3D") - 1) + "3*D" + Right(ao$, Len(ao$) - InStr(ao$, "3D") - 1)
Loop
Do Until InStr(ao$, "4D") = 0
    ao$ = Left(ao$, InStr(ao$, "4D") - 1) + "4*D" + Right(ao$, Len(ao$) - InStr(ao$, "4D") - 1)
Loop
Do Until InStr(ao$, "5D") = 0
    ao$ = Left(ao$, InStr(ao$, "5D") - 1) + "5*D" + Right(ao$, Len(ao$) - InStr(ao$, "5D") - 1)
Loop
Do Until InStr(ao$, "6D") = 0
    ao$ = Left(ao$, InStr(ao$, "6D") - 1) + "6*D" + Right(ao$, Len(ao$) - InStr(ao$, "6D") - 1)
Loop
Do Until InStr(ao$, "7D") = 0
    ao$ = Left(ao$, InStr(ao$, "7D") - 1) + "7*D" + Right(ao$, Len(ao$) - InStr(ao$, "7D") - 1)
Loop
Do Until InStr(ao$, "8D") = 0
    ao$ = Left(ao$, InStr(ao$, "8D") - 1) + "8*D" + Right(ao$, Len(ao$) - InStr(ao$, "8D") - 1)
Loop
Do Until InStr(ao$, "9D") = 0
    ao$ = Left(ao$, InStr(ao$, "9D") - 1) + "9*D" + Right(ao$, Len(ao$) - InStr(ao$, "9D") - 1)
Loop
Do Until InStr(ao$, "0D") = 0
    ao$ = Left(ao$, InStr(ao$, "0D") - 1) + "0*D" + Right(ao$, Len(ao$) - InStr(ao$, "0D") - 1)
Loop

Do Until InStr(ao$, "1EXP") = 0
    ao$ = Left(ao$, InStr(ao$, "1EXP") - 1) + "1*EXP" + Right(ao$, Len(ao$) - InStr(ao$, "1EXP") - 3)
Loop
Do Until InStr(ao$, "2EXP") = 0
    ao$ = Left(ao$, InStr(ao$, "2EXP") - 1) + "2*EXP" + Right(ao$, Len(ao$) - InStr(ao$, "2EXP") - 3)
Loop
Do Until InStr(ao$, "3EXP") = 0
    ao$ = Left(ao$, InStr(ao$, "3EXP") - 1) + "3*EXP" + Right(ao$, Len(ao$) - InStr(ao$, "3EXP") - 3)
Loop
Do Until InStr(ao$, "4EXP") = 0
    ao$ = Left(ao$, InStr(ao$, "4EXP") - 1) + "4*EXP" + Right(ao$, Len(ao$) - InStr(ao$, "4EXP") - 3)
Loop
Do Until InStr(ao$, "5EXP") = 0
    ao$ = Left(ao$, InStr(ao$, "5EXP") - 1) + "5*EXP" + Right(ao$, Len(ao$) - InStr(ao$, "5EXP") - 3)
Loop
Do Until InStr(ao$, "6EXP") = 0
    ao$ = Left(ao$, InStr(ao$, "6EXP") - 1) + "6*EXP" + Right(ao$, Len(ao$) - InStr(ao$, "6EXP") - 3)
Loop
Do Until InStr(ao$, "7EXP") = 0
    ao$ = Left(ao$, InStr(ao$, "7EXP") - 1) + "7*EXP" + Right(ao$, Len(ao$) - InStr(ao$, "7EXP") - 3)
Loop
Do Until InStr(ao$, "8EXP") = 0
    ao$ = Left(ao$, InStr(ao$, "8EXP") - 1) + "8*EXP" + Right(ao$, Len(ao$) - InStr(ao$, "8EXP") - 3)
Loop
Do Until InStr(ao$, "9EXP") = 0
    ao$ = Left(ao$, InStr(ao$, "9EXP") - 1) + "9*EXP" + Right(ao$, Len(ao$) - InStr(ao$, "9EXP") - 3)
Loop
Do Until InStr(ao$, "0EXP") = 0
    ao$ = Left(ao$, InStr(ao$, "0EXP") - 1) + "0*EXP" + Right(ao$, Len(ao$) - InStr(ao$, "0EXP") - 3)
Loop

Do Until InStr(ao$, "EXP") = 0
    ao$ = Left(ao$, InStr(ao$, "EXP") - 1) + "EP" + Right(ao$, Len(ao$) - InStr(ao$, "EXP") - 2)
Loop

Do Until InStr(ao$, "PI") = 0
    ao$ = Left(ao$, InStr(ao$, "PI") - 1) + "3.14159265358979323846264338" + Right(ao$, Len(ao$) - InStr(ao$, "PI") - 1)
Loop

translate = LCase(ao$)
End Function

Public Function Bracket(aoo$)
alb = InStr(aoo$, "("): arb = InStr(aoo$, ")")
If alb >= 0 And arb > 0 And arb < alb Then aoo$ = "(" + aoo$


l:
Lb = 0: Rb = 0
cbo$ = aoo$
For cb = 1 To Len(cbo$)
  If Left(cbo$, 1) = "(" Then Lb = Lb + 1
  If Left(cbo$, 1) = ")" Then Rb = Rb + 1
  cbo$ = Right(cbo$, Len(cbo$) - 1)
Next cb
If Lb > Rb Then aoo$ = aoo$ + ")": Beep: lr = 1: GoTo l
If Lb < Rb Then aoo$ = "(" + aoo$: Beep: lr = 1: GoTo l
qu1: lr = 0

k:
Lb = 0: Rb = 0

Bracket = aoo$
End Function
Public Function Pixel(p As Long, x1 As Double, x2 As Double, xp As Single, yp As Boolean) As Long
  Pixel = (xp - x1) / Abs(x1 - x2) * p
  If yp = True Then Pixel = p - Pixel
End Function '坐标转换为像素
