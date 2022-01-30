Attribute VB_Name = "de"
Public Function d_fx(ByVal fx As String) As String
 
 d_fx = Derivative(fx)
 d_fx = Replace(d_fx, "$", "1")
 DelBracket = DelLRBracket(d_fx)
End Function

Public Function Derivative(ByVal s As String) As String
  Dim DelBracket As Boolean
  DelBracket = DelLRBracket(s) '���s�����ַ��������������ַ�����������ɾ������
  If s = "" Then Derivative = "0": Exit Function '�Գ�0��'����������
  If NumTF(s) = True Then Derivative = "0": Exit Function '�Գ�����
  If s = "x" Then Derivative = "$": GoTo l1:  '��x��
  
 
  
  Derivative = Derivative_PM(s)
  
  Derivative = (AddlrBracket(Derivative))
  
l1:  Derivative = CleanUpExrp(Derivative)
  
End Function

Public Function Derivative_PM(ByVal s As String)  'һ���󵼣���������������ŵı��ʽ

'�������������ž�ָ���������֮���������
'������Ҫȥ����!!!!

DelBracket = DelLRBracket(s)

If PositionPM(s) = 0 Then
  Derivative_PM = Derivative_MD(s)
  Exit Function
End If

Dim Term() As String    '��̬����Term()��������������֮��Ĳ���
Dim Oper() As String       '��̬����Oper()������ţ��ӺŻ���ţ�
Dim Position_Oper() As Integer  '��̬���� Position_Oper()����ӺŻ��������λ��
Dim i, k As Integer


'��s�������������� ����ʽ�� Oper1 Term1 Oper2 Term2 Oper3 Term3  ...  Oper i Term i
If Left(s, 1) <> "+" Then s = "+" & s  '��sǰ��+�������+��

Dim p As Integer
p = PositionPM(s)

Do Until p = 0
  i = i + 1
  ReDim Preserve Oper(i + 1)  'ReDim Preserve �����±�����,���˷���Oper(0),������ͬ
  Oper(i) = mid(s, p, 1)
  ReDim Preserve Position_Oper(i + 1)
  Position_Oper(i) = p
  Mid(s, p, 1) = "#"     '��pλ�õķ����滻��#
  p = PositionPM(s)
Loop

ReDim Term(i)
If i > 1 Then
  For k = 1 To i - 1
    Term(k) = mid(s, Position_Oper(k) + 1, Position_Oper(k + 1) - Position_Oper(k) - 1)
  Next k
  Term(i) = Right(s, Len(s) - Position_Oper(i))
Else
  Term(1) = Right(s, Len(s) - 1)
End If

For k = 1 To i   '������
  Derivative_PM = Derivative_PM & Oper(k) & Derivative(Term(k))
Next k

'ȥ����ͷ��+��
Derivative_PM = Right(Derivative_PM, Len(Derivative_PM) - 1)


End Function

Public Function Derivative_MD(ByVal s As String)  '������,����������˳��ŵı��ʽ

'���������˳��ž�ָ���������֮��ĳ˳���
DelBracket = DelLRBracket(s)

Dim p As Integer
p = PositionMD(s)

If p = 0 Then
  Derivative_MD = Derivative_PE(s)
  Exit Function
End If

Dim Section() As String    '��̬����Section()���������˳���֮��Ĳ���
Dim Term() As String  '��̬����Term()���������˳���֮��Ĳ���
Dim Oper() As String       '��̬����Oper()������ţ��˺Ż���ţ�
Dim Position_Oper() As Integer  '��̬���� Position_Oper()����˺Ż��������λ��
Dim i, k As Integer
Dim Num_d As Integer  'Num_d ��¼���ŵ���Ŀ

'��s��������������,����ʽ�� Section1 Oper1 Section2 Oper2 Section3 Oper3 ... Section i, Oper i

s = s & "*"   '��s��ĩβ���*��

Do Until p = 0
  i = i + 1
  ReDim Preserve Oper(i + 1)  'ReDim Preserve �����±�����,���˷���Oper(0),������ͬ
  Oper(i) = mid(s, p, 1)
  If Oper(i) = "/" Then Num_d = Num_d + 1
  ReDim Preserve Position_Oper(i + 1)
  Position_Oper(i) = p
  Mid(s, p, 1) = "#"     '��pλ�õķ����滻��#
  p = PositionMD(s)
Loop

ReDim Section(i)
Section(1) = Left(s, Position_Oper(1) - 1)
For k = 2 To i
  Section(k) = mid(s, Position_Oper(k - 1) + 1, Position_Oper(k) - Position_Oper(k - 1) - 1)
Next k

'����,��*��������Section�ŵ�ǰ��,/��������Section�ŵ�����,����������Term().
Dim j, l As Integer
ReDim Term(i)
Term(1) = Section(1)
j = 1: l = i - Num_d
For k = 1 To i - 1
  If Oper(k) = "*" Then
    j = j + 1
    Term(j) = Section(k + 1)
  Else
    l = l + 1
    Term(l) = Section(k + 1)
  End If
Next k

For k = 1 To i
  DelBracket = DelLRBracket(Term(i))   '��ÿһ��ȥ�����ţ���������������������ģ�
Next k

Dim Numerator, Denominator, D_Numerator, D_Denominator As String

l = i - Num_d

For k = 1 To l
  Numerator = Numerator & Term(k) & "*"
Next k
Numerator = "(" & Left(Numerator, Len(Numerator) - 1) & ")"
Dim g As Boolean
Dim m, n As Integer
Dim Derivative_M As String
Dim h As String
Dim ContMulti() As String

m = 1: n = l
GoTo l3:
l1: D_Numerator = "(" & Derivative_M & ")"

If Num_d > 0 Then
  For k = l + 1 To i
    Denominator = Denominator & Term(k) & "*"
  Next k
  Denominator = "(" & Left(Denominator, Len(Denominator) - 1) & ")"
  m = l + 1: n = i: g = True
  GoTo l3:
l2:   D_Denominator = "(" & Derivative_M & ")"
  Derivative_MD = "(" & D_Numerator & "*" & Denominator & "-" & Numerator & "*" & D_Denominator & ")/" & Denominator & "^2"
Else
  Derivative_MD = D_Numerator
End If

Exit Function

l3:
     '��Term(m)��Term(n)����ʽ��
  Derivative_M = ""
  
  ReDim ContMulti(n + 1) As String
  For k = m To n
    For j = m To n
      If j = k Then
        h = Derivative(Term(j))
        If h <> "0" Then
          ContMulti(k) = ContMulti(k) & h & "*"
        Else
          ContMulti(k) = ContMulti(k) & "&" & "*"  '����0����&����
        End If
      Else
        ContMulti(k) = ContMulti(k) & "(" & Term(j) & ")*"
      End If
      
    Next j
    ContMulti(k) = Left(ContMulti(k), Len(ContMulti(k)) - 1)
  Next k
  
  
  For k = m To n
    If InStr(ContMulti(k), "&") > 0 Then ContMulti(k) = "&" '��ContMulti(k)�ǰ���&�ģ�˵��ContMulti(k)��ֵ��0��ȫ����&����
    Derivative_M = Derivative_M & ContMulti(k) & "+"
  Next k
  Derivative_M = Replace(Derivative_M, "+&", "")
  Derivative_M = Replace(Derivative_M, "&+", "")  '��&����
  If Derivative_M = "" Then Derivative_M = "0"
  If Right(Derivative_M, 1) = "+" Then Derivative_M = Left(Derivative_M, Len(Derivative_M) - 1)
If g = True Then GoTo l2 Else GoTo l1:
End Function
Public Function Derivative_PE(ByVal s As String)  '������,����������ݵı��ʽ

'��������^@�ž�ָ���������֮���^@��
DelBracket = DelLRBracket(s)

Dim p As Integer
p = PositionPE(s)

If p = 0 Then
  Derivative_PE = Derivative_Fct(s)
  Exit Function
End If

Dim Section() As String    '��̬����Section()��������^@��֮��Ĳ���
Dim Term() As String  '��̬����Term()��������^@��֮��Ĳ���
Dim Oper() As String       '��̬����Oper()�������(^�Ż�@��)
Dim Position_Oper() As Integer  '��̬���� Position_Oper()����^�Ż�@������λ��
Dim i, k As Integer


'��s��������������,����ʽ�� Section1 Oper1 Section2 Oper2 Section3 Oper3 ... Section i, Oper i

s = s & "^"   '��s��ĩβ���^��
If Left(s, 1) = "@" Then s = "2" & s '��s��@��ͷ,����sͷ�����2

Do Until p = 0
  i = i + 1
  ReDim Preserve Oper(i + 1)  'ReDim Preserve �����±�����,���˷���Oper(0),������ͬ
  Oper(i) = mid(s, p, 1)
  ReDim Preserve Position_Oper(i + 1)
  Position_Oper(i) = p
  Mid(s, p, 1) = "#"     '��pλ�õķ����滻��#
  p = PositionPE(s)
Loop

ReDim Section(i)
Section(1) = Left(s, Position_Oper(1) - 1)
For k = 2 To i
  Section(k) = mid(s, Position_Oper(k - 1) + 1, Position_Oper(k) - Position_Oper(k - 1) - 1)
Next k


Dim Left_Section As String
For k = 1 To i - 1
  Left_Section = Left_Section & Section(k) & Oper(k)
Next k
Left_Section = Left(Left_Section, Len(Left_Section) - 1)

If i = 2 Then
  If Oper(i - 1) = "^" Then
    Derivative_PE = Derivative_Power(Section(1), Section(2))
  Else
    Derivative_PE = Derivative_Extract(Section(1), Section(2))
  End If
Else
  Derivative_PE = Derivative("(" & Left_Section & ")" & Oper(i - 1) & "(" & Section(i) & ")")
End If
End Function

  

Public Function Derivative_Power(ByVal u As String, ByVal v As String) As String '��u^v��
  Dim DelBracket_u, DelBracket_v As Boolean
  DelBracket_u = DelLRBracket(u)
  DelBracket_v = DelLRBracket(v)
  If NumTF(v) = True Then
    Derivative_Power = v & "*(" & u & ")^(" & Trim(Str(Val(v) - 1)) & ")*" & Derivative(u)
  Else
    Derivative_Power = "(" & u & ")^(" & v & ")*(" & Derivative(v) & "*ln(" & u & ")+(" & v & ")/(" & u & ")*" & Derivative(u) & ")"
  End If
End Function

Public Function Derivative_Extract(ByVal u As String, ByVal v As String) As String '��u@v��)
Dim DelBracket_u, DelBracket_v As Boolean
DelBracket_u = DelLRBracket(u)
DelBracket_v = DelLRBracket(v)
If NumTF(u) = True Then
   Derivative_Extract = Trim(Str(1 / Val(u))) & "*((" & Trim(Str(1 / (1 / Val(u) - 1))) & ")@(" & v & "))*" & Derivative(v) '?????
 Else
   Derivative_Extract = Derivative_Power(v, "1/(" & u & ")")  '?????XXXXX
 End If
End Function



Private Function Derivative_Fct(s As String) As String '�ļ���,���Ժ�����
  Dim f, x, d As String
  DelBracket = DelLRBracket(s)
  f = LCase(Funname(UCase(s))) 'f�Ǻ����� ����Funname�����ַ���ǰ��ĺ�����
  x = Right(s, Len(s) - Len(f))
  
  
  Select Case f
    Case ""
    Derivative_Fct = "(" & s & ")'" '?????????
    msg = MsgBox("��������ʶ��ı��ʽ " & x, vbInformation, "����")
    Exit Function
    Case "ln", "log", "lna"
    d = "1/(" & x & ")"
    Case "lg"
    d = "1/((" & x & ")*ln10)"
    Case "exp", "ep"
    d = "exp(" & x & ")"
    Case "sin"
    d = "cos(" & x & ")"
    Case "cos"
    d = "-sin(" & x & ")"
    Case "tan", "tg"
    d = "sec(" & x & ")^2"
    Case "cot", "ctg"
    d = "-csc(" & x & ")^2"
    Case "sec"
    d = "sec(" & x & ")*tan(" & x & ")"
    Case "csc"
    d = "-csc(" & x & ")*cot(" & x & ")"
    Case "arcsin", "asin"
    d = "1/@(1-(" & x & ")^2)"
    Case "arccos", "acos"
    d = "-1/@(1-(" & x & ")^2)"
    Case "arctan", "arctg", "atn", "atan"
    d = "1/(1+(" & x & ")^2)"
    Case "arccot", "arcctg", "acot"
    d = "-1/(1+(" & x & ")^2)"
    Case "arcsec", "asec"
    d = "1/((" & x & ")*@((" & x & ")^2-1))"
    Case "arccsc", "acsc"
    d = "-1/((" & x & ")*@((" & x & ")^2-1))"
    Case "sh", "sinh"
    d = "ch(" & x & ")"
    Case "ch", "cosh"
    d = "sh(" & x & ")"
    Case "th", "tanh"
    d = "sech(" & x & ")^2"
    Case "cth", "coth"
    d = "-csch(" & x & ")^2"
    Case "sech"
    d = "-th(" & x & ")*sech(" & x & ")"
    Case "csch"
    d = "-cth(" & x & ")*csch(" & x & ")"
    Case "arsh", "asinh"
    d = "1/@(1+(" & x & ")^2)"
    Case "arch", "acosh"
    d = "+-1/@((" & x & ")^2-1)" '???????????"
    Case "arth", "atanh"
    d = "1/(1-(" & x & ")^2)"
    Case "arcth", "acoth"
    d = "1/(1-(" & x & ")^2)"
    
  End Select

  If x <> "x" Then
    d = d & "*" & Derivative(x)  '???????
  End If
  
  Derivative_Fct = d
  
End Function



Public Function PositionPM(s As String) As Integer  'ָ������������ⲿ�ĵ�һ�����Ż򸺺ŵ�λ��.
PositionPM = 0
Dim i As Integer
For i = 1 To Len(s)
  If mid(s, i, 1) = "+" Or mid(s, i, 1) = "-" Then
    If Outside_Bracket(i, s) = True Then
      PositionPM = i
      Exit For
    End If
  End If
Next i
End Function

Public Function PositionMD(s As String) As Integer 'ָ������������ⲿ�ĵ�һ���˺Ż���ŵ�λ��.
PositionMD = 0
Dim i As Integer
For i = 1 To Len(s)
  If mid(s, i, 1) = "*" Or mid(s, i, 1) = "/" Then
    If Outside_Bracket(i, s) = True Then
      PositionMD = i
      Exit For
    End If
  End If
Next i
End Function

Public Function PositionPE(s As String) As Integer 'ָ������������ⲿ�ĵ�һ���˷��򿪷��ŵ�λ��.
PositionPE = 0
Dim i As Integer
For i = 1 To Len(s)
  If mid(s, i, 1) = "^" Or mid(s, i, 1) = "@" Then
    If Outside_Bracket(i, s) = True Then
      PositionPE = i
      Exit For
    End If
  End If
Next i
End Function

Public Function Outside_Bracket(ByVal i As Integer, ByVal s As String) As Boolean   '����i���ַ��Ƿ����ַ���s����������ŵ��ⲿ�����򷵻�True
  
  If InStr(s, "(") = 0 Then
    Outside_Bracket = True
    Exit Function
  End If
  
  Dim t() As Boolean
  Dim l, k, a As Integer

  l = Len(s)
  
  If i < 1 Or i > l Then
    Outside_Bracket = True
    Exit Function
  End If
  
  ReDim t(l)    '����һ������Ϊ�ַ���s���ȵ�����
  
  Dim Flag As Boolean 'flagָʾk�Ƿ�������������ڲ�,���������������Ҳ��������������ⲿ,��Ӱ��ʵ�ʵ���Ч��.
  
  For k = 1 To l
    c = mid(s, k, 1)
    If c = "(" Then
      a = a + 1  '���������ţ�a�ͼ�1
      Flag = True
    End If
    If c = ")" Then
      a = a - 1   '���������ţ�a�ͼ�1
      Flag = True
    End If
    If Flag = True And a = 0 Then
      Flag = False  'a������ʱ��k������������ⲿ
    End If
    t(k) = Flag
  Next k
  
  If t(i) = True Then Outside_Bracket = False Else Outside_Bracket = True
 
End Function



Public Function NumTF(s As String) As Boolean '�ж��ַ����Ƿ�ȫ����������ɵ�(��������+-)
  If Trim(Str(Val(s))) = s Then NumTF = True Else NumTF = False
End Function



Public Function DelLRBracket(s As String) As Boolean  '���s���������������ģ���ɾ����������,����ֵ�����Ƿ�ɾ��������
  While OuterLeftBracket(s) = 1 And OuterRightBracket(s) = Len(s)
    s = mid(s, 2, Len(s) - 2)
    DelLRBracket = True
  Wend
End Function
Public Function AddlrBracket(s As String) As String '��s�����������
  AddlrBracket = "(" & s & ")"
End Function


Public Function OuterRightBracket(s As String) As Integer  '��������������ŵ�λ��
  Dim i, a As Integer
  Dim c As String  'c��iλ�õ��ַ�
  Dim Flag As Boolean 'flag��¼�Ƿ�����������
  For i = 1 To Len(s)
    c = mid(s, i, 1)
    If c = "(" Then a = a + 1: Flag = True
    If c = ")" Then a = a - 1: Flag = True
    If Flag = True And a = 0 Then
      OuterRightBracket = i
      Exit For
    End If
  Next i
End Function

Public Function OuterLeftBracket(s As String) As Integer '��������������ŵ�λ��
  OuterLeftBracket = InStr(s, "(")
End Function

Public Function Inside_OuterBracket(s As String) As String '��������������ڵĲ���
  If InStr(s, ")") = 0 Then
    Inside_OuterBracket = s  '��s�������ţ��򷵻�ԭ�ַ���s�����˳�����
    Exit Function
  End If
  Inside_OuterBracket = mid(s, OuterLeftBracket(s) + 1, OuterRightBracket(s) - OuterLeftBracket(s) - 1)
End Function

Public Function Left_OuterBracket(s As String) As String '���������������ߵĲ���,�ù���ʵ��û�еõ�Ӧ��,�������汸��
  Dim a As Integer
  a = OuterLeftBracket(s)
  If a = 0 Then
    Left_OuterBracket = s '�����Ų�����,�򷵻�ԭ�ַ���
  Else
    Left_OuterBracket = Left(s, a - 1)
  End If
End Function

Public Function Right_OuterBracket(s As String) As String '��������������ұߵĲ���,�ù���ʵ��û�еõ�Ӧ��,�������汸��
  If InStr(s, ")") = 0 Then
    Right_OuterBracket = s  '��s�������ţ��򷵻�ԭ�ַ���s�����˳�����
    Exit Function
  End If
  Right_OuterBracket = Right(s, Len(s) - OuterRightBracket(s))
End Function

Public Function Letter_String_Right(s As String) As String '�����ַ���s���ұߵ���ĸ���,�ù���ʵ��û�еõ�Ӧ��,�������汸��
  Dim c As String
  Dim a As Integer
  For i = 0 To Len(s)
    c = mid(s, Len(s) - i, 1)
    a = Asc(c)
    If a >= 65 And a <= 90 Or a >= 97 And a <= 122 Then
      Letter_String_Right = c & Letter_String_Right
    Else
      Exit For
    End If
  Next i
End Function



Public Function Replace(ByVal s As String, s1 As String, s2 As String) '���ַ���s�е��ַ���s1ȫ���滻���ַ���s2
  Do Until InStr(s, s1) = 0
    s = Left(s, InStr(s, s1) - 1) & s2 & Right(s, Len(s) - InStr(s, s1) - Len(s1) + 1)
  Loop
Replace = s
End Function

Public Function DelStr(ByVal s As String, del As String) ' ���ַ���s�е��ַ���delȫ��ɾ��
  DelStr = Replace(s, del, "")
End Function



Public Function CleanUpExrp(ByVal s As String) As String '���� ������ʽ



Dim i As Integer
Dim q As String
For i = 0 To 9
  q = Trim(Str(i))
  s = Replace(s, "(" & q & ")", q)
Next i
s = Replace(s, "($)", "$")
s = DelStr(s, "*$")   'ɾ�����ʽ�к��е�*$ (ʵ����*1)
s = DelStr(s, "$*")
s = Replace(s, "^1+", "+")
s = Replace(s, "^1-", "-")
s = Replace(s, "^1*", "*")
s = Replace(s, "^1/", "/")


Dim r(1 To 29) As String
r(1) = "(x)"
r(2) = "(sinx)"
r(3) = "(cosx)"
r(4) = "(tanx)"
r(5) = "(cotx)"
r(6) = "(secx)"
r(7) = "(cscx)"
r(8) = "(arcsinx)"
r(9) = "(arccosx)"
r(10) = "(arctanx)"
r(11) = "(arccotx)"
r(12) = "(arcsecx)"
r(13) = "(arccscx)"
r(14) = "(shx)"
r(15) = "(chx)"
r(16) = "(thx)"
r(17) = "(cthx)"
r(18) = "(sechx)"
r(19) = "(cschx)"
r(20) = "(arshx)"
r(21) = "(archx)"
r(22) = "(arthx)"
r(23) = "(arcthx)"
r(24) = "(arsechx)"
r(25) = "(arcschx)"
r(26) = "(lnx)"
r(27) = "(lgx)"
r(28) = "(expx)"
r(29) = "(logx)"



For i = 1 To 29
  s = Replace(s, r(i), mid(r(i), 2, Len(r(i)) - 2))
Next i





If Left(s, 2) = "0-" Then s = Right(s, Len(s) - 1)

CleanUpExrp = s
End Function
Public Function ExpChk_d(Ex As String) As String '�������ı��ʽ�Ƿ����Ų�ƥ��
While InStr(Ex, "[") > 0
    Mid$(Ex, InStr(Ex, "["), 1) = "("
Wend

While InStr(Ex, "]") > 0
    Mid$(Ex, InStr(Ex, "]"), 1) = ")"
Wend

Dim a, b As Integer
a = InStr(Ex, "("): b = InStr(Ex, ")")
If (a > 0 And b > 0 And a > b) Or (a = 0 And b > 0) Then
  ExpChk_d = "���ʽ�����© '( '"
  Exit Function
End If

Dim Lb, Rb As Integer
Lb = 0: Rb = 0

Dim cbo As String
cbo = Ex

Dim cb As Integer
For cb = 1 To Len(cbo)
  If Left(cbo, 1) = "(" Then Lb = Lb + 1
  If Left(cbo, 1) = ")" Then Rb = Rb + 1
  cbo = Right(cbo, Len(cbo) - 1)
Next cb

If Lb > Rb Then
  ExpChk_d = "���Ų�ƥ�䡣��©" & Str(Abs(Lb - Rb)) & "�� ') '�� "
  Exit Function
End If
If Lb < Rb Then
  ExpChk_d = "���Ų�ƥ�䡣��©" & Str(Abs(Lb - Rb)) & "�� '( '�� "
  Exit Function
End If

End Function

