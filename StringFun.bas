Attribute VB_Name = "StringFun"
Public Function AddNum(n) As String
On Error Resume Next
  Dim s As String
  s = Trim(Str(n))
  If Left(s, 1) <> "+" And Left(s, 1) <> "-" Then s = "+" & s
  AddNum = s
End Function
Public Function SpaceNum(sn$) As Long '空格的数量
  n$ = Trim(sn$)
  k = 1
  Do Until k = 0
    k = InStr(n$, " ")
    n$ = Trim(Right(n$, Len(n$) - k))
    j = j + 1
  Loop
  
  SpaceNum = j - 1
  

End Function

Public Function EnterNum(sn$) As Long '回车Chr(10)的数量
  n$ = Trim(sn$)
  k = 1
  Do Until k = 0
    k = InStr(n$, Chr(10))
    n$ = Trim(Right(n$, Len(n$) - k))
    j = j + 1
  Loop
  
EnterNum = j - 1
End Function
Public Function Left0(s1 As String, s2 As String) As Integer  's2左边的部分
  Left0 = Left(s1, InStr(s1, s2))
End Function
Public Function Right0(s1 As String, s2 As String) As Integer  's2右边的部分
  Right0 = Right(s1, InStr(s1, s2))
End Function
Public Function Multinomial(ByVal s As String) As String  '把只有空格和数字的字符串转化为一元n次多项式
  s = Trim(s)
  If InStr(s, " ") = 0 Then Multinomial = s: Exit Function
  Dim Term() As String
  Dim p, n As Integer
  n = 0
  Do Until InStr(s, " ") = 0
    p = InStr(s, " ")
    ReDim Preserve Term(n + 1)
    Term(n) = Left(s, p - 1)
    If Left(Term(n), 1) <> "+" And Left(Term(n), 1) <> "-" Then Term(n) = "+" & Term(n)
    s = Trim(Right(s, Len(s) - p))
    n = n + 1
  Loop
  Term(n) = s
  If Left(Term(n), 1) <> "+" And Left(Term(n), 1) <> "-" Then Term(n) = "+" & Term(n)

  Dim i As Integer
  For i = 0 To n
    Multinomial = Multinomial & Term(i) & "x^" & Trim(Str(n - i))
  Next i
  
  Multinomial = Left(Multinomial, Len(Multinomial) - 3)
  If Left(Multinomial, 1) = "+" Then Multinomial = Right(Multinomial, Len(Multinomial) - 1)
  Multinomial = Left(Multinomial, InStr(Multinomial, "^1") - 1) & Right(Multinomial, Len(Multinomial) - InStr(Multinomial, "^1") - 1)
  
End Function
