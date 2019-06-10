Attribute VB_Name = "UserFun"

Public Function Gcd(m, n) As Double
  m = Fix(m)
  n = Fix(n)
  
  If n = 0 Then Gcd = m Else Gcd = Gcd(n, m Mod n)
  
End Function

Public Function Parity(n) As Long '奇偶
Parity = 0
On Error Resume Next
If Abs(n) > 2147483640 Then Exit Function
If Fix(n) <> n Then
  Parity = 0
Else
  If n / 2 = n \ 2 Then Parity = 2 Else Parity = 1
End If
End Function

Public Function Round(n)
Round = Fix(n + 0.5)
End Function


Public Function RadXY(x, y)
p = 4 * Atn(1)
If x = 0 And y = 0 Then
  RadXY = Log(-1)
Else
  If x = 0 Then
    If y > 0 Then RadXY = p / 2 Else RadXY = 3 * p / 2
  Else
    If y = 0 Then
      If x > 0 Then RadXY = 0 Else RadXY = p
    Else
      If x > 0 Then
        If y > 0 Then RadXY = Atn(x / y) Else RadXY = Atn(x / y) + 2 * p
      Else
        RadXY = p + Atn(x / y)
      End If
    End If
  End If
End If

End Function

Public Function DegXY(x, y)
p = 4 * Atn(1)
DegXY = Degrees(RadXY(x, y))
End Function



Public Function Degrees(n)  '将弧度转换为度
p = 4 * Atn(1)
Degrees = n * 180 / p
End Function

Public Function Radians(n)   '将度转换为弧度
p = 4 * Atn(1)
Radians = n * p / 180
End Function


Public Function Arsech(n)
If n > 0 And n <= 1 Then
Arsech = Log((1 + Sqr(1 - n ^ 2)) / (1 - Sqr(1 - n ^ 2))) / 2
Else
Arsech = Log(-1)
End If
End Function
Public Function Combination(m, n) '组合
'if (m==n)return 1;
'    return Multi(m,n)/Multi(m-n,1);
If m = n Then Combination = 1 Else Combination = Multiply(m, n) / Multiply(m - n, 1)
End Function
Public Function Arrange(m, n)
If m = n Then Arrange = Multiply(m, 1) Else Arrange = Multiply(m, m - n)
End Function
Public Function Multiply(mm, nn)  '从m乘到n
'if (m==n)return 1;
' long s;
' int i=1,j;
' j=m-n-1;
' for(s=m,m--;i<=n;i++,m--)
'    {s=s/i*m;}
' for(i=1;i<=j;i++)
'    {s=s*i;}
' return s;
m = mm: n = nn
If m = n Then Multiply = 1: Exit Function
j = m - n - 1
s = m
m = m - 1
For i = 1 To n
  s = s / i * m
  m = m - 1
Next i
For i = 1 To j
 s = s * i
Next i
Multiply = s
End Function
