Attribute VB_Name = "Bt"
Public Function bracketT(Bt$) As String
Bt$ = LCase(Bt$)
Do Until InStr(Bt$, "t") = 0
    Bt$ = Left(Bt$, InStr(Bt$, "t") - 1) + "?" + Right(Bt$, Len(Bt$) - InStr(Bt$, "t"))
Loop

Do Until InStr(Bt$, "?runc") = 0
    Bt$ = Left(Bt$, InStr(Bt$, "?runc") - 1) + "trunc" + Right(Bt$, Len(Bt$) - InStr(Bt$, "?runc") - 4)
Loop
Do Until InStr(Bt$, "in?") = 0
    Bt$ = Left(Bt$, InStr(Bt$, "in?") - 1) + "int" + Right(Bt$, Len(Bt$) - InStr(Bt$, "in?") - 2)
Loop
Do Until InStr(Bt$, "?an") = 0
    Bt$ = Left(Bt$, InStr(Bt$, "?an") - 1) + "tan" + Right(Bt$, Len(Bt$) - InStr(Bt$, "?an") - 2)
Loop
Do Until InStr(Bt$, "?g") = 0
    Bt$ = Left(Bt$, InStr(Bt$, "?g") - 1) + "tg" + Right(Bt$, Len(Bt$) - InStr(Bt$, "?g") - 1)
Loop
Do Until InStr(Bt$, "co?") = 0
    Bt$ = Left(Bt$, InStr(Bt$, "co?") - 1) + "cot" + Right(Bt$, Len(Bt$) - InStr(Bt$, "co?") - 2)
Loop
Do Until InStr(Bt$, "?h") = 0
    Bt$ = Left(Bt$, InStr(Bt$, "t") - 1) + "th" + Right(Bt$, Len(Bt$) - InStr(Bt$, "?h") - 1)
Loop

Do Until InStr(Bt$, "sint") = 0
    Bt$ = Left(Bt$, InStr(Bt$, "sint") - 1) + "sin(t)" + Right(Bt$, Len(Bt$) - InStr(Bt$, "sint") - 3)
Loop
Do Until InStr(Bt$, "?") = 0
    Bt$ = Left(Bt$, InStr(Bt$, "?") - 1) + "(t)" + Right(Bt$, Len(Bt$) - InStr(Bt$, "?"))
Loop
bracketT = Bt$
End Function

