CheckStr = "[A-Za-z0-9._-]"
OutStr = ""
Index = 1
Do While True
    Index1 = VBA.InStr(Index, extractStr, "@")
    getStr = ""
    If Index1 > 0 Then
        For p = Index1 - 1 To 1 Step -1
            If Mid(extractStr, p, 1) Like CheckStr Then
                getStr = Mid(extractStr, p, 1) & getStr
            Else
                Exit For
            End If
        Next
        getStr = getStr & "@"
        For p = Index1 + 1 To Len(extractStr)
            If Mid(extractStr, p, 1) Like CheckStr Then
                getStr = getStr & Mid(extractStr, p, 1)
            Else
                Exit For
            End If
        Next
        Index = Index1 + 1
        If OutStr = "" Then
            OutStr = getStr
        Else
            OutStr = OutStr & Chr(10) & getStr
        End If
    Else
        Exit Do
    End If
Loop
ExtractEmailFun = OutStr
End Function
