Module modEnglish
    Public Function YesNo(ByVal Value As Boolean, Optional ByVal vUCase As Boolean = False, Optional ByVal SingleChar As Boolean = False) As String
        YesNo = IIf(Value, "Yes", "No")
        If vUCase Then YesNo = UCase(YesNo)
        If SingleChar Then YesNo = Left(YesNo, 1)
    End Function

    Public Function TrueFalseString(ByVal fVar As Object)
        On Error Resume Next
        TrueFalseString = IIf(TrueFalseValue(fVar), "True", "False")
    End Function

    Public Function TrueFalseValue(ByVal fVar) As Boolean
        On Error Resume Next
        Dim V As String
        If IsNothing(fVar) Then Exit Function
        V = "" & fVar
        If IsNumeric(V) Then
            TrueFalseValue = Val(V) <> 0
        Else
            TrueFalseValue = IsIn(LCase(Left(V, 1)), "t", "1", "y")
        End If
    End Function
End Module
