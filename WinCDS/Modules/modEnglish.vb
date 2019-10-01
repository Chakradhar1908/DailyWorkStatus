Module modEnglish
    Public Function YesNo(ByVal Value As Boolean, Optional ByVal vUCase As Boolean = False, Optional ByVal SingleChar As Boolean = False) As String
        YesNo = IIf(Value, "Yes", "No")
        If vUCase Then YesNo = UCase(YesNo)
        If SingleChar Then YesNo = Left(YesNo, 1)
    End Function
End Module
