Module modNumbers
    Public Function FitRange(ByVal LBnd, ByVal CHK, ByVal UBnd)
        On Error Resume Next
        If CHK < LBnd Then
            FitRange = LBnd
        ElseIf CHK > UBnd Then
            FitRange = UBnd
        Else
            FitRange = CHK
        End If
    End Function
    Public Function InRange(ByVal LBnd As Object, ByVal CHK As Object, ByVal UBnd As Object, Optional ByVal IncludeBounds As Boolean = True) As Boolean
        On Error Resume Next  ' because we're doing this as variants..
        If IncludeBounds Then
            InRange = (CHK >= LBnd) And (CHK <= UBnd)
        Else
            InRange = (CHK > LBnd) And (CHK < UBnd)
        End If
    End Function
    Public Function Random(ByVal Max as integer, Optional ByVal Min as integer = 0) as integer
        'To produce random integers in a given range, use this formula:
        'Int((upperbound - lowerbound + 1) * Rnd + lowerbound)
        Random = CLng((Max - Min + 1) * Rnd() + Min)
        If Random > Max Then Random = Max
    End Function
    Public Function MinArray(ByRef A)
        Dim L
        On Error Resume Next
        If Not IsArray(A) Then Exit Function
        If UBound(A) = -1 Then Exit Function
        MinArray = A(LBound(A))
        For Each L In A
            If Val(L) < Val(MinArray) Then MinArray = Val(L)
        Next
    End Function
    Public Function Trunc0(ByVal Number As Double) as integer
        Trunc0 = Trunc(Number, 0)
    End Function
    Public Function Trunc(ByVal Number As Double, Optional ByVal DecimalPoints as integer = 2) As Double
        Dim S As String, X as integer
        If DecimalPoints <= 0 Then
            Trunc = Int(Number)
            Exit Function
        End If
        S = "" & Number
        X = InStr(S, ".")
        Trunc = Val(Left(S, X + DecimalPoints))
    End Function
    Public Function Decimals(ByVal Number As Double) As Double
        Decimals = Number - Trunc(Number, 0)
    End Function

End Module
