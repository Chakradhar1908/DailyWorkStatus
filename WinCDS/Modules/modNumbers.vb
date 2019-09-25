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

    Public Function Max(ParamArray A())
        Dim B()
        B = A
        Max = MaxArray(B)
    End Function

    Public Function Min(ParamArray A())
        Dim B()
        B = A
        Min = MinArray(B)
    End Function

    Public Function MaxArray(ByRef A)
        '::::Decimals
        ':::SUMMARY
        ': Used to return maximum value in given Array.
        ':::DESCRIPTION
        ': This function is used to return the maximum value in given array, using loop given below.
        ':::PARAMETERS
        ': - A - Indicates the Reference Array.
        ':::RETURN
        '::: SEE ALSO
        ': - Max , Min
        Dim L
        On Error Resume Next
        If Not IsArray(A) Then Exit Function
        If UBound(A) = -1 Then Exit Function
        MaxArray = A(LBound(A))
        For Each L In A
            If Val(L) > Val(MaxArray) Then MaxArray = Val(L)
        Next
    End Function

    Public Function Pow(ByVal X As Double, ByVal Y As Integer) As Double
        '::::Pow
        ':::SUMMARY
        ': Used to return result of a number raised to a specified power.
        ':::DESCRIPTION
        ': This is a generalized exponential function.
        ': It returns the result of a number raised to a specified power.
        ':::PARAMETERS
        ': - X - Indicates the Base value.
        ': - Y - Indicates the Power value.
        ':::RETURN
        ': Double : Return Double that is x (the base) raised to the power y (the exponent).
        Dim N As Integer
        If Y < 0 Then Exit Function
        If Y = 0 Then Pow = 1 : Exit Function
        Pow = X
        For N = 2 To Y
            Pow = Pow * X
        Next
    End Function

    Public Function RoundDn(ByVal Number As Double) As Double
        '::::RoundDn
        ':::SUMMARY
        ': Used to Positivre RoundUp value or Integer part of given number.
        ':::DESCRIPTION
        ': This function is used to return the Integer part of given number if given number > 0 or Positive RoundUp value of given number.
        ':::PARAMETERS
        ': - Number - Indicates the Input Value.
        ':::RETURN
        ': Double : Return the RoundDn value as Double.

        If Number < 0 Then
            RoundDn = -RoundUp(-Number)
        Else
            RoundDn = Trunc0(Number)
        End If
    End Function

    Public Function RoundUp(ByVal Money As Decimal) As Decimal
        '::::RoundUp
        ':::SUMMARY
        ': Used to Round Up the Currency Amount.
        ':::DESCRIPTION
        ': This function is used to RoundUp the value of Currency amount to its nearest value.
        ':::PARAMETERS
        ': - Money - Indicates the Currency Value.
        ':::RETURN
        ': Currency : Return the RoundUp value as Currency.
        RoundUp = Math.Round(Money + 0.49, 0)
    End Function

End Module
