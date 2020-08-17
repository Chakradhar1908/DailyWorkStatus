Module modSplit
    Public Sub AdjustSalesSplits(ByRef A As String, ByRef B As String, ByRef C As String, ByVal Count As Integer)
        Dim Curr As Double, Rest As Double, HasD As Boolean
        Dim HasDefault As Boolean, R As Double
        Dim N As Integer

        Rest = 100.0#
        Curr = 0#
        HasDefault = False
        '  Count = IIf(Sales2 = "", 1, IIf(Sales3 = "", 2, 3))

        R = SplitValue(A)
        If R = 0 Then HasDefault = True
        Curr = Curr + R
        If Curr > 100 Then
            A = "" & Rest & "%"
            Curr = 100.0#
            Rest = 0
        Else
            Rest = Rest - R
        End If

        If Count < 2 Then
            B = "0%"
        Else
            R = SplitValue(B)
            If R = 0 Then HasDefault = True
            Curr = Curr + R
            If Curr > 100 Then
                B = "" & Rest & "%"
                Curr = 100.0#
                Rest = 0
            Else
                Rest = Rest - R
            End If
        End If

        If Count < 3 Then
            C = "0%"
        Else
            R = SplitValue(C)
            If R = 0 Then HasDefault = True
            Curr = Curr + R
            If Curr > 100 Then
                C = "" & Rest & "%"
                Curr = 100.0#
                Rest = 0
            Else
                Rest = Rest - R
            End If
        End If

        If HasDefault Then
            If SplitValue(A) = 0 Then N = N + 1
            If SplitValue(B) = 0 And Count >= 2 Then N = N + 1
            If SplitValue(C) = 0 And Count >= 3 Then N = N + 1

            If SplitValue(A) = 0 Then A = Rest / N
            If SplitValue(B) = 0 And Count >= 2 Then B = Rest / N
            If SplitValue(C) = 0 And Count >= 3 Then C = Rest / N

            Rest = 0
            Curr = 100
            HasDefault = False
        End If

        If Rest > 0 Then
            If Count = 3 Then
                C = "" & Rest + SplitValue(C) & "%"
            ElseIf Count = 2 Then
                B = "" & Rest + SplitValue(B) & "%"
            Else
                A = "0%"
            End If
        End If

        A = "" & SplitValue(A) & "%"
        B = "" & SplitValue(B) & "%"
        C = "" & SplitValue(C) & "%"
    End Sub

    'Public Function SplitCount(ByVal A As Object, ByVal B As Object, ByVal C As Object) As Integer
    '    SplitCount = IIf(B.ToString = "", 1, IIf(C.ToString = "", 2, 3))
    'End Function

    Public Function SplitCount(ByVal A As TextBox, ByVal B As TextBox, ByVal C As TextBox) As Integer
        'SplitCount = IIf(B.ToString = "", 1, IIf(C.ToString = "", 2, 3))
        SplitCount = IIf(B.Text = "", 1, IIf(C.Text = "", 2, 3))
    End Function
    Public Function GetSalesSplit(ByVal IA As String, ByVal iB As String, ByVal iC As String, ByVal Count as integer) As String
        Dim A As Double, B As Double, C As Double
        Dim dA As Boolean, dB As Boolean, DC As Boolean
        Dim X As Double

        AdjustSalesSplits(IA, iB, iC, Count)
        '  Count = IIf(Sales2 = "", 1, IIf(Sales3 = "", 2, 3))
        dA = SplitValue(IA) = 0
        dB = SplitValue(iB) = 0
        DC = SplitValue(iC) = 0

        If Count = 1 Then
            C = 0
            B = 0
            A = IIf(dA, 100, SplitValue(IA))
        ElseIf Count = 2 Then
            C = 0
            A = SplitValue(IA)
            B = SplitValue(iB)
            If dA And dB Then
                A = 50
                B = 50
            Else
                If dA Then A = 100 - B Else B = 100 - A
            End If
        Else 'Count = 3
            A = SplitValue(IA)
            B = SplitValue(iB)
            C = SplitValue(iC)
            If dA And dB And DC Then
                A = 33.33
                B = 33.33
                C = 33.33
            ElseIf dA And dB Then
                X = 100 - C
                A = X / 2
                B = X / 2
            ElseIf dA And DC Then
                X = 100 - B
                A = X / 2
                C = X / 2
            ElseIf dB And DC Then
                X = 100 - A
                B = X / 2
                C = X / 2
            ElseIf dA Then
                A = 100 - B - C
            ElseIf dB Then
                B = 100 - A - C
            ElseIf DC Then
                C = 100 - A - B
            End If
        End If
        GetSalesSplit = FormatSplit(A) & " " & FormatSplit(B) & " " & FormatSplit(C)
    End Function
    Public Sub ParseSalesSplit(ByVal SS As String, ByRef A As Double, ByRef B As Double, ByRef C As Double, ByVal Count as integer)
        Dim sA As String, sb As String, sC As String
        Dim L As Object
        SS = Trim(SS)
        If SS = "" Then
            A = 0
            B = 0
            C = 0
        Else
            L = Split(SS, " ")
            If UBound(L) >= 0 Then A = SplitValue(L(0))
            If UBound(L) >= 1 Then B = SplitValue(L(1))
            If UBound(L) >= 2 Then C = SplitValue(L(2))
        End If

        sA = A
        sb = B
        sC = C
        AdjustSalesSplits(sA, sb, sC, Count)
        A = SplitValue(sA)
        B = SplitValue(sb)
        C = SplitValue(sC)
    End Sub
    Public Function SplitValue(ByVal Split As String) As Double
        SplitValue = Val(Replace(Split, "%", ""))
    End Function
    Public Function FormatSplit(ByVal Split As Double) As String
        FormatSplit = Format(Split, "0.0#")
    End Function

End Module
