Module modRevolving
    Public Structure PaymentRecord
        Dim SaleNo As String
        Dim Amount As Decimal
    End Structure

    Public Function AddRevolvingSuffix(ByVal ArNo As String, Optional ByVal AllowMultiple As Boolean = False) As String
        If IsRevolvingCharge(ArNo) Then
            AddRevolvingSuffix = ArNo & IIf(AllowMultiple, RevolvingSuffixLetter, "")
        Else
            AddRevolvingSuffix = ArNo & RevolvingSuffix
        End If
    End Function

    Public Function IsRevolvingCharge(ByVal ArNo As String) As Boolean
        '  Dim Suf As String
        '  Suf = RevolvingSuffix()
        '  IsRevolvingCharge = ModifiedRevolvingChargeEnabled() And (Right(ArNo, Len(Suf)) = Suf)

        ' Rewritten to allow "RRRRRRR" multiple accounts per customer MJK 20140914
        Dim I As Integer, J As Integer
        If ArNo = "" Then Exit Function
        If Not ModifiedRevolvingChargeEnabled() Then Exit Function
        I = InStrRev(ArNo, " ")
        For J = I + 1 To Len(ArNo)
            If Mid(ArNo, J, 1) <> RevolvingSuffixLetter Then Exit Function
        Next
        IsRevolvingCharge = True
    End Function

    Public ReadOnly Property RevolvingSuffixLetter() As String
        Get
            RevolvingSuffixLetter = "R"
        End Get
    End Property

    Public ReadOnly Property RevolvingSuffix() As String
        Get
            RevolvingSuffix = " " & RevolvingSuffixLetter
        End Get
    End Property

    Public ReadOnly Property ModifiedRevolvingChargeEnabled() As Boolean
        Get
            ModifiedRevolvingChargeEnabled = ModifiedRevolvingChargeAllowed And StoreSettings.ModifiedRevolvingCharge
        End Get
    End Property

    Public ReadOnly Property ModifiedRevolvingChargeAllowed() As Boolean
        Get
            ModifiedRevolvingChargeAllowed = Installment()  ' Valid Installment license required 'IsDevelopment() 'Or isevridges()
        End Get
    End Property

    Public Function CalculateRevolvingPayment(ByRef NewBalance As Decimal, Optional ByRef RoundUp As Boolean = False, Optional ByRef Months As Long = 0) As Decimal
        ' Future store settings may change this calculation.  Evridge's wants 1/3 of the balance due monthly.
        ' Balance due is complicated.  See RevolvingCurrentFinancedAmount for details.
        Dim Portion As Decimal
        If Months > 0 Then
            Portion = NewBalance / Months
        Else
            Portion = NewBalance * RevolvingMinimumPaymentPercent
        End If
        If RoundUp Then
            CalculateRevolvingPayment = Math.Round(Portion + 0.49, 0)
        Else
            CalculateRevolvingPayment = Math.Round(Portion + 0.0049, 2)
        End If
    End Function

    Public Function RevolvingMinimumPaymentPercent() As Double
        'RevolvingMinimumPaymentPercent = 1 / 3
        RevolvingMinimumPaymentPercent = StoreSettings.ModifiedRevolvingMinPmt
        If RevolvingMinimumPaymentPercent > 1 Then RevolvingMinimumPaymentPercent = RevolvingMinimumPaymentPercent / 100
    End Function

End Module
