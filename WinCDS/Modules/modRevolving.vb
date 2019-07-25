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
        Dim I as integer, J as integer
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
    Public ReadOnly Property Installment() As Boolean
        Get
            Installment = InstallmentLicenseValid(InstallmentLicense)
        End Get
    End Property
    Public Function InstallmentLicenseValid(ByVal S As String) As Boolean
        'InstallmentLicenseValid = IsIn(S, LICENSE_INSTALLMENT, "TEST")
    End Function

End Module
