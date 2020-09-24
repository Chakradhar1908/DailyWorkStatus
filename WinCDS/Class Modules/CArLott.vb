Public Class CArLott
    Public PrevBalance As Decimal
    Public NewSale As Decimal
    Public Deposit As Decimal
    Public DocFee As Decimal
    Public DeferredMonths As Integer
    Public SimpleInterestRate As Double
    Public Months As Integer
    Public ApplyFinanceChargeSalesTax As Boolean
    Public WeeklyPayments As Boolean
    Public RoundUp As Boolean
    Public JointLife As Boolean
    Public JointLifeReducing As Boolean

    Public LifeInsuranceOn As TriState
    Public AccidentInsuranceOn As TriState
    Public PropertyInsuranceOn As TriState
    Public UserLifePremium As Decimal, UserAccidentPremium As Decimal, UserPropertyPremium As Decimal

    Private mDeferredInterest As Decimal
    Private mAPR As Double, mFinanceCharges As Decimal, mAmountFinanced As Decimal
    Private mFinanceChargeSalesTax As Decimal

    Public Sub Calculate()
        Dim CAF As Decimal        'Current Amount Financed
        Dim oCAF As Decimal       ' the old value
        Dim Count As Integer
        oCAF = -1
        CAF = Base

        Do While Count < 5 Or CAF <> oCAF
            oCAF = CAF
            CAF = TotalAF(CAF)
            RecalculateFinanceCharges(CAF)
            Count = Count + 1
            If Count >= 500000 Then Exit Do   ' sanity
        Loop
        mAmountFinanced = CAF
    End Sub

    Private ReadOnly Property Base() As Decimal
        Get
            Base = PrevBalance + NewSale - Deposit + DocFee
        End Get
    End Property

    Private Function TotalAF(ByVal CAF As Decimal) As Decimal
        TotalAF = Base + LifeInsurancePremium(CAF) + AccidentInsurancePremium(CAF) + PropertyInsurancePremium(CAF) + FinanceCharges + FinanceChargeSalesTax
        TotalAF = GetPrice(CurrencyFormat(TotalAF))
    End Function

    Private Sub RecalculateFinanceCharges(ByVal CAF As Decimal)
        mDeferredInterest = (SimpleInterestRate * FinanceSubTotal(CAF)) / 12 * DeferredMonths
        mFinanceCharges = ((FinanceSubTotal(CAF) * SimpleInterestRate) / 12 * Months) + DeferredInterest

        If ApplyFinanceChargeSalesTax Then
            mFinanceChargeSalesTax = CurrencyFormat(StoreSettings.SalesTax * FinanceCharges)
        Else
            mFinanceChargeSalesTax = CurrencyFormat(0)
        End If

        If (FinanceCharges) <> 0 And (Months + DeferredMonths) <> -1 And FinanceSubTotal(CAF) <> 0 Then
            mAPR = CalculateAPR(FinanceSubTotal(CAF), FinanceCharges, Months, DeferredMonths)
        Else
            mAPR = 0
        End If
    End Sub

    Public Function FinanceSubTotal(Optional ByVal CAF As Decimal = -1) As Decimal
        If CAF < 0 Then CAF = AmountFinanced
        FinanceSubTotal = Base + LifeInsurancePremium(CAF) + AccidentInsurancePremium(CAF) + PropertyInsurancePremium(CAF)
    End Function

    Public ReadOnly Property AmountFinanced() As Decimal
        Get
            AmountFinanced = GetPrice(CurrencyFormat(mAmountFinanced))
        End Get
    End Property

    Public ReadOnly Property DeferredInterest() As Decimal
        Get
            DeferredInterest = GetPrice(CurrencyFormat(mDeferredInterest))
        End Get
    End Property

    Public ReadOnly Property LifeInsurancePremium(Optional ByVal vAmountFinanced As Decimal = -1) As Decimal
        Get
            Dim Rate As Double
            Select Case LifeInsuranceOn
                Case vbTrue
                    If Thorntons Then
                        'rate = AmericanHeritage_Life(1, Months, Not JointLife, False)
                        Rate = LifeInsuranceRate
                        '        If JointLife Then
                        '          Rate = 1 / 100#
                        '        Else
                        '         Rate = 0.66 / 100#
                        '        End If
                        LifeInsurancePremium = Rate * IIf(vAmountFinanced < 0, AmountFinanced, vAmountFinanced) * Years
                    Else
                        LifeInsurancePremium = LifeInsuranceRate * IIf(vAmountFinanced < 0, AmountFinanced, vAmountFinanced) * Years
                    End If
                Case vbFalse : LifeInsurancePremium = 0
                Case Else : LifeInsurancePremium = UserLifePremium
            End Select
        End Get
    End Property

    Private ReadOnly Property Years() As Double
        Get
            Years = CDbl(Months) / 12.0#
        End Get
    End Property

    Private ReadOnly Property Thorntons() As Boolean
        Get
            Thorntons = IsThorntons
        End Get
    End Property

    Public ReadOnly Property LifeInsuranceRate() As Double
        Get
            If IsThorntons Then
                LifeInsuranceRate = IIf(Not JointLife, 0.66, 1) / 100.0#
            Else
                If JointLife Then
                    '    If JointLifeReducing Then
                    LifeInsuranceRate = 2.79 / 100.0#
                    '    Else
                    '      LifeInsuranceRate = 1.39 / 100#
                    '    End If
                Else
                    LifeInsuranceRate = 1.6 / 100.0#
                End If
            End If
            ' Debug.Print "cARLott.LifeInsuranceRate=" & LifeInsuranceRate
        End Get
    End Property

    Public ReadOnly Property AccidentInsurancePremium(Optional ByVal vAmountFinanced As Decimal = -1) As Decimal
        Get
            Select Case AccidentInsuranceOn
                Case vbTrue
                    If Thorntons Then
                        '        AccidentInsurancePremium = AmericanHeritage_Acc(1, Months, True, 7) * IIf(vAmountFinanced < 0, AmountFinanced, vAmountFinanced)
                        AccidentInsurancePremium = AccidentInsuranceRate * IIf(vAmountFinanced < 0, AmountFinanced, vAmountFinanced)
                    Else
                        AccidentInsurancePremium = AccidentInsuranceRate * IIf(vAmountFinanced < 0, AmountFinanced, vAmountFinanced) * Years
                    End If
                Case vbFalse : AccidentInsurancePremium = 0
                Case Else : AccidentInsurancePremium = UserAccidentPremium
            End Select
        End Get
    End Property

    Public ReadOnly Property AccidentInsuranceRate() As Double
        Get
            If IsThorntons Then
                AccidentInsuranceRate = ThorntonsAccidentRate
            Else
                Select Case Months
                    Case 1 To 12 : AccidentInsuranceRate = 3.0# / 100.0#
                    Case 13 To 24 : AccidentInsuranceRate = 3.8 / 100.0#
                    Case 25 To 36 : AccidentInsuranceRate = 4.6 / 100.0#
                    Case Else
                        '      MsgBox "No A & H formula available for greater than 36 months!", vbExclamation, ProgramMessageTitle
                        AccidentInsuranceRate = 0#
                End Select
            End If
            '   Debug.Print "cARLott.ThorntonsAccidentRate=" & AccidentInsuranceRate
        End Get
    End Property

    Public ReadOnly Property PropertyInsurancePremium(Optional ByVal vAmountFinanced As Decimal = -1) As Decimal
        Get
            Select Case PropertyInsuranceOn
                Case vbTrue : PropertyInsurancePremium = PropertyInsuranceRate * IIf(vAmountFinanced < 0, AmountFinanced, vAmountFinanced) * Years
                Case vbFalse : PropertyInsurancePremium = 0
                Case Else : PropertyInsurancePremium = UserPropertyPremium
            End Select
        End Get
    End Property

    Public ReadOnly Property FinanceCharges() As Decimal
        Get
            FinanceCharges = GetPrice(CurrencyFormat(mFinanceCharges))
        End Get
    End Property

    Public ReadOnly Property FinanceChargeSalesTax() As Decimal
        Get
            FinanceChargeSalesTax = GetPrice(CurrencyFormat(mFinanceChargeSalesTax))
        End Get
    End Property

    Private Function ThorntonsAccidentRate() As Double
        Select Case Months
            Case 1 : ThorntonsAccidentRate = 0.4
            Case 2 : ThorntonsAccidentRate = 0.81
            Case 3 : ThorntonsAccidentRate = 1.12
            Case 4 : ThorntonsAccidentRate = 1.36
            Case 5 : ThorntonsAccidentRate = 1.55
            Case 6 : ThorntonsAccidentRate = 1.71
            Case 7 : ThorntonsAccidentRate = 1.86
            Case 8 : ThorntonsAccidentRate = 1.98
            Case 9 : ThorntonsAccidentRate = 2.1
            Case 10 : ThorntonsAccidentRate = 2.2
            Case 11 : ThorntonsAccidentRate = 2.3
            Case 12 : ThorntonsAccidentRate = 2.39
            Case 13 : ThorntonsAccidentRate = 2.47
            Case 14 : ThorntonsAccidentRate = 2.55
            Case 15 : ThorntonsAccidentRate = 2.62
            Case 16 : ThorntonsAccidentRate = 2.69
            Case 17 : ThorntonsAccidentRate = 2.76
            Case 18 : ThorntonsAccidentRate = 2.82
            Case 19 : ThorntonsAccidentRate = 2.88
            Case 20 : ThorntonsAccidentRate = 2.94
            Case 21 : ThorntonsAccidentRate = 3
            Case 22 : ThorntonsAccidentRate = 3.05
            Case 23 : ThorntonsAccidentRate = 3.11
            Case 24 : ThorntonsAccidentRate = 3.16
            Case 25 : ThorntonsAccidentRate = 3.21
            Case 26 : ThorntonsAccidentRate = 3.26
            Case 27 : ThorntonsAccidentRate = 3.31
            Case 28 : ThorntonsAccidentRate = 3.35
            Case 29 : ThorntonsAccidentRate = 3.4
            Case 30 : ThorntonsAccidentRate = 3.44
            Case 31 : ThorntonsAccidentRate = 3.49
            Case 32 : ThorntonsAccidentRate = 3.53
            Case 33 : ThorntonsAccidentRate = 3.57
            Case 34 : ThorntonsAccidentRate = 3.61
            Case 35 : ThorntonsAccidentRate = 3.65
            Case 36 : ThorntonsAccidentRate = 3.69
            Case 37 : ThorntonsAccidentRate = 3.73
            Case 38 : ThorntonsAccidentRate = 3.77
            Case 39 : ThorntonsAccidentRate = 3.81
            Case 40 : ThorntonsAccidentRate = 3.85
            Case 41 : ThorntonsAccidentRate = 3.89
            Case 42 : ThorntonsAccidentRate = 3.92
            Case 43 : ThorntonsAccidentRate = 3.96
            Case 44 : ThorntonsAccidentRate = 3.99
            Case 45 : ThorntonsAccidentRate = 4.03
            Case 46 : ThorntonsAccidentRate = 4.06
            Case 47 : ThorntonsAccidentRate = 4.1
            Case 48 : ThorntonsAccidentRate = 4.13
            Case 49 : ThorntonsAccidentRate = 4.17
            Case 50 : ThorntonsAccidentRate = 4.2
            Case 51 : ThorntonsAccidentRate = 4.23
            Case 52 : ThorntonsAccidentRate = 4.27
            Case 53 : ThorntonsAccidentRate = 4.3
            Case 54 : ThorntonsAccidentRate = 4.33
            Case 55 : ThorntonsAccidentRate = 4.36
            Case 56 : ThorntonsAccidentRate = 4.39
            Case 57 : ThorntonsAccidentRate = 4.42
            Case 58 : ThorntonsAccidentRate = 4.45
            Case 59 : ThorntonsAccidentRate = 4.49
            Case 60 : ThorntonsAccidentRate = 4.52
            Case 61 : ThorntonsAccidentRate = 4.55
            Case 62 : ThorntonsAccidentRate = 4.59
            Case 63 : ThorntonsAccidentRate = 4.62
            Case 64 : ThorntonsAccidentRate = 4.65
            Case 65 : ThorntonsAccidentRate = 4.69
            Case 66 : ThorntonsAccidentRate = 4.72
            Case 67 : ThorntonsAccidentRate = 4.75
            Case 68 : ThorntonsAccidentRate = 4.78
            Case 69 : ThorntonsAccidentRate = 4.82
            Case 70 : ThorntonsAccidentRate = 4.85
            Case 71 : ThorntonsAccidentRate = 4.88
            Case 72 : ThorntonsAccidentRate = 4.92
            Case 73 : ThorntonsAccidentRate = 4.9525
            Case 74 : ThorntonsAccidentRate = 4.985
            Case 75 : ThorntonsAccidentRate = 5.0175
            Case 76 : ThorntonsAccidentRate = 5.05
            Case 77 : ThorntonsAccidentRate = 5.0825
            Case 78 : ThorntonsAccidentRate = 5.115
            Case 79 : ThorntonsAccidentRate = 5.1475
            Case 80 : ThorntonsAccidentRate = 5.18
            Case 81 : ThorntonsAccidentRate = 5.2125
            Case 82 : ThorntonsAccidentRate = 5.245
            Case 83 : ThorntonsAccidentRate = 5.2775
            Case 84 : ThorntonsAccidentRate = 5.31
        End Select
        ThorntonsAccidentRate = ThorntonsAccidentRate / 100
    End Function

    Public ReadOnly Property PropertyInsuranceRate() As Double
        Get
            '  PropertyInsuranceRate = 3.35 / 100#
            If IsThorntons Then
                Select Case Months
                    Case 1 : PropertyInsuranceRate = 0.16667
                    Case 2 : PropertyInsuranceRate = 0.33333
                    Case 3 : PropertyInsuranceRate = 0.49999
                    Case 4 : PropertyInsuranceRate = 0.66665
                    Case 5 : PropertyInsuranceRate = 0.83331
                    Case 6 : PropertyInsuranceRate = 0.99997
                    Case 7 : PropertyInsuranceRate = 1.16663
                    Case 8 : PropertyInsuranceRate = 1.33329
                    Case 9 : PropertyInsuranceRate = 1.49995
                    Case 10 : PropertyInsuranceRate = 1.66661
                    Case 11 : PropertyInsuranceRate = 1.83327
                    Case 12 : PropertyInsuranceRate = 2
                    Case 13 : PropertyInsuranceRate = 2.16666
                    Case 14 : PropertyInsuranceRate = 2.33332
                    Case 15 : PropertyInsuranceRate = 2.49998
                    Case 16 : PropertyInsuranceRate = 2.66664
                    Case 17 : PropertyInsuranceRate = 2.8333
                    Case 18 : PropertyInsuranceRate = 2.99996
                    Case 19 : PropertyInsuranceRate = 3.16662
                    Case 20 : PropertyInsuranceRate = 3.33328
                    Case 21 : PropertyInsuranceRate = 3.49994
                    Case 22 : PropertyInsuranceRate = 3.6666
                    Case 23 : PropertyInsuranceRate = 3.83326
                    Case 24 : PropertyInsuranceRate = 4
                    Case 25 : PropertyInsuranceRate = 4.16666
                    Case 26 : PropertyInsuranceRate = 4.33332
                    Case 27 : PropertyInsuranceRate = 4.49998
                    Case 28 : PropertyInsuranceRate = 4.66664
                    Case 29 : PropertyInsuranceRate = 4.8333
                    Case 30 : PropertyInsuranceRate = 4.99996
                    Case 31 : PropertyInsuranceRate = 5.16662
                    Case 32 : PropertyInsuranceRate = 5.33328
                    Case 33 : PropertyInsuranceRate = 5.49994
                    Case 34 : PropertyInsuranceRate = 5.6666
                    Case 35 : PropertyInsuranceRate = 5.83326
                    Case 36 : PropertyInsuranceRate = 6
                    Case 37 : PropertyInsuranceRate = 6.16666
                    Case 38 : PropertyInsuranceRate = 6.33332
                    Case 39 : PropertyInsuranceRate = 6.49998
                    Case 40 : PropertyInsuranceRate = 6.66664
                    Case 41 : PropertyInsuranceRate = 6.8333
                    Case 42 : PropertyInsuranceRate = 6.99996
                    Case 43 : PropertyInsuranceRate = 7.16662
                    Case 44 : PropertyInsuranceRate = 7.33328
                    Case 45 : PropertyInsuranceRate = 7.49994
                    Case 46 : PropertyInsuranceRate = 7.6666
                    Case 47 : PropertyInsuranceRate = 7.83326
                    Case 48 : PropertyInsuranceRate = 8
                    Case 49 : PropertyInsuranceRate = 8.16666
                    Case 50 : PropertyInsuranceRate = 8.33332
                    Case 51 : PropertyInsuranceRate = 8.49998
                    Case 52 : PropertyInsuranceRate = 8.66664
                    Case 53 : PropertyInsuranceRate = 8.8333
                    Case 54 : PropertyInsuranceRate = 8.99996
                    Case 55 : PropertyInsuranceRate = 9.16662
                    Case 56 : PropertyInsuranceRate = 9.33328
                    Case 57 : PropertyInsuranceRate = 9.49994
                    Case 58 : PropertyInsuranceRate = 9.6666
                    Case 59 : PropertyInsuranceRate = 9.83326
                    Case 60 : PropertyInsuranceRate = 10
                End Select
            Else
                PropertyInsuranceRate = 3.24 / 100.0#
            End If

            'Debug.Print "cARLott.PropertyInsuranceRate=" & PropertyInsuranceRate
        End Get
    End Property

    Public ReadOnly Property APR() As Double
        Get
            APR = mAPR
        End Get
    End Property

    Public Function MonthlyPayment() As Decimal
        Dim Op As Decimal
        If Months = 0 Then Exit Function
        If WeeklyPayments Then
            MonthlyPayment = (AmountFinanced) / (Months * 4)
        Else
            MonthlyPayment = (AmountFinanced) / Months
        End If

        If RoundUp Then
            Op = MonthlyPayment
            MonthlyPayment = MonthlyPayment - ((MonthlyPayment - CInt(MonthlyPayment)))
            If MonthlyPayment < Op Then MonthlyPayment = MonthlyPayment + 1
        End If
    End Function
End Class
