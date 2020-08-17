Public Class CLyndonLife
    Private mLoanAmount As Single
    Private mAmountFinanced As Single
    Private mTermOfLife As Integer
    Private mTermOfAcc As Single
    Private mTermOfLoanInitial As Single
    Private mDaysToFirstPayment As Single
    Private mMontlyInterestRate As Single
    Private mPremLifeRate As Single
    Private mPremAHRate As Single
    Private mSP As Single 'local copy

    Public AHBenefitType As BenefitTypes
    Public AHBenefitRate As RateTypes
    Public PropInsuranceRate As PropInsuranceTypes
    Public Periods As Single

    Public LifeInsuranceOn As Boolean
    Public AHInsuranceOn As Boolean
    Public PropertyInsuranceOn As Boolean

    Public LifeInsurance As Single
    Public AHInsurance As Single
    Public PropertyInsurance As Single
    Public TotalInsurance As Single

    Public TotalInsuranceDaily As Single
    Public TotalInsuranceMonthly As Single

    Public APR As Single

    Public Property I As Single
        Get
            I = mMontlyInterestRate
        End Get
        Set(value As Single)
            mMontlyInterestRate = value
        End Set
    End Property

    Public Property M As Integer
        Get
            M = mTermOfLife
        End Get
        Set(value As Integer)
            mTermOfLife = value
        End Set
    End Property

    Public Property LFSPR() As Single
        Get
            LFSPR = mPremLifeRate
        End Get
        Set(value As Single)
            mPremLifeRate = value
        End Set
    End Property

    Public Property LA() As Single
        Get
            LA = mLoanAmount
        End Get
        Set(value As Single)
            mLoanAmount = value
        End Set
    End Property

    Public Sub Calculate()
        On Error GoTo AnError
        AHInsurance = 0
        PropertyInsurance = 0
        TotalInsurance = 0
        AF = GetAmountFinanced(LA)
        LifeInsurance = LP2()
        AHInsurance = ((TotalPayments) / 100) * creditAHRate(M, AHBenefitRate, AHBenefitType)
        PropertyInsurance = (M / Periods) * ((TotalPayments) / 100) * propInsurance(PropInsuranceRate)

        If Not LifeInsuranceOn Then LifeInsurance = 0
        If Not AHInsuranceOn Then AHInsurance = 0
        If Not PropertyInsuranceOn Then PropertyInsurance = 0

        TotalInsurance = LifeInsurance + AHInsurance + PropertyInsurance
        TotalInsuranceDaily = Math.Round(TotalInsurance / 360, 2)
        TotalInsuranceMonthly = 30 * TotalInsuranceDaily
        '.APR = getAPRDirectRatio()
        APR = getAPRNRatio()
        Exit Sub
AnError:
        LifeInsurance = 0
        AHInsurance = 0
        PropertyInsurance = 0
        TotalInsurance = 0
        APR = 0
        TotalInsurance = 0
        Exit Sub
    End Sub

    Public ReadOnly Property FinanceCharge() As Single
        Get
            On Error Resume Next
            FinanceCharge = M * MontlyLoanPayment - AF
            FinanceCharge = Math.Round(FinanceCharge, 2)
        End Get
    End Property

    Public ReadOnly Property MontlyLoanPayment() As Single
        Get
            On Error GoTo AnError
            MontlyLoanPayment _
                    = AF _
                      * (I / Periods) _
                      / (1 - 1 / ((1 + I / Periods) ^ M))
            MontlyLoanPayment = Math.Round(MontlyLoanPayment, 2)
            Exit Property
AnError:
            MontlyLoanPayment = 0
            Exit Property
        End Get
    End Property

    Public Property AF() As Single
        Get
            AF = mAmountFinanced
        End Get
        Set(value As Single)
            mAmountFinanced = value
        End Set
    End Property

    Public Function GetAmountFinanced(ByRef LoanAmount As Single) As Single
        On Error GoTo AnError
        Dim P As Single : P = LoanAmount
        Dim N As Integer : N = M
        Dim Periods As Integer : Periods = 12
        Dim LPM As Integer : LPM = 3
        Dim I As Single : I = I
        Dim LFSPR As Single : LFSPR = LFSPR
        Dim Km As Single : Km = (I / Periods) / (1 - 1 / ((1 + I / Periods) ^ N))

        Dim V As Single
        V = (1 / (1 + (I / Periods)))
        Dim AA As Single
        AA = (1 - (V ^ N)) / (I / Periods)
        Dim Klp As Single
        Klp = (LFSPR * N / 600) * (N - AA + LPM * N * I / Periods) / ((N + 1) * (I / Periods) * AA)
        Dim Kah As Single
        Kah = creditAHRate(M, AHBenefitRate, AHBenefitType) / 100
        Dim Kpi As Single
        Kpi = (N / Periods) * propInsurance(PropInsuranceRate) / 100


        If (LifeInsuranceOn = False) Then Klp = 0
        If (AHInsuranceOn = False) Then Kah = 0
        If (PropertyInsuranceOn = False) Then Kpi = 0
        Dim Kti As Single : Kti = (Klp + Km * N * Kah + Km * N * Kpi)

        GetAmountFinanced = P / (1 - Kti)
        Exit Function
AnError:
        GetAmountFinanced = 0
        Exit Function
    End Function

    Public Function LP2(Optional ByRef N As Single = -1, Optional ByRef M As Integer = 3) As Single
        'M = 0
        On Error GoTo AnError
        If (N = -1) Then N = N
        Dim Periods As Single : Periods = 12
        LP2 _
          = (LFSPR * M / 600) _
              * (M - AA() + M * M * I / Periods) _
              / ((M + 1) * (I / Periods) * AA()) _
              * AF
        Exit Function
AnError:
        Debug.Print("Error in CNetPayout LP")
        LP2 = -9999
    End Function

    Public ReadOnly Property TotalPayments() As Single
        Get
            TotalPayments = AF + FinanceCharge
        End Get
    End Property

    Public Function getAPRNRatio() As Single
        Dim P As Single : P = AF
        Dim I As Single : I = I
        Dim M As Integer : M = Periods
        Dim N As Integer : N = M
        Dim C As Single : C = FinanceCharge '  ((Principle * InterestRate) / PeriodsPerYear * Months)
        On Error Resume Next
        getAPRNRatio = (M * (95 * N + 9) * C) / (12 * N * (N + 1) * (4 * P + C))
        'Debug.Print "getAPRNRatio APR Check : " & MontlyPaymentUsingCompound(P, getAPRNRatio, M, N)
        'Debug.Print "APR Check : " & MontlyLoanPayment
    End Function

    Public Function AA(Optional ByRef N As Single = -1) As Single
        On Error GoTo AnError
        Dim Periods As Single : Periods = 12
        Dim V As Single : V = (1 / (1 + (I / Periods)))
        AA = (1 - (V ^ M)) / (I / Periods)
        Exit Function
AnError:
        Debug.Print("Error in CNetPayout A")
        AA = -9999
    End Function

End Class
