Public Class CLyndonLife
    Private mLoanAmount As Single
    Private mAmountFinanced As Single
    Private mTermOfLife As Long
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
End Class
