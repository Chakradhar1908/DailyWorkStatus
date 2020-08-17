Public Class cArBoyd
    Public Cash As Decimal
    Public mIntr As Double
    Public N As Integer
    Public DFP As Integer
    Private Const DPY As Integer = 360
    Private Const LIFE_MAX As Decimal = 10000
    Private Const AH_MAX As Decimal = 500
    Private Const IUI_MAX As Decimal = 750
    Private Const PR_MAX As Decimal = 10000

    Private Const PR_MIN As Decimal = 3
    Private Const LIFE_MIN As Decimal = 3

    Public bLife As Boolean
    Public bAH As Boolean
    Public bIUI As Boolean
    Public bProperty As Boolean

    Public LifeR As Decimal
    Public DSr As Double
    Public AHR As Decimal
    Public IUIr As Decimal
    Public PRr As Decimal
    Public PRCov As String
    Public PRCost As Decimal

    Const BuildTable As Boolean = False

    Public Property intr As Double
        Get
            intr = mIntr
        End Get
        Set(value As Double)
            mIntr = Math.Round(value, 4)
        End Set
    End Property

    Public ReadOnly Property LifePremium() As Decimal
        Get
            If Not bLife Then Exit Property
            If BuildTable Then
                LifePremium = MinArray(LifePremiumArray)
            Else
                LifePremium = CurrencyFormat(LP * MonthlyPayment * AgtN)
            End If
            If LIFE_MIN > 0 And LifePremium < LIFE_MIN Then LifePremium = LIFE_MIN
        End Get

    End Property

    Private Function LifePremiumArray() As Object
        'Dim A(1 To 16) As Variant, I as integer
        Dim A(0 To 15) As Object, I As Integer
        For I = 0 To 15
            A(I) = LifePremiumOptions(I)
        Next
        LifePremiumArray = A
    End Function

    Public ReadOnly Property LP() As Double
        Get
            On Error Resume Next
            LP = LPR * XO * NoverAgtNallover100
            LP = Math.Round(LP, 9)
        End Get
    End Property

    Public ReadOnly Property XO() As Double
        Get
            XO = ((N + (DFP - DPY / 12) / (DPY / 12)) / 12)
            XO = Math.Round(XO, 9)
        End Get
    End Property

    Private ReadOnly Property NoverAgtNallover100() As Double
        Get
            NoverAgtNallover100 = N / AgtN / 100
        End Get
    End Property

    Public ReadOnly Property LPR() As Double
        Get
            LPR = IIf(bLife, LifeR / (1 + ((N * DSr) / 24)), 0)
            LPR = Math.Round(LPR, 9)
        End Get
    End Property

    Public ReadOnly Property MonthlyPayment() As Decimal
        Get
            On Error Resume Next
            If BuildTable Then
                MonthlyPayment = MinArray(MonthlyPaymentArray)
            Else
                MonthlyPayment = Cash / ((1 - LP - Prp) * AgtN)
            End If
            MonthlyPayment = Trunc(MonthlyPayment)
        End Get
    End Property

    Private Function MonthlyPaymentArray() As Object
        'Dim A(1 To 16) As Variant, I as integer
        Dim A(0 To 15) As Object, I As Integer
        For I = 0 To 15
            A(I) = MonthlyPaymentOptions(I)
        Next
        MonthlyPaymentArray = A
    End Function

    Private Function MonthlyPaymentOptions(ByVal O As Integer) As Decimal
        Dim R As Decimal
        Select Case O
            'Case 1
            Case 0
                R = IIf(PRCov = "P", (Cash + PropertyPremiumOptions(O)) / ((1 - LP - AHP - IUIP) * AgtN), LP / ((1 - LP - AHP - IUIP - Prp) * AgtN))
            'Case 2
            Case 1
                R = IIf(PRCov = "P", (Cash + LifePremiumOptions(O) + PropertyPremiumOptions(O)) / ((1 - AHP - IUIP) * AgtN), (Cash + LifePremiumOptions(O)) / ((1 - AHP - IUIP - Prp) * AgtN))
            'Case 3
            Case 2
                R = IIf(PRCov = "P", (Cash + AHPremiumOptions(O) + PropertyPremiumOptions(O)) / ((1 - LP - IUIP) * AgtN), (LP + AHPremiumOptions(O)) / ((1 - LP - IUIP - Prp) * AgtN))
            'Case 4
            Case 3
                R = IIf(PRCov = "P", (Cash + IUIPremiumOptions(O) + PropertyPremiumOptions(O)) / ((1 - LP - AHP) * AgtN), (Cash + IUIP) / ((1 - LP - AHP - IUIP) * AgtN))
        End Select
        MonthlyPaymentOptions = Trunc(R, 2)
    End Function

    Private Function AHPremiumOptions(ByVal O As Integer) As Decimal
        Dim R As Decimal
        If Not bAH Or AHR = 0 Then Exit Function
        Select Case O
            'Case 1
            Case 0
                R = AHP * MonthlyPaymentOptions(O) * AgtN
            'Case 16
            Case 15
                R = AHR / 100 * N * AH_MAX
        End Select
        AHPremiumOptions = Math.Round(R, 2)
    End Function

    Private Function IUIPremiumOptions(ByVal O As Integer) As Decimal
        Dim R As Decimal
        If Not bIUI Or IUIr = 0 Then Exit Function
        Select Case O
            'Case 1, 2, 3, 5, 6, 8, 10, 13
            Case 0, 1, 2, 4, 5, 7, 9, 12
                R = IUIP * MonthlyPaymentOptions(O) * AgtN
            'Case 4, 7, 9, 11, 12, 14, 15, 16
            Case 3, 6, 8, 10, 11, 13, 14, 15
                R = IUIr / 100 * N * IUI_MAX
        End Select
        IUIPremiumOptions = Math.Round(R, 2)
    End Function

    Private Function PropertyPremiumOptions(ByVal O As Integer) As Decimal
        Dim R As Decimal
        If Not bProperty Or PRr = 0 Then Exit Function
        Select Case O
            'Case 1
            Case 0
                R = IIf(PRCov = "P", Prp * PRCost, Prp * MonthlyPayment * N)
            'Case 2, 3, 4, 6, 7, 9, 11
            Case 1, 2, 3, 5, 6, 8, 10
                R = IIf(PRCov = "P", Prp * PRCost, Prp * MonthlyPaymentOptions(O) * AgtN)
            'Case 5, 8, 10, 12, 13, 14, 15, 16
            Case 4, 7, 9, 11, 12, 13, 14, 15
                If Not (PRCov = "P" And PRCost = 0) Then
                    R = IIf(PRCov = "G", Prp * (AgtN / N) * PR_MAX, Prp * PR_MAX)
                End If
        End Select
        PropertyPremiumOptions = Math.Round(R, 2)
    End Function

    Public ReadOnly Property AHP() As Double
        Get
            On Error Resume Next
            AHP = IIf(bAH, AHR * NoverAgtNallover100, 0)
            AHP = Math.Round(AHP, 9)
        End Get
    End Property

    Public ReadOnly Property IUIP() As Double
        Get
            On Error Resume Next
            IUIP = IIf(bIUI, IUIr * NoverAgtNallover100, 0)
            IUIP = Math.Round(IUIP, 9)
        End Get
    End Property

    Public ReadOnly Property Prp() As Double
        Get
            On Error Resume Next
            Prp = IIf(PRCov = "P" And PRCost = 0, 0, IIf(bProperty, PRr * XO * NoverAgtNallover100, 0))
            Prp = Math.Round(Prp, 9)
        End Get
    End Property

    Public ReadOnly Property AgtN() As Double
        Get
            On Error Resume Next
            AgtN = (1 - Pow(1 / (1 + I), N)) / I
            AgtN = AgtN * (1 + I) / (1 + ((I * DFP) / (DPY / 12)))
            AgtN = Math.Round(AgtN, 9)
        End Get
    End Property

    Public ReadOnly Property I() As Double
        Get
            I = intr / 12
        End Get
    End Property

    Private Function LifePremiumOptions(ByVal O As Integer) As Decimal
        Dim R As Decimal
        If Not bLife Or LifeR = 0 Or LP < 0 Then Exit Function
        Select Case O
            'Case 1, 3, 4, 5, 9, 10, 14, 15
            Case 0, 2, 3, 4, 8, 9, 13, 14
                R = LP * MonthlyPaymentOptions(O) * AgtN
            'Case 2, 6, 7, 8, 11, 12, 13, 16
            Case 1, 5, 6, 7, 10, 11, 12, 15
                R = LP * LIFE_MAX * AgtN / N
        End Select
        LifePremiumOptions = Math.Round(R, 2)
    End Function

    Public ReadOnly Property AHPremium() As Decimal
        Get
            If Not bAH Then Exit Property
            If BuildTable Then
                AHPremium = MinArray(AHPremiumArray)
            Else
                AHPremium = CurrencyFormat(AHP * MonthlyPayment * AgtN)
            End If
        End Get
    End Property

    Private Function AHPremiumArray() As Object
        'Dim A(1 To 16) As Variant, I as integer
        Dim A(0 To 15) As Object, I As Integer

        'For I = 1 To 16
        For I = 0 To 15
            A(I) = AHPremiumOptions(I)
        Next
        AHPremiumArray = A
    End Function

    Public ReadOnly Property PropertyPremium() As Decimal
        Get
            If Not bProperty Then Exit Property
            If BuildTable Then
                PropertyPremium = MinArray(PropertyPremiumArray)
            Else
                PropertyPremium = CurrencyFormat(Prp * MonthlyPayment * AgtN)  ' AH & IUI capped
                '    PropertyPremium = CurrencyFormat(Prp * TotalOfPayments)
            End If
            If PR_MIN > 0 And PropertyPremium < PR_MIN Then PropertyPremium = PR_MIN
        End Get
    End Property

    Private Function PropertyPremiumArray() As Object
        'Dim A(1 To 16) As Variant, I as integer
        Dim A(0 To 15) As Object, I As Integer

        'For I = 1 To 16
        For I = 0 To 15
            A(I) = PropertyPremiumOptions(I)
        Next
        PropertyPremiumArray = A
    End Function

    Public ReadOnly Property AmountFinanced() As Decimal
        Get
            AmountFinanced = CurrencyFormat(Cash) + TotalPremium
        End Get
    End Property

    Public ReadOnly Property TotalPremium() As Decimal
        Get
            TotalPremium = LifePremium + AHPremium + IUIPremium + PropertyPremium
        End Get
    End Property

    Public ReadOnly Property IUIPremium() As Decimal
        Get
            If Not bIUI Then Exit Property
            If BuildTable Then
                IUIPremium = MinArray(IUIPremiumArray)
            Else
                IUIPremium = CurrencyFormat(IUIP * MonthlyPayment * AgtN)
            End If
        End Get
    End Property

    Private Function IUIPremiumArray() As Object
        'Dim A(1 To 16) As Variant, I as integer
        Dim A(0 To 15) As Object, I As Integer
        'For I = 1 To 16
        For I = 0 To 15
            A(I) = IUIPremiumOptions(I)
        Next
        IUIPremiumArray = A
    End Function

    Public ReadOnly Property FinanceCharge() As Decimal
        Get
            FinanceCharge = TotalOfPayments - AmountFinanced
        End Get
    End Property

    Public ReadOnly Property TotalOfPayments() As Decimal
        Get
            TotalOfPayments = CurrencyFormat(MonthlyPayment * N)
        End Get
    End Property

End Class
