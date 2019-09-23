Public Class cArTreehouse
    Public CA As Decimal
    Public Ani As Double

    Public N As Long              'Term of loan, and life, A&H and Property coverage
    Private mNu As Long           'Term of the IUI coverage  (Must be equal to n, unless n is greater than 60, then nu is 60.)
    Public JointLife As Boolean

    Public bHasLife As Boolean
    Public bHasAcci As Boolean
    Public bHasProp As Boolean
    Public bHasIUI As Boolean

    Public UserLifePremium As Decimal
    Public UserAcciPremium As Decimal
    Public UserPropPremium As Decimal
    Public UserIUIPremium As Decimal

    'Public SPL as decimal        'Life rate for the term of life insurance per $100 of initial insured indebtedness
    'Public SPD as decimal        'Disability rate for the term of disability insurance per $100 of initial insured indebtedness
    Private mSPG As Decimal        'Property rate per $1000 for the term of insurance
    Private mIUSP As Decimal       'prima facie IUI rate per $100

    Public OD As Long ' Days to first payment

    Public Property Nu As Integer
        Get
            Nu = mNu
        End Get
        Set(value As Integer)
            mNu = FitRange(0, value, 60)
        End Set
    End Property

    Public Property SPG As Decimal
        Get
            SPG = IIf(bHasProp, mSPG, 0)
        End Get
        Set(value As Decimal)
            mSPG = value
        End Set
    End Property

    Public Property IUSP As Decimal
        Get
            IUSP = IIf(bHasIUI, mIUSP, 0)
        End Get
        Set(value As Decimal)
            mIUSP = value
        End Set
    End Property

    Public ReadOnly Property LifePremium() As Decimal
        Get
            '  LifePremium =      (Sp * N * (N + (OD - 30) / 30) / 1200) * (CA / (an / odf - ((Sp / 1200 * (N + (OD - 30) / 30) * N) + (iusp * Nu / 100) + (spdd * N / 100) + (spp * N / 100))))
            If UserLifePremium <> 0 Then LifePremium = UserLifePremium : Exit Property
            LifePremium = Math.Round((Sp * N * (N + (OD - 30) / 30) / 1200) * (CA / (AN / ODF - ((Sp / 1200 * (N + (OD - 30) / 30) * N) + (IUSP * Nu / 100) + (SPDD * N / 100) + (SPP * N / 1200 * ((N + (OD - 30) / 30)))))), 2)
        End Get
    End Property

    Public ReadOnly Property SPP() As Decimal
        Get
            SPP = SPG
        End Get
    End Property

    Public ReadOnly Property SPDD() As Decimal
        Get
            SPDD = SPD
        End Get
    End Property

    Public ReadOnly Property SPD() As Double
        Get
            If Not bHasAcci Then Exit Property
            SPD = DisabilityRate(N)
        End Get
    End Property

    Public ReadOnly Property ODF() As Double
        Get
            ODF = 1 + I * (OD - 30) / 30
        End Get
    End Property

    Public ReadOnly Property AN() As Double
        Get
            If I = 0 Then Exit Property
            AN = Math.Round(IIf(Ani = 0, N, (1 - ((1 + I) ^ -N)) / I), 8)
        End Get
    End Property

    Public ReadOnly Property I() As Double
        Get
            I = Ani / 12
        End Get
    End Property

    Public ReadOnly Property Sp() As Decimal
        Get
            '  Sp = IIf(ListedMonthly, SPL * 12 / N, SPL)
            Sp = SPL * 12 / N
        End Get
    End Property

    Public ReadOnly Property SPL() As Double
        Get
            If Not bHasLife Then Exit Property
            SPL = LifeRate(N, JointLife)
        End Get
    End Property

    Public ReadOnly Property LifeRate(ByVal N As Long, ByVal Joint As Boolean) As Double
        Get
            Dim A As Double
            If IsMcClure Then LifeRate = LifeRate_Table(N, Joint) : Exit Property

            A = IIf(Not Joint, 0.319, 0.477)
            LifeRate = Math.Round(N * (1 / (1 + ((0.035 * A) / 24))) * A / 12, 3)
        End Get
    End Property

    Public ReadOnly Property LifeRate_Table(ByVal N As Long, ByVal Joint As Boolean) As Double
        Get
            Select Case N
    ' Copied from Excel, so there are no errors..
                Case 1 : LifeRate_Table = IIf(Not Joint, 0.027, 0.04)
                Case 2 : LifeRate_Table = IIf(Not Joint, 0.053, 0.079)
                Case 3 : LifeRate_Table = IIf(Not Joint, 0.079, 0.119)
                Case 4 : LifeRate_Table = IIf(Not Joint, 0.106, 0.158)
                Case 5 : LifeRate_Table = IIf(Not Joint, 0.132, 0.197)
                Case 6 : LifeRate_Table = IIf(Not Joint, 0.158, 0.236)
                Case 7 : LifeRate_Table = IIf(Not Joint, 0.184, 0.275)
                Case 8 : LifeRate_Table = IIf(Not Joint, 0.21, 0.314)
                Case 9 : LifeRate_Table = IIf(Not Joint, 0.236, 0.353)
                Case 10 : LifeRate_Table = IIf(Not Joint, 0.262, 0.392)
                Case 11 : LifeRate_Table = IIf(Not Joint, 0.287, 0.43)
                Case 12 : LifeRate_Table = IIf(Not Joint, 0.313, 0.469)
                Case 13 : LifeRate_Table = IIf(Not Joint, 0.339, 0.507)
                Case 14 : LifeRate_Table = IIf(Not Joint, 0.364, 0.545)
                Case 15 : LifeRate_Table = IIf(Not Joint, 0.39, 0.584)
                Case 16 : LifeRate_Table = IIf(Not Joint, 0.415, 0.622)
                Case 17 : LifeRate_Table = IIf(Not Joint, 0.44, 0.66)
                Case 18 : LifeRate_Table = IIf(Not Joint, 0.466, 0.697)
                Case 19 : LifeRate_Table = IIf(Not Joint, 0.491, 0.735)
                Case 20 : LifeRate_Table = IIf(Not Joint, 0.516, 0.773)
                Case 21 : LifeRate_Table = IIf(Not Joint, 0.541, 0.81)
                Case 22 : LifeRate_Table = IIf(Not Joint, 0.566, 0.847)
                Case 23 : LifeRate_Table = IIf(Not Joint, 0.591, 0.885)
                Case 24 : LifeRate_Table = IIf(Not Joint, 0.615, 0.922)
                Case 25 : LifeRate_Table = IIf(Not Joint, 0.64, 0.959)
                Case 26 : LifeRate_Table = IIf(Not Joint, 0.665, 0.996)
                Case 27 : LifeRate_Table = IIf(Not Joint, 0.689, 1.033)
                Case 28 : LifeRate_Table = IIf(Not Joint, 0.714, 1.07)
                Case 29 : LifeRate_Table = IIf(Not Joint, 0.738, 1.106)
                Case 30 : LifeRate_Table = IIf(Not Joint, 0.763, 1.143)
                Case 31 : LifeRate_Table = IIf(Not Joint, 0.787, 1.179)
                Case 32 : LifeRate_Table = IIf(Not Joint, 0.811, 1.216)
                Case 33 : LifeRate_Table = IIf(Not Joint, 0.836, 1.252)
                Case 34 : LifeRate_Table = IIf(Not Joint, 0.86, 1.288)
                Case 35 : LifeRate_Table = IIf(Not Joint, 0.884, 1.324)
                Case 36 : LifeRate_Table = IIf(Not Joint, 0.908, 1.36)
                Case 37 : LifeRate_Table = IIf(Not Joint, 0.932, 1.396)
                Case 38 : LifeRate_Table = IIf(Not Joint, 0.956, 1.431)
                Case 39 : LifeRate_Table = IIf(Not Joint, 0.979, 1.467)
                Case 40 : LifeRate_Table = IIf(Not Joint, 1.003, 1.503)
                Case 41 : LifeRate_Table = IIf(Not Joint, 1.027, 1.538)
                Case 42 : LifeRate_Table = IIf(Not Joint, 1.05, 1.573)
                Case 43 : LifeRate_Table = IIf(Not Joint, 1.074, 1.609)
                Case 44 : LifeRate_Table = IIf(Not Joint, 1.097, 1.644)
                Case 45 : LifeRate_Table = IIf(Not Joint, 1.121, 1.679)
                Case 46 : LifeRate_Table = IIf(Not Joint, 1.144, 1.714)
                Case 47 : LifeRate_Table = IIf(Not Joint, 1.167, 1.749)
                Case 48 : LifeRate_Table = IIf(Not Joint, 1.191, 1.784)
                Case 49 : LifeRate_Table = IIf(Not Joint, 1.214, 1.818)
                Case 50 : LifeRate_Table = IIf(Not Joint, 1.237, 1.853)
                Case 51 : LifeRate_Table = IIf(Not Joint, 1.26, 1.887)
                Case 52 : LifeRate_Table = IIf(Not Joint, 1.283, 1.922)
                Case 53 : LifeRate_Table = IIf(Not Joint, 1.306, 1.956)
                Case 54 : LifeRate_Table = IIf(Not Joint, 1.329, 1.99)
                Case 55 : LifeRate_Table = IIf(Not Joint, 1.351, 2.024)
                Case 56 : LifeRate_Table = IIf(Not Joint, 1.374, 2.058)
                Case 57 : LifeRate_Table = IIf(Not Joint, 1.397, 2.092)
                Case 58 : LifeRate_Table = IIf(Not Joint, 1.419, 2.126)
                Case 59 : LifeRate_Table = IIf(Not Joint, 1.442, 2.16)
                Case 60 : LifeRate_Table = IIf(Not Joint, 1.464, 2.194)
            End Select
        End Get
    End Property

    Public ReadOnly Property DisabilityRate(ByVal N As Integer, Optional ByVal R As Integer = 0) As Double
        Get
            'Dim X(1 To 120) As Variant
            Dim X(0 To 119) As Object    'Replaced above line aray with 0 to 119. Because vb.net will not accept 1 as lbound. It must be zero only.
            R = FitRange(0, R, 7)
            'N = FitRange(1, N, 120)
            N = FitRange(0, N, 119)     'Replaced above line because array start must be zero, not 1.

            'X(0) = Array(0.31, 0.47, 0.24, 0.36, 0.21, 0.32, 0.13, 0.2)
            X(0) = New String() {0.31, 0.47, 0.24, 0.36, 0.21, 0.32, 0.13, 0.2}
            X(1) = New String() {0.61, 0.92, 0.47, 0.71, 0.42, 0.63, 0.26, 0.39}
            X(2) = New String() {0.92, 1.38, 0.71, 1.07, 0.63, 0.95, 0.4, 0.6}
            X(3) = New String() {1.23, 1.84, 0.93, 1.4, 0.84, 1.26, 0.53, 0.8}
            X(4) = New String() {1.52, 2.28, 1.16, 1.74, 1.05, 1.58, 0.66, 0.99}
            X(5) = New String() {1.74, 2.61, 1.39, 2.09, 1.26, 1.89, 0.79, 1.18}
            X(6) = New String() {1.84, 2.76, 1.57, 2.35, 1.38, 2.07, 0.9, 1.35}
            X(7) = New String() {1.94, 2.91, 1.66, 2.49, 1.48, 2.22, 0.99, 1.48}
            X(8) = New String() {2.01, 3.02, 1.73, 2.6, 1.58, 2.37, 1.08, 1.62}
            X(9) = New String() {2.1, 3.15, 1.81, 2.71, 1.67, 2.5, 1.15, 1.73}
            X(10) = New String() {2.16, 3.24, 1.88, 2.82, 1.71, 2.57, 1.24, 1.86}
            X(11) = New String() {2.21, 3.32, 1.93, 2.89, 1.78, 2.66, 1.29, 1.94}
            X(12) = New String() {2.27, 3.41, 1.99, 2.99, 1.8, 2.7, 1.35, 2.03}
            X(13) = New String() {2.32, 3.48, 2.05, 3.08, 1.85, 2.77, 1.41, 2.12}
            X(14) = New String() {2.38, 3.57, 2.1, 3.15, 1.88, 2.82, 1.46, 2.19}
            X(15) = New String() {2.43, 3.64, 2.15, 3.22, 1.91, 2.86, 1.51, 2.27}
            X(16) = New String() {2.47, 3.71, 2.19, 3.29, 1.94, 2.91, 1.56, 2.34}
            X(17) = New String() {2.52, 3.78, 2.23, 3.34, 1.98, 2.96, 1.62, 2.43}
            X(18) = New String() {2.56, 3.83, 2.29, 3.43, 1.99, 2.98, 1.66, 2.49}
            X(19) = New String() {2.6, 3.9, 2.31, 3.47, 2.02, 3.03, 1.69, 2.54}
            X(20) = New String() {2.64, 3.95, 2.36, 3.54, 2.06, 3.08, 1.73, 2.59}
            X(21) = New String() {2.67, 4.01, 2.39, 3.59, 2.07, 3.1, 1.75, 2.63}
            X(22) = New String() {2.72, 4.08, 2.43, 3.64, 2.09, 3.13, 1.76, 2.64}
            X(23) = New String() {2.74, 4.11, 2.46, 3.69, 2.11, 3.17, 1.78, 2.68}
            X(24) = New String() {2.78, 4.18, 2.5, 3.74, 2.12, 3.18, 1.81, 2.71}
            X(25) = New String() {2.81, 4.21, 2.53, 3.8, 2.17, 3.25, 1.84, 2.76}
            X(26) = New String() {2.84, 4.26, 2.56, 3.85, 2.18, 3.26, 1.85, 2.78}
            X(27) = New String() {2.86, 4.29, 2.59, 3.88, 2.19, 3.28, 1.87, 2.81}
            X(28) = New String() {2.91, 4.36, 2.62, 3.93, 2.21, 3.31, 1.88, 2.83}
            X(29) = New String() {2.91, 4.37, 2.65, 3.98, 2.23, 3.34, 1.91, 2.86}
            X(30) = New String() {2.96, 4.44, 2.69, 4.03, 2.25, 3.38, 1.93, 2.89}
            X(31) = New String() {2.99, 4.49, 2.7, 4.04, 2.26, 3.39, 1.94, 2.91}
            X(32) = New String() {3, 4.5, 2.73, 4.09, 2.27, 3.4, 1.95, 2.92}
            X(33) = New String() {3.03, 4.55, 2.76, 4.14, 2.3, 3.46, 1.98, 2.97}
            X(34) = New String() {3.06, 4.58, 2.78, 4.17, 2.31, 3.47, 2, 3.01}
            X(35) = New String() {3.08, 4.61, 2.82, 4.22, 2.33, 3.5, 2.03, 3.04}
            X(36) = New String() {3.11, 4.66, 2.84, 4.26, 2.34, 3.52, 2.02, 3.03}
            X(37) = New String() {3.13, 4.69, 2.87, 4.3, 2.35, 3.53, 2.03, 3.05}
            X(38) = New String() {3.16, 4.74, 2.88, 4.32, 2.37, 3.56, 2.05, 3.08}
            X(39) = New String() {3.17, 4.75, 2.9, 4.35, 2.37, 3.56, 2.06, 3.1}
            X(40) = New String() {3.2, 4.8, 2.93, 4.4, 2.39, 3.59, 2.07, 3.11}
            X(41) = New String() {3.23, 4.85, 2.95, 4.43, 2.41, 3.62, 2.09, 3.14}
            X(42) = New String() {3.24, 4.86, 2.97, 4.46, 2.43, 3.65, 2.12, 3.17}
            X(43) = New String() {3.26, 4.89, 2.99, 4.49, 2.43, 3.65, 2.13, 3.19}
            X(44) = New String() {3.29, 4.94, 3.01, 4.52, 2.45, 3.68, 2.13, 3.2}
            X(45) = New String() {3.31, 4.97, 3.03, 4.55, 2.46, 3.69, 2.14, 3.22}
            X(46) = New String() {3.33, 5, 3.05, 4.58, 2.48, 3.72, 2.17, 3.25}
            X(47) = New String() {3.35, 5.03, 3.07, 4.61, 2.49, 3.74, 2.17, 3.26}
            X(48) = New String() {3.36, 5.04, 3.11, 4.66, 2.5, 3.75, 2.2, 3.29}
            X(49) = New String() {3.39, 5.09, 3.13, 4.69, 2.51, 3.76, 2.21, 3.31}
            X(50) = New String() {3.41, 5.12, 3.13, 4.7, 2.52, 3.78, 2.2, 3.3}
            X(51) = New String() {3.42, 5.13, 3.15, 4.73, 2.53, 3.79, 2.21, 3.32}
            X(52) = New String() {3.44, 5.16, 3.19, 4.78, 2.55, 3.82, 2.23, 3.35}
            X(53) = New String() {3.47, 5.21, 3.19, 4.79, 2.54, 3.81, 2.24, 3.36}
            X(54) = New String() {3.48, 5.22, 3.2, 4.8, 2.56, 3.85, 2.25, 3.38}
            X(55) = New String() {3.51, 5.26, 3.23, 4.85, 2.58, 3.88, 2.27, 3.41}
            X(56) = New String() {3.52, 5.28, 3.24, 4.86, 2.59, 3.89, 2.28, 3.42}
            X(57) = New String() {3.52, 5.29, 3.27, 4.91, 2.59, 3.88, 2.29, 3.43}
            X(58) = New String() {3.56, 5.33, 3.29, 4.94, 2.6, 3.9, 2.3, 3.45}
            X(59) = New String() {3.57, 5.36, 3.3, 4.95, 2.61, 3.91, 2.31, 3.46}
            X(60) = New String() {3.58, 5.37, 3.31, 4.96, 2.63, 3.94, 2.32, 3.47}
            X(61) = New String() {3.59, 5.38, 3.33, 4.99, 2.65, 3.97, 2.34, 3.51}
            X(62) = New String() {3.61, 5.41, 3.35, 5.02, 2.66, 3.98, 2.36, 3.54}
            X(63) = New String() {3.63, 5.44, 3.36, 5.05, 2.66, 4, 2.37, 3.55}
            X(64) = New String() {3.63, 5.45, 3.36, 5.04, 2.68, 4.03, 2.37, 3.56}
            X(65) = New String() {3.65, 5.48, 3.38, 5.07, 2.7, 4.06, 2.4, 3.59}
            X(66) = New String() {3.66, 5.49, 3.4, 5.1, 2.71, 4.07, 2.42, 3.62}
            X(67) = New String() {3.68, 5.52, 3.41, 5.11, 2.73, 4.1, 2.42, 3.64}
            X(68) = New String() {3.69, 5.53, 3.43, 5.14, 2.75, 4.13, 2.44, 3.67}
            X(69) = New String() {3.7, 5.56, 3.44, 5.17, 2.76, 4.14, 2.45, 3.68}
            X(70) = New String() {3.72, 5.58, 3.45, 5.18, 2.77, 4.15, 2.47, 3.71}
            X(71) = New String() {3.73, 5.59, 3.46, 5.19, 2.79, 4.18, 2.48, 3.72}
            X(72) = New String() {3.74, 5.6, 3.48, 5.22, 2.81, 4.21, 2.5, 3.75}
            X(73) = New String() {3.75, 5.63, 3.5, 5.24, 2.83, 4.24, 2.52, 3.78}
            X(74) = New String() {3.77, 5.66, 3.52, 5.27, 2.82, 4.24, 2.53, 3.8}
            X(75) = New String() {3.79, 5.69, 3.51, 5.27, 2.84, 4.27, 2.54, 3.81}
            X(76) = New String() {3.8, 5.7, 3.53, 5.29, 2.86, 4.3, 2.56, 3.84}
            X(77) = New String() {3.81, 5.71, 3.55, 5.32, 2.88, 4.32, 2.58, 3.87}
            X(78) = New String() {3.82, 5.74, 3.57, 5.35, 2.89, 4.34, 2.6, 3.9}
            X(79) = New String() {3.84, 5.76, 3.57, 5.36, 2.9, 4.35, 2.61, 3.91}
            X(80) = New String() {3.85, 5.77, 3.59, 5.39, 2.92, 4.38, 2.62, 3.92}
            X(81) = New String() {3.87, 5.8, 3.6, 5.4, 2.93, 4.39, 2.64, 3.95}
            X(82) = New String() {3.87, 5.81, 3.62, 5.43, 2.95, 4.42, 2.66, 3.98}
            X(83) = New String() {3.88, 5.82, 3.62, 5.44, 2.96, 4.45, 2.66, 4}
            X(84) = New String() {3.9, 5.85, 3.64, 5.47, 2.98, 4.48, 2.68, 4.03}
            X(85) = New String() {3.92, 5.87, 3.66, 5.49, 2.98, 4.47, 2.69, 4.04}
            X(86) = New String() {3.93, 5.9, 3.66, 5.49, 3, 4.5, 2.7, 4.05}
            X(87) = New String() {3.94, 5.91, 3.68, 5.51, 3.02, 4.53, 2.72, 4.08}
            X(88) = New String() {3.95, 5.92, 3.69, 5.54, 3.04, 4.56, 2.74, 4.11}
            X(89) = New String() {3.96, 5.95, 3.71, 5.57, 3.05, 4.57, 2.76, 4.14}
            X(90) = New String() {3.98, 5.97, 3.72, 5.58, 3.05, 4.58, 2.77, 4.15}
            X(91) = New String() {3.99, 5.98, 3.74, 5.6, 3.07, 4.61, 2.77, 4.16}
            X(92) = New String() {4.01, 6.01, 3.74, 5.61, 3.09, 4.64, 2.79, 4.19}
            X(93) = New String() {4.01, 6.02, 3.76, 5.64, 3.1, 4.65, 2.81, 4.22}
            X(94) = New String() {4.03, 6.05, 3.77, 5.65, 3.12, 4.68, 2.82, 4.23}
            X(95) = New String() {4.04, 6.06, 3.79, 5.68, 3.14, 4.7, 2.84, 4.26}
            X(96) = New String() {4.05, 6.08, 3.8, 5.71, 3.14, 4.72, 2.85, 4.27}
            X(97) = New String() {4.07, 6.11, 3.81, 5.72, 3.15, 4.73, 2.87, 4.3}
            X(98) = New String() {4.08, 6.12, 3.82, 5.73, 3.17, 4.75, 2.87, 4.31}
            X(99) = New String() {4.08, 6.13, 3.83, 5.75, 3.19, 4.78, 2.89, 4.34}
            X(100) = New String() {4.1, 6.15, 3.85, 5.78, 3.21, 4.81, 2.91, 4.37}
            X(101) = New String() {4.12, 6.18, 3.87, 5.81, 3.2, 4.8, 2.92, 4.38}
            X(102) = New String() {4.14, 6.2, 3.88, 5.82, 3.22, 4.83, 2.93, 4.39}
            X(103) = New String() {4.14, 6.21, 3.88, 5.82, 3.24, 4.86, 2.95, 4.42}
            X(104) = New String() {4.15, 6.22, 3.9, 5.85, 3.25, 4.87, 2.96, 4.45}
            X(105) = New String() {4.17, 6.25, 3.92, 5.88, 3.27, 4.9, 2.98, 4.48}
            X(106) = New String() {4.17, 6.26, 3.92, 5.89, 3.28, 4.93, 2.99, 4.49}
            X(107) = New String() {4.19, 6.28, 3.94, 5.91, 3.29, 4.94, 3, 4.5}
            X(108) = New String() {4.21, 6.31, 3.95, 5.92, 3.3, 4.95, 3.02, 4.53}
            X(109) = New String() {4.21, 6.32, 3.95, 5.93, 3.32, 4.97, 3.02, 4.54}
            X(110) = New String() {4.22, 6.33, 3.97, 5.96, 3.33, 5, 3.04, 4.57}
            X(111) = New String() {4.24, 6.35, 3.99, 5.98, 3.35, 5.03, 3.06, 4.59}
            X(112) = New String() {4.25, 6.38, 4.01, 6.01, 3.35, 5.02, 3.07, 4.6}
            X(113) = New String() {4.27, 6.4, 4, 6, 3.37, 5.05, 3.08, 4.61}
            X(114) = New String() {4.28, 6.41, 4.02, 6.03, 3.38, 5.08, 3.09, 4.64}
            X(115) = New String() {4.28, 6.42, 4.04, 6.05, 3.4, 5.1, 3.11, 4.67}
            X(116) = New String() {4.3, 6.45, 4.05, 6.08, 3.41, 5.11, 3.13, 4.7}
            X(117) = New String() {4.31, 6.47, 4.06, 6.09, 3.43, 5.14, 3.14, 4.71}
            X(118) = New String() {4.32, 6.48, 4.08, 6.11, 3.43, 5.15, 3.15, 4.72}
            X(119) = New String() {4.34, 6.51, 4.08, 6.12, 3.45, 5.18, 3.16, 4.75}

            DisabilityRate = X(N)(R)

        End Get
    End Property

    Public ReadOnly Property AccidentPremium() As Decimal
        Get
            '  AccidentPremium =Round((spdd * N / 100) * (CA / (an / odf - ((Sp / 1200 * (N + (OD - 30) / 30) * N) + (iusp * Nu / 100) + (spdd * N / 100) + (spp * N / 100)))), 2)
            If UserAcciPremium <> 0 Then AccidentPremium = UserAcciPremium : Exit Property
            AccidentPremium = Math.Round((SPDD * N / 100) * (CA / (AN / ODF - ((Sp / 1200 * (N + (OD - 30) / 30) * N) + (IUSP * Nu / 100) + (SPDD * N / 100) + (SPP * N / 1200 * ((N + (OD - 30) / 30)))))), 2)
        End Get
    End Property

    Public ReadOnly Property PropertyPremium() As Decimal
        Get
            Dim K As Double
            If UserPropPremium <> 0 Then PropertyPremium = UserPropPremium : Exit Property
            'PropertyPremium = (spp * N / 100) * (CA / (an / odf - ((Sp / 1200 * (N + (OD - 30) / 30) * N) + (iusp * Nu / 100) + (spdd * N / 100) + (spp * N / 100))))
            K = (N + (OD - 30) / 30)
            '  propertypremium =
            PropertyPremium = SPP * N / 1200 * (K * (CA / (AN / ODF - ((Sp / 1200 * K * N) + (IUSP * Nu / 100) + (SPDD * N / 100) + (SPP * N / 1200 * (K))))))
            PropertyPremium = Math.Round(PropertyPremium, 2)
        End Get
    End Property

    Public ReadOnly Property IUIPremium() As Decimal
        Get
            If UserIUIPremium <> 0 Then IUIPremium = UserIUIPremium : Exit Property
            'IUIPremium =      (iusp * Nu / 100) * (CA / (an / odf - ((Sp / 1200 * (N + (OD - 30) / 30) * N) + (iusp * Nu / 100) + (spdd * N / 100) + (spp * N / 100))))
            IUIPremium = Math.Round((IUSP * Nu / 100) * (CA / (AN / ODF - ((Sp / 1200 * (N + (OD - 30) / 30) * N) + (IUSP * Nu / 100) + (SPDD * N / 100) + (SPP * N / 1200 * ((N + (OD - 30) / 30)))))), 2)
        End Get
    End Property

    Public ReadOnly Property AmountFinanced() As Decimal
        Get
            AmountFinanced = CA + LifePremium + AccidentPremium + PropertyPremium + IUIPremium
        End Get
    End Property

    Public ReadOnly Property FC() As Decimal
        Get
            FC = B - AmountFinanced
        End Get
    End Property

    Public ReadOnly Property B() As Decimal
        Get
            B = MonthlyPayment * N
        End Get
    End Property

    Public ReadOnly Property MonthlyPayment() As Decimal
        Get
            MonthlyPayment = (CA + LifePremium + AccidentPremium + PropertyPremium + IUIPremium) / (AN / ODF)
        End Get
    End Property

End Class
