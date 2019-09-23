Imports Microsoft.VisualBasic.Interaction
Public Class ARPaySetUp
    Dim WithEvents mDBAccess As CDbAccessGeneral
    Dim WithEvents mDBAccessTransactions As CDbAccessGeneral

    Dim PrevousBal As Decimal
    Dim Contract As Decimal
    Dim DocFee As Decimal
    Dim SendNotice As String
    Public Rate As Double
    Dim InterestRate As Double
    Dim SIR As Double
    Dim NewBalance As Decimal
    Dim FinanceCharge As Decimal
    Dim FinanceChargeSalesTax As Decimal
    Dim Months As Integer
    Dim DeferredMonths As Integer
    Dim DeferredInt As Decimal
    Dim CashOpt As Integer
    Dim Payment As Decimal
    Dim APR As Double
    Dim LastPay As Decimal
    Public NoMonths As Object
    Dim DueOn As Integer
    Public ArNo As String
    Dim Status As String
    Dim mArNo As String
    Dim TransType As String
    Dim Charges As Decimal
    Dim Credits As Decimal
    Dim Balance As Decimal
    Dim LifeCredit As Decimal
    Dim AccidentCredit As Decimal
    Dim PropertyCredit As Decimal
    Dim IUICredit As Decimal
    Dim InterestCredit As Decimal
    Public FirstPayment As String
    Dim TotPaid As Decimal
    Dim LateChargeBal As Decimal
    Dim LateCharge As Decimal
    Dim Copies As Integer
    Dim Counter As Integer
    Dim Telephone As String
    Dim DateDue As String
    Dim mBalance As String
    Dim AccountArray(12) As Object
    Dim Zz As Integer                        ' Used in mDBAccess_GetRecordEvent.
    Dim AddOn As String
    Dim AddOnRecordAccount As String

    Public MailRec As Integer                   ' Used by ArCard
    Public AccountFound As String            ' Used by MainMenu
    Public UnloadARPaySetUp As String        ' Used by OrdSelect
    Public INTEREST As Decimal              ' Called by ArCard.
    Public InterestTax As Decimal
    Public Life As Decimal                  ' Called by ArCard.
    Public Accident As Decimal              ' Called by ArCard.
    Public Prop As Decimal                  ' Called by ArCard.
    Public IUI As Decimal
    Dim DBInterest As Decimal

    Private ReprintMailIndex As Integer, ReprintSaleNo As String
    Private NoAdjust As Boolean

    Public NeedPayOff As Boolean             ' Because we moved payoff until the OK button

    Private sBOS2_ProcessSale As Boolean, sBOS2_NextSale As Boolean, sBOS2_Clear As Boolean, sBOS2_MainMenu As Boolean, sBOS2_Notes_Open As Boolean
    Private LastCheckedArNo As String

    Private ReadOnly Property IStorename() As String
        Get
            'IStorename = Switch(IsLouisOfHartford, "C & C Sales Finance Plan", True, StoreSettings.Name)
            IStorename = StoreSettings.Name
        End Get
    End Property

    Private Sub cboCashOption_Click(sender As Object, e As EventArgs) Handles cboCashOption.Click
        OneOrTheOther(False)
    End Sub

    Private Sub OneOrTheOther(ByVal Deferred As Boolean)
        Static IsInIt As Boolean
        If IsInIt Then Exit Sub
        IsInIt = True
        If Deferred Then
            'cboCashOption.ListIndex = 0
            cboCashOption.SelectedIndex = 0

        Else
            'cboDeferred.ListIndex = 0
            cboDeferred.SelectedIndex = 0
        End If
        IsInIt = False
    End Sub

    Private Sub chkRoundUp_Click(sender As Object, e As EventArgs) Handles chkRoundUp.Click
        Recalculate()
    End Sub

    Private ReadOnly Property IStoreAddress() As String
        Get
            'IStoreAddress = Switch(IsLouisOfHartford, "P.O. Box 260010", True, StoreSettings.Address)
            IStoreAddress = StoreSettings.Address
        End Get
    End Property

    Private ReadOnly Property IStoreCity() As String
        Get
            'IStoreCity = Switch(IsLouisOfHartford, "Hartford, CT 06126", True, StoreSettings.City)
            IStoreCity = StoreSettings.City
        End Get
    End Property

    Private ReadOnly Property IStorePhone() As String
        Get
            'IStorePhone = Switch(IsLouisOfHartford, "(860) 247-9806", True, StoreSettings.Phone)
            IStorePhone = StoreSettings.Phone
        End Get
    End Property

    Private ReadOnly Property UseThorntonsInsurance() As Boolean
        Get
            UseThorntonsInsurance = IsThorntons
        End Get
    End Property

    Private Sub HandleBillOSaleControls(ByVal revert As Boolean)
        cmdApply.Enabled = revert
        cmdPrint.Enabled = revert
        If Not revert Then
            sBOS2_ProcessSale = BillOSale.cmdProcessSale.Enabled
            sBOS2_NextSale = BillOSale.cmdNextSale.Enabled
            sBOS2_Clear = BillOSale.cmdClear.Enabled
            sBOS2_MainMenu = BillOSale.cmdMainMenu.Enabled
            sBOS2_Notes_Open = BillOSale.Notes_Open.Enabled

            BillOSale.cmdProcessSale.Enabled = False
            BillOSale.cmdNextSale.Enabled = False
            BillOSale.cmdClear.Enabled = False
            BillOSale.cmdMainMenu.Enabled = False
            BillOSale.Notes_Open.Enabled = False
        Else
            BillOSale.cmdProcessSale.Enabled = sBOS2_ProcessSale
            BillOSale.cmdNextSale.Enabled = sBOS2_NextSale
            BillOSale.cmdClear.Enabled = sBOS2_Clear
            BillOSale.cmdMainMenu.Enabled = sBOS2_MainMenu
            BillOSale.Notes_Open.Enabled = sBOS2_Notes_Open
        End If
        Application.DoEvents()
    End Sub

    Public Sub Recalculate()
        Dim XI As Integer, TI As Decimal, X As Decimal
        Dim A As CArLott
        Dim MaxAPR As Double

        Life = 0
        Accident = 0
        Prop = 0
        IUI = 0

        txtSubTotal.Text = CurrencyFormat(GetPrice(txtPrevBalance.Text) + GetPrice(txtGrossSale.Text) - GetPrice(txtOrigDeposit.Text))

        If IsElmore Then
            If GetPrice(txtSubTotal.Text) <= 1500.0# Then
                Rate = 24 : APR = Rate
            ElseIf txtSubTotal.Text > 1501 And txtSubTotal.Text <= 2000 Then
                Rate = 22 : APR = Rate
            ElseIf txtSubTotal.Text > 2001 And txtSubTotal.Text <= 3000 Then
                Rate = 20 : APR = Rate
            ElseIf txtSubTotal.Text > 3000 Then
                Rate = 18 : APR = Rate
            End If
        ElseIf IsKenLu Then
            'BFH20170615 - Added KenLu Clause (just a little different than Elmore)
            'New customer with Installment Financing.  Their states require a graduated APR interest rate.
            '          Less than 1488.89 = 24% APR
            '          Less than 1999.99 = 22%
            '          Less than 2499.99 = 20%
            '          Greater than 2500.00 = 18%
            '
            'Comment that this is for North Carolina
            'Make sure interest rate shows up on the new account set up form.    If GetPrice(txtFinanceAmount) <= 1500# Then
            If False Then
            ElseIf GetPrice(txtSubTotal.Text) <= 1488.89 Then
                Rate = 24 : APR = Rate
            ElseIf GetPrice(txtSubTotal.Text) < 1999.99 Then
                Rate = 22 : APR = Rate
            ElseIf GetPrice(txtSubTotal.Text) < 2499.99 Then
                Rate = 20 : APR = Rate
            Else
                Rate = 18 : APR = Rate
            End If
            If Val(txtRate) <> 0 Then
                Rate = Val(txtRate) : APR = Rate
            End If
        End If

        DocFee = GetPrice(txtDocFee.Text)
        txtDocFee.Text = CurrencyFormat(DocFee)
        Months = Val(txtMonthsToFinance.Text)


        ' transplanted from optJointLife_onClick
        If IsRevolvingCharge(txtArNo.Text) Then
            FinanceCharge = INTEREST
            FinanceChargeSalesTax = 0
            'Payment = CalculateRevolvingPayment(RevolvingCurrentFinancedAmount(txtArNo) + GetPrice(txtTotalBalance) - GetPrice(txtPrevBalance), chkRoundUp.Value)
            txtFinanceAmount.Text = CurrencyFormat(txtSubTotal.Text + GetPrice(txtDocFee.Text) + GetPrice(txtLifeInsurance.Text) + GetPrice(txtAccidentInsurance.Text) + GetPrice(txtPropertyInsurance.Text) + GetPrice(txtUnemploymentInsurance.Text))
            txtTotalBalance.Text = CurrencyFormat(GetPrice(txtFinanceAmount.Text) + GetPrice(txtFinanceCharges.Text))
            Payment = CalculateRevolvingPayment(GetPrice(txtTotalBalance.Text), chkRoundUp.Checked, CLng(Months))
            APR = StoreSettings.ModifiedRevolvingRate
            If APR Then
                InterestRate = Rate * 0.01 / 12 'CalculateSIR(NewBalance, Rate, Months)
            Else
                InterestRate = Rate * 0.01
            End If
        Else
            txtFinanceAmount.Text = CurrencyFormat(GetPrice(txtSubTotal.Text) + GetPrice(txtDocFee.Text) + GetPrice(txtLifeInsurance.Text) + GetPrice(txtAccidentInsurance.Text) + GetPrice(txtPropertyInsurance.Text)) + GetPrice(txtUnemploymentInsurance.Text)
            If optWeekly.Checked = True Then
                Payment = (GetPrice(txtFinanceAmount.Text) + FinanceCharge) / (Max(Months, 1) * 4)
            Else
                Payment = (GetPrice(txtFinanceAmount.Text) + FinanceCharge) / Max(Months, 1)
            End If
            If StoreSettings.bAPR Then
                InterestRate = CalculateSIR(GetPrice(txtFinanceAmount.Text), Rate, Months)
            Else
                InterestRate = Rate * 0.01
            End If
            txtFinanceAmount.Text = CurrencyFormat(GetPrice(txtSubTotal.Text) + GetPrice(txtDocFee.Text) + GetPrice(txtLifeInsurance.Text) + GetPrice(txtAccidentInsurance.Text) + GetPrice(txtPropertyInsurance.Text)) + GetPrice(txtUnemploymentInsurance.Text)
            FinanceCharge = ((txtFinanceAmount.Text * InterestRate) / 12 * Months)
        End If
        SIR = InterestRate
        txtFinanceCharges.Text = CurrencyFormat(FinanceCharge)
        txtPaymentWillBe.Text = CurrencyFormat(Payment)


        If IsLott Then ' Or IsLott   ' Or IsMidSouth         ' Lott is now completely reverse driven through the new class module
            A = New CArLott

            If StoreSettings.bAPR Then
                InterestRate = CalculateSIR(NewBalance, Rate, Months)
            Else
                InterestRate = Rate * 0.01
            End If
            SIR = InterestRate
            'A.JointLife = optJointLife(1)
            A.JointLife = optJointLife1.Checked
            A.JointLifeReducing = False

            A.PrevBalance = GetPrice(txtPrevBalance.Text)
            A.NewSale = GetPrice(txtGrossSale.Text)
            A.Deposit = GetPrice(txtOrigDeposit.Text)
            A.DocFee = GetPrice(txtDocFee.Text)

            'A.LifeInsuranceOn = Switch(chkLife = vbChecked, vbTrue, chkLife = vbUnchecked, vbFalse, chkLife = vbGrayed, vbUseDefault)
            A.LifeInsuranceOn = Switch(chkLife.CheckState = CheckState.Checked, vbTrue, chkLife.CheckState = CheckState.Unchecked, vbFalse, chkLife.CheckState = CheckState.Indeterminate, vbUseDefault)
            'A.UserLifePremium = IIf(chkLife = vbGrayed, GetPrice(txtLifeInsurance), 0)
            A.UserLifePremium = IIf(chkLife.CheckState = CheckState.Indeterminate, GetPrice(txtLifeInsurance.Text), 0)
            'A.AccidentInsuranceOn = Switch(chkAccident = vbChecked, vbTrue, chkAccident = vbUnchecked, vbFalse, chkAccident = vbGrayed, vbUseDefault)
            A.AccidentInsuranceOn = Switch(chkAccident.CheckState = CheckState.Checked, vbTrue, chkAccident.CheckState = CheckState.Unchecked, vbFalse, chkAccident.CheckState = CheckState.Indeterminate, vbUseDefault)
            'A.UserAccidentPremium = IIf(chkAccident = vbGrayed, GetPrice(txtAccidentInsurance), 0)
            A.UserAccidentPremium = IIf(chkAccident.CheckState = CheckState.Indeterminate, GetPrice(txtAccidentInsurance.Text), 0)
            'A.PropertyInsuranceOn = Switch(chkProperty = vbChecked, vbTrue, chkProperty = vbUnchecked, vbFalse, chkProperty = vbGrayed, vbUseDefault)
            A.PropertyInsuranceOn = Switch(chkProperty.CheckState = CheckState.Checked, vbTrue, chkProperty.CheckState = CheckState.Unchecked, vbFalse, chkProperty.CheckState = CheckState.Indeterminate, vbUseDefault)
            'A.UserPropertyPremium = IIf(chkProperty = vbGrayed, GetPrice(txtPropertyInsurance), 0)
            A.UserPropertyPremium = IIf(chkProperty.CheckState = CheckState.Indeterminate, GetPrice(txtPropertyInsurance.Text), 0)

            A.DeferredMonths = Val(cboDeferred.Text)
            A.SimpleInterestRate = SIR
            A.Months = Val(txtMonthsToFinance.Text)
            A.ApplyFinanceChargeSalesTax = (StoreSettings.bInstallmentInterestIsTaxable)
            A.WeeklyPayments = Not (StoreSettings.bPaymentBooksMonthly)
            'A.RoundUp = (chkRoundUp.Value = 1)
            A.RoundUp = (chkRoundUp.Checked = True)

            A.Calculate

            APR = A.APR
            UpdateAPRLabel()
            txtLifeInsurance.Text = CurrencyFormat(A.LifeInsurancePremium)
            txtAccidentInsurance.Text = CurrencyFormat(A.AccidentInsurancePremium)
            txtPropertyInsurance.Text = CurrencyFormat(A.PropertyInsurancePremium)

            NewBalance = A.FinanceSubTotal
            txtFinanceAmount.Text = CurrencyFormat(A.FinanceSubTotal)

            DeferredInt = A.DeferredInterest
            txtDeferredInt.Text = CurrencyFormat(A.DeferredInterest)

            FinanceCharge = A.FinanceCharges
            txtFinanceCharges.Text = CurrencyFormat(A.FinanceCharges)

            FinanceChargeSalesTax = A.FinanceChargeSalesTax
            txtFinanceChargeSalesTax.Text = CurrencyFormat(A.FinanceChargeSalesTax)

            Payment = A.MonthlyPayment
            txtPaymentWillBe.Text = CurrencyFormat(A.MonthlyPayment)

            CalculateLateCharge()
            txtTotalBalance.Text = GetPrice(txtFinanceAmount.Text) + GetPrice(txtFinanceCharges.Text)
            CalculateMath
            Exit Sub

            '  ElseIf ...
        ElseIf UseThorntonsInsurance Then             ' Thorntons, like Lott, is now completely reverse driven through the new class module.. but of course, is slightly different
            A = New CArLott

            If StoreSettings.bAPR Then
                InterestRate = CalculateSIR(NewBalance, Rate, Months)
            Else
                InterestRate = Rate * 0.01
            End If
            SIR = InterestRate
            'A.JointLife = optJointLife(1)
            A.JointLife = optJointLife1.Checked
            A.JointLifeReducing = False

            If True Or (GetPrice(txtGrossSale.Text) >= 0 And Not IsDevelopment()) Then
                A.PrevBalance = GetPrice(txtPrevBalance.Text)
                A.NewSale = GetPrice(txtGrossSale.Text)
                A.Deposit = GetPrice(txtOrigDeposit.Text)
                A.DocFee = GetPrice(txtDocFee.Text)

                'A.LifeInsuranceOn = Switch(chkLife = vbChecked, vbTrue, chkLife = vbUnchecked, vbFalse, chkLife = vbGrayed, vbUseDefault)
                A.LifeInsuranceOn = Switch(chkLife.CheckState = CheckState.Checked, vbTrue, chkLife.CheckState = CheckState.Unchecked, vbFalse, chkLife.CheckState = CheckState.Indeterminate, vbUseDefault)
                'A.UserLifePremium = IIf(chkLife = vbGrayed, GetPrice(txtLifeInsurance), 0)
                A.UserLifePremium = IIf(chkLife.CheckState = CheckState.Indeterminate, GetPrice(txtLifeInsurance.Text), 0)
                'A.AccidentInsuranceOn = Switch(chkAccident = vbChecked, vbTrue, chkAccident = vbUnchecked, vbFalse, chkAccident = vbGrayed, vbUseDefault)
                A.AccidentInsuranceOn = Switch(chkAccident.CheckState = CheckState.Checked, vbTrue, chkAccident.CheckState = CheckState.Unchecked, vbFalse, chkAccident.CheckState = CheckState.Indeterminate, vbUseDefault)
                'A.UserAccidentPremium = IIf(chkAccident = vbGrayed, GetPrice(txtAccidentInsurance), 0)
                A.UserAccidentPremium = IIf(chkAccident.CheckState = CheckState.Indeterminate, GetPrice(txtAccidentInsurance.Text), 0)
                A.PropertyInsuranceOn = vbUseDefault 'Switch(chkProperty = vbChecked, vbTrue, chkProperty = vbUnchecked, vbFalse, chkProperty = vbGrayed, vbUseDefault)
                'BFH20081113 If we don't have the chkproperty<>gray check, we can't have user-entered amounts
                If chkProperty.CheckState <> CheckState.Indeterminate And chkProperty.Checked = True Then
                    'txtPropertyInsurance = CurrencyFormat(AmericanHeritage_Prop(A.NewSale + A.PrevBalance, Val(txtMonthsToFinance)))
                    txtPropertyInsurance.Text = CurrencyFormat(ThorntonsPropertyRate() * (A.NewSale + A.PrevBalance))  'AmericanHeritage_Prop(A.NewSale + A.PrevBalance, Val(txtMonthsToFinance)))
                End If
                'A.UserPropertyPremium = IIf(chkProperty <> vbUnchecked, GetPrice(txtPropertyInsurance), 0)
                A.UserPropertyPremium = IIf(chkProperty.CheckState <> CheckState.Unchecked, GetPrice(txtPropertyInsurance.Text), 0)

                A.DeferredMonths = Val(cboDeferred.Text)
                A.SimpleInterestRate = SIR
                A.Months = Val(txtMonthsToFinance.Text)
                A.ApplyFinanceChargeSalesTax = StoreSettings.bInstallmentInterestIsTaxable
                A.WeeklyPayments = Not (StoreSettings.bPaymentBooksMonthly)
                'A.RoundUp = (chkRoundUp.Value = 1)
                A.RoundUp = (chkRoundUp.Checked = True)

                A.Calculate()

                APR = A.APR
                UpdateAPRLabel()

                txtLifeInsurance.Text = CurrencyFormat(A.LifeInsurancePremium)
                txtAccidentInsurance.Text = CurrencyFormat(A.AccidentInsurancePremium)
                '    txtPropertyInsurance = CurrencyFormat(A.PropertyInsurancePremium)

                NewBalance = A.FinanceSubTotal
                txtFinanceAmount.Text = CurrencyFormat(A.FinanceSubTotal)

                DeferredInt = A.DeferredInterest
                txtDeferredInt.Text = CurrencyFormat(A.DeferredInterest)

                FinanceCharge = A.FinanceCharges
                txtFinanceCharges.Text = CurrencyFormat(A.FinanceCharges)

                FinanceChargeSalesTax = A.FinanceChargeSalesTax
                txtFinanceChargeSalesTax.Text = CurrencyFormat(A.FinanceChargeSalesTax)

                Payment = A.MonthlyPayment
                txtPaymentWillBe.Text = CurrencyFormat(A.MonthlyPayment)

                CalculateLateCharge()
                txtTotalBalance.Text = CurrencyFormat(GetPrice(txtFinanceAmount.Text) + GetPrice(txtFinanceCharges.Text))
                CalculateMath()
                Exit Sub
            End If

            ' negative, reverse sale
            Dim P As Decimal, S As Decimal, D As Decimal, Doc As Decimal
            Dim tL As Decimal, TA As Decimal, Tp As Decimal, tU As Decimal, tInt As Decimal, tST As Decimal
            GetPreviousContractTerms(txtArNo.Text, 0, P, S, D, Doc, tL, TA, Tp, tU, tInt, tST)

            A.PrevBalance = GetPrice(txtPrevBalance.Text)
            A.NewSale = GetPrice(txtGrossSale.Text)
            A.Deposit = GetPrice(txtOrigDeposit.Text)
            A.DocFee = GetPrice(txtDocFee.Text)

            A.LifeInsuranceOn = vbTrue
            A.UserLifePremium = 0
            A.AccidentInsuranceOn = vbTrue
            A.UserAccidentPremium = 0
            A.PropertyInsuranceOn = vbUseDefault

            'BFH20081113 If we don't have the chkproperty<>gray check, we can't have user-entered amounts
            If chkProperty.CheckState <> CheckState.Indeterminate And chkProperty.Checked = True Then
                txtPropertyInsurance.Text = CurrencyFormat(AmericanHeritage_Prop(A.NewSale + A.PrevBalance, Val(txtMonthsToFinance.Text)))
            End If
            A.UserPropertyPremium = IIf(chkProperty.CheckState <> CheckState.Indeterminate, GetPrice(txtPropertyInsurance.Text), 0)

            A.DeferredMonths = Val(cboDeferred.Text)
            A.SimpleInterestRate = SIR
            A.Months = Val(txtMonthsToFinance.Text)
            A.ApplyFinanceChargeSalesTax = StoreSettings.bInstallmentInterestIsTaxable
            A.WeeklyPayments = Not (StoreSettings.bPaymentBooksMonthly)
            A.RoundUp = (chkRoundUp.Checked = True)

            A.Calculate

            '    txtDocFee = CurrencyFormat(-Doc)
            txtPropertyInsurance.Text = CurrencyFormat(AmericanHeritage_Prop(A.NewSale + A.PrevBalance, Val(txtMonthsToFinance.Text)) - Tp)
            txtAccidentInsurance.Text = CurrencyFormat(A.AccidentInsurancePremium - TA)
            txtLifeInsurance.Text = CurrencyFormat(A.LifeInsurancePremium - tL)
            txtUnemploymentInsurance.Text = CurrencyFormat(0)

            txtFinanceCharges.Text = CurrencyFormat(-tInt)
            txtFinanceAmount.Text = P + tInt
            txtTotalBalance.Text = GetPrice(txtFinanceAmount.Text) + GetPrice(txtFinanceCharges.Text)
            txtPaymentWillBe.Text = CurrencyFormat(GetPrice(txtTotalBalance.Text) / Months)

            NewBalance = GetPrice(txtPrevBalance.Text) + GetPrice(txtGrossSale.Text) + GetPrice(txtOrigDeposit.Text) + GetPrice(txtDocFee.Text) + GetPrice(txtPropertyInsurance.Text) + GetPrice(txtAccidentInsurance.Text) + GetPrice(txtLifeInsurance.Text)
            DeferredInt = 0
            FinanceCharge = -tInt
            FinanceChargeSalesTax = -tST
            Payment = GetPrice(txtPaymentWillBe.Text)

            CalculateLateCharge()
            CalculateMath
            APR = 0
            UpdateAPRLabel
            Exit Sub
        ElseIf UseThorntonsInsurance Then
            txtPropertyInsurance.Text = CurrencyFormat(ThorntonsPropertyRate() * GetPrice(txtSubTotal.Text))
            txtLifeInsurance.Text = CurrencyFormat(ThorntonsLifeRate(optJointLife1.Checked) * GetPrice(txtSubTotal.Text))
            txtAccidentInsurance.Text = CurrencyFormat(ThorntonsAccidentRate * GetPrice(txtSubTotal.Text))

        ElseIf IsBoyd Then
            Dim B As cArBoyd
            B = New cArBoyd
            If StoreSettings.bAPR Then
                InterestRate = Rate 'CalculateSIR(NewBalance, Rate, Months)
            Else
                InterestRate = Rate '* 0.01
            End If
            B.Cash = GetPrice(txtPrevBalance.Text) + GetPrice(txtGrossSale.Text) - GetPrice(txtOrigDeposit.Text) + GetPrice(txtDocFee.Text)

            B.DFP = DateDiff("d", dteDate1, dteDate2) 'Val(cboDeferred.Text) * 30
            B.N = Val(txtMonthsToFinance.Text)
            B.Intr = InterestRate / 100

            B.bLife = chkLife.Checked = True
            B.bAH = chkAccident.Checked = True
            B.bProperty = chkProperty.Checked = True

            B.LifeR = IIf(chkLife.Checked = True, 0.55, 0)
            B.DSr = IIf(chkAccident.Checked = True, 0, 0)         ' rate ????
            B.IUIr = 0                                 ' unemployment??
            B.PRr = IIf(chkProperty.Checked = True, 2.9, 0)       ' BFH20140505 - To 2.9 from 3#

            txtLifeInsurance.Text = CurrencyFormat(B.LifePremium)
            txtAccidentInsurance.Text = CurrencyFormat(B.AHPremium)
            txtPropertyInsurance.Text = CurrencyFormat(B.PropertyPremium)

            txtFinanceAmount.Text = CurrencyFormat(B.AmountFinanced)
            txtFinanceCharges.Text = CurrencyFormat(B.FinanceCharge)
            txtTotalBalance.Text = CurrencyFormat(B.TotalOfPayments)
            txtPaymentWillBe.Text = CurrencyFormat(B.MonthlyPayment)
            Payment = B.MonthlyPayment
            LastPay = B.AmountFinanced - (B.N - 1) * B.MonthlyPayment

            APR = CalculateAPR(B.AmountFinanced, B.FinanceCharge, Val(txtMonthsToFinance), Val(cboDeferred.Text))
            If APR < 0 Then APR = 0
            Dim Sanity As Integer
            If StoreSettings.bAPR And APR > 0 Then
                Dim Delta As Double, Previous As Integer
                Delta = 0.1 * IIf(APR < (Rate), 1, -1)
                Previous = IIf(APR < (Rate), 1, -1)
                Do While True
                    Sanity = Sanity + 1
                    If Sanity > 3000 Then Exit Do
                    If APR > (Rate) And Previous = 1 Or APR < (Rate) And Previous = -1 Then
                        Delta = Delta / -10
                        Previous = Previous * -1
                    End If
                    If APR = Rate Then Exit Do
                    B.Intr = B.Intr + Delta

                    APR = CalculateAPR(B.AmountFinanced, B.FinanceCharge, Val(txtMonthsToFinance), Val(cboDeferred.Text))
                    SIR = B.intr
                    If Math.Abs(Delta) < 0.00000000001 And Previous = 1 Then Exit Do
                Loop
                txtLifeInsurance.Text = CurrencyFormat(B.LifePremium)
                txtAccidentInsurance.Text = CurrencyFormat(B.AHPremium)
                txtPropertyInsurance.Text = CurrencyFormat(B.PropertyPremium)

                txtFinanceAmount.Text = CurrencyFormat(B.AmountFinanced)
                txtFinanceCharges.Text = CurrencyFormat(B.FinanceCharge)
                txtTotalBalance.Text = CurrencyFormat(B.TotalOfPayments)
                txtPaymentWillBe.Text = CurrencyFormat(B.MonthlyPayment)
                Payment = B.MonthlyPayment
                LastPay = B.AmountFinanced - (B.N - 1) * B.MonthlyPayment
            End If

            '      APR = CalculateAPR(B.Cash, B.FinanceCharge, Val(txtMonthsToFinance), Val(cboDeferred.Text))
            UpdateAPRLabel
            NewBalance = B.AmountFinanced
            FinanceCharge = B.FinanceCharge

            B = Nothing
            CalculateLateCharge()
            txtTotalBalance.Text = GetPrice(txtFinanceAmount.Text) + GetPrice(txtFinanceCharges.Text)
            CalculateMath()
            Exit Sub
        ElseIf UseAmericanNationalInsurance Then
            Dim C As cArTreehouse
            C = New cArTreehouse
            InterestRate = Rate '* 0.01
            C.CA = GetPrice(txtPrevBalance.Text) + GetPrice(txtGrossSale.Text) - GetPrice(txtOrigDeposit.Text) + GetPrice(txtDocFee.Text)
            C.Ani = Rate * 0.01
            C.N = Val(txtMonthsToFinance.Text)
            NoMonths = C.N
            C.Nu = C.N
            C.OD = DateDiff("d", dteDate1, dteDate2)    ' days to first pay
            'C.JointLife = optJointLife(1)
            C.JointLife = optJointLife1.Checked
            C.bHasAcci = chkAccident.Checked = True
            C.bHasLife = chkLife.Checked = True
            C.bHasProp = chkProperty.Checked = True
            C.bHasIUI = chkUnemployment.Checked = True

            'BFH20170213 - Added these so that the C.AmountFinanced reflects these for Old Account Setup
            'C.UserLifePremium = IIf(chkLife = vbGrayed, GetPrice(txtLifeInsurance), 0)
            C.UserLifePremium = IIf(chkLife.CheckState = CheckState.Indeterminate, GetPrice(txtLifeInsurance.Text), 0)
            'C.UserAcciPremium = IIf(chkAccident = vbGrayed, GetPrice(txtAccidentInsurance), 0)
            C.UserAcciPremium = IIf(chkAccident.CheckState = CheckState.Indeterminate, GetPrice(txtAccidentInsurance.Text), 0)
            'C.UserPropPremium = IIf(chkProperty = vbGrayed, GetPrice(txtPropertyInsurance), 0)
            C.UserPropPremium = IIf(chkProperty.CheckState = CheckState.Indeterminate, GetPrice(txtPropertyInsurance.Text), 0)
            'C.UserIUIPremium = IIf(chkUnemployment = vbGrayed, GetPrice(txtUnemploymentInsurance), 0)
            C.UserIUIPremium = IIf(chkUnemployment.CheckState = CheckState.Indeterminate, GetPrice(txtUnemploymentInsurance.Text), 0)

            '      C.SPL = IIf(chkLife = vbchecked, 0.313, 0) ' life ' 1.753
            '      C.SPD = IIf(chkAccident = vbchecked, 2.21, 0) ' accident  ' 3.8
            C.SPG = IIf(chkProperty.CheckState = CheckState.Checked, 5.49, 0)  ' property
            C.IUSP = IIf(chkUnemployment.CheckState = CheckState.Checked, 2.75, 0) ' unemployment


            If chkLife.CheckState <> CheckState.Indeterminate Then txtLifeInsurance.Text = CurrencyFormat(Max(0, C.LifePremium))
            If chkAccident.CheckState <> CheckState.Indeterminate Then txtAccidentInsurance.Text = CurrencyFormat(Max(0, C.AccidentPremium))
            If chkProperty.CheckState <> CheckState.Indeterminate Then txtPropertyInsurance.Text = CurrencyFormat(Max(0, C.PropertyPremium))
            If chkUnemployment.CheckState <> CheckState.Indeterminate Then txtUnemploymentInsurance.Text = CurrencyFormat(Max(0, C.IUIPremium))

            txtFinanceAmount.Text = CurrencyFormat(C.AmountFinanced)
            txtFinanceCharges.Text = CurrencyFormat(C.FC)
            txtTotalBalance.Text = CurrencyFormat(C.B)
            If chkRoundUp.CheckState = CheckState.Checked Then
                Payment = Trunc(C.MonthlyPayment, 0) + 1
                LastPay = C.B - Payment * (C.N - 1)
            Else
                Payment = C.MonthlyPayment
                LastPay = C.MonthlyPayment
            End If
            txtPaymentWillBe.Text = CurrencyFormat(Payment)

            APR = InterestRate ' Val(txtRate) ' CalculateAPR(C.AmountFinanced, C.FC, Val(txtMonthsToFinance), Val(cboDeferred.Text))
            UpdateAPRLabel
            NewBalance = C.AmountFinanced

            FinanceCharge = C.FC

            DisposeDA(C)

            CalculateLateCharge()
            txtTotalBalance.Text = GetPrice(txtFinanceAmount.Text) + GetPrice(txtFinanceCharges.Text)
            CalculateMath()
            Exit Sub
        End If


        If Not IsElmore Then
            'If chkLife.Value <> 0 Then Life = GetLife
            If chkLife.Checked <> False Then Life = GetLife()
            'If chkAccident.Value <> 0 Then Accident = GetAcc
            If chkAccident.Checked <> False Then Accident = GetAcc()
            'If chkProperty.Value <> 0 Then Prop = GetProp
            If chkProperty.Checked <> False Then Prop = GetProp()
            'If chkUnemployment.Value <> 0 Then IUI = GetIUI
            If chkUnemployment.Checked <> False Then IUI = GetIUI()
            NewBalance = txtSubTotal.Text + DocFee + Life + Accident + Prop + IUI

            RecalculateFinancing

            If IsYeatts Then
                txtRate.Text = ""
                MaxAPR = modCustomizations.Yeatts_MaximumAPR(NewBalance)
                If APR > MaxAPR Or (txtRate.Text = "" And NewBalance > 0) Then
                    '        If txtRate <> "" Then MsgBox "APR was calculated to be " & APR & "%." & vbCrLf & "OK Maximum APR for " & FormatCurrency(NewBalance) & " is " & MaxAPR & "%." & vbCrLf & "Simple Interest Rate will be adjusted.", vbExclamation, "Automatic Simple Interest Rate Adjustment"
                    InterestRate = 0.5
                    RecalculateFinancing()

                    Do While APR > MaxAPR And Rate > 0
                        InterestRate = InterestRate - 0.0001
                        RecalculateFinancing()
                    Loop
                    txtRate.Text = Format(InterestRate * 100, "#.00")
                    SIR = InterestRate * 100
                    UpdateAPRLabel()
                End If
            End If
        End If      ' END:  If Not Elmore Then

        CalculateLateCharge

        If IsElmore Then
            Dim T As CLyndonLife
            T = New CLyndonLife

            T.I = Rate * 0.01    ' interest
            T.M = txtMonthsToFinance.Text        ' term in months
            T.LFSPR = 0.5          ' factor
            T.LA = GetPrice(txtSubTotal.Text) + GetPrice(txtDocFee.Text)         ' loan amount
            T.AHBenefitType = BenefitTypes.Retro_7day
            T.AHBenefitRate = RateTypes.SingleRate
            T.PropInsuranceRate = PropInsuranceTypes.DualInterest

            T.LifeInsuranceOn = (chkLife.Value = 1)
            T.AHInsuranceOn = (chkAccident.Value = 1)
            T.PropertyInsuranceOn = (chkProperty.Value = 1)

            T.Calculate
            txtLifeInsurance.Text = CurrencyFormat(Math.Round(T.LifeInsurance, 2))
            txtAccidentInsurance.Text = CurrencyFormat(Math.Round(T.AHInsurance, 2))
            txtPropertyInsurance.Text = CurrencyFormat(Math.Round(T.PropertyInsurance, 2))
            txtFinanceCharges.Text = CurrencyFormat(Math.Round(T.FinanceCharge, 2))
            txtPaymentWillBe.Text = CurrencyFormat(Math.Round(T.MontlyLoanPayment, 2))
            FinanceCharge = GetPrice(txtFinanceCharges.Text)

            txtFinanceAmount = CurrencyFormat(txtSubTotal + GetPrice(txtDocFee) + GetPrice(txtLifeInsurance) + GetPrice(txtAccidentInsurance) + GetPrice(txtPropertyInsurance) + GetPrice(txtUnemploymentInsurance))
            NewBalance = txtFinanceAmount
        End If    ' END:  If Elmore Then
        txtFinanceAmount = CurrencyFormat(txtSubTotal + GetPrice(txtDocFee) + GetPrice(txtLifeInsurance) + GetPrice(txtAccidentInsurance) + GetPrice(txtPropertyInsurance) + GetPrice(txtUnemploymentInsurance))
        txtTotalBalance = CurrencyFormat(GetPrice(txtFinanceAmount) + GetPrice(txtFinanceCharges))
        CalculateMath
    End Sub
    '
    'Private Function ReverseRateCalculator(ByVal Total as decimal, ByVal Rate As Double) as decimal
    '  If Rate <= 0 Or Rate = 1 Then Exit Function
    '  ReverseRateCalculator = Rate * Total / (1 - Rate)
    'End Function

    Private Sub UpdateAPRLabel()
        'Debug.Print "APR = " & APR
        lblAPR = Format(APR, "#0.00")
    End Sub

    Private Sub CalculateLateCharge()
        'calculate late charge
        If optWeekly = True Then
            LateCharge = CurrencyFormat((StoreSettings.LateChargePer * 0.01) * GetPrice(txtPaymentWillBe * 4))
        Else
            LateCharge = CurrencyFormat((StoreSettings.LateChargePer * 0.01) * GetPrice(txtPaymentWillBe))
        End If
        If StoreSettings.MaxLateCharge <> 0 Then
            If LateCharge > StoreSettings.MaxLateCharge Then LateCharge = StoreSettings.MaxLateCharge
        End If

        If StoreSettings.MinLateCharge > 0 Then  'There is a minimum late charge
            If LateCharge < StoreSettings.MinLateCharge Then LateCharge = StoreSettings.MinLateCharge
        End If
    End Sub

    Public Sub CalculateMath()
        lblMathMonthly = FormatCurrency(GetPrice(txtPaymentWillBe))
        lblMathMontlyMonths = IIf(Val(txtMonthsToFinance) > 1, Val(txtMonthsToFinance) - 1, 0)
        lblMathMonthlyTotal = FormatCurrency(GetPrice(lblMathMonthly) * Val(lblMathMontlyMonths))

        lblMathLastPay = FormatCurrency(GetPrice(txtFinanceAmount) + GetPrice(txtFinanceCharges) + GetPrice(txtFinanceChargeSalesTax) - GetPrice(lblMathMonthlyTotal))
        LastPay = GetPrice(lblMathLastPay)
        lblMathTotal = FormatCurrency(GetPrice(lblMathMonthlyTotal) + GetPrice(lblMathLastPay))
    End Sub

    Private Function ThorntonsPropertyRate() As Double
        Select Case Months
            Case 1 : ThorntonsPropertyRate = 0.16667
            Case 2 : ThorntonsPropertyRate = 0.33333
            Case 3 : ThorntonsPropertyRate = 0.49999
            Case 4 : ThorntonsPropertyRate = 0.66665
            Case 5 : ThorntonsPropertyRate = 0.83331
            Case 6 : ThorntonsPropertyRate = 0.99997
            Case 7 : ThorntonsPropertyRate = 1.16663
            Case 8 : ThorntonsPropertyRate = 1.33329
            Case 9 : ThorntonsPropertyRate = 1.49995
            Case 10 : ThorntonsPropertyRate = 1.66661
            Case 11 : ThorntonsPropertyRate = 1.83327
            Case 12 : ThorntonsPropertyRate = 2
            Case 13 : ThorntonsPropertyRate = 2.16666
            Case 14 : ThorntonsPropertyRate = 2.33332
            Case 15 : ThorntonsPropertyRate = 2.49998
            Case 16 : ThorntonsPropertyRate = 2.66664
            Case 17 : ThorntonsPropertyRate = 2.8333
            Case 18 : ThorntonsPropertyRate = 2.99996
            Case 19 : ThorntonsPropertyRate = 3.16662
            Case 20 : ThorntonsPropertyRate = 3.33328
            Case 21 : ThorntonsPropertyRate = 3.49994
            Case 22 : ThorntonsPropertyRate = 3.6666
            Case 23 : ThorntonsPropertyRate = 3.83326
            Case 24 : ThorntonsPropertyRate = 4
            Case 25 : ThorntonsPropertyRate = 4.16666
            Case 26 : ThorntonsPropertyRate = 4.33332
            Case 27 : ThorntonsPropertyRate = 4.49998
            Case 28 : ThorntonsPropertyRate = 4.66664
            Case 29 : ThorntonsPropertyRate = 4.8333
            Case 30 : ThorntonsPropertyRate = 4.99996
            Case 31 : ThorntonsPropertyRate = 5.16662
            Case 32 : ThorntonsPropertyRate = 5.33328
            Case 33 : ThorntonsPropertyRate = 5.49994
            Case 34 : ThorntonsPropertyRate = 5.6666
            Case 35 : ThorntonsPropertyRate = 5.83326
            Case 36 : ThorntonsPropertyRate = 6
            Case 37 : ThorntonsPropertyRate = 6.16666
            Case 38 : ThorntonsPropertyRate = 6.33332
            Case 39 : ThorntonsPropertyRate = 6.49998
            Case 40 : ThorntonsPropertyRate = 6.66664
            Case 41 : ThorntonsPropertyRate = 6.8333
            Case 42 : ThorntonsPropertyRate = 6.99996
            Case 43 : ThorntonsPropertyRate = 7.16662
            Case 44 : ThorntonsPropertyRate = 7.33328
            Case 45 : ThorntonsPropertyRate = 7.49994
            Case 46 : ThorntonsPropertyRate = 7.6666
            Case 47 : ThorntonsPropertyRate = 7.83326
            Case 48 : ThorntonsPropertyRate = 8
            Case 49 : ThorntonsPropertyRate = 8.16666
            Case 50 : ThorntonsPropertyRate = 8.33332
            Case 51 : ThorntonsPropertyRate = 8.49998
            Case 52 : ThorntonsPropertyRate = 8.66664
            Case 53 : ThorntonsPropertyRate = 8.8333
            Case 54 : ThorntonsPropertyRate = 8.99996
            Case 55 : ThorntonsPropertyRate = 9.16662
            Case 56 : ThorntonsPropertyRate = 9.33328
            Case 57 : ThorntonsPropertyRate = 9.49994
            Case 58 : ThorntonsPropertyRate = 9.6666
            Case 59 : ThorntonsPropertyRate = 9.83326
            Case 60 : ThorntonsPropertyRate = 10
        End Select
        ThorntonsPropertyRate = ThorntonsPropertyRate / 100.0#
    End Function

    Private Function ThorntonsLifeRate(ByVal JointLife As Boolean) As Double
        If JointLife Then
            ThorntonsLifeRate = 2.79 / 100.0#
        Else
            ThorntonsLifeRate = 1.6 / 100.0#
        End If
    End Function

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

    Private Function GetLife() As Currency
        Dim X As Currency
        If chkLife.Value = vbGrayed Then GetLife = GetPrice(txtLifeInsurance) : Exit Function
        X = (GetPrice(txtSubTotal) + GetPrice(txtDocFee))
        If IsBoyd Or IsUFO() Then
            txtLifeInsurance = 0.55 * (X / 100)
            txtLifeInsurance = txtLifeInsurance * (txtMonthsToFinance / 12)
            txtLifeInsurance = CurrencyFormat(txtLifeInsurance)
            If GetPrice(txtLifeInsurance) < 3 Then txtLifeInsurance = "3.00"
        ElseIf IsLott Then
            ' $1.60 per $100.00 per 12 months.  it is rebated at a pro-rata rate.
            txtLifeInsurance = CurrencyFormat(1.6 * (X / 100.0#) * Val(txtMonthsToFinance) / 12.0#)
        ElseIf IsCarroll Then
            txtLifeInsurance = CurrencyFormat(0.008 * X * (CDbl(txtMonthsToFinance) / 12.0#))
        ElseIf IsShaw Or IsWesternDiscount Then
            txtLifeInsurance = CurrencyFormat(0.015 * X * (CDbl(txtMonthsToFinance) / 12.0#))
        ElseIf UseThorntonsInsurance Then
            txtLifeInsurance = CurrencyFormat(AmericanHeritage_Life(X, txtMonthsToFinance, optJointLife(0), False))
        End If

        If IsElmore Then Recalculate()
        GetLife = GetPrice(txtLifeInsurance)
    End Function

    Private Function GetAcc() As Currency
        Dim X As Currency
        If chkAccident.Value = vbGrayed Then GetAcc = GetPrice(txtAccidentInsurance) : Exit Function

        X = (GetPrice(txtSubTotal) + GetPrice(txtDocFee))
        If IsLott Then
            ' computed at $3.00 per $100 per 12 months on contracts from 1 to 12 months
            ' On contracts that run from 13 to 24 months, it is $3.80 a month.
            ' We write very little A & H.  I would like the insurance defaults to automatically
            ' figure the life and property only.  Of course, I would like the option to add
            ' A & H if we like to.
            If Val(txtMonthsToFinance) <= 12 Then
                txtAccidentInsurance = CurrencyFormat(3.0# * (X / 100.0#) * Val(txtMonthsToFinance) / 12.0#)
            ElseIf Val(txtMonthsToFinance) <= 24 Then
                txtAccidentInsurance = CurrencyFormat(3.8 * (X / 100.0#) * Val(txtMonthsToFinance) / 12.0#)
            ElseIf Val(txtMonthsToFinance) <= 36 Then
                txtAccidentInsurance = CurrencyFormat(4.6 * (X / 100.0#) * Val(txtMonthsToFinance) / 12.0#)
            Else
                MsgBox "No A & H formula available for contracts greater than 36 months!", vbExclamation, ProgramMessageTitle
    End If
        End If
        If UseThorntonsInsurance Then
            txtAccidentInsurance = CurrencyFormat(AmericanHeritage_Acc(X, txtMonthsToFinance, True, 7))
        End If

        If IsElmore Then Recalculate()
        GetAcc = GetPrice(txtAccidentInsurance)
    End Function

    Private Function GetProp() As Currency
        Dim X As Currency
        If chkProperty.Value = vbGrayed Then GetProp = GetPrice(txtPropertyInsurance) : Exit Function
        X = (GetPrice(txtSubTotal) + GetPrice(txtOrigDeposit) + GetPrice(txtDocFee)) '11-26-07 added deposit so insurance is on gross
        If IsUFO() Then
            txtPropertyInsurance = 3.0# * (X / 100)
            txtPropertyInsurance = txtPropertyInsurance * (txtMonthsToFinance / 12)
            txtPropertyInsurance = CurrencyFormat(txtPropertyInsurance)
            If Val(txtPropertyInsurance) < 3 Then txtPropertyInsurance = "3.00"
        ElseIf IsBoyd Then
            'BFH2014050 - This one doesn't actually get used apparently.. it's in the Recalulate function
            txtPropertyInsurance = 2.9 * (X / 100)
            txtPropertyInsurance = txtPropertyInsurance * (txtMonthsToFinance / 12)
            txtPropertyInsurance = CurrencyFormat(txtPropertyInsurance)
            If Val(txtPropertyInsurance) < 3 Then txtPropertyInsurance = "3.00"
        ElseIf IsLott Then
            ' $3.35 per $100.00 per 12 months.  It is rebated on a rule of 78.' changed 10-8-2007
            '11-14-2007 should be based on number payment x payment like Elmore
            txtPropertyInsurance = CurrencyFormat(3.35 * (X / 100.0#) * Val(txtMonthsToFinance) / 12.0#)
        ElseIf IsCarroll Then
            txtPropertyInsurance = CurrencyFormat(X * 0.03 * ((txtMonthsToFinance) / 12.0#))
        ElseIf IsShaw Or IsWesternDiscount Then
            txtPropertyInsurance = CurrencyFormat(0.03 * X * (CDbl(txtMonthsToFinance) / 12.0#))
        ElseIf UseThorntonsInsurance Then
            txtPropertyInsurance = CurrencyFormat(AmericanHeritage_Prop(X - GetPrice(txtDocFee), txtMonthsToFinance))
        End If

        If IsElmore Then Recalculate()
        GetProp = GetPrice(txtPropertyInsurance)
    End Function

    Private Function GetIUI() As Currency
        txtUnemploymentInsurance = CurrencyFormat(0)
        GetIUI = GetPrice(txtUnemploymentInsurance)
    End Function

    Public Sub RecalculateFinancing(Optional ByVal EditingFinancing As Boolean = False)
        DeferredMonths = cboDeferred.ListIndex
        DeferredInt = (InterestRate * NewBalance) / 12 * DeferredMonths
        txtDeferredInt.Text = CurrencyFormat(DeferredInt)

        If Months = 0 Then
            If EditingFinancing Then
                FinanceCharge = GetPrice(txtFinanceCharges)
            Else
                FinanceCharge = 0 + DeferredInt
            End If
            Payment = 0
        Else
            If EditingFinancing Then
                FinanceCharge = GetPrice(txtFinanceCharges)
            Else
                If Not IsRevolvingCharge(txtArNo) And UseAlabamaSection5_19_3 Then
                    FinanceCharge = AlabamaFinanceCharges(NewBalance)
                Else
                    FinanceCharge = ((NewBalance * InterestRate) / 12 * Months)
                End If
            End If

            If StoreSettings.bInstallmentInterestIsTaxable Then
                FinanceChargeSalesTax = CurrencyFormat(StoreSettings.SalesTax * FinanceCharge)
            Else
                FinanceChargeSalesTax = CurrencyFormat(0)
            End If

            If optWeekly Then
                Payment = (NewBalance + FinanceCharge + FinanceChargeSalesTax) / (Months * 4)
            Else
                Payment = (NewBalance + FinanceCharge + FinanceChargeSalesTax) / Months
            End If
        End If

        If chkRoundUp.Value = 1 Then
            Dim Op As Currency
            Op = Payment
            Payment = Payment - ((Payment - Round(Payment, 0)))
            If Payment < Op Then Payment = Payment + 1
        End If

        If IsRevolvingCharge(txtArNo) Then
            FinanceCharge = INTEREST
            FinanceChargeSalesTax = 0
            'Payment = CalculateRevolvingPayment(RevolvingCurrentFinancedAmount(txtArNo) + GetPrice(txtTotalBalance) - GetPrice(txtPrevBalance), chkRoundUp.Value)
            Payment = CalculateRevolvingPayment(GetPrice(txtTotalBalance), chkRoundUp.Value, CLng(Months))
            APR = StoreSettings.ModifiedRevolvingRate
        Else
            If (FinanceCharge) <> 0 And (Months + DeferredMonths) <> -1 And NewBalance <> 0 Then
                APR = CalculateAPR(NewBalance, FinanceCharge, Months, DeferredMonths)
            Else
                APR = 0
            End If
        End If

        txtFinanceCharges = CurrencyFormat(FinanceCharge)
        txtFinanceChargeSalesTax = CurrencyFormat(FinanceChargeSalesTax)
        txtPaymentWillBe = CurrencyFormat(Payment)

        UpdateAPRLabel()
        UpdateTotalCaption
    End Sub

End Class
