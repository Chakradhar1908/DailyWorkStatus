Imports Microsoft.VisualBasic.Interaction
Imports VBRUN
Imports Microsoft.VisualBasic.PowerPacks.Printing.Compatibility.VB6
Imports VBA

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
    Public FromARPaySetUpForm As Boolean
    Public DBAccess_SetRecordEvent As Boolean
    Public DBAccessTransactions_SetRecordEvent As Boolean

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
            If Val(txtRate.Text) <> 0 Then
                Rate = Val(txtRate.Text) : APR = Rate
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

            A.Calculate()

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
            CalculateMath()
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

            A.Calculate()

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
            CalculateMath()
            APR = 0
            UpdateAPRLabel()
            Exit Sub
        ElseIf UseThorntonsInsurance Then
            txtPropertyInsurance.Text = CurrencyFormat(ThorntonsPropertyRate() * GetPrice(txtSubTotal.Text))
            txtLifeInsurance.Text = CurrencyFormat(ThorntonsLifeRate(optJointLife1.Checked) * GetPrice(txtSubTotal.Text))
            txtAccidentInsurance.Text = CurrencyFormat(ThorntonsAccidentRate() * GetPrice(txtSubTotal.Text))

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
            B.intr = InterestRate / 100

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

            APR = CalculateAPR(B.AmountFinanced, B.FinanceCharge, Val(txtMonthsToFinance.Text), Val(cboDeferred.Text))
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
                    B.intr = B.intr + Delta

                    APR = CalculateAPR(B.AmountFinanced, B.FinanceCharge, Val(txtMonthsToFinance.Text), Val(cboDeferred.Text))
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

            T.LifeInsuranceOn = (chkLife.Checked = True)
            T.AHInsuranceOn = (chkAccident.Checked = True)
            T.PropertyInsuranceOn = (chkProperty.Checked = True)

            T.Calculate
            txtLifeInsurance.Text = CurrencyFormat(Math.Round(T.LifeInsurance, 2))
            txtAccidentInsurance.Text = CurrencyFormat(Math.Round(T.AHInsurance, 2))
            txtPropertyInsurance.Text = CurrencyFormat(Math.Round(T.PropertyInsurance, 2))
            txtFinanceCharges.Text = CurrencyFormat(Math.Round(T.FinanceCharge, 2))
            txtPaymentWillBe.Text = CurrencyFormat(Math.Round(T.MontlyLoanPayment, 2))
            FinanceCharge = GetPrice(txtFinanceCharges.Text)

            txtFinanceAmount.Text = CurrencyFormat(txtSubTotal.Text + GetPrice(txtDocFee.Text) + GetPrice(txtLifeInsurance.Text) + GetPrice(txtAccidentInsurance.Text) + GetPrice(txtPropertyInsurance.Text) + GetPrice(txtUnemploymentInsurance.Text))
            NewBalance = txtFinanceAmount.Text
        End If    ' END:  If Elmore Then
        txtFinanceAmount.Text = CurrencyFormat(txtSubTotal.Text + GetPrice(txtDocFee.Text) + GetPrice(txtLifeInsurance.Text) + GetPrice(txtAccidentInsurance.Text) + GetPrice(txtPropertyInsurance.Text) + GetPrice(txtUnemploymentInsurance.Text))
        txtTotalBalance.Text = CurrencyFormat(GetPrice(txtFinanceAmount.Text) + GetPrice(txtFinanceCharges.Text))
        CalculateMath()
    End Sub
    '
    'Private Function ReverseRateCalculator(ByVal Total as decimal, ByVal Rate As Double) as decimal
    '  If Rate <= 0 Or Rate = 1 Then Exit Function
    '  ReverseRateCalculator = Rate * Total / (1 - Rate)
    'End Function

    Private Sub UpdateAPRLabel()
        'Debug.Print "APR = " & APR
        lblAPR.Text = Format(APR, "#0.00")
    End Sub

    Private Sub CalculateLateCharge()
        'calculate late charge
        If optWeekly.Checked = True Then
            LateCharge = CurrencyFormat((StoreSettings.LateChargePer * 0.01) * GetPrice(txtPaymentWillBe.Text * 4))
        Else
            LateCharge = CurrencyFormat((StoreSettings.LateChargePer * 0.01) * GetPrice(txtPaymentWillBe.Text))
        End If
        If StoreSettings.MaxLateCharge <> 0 Then
            If LateCharge > StoreSettings.MaxLateCharge Then LateCharge = StoreSettings.MaxLateCharge
        End If

        If StoreSettings.MinLateCharge > 0 Then  'There is a minimum late charge
            If LateCharge < StoreSettings.MinLateCharge Then LateCharge = StoreSettings.MinLateCharge
        End If
    End Sub

    Public Sub CalculateMath()
        lblMathMonthly.Text = FormatCurrency(GetPrice(txtPaymentWillBe.Text))
        lblMathMontlyMonths.Text = IIf(Val(txtMonthsToFinance.Text) > 1, Val(txtMonthsToFinance.Text) - 1, 0)
        lblMathMonthlyTotal.Text = FormatCurrency(GetPrice(lblMathMonthly.Text) * Val(lblMathMontlyMonths.Text))

        lblMathLastPay.Text = FormatCurrency(GetPrice(txtFinanceAmount.Text) + GetPrice(txtFinanceCharges.Text) + GetPrice(txtFinanceChargeSalesTax.Text) - GetPrice(lblMathMonthlyTotal.Text))
        LastPay = GetPrice(lblMathLastPay.Text)
        lblMathTotal.Text = FormatCurrency(GetPrice(lblMathMonthlyTotal.Text) + GetPrice(lblMathLastPay.Text))
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

    Private Function GetLife() As Decimal
        Dim X As Decimal
        If chkLife.CheckState = CheckState.Indeterminate Then GetLife = GetPrice(txtLifeInsurance.Text) : Exit Function
        X = (GetPrice(txtSubTotal.Text) + GetPrice(txtDocFee.Text))
        If IsBoyd Or IsUFO() Then
            txtLifeInsurance.Text = 0.55 * (X / 100)
            txtLifeInsurance.Text = txtLifeInsurance.Text * (txtMonthsToFinance.Text / 12)
            txtLifeInsurance.Text = CurrencyFormat(txtLifeInsurance.Text)
            If GetPrice(txtLifeInsurance.Text) < 3 Then txtLifeInsurance.Text = "3.00"
        ElseIf IsLott Then
            ' $1.60 per $100.00 per 12 months.  it is rebated at a pro-rata rate.
            txtLifeInsurance.Text = CurrencyFormat(1.6 * (X / 100.0#) * Val(txtMonthsToFinance.Text) / 12.0#)
        ElseIf IsCarroll Then
            txtLifeInsurance.Text = CurrencyFormat(0.008 * X * (CDbl(txtMonthsToFinance.Text) / 12.0#))
        ElseIf IsShaw Or IsWesternDiscount Then
            txtLifeInsurance.Text = CurrencyFormat(0.015 * X * (CDbl(txtMonthsToFinance.Text) / 12.0#))
        ElseIf UseThorntonsInsurance Then
            txtLifeInsurance.Text = CurrencyFormat(AmericanHeritage_Life(X, txtMonthsToFinance.Text, optJointLife0.Checked, False))
        End If

        If IsElmore Then Recalculate()
        GetLife = GetPrice(txtLifeInsurance.Text)
    End Function

    Private Sub ARPaySetUp_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'large portions moved to SetDefaultsInstallment MJK20140218
        tmrLoad.Enabled = True
        NeedPayOff = False
        AddOnAcc.Typee = ArAddOn_Nil   ' clear previous results of AddOnAcc.
        AccountFound = ""
        AddOn = ArAddOn_Nil
        FirstPayment = ""

        If IsBoyd Then
            chkRoundUp.Visible = False
        Else
            chkRoundUp.Visible = True
        End If

        txtArNo.Text = ""
        txtGrossSale.Text = ""
        txtOrigDeposit.Text = ""
        txtFinanceAmount.Text = ""
        PrevousBal = 0
        Balance = 0
        NewBalance = 0
        Contract = 0
        Zz = 0


        'lblLateChargesApplied.ToolTipText = "Grace Period: " & StoreSettings.GracePeriod & " day(s)"
        ToolTip1.SetToolTip(lblLateChargesApplied, "Grace Period: " & StoreSettings.GracePeriod & " day(s)")
        optLate6.Checked = True

        'dteDate1.Value = DateFormat(Now)
        dteDate1.Value = Date.Parse(DateFormat(Now), Globalization.CultureInfo.InvariantCulture)
        AdjustFirstPay()

        txtArNo.Visible = False 'account No
        chkAutoARNO.Visible = False

        txtPrevBalance.Visible = False 'orig balance
        lblPrevBal.Visible = False

        txtAddlPaymentsMade.Visible = False
        lblAddlPayments.Visible = False
        txtBalDueLateCharge.Visible = False
        lblBalDueLateCharge.Visible = False
        lblAcctNo.Visible = False
        lblAccountNo.Visible = False


        SetDefaultsInstallment ' MJK20140218

        If OrderMode("A", "D") Then   'New Sale or Payment - Direct Deliver Sales
            cmdCancel.Text = "Cancel Set-Up"

            If IsDate(BillOSale.lblDelDate.Text) Then
                dteDate1.Value = BillOSale.lblDelDate.Text  'delivery date
            Else
                If IsDate(BillOSale.dteSaleDate.Value) Then
                    dteDate1.Value = BillOSale.dteSaleDate.Value 'written
                Else
                    dteDate1.Value = Today
                    'MsgBox "Can't find Delivery or Written dates.  Pay special attention to the date box on this form.", vbCritical, "Warning"
                    MessageBox.Show("Can't find Delivery or Written dates.  Pay special attention to the date box on this form.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                End If
            End If

            dteDate2.Value = DateAdd("m", 1, dteDate1.Value)
            '    dteDate2 = Format(dteDate2, "mm/dd/yyyy")
            CheckLateDay()

            BillOSale.Recalculate()
            txtGrossSale.Text = CurrencyFormat(BillOSale.Written)
            txtOrigDeposit.Text = CurrencyFormat(BillOSale.Deposit)
            txtFinanceAmount.Text = CurrencyFormat(BillOSale.Sale)
            txtArNo.Visible = True
            chkAutoARNO.Visible = True
        End If

        If OrderMode("B") Then   'deliver sales
            cmdCancel.Text = "Cancel Set-Up"
            txtArNo.Visible = True
            chkAutoARNO.Visible = True
            dteDate1.Value = OrdPay.dtePayDate.Value  'delivery date
            dteDate2.Value = DateFormat(DateAdd("D", 30, dteDate1.Value))
            txtGrossSale.Text = CurrencyFormat(OrdPay.Sale)
            txtOrigDeposit.Text = CurrencyFormat(OrdPay.TotDeposit)
            txtFinanceAmount.Text = CurrencyFormat(GetPrice(txtPrevBalance.Text) + OrdPay.Sale - OrdPay.TotDeposit)
        End If

        If ArMode("E") Then ' contract estimator
            Me.Text = "Contract Estimator"
            txtPrevBalance.Visible = True
            'txtPrevBalance.TabIndex = 1
            txtPrevBalance.Text = "0.00"
            lblPrevBal.Visible = True

            txtGrossSale.Enabled = True
            txtOrigDeposit.Enabled = True
            txtFinanceAmount.Enabled = True
            'lblCashOpt.Visible = False
            cboCashOption.Visible = True
            cmdApply.Enabled = False
        End If

        If OrderMode("A", "B", "D") Then
            txtGrossSale.Enabled = False
            txtOrigDeposit.Enabled = True 'False
            txtFinanceAmount.Enabled = False
            txtMonthsToFinance.TabIndex = 1
        End If

        If ArMode("S") Then ' Set up old accounts
            txtArNo.Visible = True  'account No
            txtArNo.TabIndex = 1
            chkAutoARNO.Visible = True
            txtAddlPaymentsMade.Visible = True
            lblAddlPayments.Visible = True
            txtBalDueLateCharge.Visible = True
            lblBalDueLateCharge.Visible = True
            txtGrossSale.Enabled = True
            txtOrigDeposit.Enabled = True
            txtFinanceAmount.Enabled = True
            txtPrevBalance.Text = "0.00"

            lblTotal.Visible = True
            lblTotalCaption.Visible = True
        End If
        If ArMode("REPRINT") Then
            txtArNo.Visible = True  'account No
            txtArNo.TabIndex = 1
            txtPrevBalance.Visible = True
            lblPrevBal.Visible = True
        End If

        If OrderMode("A", "B", "D", "Credit") Or ArMode("S") Then
            'new sales, deliver sales, or old account set up
            If OrderMode("A", "B", "D") Then
                MailCheck.Index = BillOSale.Index
                BillOSale.Index = BillOSale.Index
            ElseIf OrderMode("Credit") Then
                MailCheck.Index = OnScreenReport.Index
            Else
                MailCheck.Index = BillOSale.Index
            End If
            If MailCheck.Index = "" Then MailCheck.Index = "0"

            ' This checks for previous AR Account.
            mDBAccess_Init(MailCheck.Index)
            mDBAccess.SQL = "SELECT * From InstallmentInfo Where MailIndex = " & MailCheck.Index
            mDBAccess.GetRecord()
            mDBAccess.dbClose()
            mDBAccess = Nothing

            'If AddOnAcc.lstAccounts.ListCount > 0 Then
            If AddOnAcc.lstAccounts.Items.Count > 0 Then
                If AddOnAcc.lstAccounts.Items.Count = 1 Then AddOnAcc.lstAccounts.SelectedIndex = 0
                'AddOnAcc.Show vbModal
                AddOnAcc.ShowDialog()
                AddOn = AddOnAcc.Typee

                Select Case AddOn
                    Case ArAddOn_New
                        If OrderMode("B") Then
                            Show() 'vbModal
                        Else
                            Show()
                        End If

                        txtArNo.Visible = True  'account No
                        chkAutoARNO.Visible = True
                        If IsRevolvingCharge(mArNo) Then
                            txtArNo.Text = mArNo & RevolvingSuffixLetter
                            SetDefaultsRevolving
                        Else
                            txtArNo.Text = ArAddOnAccount(mArNo)  ' mArNo & "A"
                        End If
                        ArNo = txtArNo.Text
                        mDBAccess_Init(ArNo)
                        mDBAccess.GetRecord()    ' this gets the record
                        mDBAccess.dbClose()
                        mDBAccess = Nothing
                        txtArNo.Enabled = True
                        Recalculate()
                    Case ArAddOn_Add, ArAddOn_AdT
                        'gets selected record from the list
                        Zz = 0
                        ArNo = AccountArray(AddOnAcc.lstAccounts.SelectedIndex)
                        If AddOnAcc.Revolved = True Then ArNo = AddRevolvingSuffix(ArNo)
                        mDBAccess_Init(ArNo)
                        mDBAccessTransactions_Init(ArNo)
                        mDBAccess.GetRecord()    ' this gets the record
                        mDBAccess.dbClose()
                        mDBAccess = Nothing
                        mDBAccessTransactions.dbClose()
                        mDBAccessTransactions = Nothing

                        ' BFH20051229 - we moved the payoff to arcard's OK button
                        NeedPayOff = True
                        '        Payoff

                        If OrderMode("B") Then
                            ArCard.Show()
                        Else
                            ArCard.ShowArCardForDisplayOnly(ArCard.ArNo, False)
                        End If

                        If OrderMode("B") Then
                            Show()
                        Else
                            Show()
                        End If
                        txtPrevBalance.Visible = True
                        lblPrevBal.Visible = True

                        txtArNo.Text = IIf(AddOn = ArAddOn_AdT, ArAddOnAccount(ArNo), ArNo)

                        If IsRevolvingCharge(ArNo) Then
                            txtRate.Text = StoreSettings.ModifiedRevolvingRate
                            txtMonthsToFinance.Text = StoreSettings.ModifiedRevolvingSameAsCash
                            txtDocFee.Text = ""
                        End If

                        If GetPrice(ArCard.lblTotalPayoff.Text) < GetPrice(txtPrevBalance.Text) Then txtPrevBalance.Text = GetPrice(ArCard.lblTotalPayoff.Text)
                        Recalculate()

                        'Unload AddOnAcc
                        AddOnAcc.Close()
                        AddOnAcc = Nothing
                        Erase AccountArray
                        Zz = 0
                        txtPrevBalance.Visible = True 'Previous balance
                        lblPrevBal.Visible = True
                        If OrderMode("Credit") Then
                            txtGrossSale.Text = OnScreenReport.Balance + OnScreenReport.TotTax
                        Else
                            txtPrevBalance.Enabled = False
                        End If
                        txtArNo.Enabled = False
                        chkAutoARNO.Enabled = False
                        '        Case "AddToNew"
                        '        '
                    Case Else : DevErr("Unknown AddOnAcc.Typee: " & AddOnAcc.Typee)
                End Select
            End If
        End If

        '  cboCashOption.Clear
        '  cboDeferred.Clear
        Dim Z1 As Integer
        '  For Z1 = 0 To 36
        '    cboCashOption.AddItem Z1, Z1
        '    cboDeferred.AddItem CStr(Z1) & " mos.", Z1
        '  Next Z1
        '  cboCashOption.ListIndex = 0
        '  cboDeferred.ListIndex = 0

        If False Then
            '
            '  ElseIf IsElmore Or IsLott Or IsCarroll Then
            '    chkLife.Value = 1
            '    chkAccident.Value = 1
            '    chkProperty.Value = 1
            '    chkUnemployment = 0
            '  ElseIf IsBoyd Then
            '    chkLife.Value = 1
            '    chkAccident.Value = 0
            '    chkProperty.Value = 1
            '    chkUnemployment = 0
            '  ElseIf IsMidSouth Then
            '    chkLife.Value = 0
            '    chkAccident.Value = 0
            '    chkProperty.Value = 0
            '    chkUnemployment = 0
            '  ElseIf IsShaw() Or IsWesternDiscount() Then
            '    chkLife.Value = 1
            '    chkAccident.Value = 0
            '    chkProperty.Value = 1
            '    chkUnemployment = 0
            '  ElseIf IsThorntons Then
            '    chkLife.Value = 1
            '    chkAccident.Value = 1
            '    chkProperty.Value = 1
            '    chkUnemployment = 0
        ElseIf UseAmericanNationalInsurance Then
            If IsMcClure Then
                chkLife.Checked = True
                chkAccident.Checked = True
                chkProperty.Checked = True
                chkUnemployment.Checked = False
            End If
            '  ElseIf True Then
            '    chkLife.Value = 0
            '    chkAccident.Value = 0
            '    chkProperty.Value = 0
            '    chkUnemployment = 0
        Else
            '
        End If



        If True Then
            PrevousBal = GetPrice(txtPrevBalance.Text)
            PrevousBal = CurrencyFormat(PrevousBal)
            txtFinanceAmount.Text = CurrencyFormat(GetPrice(txtSubTotal.Text) + GetPrice(txtDocFee.Text) + GetPrice(txtLifeInsurance.Text) + GetPrice(txtAccidentInsurance.Text) + GetPrice(txtPropertyInsurance.Text) + GetPrice(txtUnemploymentInsurance.Text))
        End If
        Recalculate()
    End Sub

    Private Function GetAcc() As Decimal
        Dim X As Decimal
        If chkAccident.CheckState = CheckState.Indeterminate Then GetAcc = GetPrice(txtAccidentInsurance.Text) : Exit Function

        X = (GetPrice(txtSubTotal.Text) + GetPrice(txtDocFee.Text))
        If IsLott Then
            ' computed at $3.00 per $100 per 12 months on contracts from 1 to 12 months
            ' On contracts that run from 13 to 24 months, it is $3.80 a month.
            ' We write very little A & H.  I would like the insurance defaults to automatically
            ' figure the life and property only.  Of course, I would like the option to add
            ' A & H if we like to.
            If Val(txtMonthsToFinance.Text) <= 12 Then
                txtAccidentInsurance.Text = CurrencyFormat(3.0# * (X / 100.0#) * Val(txtMonthsToFinance.Text) / 12.0#)
            ElseIf Val(txtMonthsToFinance.Text) <= 24 Then
                txtAccidentInsurance.Text = CurrencyFormat(3.8 * (X / 100.0#) * Val(txtMonthsToFinance.Text) / 12.0#)
            ElseIf Val(txtMonthsToFinance.Text) <= 36 Then
                txtAccidentInsurance.Text = CurrencyFormat(4.6 * (X / 100.0#) * Val(txtMonthsToFinance.Text) / 12.0#)
            Else
                'MsgBox "No A & H formula available for contracts greater than 36 months!", vbExclamation, ProgramMessageTitle
                MessageBox.Show("No A & H formula available for contracts greater than 36 months!", ProgramMessageTitle, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            End If
        End If
        If UseThorntonsInsurance Then
            txtAccidentInsurance.Text = CurrencyFormat(AmericanHeritage_Acc(X, txtMonthsToFinance.Text, True, 7))
        End If

        If IsElmore Then Recalculate()
        GetAcc = GetPrice(txtAccidentInsurance.Text)
    End Function

    Private Function GetProp() As Decimal
        Dim X As Decimal
        If chkProperty.CheckState = CheckState.Indeterminate Then GetProp = GetPrice(txtPropertyInsurance.Text) : Exit Function
        X = (GetPrice(txtSubTotal.Text) + GetPrice(txtOrigDeposit.Text) + GetPrice(txtDocFee.Text)) '11-26-07 added deposit so insurance is on gross
        If IsUFO() Then
            txtPropertyInsurance.Text = 3.0# * (X / 100)
            txtPropertyInsurance.Text = txtPropertyInsurance.Text * (txtMonthsToFinance.Text / 12)
            txtPropertyInsurance.Text = CurrencyFormat(txtPropertyInsurance.Text)
            If Val(txtPropertyInsurance.Text) < 3 Then txtPropertyInsurance.Text = "3.00"
        ElseIf IsBoyd Then
            'BFH2014050 - This one doesn't actually get used apparently.. it's in the Recalulate function
            txtPropertyInsurance.Text = 2.9 * (X / 100)
            txtPropertyInsurance.Text = txtPropertyInsurance.Text * (txtMonthsToFinance.Text / 12)
            txtPropertyInsurance.Text = CurrencyFormat(txtPropertyInsurance.Text)
            If Val(txtPropertyInsurance.Text) < 3 Then txtPropertyInsurance.Text = "3.00"
        ElseIf IsLott Then
            ' $3.35 per $100.00 per 12 months.  It is rebated on a rule of 78.' changed 10-8-2007
            '11-14-2007 should be based on number payment x payment like Elmore
            txtPropertyInsurance.Text = CurrencyFormat(3.35 * (X / 100.0#) * Val(txtMonthsToFinance.Text) / 12.0#)
        ElseIf IsCarroll Then
            txtPropertyInsurance.Text = CurrencyFormat(X * 0.03 * ((txtMonthsToFinance.Text) / 12.0#))
        ElseIf IsShaw Or IsWesternDiscount Then
            txtPropertyInsurance.Text = CurrencyFormat(0.03 * X * (CDbl(txtMonthsToFinance.Text) / 12.0#))
        ElseIf UseThorntonsInsurance Then
            txtPropertyInsurance.Text = CurrencyFormat(AmericanHeritage_Prop(X - GetPrice(txtDocFee.Text), txtMonthsToFinance.Text))
        End If

        If IsElmore Then Recalculate()
        GetProp = GetPrice(txtPropertyInsurance.Text)
    End Function

    Private Function GetIUI() As Decimal
        txtUnemploymentInsurance.Text = CurrencyFormat(0)
        GetIUI = GetPrice(txtUnemploymentInsurance.Text)
    End Function

    Public Sub RecalculateFinancing(Optional ByVal EditingFinancing As Boolean = False)
        DeferredMonths = cboDeferred.SelectedIndex
        DeferredInt = (InterestRate * NewBalance) / 12 * DeferredMonths
        txtDeferredInt.Text = CurrencyFormat(DeferredInt)

        If Months = 0 Then
            If EditingFinancing Then
                FinanceCharge = GetPrice(txtFinanceCharges.Text)
            Else
                FinanceCharge = 0 + DeferredInt
            End If
            Payment = 0
        Else
            If EditingFinancing Then
                FinanceCharge = GetPrice(txtFinanceCharges.Text)
            Else
                If Not IsRevolvingCharge(txtArNo.Text) And UseAlabamaSection5_19_3 Then
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

            If optWeekly.Checked = True Then
                Payment = (NewBalance + FinanceCharge + FinanceChargeSalesTax) / (Months * 4)
            Else
                Payment = (NewBalance + FinanceCharge + FinanceChargeSalesTax) / Months
            End If
        End If

        If chkRoundUp.Checked = True Then
            Dim Op As Decimal
            Op = Payment
            Payment = Payment - ((Payment - Math.Round(Payment, 0)))
            If Payment < Op Then Payment = Payment + 1
        End If

        If IsRevolvingCharge(txtArNo.Text) Then
            FinanceCharge = INTEREST
            FinanceChargeSalesTax = 0
            'Payment = CalculateRevolvingPayment(RevolvingCurrentFinancedAmount(txtArNo) + GetPrice(txtTotalBalance) - GetPrice(txtPrevBalance), chkRoundUp.Value)
            Payment = CalculateRevolvingPayment(GetPrice(txtTotalBalance.Text), chkRoundUp.Checked, CLng(Months))
            APR = StoreSettings.ModifiedRevolvingRate
        Else
            If (FinanceCharge) <> 0 And (Months + DeferredMonths) <> -1 And NewBalance <> 0 Then
                APR = CalculateAPR(NewBalance, FinanceCharge, Months, DeferredMonths)
            Else
                APR = 0
            End If
        End If

        txtFinanceCharges.Text = CurrencyFormat(FinanceCharge)
        txtFinanceChargeSalesTax.Text = CurrencyFormat(FinanceChargeSalesTax)
        txtPaymentWillBe.Text = CurrencyFormat(Payment)

        UpdateAPRLabel()
        UpdateTotalCaption()
    End Sub

    Private Sub UpdateTotalCaption()
        lblTotal.Text = FormatCurrency(GetPrice(txtFinanceAmount.Text) - GetPrice(txtAddlPaymentsMade.Text) + GetPrice(txtBalDueLateCharge.Text) + GetPrice(txtFinanceCharges.Text))
    End Sub

    Private Sub AdjustFirstPay(Optional ByVal N As Integer = 0)
        Dim T As Date
        'bfh20090309 - changed to +1 month instead of +30 days (first pay for 2/10/2009 came out to 3/12/2009 otherwise, which was wrong)
        If IsRevolvingCharge(txtArNo.Text) Then
            '    T = DateAdd("d", RevolvingStatementDay - Day(dteDate2.Value), dteDate2.Value)
            T = DateAdd("m", 1, dteDate1)
            dteDate2.Value = fAdjustFirstPay(T, cboDeferred.SelectedIndex)
        ElseIf IsPricesFurniture Then
            T = dteDate2.Value
            dteDate2.Value = fAdjustFirstPay(T, cboDeferred.SelectedIndex)
        Else
            T = DateAdd("m", 1, dteDate1.Value) ' Adjust First Payment date to Delivery+30.
            dteDate2.Value = fAdjustFirstPay(T, cboDeferred.SelectedIndex)
        End If

        CheckLateDay(N)
    End Sub

    Private Function fAdjustFirstPay(ByVal X As Date, ByVal Deferrment As Integer) As Date
        fAdjustFirstPay = IIf(Deferrment <= 0, X, DateAdd("m", Deferrment, X))
    End Function

    Private Sub SetDefaultsInstallment()
        txtDocFee.Enabled = True
        chkLife.Enabled = True
        txtLifeInsurance.Enabled = True
        chkAccident.Enabled = True
        txtAccidentInsurance.Enabled = True
        chkProperty.Enabled = True
        txtPropertyInsurance.Enabled = True
        txtMonthsToFinance.Enabled = True
        optMonthly.Checked = True
        optMonthly.Enabled = True
        optWeekly.Enabled = True
        RecalculateFinancing(True)
        cmdPrint.Enabled = True
        chkAutoARNO.Visible = (chkAutoARNO.Tag = "visible")

        If UseIUI() Then
            chkUnemployment.Visible = True
            lblUnemploymentInsurance.Visible = True
            txtUnemploymentInsurance.Visible = True
        Else
            chkUnemployment.Visible = False
            lblUnemploymentInsurance.Visible = False
            txtUnemploymentInsurance.Visible = False
        End If

        optLate16.Enabled = True
        optLate26.Enabled = True
        UpdateLateCaptions()
        '  optLate6.Caption = "Due on " & QueryDueDate(1, dteDate2, , True) & ", Late on " & QueryLateDate(1, dteDate2, , True)
        '  optLate16.Caption = "Due on " & QueryDueDate(10, dteDate2, , True) & ", Late on " & QueryLateDate(10, dteDate2, , True)
        '  optLate26.Caption = "Due on " & QueryDueDate(20, dteDate2, , True) & ", Late on " & QueryLateDate(20, dteDate2, , True)
        optLate16.Visible = True
        optLate26.Visible = True

        txtMonthsToFinance.Text = ""
        txtDocFee.Text = ""
        txtLifeInsurance.Text = ""
        txtAccidentInsurance.Text = ""
        txtPropertyInsurance.Text = ""
        txtUnemploymentInsurance.Text = ""
        txtFinanceCharges.Text = ""
        txtPaymentWillBe.Text = ""
        txtFinanceChargeSalesTax.Text = ""

        FinanceCharge = 0
        CashOpt = 0
        Payment = 0

        'use for previous balance on Add On & New accounts
        lblCashOpt.Visible = True
        cboCashOption.Visible = True
        cboDeferred.Visible = True

        optJointLife0.Checked = True
        fraJointLife.Visible = (UseAmericanNationalInsurance Or UseThorntonsInsurance Or IsLott) ' Or IsLott ' IsMidSouth Or

        txtFinanceChargeSalesTax.Visible = StoreSettings.bInstallmentInterestIsTaxable
        lblFinanceChargeSalesTax.Visible = StoreSettings.bInstallmentInterestIsTaxable

        ' BFH20111210 - Added Ordermode(A)
        lblTotalBalance.Visible = ArMode("S") Or OrderMode("A")
        txtTotalBalance.Visible = ArMode("S") Or OrderMode("A")

        'defaults
        Rate = Val(StoreSettings.SimpleInterestRate) 'Interest rate

        txtDocFee.Text = CurrencyFormat(StoreSettings.DocFee) 'file fee default
        txtMonthsToFinance.Text = "12" ' months to finance default
        If StoreSettings.bPaymentBooksMonthly Then
            optMonthly.Checked = True
        Else
            optWeekly.Checked = True
        End If
        Months = txtMonthsToFinance.Text

        cboCashOption.Items.Clear()
        cboDeferred.Items.Clear()
        Dim Z1 As Integer
        For Z1 = 0 To 36
            'cboCashOption.AddItem Z1 & " mos.", Z1
            cboCashOption.Items.Insert(Z1, Z1 & " mos.")
            'cboDeferred.AddItem CStr(Z1) & " mos.", Z1
            cboDeferred.Items.Insert(Z1, CStr(Z1) & " mos.")
        Next
        cboCashOption.SelectedIndex = 0
        cboDeferred.SelectedIndex = 0

        If False Then
            '
        ElseIf IsElmore Or IsCarroll Or IsLott Then ' Or IsLott
            chkLife.Checked = True
            chkAccident.Checked = True
            chkProperty.Checked = True
            chkUnemployment.Checked = False
        ElseIf IsBoyd Then
            chkLife.Checked = True
            chkAccident.Checked = False
            chkProperty.Checked = True
            chkUnemployment.Checked = False
            '  ElseIf IsMidSouth Then
            '    chkLife.Value = 0
            '    chkAccident.Value = 0
            '    chkProperty.Value = 0
            '    chkUnemployment = 0
        ElseIf IsShaw() Or IsWesternDiscount() Then
            chkLife.Checked = True
            chkAccident.Checked = False
            chkProperty.Checked = True
            chkUnemployment.Checked = False
        ElseIf UseThorntonsInsurance Then
            chkLife.Checked = True
            chkAccident.Checked = True
            chkProperty.Checked = True
            chkUnemployment.Checked = False
        ElseIf UseAmericanNationalInsurance Then
            chkLife.Checked = True
            chkAccident.Checked = True
            chkProperty.Checked = True
            chkUnemployment.Checked = True
        Else
            chkLife.Checked = False
            chkAccident.Checked = False
            chkProperty.Checked = False
            chkUnemployment.Checked = False
        End If

        If True Then
            PrevousBal = GetPrice(txtPrevBalance.Text)
            PrevousBal = CurrencyFormat(PrevousBal)
            txtFinanceAmount.Text = CurrencyFormat(GetPrice(txtSubTotal.Text) + GetPrice(txtDocFee.Text) + GetPrice(txtLifeInsurance.Text) + GetPrice(txtAccidentInsurance.Text) + GetPrice(txtPropertyInsurance.Text) + GetPrice(txtUnemploymentInsurance.Text))
        End If
        chkRoundUp.Checked = IIf(StoreSettings.bInstallmentRoundUp, True, False)
        Recalculate()
        CheckLateDay()
    End Sub

    Private Sub CheckLateDay(Optional ByVal vDueOn As Integer = 0)
        NoAdjust = True

        If vDueOn = 0 Then
            If IsRevolvingCharge(txtArNo.Text) Then
                optLate6.Checked = True
                '      DueOn = RevolvingStatementDay
                'DueOn = Day(dteDate1)
                DueOn = DateAndTime.Day(dteDate1.Value)
                'ElseIf Day(dteDate2) > 1 And Day(dteDate2) <= 10 Then
            ElseIf DateAndTime.Day(dteDate2.Value) > 1 And DateAndTime.Day(dteDate2.Value) <= 10 Then
                optLate16.Checked = True
                DueOn = 10
            ElseIf DateAndTime.Day(dteDate2.Value) >= 11 And DateAndTime.Day(dteDate2.Value) <= 20 Then
                optLate26.Checked = True
                DueOn = 20
            ElseIf DateAndTime.Day(dteDate2.Value) >= 21 And DateAndTime.Day(dteDate2.Value) <= 31 Or DateAndTime.Day(dteDate2.Value) = 1 Then
                optLate6.Checked = True
                DueOn = 1
            End If
        Else
            DueOn = vDueOn
        End If
        Application.DoEvents()

        Do While DateAndTime.Day(dteDate2.Value) <> DueOn
            dteDate2.Value = DateAdd("d", 1, dteDate2.Value)
        Loop

        NoAdjust = False
    End Sub

    Private Sub mDBAccess_Init(Tid As String)
        mDBAccess = New CDbAccessGeneral
        mDBAccess.dbOpen(GetDatabaseAtLocation(StoresSld))
        If AddOnAcc.Typee = ArAddOn_New Then
            mDBAccess.SQL = "SELECT * From InstallmentInfo WHERE (((ArNo)  =""" & ProtectSQL(Tid) & """))"
        Else 'checks for old accounts
            mDBAccess.SQL = "SELECT * From InstallmentInfo" _
          & " WHERE Status <> '" & arST_Void & "'" & " and ArNo=""" & ProtectSQL(Tid) & """"
        End If
    End Sub

    Private Sub mDBAccess_GetRecordNotFound()   ' called if not found
        'MsgBox "Record not found"
        mArNo = -1
        DBInterest = 0
    End Sub

    Private Sub mDBAccess_GetRecordEvent(RS As ADODB.Recordset)
        ' finds old record

        '**** show old account on screen and choose new account or Add On *****
        AccountFound = "Y"

        Do While Not RS.EOF
            Application.DoEvents()
            Dim Tid As String
            ArCard.lblBalance.Text = ""
            ArCard.lblLateCharge.Text = ""

            Status = IfNullThenNilString(RS("Status").Value)
            mArNo = IfNullThenNilString(RS("ArNo").Value)
            ArCard.lblAccount.Text = mArNo
            Telephone = IfNullThenNilString(RS("Telephone").Value)
            mBalance = IfNullThenZeroCurrency(RS("Balance").Value)

            If AddOnAcc.Typee <> ArAddOn_Nil Then
                'ArCard.Text20 = IfNullThenNilString(rs!Status)
                ArCard.ArNo = RS("ArNo").Value
                ArCard.Status = RS("Status").Value
                MailRec = IfNullThenNilString(RS("MailIndex").Value)
                ArCard.txtFinanced.Text = CurrencyFormat(IfNullThenZeroCurrency(RS("Financed").Value))
                ArCard.txtMonths.Text = IfNullThenNilString(RS("Months").Value)
                ArCard.txtRate.Text = IfNullThenNilString(RS("Rate").Value)
                ArCard.txtMonthlyPayment.Text = CurrencyFormat(IfNullThenZeroCurrency(RS("PerMonth").Value))
                ArCard.txtLateChargeAmount.Text = CurrencyFormat(IfNullThenZeroCurrency(RS("LateCharge").Value))

                ArCard.txtPaidBy.Text = IfNullThenNilString(RS("LateDueOn").Value)
                ArCard.txtDelivery.Text = DateFormat(IfNullThenNilString(RS("DeliveryDate").Value))
                ArCard.txtFirstPay.Text = DateFormat(IfNullThenNilString(RS("FirstPayment").Value))
                ArCard.txtLastPay.Text = DateAdd("m", Val(ArCard.txtMonths.Text) - 1, ArCard.txtFirstPay.Text)
                ArCard.txtSameAsCash.Text = IfNullThenNilString(RS("CashOpt").Value)
                ArCard.lblBalance.Text = IfNullThenZeroCurrency(RS("Balance").Value)
                'ArCard.TotPaid = IfNullThenNilString(rs!TotPaid)
                'allow for opening late charge balance
                ArCard.lblLateCharge.Text = IfNullThenZeroCurrency(RS("LateChargeBal").Value)
                INTEREST = IfNullThenNilString(RS("INTEREST").Value)
                DBInterest = GetPrice(INTEREST)
                txtPrevBalance.Text = CurrencyFormat(mBalance) - IIf(IsRevolvingCharge(ArNo), INTEREST, 0)
                InterestTax = IfNullThenZeroCurrency(RS("InterestSalesTax").Value)
                FinanceChargeSalesTax = IfNullThenZeroCurrency(RS("InterestSalesTax").Value)
                Life = IfNullThenZeroCurrency(RS("Life").Value)
                Accident = IfNullThenZeroCurrency(RS("Accident").Value)
                Prop = IfNullThenZeroCurrency(RS("Prop").Value)
                IUI = IfNullThenZeroCurrency(RS("IUI").Value)

                SendNotice = IfNullThenNilString(RS("SendNotice").Value)

                ArCard.GetCustomer()
                ArCard.GetCust()
                ArCard.GetPayoff()
                ArCard.GetAgeing()
            End If

            If Status <> arST_Void And Not ArNoIsAddOnRecord(mArNo) Then
                'AddOnAcc.lstAccounts.AddItem " " & mArNo & "        " & Telephone & "  Balance: " & CurrencyFormat(mBalance)
                AddOnAcc.lstAccounts.Items.Add(" " & mArNo & "        " & Telephone & "  Balance: " & CurrencyFormat(mBalance))
                AccountArray(Zz) = Trim(mArNo)
                Zz = Zz + 1
            End If

            RS.MoveNext()
        Loop
    End Sub

    Public Sub mDBAccess_SetRecordEvent(RS As ADODB.Recordset)
        On Error GoTo ErrorHandler
        'called to write the record to info file

        If IsIn(AddOn, ArAddOn_Add) Then
            'BFH20170612 - This will back up the [InstallmentInfo] information for a given account when doing an Add-On
            AddOnRecordAccount = ArAddOnCreateContractHistoryAccount(ArNo, StoresSld)
        ElseIf IsIn(AddOn, ArAddOn_AdT) Then
            'BFH20170707 - Add On to New Account:  Payoff and close existing account, add it to a NEW ACCOUNT as a previous balance
            ArAddOnToNewCloseOutAccount(StoresSld, mArNo, ArNo, GetPrice(txtPrevBalance.Text))
        End If

        If ArMode("S") Or OrderMode("A", "B") Then
            BillOSale.Index = BillOSale.Index
        End If

        RS("ArNo").Value = Trim(ArNo)
        RS("MailIndex").Value = Trim(BillOSale.Index)
        RS("LastName").Value = Trim(BillOSale.CustomerLast.Text)

        RS("Telephone").Value = CleanAni(BillOSale.CustomerPhone1.Text)
        RS("Financed").Value = CurrencyFormat(NewBalance + FinanceCharge + FinanceChargeSalesTax)
        RS("Months").Value = Trim(Months)
        RS("Rate").Value = Format(SIR * 100, "0.00")
        RS("APR").Value = Format(APR, "0.00")
        RS("PerMonth").Value = Trim(txtPaymentWillBe.Text)
        RS("LateDueOn").Value = Trim(DueOn)
        If Val(cboCashOption.SelectedIndex) < 0 Then cboCashOption.SelectedIndex = 0
        RS("CashOpt").Value = cboCashOption.SelectedIndex

        'calculate late charge
        LateCharge = CurrencyFormat((StoreSettings.LateChargePer * 0.01) * GetPrice(txtPaymentWillBe.Text))
        If StoreSettings.MaxLateCharge <> 0 Then
            If LateCharge > StoreSettings.MaxLateCharge Then LateCharge = StoreSettings.MaxLateCharge
        End If

        If StoreSettings.MinLateCharge > 0 Then 'There is a minimum late charge
            If LateCharge < StoreSettings.MinLateCharge Then LateCharge = StoreSettings.MinLateCharge
        End If

        RS("LateCharge").Value = LateCharge
        RS("DeliveryDate").Value = dteDate1.Value

        '  FirstPayment = DateAdd("m", 1, Format(dteDate1.Value, "mm/" & DueOn & "/yyyy"))
        '  FirstPayment = fAdjustFirstPay(FirstPayment, cboDeferred.ListIndex) ' BFH20061026 - Added b/c it was wrong.. contracts and coupons were broken
        '  If DateDiff("d", dteDate2, FirstPayment) < 0 Then
        '    FirstPayment = DateAdd("m", 1, FirstPayment)
        '  End If

        RS("FirstPayment").Value = dteDate2.Value ' FirstPayment
        RS("Balance").Value = Math.Round(NewBalance + FinanceCharge + FinanceChargeSalesTax, 2)
        RS("LateChargeBal").Value = "0.00"
        RS("Status").Value = "O"
        RS("Interest").Value = GetPrice(txtFinanceCharges.Text)
        RS("InterestSalesTax").Value = GetPrice(txtFinanceChargeSalesTax.Text)
        RS("Life").Value = GetPrice(txtLifeInsurance.Text)
        RS("LifeType").Value = IIf(optJointLife1.Checked, 1, 0)
        RS("Accident").Value = GetPrice(txtAccidentInsurance.Text)
        RS("Prop").Value = GetPrice(txtPropertyInsurance.Text)
        RS("IUI").Value = GetPrice(txtUnemploymentInsurance.Text)
        RS("TotPaid").Value = GetPrice("0")

        RS("Period").Value = Switch(optMonthly.Checked = True, "M", optWeekly, "W", True, IIf(StoreSettings.bPaymentBooksMonthly, "M", "W"))

        If ArMode("S") Then
            'opening accounts previous payments
            If Val(txtAddlPaymentsMade.Text) > 0 Then
                '      rs("Balance") = Round(CurrencyFormat(NewBalance + FinanceCharge + FinanceChargeSalesTax - GetPrice(txtAddlPaymentsMade)), 2)
                RS("TotPaid").Value = CurrencyFormat(GetPrice(txtAddlPaymentsMade.Text))
            End If

            'allow for opening late charge balance
            If Val(txtBalDueLateCharge.Text) > 0 Then
                '      rs("Balance") = Round(CurrencyFormat(NewBalance + FinanceCharge + FinanceChargeSalesTax - GetPrice(txtAddlPaymentsMade) + GetPrice(txtBalDueLateCharge)), 2)
                RS("LateChargeBal").Value = CurrencyFormat(GetPrice(txtBalDueLateCharge.Text))
            End If

            RS("Balance").Value = Math.Round(CDec(CurrencyFormat(NewBalance + FinanceCharge + FinanceChargeSalesTax)) - GetPrice(txtAddlPaymentsMade.Text) + GetPrice(txtBalDueLateCharge.Text), 2)
        End If
        Exit Sub
ErrorHandler:
        If IsDevelopment() Then MessageBox.Show("Trouble with save:" & Err.Description, "Developer Alert", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        Err.Clear()
        Resume Next
    End Sub

    Public Sub mDBAccessTransactions_SetRecordEvent(RS As ADODB.Recordset)   ' called to write the record
        '    RS("ArNo") = mArNo
        RS("ArNo").Value = ArNo
        If BillOSale.Index = 0 Then BillOSale.Index = MailRec 'this is for Add on sales from old account set up
        RS("MailIndex").Value = BillOSale.Index
        RS("LastName").Value = BillOSale.CustomerLast.Text
        RS("TransDate").Value = dteDate1.Value

        RS("Type").Value = TransType
        RS("Charges").Value = CurrencyFormat(Charges)
        RS("Credits").Value = CurrencyFormat(Credits)
        RS("Balance").Value = Math.Round(CDec(CurrencyFormat(Balance)), 2)

        If IsIn(AddOn, ArAddOn_Add, ArAddOn_AdT) And AddOnRecordAccount <> "" Then
            RS("Receipt").Value = AddOnRecordAccount        ' On an add-on, record the history account #
            AddOnRecordAccount = ""
        End If
    End Sub

    Private Sub CalcFirstPayment()
    End Sub

    Private Sub SetDefaultsRevolving()
        txtDocFee.Text = "0.00"
        txtDocFee.Enabled = False
        chkLife.Checked = False
        chkLife.Enabled = False
        txtLifeInsurance.Text = "0.00"
        txtLifeInsurance.Enabled = False
        chkAccident.Checked = False
        chkAccident.Enabled = False
        txtAccidentInsurance.Text = "0.00"
        txtAccidentInsurance.Enabled = False
        chkProperty.Checked = False
        chkProperty.Enabled = False
        txtPropertyInsurance.Text = "0.00"
        txtPropertyInsurance.Enabled = False
        txtMonthsToFinance.Text = "3"
        txtMonthsToFinance.Enabled = False
        Months = 3
        chkUnemployment.Checked = False
        chkUnemployment.Visible = False
        lblUnemploymentInsurance.Visible = False
        txtUnemploymentInsurance.Visible = False
        optMonthly.Checked = False
        optMonthly.Enabled = False
        optWeekly.Checked = False
        optWeekly.Enabled = False
        cboCashOption.SelectedIndex = RevolvingSameAsCash()
        FinanceCharge = DBInterest ' Must carry through for existing sales, but be zero for new revolving accounts
        INTEREST = FinanceCharge
        txtFinanceCharges.Text = CurrencyFormat(FinanceCharge)
        NewBalance = GetPrice(txtSubTotal.Text)
        txtFinanceAmount.Text = CurrencyFormat(GetPrice(txtSubTotal.Text) + GetPrice(txtDocFee.Text) + GetPrice(txtLifeInsurance.Text) + GetPrice(txtAccidentInsurance.Text) + GetPrice(txtPropertyInsurance.Text) + GetPrice(txtUnemploymentInsurance.Text))

        Rate = StoreSettings.ModifiedRevolvingRate
        If StoreSettings.ModifiedRevolvingAPR Then
            InterestRate = Rate * 0.01 / 12 'CalculateSIR(NewBalance, Rate, Months)
        Else
            InterestRate = Rate * 0.01
        End If
        SIR = InterestRate

        txtTotalBalance.Text = CurrencyFormat(GetPrice(txtSubTotal.Text))
        cmdPrint.Enabled = False
        optLate6.Text = "Due on the Delivery Day"
        '  optLate6.Caption = "Due on the " & Ordinal(RevolvingStatementDay) & ", Late on " & Ordinal(RevolvingStatementDay)
        optLate16.Visible = False
        optLate26.Visible = False
        '  optLate6.Value = True
        '  DueOn = RevolvingStatementDay
        optLate16.Enabled = False
        optLate26.Enabled = False
        AdjustFirstPay()

        If chkAutoARNO.Tag = "" Then
            If chkAutoARNO.Visible Then
                chkAutoARNO.Tag = "visible"
            Else
                chkAutoARNO.Tag = "invisible"
            End If
        End If
        chkAutoARNO.Visible = False

        Recalculate() 'Financing True
        CheckLateDay()
    End Sub

    Private Sub mDBAccessTransactions_Init(ByVal Tid As String)
        mDBAccessTransactions = New CDbAccessGeneral
        mDBAccessTransactions.dbOpen(GetDatabaseAtLocation())
        mDBAccessTransactions.SQL = "SELECT * From Transactions WHERE (((ArNo)=""" & ProtectSQL(Tid) & """))"  ' Changed to tid from mArNo - MJK 20041122
    End Sub

    Private Sub UpdateLateCaptions()
        'BFH20170713 - Changed to show Grace applied LATE DATE
        ' I believe these should show the LATE date...  typically, GRACE is only applied on late notices, and that is what this is about
        optLate6.Text = "Due on 1st, Late on " & QueryLateDate(1, dteDate2.Value, , False)
        optLate16.Text = "Due on 10th, Late on " & QueryLateDate(10, dteDate2.Value, , False)
        optLate26.Text = "Due on 20th, Late on " & QueryLateDate(20, dteDate2.Value, , False)
    End Sub

    Private Function NeedCreditApp() As Boolean
        Dim SQL As String, RS As ADODB.Recordset

        NeedCreditApp = False
        If Not UseIUI() Then Exit Function
        If ArMode("E") Then Exit Function       ' not in estimator

        SQL = "SELECT HisAge, HisSS, CoName, CoAddress, CoCityState, CoSS, CoAge FROM [ArApp] WHERE MailIndex='" & BillOSale.MailIndex & "'"
        RS = GetRecordsetBySQL(SQL, , GetDatabaseAtLocation())
        If RS.RecordCount = 0 Then NeedCreditApp = True
        If Not NeedCreditApp Then
            If MsgBox("Edit Credit Application?", vbQuestion + vbYesNo + vbDefaultButton2, "Credit App On File") = vbYes Then NeedCreditApp = True
        End If
    End Function

    Private Sub cmdApply_Click(sender As Object, e As EventArgs) Handles cmdApply.Click
        'Handle failure cases before changing mouse pointer or other things.
        Working(True)

        If NeedCreditApp() Then CreditApp()

        If NeedPayOff Then Payoff()

        If GetPrice(txtMonthsToFinance.Text) <= 0 And Not IsRevolvingCharge(txtArNo.Text) Then
            'MsgBox "Please enter the length of the financing period.", vbExclamation
            MessageBox.Show("Please enter the length of the financing period.")
            txtMonthsToFinance.Select()
            SelectContents(txtMonthsToFinance)
            Working(False)
            Exit Sub
        End If

        '  If SIR <= 0 Then
        '    MsgBox "Please enter the interest rate.", vbExclamation
        '    txtRate.SetFocus
        '    SelectContents txtRate
        '    Working False
        '    Exit Sub
        '  End If

        'deliver sales AddOn
        If OrderMode("A", "B") Or (OrderMode("D") And AddOnAcc.Typee <> ArAddOn_Add) Or ArMode("S") Then
            If Trim(txtArNo.Text) = "" Then ' Or Val(txtArNo) = 0 Then ' added, then removed again:  bfh20061113
                'MsgBox "You Must Enter An Account Number!", vbCritical
                MessageBox.Show("You Must Enter An Account Number!")
                On Error Resume Next
                txtArNo.Select()
                Working(False)
                Exit Sub
            ElseIf ArNoExists(txtArNo.Text) <> IsIn(AddOn, ArAddOn_Add) Then
                If IsIn(AddOn, ArAddOn_Add) Then
                    'MsgBox "This account number could not be found, so it can't be added on to.", vbCritical
                    MessageBox.Show("This account number could not be found, so it can't be added on to.")
                    On Error Resume Next
                    txtArNo.Select()
                    Working(False)
                    Exit Sub
                Else
                    'MsgBox "This account number is already in use.  Please enter another.", vbCritical
                    MessageBox.Show("This account number is already in use.  Please enter another.")
                    On Error Resume Next
                    txtArNo.Select()
                    Working(False)
                    Exit Sub
                End If
            End If
        End If

        If OrderMode("A") Then     ' Save txtArNo as a note on the bill of sale.
            BillOSale.SetDesc(BillOSale.NewStyleLine - 1, BillOSale.QueryDesc(BillOSale.NewStyleLine - 1) & " Account #" & txtArNo.Text)
            BillOSale.InstallmentTotal = GetPrice(txtFinanceAmount.Text) + GetPrice(txtFinanceCharges.Text) + GetPrice(txtFinanceChargeSalesTax.Text)

            Dim PSRes As Boolean
            '    If BillOSale.UseNewProcessSale Then
            PSRes = BillOSale.ProcessSale2
            '    Else
            '      PSRes = BillOSale.ProcessSale
            '    End If
            If Not PSRes Then
                'MsgBox "Can't process Store Finance because Sale failed to save.", vbCritical, "Error",
                MessageBox.Show("Can't process Store Finance because Sale failed to save.")
                Working(False)
                Exit Sub
            End If
        End If

        'no contract - Post to data base
        If ArMode("S") Then  ' Orig accounts
            '    If Val(txtGrossSale) = 0 Then 'no Orig sale amount
            '     If MsgBox("This Contract Is Incomplete!  You Must Enter The Gross Sale W/Tax", vbExclamation, "Incomplete Sale") = vbOK Then
            '       txtGrossSale.SetFocus
            '       cmdPrint.Enabled = True
            '       cmdApply.Enabled = True
            '       Working False
            '       Exit Sub
            '     End If
            '    End If

            If txtAddlPaymentsMade.Text = "" Then
                'If MsgBox("Do You Want To Add Any Previous Payments!", vbExclamation + vbYesNo) = vbYes Then
                If MessageBox.Show("Do You Want To Add Any Previous Payments!", "", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) = DialogResult.Yes Then
                    cmdPrint.Enabled = True
                    cmdApply.Enabled = True
                    FocusControl(txtAddlPaymentsMade)
                    Working(False)
                    Exit Sub
                End If
            End If
        End If

        'MousePointer = vbHourglass
        Cursor = Cursors.WaitCursor
        HandleBillOSaleControls(False)
        cmdCancel.Text = "Done"

        UnloadARPaySetUp = False 'processed sale for

        If OrderMode("A") Then   'deliver sales
            cmdCancel.Enabled = True
            'Unload ArCard
            ArCard.Close()
            Show()
        End If

        Dim Printed As Boolean
        If Not ArMode("E") Then   ' No contract estimator
            ProcessAccount()

            Printed = True
            'If cmdPrint.Value Then  ---> In vb6 code, cmdPrint is a command button. It will be replaced with checkbox in vb.net. Cause If button.value conditional using is not possible in vb.net
            If cmdPrint.Checked = True Then
                PrintContract()
                If StoreSettings.bPrintPaymentBooks Then PrintCoupons()
            End If


            Dim WasB As Boolean
            WasB = OrderMode("B")

            If OrderMode("D") Then
                OrdPay.FinanceOnAccount(ArNo)
                'Unload ArCard
                ArCard.Close()
                'Unload Me
                Me.Close()
                Exit Sub
            Else
                OrdPay.FinishRoutine(False)
            End If

            HandleBillOSaleControls(True)

            Working(False, False)
            '    cmdPrint.Enabled = False 'print & apply
            '    cmdApply.Enabled = False

            If WasB Or OrderMode("B") Then Exit Sub


            If OrderMode("Credit") Then
                'Unload ArCard
                ArCard.Close()
            End If

            If Not ArMode("S") Then  'exit out of this sub from order "B"
                'MousePointer = 0
                Cursor = Cursors.Default
                'Unload Me ' MJK 20140220
                Me.Close()
                Exit Sub
            End If
        End If

        If ArMode("S") Then
            'If MsgBox("Any More New Accounts To Open?", vbQuestion + vbYesNo) = vbYes Then
            If MessageBox.Show("Any More New Accounts To Open?", "", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
                Domain_exit()
                'Unload Me
                Me.Close()
                'Unload ArCard
                ArCard.Close()
                'Unload OrdPay
                OrdPay.Close()
                BillOSale.Show()
                MailCheck.optTelephone.Checked = True
                'MailCheck.Show vbModal, BillOSale
                MailCheck.ShowDialog(BillOSale)
                Exit Sub
            Else
                Domain_exit()
                'Unload BillOSale
                BillOSale.Close()
                modProgramState.ArSelect = ""
                'Unload ArCard
                ArCard.Close()
                'Unload OrdPay
                OrdPay.Close()
                'Unload ARPaySetUp
                Me.Close()
                MainMenu.Show()
                Exit Sub
            End If
        End If

        If Not OrderMode("B", "A") Then
            'If cmdPrint.Value = True And Not Printed Then  --> Replaced button with checkbox. Cause If cmdPrint.Value = True is not possible in vb.net with button.
            If cmdPrint.Checked = True And Not Printed Then
                PrintContract()
                If StoreSettings.bPrintPaymentBooks Then PrintCoupons()
            End If

            'MousePointer = 0
            Cursor = Cursors.Default
            'Unload Me
            Me.Close()
            ArSelect = ""
        Else
            'MousePointer = 0
            Cursor = Cursors.Default
        End If
    End Sub

    Private Sub CreditApp()
        'Load ArApp    ------> Load and Unload methods are not valid in vb.net. For Load formname, use formname.show only.
        ArApp.Show()
        ArApp.GetApp(BillOSale.MailIndex, txtArNo.Text)

        If Trim(ArApp.txtFirstName.Text) = "" And Trim(ArApp.txtLastName.Text) = "" Then
            If Not UseAmericanNationalInsurance Then
                ArApp.txtFirstName = BillOSale.CustomerFirst
                ArApp.txtLastName = BillOSale.CustomerLast
            Else
                If InStr(BillOSale.CustomerFirst.Text, "&") Then
                    ArApp.txtFirstName.Text = SplitWord(BillOSale.CustomerFirst.Text, 1, "&")
                    ArApp.txtLastName.Text = BillOSale.CustomerLast.Text
                    ArApp.txtCoName.Text = SplitWord(BillOSale.CustomerFirst.Text, 2, "&") & " " & BillOSale.CustomerLast.Text
                End If
            End If
        End If


        If Trim(ArApp.txtAddress.Text) = "" Then ArApp.txtAddress.Text = BillOSale.CustomerAddress.Text
        If Trim(ArApp.txtCity.Text) = "" Then ArApp.txtCity.Text = BillOSale.CustomerCity.Text
        If Trim(ArApp.txtZip.Text) = "" Then ArApp.txtZip.Text = BillOSale.CustomerZip.Text
        If Trim(ArApp.txtTele1.Text) = "" Then ArApp.txtTele1.Text = BillOSale.CustomerPhone1.Text
        If Trim(ArApp.txtTele2.Text) = "" Then ArApp.txtTele2.Text = BillOSale.CustomerPhone2.Text

        ArApp.txtAccount.Text = txtArNo.Text
        'ArApp.Show 1
        ArApp.ShowDialog()
        If IsFormLoaded("ArApp") Then
            'Unload ArApp
            ArApp.Close()
        End If
    End Sub

    Private Sub Working(Optional ByVal Working As Boolean = False, Optional ByVal CanSubmit As Boolean = True)
        If Working Then CanSubmit = False

        'MousePointer = IIf(Working, vbHourglass, vbDefault)
        Me.Cursor = IIf(Working, Cursors.WaitCursor, Cursors.Default)
        EnableFrame(Me, fraARPaySetup, Not Working)
        'cmdCancel.Enabled = Not Working

        cmdApply.Enabled = CanSubmit
        cmdPrint.Enabled = CanSubmit
    End Sub

    Public Sub Payoff()
        If mBalance = 0 Then Exit Sub

        mDBAccessTransactions_Init(ArNo)
        mDBAccessTransactions.GetRecord()    ' this gets the record

        mDBAccessTransactions.SQL = "SELECT * From Transactions WHERE ArNo=""-1"""
        ArNo = mArNo

        '  If Val(ArCard.DocCredit) > 0 Then  'Doc
        '    TransType = "Doc Payoff"
        '    Charges = "0"
        '    Credits = ArCard.DocCredit
        '    Balance = ArCard.lblBalance - ArCard.DocCredit
        '    mDBAccessTransactions.SetRecord
        '  End If
        '
        If Val(ArCard.LifeCredit) > 0 Then  'Life
            TransType = arPT_poLif
            Charges = "0"
            Credits = ArCard.LifeCredit
            Balance = GetPrice(ArCard.lblBalance.Text) - ArCard.LifeCredit
            mDBAccessTransactions.SetRecord()
        End If

        If Val(ArCard.AccidentCredit) > 0 Then  'Acc
            TransType = arPT_poAcc
            Charges = "0"
            Credits = ArCard.AccidentCredit
            Balance = GetPrice(ArCard.lblBalance.Text) - ArCard.LifeCredit - ArCard.AccidentCredit
            mDBAccessTransactions.SetRecord()
        End If

        If Val(ArCard.PropertyCredit) > 0 Then  'Prop
            TransType = arPT_poPro
            Charges = "0"
            Credits = ArCard.PropertyCredit
            Balance = GetPrice(ArCard.lblBalance.Text) - ArCard.LifeCredit - ArCard.AccidentCredit - ArCard.PropertyCredit
            mDBAccessTransactions.SetRecord()
        End If

        If Val(ArCard.IUICredit) > 0 Then  'IUI
            TransType = arPT_poIUI
            Charges = "0"
            Credits = ArCard.IUICredit
            Balance = GetPrice(ArCard.lblBalance.Text) - ArCard.LifeCredit - ArCard.AccidentCredit - ArCard.PropertyCredit - ArCard.IUICredit
            mDBAccessTransactions.SetRecord()
        End If

        If Val(ArCard.InterestCredit) > 0 Then  'C/C
            TransType = arPT_poInt
            Charges = "0"
            Credits = ArCard.InterestCredit
            Dim lblARC8 As Decimal
            lblARC8 = IIf(GetPrice(ArCard.lblBalance.Text) = 0, 0, GetPrice(ArCard.lblBalance.Text))
            Balance = GetPrice(lblARC8) - ArCard.LifeCredit - ArCard.AccidentCredit - ArCard.PropertyCredit - ArCard.IUICredit - ArCard.InterestCredit
            mDBAccessTransactions.SetRecord()
        End If

        If Val(ArCard.InterestTaxCredit) > 0 Then  'C/C
            TransType = arPT_poTax
            Charges = "0"
            Credits = ArCard.InterestTaxCredit
            Balance = Balance - ArCard.InterestTaxCredit
            mDBAccessTransactions.SetRecord()
        End If

        'bfh20060905
        Balance = GetPrice(ArCard.lblBalance.Text) - ArCard.LifeCredit - ArCard.AccidentCredit - ArCard.PropertyCredit - ArCard.IUICredit - ArCard.InterestCredit - ArCard.InterestTaxCredit

        mDBAccessTransactions.dbClose()
        mDBAccessTransactions = Nothing
        txtPrevBalance.Text = Balance
    End Sub

    Private Sub ProcessAccount()
        ArNo = txtArNo.Text

        'BFH20170707 - Removed the following, because.... It was done on the line above it no matter what.
        '  If OrderMode("B", "A") And IsNotIn(AddOn, ArAddOn_Add, ArAddOn_AdT) Then 'deliver sales/Add on
        '    ArNo = txtArNo
        '  End If
        '
        'addon account put in the A/R no automatically
        If ArMode("S") Then    'old customer accounts / Add on
            ArNo = txtArNo.Text
            If Trim(ArNo) = "" Then Exit Sub
        End If

        mDBAccess_Init(Trim(ArNo))
        'FromARPaySetUpForm = True
        DBAccess_SetRecordEvent = True
        mDBAccess.SetRecord()
        'FromARPaySetUpForm = False
        DBAccess_SetRecordEvent = False
        mDBAccess.dbClose()
        mDBAccess = Nothing


        If BillOSale.BillOfSale.Text = "" Then
            TransType = arPT_New & " " & GetSaleNoFromArNo(ArNo)
        Else
            TransType = arPT_New & " " & BillOSale.BillOfSale.Text
        End If
        Charges = GetPrice(txtGrossSale.Text)
        ' Only add previous balance into contract amount if is an Add-OnTo (Add-Ons just use a running total from the previous account)
        If IsIn(AddOn, ArAddOn_AdT) Then Charges = Charges + GetPrice(txtPrevBalance.Text)
        Credits = GetPrice(txtOrigDeposit.Text)
        Balance = (GetPrice(txtPrevBalance.Text) + GetPrice(txtGrossSale.Text)) - GetPrice(txtOrigDeposit.Text)

        mDBAccessTransactions_Init("-1")
        DBAccessTransactions_SetRecordEvent = True
        mDBAccessTransactions.SetRecord()
        DBAccessTransactions_SetRecordEvent = False

        If GetPrice(txtDocFee.Text) <> 0 Then  'doc fee
            TransType = arPT_Doc
            Charges = GetPrice(txtDocFee.Text)
            Credits = 0
            Balance = (GetPrice(txtPrevBalance.Text) + GetPrice(txtGrossSale.Text) + txtDocFee.Text) - GetPrice(txtOrigDeposit.Text)
            DBAccessTransactions_SetRecordEvent = True
            mDBAccessTransactions.SetRecord()
            DBAccessTransactions_SetRecordEvent = False
            '    Cash BillOSale.BillOfSale, GetPrice(txtDocFee), "40320", "Doc Fee"
        End If

        If GetPrice(txtLifeInsurance.Text) > 0 Then  'Life
            TransType = arPT_Lif
            Charges = GetPrice(txtLifeInsurance.Text)
            Credits = 0
            Balance = (GetPrice(txtPrevBalance.Text) + GetPrice(txtGrossSale.Text) + txtDocFee.Text + GetPrice(txtLifeInsurance.Text)) - GetPrice(txtOrigDeposit.Text)
            DBAccessTransactions_SetRecordEvent = True
            mDBAccessTransactions.SetRecord()
            DBAccessTransactions_SetRecordEvent = False
            '    Cash BillOSale.BillOfSale, GetPrice(txtLifeInsurance), "40330", "Life Ins."  ' 40330 == life single, 40340 == life joint
        End If

        If GetPrice(txtAccidentInsurance.Text) > 0 Then  'Acc
            TransType = arPT_Acc
            Charges = GetPrice(txtAccidentInsurance.Text)
            Credits = 0
            Balance = (GetPrice(txtPrevBalance.Text) + GetPrice(txtGrossSale.Text) + txtDocFee.Text + GetPrice(txtLifeInsurance.Text) + GetPrice(txtAccidentInsurance.Text)) - GetPrice(txtOrigDeposit.Text)
            DBAccessTransactions_SetRecordEvent = True
            mDBAccessTransactions.SetRecord()
            DBAccessTransactions_SetRecordEvent = False
            '    Cash BillOSale.BillOfSale, GetPrice(txtAccidentInsurance), "40350", "Acc. Ins."
        End If

        If GetPrice(txtPropertyInsurance.Text) > 0 Then  'Prop
            TransType = arPT_Pro
            Charges = GetPrice(txtPropertyInsurance.Text)
            Credits = 0
            Balance = (GetPrice(txtPrevBalance.Text) + GetPrice(txtGrossSale.Text) + txtDocFee.Text + GetPrice(txtLifeInsurance.Text) + GetPrice(txtAccidentInsurance.Text) + GetPrice(txtPropertyInsurance.Text)) - GetPrice(txtOrigDeposit.Text)
            DBAccessTransactions_SetRecordEvent = True
            mDBAccessTransactions.SetRecord()
            DBAccessTransactions_SetRecordEvent = False
            '    Cash BillOSale.BillOfSale, GetPrice(txtPropertyInsurance), "40360", "Prop. Ins."
        End If

        If GetPrice(txtUnemploymentInsurance.Text) > 0 Then  'IUI
            TransType = arPT_IUI
            Charges = GetPrice(txtUnemploymentInsurance.Text)
            Credits = 0
            Balance = (GetPrice(txtPrevBalance.Text) + GetPrice(txtGrossSale.Text) + txtDocFee.Text + GetPrice(txtLifeInsurance.Text) + GetPrice(txtAccidentInsurance.Text) + GetPrice(txtPropertyInsurance.Text) + GetPrice(txtUnemploymentInsurance.Text)) - GetPrice(txtOrigDeposit.Text)
            DBAccessTransactions_SetRecordEvent = True
            mDBAccessTransactions.SetRecord()
            DBAccessTransactions_SetRecordEvent = False
            '    Cash BillOSale.BillOfSale, GetPrice(txtUnemploymentInsurance), "40360", "IUI Ins."
        End If

        If GetPrice(txtFinanceCharges.Text) > 0 Then  'C/C
            TransType = arPT_Int
            Charges = GetPrice(txtFinanceCharges.Text)
            Credits = 0
            Balance = (GetPrice(txtPrevBalance.Text) + GetPrice(txtGrossSale.Text) + txtDocFee.Text + GetPrice(txtLifeInsurance.Text) + GetPrice(txtAccidentInsurance.Text) + GetPrice(txtPropertyInsurance.Text) + GetPrice(txtUnemploymentInsurance.Text) + GetPrice(txtFinanceCharges.Text)) - GetPrice(txtOrigDeposit.Text)
            DBAccessTransactions_SetRecordEvent = True
            mDBAccessTransactions.SetRecord()
            DBAccessTransactions_SetRecordEvent = False
            '    Cash BillOSale.BillOfSale, GetPrice(txtFinanceCharges), "40370", "Interest Chg."
        End If

        If FinanceChargeSalesTax > 0 Then  'Finance Charge Sales Tax
            TransType = arPT_Tax
            Charges = GetPrice(FinanceChargeSalesTax)
            Credits = "0"
            Balance = (GetPrice(txtPrevBalance.Text) + GetPrice(txtGrossSale.Text) + txtDocFee.Text + GetPrice(txtLifeInsurance.Text) + GetPrice(txtAccidentInsurance.Text) + GetPrice(txtPropertyInsurance.Text) + GetPrice(txtUnemploymentInsurance.Text) + GetPrice(txtFinanceCharges.Text) + GetPrice(FinanceChargeSalesTax)) - GetPrice(txtOrigDeposit.Text)
            DBAccessTransactions_SetRecordEvent = True
            mDBAccessTransactions.SetRecord()
            DBAccessTransactions_SetRecordEvent = False
            '    Cash BillOSale.BillOfSale, GetPrice(FinanceChargeSalesTax), "40380", "Int. Sls Tax"
        End If

        If GetPrice(txtAddlPaymentsMade.Text) > 0 Then  'Old account payments
            TransType = arPT_Prv
            Charges = 0
            Credits = GetPrice(txtAddlPaymentsMade.Text)
            Balance = (GetPrice(txtPrevBalance.Text) + GetPrice(txtGrossSale.Text) + txtDocFee.Text + GetPrice(txtLifeInsurance.Text) + GetPrice(txtAccidentInsurance.Text) + GetPrice(txtPropertyInsurance.Text) + GetPrice(txtUnemploymentInsurance.Text) + GetPrice(txtFinanceCharges.Text) + GetPrice(FinanceChargeSalesTax) - GetPrice(txtOrigDeposit.Text) - GetPrice(txtAddlPaymentsMade.Text))
            DBAccessTransactions_SetRecordEvent = True
            mDBAccessTransactions.SetRecord()
            DBAccessTransactions_SetRecordEvent = False
        End If

        If GetPrice(txtBalDueLateCharge.Text) > 0 Then  'old account latecharge balance
            TransType = arPT_PLC
            Charges = GetPrice(txtBalDueLateCharge.Text)
            Credits = 0
            Balance = (GetPrice(txtPrevBalance.Text) + GetPrice(txtGrossSale.Text) + txtDocFee.Text + GetPrice(txtLifeInsurance.Text) + GetPrice(txtAccidentInsurance.Text) + GetPrice(txtPropertyInsurance.Text) + GetPrice(txtUnemploymentInsurance.Text) + GetPrice(txtFinanceCharges.Text) + GetPrice(FinanceChargeSalesTax) - GetPrice(txtOrigDeposit.Text) - GetPrice(txtAddlPaymentsMade.Text)) + GetPrice(txtBalDueLateCharge.Text)
            DBAccessTransactions_SetRecordEvent = True
            mDBAccessTransactions.SetRecord()
            DBAccessTransactions_SetRecordEvent = False
        End If

        ' BFH20060911
        Balance = (GetPrice(txtPrevBalance.Text) + GetPrice(txtGrossSale.Text) + txtDocFee.Text + GetPrice(txtLifeInsurance.Text) + GetPrice(txtAccidentInsurance.Text) + GetPrice(txtPropertyInsurance.Text) + GetPrice(txtUnemploymentInsurance.Text) + GetPrice(txtFinanceCharges.Text) + GetPrice(FinanceChargeSalesTax) - GetPrice(txtOrigDeposit.Text) - GetPrice(txtAddlPaymentsMade.Text)) + GetPrice(txtBalDueLateCharge.Text)
        mDBAccessTransactions.dbClose()
        mDBAccessTransactions = Nothing
    End Sub

    Public Sub PrintContract(Optional ByVal JustOnePlease As Boolean = False)
        Dim BSN As String
        On Error GoTo ErrorHandler


        If IsRevolvingCharge(txtArNo.Text) Then Exit Sub ' No contract for revolving accounts MJK20140218
        OutputObject = Printer

        If IsDevelopmentMANUAL() Then GoTo JUST1FORM  'BFH20080301 - DEBUGGING

        If UseAmericanNationalInsurance Then InsuranceFormTreeHouse : Exit Sub

        If IsBoyd() Or IsUFO() Then Counter = 3 Else Counter = 2
        If JustOnePlease Then Counter = 1

        If FirstPayment = "" Then FirstPayment = dteDate2.Value

        'Print contract & Post
        For Copies = 1 To Counter
            Printer.FontName = "Arial"
            Printer.CurrentX = 0
            Printer.FontSize = 14
            Printer.CurrentY = 200
            Printer.Print("   Retail Installment & Security Agreement")

            Printer.CurrentY = 1000
            Printer.CurrentX = 0

            Printer.FontSize = 16
            Printer.FontBold = True
            Printer.CurrentX = 500
            Printer.Print(IStorename)

            Printer.FontSize = 14
            Printer.FontBold = False
            Printer.CurrentX = 500
            Printer.Print(IStoreAddress)
            Printer.CurrentX = 500
            Printer.Print(IStoreCity)
            Printer.CurrentX = 500
            Printer.Print(IStorePhone)

            Printer.CurrentX = 8000
            Printer.CurrentY = 100 '500 '1200
            Printer.FontSize = 12
            If IsEvridge Then Printer.FontSize = 11

            PrintTo(OutputObject, "Delivery Date:", 125, AlignConstants.vbAlignRight, False)
            PrintTo(OutputObject, dteDate1.Value, 142, AlignConstants.vbAlignRight, True)

            PrintTo(OutputObject, "1st. Payment Due:", 125, AlignConstants.vbAlignRight, False)
            PrintTo(OutputObject, FirstPayment, 142, AlignConstants.vbAlignRight, True)

            If IsEvridge Then
                Dim InterestPayoff As Decimal
                InterestPayoff = GetPrice(GetValueBySQL("SELECT TOP 1 [Credits] FROM Transactions WHERE ArNo='" & ArNo & "' AND [Type]='Interest Payoff' ORDER BY [TransactionID] DESC", , GetDatabaseAtLocation))

                PrintTo(OutputObject, "Prior Balance:", 125, AlignConstants.vbAlignRight, False)
                PrintTo(OutputObject, PriceFormatFunc(GetPrice(txtPrevBalance.Text) + InterestPayoff), 142, AlignConstants.vbAlignRight, True)

                PrintTo(OutputObject, "Interest Payoff:", 125, AlignConstants.vbAlignRight, False)
                PrintTo(OutputObject, PriceFormatFunc(InterestPayoff), 142, AlignConstants.vbAlignRight, True)
            End If

            PrintTo(OutputObject, "Previous Balance:", 125, AlignConstants.vbAlignRight, False)
            PrintTo(OutputObject, PriceFormatFunc(txtPrevBalance), 142, AlignConstants.vbAlignRight, True)

            PrintTo(OutputObject, "New Sale W/Tax:", 125, AlignConstants.vbAlignRight, False)
            PrintTo(OutputObject, PriceFormatFunc(txtGrossSale), 142, AlignConstants.vbAlignRight, True)

            PrintTo(OutputObject, "Total Deposit:", 125, AlignConstants.vbAlignRight, False)
            PrintTo(OutputObject, PriceFormatFunc(txtOrigDeposit), 142, AlignConstants.vbAlignRight, True)

            PrintTo(OutputObject, "Sub Total:", 125, AlignConstants.vbAlignRight, False)
            PrintTo(OutputObject, PriceFormatFunc(txtSubTotal), 142, AlignConstants.vbAlignRight, True)

            PrintTo(OutputObject, "Documentation Fee:", 125, AlignConstants.vbAlignRight, False)
            PrintTo(OutputObject, PriceFormatFunc(txtDocFee), 142, AlignConstants.vbAlignRight, True)

            PrintTo(OutputObject, "Life Insurance:", 125, AlignConstants.vbAlignRight, False)
            PrintTo(OutputObject, PriceFormatFunc(txtLifeInsurance), 142, AlignConstants.vbAlignRight, True)

            PrintTo(OutputObject, "Accident Insurance:", 125, AlignConstants.vbAlignRight, False)
            PrintTo(OutputObject, PriceFormatFunc(txtAccidentInsurance), 142, AlignConstants.vbAlignRight, True)

            PrintTo(OutputObject, "Property Insurance:", 125, AlignConstants.vbAlignRight, False)
            PrintTo(OutputObject, PriceFormatFunc(txtPropertyInsurance), 142, AlignConstants.vbAlignRight, True)

            '### UNEMPLOYMENT INSURANCE???  BFH20090722

            PrintTo(OutputObject, "Months Financed:", 125, AlignConstants.vbAlignRight, False)
            PrintTo(OutputObject, txtMonthsToFinance, 142, AlignConstants.vbAlignRight, True)

            PrintTo(OutputObject, "Amount Financed:", 125, AlignConstants.vbAlignRight, False)
            PrintTo(OutputObject, PriceFormatFunc(txtFinanceAmount), 142, AlignConstants.vbAlignRight, True)

            PrintTo(OutputObject, "Payment Deferred:", 125, AlignConstants.vbAlignRight, False)
            PrintTo(OutputObject, cboDeferred.SelectedIndex & " month(s)", 142, AlignConstants.vbAlignRight, True)

            '    printto( OutputObject, "Deferment Interest:", 125, alignconstants.vbalignright, False
            '    printto( OutputObject, PriceFormatFunc(txtDeferredInt), 142, alignconstants.vbalignright, True

            PrintTo(OutputObject, "Finance Charge:", 125, AlignConstants.vbAlignRight, False)
            PrintTo(OutputObject, PriceFormatFunc(txtFinanceCharges), 142, AlignConstants.vbAlignRight, True)

            If StoreSettings.bInstallmentInterestIsTaxable Then
                PrintTo(OutputObject, "Finance Charge Sales Tax:", 125, AlignConstants.vbAlignRight, False)
                PrintTo(OutputObject, PriceFormatFunc(FinanceChargeSalesTax), 142, AlignConstants.vbAlignRight, True)
            End If

            PrintTo(OutputObject, "Total Financed:", 125, AlignConstants.vbAlignRight, False)
            PrintTo(OutputObject, PriceFormatFunc(NewBalance + FinanceCharge + FinanceChargeSalesTax), 142, AlignConstants.vbAlignRight, True)

            Printer.FontSize = 12
            Printer.CurrentY = 3000 '3300
            Printer.CurrentX = 0

            BSN = GetBSNumList(ArNo)
            If ArMode("REPRINT") Then
                Printer.Print(TAB(10), Trim(ArCard.lblLastName.Text), ",  ", ArCard.lblFirstName.Text)
                Printer.Print(TAB(10), Trim(ArCard.lblAddress.Text))
                Printer.Print(TAB(10), Trim(ArCard.lblCity.Text), "   ", Trim(ArCard.lblZip.Text))
                Printer.Print(TAB(10), DressAni(CleanAni(ArCard.lblTele1.Text)), "  ", DressAni(CleanAni(ArCard.lblTele2.Text)))
                Printer.Print(TAB(10), "Sale No(s): ", BSN, "    Account: ", ArNo)
            ElseIf Not ArMode("E") Then  ' contract estimator
                Printer.Print(TAB(10), Trim(BillOSale.CustomerLast.Text), ",  ", BillOSale.CustomerFirst.Text)
                Printer.Print(TAB(10), Trim(BillOSale.CustomerAddress.Text))
                Printer.Print(TAB(10), Trim(BillOSale.CustomerCity.Text), "   ", Trim(BillOSale.CustomerZip.Text))
                Printer.Print(TAB(10), DressAni(CleanAni(BillOSale.CustomerPhone1.Text)), "  ", DressAni(CleanAni(BillOSale.CustomerPhone2.Text)))
                Printer.Print(TAB(10), "Sale No(s): ", BSN, "    Account: ", ArNo)
            End If

            Printer.DrawWidth = 7
            'Printer.Line(0, 4600) - Step(11375, 1300), QBColor(0), B   'set large boxed for truth & Lending
            Printer.Line(0, 4600, 11375, 1300, QBColor(0), True)   'set large boxed for truth & Lending

            'Printer.Line(0, 4600)-Step(11375, 1300), QBColor(0), B   'set large boxed for truth & Lending
            Printer.Line(0, 4600, 11375, 1300, QBColor(0), True)   'set large boxed for truth & Lending
            'Printer.Line(2500, 4600)-Step(0, 1300)
            Printer.Line(2500, 4600, 0, 1300)
            'Printer.Line(4500, 4600)-Step(0, 1300)
            Printer.Line(4500, 4600, 0, 1300)
            'Printer.Line(6800, 4600)-Step(0, 1300)
            Printer.Line(6800, 4600, 0, 1300)
            'Printer.Line(8900, 4600)-Step(0, 1300)
            Printer.Line(8900, 4600, 0, 1300)

            Printer.CurrentX = 50
            Printer.CurrentY = 4700 '5100
            Printer.FontBold = True
            Printer.FontSize = 8
            Printer.Print(" ANNUAL PERCENTAGE RATE")
            Printer.FontBold = False
            Printer.CurrentX = 50
            Printer.Print(" The cost of your credit as a")
            Printer.CurrentX = 50
            Printer.Print("yearly rate.")

            Printer.CurrentX = 2800
            Printer.CurrentY = 4700 '5100
            Printer.FontBold = True
            Printer.Print("FINANCE CHARGE")
            Printer.FontBold = False
            Printer.CurrentX = 2800
            Printer.Print(" The dollar amount the")
            Printer.CurrentX = 2800
            Printer.Print("Credit will cost you.")

            Printer.CurrentX = 4900
            Printer.CurrentY = 4700 '5100
            Printer.FontBold = True
            Printer.Print("AMOUNT FINANCED")
            Printer.FontBold = False
            Printer.CurrentX = 4600
            Printer.Print("The amount of credit pro-")
            Printer.CurrentX = 4600
            Printer.Print("vided to you on your behalf")

            Printer.CurrentX = 6900
            Printer.CurrentY = 4700 '5100
            Printer.FontBold = True
            Printer.Print(" TOTAL OF PAYMENTS")
            Printer.FontBold = False
            Printer.CurrentX = 6900
            Printer.Print(" The amount you will have")
            Printer.CurrentX = 6900
            Printer.Print("paid after you have made all")
            Printer.CurrentX = 6900
            Printer.Print("payments as scheduled.")

            Printer.CurrentX = 9200
            Printer.CurrentY = 4700 '5100
            Printer.FontBold = True
            Printer.Print("  TOTAL SALE PRICE")
            Printer.FontBold = False
            Printer.CurrentX = 9000
            Printer.Print("The total cost of your purchases")
            Printer.CurrentX = 9000
            Printer.Print("on credit; including your down-")
            Printer.CurrentX = 9000
            Printer.Print("payment of: ")

            Printer.FontBold = False
            Printer.FontSize = 8

            CalculateLastPay

            'Deposit
            Printer.CurrentX = 10000
            Printer.CurrentY = 5300 ' 5700
            Printer.Print(CurrencyFormat(txtOrigDeposit.Text)) 'deposit

            Printer.CurrentX = 800
            Printer.CurrentY = 5500 '5900

            Printer.FontSize = 14
            Printer.Print(Format(APR, "#0.00"))
            Printer.CurrentX = 3100
            Printer.Print(Format(txtFinanceCharges, "$###,##0.00"))
            Printer.CurrentX = 5200
            Printer.Print(Format(txtFinanceAmount, "$###,##0.00"))
            Printer.CurrentX = 7400
            Printer.Print(Format(NewBalance + FinanceCharge + FinanceChargeSalesTax, "$###,##0.00"))
            Printer.CurrentX = 9700
            Printer.Print(Format(NewBalance + FinanceCharge + FinanceChargeSalesTax + GetPrice(txtOrigDeposit.Text), "$###,##0.00"))

            PrintContractBody(BSN, dteDate1.Value)

            Printer.EndDoc()
        Next

        If IsElmore Then
            If GetPrice(txtLifeInsurance.Text) > 0 Or GetPrice(txtAccidentInsurance.Text) > 0 Or GetPrice(txtPropertyInsurance.Text) > 0 Or GetPrice(txtUnemploymentInsurance.Text) > 0 Then
                InsuranceForm
                InsuranceForm
            End If

        ElseIf IsLott Then ' IsMidSouth Then ' Or IsLott
JUSTFORM:
            If Not (chkLife.Checked = False And chkAccident.Checked = False And chkProperty.Checked = False And chkUnemployment.Checked = False) Then
                If LegalContractPrinter <> "" Then
                    'If MsgBox("Printing Insurance Contracts..." & vbCrLf & "This will print to: " & LegalContractPrinter & vbCrLf2 & "Would you like to use this printer?", vbQuestion + vbYesNo, "Legal Contract Printer") = vbNo Then
                    If MessageBox.Show("Printing Insurance Contracts..." & vbCrLf & "This will print to: " & LegalContractPrinter & vbCrLf2 & "Would you like to use this printer?", "Legal Contract Printer", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.No Then
                        LegalContractPrinter = ""
                    End If
                End If
                If LegalContractPrinter = "" Then
                    Dim DN As String
                    PrinterSetupDialog(Me, DN, "", False)
                    LegalContractPrinter = DN
                End If
                InsuranceForm_New(1)
                '        If IsMidSouth Then
                '          InsuranceForm_New 2
                '          InsuranceForm_New 3
                '          InsuranceForm_New 4
                '        End If
            End If

        ElseIf IsCarroll Then
            InsuranceFormCarroll()
        End If

        If IsJeffros() Or IsChicago() Or IsCarpet() Then
            WageAssignment
            WageAssignment
        End If

        Exit Sub


JUST1FORM:
        InsuranceForm_New(1)
        Exit Sub

ErrorHandler:
        CheckStandardErrors("Installment Print Contract")
        Working(False, False)
        Exit Sub
    End Sub

    Public Sub PrintCoupons()
        Dim Grace As Integer, Pages As Integer, NoPayments As Integer
        Dim Z1 As Integer, YY As Integer, Y As Integer
        Dim O As Object

        On Error GoTo ErrorHandler

        If IsRevolvingCharge(txtArNo.Text) Then Exit Sub ' No coupons for revolving accounts MJK20140218
        If ArMode("E") Then NoMonths = Val(txtMonthsToFinance.Text)
        O = Printer
        Pages = IIf(optWeekly.Checked = True, NoMonths, Trunc(NoMonths / 4 + 0.9, 0)) ' math trick ==> trunc(3.0 + .9) = 3, trunc(3.25 + .9) = 4, trunc(3.75 + .9) = 4... it's what we want
        NoPayments = IIf(optWeekly.Checked = True, NoMonths * 4, NoMonths)
        Counter = 1 - Pages
        Grace = AdjustedGracePeriod(DueOn)

        For YY = 1 To Pages
            For Z1 = 1 To 4  'no per page
                Counter = ((Z1 - 1) * Pages + 1) + (YY - 1) ' + (YY - 1) * 4 + (Z1 - 1) ' 0 to n-1
                If Counter > NoPayments Then Exit For        ' for monthly, non-multiple of 4
                Y = Choose(Z1, 800, 4800, 8800, 12400)

                O.FontSize = 8
                O.FontBold = False
                O.CurrentY = 0
                O.CurrentX = 0

                O.DrawWidth = 4

                'left stub
                O.CurrentY = Y
                O.CurrentX = 500
                O.print("Payment No: ", Counter, " of ", NoPayments)
                O.CurrentX = 500

                'calculate first payment
                '              If ArMode("E") Then
                '                FirstPayment = DateAdd("m", IIf(DueOn = 1, 1, 0), Format(dteDate2, "mm/" & DueOn & "/yyyy"))
                '              Else
                FirstPayment = dteDate2.Value
                '              End If
                If Not IsDate(FirstPayment) Then MessageBox.Show("Could not get First Payment date for Coupons.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning) : Exit Sub
                If optWeekly.Checked = True Then
                    DateDue = DateAdd("ww", Counter - 1, FirstPayment)
                Else
                    DateDue = DateAdd("m", Counter - 1, FirstPayment)
                End If

                O.print("   Due Date: ", DateDue)
                O.CurrentX = 500
                O.print("       Amount: ", FormatCurrency(GetPrice(IIf(Counter <> NoMonths, Payment, LastPay))))  ' check last payment

                O.print
                O.CurrentX = 200
                PrintToPosition(O, "Amount Paid:  $______________", 2600, AlignConstants.vbAlignRight, True)
                O.print
                PrintToPosition(O, "Date Paid:  $______________", 2600, AlignConstants.vbAlignRight, True)
                O.print
                PrintToPosition(O, "Check No:  $______________", 2600, AlignConstants.vbAlignRight, True)

                'lines seperating coupons
                O.Line(3000, Choose(Z1, 700, 4500, 8500, 12500), 3000, Choose(Z1, 3200, 7000, 11000, 15000))

                'body of coupon

                O.FontSize = 10
                O.CurrentY = Y - IIf(Z1 <> 4, 800, 300)

                O.CurrentX = 3200
                O.FontBold = True
                O.print("Your payment must be received by Due Date!")
                O.FontBold = False

                O.print
                O.print

                O.CurrentY = Y - IIf(Z1 <> 4, 500, 0)
                O.CurrentX = 8000
                O.print("   Payment No: ", Counter, " of ", NoPayments)
                O.print

                O.FontBold = True

                PrintToPosition(O, "Due On:", 7100, AlignConstants.vbAlignRight)
                PrintToPosition(O, DateDue, 8300, AlignConstants.vbAlignRight)

                PrintToPosition(O, "Late On:", 10000, AlignConstants.vbAlignRight)
                PrintToPosition(O, DateAdd("D", Grace, DateDue), 11100, AlignConstants.vbAlignRight, True)

                PrintToPosition(O, "Amount:", 7100, AlignConstants.vbAlignRight)
                PrintToPosition(O, Format(IIf(Counter <> Val(NoMonths), txtPaymentWillBe, LastPay), "$###,##0.00"), 8300, AlignConstants.vbAlignRight)
                PrintToPosition(O, "Late Charge:", 10000, AlignConstants.vbAlignRight)
                PrintToPosition(O, FormatCurrency(LateCharge), 11100, AlignConstants.vbAlignRight, True)

                PrintToPosition(O, "Amount:", 10000, AlignConstants.vbAlignRight)
                PrintToPosition(O, FormatCurrency(IIf(Counter <> Val(NoMonths), txtPaymentWillBe, LastPay) + LateCharge), 11100, AlignConstants.vbAlignRight, True)

                O.FontBold = False
                O.print

                If Z1 = 4 Then O.CurrentY = O.CurrentY - 100  'for Arno to fit

                O.CurrentX = 7990
                PrintToPosition(O, "Amount Paid: $_____________", 11100, AlignConstants.vbAlignRight, True)
                O.CurrentX = 4000

                O.print("    Mail To: ")
                O.CurrentX = 4000
                O.print(IStorename)
                O.CurrentX = 8000
                O.print(Trim(BillOSale.CustomerFirst.Text & " " & BillOSale.CustomerLast.Text))

                O.CurrentX = 4000
                O.print(IStoreAddress)
                O.CurrentX = 8000
                O.print(BillOSale.CustomerAddress.Text)

                If Trim(BillOSale.AddAddress.Text) <> "" Then
                    O.CurrentX = 8000
                    O.print(BillOSale.AddAddress.Text)
                End If

                O.CurrentX = 4000
                O.print(IStoreCity)
                O.CurrentX = 8000
                O.print(BillOSale.CustomerCity.Text & "  " & BillOSale.CustomerZip.Text)

                O.CurrentX = 8000
                O.print("Account Number: ", ArNo)
                O.CurrentY = Choose(Z1, 3600, 7600, 11600, 0)

                If Z1 <> 4 Then
                    O.print("-----------------------------------------------------------------------------------------------------------------------------------------------------------------------")
                End If

            Next
            O.NewPage
        Next
        O.EndDoc

        Exit Sub

ErrorHandler:
        CheckStandardErrors("Installment Print Coupons")
        Working(False, False)
        Exit Sub

    End Sub

    Private Sub InsuranceFormTreeHouse()
        Dim Op As Object, R As VbMsgBoxResult
        Dim PName As String, I As Integer
        Dim Page1Copies As Integer, Page2Copies As Integer, Page3Copies As Integer, Page4Copies As Integer
        Dim X_SCALE As Double, Y_SCALE As Double
        Dim First As String, Last As String, Add As String, City As String, Zip As String, Sales1 As String
        Dim SS As String, DOB As Date, hAge As String, HasCo As Boolean

        Op = Printer.DeviceName
        FirstPayment = DateAdd("m", 1, Format(dteDate1.Value, "mm/" & DueOn & "/yyyy"))
        FirstPayment = fAdjustFirstPay(FirstPayment, cboDeferred.SelectedIndex) ' BFH20061026 - Added b/c it was wrong.. contracts and coupons were broken
        If DateDiff("d", dteDate2, FirstPayment) < 0 Then
            FirstPayment = DateAdd("m", 1, FirstPayment)
        End If

        If chkLife.Checked = False And chkProperty.Checked = False And chkAccident.Checked = False And chkUnemployment.Checked = False Then
            Page1Copies = 0
            Page2Copies = 0
            Page3Copies = 2
            Page4Copies = 2
        Else
            Page1Copies = 3
            Page2Copies = 1  ' refund
            Page3Copies = 3
            Page4Copies = 3
        End If

        ' Auto Dell 2335dn MFP on BLUESKY2
        ' Brother HL-5370DW series  Treehouse
        ' Brother MFC-8640D USB'  my printer


        If False Then
            '
            '  ElseIf IsBlueSky Then
            '    PName = "Auto Dell 2335dn MFP on BLUESKY2"
        ElseIf IsTreehouse Then
            PName = "Brother HL-5370DW series"
        Else
            PName = "Brother MFC-8640D USB" ' me
        End If

        'If MsgBox("Please make sure your " & PName & " is ready to print on legal sized paper.", vbExclamation + vbOKCancel) = vbCancel Then Exit Sub
        If MessageBox.Show("Please make sure your " & PName & " is ready to print on legal sized paper.", "", MessageBoxButtons.OKCancel, MessageBoxIcon.Exclamation) = DialogResult.Cancel Then Exit Sub
        If True Or Not IsDevelopment() Then
            If Not SetPrinter(PName) Then
                If True Or Not IsDevelopment() Then
                    MessageBox.Show("Could not connect to " & PName & ".", "Printer Selection Failed", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                    '        Exit Sub
                End If
            End If
        End If

        '  vbPRBNUpper         - Use paper from the upper bin.
        '  vbPRBNLower         - Use paper from the lower bin.
        '  vbPRBNMiddle        - Use paper from the middle bin.
        '  vbPRBNManual        - Wait for manual insertion of each sheet of paper.
        '  vbPRBNEnvelope      - Use envelopes from the envelope feeder.
        '  vbPRBNEnvManual     - Use envelopes from the envelope feeder, but wait for manual insertion.
        '  vbPRBNAuto          - Use paper from the current default bin. (This is the default.)
        '  vbPRBNTractor       - Use paper fed from the tractor feeder.
        '  vbPRBNSmallFmt      - Use paper from the small paper feeder.
        '  vbPRBNLargeFmt      - Use paper from the large paper bin.
        '  vbPRBNLargeCapacity - Use paper from the large capacity feeder.
        '  vbPRBNCassette      - Use paper from the attached cassette cartridge.
        On Error Resume Next
        ' bluesky has 2 trays... tray #2 is legal sized paper..
        Printer.EndDoc()  ' jic

        '  If IsBlueSky Then Printer.PaperBin = vbPRBNLower
        If IsTreehouse Then Printer.PaperBin = vbPRBNLower 'vbPRBNLower

        Printer.FontName = "Arial"
        Printer.FontSize = 8
        'GoTo HERE
        For I = 1 To Page1Copies

            Printer.PaperSize = vbPRPSLegal
            X_SCALE = Printer.ScaleWidth / 12240
            Y_SCALE = Printer.ScaleHeight / 15840

            'picPicture.Picture = LoadPictureStd(FXFile("FNI-Blank-1.gif"))
            picPicture.Image = LoadPictureStd(FXFile("FNI-Blank-1.gif"))
            Printer.PaintPicture(picPicture.Image, 0, 0, Printer.ScaleWidth, Printer.ScaleHeight)

            Dim MailIndex As Integer
            'Dim First As String, Last As String, Add As String, City As String, Zip As String, Sales1 As String
            Dim SQL As String, RS As ADODB.Recordset
            'Dim SS As String, DOB As Date, hAge As String, HasCo As Boolean
            Dim CoName As String, CoAddress As String, CoCity As String, CoSS As String, CoDOB As Date, CoAge As String
            Dim expireDate As Date

            If IsFormLoaded("BillOSale") Then
                MailIndex = BillOSale.MailIndex
                First = BillOSale.CustomerFirst.Text
                Last = BillOSale.CustomerLast.Text
                Add = BillOSale.CustomerAddress.Text
                City = BillOSale.CustomerCity.Text
                Zip = BillOSale.CustomerZip.Text
                Sales1 = BillOSale.Sales1.Text
            Else
                MailIndex = ReprintMailIndex
                Dim M As MailNew
                GetMailNewByIndex(MailIndex, M, StoresSld)
                First = M.First
                Last = M.Last
                Add = M.Address
                City = M.City
                Zip = M.Zip
                Sales1 = ""
            End If

            If MailIndex <> 0 Then
                SQL = "SELECT DOB, HisSS, CoName, CoAddress, CoCityState, CoSS, CoDOB FROM [ArApp] WHERE MailIndex='" & MailIndex & "'"
                RS = GetRecordsetBySQL(SQL, , GetDatabaseAtLocation())
                If RS.RecordCount >= 1 Then
                    DOB = IfNullThenZeroDate(RS("DOB").Value)
                    If IsDate(DOB) Then hAge = Age(DOB) Else hAge = ""
                    SS = Trim(IfNullThenNilString(RS("HisSS").Value))
                    CoName = Trim(IfNullThenNilString(RS("CoName").Value))
                    CoAddress = Trim(IfNullThenNilString(RS("CoAddress").Value))
                    CoCity = Trim(IfNullThenNilString(RS("CoCityState").Value))
                    CoSS = Trim(IfNullThenNilString(RS("CoSS").Value))
                    CoDOB = IfNullThenZeroDate(RS("CoDOB").Value)
                    If IsDate(CoDOB) Then CoAge = Age(CoDOB) Else CoAge = ""
                    HasCo = CoName <> "" Or CoSS <> ""
                End If
                RS = Nothing
            End If

            expireDate = DateAdd("m", Val(txtMonthsToFinance), dteDate2)

            Printer.FontName = "Arial"
            Printer.FontSize = 8

            Printer.FontBold = True

            Printer.CurrentX = 1300 * X_SCALE
            Printer.CurrentY = 1200 * Y_SCALE
            Printer.Print(txtArNo.Text)

            ' PRIMARY

            Printer.CurrentX = 1000 * X_SCALE
            Printer.CurrentY = 1930 * Y_SCALE
            If InStr(First, "&") Then
                If CoName = SplitWord(First, 2, "&") & " " & Last Then
                    Printer.Print(SplitWord(First, 1, "&") & " " & Last)
                ElseIf CoName = SplitWord(First, 1, "&") & " " & Last Then
                    Printer.Print(SplitWord(First, 2, "&") & " " & Last)
                Else
                    Printer.Print(First & " " & Last)
                End If
            Else
                Printer.Print(First & " " & Last)
            End If


            Printer.CurrentX = 5225 * X_SCALE
            Printer.CurrentY = 1930 * Y_SCALE
            Printer.Print(hAge)

            Printer.CurrentX = 6050 * X_SCALE
            Printer.CurrentY = 1930 * Y_SCALE
            Printer.Print(SS)

            Printer.CurrentX = 7600 * X_SCALE
            Printer.CurrentY = 1930 * Y_SCALE
            Printer.Print(Add & ";", City & " " & Zip)

            ' CO-BUYER

            Printer.CurrentX = 1000 * X_SCALE
            Printer.CurrentY = 2700 * Y_SCALE
            Printer.Print(CoName)

            Printer.CurrentX = 5225 * X_SCALE
            Printer.CurrentY = 2700 * Y_SCALE
            Printer.Print(CoAge)

            Printer.CurrentX = 6050 * X_SCALE
            Printer.CurrentY = 2700 * Y_SCALE
            Printer.Print(CoSS)

            Printer.CurrentX = 7600 * X_SCALE
            Printer.CurrentY = 2700 * Y_SCALE
            Printer.Print(CoAddress & ";", CoCity)

            ' STORE

            Printer.CurrentX = 2100 * X_SCALE
            Printer.CurrentY = 3550 * Y_SCALE
            Printer.Print(UCase(StoreSettings.Name))

            Printer.CurrentX = 5500 * X_SCALE
            Printer.CurrentY = 3370 * Y_SCALE
            Printer.Print(UCase(StoreSettings.Address))

            Printer.CurrentX = 5500 * X_SCALE
            Printer.CurrentY = 3550 * Y_SCALE
            Printer.Print(UCase(StoreSettings.City))

            Printer.CurrentX = 8000 * X_SCALE
            Printer.CurrentY = 3550 * Y_SCALE
            Printer.Print(UCase("estate"))

            ' loan information

            Printer.CurrentX = 4800 * X_SCALE
            Printer.CurrentY = 4150 * Y_SCALE
            Printer.Print(lblAPR.Text)

            Printer.CurrentX = 7500 * X_SCALE
            Printer.CurrentY = 4150 * Y_SCALE
            Printer.Print(txtFinanceAmount.Text)

            Printer.CurrentX = 11000 * X_SCALE
            Printer.CurrentY = 4150 * Y_SCALE
            Printer.Print(txtMonthsToFinance.Text)

            ' GROSS DECREASING
            Printer.CurrentX = IIf(optJointLife1.Checked, 3560, 2880) * X_SCALE
            Printer.CurrentY = 5260 * Y_SCALE
            Printer.Print("X")

            Printer.CurrentX = 5000 * X_SCALE
            Printer.CurrentY = 5300 * Y_SCALE
            Printer.Print(dteDate1.Value)

            Printer.CurrentX = 6500 * X_SCALE
            Printer.CurrentY = 5300 * Y_SCALE
            Printer.Print(txtMonthsToFinance.Text)

            Printer.CurrentX = 7450 * X_SCALE
            Printer.CurrentY = 5300 * Y_SCALE
            Printer.Print(expireDate)

            Printer.CurrentX = 8700 * X_SCALE
            Printer.CurrentY = 5300 * Y_SCALE
            Printer.Print(CurrencyFormat(GetPrice(Balance)))

            Printer.CurrentX = 10900 * X_SCALE
            Printer.CurrentY = 5300 * Y_SCALE
            Printer.Print(txtLifeInsurance.Text)

            ' ACCIDENT
            Printer.CurrentX = 1000 * X_SCALE
            Printer.CurrentY = 6750 * Y_SCALE
            Printer.Print("1st")

            Printer.CurrentX = 2500 * X_SCALE
            Printer.CurrentY = 6950 * Y_SCALE
            Printer.Print("14")

            '2880
            '3560
            Printer.CurrentX = 2880 * X_SCALE
            Printer.CurrentY = 7450 * Y_SCALE
            Printer.Print("X")

            Printer.CurrentX = 5000 * X_SCALE
            Printer.CurrentY = 6600 * Y_SCALE
            Printer.Print(dteDate1.Value)

            Printer.CurrentX = 6500 * X_SCALE
            Printer.CurrentY = 6600 * Y_SCALE
            Printer.Print(txtMonthsToFinance.Text)

            Printer.CurrentX = 7450 * X_SCALE
            Printer.CurrentY = 6600 * Y_SCALE
            Printer.Print(expireDate)

            Printer.CurrentX = 9900 * X_SCALE
            Printer.CurrentY = 6600 * Y_SCALE
            Printer.Print(Math.Round(Payment, 2))

            Printer.CurrentX = 10900 * X_SCALE
            Printer.CurrentY = 6600 * Y_SCALE
            Printer.Print(txtAccidentInsurance.Text)

            ' PROPERTY
            Printer.CurrentX = 5000 * X_SCALE
            Printer.CurrentY = 8150 * Y_SCALE
            Printer.Print(dteDate1.Value)

            Printer.CurrentX = 6500 * X_SCALE
            Printer.CurrentY = 8150 * Y_SCALE
            Printer.Print(txtMonthsToFinance.Text)

            Printer.CurrentX = 7450 * X_SCALE
            Printer.CurrentY = 8150 * Y_SCALE
            Printer.Print(expireDate)

            Printer.CurrentX = 8700 * X_SCALE
            Printer.CurrentY = 8150 * Y_SCALE
            Printer.Print(CurrencyFormat(GetPrice(Balance)))

            Printer.CurrentX = 10900 * X_SCALE
            Printer.CurrentY = 8150 * Y_SCALE
            Printer.Print(txtPropertyInsurance.Text)

            'IUI
            '2880
            '3560
            Printer.CurrentX = 2880 * X_SCALE
            Printer.CurrentY = 9250 * Y_SCALE
            Printer.Print("X")

            Printer.CurrentX = 5000 * X_SCALE
            Printer.CurrentY = 9100 * Y_SCALE
            Printer.Print(dteDate1.Value)

            Printer.CurrentX = 6500 * X_SCALE
            Printer.CurrentY = 9100 * Y_SCALE
            Printer.Print(txtMonthsToFinance.Text)

            Printer.CurrentX = 7450 * X_SCALE
            Printer.CurrentY = 9100 * Y_SCALE
            Printer.Print(expireDate)

            Printer.CurrentX = 9900 * X_SCALE
            Printer.CurrentY = 9100 * Y_SCALE
            Printer.Print(Math.Round(Payment, 2))

            Printer.CurrentX = 10900 * X_SCALE
            Printer.CurrentY = 9100 * Y_SCALE
            Printer.Print(txtUnemploymentInsurance.Text)

            ' OTHER

            Printer.CurrentX = 10900 * X_SCALE
            Printer.CurrentY = 9700 * Y_SCALE
            Printer.Print(CurrencyFormat(GetPrice(txtLifeInsurance.Text) + GetPrice(txtAccidentInsurance.Text) + GetPrice(txtPropertyInsurance.Text) + GetPrice(txtUnemploymentInsurance.Text)))

            Printer.CurrentX = 3560 * X_SCALE
            Printer.CurrentY = 10200 * Y_SCALE
            Printer.Print("HOUSEHOLD GOODS")

            '
            '  If I <= Page2Copies Then
            '    Printer.NewPage
            '    Printer.PaperSize = vbPRPSLegal
            '    picPicture.Picture = LoadPicture(fxfile( "FNI-Blank-2.gif"))
            '    Printer.PaintPicture picPicture.Picture, 0, 0, Printer.ScaleWidth, Printer.ScaleHeight
            '  End If

            Printer.NewPage()
        Next

        For I = 1 To Page2Copies

            '    If IsBlueSky Then Printer.PaperBin = vbPRBNLower
            If IsTreehouse Then Printer.PaperBin = vbPRBNLower 'vbPRBNLower

            Printer.PaperSize = vbPRPSLegal
            picPicture.Image = LoadPictureStd(FXFile("FNI-Blank-2.gif"))
            Printer.PaintPicture(picPicture.Image, 0, 0, Printer.ScaleWidth, Printer.ScaleHeight)
            Printer.NewPage()
        Next

        Printer.EndDoc()


HERE:

        For I = 1 To Page3Copies
            Printer.Duplex = vbPRDPHorizontal 'vbPRDPSimplex 'vbPRDPHorizontal    ' do it if it's supported..
            '  If IsBlueSky Then Printer.PaperBin = vbPRBNLower
            If IsTreehouse Then Printer.PaperBin = vbPRBNLower 'vbPRBNLower
            Printer.PaperSize = vbPRPSLegal

            picPicture.Image = LoadPictureStd(FXFile("FNI-Burrell-1.gif"))
            Printer.PaintPicture(picPicture.Image, 0, 0, Printer.ScaleWidth, Printer.ScaleHeight)
            Printer.FontName = "Arial"
            Printer.FontSize = 8

            PrintOut_Alt2(OutObj:=Printer, X:=10000, Y:=400, Text:=dteDate1.Value) ' Date

            PrintOut_Alt2(OutObj:=Printer, X:=4100, Y:=800, Text:=StoreSettings.Name)
            PrintOut_Alt2(OutObj:=Printer, X:=4100, Y:=950, Text:=StoreSettings.Address)
            PrintOut_Alt2(OutObj:=Printer, X:=4100, Y:=1100, Text:=StoreSettings.City)
            PrintOut_Alt2(OutObj:=Printer, X:=4100, Y:=1250, Text:=StoreSettings.Phone)

            PrintOut_Alt2(OutObj:=Printer, X:=2600, Y:=2125, Text:=Math.Round(APR, 2))
            PrintOut_Alt2(OutObj:=Printer, X:=3750, Y:=2125, Text:=CurrencyFormat(txtFinanceCharges.Text))
            PrintOut_Alt2(OutObj:=Printer, X:=5500, Y:=2125, Text:=CurrencyFormat(txtFinanceAmount.Text))
            PrintOut_Alt2(OutObj:=Printer, X:=7300, Y:=2125, Text:=CurrencyFormat(txtTotalBalance.Text))
            PrintOut_Alt2(OutObj:=Printer, X:=10200, Y:=1900, FontSize:=6, Text:=CurrencyFormat(txtOrigDeposit.Text))
            PrintOut_Alt2(OutObj:=Printer, X:=9100, Y:=2125, Text:=CurrencyFormat(txtGrossSale.Text))

            PrintOut_Alt2(OutObj:=Printer, X:=2800, Y:=3000, Text:=txtMonthsToFinance.Text)
            PrintOut_Alt2(OutObj:=Printer, X:=3800, Y:=3000, Text:=Math.Round(Payment, 2))
            PrintOut_Alt2(OutObj:=Printer, X:=5000, Y:=3000, Text:=Switch(optLate6, "1st of every month", optLate16, "10th of every month", True, "20th of every month"))

            PrintOut_Alt2(OutObj:=Printer, X:=3000, Y:=4500, Text:=First & " " & Last)
            PrintOut_Alt2(OutObj:=Printer, X:=3000, Y:=4650, Text:=Add)
            PrintOut_Alt2(OutObj:=Printer, X:=3000, Y:=4800, Text:=City & " " & Zip)

            PrintOut_Alt2(OutObj:=Printer, X:=2000, Y:=5250, Text:=Sales1)
            PrintOut_Alt2(OutObj:=Printer, X:=2700, Y:=5250, Text:="Yes")
            PrintOut_Alt2(OutObj:=Printer, X:=3500, Y:=5250, Text:="No")
            PrintOut_Alt2(OutObj:=Printer, X:=4400, Y:=5250, Text:=YesNo(Not chkAutoARNO.Visible))
            '  PrintOut_alt2 OutObj:=Printer, X:=5200, Y:=5250, Text:="NO MEMO"
            '  PrintOut_alt2 OutObj:=Printer, X:=8450, Y:=5250, Text:="NO OTHER"

            Printer.FontSize = 6

            Dim J As Integer, gY As Integer
            gY = 5600

            Dim CS As sSale
            CS = New sSale

            If IsFormLoaded("BillOSale") Then
                CS.LoadFromBillOSale()
            Else
                CS.LoadSaleNo(ReprintSaleNo)
            End If

            For J = 0 To CS.ItemCount - 1
                'BFH20090910 - via Jerry's email of 9/8/2009...
                ' On the attached contract, the sale part of the contract must add up to the total plus tax.
                ' Therefore, it must have Notes, Kits, Delivery, Labor, Stain Protection etc in that section of the contract.
                '    If IsItem(BillOSale.QueryStyle(J)) Then
                Dim Xa As String, Xy As String
                Xa = CS.Item(J).Status
                Xy = CS.Item(J).Style
                If IsVoid(Xa) Or IsReturned(Xa) Then GoTo Skip
                If IsSub(Xy) Or IsADJ(Xy) Or IsPayment(Xy) Then GoTo Skip
                If IsNote(Xy) And GetPrice(CS.Item(J).Price) = 0 Then GoTo Skip
                If IsTax(Xy) Then GoTo Skip

                PrintOut_Alt2(OutObj:=Printer, X:=2000, Y:=gY, Text:=CS.Item(J).Quantity)
                PrintOut_Alt2(OutObj:=Printer, X:=2700, Y:=gY, Text:=CS.Item(J).VendorNo)
                PrintOut_Alt2(OutObj:=Printer, X:=3400, Y:=gY, Text:=CS.Item(J).Desc)
                PrintOut_Alt2(OutObj:=Printer, X:=8450, Y:=gY, Text:=CS.Item(J).Style)
                PrintOut_Alt2(OutObj:=Printer, X:=9850, Y:=gY, Text:=CurrencyFormat(CS.Item(J).Price))
                gY = gY + 100
Skip:
            Next

            PrintOut_Alt2(OutObj:=Printer, X:=9850, Y:=8200, Text:=CurrencyFormat(CS.SubTotal("TAX1") + CS.SubTotal("TAX2")))
            PrintOut_Alt2(OutObj:=Printer, X:=9850, Y:=8375, Text:=txtGrossSale.Text)

            Printer.FontSize = 8

            If Not ArMode("REPRINT") Then
                ' left summary
                If GetPrice(txtPrevBalance.Text) <> 0 Then
                    PrintOut_Alt2(OutObj:=Printer, X:=4100, Y:=8370, Text:=CurrencyFormat(txtPrevBalance.Text))
                    '    If IsFormLoaded("ArCard") Then
                    PrintOut_Alt2(OutObj:=Printer, X:=4100, Y:=8560, Text:=CurrencyFormat(ArCard.InterestCredit))
                    PrintOut_Alt2(OutObj:=Printer, X:=4100, Y:=8750, Text:=CurrencyFormat(ArCard.LifeCredit))
                    PrintOut_Alt2(OutObj:=Printer, X:=4100, Y:=8940, Text:=CurrencyFormat(ArCard.AccidentCredit))
                    PrintOut_Alt2(OutObj:=Printer, X:=4100, Y:=9150, Text:=CurrencyFormat(ArCard.PropertyCredit))
                    PrintOut_Alt2(OutObj:=Printer, X:=4100, Y:=9330, Text:=CurrencyFormat(ArCard.IUICredit))
                    PrintOut_Alt2(OutObj:=Printer, X:=4100, Y:=9520, Text:=CurrencyFormat(GetPrice(txtPrevBalance.Text) + ArCard.InterestCredit + ArCard.LifeCredit + ArCard.AccidentCredit + ArCard.PropertyCredit + ArCard.IUICredit))
                    '    End If
                End If
            Else
                PrintOut_Alt2(OutObj:=Printer, X:=4100, Y:=8370, Text:="N/A")
                PrintOut_Alt2(OutObj:=Printer, X:=4100, Y:=9520, Text:="N/A")
            End If

            ' left, prop summary
            If chkProperty.Checked = True Then
                PrintOut_Alt2(OutObj:=Printer, X:=1950, Y:=9900, Text:="X")
                PrintOut_Alt2(OutObj:=Printer, X:=6200, Y:=10575, Text:=txtPropertyInsurance.Text)
            End If
            If False Then
                PrintOut_Alt2(OutObj:=Printer, X:=1950, Y:=10820, Text:="X")
            End If

            ' left, other summary
            If optJointLife0.Checked = True Then
                PrintOut_Alt2(OutObj:=Printer, X:=3850, Y:=11650, Text:=txtLifeInsurance.Text)      ' single term premium
            Else
                PrintOut_Alt2(OutObj:=Printer, X:=6200, Y:=11650, Text:=txtLifeInsurance.Text)      ' joint term premium
            End If
            PrintOut_Alt2(OutObj:=Printer, X:=3850, Y:=11850, Text:=txtAccidentInsurance.Text)
            PrintOut_Alt2(OutObj:=Printer, X:=3850, Y:=12050, Text:=txtUnemploymentInsurance.Text)

            If chkLife.Checked = True Then PrintOut_Alt2(OutObj:=Printer, X:=1950, Y:=12230, Text:="X")
            If chkAccident.Checked = True Then PrintOut_Alt2(OutObj:=Printer, X:=4075, Y:=12230, Text:="X")
            If chkUnemployment.Checked = True Then PrintOut_Alt2(OutObj:=Printer, X:=5690, Y:=12230, Text:="X")
            If chkLife.Checked = True And chkAccident.Checked = False And chkUnemployment.Checked = False Then PrintOut_Alt2(OutObj:=Printer, X:=1950, Y:=12430, Text:="X")

            PrintOut_Alt2(OutObj:=Printer, X:=2100, Y:=12650, Text:=hAge)

            If chkLife.Checked = True Then
                PrintOut_Alt2(OutObj:=Printer, X:=1925, Y:=12900, Text:="X")
            Else
                PrintOut_Alt2(OutObj:=Printer, X:=4350, Y:=12900, Text:="X")
            End If

            PrintOut_Alt2(OutObj:=Printer, X:=2100, Y:=13090, Text:=hAge)

            ' right summary
            PrintOut_Alt2(OutObj:=Printer, X:=9750, Y:=8760, Text:=CurrencyFormat(txtGrossSale.Text))
            PrintOut_Alt2(OutObj:=Printer, X:=9200, Y:=8950, Text:=CurrencyFormat(txtOrigDeposit.Text))
            'PrintOut_alt2 OutObj:=Printer, X:=9200, Y:=9160, Text:="0.00" ' trade in
            PrintOut_Alt2(OutObj:=Printer, X:=9750, Y:=9360, Text:=CurrencyFormat(txtOrigDeposit.Text))
            PrintOut_Alt2(OutObj:=Printer, X:=9750, Y:=9550, Text:=CurrencyFormat(GetPrice(txtGrossSale.Text) - GetPrice(txtOrigDeposit.Text)))

            PrintOut_Alt2(OutObj:=Printer, X:=9750, Y:=9740, Text:=CurrencyFormat(txtPrevBalance.Text))
            PrintOut_Alt2(OutObj:=Printer, X:=9750, Y:=9930, Text:=CurrencyFormat(txtSubTotal.Text))

            PrintOut_Alt2(OutObj:=Printer, X:=9200, Y:=10320, Text:=CurrencyFormat(txtLifeInsurance.Text))
            PrintOut_Alt2(OutObj:=Printer, X:=9200, Y:=10510, Text:=CurrencyFormat(txtAccidentInsurance.Text))
            PrintOut_Alt2(OutObj:=Printer, X:=9200, Y:=10700, Text:=CurrencyFormat(txtPropertyInsurance.Text))
            PrintOut_Alt2(OutObj:=Printer, X:=9200, Y:=10890, Text:=CurrencyFormat(txtUnemploymentInsurance.Text))
            PrintOut_Alt2(OutObj:=Printer, X:=9200, Y:=11080, Text:="0.00")
            PrintOut_Alt2(OutObj:=Printer, X:=9200, Y:=11270, Text:=CurrencyFormat(txtDocFee))

            PrintOut_Alt2(OutObj:=Printer, X:=9750, Y:=11480, Text:=CurrencyFormat(GetPrice(txtLifeInsurance.Text) + GetPrice(txtAccidentInsurance.Text) + GetPrice(txtPropertyInsurance.Text) + GetPrice(txtUnemploymentInsurance.Text) + GetPrice(txtDocFee.Text)))
            PrintOut_Alt2(OutObj:=Printer, X:=9750, Y:=11670, Text:=CurrencyFormat(txtFinanceAmount.Text))
            PrintOut_Alt2(OutObj:=Printer, X:=9750, Y:=11870, Text:=CurrencyFormat(txtFinanceCharges.Text))
            PrintOut_Alt2(OutObj:=Printer, X:=9750, Y:=12070, Text:=CurrencyFormat(txtTotalBalance.Text))

            ' payment schedule

            Printer.FontSize = 6
            PrintOut_Alt2(OutObj:=Printer, X:=10325, Y:=13800, Text:=txtMonthsToFinance.Text & " " & IIf(StoreSettings.bPaymentBooksMonthly, "monthly", "weekly"))
            PrintOut_Alt2(OutObj:=Printer, X:=9750, Y:=14000, Text:=Math.Round(Payment, 2))
            PrintOut_Alt2(OutObj:=Printer, X:=8400, Y:=14200, Text:=LastPay)
            PrintOut_Alt2(OutObj:=Printer, X:=7450, Y:=14400, Text:=FirstPayment) ' dteDate2
            Printer.FontSize = 8

            If I <= Page4Copies Then
                Printer.NewPage()
                '    If IsBlueSky Then Printer.PaperBin = vbPRBNLower
                If IsTreehouse Then Printer.PaperBin = vbPRBNLower 'vbPRBNLower
                Printer.PaperSize = vbPRPSLegal
                picPicture.Image = LoadPictureStd(FXFile("FNI-Burrell-2.gif"))
                Printer.PaintPicture(picPicture.Image, 0, 0, Printer.ScaleWidth, Printer.ScaleHeight)
            End If

            Printer.NewPage()
        Next
        Printer.EndDoc()


        On Error Resume Next
        If IsTreehouse Or IsBlueSky Then Printer.PaperBin = vbPRBNAuto 'vbPRBNUpper
        SetPrinter(Op)
    End Sub

    Private Sub InsuranceFormCarroll() 'Carroll
        Dim Op As Object, R As VbMsgBoxResult
        Op = Printer.DeviceName

#If False Then
  picPicture.Picture = LoadPictureStd(FXFile("Insurance Contract Carroll.gif"))
  Printer.PaintPicture picPicture.Picture, 0, 0, Printer.ScaleWidth, Printer.ScaleHeight
#Else
        '  R = MsgBox("Be sure contract form is ready in tractor printer.", vbExclamation + vbOKCancel, "Printing contract")
        '  If R = vbCancel Then Exit Sub
        If Not PrinterSetupDialog(Me, "Lexmark 2380 Plus", "") Then Exit Sub
#End If

        Printer.FontName = "Arial"
        Printer.FontSize = 12

        Printer.CurrentX = 6000
        Printer.CurrentY = 2075 '1175
        Printer.Print(StoreSettings.Name)

        Printer.CurrentX = 10050
        Printer.Print(ArNo) 'Creditor's No.

        Printer.FontSize = 8
        Printer.CurrentX = 200
        Printer.CurrentY = 2575
        Printer.Print(Trim(BillOSale.CustomerFirst.Text), "  ", Trim(BillOSale.CustomerLast.Text))
        Printer.CurrentX = 200
        Printer.Print(Trim(BillOSale.CustomerAddress.Text), IIf(Len(Trim(BillOSale.CustomerAddress.Text)) > 0, ",  ", ""), Trim(BillOSale.CustomerCity.Text), " ", Trim(BillOSale.CustomerZip.Text))

        Printer.CurrentX = 200
        Printer.CurrentY = 3300
        Printer.FontSize = 12
        Printer.Print(dteDate1.Value) 'Date

        Printer.CurrentX = 4000
        '.CurrentY = 2880
        Printer.Print(txtMonthsToFinance.Text) 'Term (Months)

        Printer.CurrentX = 3050
        Printer.CurrentY = 3700
        Printer.FontSize = 12
        Printer.Print("X") '#1 Joint Life Net Decreasing Term

        Printer.CurrentY = 4280
        Printer.CurrentX = 7750

        Printer.Print(CurrencyFormat(txtFinanceAmount.Text))
        Printer.CurrentX = 10050
        Printer.Print(CurrencyFormat(txtLifeInsurance.Text))

        Printer.CurrentY = 5200
        Printer.CurrentX = 500
        Printer.Print("7")
        Printer.CurrentY = 5650
        Printer.CurrentX = 10050
        Printer.Print(CurrencyFormat(txtAccidentInsurance.Text)) 'Disablity Premium

        Printer.CurrentX = 8850
        Printer.CurrentY = 7000
        Printer.Print("X") 'Dual Interest

        Printer.CurrentY = 7050
        Printer.CurrentX = 10050
        Printer.Print(CurrencyFormat(txtPropertyInsurance.Text))

        Printer.EndDoc()
        On Error Resume Next
        SetPrinter(Op)
    End Sub

    Private Function GetBSNumList(ByVal ArAcctNum As String) As String
        Dim RS As ADODB.Recordset, SQL As String
        Dim Current As String, UsedCurrent As Boolean, T As String
        SQL = ""
        SQL = SQL & "SELECT DISTINCT([SaleNo]) AS SN "
        SQL = SQL & "FROM [GrossMargin] "
        SQL = SQL & "WHERE [Style] IN ('NOTES','PAYMENT') "
        SQL = SQL & "AND [Desc] LIKE 'STORE FINANCE%Account [#]" & ArAcctNum & "' "
        SQL = SQL & "ORDER BY SaleNo"
        RS = GetRecordsetBySQL(SQL, , GetDatabaseAtLocation())

        If Not ArMode("E") Then  ' contract estimator
            Current = BillOSale.BillOfSale.Text
        Else
            Current = ""
        End If

        Do While Not RS.EOF
            T = IfNullThenNilString(RS("SN"))
            GetBSNumList = GetBSNumList & IIf(Len(GetBSNumList) > 0, ",", "") & T
            If Current = T Then UsedCurrent = True
            RS.MoveNext()
        Loop
        DisposeDA(RS)

        If Current <> "" And Not UsedCurrent Then GetBSNumList = GetBSNumList & IIf(Len(GetBSNumList) > 0, ",", "") & Current

        If GetBSNumList = "" Then   ' just in case....
            If Not ArMode("E") Then  ' contract estimator
                GetBSNumList = BillOSale.BillOfSale.Text
            Else
                GetBSNumList = ""
            End If
        End If
    End Function

    Private Sub CalculateLastPay()
        Dim PWB As Decimal, Tot As Decimal, Adj As Decimal
        PWB = GetPrice(txtPaymentWillBe.Text)
        Tot = Math.Round(GetPrice(NewBalance) + GetPrice(FinanceCharge) + GetPrice(FinanceChargeSalesTax), 2)
        LastPay = 0
        Payment = 0
        NoMonths = 0

        ' BFH20060505
        ' there isn't a need for checking weekly/monthly here b/c
        ' it is based off txtPaymentWillBe which already checked it.
        NoMonths = Val(txtMonthsToFinance)
        '  If GetPrice(PWB) = 0 Then NoMonths = 0 Else NoMonths = Format(Tot / PWB, "##")
        Adj = Val(NoMonths) * PWB
        Payment = PWB                   ' normal payment remains the same

        ' We could use the Switch() statement here, but it's not that widely used in WinCDS code
        If Adj = Tot Then               ' If X payments of $Y is exactly what we want, ...
            LastPay = PWB                 ' All the payments (in particularly the last one) are the same
        ElseIf Adj > Tot Then           ' if paying X payments of $Y is greater than total, ...
            LastPay = PWB - (Adj - Tot)   ' last payment is normal payment minus overage
        ElseIf Adj < Tot Then           ' Finally, if full payment isn't quite enough, ...
            LastPay = PWB + (Tot - Adj)   ' Add the extra to the normal payment for lastpay
        End If
    End Sub

    Private Sub PrintContractBody(Optional ByVal BSNum As String = "", Optional ByVal BSDate As String = "")
        If Trim(BSNum) = "" Then BSNum = New String("_", 16)
        If Trim(BSDate) = "" Then BSDate = New String("_", 16)
        Printer.FontName = "Arial"
        Printer.FontSize = 6
        Printer.CurrentX = 0
        Printer.CurrentY = 6100 '6500

        '    printer.print( "Buyer acknowledges that the seller has offered to sell the above described merchandise for the cash price indicated, but buyer has elected to purchase on the terms and conditions of this agreement.  The"
        Printer.Print("Buyer acknowledges that the seller has offered to sell the merchandise on ")
        Printer.FontBold = True : Printer.Print("Bill of Sale No(s) ", BSNum) : Printer.FontBold = False
        If IsChicago() Then
            Printer.Print(" and itemized below")
        End If
        Printer.Print(", Dated ")
        Printer.FontBold = True : Printer.Print(BSDate)
        Printer.FontBold = False : Printer.Print(" for the cash price indicated, but buyer has elected to purchase ")
        Printer.Print("on the terms and conditions of this agreement.  The undersigned (BUYER) purchases subject to the terms and conditions as set forth below, from the seller as named above.")
        Printer.Print("")
        Printer.Print("The buyer agrees to pay  ")

        Printer.FontBold = True
        Printer.FontSize = 8

        If optWeekly.Checked = True Then
            Printer.Print((txtMonthsToFinance.Text * 4) - 1)
        Else
            Printer.Print(txtMonthsToFinance.Text - 1)
        End If

        Printer.FontBold = False
        Printer.FontSize = 6

        If optWeekly.Checked = True Then
            Printer.Print(TAB(36), "consecutive weekly payments beginning  ")
        Else
            Printer.Print(TAB(36), "consecutive monthly payments beginning  ")
        End If

        Printer.FontBold = True
        Printer.FontSize = 8
        Printer.Print(FirstPayment)
        Printer.FontBold = False
        Printer.FontSize = 6

        Printer.Print(TAB(95), ".  Each payment shall be   ")

        Printer.FontBold = True
        Printer.FontSize = 8

        Printer.Print(Format(Payment, "$###,##0.00"))

        Printer.FontBold = False
        Printer.FontSize = 6

        Printer.Print(TAB(140), " and an additional final payment which shall be")
        Printer.Print("")

        Printer.FontBold = True
        Printer.FontSize = 8
        Printer.Print("    ", Format(LastPay, "$###,##0.00"))
        Printer.FontBold = False
        Printer.FontSize = 6

        Printer.Print(".  You shall make your payments to us at our office or the address of anyone to whom we may transfer this account to.")
        Printer.Print("")
        Printer.Print("SECURITY: To protect us, you give us a purchase money security interest created under the Uniform Commercial Code of this state in the Property sold and described under ", Chr(34), "Description Of Merchandise", Chr(34), ".")
        Printer.Print("You also give us a security interest in the proceeds from any unauthorized sale of the property, and the proceeds of any insurance you requested.  We waive any other security interest or lien which may arise by")
        Printer.Print("operation of law, except the lien of any judgment which we may obtain if this contract is not paid in accordance with its terms.")
        Printer.Print("Any note given in connection with this contract is understood to be as evidence of, and not in payment of, the obligation hereunder and may be negotiated without waiving any conditions thereof.")
        Printer.Print("")
        Printer.Print("USE AND LOCATION OF PROPERTY:  If you are buying the property primarily for personal, family, of household use, you agree not to use the property in violation of the law.  The property must remain at your")
        Printer.Print("address shown.  You must obtain written permission in advance to move the merchandise.")
        Printer.Print("")
        Printer.Print("DEFAULT:  You are in default if:")
        Printer.Print("A.  We do not receive an installment payment from you on or before due date.")
        Printer.Print("B.  You break one or more of your promises under this contract.")
        Printer.Print("C.  You make any statement or representation in connection with this contract which is false in any material respect.")
        Printer.Print("D.  Insolvency actions are begun by or against you; insolvency includes situations where you are unable to pay all of your debts as they become due.")
        Printer.Print("")
        Printer.Print("ACCELERATION:    If you are in default, we may demand immediate payment of the entire amount you owe.  This includes all the remaining monthly payments you must pay.  We shall have all rights and")
        Printer.Print("remedies given by the Uniform Commercial Code.  This includes the right to retain property.")
        Printer.Print("")
        Printer.Print("REPOSSESSION:  If we retake the property, we have the right to sell it at public or private sale and apply the proceeds of the sale to what you owe, less selling expense.  You agree to pay the difference between")
        Printer.Print("the sale proceeds and what you owe.  We are permitted by law to collect the difference from you.  If we receive more money from the sale than you owe, we will pay the surplus amount to you.")
        Printer.Print("")
        Printer.Print("ATTORNEY'S FEES AND COURT COST:  If this Contract is given to an attorney for collection, you shall pay reasonable attorney's fees, as provided by the laws of this state in which the contract is executed.")
        Printer.Print("You will also pay any court costs if permitted by the law.")
        Printer.Print("")
        Printer.Print("ENTIRE CONTRACT:  No oral promises or statement are part of this contract.  No warranties or representations, whether they are written or arise by operation of law are part of this Contract unless we give")
        Printer.Print("you a written warranty in connection with this contract.")
        Printer.Print("")
        Printer.Print("SIGNERS OF CONTRACT:  If there are more than one of you signing this contract, each of you is individually responsible to see that you fully perform all obligations under this contract.  It is your responsibility")
        Printer.Print("to know whether this contract is in default or that payments have been missed.  We are not responsible for notifying you of late payments, or any default proceedings.")
        Printer.Print("")
        Printer.Print("FINANCING STATEMENT:  You will sign financing statements showing our security interest in the Property which we can file from time to time in any filing office we think appropriate.")
        Printer.Print("")
        Printer.Print("FILING FEES FOR FINANCING THIS STATEMENT:  You will pay any required filing fees on these statements.")
        Printer.Print("")
        Printer.Print("NO WAIVER OF RIGHTS:  We do not waive our right to have future payments made when due if we accept a late or partial payment or delay the enforcement of our rights on any occasion.")
        Printer.Print("")
        Printer.Print("LAW APPLICABLE:  This contract is governed by the law of this state in which it is executed.")
        Printer.Print("")
        Printer.Print("INVALID PROVISIONS:  If any part of this contract becomes invalid or unenforceable the remainder of the Contract will be enforceable.")
        Printer.Print("")
        Printer.Print("NOT PAID IN TIME LIMIT:  If contract is not paid in the months shown on this contract, then interest will continue to be charged at the contract rate.")
        Printer.Print("")
        If Not IsLott Then 'And Not IsMidSouth Then  ' requested this removed
            Printer.Print("BAD CHECKS:  All returned checks fees will be charged $30.00 to the customer.")
            Printer.Print("")
        End If
        Printer.Print("LATE CHARGES:  A Late Charge Fee of:  ")
        Printer.FontBold = True
        Printer.FontSize = 10

        Printer.Print(FormatCurrency(LateCharge))
        Printer.FontBold = False
        Printer.FontSize = 6
        Printer.Print("   will be added automatically to the balance on this account if any payment is received after " & StoreSettings.GracePeriod & " day(s) from the payment due date.")
        Printer.Print("")
        Printer.FontSize = 8

        'Custom area
        '    printer.print( "To contact "; Trim(StoreSettings.Name); " about this account call "; StoreSettings.Phone; "  This contract is subject in whole or part to Texas law which is enforced"
        '    printer.print( "by the Consumer Credit Commission, 2601 N. Lamar Blvd, Austin, Texas  78705-4207; (800) 538-1579; (512) 936-7600, and can be contacted relative"
        '    printer.print( "to any inquiried or complaints."

        If IsChicago() Then
            Printer.Print("All Sales Final!  We do not accept returns or exchanges for merchandise ordered.  Any Special Orders Cancelled will be subject to 25% restocking fee.")
            Printer.Print("The First Payment of an installmant contract will be due 30 days after delivery.  Current payment still due on existing balance.")
            Printer.Print("Received in good condition:")



            '   ElseIf UseAmericanNationalInsurance Then
            '    Printer.FontSize = 6
            '    printer.print( "Sales tax and official fees will be paid by Seller to the appropriate governmental agencies. Fire, extended coverage, credit and Involuntary Unemployment Insurance premiums will be paid by Seller to insurance"
            '    printer.print( "companies. If a Net Balance of Prior Contract is shown above, Buyer's outstanding debt to Seller has been reduced by the amount of finance charges and insurance premiums unearned, and the net balance"
            '    printer.print( "outstanding is included in the Amount Financed of this contract."


        ElseIf IsBoyd Then
            Printer.Print("Copy of Insurance Policy available upon request!")
        ElseIf IsChicago() Then
            Do While Printer.CurrentY + Printer.TextHeight("_") < 14000
                Printer.Print(New String("_", Printer.ScaleWidth \ Printer.TextWidth("_")))
            Loop

        ElseIf IsLott Then ' IsMidSouth Then ' Or IsLott
            Printer.FontBold = True
            Printer.Print("Insurance: ")
            Printer.FontBold = False
            Printer.Print("Credit life insurance and credit disability insurance are not required to obtain credit, ")
            Printer.Print("and will not be provided unless I sign and agree to pay the additional cost.")
            Printer.Print("")
            Printer.Print("")
            Printer.Print("   I Want Credit Life Insurance: _______________________                                       I Want Credit Life and Disability Insurance: _______________________")

        ElseIf IsCarroll Then
            Printer.FontSize = 6
            Printer.FontBold = True
            Printer.Print("Insurance: ")
            Printer.FontBold = False
            Printer.Print("Credit life insurance and credit disability insurance are not required to obtain credit, and will not be provided ")
            Printer.Print("unless I sign and agree to pay the additional cost.  Insurance, if provided, is for the term of the credit sale.")
            Printer.Print("I may obtain required property insurance from anyone I want that is acceptable to you.")
            Printer.Print("")
            Printer.Print("   I Want Credit Life Insurance: _______________________                                       I Want Property Insurance: _______________________")
            Printer.FontSize = 8
        End If

        If IsMichaels Then
            Printer.FontSize = 6
            '    Printer.CurrentY = 13700
            Printer.Print(TAB(5), "NOTICE TO BUYER. ")
            Printer.Print(TAB(5), "ANY HOLDER OF THIS CONSUMER CREDIT CONTRACT IS SUBJECT TO ALL CLAIMS AND DEFENSES WHICH THE DEBTOR ")
            Printer.Print(TAB(5), "COULD ASSERT AGAINST THE SELLER OF GOODS OR SERVICES OBTAINED PURSUANT HERETO OR WITH THE PROCEEDS HEREOF RECOVERY HERE ")
            Printer.Print(TAB(5), "UNDER BY THE DEBTOR SHALL NOT EXCEED AMOUNTS PAID BY THE DEBTOR HERE UNDER. DO NOT SIGN THIS CONTRACT BEFORE YOU READ IT ")
            Printer.Print(TAB(5), "OR IF IT CONTAINS BLANK SPACES, YOU ARE ENTITLED TO A COPY OF THE CONTRACT YOU SIGN. UNDER LAW YOU HAVE THE RIGHT TO PAY ")
            Printer.Print(TAB(5), "OFF IN ADVANCE THE FULL AMOUNT DUE AND UNDER CERTAIN CONDITIONS MAY OBTAIN A PARTIAL REFUND OF THE FINANCE CHARGES.")
            Printer.Print(TAB(5), "KEEP THIS CONTRACT TO PROTECT YOUR LEGAL RIGHTS.")

        Else
            Printer.FontSize = 8
            Printer.CurrentY = 14000 '13900
            Printer.Print(TAB(5), "  NO INSURANCE IS INCLUDED ON FURNITURE;  IT IS THE BUYER'S RESPONSIBILITY FOR ANY LOSS OF OR DAMAGE TO THE MERCHANDISE")
        End If

        If Val(cboCashOption.SelectedIndex) >= 1 Then
            Printer.FontBold = True
            Printer.FontSize = 10
            Printer.Print(cboCashOption.SelectedIndex, " Months Same As Cash Option!")
            Printer.FontBold = False
        Else
            Printer.Print("")
        End If

        Printer.FontSize = 8

        If Val(cboCashOption.SelectedIndex) = 0 Then
            Printer.Print("")
        End If

        Printer.Print("You are entitled to an exact copy of the agreement you sign")
        Printer.Print("Do not sign this agreement before you read it or if it contains")
        Printer.Print("any blanks not filled in.")

        Printer.CurrentY = 14200 '13900
        Printer.FontSize = 14
        PrintToPosition(Printer, "Buyer: ____________________________", Printer.ScaleWidth - 100, AlignConstants.vbAlignRight, False)

        Printer.CurrentY = 14600 '14500
        PrintToPosition(Printer, "Buyer: ____________________________", Printer.ScaleWidth - 100, AlignConstants.vbAlignRight, False)

        Printer.CurrentY = 15000 '14500
        PrintToPosition(Printer, "Guarantor: ____________________________", Printer.ScaleWidth - 100, AlignConstants.vbAlignRight, False)
    End Sub

    Private Sub cmdCancel_Click(sender As Object, e As EventArgs) Handles cmdCancel.Click
        'cancel / done
        If ArMode("REPRINT") Then Hide() : 
        Exit Sub

        If OrderMode("A") Then
            If cmdCancel.Text = "Cancel Set-Up" Then
                Dim X As Integer
                X = BillOSale.X '+ 1
                BillOSale.SetDesc(X, "")
                BillOSale.SetStyle(X, "")
                BillOSale.SetQuan(X, "")
                BillOSale.SetStatus(X, "")
                BillOSale.SetPrice(X, "")
                BillOSale.X = BillOSale.X - 1
                BillOSale.SetDesc(X - 1, "")
                BillOSale.SetStyle(X - 1, "")
                BillOSale.SetQuan(X - 1, "")
                BillOSale.SetStatus(X - 1, "")
                BillOSale.SetPrice(X - 1, "")
                BillOSale.NewStyleLine = BillOSale.NewStyleLine - 2
                OrdSelect.ArStatus = "" 'C
                BillOSale.GetGrid.MoveRow(BillOSale.GetGrid.Row - 2)  '***does not work in deliver mode
                BillOSale.GetGrid.Refresh(True)
            End If
            'Unload Me
            Me.Close()
        ElseIf OrderMode("D") Then
            OrdPay.cmdOk.Enabled = True
            OrdPay.cmdCancel.Enabled = True
            'Unload Me
            Me.Close()
        ElseIf OrderMode("Credit") Then
            MessageBox.Show("Add On was cancelled." & vbCrLf & "Adjustments was made on this Installment Sale." & vbCrLf & "To make it balance, make a payment to the sale or reverse the adjustment back to its original version.")
            OrdPay.cmdOk.Enabled = True
            OrdPay.cmdCancel.Enabled = True
            'Unload Me
            Me.Close()
        Else
            'Unload Me
            Me.Close()
            'Unload OrdSelect
            OrdSelect.Close()
            modProgramState.ArSelect = ""
        End If

        If IsIn(AddOn, ArAddOn_Add) Then
            'Unload ArCard
            ArCard.Close()
        End If

        OrdSelect.ArStatus = ""
    End Sub

    Private Sub ARPaySetUp_Activated(sender As Object, e As EventArgs) Handles MyBase.Activated
        If DateBetween(Today, #9/10/2016#, #8/28/2016#) Then
            cboCashOption.BackColor = Color.FromArgb(256, 256, 128)
            '    lblCashOpt.BackColor = RGB(256, 256, 128)
            'lblCashOpt.FontUnderline = True
            lblCashOpt.Font = New Font(lblCashOpt.Font.Name, lblCashOpt.Font.Size, FontStyle.Underline)
            cboDeferred.BackColor = Color.FromArgb(150, 150, 256)
            '    lblDeferred.BackColor = RGB(150, 150, 256)
            'lblDeferred.FontUnderline = True
            lblDeferred.Font = New Font(lblDeferred.Font.Name, lblDeferred.Font.Size, FontStyle.Underline)
        End If
    End Sub

    Private Sub ARPaySetUp_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        'If no form cancels the QueryUnload event, the Unload event occurs first in all other forms and then in an MDI form. 
        'When a child form or a Form object closes, the QueryUnload event in that form occurs before the form's Unload event

        '--------------
        'This code block is for Query unload event of vb6.0
        If OrderMode("B") Then  'deliver sale
            'If UnloadMode = vbFormControlMenu Then Cancel = True
            If e.CloseReason = CloseReason.UserClosing Then e.Cancel = True
            Exit Sub
        End If

        'If UnloadMode = vbFormControlMenu Then
        If e.CloseReason = CloseReason.UserClosing Then
            UnloadARPaySetUp = True
        End If
        '------------------------

        'This block of code is for Form Unload event of vb6.0. Cause in vb.net, there are no two separate events unload and queryunload. So joined both events 
        'in FormClosing event of vb.net
        On Error Resume Next

        If Not ArMode("E") Then ' contract estimator
            mDBAccess.dbClose()
            mDBAccessTransactions.dbClose()
        Else
            modProgramState.ArSelect = ""
            'Unload Me
            Me.Close()
            MainMenu.Show()
        End If
        mDBAccess = Nothing
        mDBAccessTransactions = Nothing

        If ArMode("S") Then
            'Unload AddOnAcc  ' added: bfh20050803
            AddOnAcc.Close()
            'Unload MailCheck ' added: bfh20050803
            MailCheck.Close()
            'Unload BillOSale
            BillOSale.Close()
            MainMenu.Show()    ' added: bfh20050803
        End If
    End Sub

    Private Sub optLate16_Click(sender As Object, e As EventArgs) Handles optLate16.Click
        DueOn = 10
        If Not NoAdjust Then AdjustFirstPay(10)
        UpdateLateCaptions()
    End Sub

    Private Sub optLate26_Click(sender As Object, e As EventArgs) Handles optLate26.Click
        DueOn = 20
        If Not NoAdjust Then AdjustFirstPay(20)
        UpdateLateCaptions()
    End Sub

    Private Sub optMonthly_Click(sender As Object, e As EventArgs) Handles optMonthly.Click
        Recalculate()
    End Sub

    Private Sub optWeekly_Click(sender As Object, e As EventArgs) Handles optWeekly.Click
        Recalculate()
    End Sub

    Private Sub tmrLoad_Tick(sender As Object, e As EventArgs) Handles tmrLoad.Tick
        On Error Resume Next
        If ArMode("E") Then txtPrevBalance.Select()
        tmrLoad.Enabled = False
    End Sub

    Private Sub txtArNo_TextChanged(sender As Object, e As EventArgs) Handles txtArNo.TextChanged
        Dim R1 As Boolean, R2 As Boolean
        R1 = IsRevolvingCharge(LastCheckedArNo)
        R2 = IsRevolvingCharge(txtArNo.Text)
        LastCheckedArNo = txtArNo.Text
        If R1 <> R2 Then
            If R2 Then
                SetDefaultsRevolving()
            Else
                SetDefaultsInstallment()
            End If
        End If
    End Sub

    Private Sub txtFinanceCharges_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles txtFinanceCharges.Validating
        '  If MsgBox("Are you sure you wish to change the Finance Charges?" & vbCrLf & "This will alter the actual APR rate if you do.", vbQuestion + vbOKCancel, "Change Finance Charges") = vbOK Then
        '    MsgBox "Note that if you change one of the other values after editing this field, it will be reset.", vbInformation, "Note"
        FinanceCharge = GetPrice(txtFinanceCharges.Text)
        txtTotalBalance.Text = CurrencyFormat(NewBalance + FinanceCharge + FinanceChargeSalesTax)
        If Not ArMode("S") And (UseAmericanNationalInsurance) Then Exit Sub
        RecalculateFinancing(True)
        '  Else
        '    Cancel = True
        '  End If
    End Sub

    Private Sub txtPaymentWillBe_TextChanged(sender As Object, e As EventArgs) Handles txtPaymentWillBe.TextChanged
        CalculateMath()
    End Sub

    Private Sub txtPaymentWillBe_DoubleClick(sender As Object, e As EventArgs) Handles txtPaymentWillBe.DoubleClick
        Recalculate()

        If fraMath.Visible Then
            fraMath.Visible = False
        Else
            CalculateMath()
            fraMath.Visible = True
            'fraMath.ZOrder 0
            fraMath.BringToFront()
        End If
    End Sub

    Private Sub txtPrevBalance_DoubleClick(sender As Object, e As EventArgs) Handles txtPrevBalance.DoubleClick
        If IsDevelopment() Then InsuranceFormTreeHouse()
    End Sub

    Private Sub txtArNo_DoubleClick(sender As Object, e As EventArgs) Handles txtArNo.DoubleClick
        If IsDevelopment() Then InsuranceFormTreeHouse()
    End Sub

    Private Sub txtRate_DoubleClick(sender As Object, e As EventArgs) Handles txtRate.DoubleClick
        If Trim(txtRate.Text) <> "" Then txtRate.Text = Format(Rate, "##.00")
        'SelectContents(txtRate.Text)
        SelectContents(txtRate)
    End Sub

    Private Sub txtRate_MouseDown(sender As Object, e As MouseEventArgs) Handles txtRate.MouseDown
        'If Button = vbRightButton Then DebugAPR
        If e.Button = MouseButtons.Right Then DebugAPR()
    End Sub

    Private Sub DebugAPR()
        Dim M As String, Pattern As String
        Pattern = IIf(IsDevelopment, "0.00000000%", "0.00%")
        M = ""
        M = M & "APR    = " & Format(APR / 100, Pattern) & vbCrLf
        M = M & "Simple =" & Format(SIR, Pattern)
        MessageBox.Show(M, "APR INFORMATION", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub

    Private Sub lblAPR_DoubleClick(sender As Object, e As EventArgs) Handles lblAPR.DoubleClick
        DebugAPR()
    End Sub

    Private Sub txtRate_TextChanged(sender As Object, e As EventArgs) Handles txtRate.TextChanged
        'txtRate_LostFocus
    End Sub

    Private Sub txtRate_Leave(sender As Object, e As EventArgs) Handles txtRate.Leave
        'BFH20170428 - Re-instantiated the Rate = ... line because otherwise tab-off of rate wouldn't update financing...
        'disabled following line to kill the abnormal behavior of the apr value changing for 21.00 to 0 wkd 17 MAR 17
        Rate = Format(GetPrice(txtRate.Text), "##.00")
        Recalculate()
        lblAPR.Text = Format(APR, "#0.00")
    End Sub

    Private Sub txtPaymentWillBe_Leave(sender As Object, e As EventArgs) Handles txtPaymentWillBe.Leave
        txtPaymentWillBe.Text = CurrencyFormat(GetPrice(txtPaymentWillBe.Text))
    End Sub

    Private Sub txtPrevBalance_Leave(sender As Object, e As EventArgs) Handles txtPrevBalance.Leave
        txtPrevBalance.Text = CurrencyFormat(GetPrice(txtPrevBalance.Text))
        Recalculate()
    End Sub

    Private Sub txtAddlPaymentsMade_Leave(sender As Object, e As EventArgs) Handles txtAddlPaymentsMade.Leave
        txtAddlPaymentsMade.Text = CurrencyFormat(GetPrice(txtAddlPaymentsMade.Text))
    End Sub

    Private Sub txtBalDueLateCharge_Leave(sender As Object, e As EventArgs) Handles txtBalDueLateCharge.Leave
        txtBalDueLateCharge.Text = CurrencyFormat(GetPrice(txtBalDueLateCharge.Text))
    End Sub

    Private Sub txtGrossSale_Leave(sender As Object, e As EventArgs) Handles txtGrossSale.Leave
        txtGrossSale.Text = CurrencyFormat(GetPrice(txtGrossSale.Text))
        Recalculate()
    End Sub

    Private Sub txtGrossSale_Enter(sender As Object, e As EventArgs) Handles txtGrossSale.Enter
        'SelectContents(txtGrossSale.Text)
        SelectContents(txtGrossSale)
        If ArMode("S") Then
            If Trim(txtArNo.Text) = "" Then
                If MessageBox.Show("Select Auto Account or Enter Manual Account!", "", MessageBoxButtons.OK, MessageBoxIcon.Warning) = DialogResult.OK Then
                    txtArNo.Select()
                    Exit Sub
                End If
            End If
        End If
    End Sub

    Private Sub txtOrigDeposit_Leave(sender As Object, e As EventArgs) Handles txtOrigDeposit.Leave
        txtOrigDeposit.Text = CurrencyFormat(txtOrigDeposit.Text)
        Recalculate()
    End Sub

    Private Sub txtMonthsToFinance_Leave(sender As Object, e As EventArgs) Handles txtMonthsToFinance.Leave
        Recalculate()
    End Sub

    Private Sub txtMonthsToFinance_Enter(sender As Object, e As EventArgs) Handles txtMonthsToFinance.Enter
        'SelectContents(txtMonthsToFinance.Text)
        SelectContents(txtMonthsToFinance)
    End Sub

    Private Sub txtDocFee_Leave(sender As Object, e As EventArgs) Handles txtDocFee.Leave
        Recalculate()
    End Sub

    Private Sub txtAccidentInsurance_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtAccidentInsurance.KeyPress
        'chkAccident.Value = vbGrayed
        chkAccident.CheckState = CheckState.Indeterminate
    End Sub

    Private Sub txtLifeInsurance_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtLifeInsurance.KeyPress
        'chkLife.Value = vbGrayed
        chkLife.CheckState = CheckState.Indeterminate
    End Sub

    Private Sub txtPropertyInsurance_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtPropertyInsurance.KeyPress
        'chkProperty.Value = vbGrayed
        chkProperty.CheckState = CheckState.Indeterminate
    End Sub

    Private Sub txtUnemploymentInsurance_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtUnemploymentInsurance.KeyPress
        'chkUnemployment.Value = vbGrayed
        chkUnemployment.CheckState = CheckState.Indeterminate
    End Sub

    Private Sub txtPrevBalance_Enter(sender As Object, e As EventArgs) Handles txtPrevBalance.Enter
        'SelectContents(txtPrevBalance.Text)
        SelectContents(txtPrevBalance)
    End Sub

    Private Sub txtFinanceChargeSalesTax_TextChanged(sender As Object, e As EventArgs) Handles txtFinanceChargeSalesTax.TextChanged
        'SelectContents(txtFinanceChargeSalesTax.Text)
        SelectContents(txtFinanceChargeSalesTax)
    End Sub

    Private Sub txtOrigDeposit_Enter(sender As Object, e As EventArgs) Handles txtOrigDeposit.Enter
        'SelectContents(txtOrigDeposit.Text)
        SelectContents(txtOrigDeposit)
    End Sub

    Private Sub txtDocFee_Enter(sender As Object, e As EventArgs) Handles txtDocFee.Enter
        'SelectContents(txtDocFee.Text)
        SelectContents(txtDocFee)
    End Sub

    Private Sub txtLifeInsurance_Enter(sender As Object, e As EventArgs) Handles txtLifeInsurance.Enter
        'SelectContents(txtLifeInsurance.Text)
        SelectContents(txtLifeInsurance)
    End Sub

    Private Sub txtAccidentInsurance_Enter(sender As Object, e As EventArgs) Handles txtAccidentInsurance.Enter
        'SelectContents(txtAccidentInsurance.Text)
        SelectContents(txtAccidentInsurance)
    End Sub

    Private Sub txtPropertyInsurance_Enter(sender As Object, e As EventArgs) Handles txtPropertyInsurance.Enter
        'SelectContents(txtPropertyInsurance.Text)
        SelectContents(txtPropertyInsurance)
    End Sub

    Private Sub txtUnemploymentInsurance_Enter(sender As Object, e As EventArgs) Handles txtUnemploymentInsurance.Enter
        'SelectContents(txtUnemploymentInsurance.Text)
        SelectContents(txtUnemploymentInsurance)
    End Sub

    Private Sub txtLifeInsurance_Leave(sender As Object, e As EventArgs) Handles txtLifeInsurance.Leave
        txtLifeInsurance.Text = CurrencyFormat(GetPrice(txtLifeInsurance.Text))
        Recalculate()
    End Sub

    Private Sub txtAccidentInsurance_Leave(sender As Object, e As EventArgs) Handles txtAccidentInsurance.Leave
        txtAccidentInsurance.Text = CurrencyFormat(GetPrice(txtAccidentInsurance.Text))
        Recalculate()
    End Sub

    Private Sub txtPropertyInsurance_Leave(sender As Object, e As EventArgs) Handles txtPropertyInsurance.Leave
        txtPropertyInsurance.Text = CurrencyFormat(GetPrice(txtPropertyInsurance.Text))
        Recalculate()
    End Sub

    Private Sub txtUnemploymentInsurance_Leave(sender As Object, e As EventArgs) Handles txtUnemploymentInsurance.Leave
        txtUnemploymentInsurance.Text = CurrencyFormat(GetPrice(txtUnemploymentInsurance.Text))
        Recalculate()
    End Sub

    Private Sub txtFinanceCharges_Enter(sender As Object, e As EventArgs) Handles txtFinanceCharges.Enter
        'SelectContents(txtFinanceCharges.Text)
        SelectContents(txtFinanceCharges)
    End Sub

    Private Sub txtSubTotal_Enter(sender As Object, e As EventArgs) Handles txtSubTotal.Enter
        'SelectContents(txtSubTotal.Text)
        SelectContents(txtSubTotal)
    End Sub

    Private Sub dteDate1_ValueChanged(sender As Object, e As EventArgs) Handles dteDate1.ValueChanged
        AdjustFirstPay()
        CheckLateDay()
        UpdateLateCaptions()
    End Sub

    Private Sub dteDate2_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles dteDate2.Validating
        If Not NoAdjust Then CheckLateDay()  ' No changes needed.
    End Sub

    Private Sub cboDeferred_Click(sender As Object, e As EventArgs) Handles cboDeferred.Click
        Dim tRate As Double
        OneOrTheOther(True)
        AdjustFirstPay()

        tRate = APR
        Recalculate()
    End Sub

    Private Sub InsuranceForm() 'Elmore, Lott
        Dim Op As Object
        Dim FA As Decimal
        '  If IsMidSouth Then ' Or IsLott
        '    FA = GetPrice(txtFinanceAmount) + GetPrice(txtFinanceCharges) + GetPrice(txtFinanceChargeSalesTax)
        '  Else
        FA = GetPrice(txtFinanceAmount.Text)
        '  End If
        Op = Printer.DeviceName

        picPicture.Image = LoadPictureStd(FXFile("Insurance Contract 2.gif"))
        Printer.PaintPicture(picPicture.Image, 0, 0, Printer.ScaleWidth, Printer.ScaleHeight)

        Printer.FontName = "Arial"
        Printer.FontSize = 12

        Printer.CurrentX = 5850
        Printer.CurrentY = 1175
        Printer.Print(StoreSettings.Name)

        Printer.CurrentX = 10100
        Printer.Print(ArNo) 'Creditor's No.

        Printer.FontSize = 8
        Printer.CurrentX = 200
        Printer.CurrentY = 1675
        Printer.Print(Trim(BillOSale.CustomerFirst.Text), "  ", Trim(BillOSale.CustomerLast.Text))
        Printer.CurrentX = 200
        Printer.Print(Trim(BillOSale.CustomerAddress.Text), ",  ", Trim(BillOSale.CustomerCity.Text), " ", Trim(BillOSale.CustomerZip.Text))

        Printer.CurrentX = 200
        Printer.CurrentY = 2500
        Printer.FontSize = 12
        Printer.Print(dteDate2.Value) 'Date

        Printer.CurrentX = 3800
        '.CurrentY = 2880
        Printer.Print(txtMonthsToFinance.Text) 'Term (Months)

        Printer.CurrentX = 3060
        Printer.FontSize = 12
#If False Then
  Printer.CurrentY = 3250    '#2 Joint Life Net Decreasing Term
#Else
        Printer.CurrentY = 3500    '#3 Single Life Level Term
#End If
        Printer.Print("X")

        Printer.CurrentX = 6300
        Printer.CurrentY = 3575
        Printer.Print(Format(APR, "##.##"))

        Printer.CurrentY = 3500
        Printer.CurrentX = 8000
        Printer.Print(CurrencyFormat(FA))
        Printer.CurrentX = 10100
        Printer.Print(CurrencyFormat(txtLifeInsurance.Text))

        Printer.CurrentY = 4550
        Printer.CurrentX = 400
        If IsLott Then ' IsMidSouth Then ' Or IsLott
            Printer.Print("14")      ' BFH20071202
        Else
            Printer.Print("7")
        End If
        Printer.CurrentX = 6075
        Printer.Print("X") 'Single
        Printer.CurrentX = 8000
        Printer.CurrentY = 5000
        Printer.Print(txtPaymentWillBe.Text) 'Monthly Disability Proceeds
        Printer.CurrentX = 10100
        Printer.Print(CurrencyFormat(txtAccidentInsurance.Text)) 'Disablity Premium

        Printer.CurrentX = 2930
        Printer.CurrentY = 6100
        Printer.Print("X") 'Dual Interest

        Printer.CurrentX = 3900
        Printer.CurrentY = 6600
        Printer.Print(CurrencyFormat(FA))
        Printer.CurrentX = 6400
        Printer.Print(CurrencyFormat(FA))
        Printer.CurrentX = 10100
        Printer.Print(CurrencyFormat(txtPropertyInsurance))

        Printer.EndDoc()
        On Error Resume Next
        SetPrinter(Op)
    End Sub

    Private Sub InsuranceForm_New(ByVal N As Integer) 'Lott, 20080311
        Dim Op As Object, OPS As Object, O As Object, Adj As Single
        Dim FA As Decimal
        Adj = -0.69

        Op = Printer.DeviceName
        SetPrinter(LegalContractPrinter)
        OPS = Printer.PaperSize
        O = Printer

        If IsLott Then ' IsMidSouth Then ' Or IsLott
            FA = GetPrice(txtFinanceAmount.Text) + GetPrice(txtFinanceCharges.Text) + GetPrice(txtFinanceChargeSalesTax.Text)
        Else
            FA = GetPrice(txtFinanceAmount.Text)
        End If

        Dim ExpDate As Date
        Dim tLifPre As Decimal, tDisPre As Decimal, tProPre As Decimal, tIUIPre As Decimal
        Dim bLif As Boolean, bDis As Boolean, bPro As Boolean, bIUI As Boolean
        bLif = chkLife.Checked = True
        bDis = chkAccident.Checked = True
        bPro = chkProperty.Checked = True
        bIUI = chkUnemployment.Checked = True
        If True Or Val(txtMonthsToFinance.Text) <> 0 Then
            tLifPre = IIf(bLif, GetPrice(txtLifeInsurance.Text), 0) '/ Val(txtMonthsToFinance)
            tDisPre = IIf(bDis, GetPrice(txtAccidentInsurance.Text), 0) '/ Val(txtMonthsToFinance)
            tProPre = IIf(bPro, GetPrice(txtPropertyInsurance.Text), 0) '/ Val(txtMonthsToFinance)
            tIUIPre = IIf(bIUI, GetPrice(txtUnemploymentInsurance.Text), 0)
        End If

        ExpDate = DateAdd("m", CDbl(txtMonthsToFinance.Text), dteDate1.Value)

        ' 8.5"x14"
        Printer.PaperSize = vbPRPSLegal
        Printer.ScaleMode = vbInches
        picPicture.Image = LoadPictureStd(FXFile("CentralInsurance" & N & ".gif"))
        Printer.PaintPicture(picPicture.Image, 0, 0, Printer.ScaleWidth, Printer.ScaleHeight)

        Printer.FontName = "Arial"
        Printer.FontSize = 8

        PrintOut_Alt(OutObj:=O, X:=3.83, Y:=1.57 + Adj, FontSize:=18, FontBold:=True, Text:=ArNo)

        PrintOut_Alt(OutObj:=O, X:=0.5, Y:=2.09 + Adj, FontSize:=8, Text:=Trim(BillOSale.CustomerFirst.Text) & " " & Trim(BillOSale.CustomerLast.Text))
        '    PrintOut_Alt( OutObj:=O, X:=2.9, Y:=2+adj, FontSize:=8, Text:="AGE" ' Age
        '
        '    PrintOut_Alt( OutObj:=O, X:=3.53, Y:=2+adj, FontSize:=8, Text:="DM" ' DOB
        '    PrintOut_Alt( OutObj:=O, X:=3.94, Y:=2+adj, FontSize:=8, Text:="DD" ' DOB
        '    PrintOut_Alt( OutObj:=O, X:=4.24, Y:=2+adj, FontSize:=8, Text:="DBYY" ' DOB
        '
        PrintOut_Alt(OutObj:=O, X:=0.5, Y:=2.55 + Adj, FontSize:=8, Text:=Trim(BillOSale.CustomerAddress.Text) & ",  " & Trim(BillOSale.CustomerCity.Text) & " " & Trim(BillOSale.CustomerZip.Text))

        If True Then
            '      PrintOut_Alt( OutObj:=O, X:=0.5, Y:=3.04 + Adj, FontSize:=8, Text:="Co-Buyer Name"
            '      PrintOut_Alt( OutObj:=O, X:=0.5, Y:=3.49 + Adj, FontSize:=8, Text:="First Beneficiary"
        End If


        ' life/dis
        If bLif Then PrintOut_Alt(OutObj:=O, X:=1.3, Y:=4.15 + Adj, FontSize:=14, FontBold:=True, Text:=txtMonthsToFinance.Text)
        If bDis Then PrintOut_Alt(OutObj:=O, X:=2.54, Y:=4.15 + Adj, FontSize:=14, FontBold:=True, Text:=txtMonthsToFinance.Text)
        PrintOut_Alt(OutObj:=O, X:=3.89, Y:=4.15 + Adj, FontSize:=14, FontBold:=True, Text:=txtMonthsToFinance.Text)
        If bLif Or bDis Then
            PrintOut_Alt(OutObj:=O, X:=5.21, Y:=4.15 + Adj, FontSize:=14, FontBold:=True, Text:=DateDiff("d", dteDate1.Value, dteDate2.Value))
            PrintOut_Alt(OutObj:=O, X:=6.25, Y:=4.21 + Adj, FontSize:=10, FontBold:=True, Text:=Format(dteDate1, "mm dd yy"))
            PrintOut_Alt(OutObj:=O, X:=7.53, Y:=4.21 + Adj, FontSize:=10, FontBold:=True, Text:=Format(ExpDate, "mm dd yy"))
        End If

        If optJointLife0.Checked = True Then
            PrintOut_Alt(OutObj:=O, X:=1.44, Y:=4.46 + Adj, FontSize:=7, FontBold:=True, Text:="X")  ' single
        Else
            PrintOut_Alt(OutObj:=O, X:=2.25, Y:=4.46 + Adj, FontSize:=7, FontBold:=True, Text:="X")  ' joint
        End If

        If True Then
            PrintOut_Alt(OutObj:=O, X:=0.41, Y:=5.17 + Adj, FontSize:=7, FontBold:=True, Text:="X") ' joint-level
        Else
            '      PrintOut_Alt( OutObj:=O, X:=0.41, Y:=4.7 + Adj, FontSize:=7, FontBold:=True, Text:="X"  ' Gross Decreasing Term Life
            '      PrintOut_Alt( OutObj:=O, X:=0.41, Y:=4.87 + Adj, FontSize:=7, FontBold:=True, Text:="X"  ' Net Pay Decreasing Term Life
            '      PrintOut_Alt( OutObj:=O, X:=0.41, Y:=5.04 + Adj, FontSize:=7, FontBold:=True, Text:="X"  ' Life With Dismemberment
        End If

        If bLif Then
            PrintOut_Alt(OutObj:=O, X:=5.85, Y:=4.76 + Adj, FontSize:=10, Text:=CurrencyFormat(FA))
            If tLifPre <> 0 Then PrintOut_Alt(OutObj:=O, X:=7.13, Y:=4.76 + Adj, FontSize:=10, Text:=CurrencyFormat(tLifPre))
        End If


        If bDis Then
            PrintOut_Alt(OutObj:=O, X:=5.85, Y:=6.24 + Adj, FontSize:=10, Text:=txtPaymentWillBe.Text)
            If tDisPre <> 0 Then PrintOut_Alt(OutObj:=O, X:=7.13, Y:=6.24 + Adj, FontSize:=10, Text:=CurrencyFormat(tDisPre))
        End If
        PrintOut_Alt(OutObj:=O, X:=7.13, Y:=6.55 + Adj, FontSize:=10, Text:=CurrencyFormat(tLifPre + tDisPre))


        If bPro Then
            PrintOut_Alt(OutObj:=O, X:=0.92, Y:=10.23 + Adj, FontSize:=8, Text:=Month(dteDate1.Value))
            PrintOut_Alt(OutObj:=O, X:=1.36, Y:=10.23 + Adj, FontSize:=8, Text:=DateAndTime.Day(dteDate1.Value))
            PrintOut_Alt(OutObj:=O, X:=1.86, Y:=10.23 + Adj, FontSize:=8, Text:=Year(dteDate1.Value))


            PrintOut_Alt(OutObj:=O, X:=2.92, Y:=10.23 + Adj, FontSize:=8, Text:=Month(ExpDate))
            PrintOut_Alt(OutObj:=O, X:=3.38, Y:=10.23 + Adj, FontSize:=8, Text:=DateAndTime.Day(ExpDate))
            PrintOut_Alt(OutObj:=O, X:=3.9, Y:=10.23 + Adj, FontSize:=8, Text:=Year(ExpDate))

            PrintOut_Alt(OutObj:=O, X:=4.63, Y:=10.18 + Adj, FontSize:=11, Text:=CurrencyFormat(FA))
            PrintOut_Alt(OutObj:=O, X:=7.19, Y:=10.18 + Adj, FontSize:=10, FontBold:=True, Text:=txtMonthsToFinance.Text)

            PrintOut_Alt(OutObj:=O, X:=4.51, Y:=10.41 + Adj, FontSize:=7, FontBold:=True, Text:="X")
            If tProPre <> 0 Then PrintOut_Alt(OutObj:=O, X:=6.49, Y:=11.41 + Adj, FontSize:=10, Text:=CurrencyFormat(tProPre))
        End If


        '    PrintOut_Alt( OutObj:=O, X:=0, Y:=7000, FontSize:=10, Text:=dteDate2
        '    PrintOut_Alt( OutObj:=O, X:=0, Y:=7000, FontSize:=10, Text:=dteDate2
        '    PrintOut_Alt( OutObj:=O, X:=0, Y:=7000, FontSize:=10, Text:=dteDate2


        Printer.EndDoc()
        On Error Resume Next
        SetPrinter(Op)
        Printer.PaperSize = OPS
        Printer.ScaleMode = vbTwips ' default
    End Sub

    Private Sub cmdPrint_Click(sender As Object, e As EventArgs) Handles cmdPrint.Click
        'cmdApply.Value = True
        cmdApply_Click(cmdApply, New EventArgs)
    End Sub

    Private Sub GetArNo()
        If AccountFound <> "Y" Or Status = "V" Then 'Addon
            'bfh20051206
            '    ArNo = GetFileAutonumber(frmSetup .StoreOrdDrv + "NewOrder\ArNo.Dat", 2000)
            ArNo = GetFileAutonumber(ArNoFile, 2000)
            lblAcctNo.Visible = True
            lblAccountNo.Text = ArNo
            lblAccountNo.Visible = True
        End If
    End Sub

    Private Sub txtArNo_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtArNo.KeyPress
        'KeyAscii = Asc(UCase(Chr(KeyAscii)))
        e.KeyChar = UCase(e.KeyChar)
    End Sub

    Private Sub chkAutoARNO_Click(sender As Object, e As EventArgs) Handles chkAutoARNO.Click
        'Arno
        If chkAutoARNO.Checked = True Then
            GetArNo()
            txtArNo.Text = ArNo
        End If
        If chkAutoARNO.Checked = False Then 'uncheck
            txtArNo.Text = ""
        End If
    End Sub

    Private Sub chkLife_Click(sender As Object, e As EventArgs) Handles chkLife.Click
        'turn on life
        If chkLife.CheckState = CheckState.Indeterminate Then Exit Sub
        If chkLife.CheckState = CheckState.Checked Then
            GetLife()
        Else
            txtLifeInsurance.Text = ""
        End If

        If IsElmore Or IsBoyd Or UseAmericanNationalInsurance Or UseThorntonsInsurance Or IsLott Then ' Or IsLott Or IsMidSouth
            Recalculate()
        Else
            txtFinanceAmount.Text = CurrencyFormat(GetPrice(txtSubTotal.Text) + GetPrice(txtDocFee.Text) + GetPrice(txtLifeInsurance.Text) + GetPrice(txtAccidentInsurance.Text) + GetPrice(txtPropertyInsurance.Text) + GetPrice(txtUnemploymentInsurance.Text))
            FinanceCharge = ((txtFinanceAmount.Text * InterestRate) / 12 * Months)
            txtFinanceCharges.Text = CurrencyFormat(FinanceCharge)
            If optWeekly.Checked = True Then
                Payment = (GetPrice(txtFinanceAmount.Text) + FinanceCharge) / (Months * 4)
            Else
                Payment = (GetPrice(txtFinanceAmount.Text) + FinanceCharge) / Months
            End If
            txtPaymentWillBe.Text = CurrencyFormat(Payment)
        End If
    End Sub

    Private Sub optJointLife0_Click(sender As Object, e As EventArgs) Handles optJointLife0.Click
        'turn on life
        If chkLife.Checked = True Then
            GetLife()
        End If

        Recalculate()

        ' Moved into Recalculate.  This was giving different answers depending what was clicked.
        '  If IsElmore Or IsLott Or IsMidSouth Or IsBoyd Or IsTreehouse Or IsBlueSky Then
        '    Recalculate
        '  Else
        '    txtFinanceAmount = CurrencyFormat(GetPrice(txtSubTotal) + GetPrice(txtDocFee) + GetPrice(txtLifeInsurance) + GetPrice(txtAccidentInsurance) + GetPrice(txtPropertyInsurance)) + GetPrice(txtUnemploymentInsurance)
        '    FinanceCharge = ((txtFinanceAmount * InterestRate) / 12 * Months)
        '    txtFinanceCharges = CurrencyFormat(FinanceCharge)
        '    If optWeekly Then
        '      Payment = (GetPrice(txtFinanceAmount) + FinanceCharge) / (Months * 4)
        '    Else
        '      Payment = (GetPrice(txtFinanceAmount) + FinanceCharge) / Months
        '    End If
        '    txtPaymentWillBe = CurrencyFormat(Payment)
        ' End If
    End Sub

    Private Sub optJointLife1_Click(sender As Object, e As EventArgs) Handles optJointLife1.Click
        'turn on life
        If chkLife.Checked = True Then
            GetLife()
        End If

        Recalculate()

        ' Moved into Recalculate.  This was giving different answers depending what was clicked.
        '  If IsElmore Or IsLott Or IsMidSouth Or IsBoyd Or IsTreehouse Or IsBlueSky Then
        '    Recalculate
        '  Else
        '    txtFinanceAmount = CurrencyFormat(GetPrice(txtSubTotal) + GetPrice(txtDocFee) + GetPrice(txtLifeInsurance) + GetPrice(txtAccidentInsurance) + GetPrice(txtPropertyInsurance)) + GetPrice(txtUnemploymentInsurance)
        '    FinanceCharge = ((txtFinanceAmount * InterestRate) / 12 * Months)
        '    txtFinanceCharges = CurrencyFormat(FinanceCharge)
        '    If optWeekly Then
        '      Payment = (GetPrice(txtFinanceAmount) + FinanceCharge) / (Months * 4)
        '    Else
        '      Payment = (GetPrice(txtFinanceAmount) + FinanceCharge) / Months
        '    End If
        '    txtPaymentWillBe = CurrencyFormat(Payment)
        ' End If
    End Sub

    Private Sub chkAccident_Click(sender As Object, e As EventArgs) Handles chkAccident.Click
        'turn on Accident
        If chkAccident.CheckState = CheckState.Indeterminate Then Exit Sub
        If chkAccident.Checked = True Then
            GetAcc()
        Else
            txtAccidentInsurance.Text = ""
        End If

        If IsElmore Or IsBoyd Or UseAmericanNationalInsurance Or UseThorntonsInsurance Or IsLott Then  ' Or IsLott Or IsMidSouth
            Recalculate()
        Else
            txtFinanceAmount.Text = CurrencyFormat(GetPrice(txtSubTotal.Text) + GetPrice(txtDocFee.Text) + GetPrice(txtLifeInsurance.Text) + GetPrice(txtAccidentInsurance.Text) + GetPrice(txtPropertyInsurance.Text) + GetPrice(txtUnemploymentInsurance.Text))
            If Months = 0 Then
                FinanceCharge = 0
            Else
                FinanceCharge = ((txtFinanceAmount.Text * InterestRate) / 12 * Months)
            End If
            txtFinanceCharges.Text = CurrencyFormat(FinanceCharge)
            If Months = 0 Then
                Payment = 0
            Else
                If optWeekly.Checked = True Then
                    Payment = (GetPrice(txtFinanceAmount.Text) + FinanceCharge) / (Months * 4)
                Else
                    Payment = (GetPrice(txtFinanceAmount.Text) + FinanceCharge) / Months
                End If
            End If
            txtPaymentWillBe.Text = CurrencyFormat(Payment)
        End If
    End Sub

    Private Sub chkProperty_Click(sender As Object, e As EventArgs) Handles chkProperty.Click
        'turn off property
        If chkProperty.CheckState = CheckState.Indeterminate Then Exit Sub
        If chkProperty.Checked = True Then
            GetProp()
        Else
            txtPropertyInsurance.Text = ""
        End If

        If IsElmore Or IsBoyd Or UseAmericanNationalInsurance Or UseThorntonsInsurance Or IsLott Then  ' Or IsLott Or IsMidSouth
            Recalculate()
        Else
            txtFinanceAmount.Text = CurrencyFormat(GetPrice(txtSubTotal.Text) + GetPrice(txtDocFee.Text) + GetPrice(txtLifeInsurance.Text) + GetPrice(txtAccidentInsurance.Text) + GetPrice(txtPropertyInsurance.Text) + GetPrice(txtUnemploymentInsurance.Text))
            If Months = 0 Then
                FinanceCharge = 0
            Else
                FinanceCharge = ((txtFinanceAmount.Text * InterestRate) / 12 * Months)
            End If
            txtFinanceCharges.Text = CurrencyFormat(FinanceCharge)
            If Months = 0 Then

                Payment = 0
            Else
                Payment = (GetPrice(txtFinanceAmount.Text) + FinanceCharge) / Months
            End If
            txtPaymentWillBe.Text = CurrencyFormat(Payment)
        End If

    End Sub

    Private Sub chkUnemployment_Click(sender As Object, e As EventArgs) Handles chkUnemployment.Click
        'turn off property
        If chkUnemployment.CheckState = CheckState.Indeterminate Then Exit Sub
        If chkUnemployment.Checked = True Then
            GetIUI()
        Else
            txtUnemploymentInsurance.Text = ""
        End If

        If IsElmore Or IsBoyd Or UseAmericanNationalInsurance Or UseThorntonsInsurance Or IsLott Then  ' Or IsLott Or IsMidSouth
            Recalculate()
        Else
            txtFinanceAmount.Text = CurrencyFormat(GetPrice(txtSubTotal.Text) + GetPrice(txtDocFee.Text) + GetPrice(txtLifeInsurance.Text) + GetPrice(txtAccidentInsurance.Text) + GetPrice(txtPropertyInsurance.Text) + GetPrice(txtUnemploymentInsurance.Text))
            If Months = 0 Then
                FinanceCharge = 0
            Else
                FinanceCharge = ((txtFinanceAmount.Text * InterestRate) / 12 * Months)
            End If
            txtFinanceCharges.Text = CurrencyFormat(FinanceCharge)
            If Months = 0 Then

                Payment = 0
            Else
                Payment = (GetPrice(txtFinanceAmount.Text) + FinanceCharge) / Months
            End If
            txtPaymentWillBe.Text = CurrencyFormat(Payment)
        End If

    End Sub

    Private Function PrintOut_Alt2(Optional ByVal X As Single = -1, Optional ByVal Y As Single = -1, Optional ByVal Text As String = "" _
   , Optional ByVal XCenter As Boolean = False _
  , Optional ByVal FontName As String = "", Optional ByVal FontBold As Boolean = False, Optional ByVal FontSize As String = "" _
  , Optional ByVal DrawWidth As Integer = -1, Optional ByVal NewPage As Boolean = False, Optional ByVal BlankLines As Integer = -1 _
  , Optional ByVal Orientation As Integer = -1, Optional ByVal OutObj As Object = Nothing)

        Dim X_SCALE As Double, Y_SCALE As Double
        X_SCALE = Printer.ScaleWidth / 12240 * 0.946153846153846
        Y_SCALE = Printer.ScaleHeight / 15840 * 0.974093264248705

        X = X * X_SCALE
        Y = Y * Y_SCALE
        PrintOut(X, Y, Text, XCenter, FontName, FontBold, FontSize, DrawWidth, NewPage, BlankLines, Orientation, OutObj)
    End Function

    Private Sub txtAddlPaymentsMade_TextChanged(sender As Object, e As EventArgs) Handles txtAddlPaymentsMade.TextChanged
        UpdateTotalCaption()
    End Sub

    Private Sub txtBalDueLateCharge_TextChanged(sender As Object, e As EventArgs) Handles txtBalDueLateCharge.TextChanged
        UpdateTotalCaption()
    End Sub

    Private Sub txtFinanceAmount_TextChanged(sender As Object, e As EventArgs) Handles txtFinanceAmount.TextChanged
        UpdateTotalCaption()
    End Sub

    'NOTE: COMMENTED BELOW CODE, CAUSE RADIO BUTTON CLICK EVENT OF VB6.0 WILL NOT WORK IN VB.NET. REPLACED IT WITH CHECKEDCHANGED EVENT.
    'Private Sub optLate6_Click(sender As Object, e As EventArgs) Handles optLate6.Click
    '    If IsRevolvingCharge(txtArNo.Text) Then
    '        '    DueOn = RevolvingStatementDay
    '    Else
    '        DueOn = 1
    '    End If
    '    If Not NoAdjust Then AdjustFirstPay(1)
    '    UpdateLateCaptions()
    'End Sub

    Private Sub optLate6_CheckedChanged(sender As Object, e As EventArgs) Handles optLate6.CheckedChanged
        If IsRevolvingCharge(txtArNo.Text) Then
            '    DueOn = RevolvingStatementDay
        Else
            DueOn = 1
        End If
        If Not NoAdjust Then AdjustFirstPay(1)
        UpdateLateCaptions()
    End Sub

    Private Sub WageAssignment() ' New Age & Jeffro Furniture
        Printer.FontName = "Arial"
        Printer.FontSize = 24
        Printer.FontBold = True
        Printer.CurrentX = 0
        Printer.CurrentY = 1000

        PrintCentered("WAGE ASSIGNMENT")

        Printer.FontBold = False
        Printer.FontSize = 10
        Printer.CurrentX = 0

        Printer.FontBold = True
        'Printer.Print Tab(7); BillOSale.dteSaleDate
        Printer.FontBold = False
        Printer.CurrentY = 1600
        Printer.Print("__________________")
        Printer.FontSize = 8
        Printer.Print("             Date")
        Printer.FontSize = 10
        Printer.CurrentY = 1500
        Printer.Print(TAB(7), BillOSale.dteSaleDate.Value)

        ''''''''''''''''''''''''''''''''''
        ' Buyer Information
        ''''''''''''''''''''''''''''''''''

        Printer.FontBold = True
        Printer.CurrentY = 2400 '2350
        Printer.Print("____________________________________")
        Printer.FontSize = 8
        Printer.Print("                           Buyer's Name")
        Printer.FontSize = 10
        Printer.CurrentY = 2300
        Printer.Print(TAB(5), Trim(BillOSale.CustomerFirst.Text), " ", BillOSale.CustomerLast.Text & "   Account: ", ArNo)
        Printer.FontBold = False

        Printer.CurrentY = 3200
        Printer.Print("____________________________________")
        Printer.FontSize = 8
        Printer.Print("                                   Address")
        Printer.FontSize = 10
        Printer.CurrentY = 3100
        Printer.Print(TAB(5), BillOSale.CustomerAddress.Text)

        Printer.CurrentY = 4000 '4000 '3875
        Printer.Print("____________________________________")
        Printer.FontSize = 8
        Printer.Print("     City                       State                   Zip")
        Printer.FontSize = 10
        Printer.Print()
        Printer.CurrentY = 3900
        Printer.Print(TAB(5), Trim(BillOSale.CustomerCity.Text), ", ", BillOSale.CustomerZip.Text)

        ''''''''''''''''''''''''''''''''''
        ' Co-Buyer Information
        ''''''''''''''''''''''''''''''''''

        Dim HasCoBuyer As Boolean, dX As Integer, ddX As Integer
        Dim CoBuyerName As String, CoBuyerAddress As String, CoBuyerCity As String, CoBuyerSS As String
        Dim RS As ADODB.Recordset, SQL As String
        dX = 5000
        ddX = 200

        If BillOSale.MailIndex <> 0 Then
            SQL = "SELECT CoName, CoAddress, CoCityState, CoSS FROM [ArApp] WHERE MailIndex='" & BillOSale.MailIndex & "'"
            RS = GetRecordsetBySQL(SQL, , GetDatabaseAtLocation())
            If RS.RecordCount >= 1 Then
                CoBuyerName = Trim(IfNullThenNilString(RS("CoName")))
                CoBuyerAddress = Trim(IfNullThenNilString(RS("CoAddress")))
                CoBuyerCity = Trim(IfNullThenNilString(RS("CoCityState")))
                CoBuyerSS = Trim(IfNullThenNilString(RS("CoSS")))
                HasCoBuyer = CoBuyerName <> "" Or CoBuyerSS <> ""
                RS = Nothing
            End If
        End If

        If HasCoBuyer Then
            Printer.FontBold = True
            Printer.CurrentY = 2400 '2350
            Printer.CurrentX = dX
            Printer.Print("____________________________________")
            Printer.FontSize = 8
            Printer.CurrentX = dX
            Printer.Print("                         Co-Buyer's Name")
            Printer.FontSize = 10
            Printer.CurrentY = 2300
            Printer.CurrentX = dX + ddX
            Printer.Print(CoBuyerName)
            Printer.FontBold = False

            Printer.CurrentY = 3200
            Printer.CurrentX = dX
            Printer.Print("____________________________________")
            Printer.FontSize = 8
            Printer.CurrentX = dX
            Printer.Print("                                   Address")
            Printer.FontSize = 10
            Printer.CurrentY = 3100
            Printer.CurrentX = dX + ddX
            Printer.Print(CoBuyerAddress)

            Printer.CurrentY = 4000 '4000 '3875
            Printer.CurrentX = dX
            Printer.Print("____________________________________")
            Printer.FontSize = 8
            Printer.CurrentX = dX
            Printer.Print("     City                       State                   Zip")
            Printer.FontSize = 10
            Printer.Print()
            Printer.CurrentY = 3900
            Printer.CurrentX = dX + ddX
            Printer.Print(Trim(CoBuyerCity))
        End If


        ''''''''''''''''''''''''''''''''''
        ' Seller Information
        ''''''''''''''''''''''''''''''''''

        Printer.Print()
        Printer.CurrentY = 4800
        Printer.Print("____________________________________")
        Printer.FontSize = 8
        Printer.Print("                     Seller (Assignee)")
        Printer.FontSize = 10
        Printer.Print()
        Printer.CurrentY = 4700
        Printer.Print(TAB(5), StoreSettings.Name)

        Printer.CurrentY = 5600
        Printer.Print("____________________________________")
        Printer.FontSize = 8
        Printer.Print("                     Seller's Address")
        Printer.FontSize = 10
        Printer.Print()
        Printer.CurrentY = 5500
        Printer.Print(TAB(5), StoreSettings.Address)

        Printer.CurrentY = 6200
        Printer.Print("____________________________________")
        Printer.FontSize = 8
        Printer.Print("     City                       State                   Zip")
        Printer.FontSize = 10
        Printer.Print()
        Printer.CurrentY = 6100
        Printer.Print(TAB(5), StoreSettings.City)
        Printer.Print()

        Printer.FontName = "Garamond"
        Printer.Print()
        Printer.Print()
        Printer.FontSize = 8
        Printer.Print("Amount of Debt:  ")

        Printer.FontBold = True
        Printer.FontSize = 12
        Printer.Print(Format(NewBalance + FinanceCharge, "$###,##0.00"))

        Printer.FontBold = False
        Printer.FontSize = 8
        Printer.Print("          payable in successive installments of:  ")

        Printer.FontBold = True
        Printer.FontSize = 12
        Printer.Print(Format(txtPaymentWillBe, "$##,##0.00"))

        Printer.FontBold = False
        Printer.FontSize = 8
        Printer.Print("      each, beginning:   ")

        Printer.FontBold = True
        Printer.FontSize = 12
        Printer.Print(FirstPayment, " .    ")


        Printer.FontBold = False
        Printer.FontSize = 8
        Printer.Print("       If Default be")
        Printer.Print()

        Printer.Print("made in the payment of any said installments, then all unpaid installments shall, at the assignee's option, become immediately due and payable without notice or demand.")
        Printer.Print("Time Price Differential (Finance Charge): ", "       .")
        Printer.Print("    As security of the above-described debt, which is the same time balance (Total of Payments) due on a retail installment contract, each of the undersigned hereby assigns, transfers and")
        Printer.Print("sets over the above-named assignee, wages, salary, commissions and bonuses due or subsequently earned from his present employer for a period of (3) years from the date")
        Printer.Print("hereof and from any future employer within a period of two (2) years from the date of execution hereof.  Any undersigned Debtor may revoke his assignment of wages by written")
        Printer.Print("notification to the holder.  This assignment shall remain effective as to all the undersigned Debtors not electing to revoke their assignments.")
        Printer.Print("    The amount may be collected by the assignee hereon shall not exceed the lesser of (1) 15% of the gross amount paid assignor for any week, or 2) the amount which disposable")
        Printer.Print("earnings for a week exceed forty-five times the Federal Minimum Hourly Wage in effect at the time the amounts are payable; and shall be collected until the total amount due under this")
        Printer.Print("assignment is paid or until expiration of employer's payroll period ending immediately prior to 84 days after service of the demand hereon, whichever first occurs.  This Wage Assignment")
        Printer.Print("shall be valid for a period of three years from the date hereof.")
        Printer.Print("    The term disposable earnings means that the part of the earnings remaining after deduction of any amount required by law to be withheld.")
        Printer.Print("    The assignors) hereby authorize, empower and direct his/their said employer(s) to pay to assign any and all monies due or to become due assignor(s) hereon, authorize assignee")
        Printer.Print("to receipt for the same and release and discharge employer from all liabilities to assignor(s) on account of monies paid in accordance herewith.  No copy hereof shall be served on")
        Printer.Print("employer(s) except in conformity with applicable law.")
        Printer.Print()
        Printer.Print("Each assignor acknowledges receipt of an exact copy of this Wage Assignment.")

        Printer.FontSize = 24
        Printer.FontBold = True
        Printer.FontName = "Arial"

        Printer.Print()
        PrintCentered("WAGE ASSIGNMENT")
        Printer.FontBold = False

        Printer.FontSize = 12
        Printer.CurrentX = 0
        Printer.Print()
        Printer.Print("___________________________________ _________________________ ___________________________")
        Printer.Print("                  Present Employer                                            S/S                                          Assignor")
        Printer.Print()
        Printer.Print("___________________________________ _________________________ ___________________________")
        Printer.Print("                  Present Employer                                            S/S                                           Assignor")

        If HasCoBuyer Then
            Printer.Print()
            Printer.Print()
            Printer.Print("___________________________________ _________________________ ___________________________")
            Printer.Print("                  Present Employer                                            S/S                                           Co-Buyer")
        End If

        Printer.EndDoc()
    End Sub

    Private Function PrintOut_Alt(Optional ByVal X As Single = -1, Optional ByVal Y As Single = -1, Optional ByVal Text As String = "" _
   , Optional ByVal XCenter As Boolean = False _
  , Optional ByVal FontName As String = "", Optional ByVal FontBold As Boolean = False, Optional ByVal FontSize As String = "" _
  , Optional ByVal DrawWidth As Integer = -1, Optional ByVal NewPage As Boolean = False, Optional ByVal BlankLines As Integer = -1 _
  , Optional ByVal Orientation As Integer = -1, Optional ByVal OutObj As Object = Nothing)

#If True Then
        Const X_OFFSET As Single = 0#
        Const Y_OFFSET As Single = 0#
        Const X_SCALE As Double = 0.946153846153846
        Const Y_SCALE As Double = 0.974093264248705
#Else
  Const X_OFFSET As Single = 0#
  Const Y_OFFSET As Single = 0#
  Const X_SCALE As Double = 1#
  Const Y_SCALE As Double = 1#
#End If


        X = X * X_SCALE + X_OFFSET
        Y = Y * Y_SCALE + Y_OFFSET
        PrintOut(X, Y, Text, XCenter, FontName, FontBold, FontSize, DrawWidth, NewPage, BlankLines, Orientation, OutObj)
    End Function

    Public Sub LoadAdjustmentContract(ByVal vSaleNo As String, ByVal BalanceDue As Decimal)
        Dim I As cInstallment, T As cTransaction

        AddOn = ArAddOn_Add

        SetDefaultsInstallment()
        dteDate1.Value = Today
        'dteDate1_Change
        dteDate1_ValueChanged(dteDate1, New EventArgs)

        T = New cTransaction
        T.DataAccess.DataBase = GetDatabaseAtLocation()
        T.DataAccess.Records_OpenSQL("SELECT * FROM [Transactions] WHERE [Type]='NewSale " & vSaleNo & "'")
        If T.DataAccess.Records_Available Then
            txtArNo.Text = T.ArNo
        Else
            MessageBox.Show("Could not load contract!", "Account Not Found", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            DisposeDA(T)
            Exit Sub
        End If
        ArNo = Val(txtArNo)
        ReprintMailIndex = Val(T.MailIndex)

        I = New cInstallment
        I.Load(txtArNo.Text, "ArNo")

        '  dteDate2 = Date
        '  FirstPayment = FirstPay
        '  Select Case Val(PaidBy)
        '    Case 10:   optLate16 = True
        '    Case 20:   optLate26 = True
        '    Case Else: optLate6 = True
        '  End Select
        txtMonthsToFinance.Text = I.Months
        '  vMonths = Val(txtMonthsToFinance)
        txtRate.Text = I.APR

#If False Then
  txtPrevBalance = CurrencyFormat(PrevBal)
  txtOrigDeposit = CurrencyFormat(OrigDeposit)
  txtSubTotal = CurrencyFormat(GetPrice(PrevBal) + GetPrice(GrossSale) - GetPrice(OrigDeposit))
#End If

        txtGrossSale.Text = BalanceDue

        Recalculate()

        DisposeDA(T, I)


        '  'SubTotal = GetPrice(txtSubTotal)
        '
        '  txtDocFee = CurrencyFormat(DocFee)
        '  DocFee = GetPrice(txtDocFee)
        '
        '  chkLife = IIf(Life <> 0, 1, 0)
        '  txtLifeInsurance = CurrencyFormat(Life)
        '  chkAccident = IIf(Accident <> 0, 1, 0)
        '  txtAccidentInsurance = CurrencyFormat(Accident)
        '  chkProperty = IIf(Property <> 0, 1, 0)
        '  txtPropertyInsurance = CurrencyFormat(Property)
        '  chkUnemployment = IIf(IUI <> 0, 1, 0)
        '  txtUnemploymentInsurance = CurrencyFormat(IUI)
        '
        '  txtFinanceAmount = CurrencyFormat(GetPrice(txtSubTotal) + DocFee + Life + Accident + Property + IUI)
        '  NewBalance = GetPrice(txtFinanceAmount)
        '  txtFinanceCharges = CurrencyFormat(InterestCharged)
        '  txtTotalBalance = CurrencyFormat(GetPrice(txtFinanceAmount) + GetPrice(txtFinanceCharges))
        '
        '  Payment = GetPrice(txtTotalBalance) / Months
        '  LastPay = GetPrice(txtTotalBalance) - Payment * (Months - 1)
        '  CalculateLateCharge
        '
        '
    End Sub

    Public Sub LoadReverseContract(
    ByVal vArNo As String,
    ByVal SaleNo As String, MailIndex As Integer,
    ByVal Delivery As String, ByVal FirstPay As String, ByVal PaidBy As String,
    ByVal vMonths As Integer, ByVal vAPR As String, ByVal vPerMonth As Decimal,
    ByVal PrevBal As Decimal, ByVal GrossSale As Decimal, ByVal OrigDeposit As Decimal,
    ByVal DocFee As Decimal, ByVal Life As Decimal, ByVal Accident As Decimal, ByVal Property1 As Decimal, ByVal IUI As Decimal,
    ByVal InterestCharged As Decimal)

        txtArNo.Text = vArNo
        ArNo = vArNo

        ReprintMailIndex = MailIndex
        ReprintSaleNo = SaleNo

        dteDate1.Value = Delivery
        dteDate2.Value = FirstPay
        FirstPayment = FirstPay
        Select Case Val(PaidBy)
            Case 10 : optLate16.Checked = True
            Case 20 : optLate26.Checked = True
            Case Else : optLate6.Checked = True
        End Select
        txtMonthsToFinance.Text = vMonths
        vMonths = Val(txtMonthsToFinance.Text)
        txtRate.Text = vAPR
        APR = vAPR

        txtPrevBalance.Text = CurrencyFormat(PrevBal)
        txtGrossSale.Text = CurrencyFormat(GrossSale)
        txtOrigDeposit.Text = CurrencyFormat(OrigDeposit)
        txtSubTotal.Text = CurrencyFormat(GetPrice(PrevBal) + GetPrice(GrossSale) - GetPrice(OrigDeposit))
        'SubTotal = GetPrice(txtSubTotal)

        txtDocFee.Text = CurrencyFormat(DocFee)
        DocFee = GetPrice(txtDocFee.Text)

        chkLife = IIf(Life <> 0, 1, 0)
        txtLifeInsurance.Text = CurrencyFormat(Life)
        chkAccident = IIf(Accident <> 0, 1, 0)
        txtAccidentInsurance.Text = CurrencyFormat(Accident)
        chkProperty.Checked = IIf(Property1 <> 0, True, False)
        txtPropertyInsurance.Text = CurrencyFormat(Property1)
        chkUnemployment.Checked = IIf(IUI <> 0, True, False)
        txtUnemploymentInsurance.Text = CurrencyFormat(IUI)

        txtFinanceAmount.Text = CurrencyFormat(GetPrice(txtSubTotal.Text) + DocFee + Life + Accident + Property1 + IUI)
        NewBalance = GetPrice(txtFinanceAmount.Text)
        txtFinanceCharges.Text = CurrencyFormat(InterestCharged)
        txtTotalBalance.Text = CurrencyFormat(GetPrice(txtFinanceAmount.Text) + GetPrice(txtFinanceCharges.Text))

        Payment = GetPrice(txtTotalBalance.Text) / Months
        If Payment <> vPerMonth Then Payment = vPerMonth
        LastPay = GetPrice(txtTotalBalance.Text) - Payment * (Months - 1)
        CalculateLateCharge()
    End Sub

    Public Function DeveloperEx() As String
        Dim S As String
        S = ""
        S = S & "DCk Mthly Pmt" & vbCrLf
        S = S & "  Pmt BrkDn" & vbCrLf
        DeveloperEx = S
    End Function
End Class
