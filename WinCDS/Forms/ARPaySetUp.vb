﻿Imports Microsoft.VisualBasic.Interaction
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
                    B.Intr = B.Intr + Delta

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

        dteDate1.Value = DateFormat(Now)
        AdjustFirstPay

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

End Class
