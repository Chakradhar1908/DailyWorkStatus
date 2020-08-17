Public Class ArCard
    Private Const FRM_W as integer = 10525
    Private Const FRM_H_MIN as integer = 7290
    Private Const FRM_H_MAX as integer = 13000

    Private WithEvents mDBAccess As CDbAccessGeneral
    Private WithEvents mDBAccessTransactions As CDbAccessGeneral

    Public PayCount as integer
    Private PayLog() As PaymentRecord ' Array for tracking payments, for receipts

    Private altPayType As String    ' used for a couple specialty operations like payoffs
    Public ArNo As String
    Public Status As String
    Public MailRec as integer
    Public DocCredit As Decimal
    Public LifeCredit As Decimal
    Public AccidentCredit As Decimal
    Public PropertyCredit As Decimal
    Public IUICredit As Decimal
    Public InterestCredit As Decimal
    Public InterestTaxCredit As Decimal

    ' These need to be eliminated, to speed up the whole program.
    'GDM Form variable to indicate that we are processing a move button
    Public UsingMoveButton As Boolean
    Public f_strDirection As String 'GDM

    Private Const NON_ALERT_COLOR as integer = &H6666CC
    Private Const ALERT_COLOR as integer = &H6666CC ' '&HC0&

    Dim Mail As MailNew
    Dim Mail2 As MailNew2

    Dim mArNo As String
    Dim CustRec as integer
    Dim CashOpt as integer

    Public INTEREST As Decimal
    Public InterestTax As Decimal
    Public DocFee As Decimal
    Public Life As Decimal
    Public Accident As Decimal
    Public Prop As Decimal
    Public IUI As Decimal

    Dim Charges As Decimal
    Dim Credits As Decimal
    Dim Balance As Decimal
    Dim TotPaid As Decimal
    Dim Financed As Decimal
    Dim TransType As String
    Dim Payoff As String
    Dim PayoffSameAsCash As Boolean
    Dim StatusChg As String
    Dim Receipt As String
    Dim NewTypee As String

    Dim TransDate As String
    Dim LastPayDate As String
    Dim LastPay As String
    Dim LateChargeBal As String
    Dim Months As String
    Dim SendNotice As String
    Dim Counter as integer

    Dim Approval As String

    Dim OpenFormAs As String
    Dim InterestDebit As Decimal  ' For Bankruptcy transactions
    Dim InterestCreditRevolving As Decimal

    Dim DoRecordAccountClosed As Boolean

    Dim cmdApplyValue As Boolean                 ' Used to determine whether this button has been clicked.
    Dim cmdReceiptValue As Boolean               ' Future Languages do not support command button value property

    Public Sub ShowArCardForDisplayOnly(ByVal nArNo As String, Optional ByVal Modal As Boolean = True, Optional ByVal AllowClose As Boolean = False, Optional ByVal AllowContractChange As Boolean = False)
        Dim OldAR As String
        If nArNo <> "" Then LoadArNo(nArNo)

        cmdApply.Enabled = AllowContractChange
        'cmdCreditApp.Enabled = False
        'cmdDetail.Enabled = False
        cmdEdit.Enabled = False
        cmdExport.Enabled = False
        cmdFields.Enabled = False
        cmdMakeSameAsCash.Enabled = False
        cmdMoveFirst.Enabled = False
        cmdMoveLast.Enabled = False
        cmdMoveNext.Enabled = False
        cmdMovePrevious.Enabled = False
        cmdPayoff.Enabled = False
        cmdReceipt.Enabled = False
        cmdReprintContract.Enabled = False
        'Notes_Open.Enabled = False
        cmdPrint.Enabled = False
        'cmdPrintCard.Enabled = False
        'cmdPrintLabel.Enabled = False
        cmdCancel.Enabled = AllowClose

        fraPaymentOptions.Visible = False
        fraEditOptions.Visible = False

        'Show IIf(Modal, 1, 0)
        If Modal = True Then
            Me.ShowDialog()
        Else
            Me.Show()
        End If

        If OldAR <> "" Then ArSelect = OldAR

        '  DisposeDA RS
    End Sub

    Public Function LoadArNo(Optional ByVal vArNo As String = "") As Boolean
        If IsRevolvingCharge(vArNo) Then DoRevolvingProcessAccount(Today, vArNo, False) ' in case interest is due..  applies interest, no statement until statement date

        mDBAccess_Init(vArNo)
        mDBAccess.GetRecord()    ' this gets the record
        mDBAccess.dbClose()
        mDBAccess = Nothing

        If mArNo <> "-1" Then 'not found
            GetCustomer()
            mDBAccessTransactions_Init(mArNo)
            mDBAccessTransactions.GetRecord()    ' this gets the record
            mDBAccessTransactions.dbClose()
            mDBAccessTransactions = Nothing
            GetPayoff
            GetAgeing
            filFile_Click
        End If

        txtPaymentHistory.Text = WrapLongText(GetArCreditHistory(mArNo, Today, 24), 12) ' GetPaymentHistorySimple(vArNo, Date, 24)
    End Function

    Private Sub mDBAccess_Init(ByVal Tid As String)
        Dim SQL As String

        mDBAccess = New CDbAccessGeneral
        mDBAccess.dbOpen(GetDatabaseAtLocation())
        SQL = ""
        SQL = SQL & "SELECT InstallmentInfo.*, ArApp.HisSS, ArApp.ApprovalTerms, ArApp.CreditLimit"
        SQL = SQL & " FROM InstallmentInfo LEFT JOIN ArApp ON (InstallmentInfo.MailIndex=Val(iif(isnull(ArApp.MailIndex),0,ArApp.MailIndex)))"
        SQL = SQL & " WHERE (Status<>'V')"  'BFH20080215 - ALLOW OPENING VOIDS || 20080305 - removed again
        '    SQL = SQL & " WHERE Status<>'V'"
        ' GDM BEGINNING OF CHANGE 3/29/2001
        ' Change the sql string based on the move option (f_strDirection)
        If UsingMoveButton Then
            Select Case f_strDirection
                Case "First"    'Get first record in Table
                    mDBAccess.SQL = SQL & " ORDER BY InstallmentInfo.ArNo"
                    Exit Sub
                Case "Last"     'Get current record and all records beyond
                    mDBAccess.SQL = SQL & " AND InstallmentInfo.ArNo  >=""" & ProtectSQL(Tid) & """ ORDER BY InstallmentInfo.ArNo"
                    Exit Sub
                Case "Previous" 'Get all records up to and including current record
                    mDBAccess.SQL = SQL & " AND InstallmentInfo.ArNo  <=""" & ProtectSQL(Tid) & """ ORDER BY InstallmentInfo.ArNo"
                    Exit Sub
                Case "Next"     'Get current record and next record only
                    mDBAccess.SQL = SQL & " AND InstallmentInfo.ArNo  >=""" & ProtectSQL(Tid) & """ ORDER BY InstallmentInfo.ArNo"
                    Exit Sub
            End Select
        End If
        ' GDM END OF CHANGE
        mDBAccess.SQL = SQL & " AND InstallmentInfo.ArNo  =""" & ProtectSQL(Tid) & """"
    End Sub

    Public Sub GetCustomer()
        'MousePointer = 11
        Me.Cursor = Cursors.WaitCursor

        On Error GoTo HandleErr

        If ARPaySetUp.AccountFound = "Y" Then MailRec = ARPaySetUp.MailRec

        Dim RS As ADODB.Recordset
        RS = getRecordsetByTableLabelIndexNumber("Mail", "Index", CStr(MailRec))
        If (RS.RecordCount <> 0) Then CopyMailRecordsetToMailNew(RS, Mail)
        RS.Close()
        RS = Nothing

        CopyMailRecordsetToMailNew2(Nothing, Mail2)
        RS = getRecordsetByTableLabelIndexNumber("MailShipTo", "Index", CStr(Mail.Index))
        If (RS.RecordCount <> 0) Then CopyMailRecordsetToMailNew2(RS, Mail2)
        RS.Close()
        RS = Nothing

        GetCust
        ClearPayments()
        'MousePointer = 0
        Me.Cursor = Cursors.Default
        Exit Sub

        'Does Not Find Customer
        If MsgBox("Name Not In Data Base:  Try Again?", vbYesNo + vbExclamation) = vbYes Then
            ' Retry
            Exit Sub
        End If
        Exit Sub

HandleErr:
        If Err.Number = 53 Then
            CustRec = "999"
            Resume Next
        ElseIf Err.Number = 52 Then
            Resume Next
        End If
    End Sub

    Private Sub ClearPayments()
        'ReDim PayLog(0)
        Erase PayLog
        PayCount = 0
    End Sub

    Private Sub mDBAccessTransactions_Init(Tid As String)
        mDBAccessTransactions = New CDbAccessGeneral
        mDBAccessTransactions.dbOpen(GetDatabaseAtLocation())
        mDBAccessTransactions.SQL =
            "SELECT Transactions.*" _
            & " From Transactions" _
            & " WHERE (((Transactions.ArNo)=""" & ProtectSQL(mArNo) & """))" _
            & " ORDER BY  Transactions.ArNo, Transactions.transactionId"
    End Sub

    Public Sub GetPayoff(Optional ByVal AsOfDate As Date = NullDate)
        Dim ppoLife As Boolean, ppoAcc As Boolean, ppoProp As Boolean, ppoIUI As Boolean
        Dim C As cArTreehouse, N As Integer, RN As Integer, A As Double, B As Double
        Dim LA As String, SAC As Boolean

        CheckNullDate(AsOfDate, PayoffAsOfDate)

        If IsRevolvingCharge(ArNo) Then
            Dim InstAcct As New cInstallment
            InstAcct.Load(ArNo)
            lblTotalPayoff.Text = CurrencyFormat(InstAcct.GetPayoffRevolving(AsOfDate))
            DisposeDA(InstAcct)
            Exit Sub
        End If

        InterestCredit = 0
        InterestTaxCredit = 0
        DocCredit = 0
        LifeCredit = 0
        AccidentCredit = 0
        PropertyCredit = 0
        IUICredit = 0

        GetPreviousPayoff(ArNo, ppoLife, ppoAcc, ppoProp, ppoIUI)

        If ARPaySetUp.AccountFound = "Y" Then
            INTEREST = ARPaySetUp.INTEREST
            Life = ARPaySetUp.Life
            Accident = ARPaySetUp.Accident
            Prop = ARPaySetUp.Prop
            IUI = ARPaySetUp.IUI
            InterestTax = ARPaySetUp.InterestTax
        End If

        If AlreadyMadeSameAsCash() Then
            DocCredit = 0
            LifeCredit = 0
            AccidentCredit = 0
            PropertyCredit = 0
            IUICredit = 0
            InterestCredit = 0
            InterestTaxCredit = 0
            lblTotalPayoff.Text = CurrencyFormat(GetPrice(lblBalance.Text))
            If AddOnAcc.Typee = ArAddOn_New Then ARPaySetUp.txtPrevBalance.Text = Format((GetPrice(lblBalance.Text)) - (LifeCredit) - (AccidentCredit) - (PropertyCredit) - (IUICredit) - (InterestCredit) - (InterestTaxCredit), "###,###.00")
            Exit Sub
        End If

        If CashOptPaidOff(True) And Status = "O" And Not AlreadyMadeSameAsCash() Then
            SAC = True    ' same as cash, for interest and interest tax
            If (Not PayInsuranceAfter30Days() Or DateDiff("d", DateValue(txtDelivery.Text), Today) <= 30) Then
                LifeCredit = Life
                AccidentCredit = Accident
                PropertyCredit = Prop
                IUICredit = IUI
                InterestCredit = INTEREST
                InterestTaxCredit = InterestTax
                lblTotalPayoff.Text = CurrencyFormat((GetPrice(lblBalance.Text)) - (LifeCredit) - (AccidentCredit) - (PropertyCredit) - (IUICredit) - (InterestCredit) - (InterestTaxCredit))
                If AddOnAcc.Typee = ArAddOn_New Then ARPaySetUp.txtPrevBalance.Text = Format((GetPrice(lblBalance.Text)) - (LifeCredit) - (AccidentCredit) - (PropertyCredit) - (IUICredit) - (InterestCredit) - (InterestTaxCredit), "###,###.00")
                Exit Sub
            End If
        End If


        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

        '  If Val(DocFee) > 0 And getprice(lblBalance) > 0 Then
        '    DocCredit = ProRata(DocFee, Val(txtMonths), txtLastPay)
        '  End If

        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

        If Val(Life) > 0 And GetPrice(lblBalance.Text) > 0 Then
            Select Case LifePayoffMethod()
                Case ArPayoffMethod_ProRata : LifeCredit = ProRata(Life, Val(txtMonths.Text), txtLastPay.Text, AsOfDate)
                Case ArPayoffMethod_Rule_78 : LifeCredit = Rule78(DateValue(txtDelivery.Text), GetPrice(Life), Val(txtMonths.Text), , AsOfDate)
                Case ArPayoffMethod_Anticip
                    N = Val(txtMonths)
                    RN = CountMonths(AsOfDate, DateValue(txtLastPay.Text), True)
                    C = New cArTreehouse
                    A = C.LifeRate(N, False)
                    B = C.LifeRate(RN, False)
                    C = Nothing
                    LifeCredit = RuleOfAnticipationForTreehouse(GetPrice(txtFinanced.Text), GetPrice(txtMonthlyPayment.Text), Life, N, RN, A, B)

                    If txtPayMemo.Text = "*" And IsDevelopment() Then
                        LA = ""
                        LA = LA & "Orig Loan: " & txtFinanced.Text & vbCrLf
                        LA = LA & "Month Pmt: " & txtMonthlyPayment.Text & vbCrLf
                        LA = LA & "Orig Prem: " & Life & vbCrLf
                        LA = LA & "Orig Term: " & N & vbCrLf
                        LA = LA & "Remain Tm: " & RN & vbCrLf
                        LA = LA & "Orig Rate: " & A & vbCrLf
                        LA = LA & "Remain Rt: " & B & vbCrLf
                        LA = LA & vbCrLf
                        LA = LA & "LIFE CREDIT = " & LifeCredit & vbCrLf
                        'MsgBox LA
                        MessageBox.Show(LA)
                    End If

                Case Else : If IsDevelopment() Then MessageBox.Show("Unknown LifePayoffMethod: " & LifePayoffMethod())
            End Select
        End If
        If LifeCredit > Life Then LifeCredit = Life
        If MinLifeCredit() > 0 And GetPrice(LifeCredit) < MinLifeCredit() Then LifeCredit = 0  ' minimum refund amount
        If ppoLife Then LifeCredit = 0

        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        If Val(Accident) > 0 And GetPrice(lblBalance.Text) > 0 Then
            Select Case AccPayoffMethod()
                Case ArPayoffMethod_ProRata : AccidentCredit = ProRata(Accident, Val(txtMonths), txtLastPay.Text, AsOfDate)
                Case ArPayoffMethod_Rule_78 : AccidentCredit = Rule78(DateValue(txtDelivery.Text), Accident, Val(txtMonths.Text), , AsOfDate)
                Case ArPayoffMethod_Rule78b : AccidentCredit = Rule78(DateValue(txtDelivery.Text), Accident, Val(txtMonths.Text), True, AsOfDate)
                Case ArPayoffMethod_Anticip
                    N = Val(txtMonths)
                    RN = CountMonths(AsOfDate, DateValue(txtLastPay.Text), True)
                    C = New cArTreehouse
                    A = C.DisabilityRate(N)
                    B = C.DisabilityRate(RN)
                    C = Nothing
                    AccidentCredit = RuleOfAnticipationForTreehouse(GetPrice(txtFinanced.Text), GetPrice(txtMonthlyPayment.Text), Accident, N, RN, A, B)

                    If txtPayMemo.Text = "*" And IsDevelopment() Then
                        LA = ""
                        LA = LA & "Orig Loan: " & txtFinanced.Text & vbCrLf
                        LA = LA & "Month Pmt: " & txtMonthlyPayment.Text & vbCrLf
                        LA = LA & "Orig Prem: " & Accident & vbCrLf
                        LA = LA & "Orig Term: " & N & vbCrLf
                        LA = LA & "Remain Tm: " & RN & vbCrLf
                        LA = LA & "Orig Rate: " & A & vbCrLf
                        LA = LA & "Remain Rt: " & B & vbCrLf
                        LA = LA & vbCrLf
                        LA = LA & "ACC CREDIT = " & AccidentCredit & vbCrLf
                        'MsgBox LA
                        MessageBox.Show(LA)
                    End If

                Case Else : If IsDevelopment() Then MessageBox.Show("Unknown AccPayoffMethod: " & AccPayoffMethod())
            End Select
        End If
        If AccidentCredit > Accident Then AccidentCredit = Accident
        If MinAccCredit() > 0 And GetPrice(AccidentCredit) < MinAccCredit() Then AccidentCredit = 0  ' minimum refund amount
        If ppoAcc Then AccidentCredit = 0

        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        If Val(Prop) > 0 And GetPrice(lblBalance.Text) > 0 Then
            Select Case PropPayoffMethod()
                Case ArPayoffMethod_ProRata : PropertyCredit = ProRata(Prop, Val(txtMonths.Text), txtLastPay.Text, AsOfDate)
                Case ArPayoffMethod_Rule_78 : PropertyCredit = Rule78(DateValue(txtDelivery.Text), Prop, Val(txtMonths.Text), , AsOfDate)
                Case ArPayoffMethod_Rule78b : PropertyCredit = Rule78(DateValue(txtDelivery.Text), Prop, Val(txtMonths.Text), True, AsOfDate)
                Case Else : If IsDevelopment() Then MessageBox.Show("Unknown PropPayoffMethod: " & PropPayoffMethod())
            End Select
        End If
        If PropertyCredit > Prop Then PropertyCredit = Prop
        If MinPropCredit() > 0 And GetPrice(PropertyCredit) < MinPropCredit() Then PropertyCredit = 0  ' minimum refund amount
        If ppoProp Then PropertyCredit = 0

        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        If Val(IUI) > 0 And GetPrice(lblBalance.Text) > 0 Then
            Select Case IUIPayoffMethod()
                Case ArPayoffMethod_ProRata : IUICredit = ProRata(IUI, Val(txtMonths.Text), txtLastPay.Text, AsOfDate)
                Case ArPayoffMethod_Rule_78 : IUICredit = Rule78(DateValue(txtDelivery.Text), IUI, Val(txtMonths.Text), , AsOfDate)
                Case ArPayoffMethod_Rule78b : IUICredit = Rule78(DateValue(txtDelivery.Text), IUI, Val(txtMonths.Text), True, AsOfDate)
                Case Else : If IsDevelopment() Then MessageBox.Show("Unknown IUIPayoffMethod: " & IUIPayoffMethod())
            End Select
        End If
        If IUICredit > IUI Then IUICredit = IUI
        If MinIUICredit() > 0 And GetPrice(IUICredit) < MinIUICredit() Then IUICredit = 0  ' minimum refund amount
        If ppoIUI Then IUICredit = 0

        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        If CLng(GetPrice(lblBalance.Text)) <> 0 And GetPrice(lblBalance.Text) > 0 Then  'added 12-18-2002 for closed accounts
            If Val(INTEREST) >= 0 Then
                InterestCredit = Rule78(DateValue(txtDelivery.Text), GetPrice(INTEREST), Val(txtMonths.Text), , AsOfDate)
            End If
        End If

        If IsElmore And GetPrice(txtRate.Text) <> 0 Then
            '10% of Balance
            If (lblBalance.Text * 0.1) > 25.0# Then
                InterestCredit = InterestCredit - 25.0#
            Else
                InterestCredit = (lblBalance.Text * 0.1)
            End If
        End If
        If SAC Then InterestCredit = INTEREST
        If InterestCredit > INTEREST Then InterestCredit = INTEREST
        If InterestCredit < 0 Then InterestCredit = 0   'no negative numbers

        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        If CLng(GetPrice(lblBalance.Text)) <> 0 Then
            If Val(InterestTax) >= 0 Then
                InterestTaxCredit = ProRata(InterestTax, Val(txtMonths), txtLastPay.Text, AsOfDate)  ' BFH20080703 - this should probably refund the full amount??
            End If
        End If
        If SAC Then InterestTaxCredit = InterestTax
        If InterestTaxCredit < 0 Then InterestTaxCredit = 0

        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        lblTotalPayoff.Text = CurrencyFormat((GetPrice(lblBalance.Text)) - (LifeCredit) - (AccidentCredit) - (PropertyCredit) - (IUICredit) - (InterestCredit) - (InterestTaxCredit))
        If GetPrice(lblTotalPayoff.Text) < 0 Then lblTotalPayoff.Text = CurrencyFormat(0)

        If AddOnAcc.Typee = ArAddOn_New Then
            ARPaySetUp.txtPrevBalance.Text = Format((GetPrice(lblBalance.Text)) - (LifeCredit) - (AccidentCredit) - (PropertyCredit) - (IUICredit) - (InterestCredit) - (InterestTaxCredit), "###,###.00")
            'BFH20071219 - Negative Previous Balances Allowed...
            '    If ARPaySetUp.txtPrevBalance.Text < 0 Then ARPaySetUp.txtPrevBalance = ".00"
        End If
    End Sub

    Private Function MinIUICredit() As Decimal
        If IsBoyd Then MinIUICredit = 2
        If IsTreehouse Or IsBlueSky Then MinIUICredit = 1
    End Function

    Private Function IUIPayoffMethod() As String
        IUIPayoffMethod = ArPayoffMethod_Rule_78
        If IsTreehouse Or IsBlueSky Then IUIPayoffMethod = ArPayoffMethod_ProRata
    End Function

    Private Function MinPropCredit() As Decimal
        If IsBoyd Then MinPropCredit = 2
        If IsTreehouse Or IsBlueSky Then MinPropCredit = 1
    End Function

    Private Function PropPayoffMethod() As String
        PropPayoffMethod = ArPayoffMethod_Rule_78
        '  If IsMidSouth Then PropPayoffMethod = "Rule 78b" ' Or IsLott
        If IsTreehouse Or IsBlueSky Then PropPayoffMethod = ArPayoffMethod_ProRata
        If UseAmericanNationalInsurance Then PropPayoffMethod = ArPayoffMethod_ProRata
    End Function

    Private Function MinAccCredit() As Decimal
        If IsBoyd Then MinAccCredit = 2
        If IsTreehouse Or IsBlueSky Then MinAccCredit = 1
    End Function

    Private Function AccPayoffMethod() As String
        AccPayoffMethod = ArPayoffMethod_Rule_78
        If IsTreehouse Or IsBlueSky Then AccPayoffMethod = ArPayoffMethod_Anticip
        '  If IsMidSouth Then AccPayoffMethod = "Rule 78b" ' Or IsLott
    End Function

    Private Function MinLifeCredit() As Decimal
        If IsBoyd Then MinLifeCredit = 2
        If IsTreehouse Or IsBlueSky Then MinLifeCredit = 1
    End Function

    Private Function LifePayoffMethod() As String
        LifePayoffMethod = ArPayoffMethod_Rule_78
        '  If IsMidSouth Then LifePayoffMethod = "ProRata" ' Or IsLott
        If IsTreehouse Or IsBlueSky Then LifePayoffMethod = ArPayoffMethod_Anticip
    End Function

    Private Function PayInsuranceAfter30Days() As Boolean
        '  PayInsuranceAfter30Days = IsMidSouth ' Or IsLott
    End Function

    Private Function CashOptPaidOff(Optional ByVal AssumePaid As Boolean = False) As Boolean
        Dim CashOptOKDate As Date, PayOffBal As Decimal
        If Val(txtSameAsCash.Text) = 0 Then Exit Function
        CashOptOKDate = DateAdd("m", Val(txtSameAsCash), DateValue(txtDelivery.Text)) 'txtFirstPay)) ' MJK 20140118
        If DateAfter(Today, CashOptOKDate) Then Exit Function
        If AssumePaid Then CashOptPaidOff = True : Exit Function
        PayOffBal = GetPrice(txtFinanced.Text) - Life - Accident - Prop - IUI - INTEREST - InterestTax
        CashOptPaidOff = PayOffBal <= TotPaid
    End Function

    Public Sub GetAgeing()
        Dim LateChargeBal As Decimal
        LateChargeBal = GetPrice(lblLateCharge.Text)

        Dim AR As Decimal, L0 As Decimal, L30 As Decimal, L60 As Decimal, L90 As Decimal
        Dim NDD As String
        If IsRevolvingCharge(ArNo) Then
            ' Aging needs to be rethought for revolving accounts; amounts due are based on the date of each sale.
            SetAgeingVisible(False)
        Else
            ComputeAgeing(Today, DateValue(txtFirstPay.Text), Val(txtMonths.Text), Val(txtPaidBy.Text), txtPayPeriod.Text = "W",
      GetPrice(txtMonthlyPayment.Text), TotPaid, Financed, Balance, False, False,
      AR, L0, L30, L60, L90, , , , , , , NDD)
            SetAgeingVisible(True)
            txtNextDue.Text = NDD
        End If

        If AR < 0 Then
            '    AR = 0   ' BFH20140227
            lblTotDue.Text = "TotDue: " & CurrencyFormat(AR)
        ElseIf -AR > Balance Then
            lblTotDue.Text = "TotDue: " & CurrencyFormat(AR)
        Else
            lblTotDue.Text = "TotDue: " & CurrencyFormat(AR + LateChargeBal)
        End If
        lblArrearages.Text = "Arrearages: " & CurrencyFormat(AR)

        lblLate0.Text = CurrencyFormat(L0)
        lblLate31.Text = CurrencyFormat(L30)
        lblLate61.Text = CurrencyFormat(L60)
        lblLate91.Text = CurrencyFormat(L90)
        lblBalance.Text = CurrencyFormat(lblBalance)
        lblLateCharge.Text = CurrencyFormat(lblLateCharge)
    End Sub

    Public Sub SetAgeingVisible(ByVal Vis As Boolean)
        lbl0030.Visible = Vis
        lblLate0.Visible = Vis
        lbl3160.Visible = Vis
        lblLate31.Visible = Vis
        lbl6190.Visible = Vis
        lblLate61.Visible = Vis
        lblOver91.Visible = Vis
        lblLate91.Visible = Vis
        lblArrearages.Visible = Vis
        lblTotDue.Visible = Vis
    End Sub

    Private Sub filFile_Click()
        On Error GoTo ClearLetter
        cmdEdit.Enabled = True
        cmdPrint.Enabled = True
        cmdExport.Enabled = True

        rtfFile.LoadFile(filFile.Path & "\" & filFile.FileName)
        ReplaceLetterTokens(rtfFile)

        Exit Sub
ClearLetter:
        rtfFile.SelectionStart = 0
        rtfFile.SelectionLength = Len(rtfFile)
        rtfFile.SelectedText = ""
        rtfFile.Tag = ""

    End Sub

    Public Sub ReplaceLetterTokens(ByRef rtb As RichTextBox)
        Dim L As Object, I As Object, Op As Object
        'Op = MousePointer
        Op = Cursor
        'MousePointer = vbHourglass
        Cursor = Cursors.WaitCursor
        '01-04:  store (name, add, city, phone)
        '05-08:  lblaccount, [txtlastpay], Trim(lblFirstName), Trim(lblLastName)
        '09-14:  lblAddress, city, zip, tele1, tele2, lblSSN (BFH20050516: was ArApp.SS)
        '15:     Format(lblBalance, "$#,##0.00")    'Balance
        '16:     txtPaidBy
        '17:     Format(txtMonthlyPayment, "$#,##0.00")    'Payment (T&C)
        '18:     Format(txtLateChargeAmount, "$#,##0.00")     'Late Charge (T&C)
        '19:     Format(txtFinanced, "$#,##0.00")     'Amt Financed
        '20-22:  txtMonths, LastPayDate, LastPay
        '23:     Format(GetPrice(lblLate31) + GetPrice(lblLate61) + GetPrice(lblLate91), "$#,##0.00")
        '24:     Format(GetPrice(txtMonthlyPayment) + GetPrice(txtLateChargeAmount), "$#,##0.00")  ' Payment + Late Charge
        '25-26:  txtPaidBy, lblArrearages
        '27:     Format(GetPrice(lblLate0) + GetPrice(lblLate31) + GetPrice(lblLate61) + GetPrice(lblLate91), "$#,##0.00") ' Total due
        '28:     dateformat(Now)
        '29:     txtDelivery
        '30-31:  Credit Limit, Approval Terms (both from ArApp, like SSN was)
        '32:     Last Payment Made (Payment History)
        '33:     Last Payment Made Date (Payment History)
        '34:     LastPayment Made Type (Payment History)

        'ReDim L(1 To 34)
        ReDim L(0 To 33)
        L(0) = StoreSettings.Name : L(1) = StoreSettings.Address : L(2) = StoreSettings.City : L(3) = StoreSettings.Phone
        L(4) = lblAccount.Text : L(5) = cboStatus.Text : L(6) = Trim(lblFirstName.Text) : L(7) = Trim(lblLastName.Text)
        L(8) = lblAddress.Text : L(9) = lblCity.Text : L(10) = lblZip.Text : L(11) = lblTele1.Text : L(12) = lblTele2.Text : L(13) = lblSSN.Text
        L(14) = CurrencyFormat(lblBalance.Text, , True)
        L(15) = txtPaidBy.Text
        L(16) = CurrencyFormat(txtMonthlyPayment.Text, , True)
        L(17) = CurrencyFormat(txtLateChargeAmount.Text, , True)
        L(18) = CurrencyFormat(txtFinanced.Text, , True)
        L(19) = txtMonths.Text : L(20) = txtLastPay.Text : L(21) = GetPrice(txtFinanced.Text) - GetPrice(txtMonthlyPayment.Text) * (GetPrice(txtMonths.Text) - 1)
        L(22) = CurrencyFormat(GetPrice(lblLate31.Text) + GetPrice(lblLate61.Text) + GetPrice(lblLate91.Text), , True)
        L(23) = CurrencyFormat(GetPrice(txtMonthlyPayment.Text) + GetPrice(txtLateChargeAmount.Text), , True)
        L(24) = "" : L(25) = Mid(lblArrearages.Text, 13)
        L(26) = CurrencyFormat(GetPrice(lblLate0.Text) + GetPrice(lblLate31.Text) + GetPrice(lblLate61.Text) + GetPrice(lblLate91.Text) + GetPrice(LateChargeBal), , True)
        L(27) = DateFormat(Now)
        L(28) = txtDelivery.Text
        L(29) = lblCreditLimit.Text : L(30) = lblApprovalTerms.Text

        Dim LPAmt As Decimal, LPTyp As String, LPDat As String
        If GetArNoLastPayment(L(5), LPAmt, LPTyp, LPDat) Then
            L(31) = CurrencyFormat(LPAmt)
            L(32) = LPDat
            L(33) = LPTyp
        Else
            L(31) = "[NONE]"
            L(32) = "[NEVER]"
            L(33) = "[N/A]"
        End If

        '  L = Array(frmSetup .StoreName, frmSetup .StoreAddress, frmSetup .StoreCity, frmSetup .StorePhone, _
        '            lblAccount, "", Trim(lblFirstName), Trim(lblLastName), _
        '            lblAddress, lblCity, lblZip, lblTele1, lblTele2, lblSSN, _
        '            Format(lblBalance, "$#,##0.00"), _
        '            txtPaidBy, _
        '            Format(txtMonthlyPayment, "$#,##0.00"), Format(txtLateChargeAmount, "$#,##0.00"), _
        '            Format(txtFinanced, "$#,##0.00"), _
        '            txtMonths, LastPayDate, LastPay, _
        '            Format(GetPrice(lblLate31) + GetPrice(lblLate61) + GetPrice(lblLate91), "$#,##0.00"), _
        '            Format(GetPrice(txtMonthlyPayment) + GetPrice(txtLateChargeAmount), "$#,##0.00"), _
        '            txtPaidBy, lblArrearages, _
        '            Format(GetPrice(lblLate0) + GetPrice(lblLate31) + GetPrice(lblLate61) + GetPrice(lblLate91), "$#,##0.00"), _
        '            DateFormat(Now), _
        '            txtDelivery, _
        '            lblCreditLimit, lblApprovalTerms _
        '            )

        rtb.SelectionLength = 0
        For I = LBound(L) To UBound(L)
            Do While rtb.Find("%" & Format(I, "00"), 1, -1, RichTextBoxFinds.WholeWord) <> -1
                'rtb.SelText = L(I)
                rtb.SelectedText = L(I)
                'rtb.SelLength = 0
                rtb.SelectionLength = 0
            Loop
        Next

        'MousePointer = Op
        Cursor = Op
        Exit Sub
ErrorHandler:
        'MousePointer = Op
        Cursor = Op
        'rtb.SelStart = 0
        rtb.SelectionStart = 0
        'rtb.SelLength = Len(rtb)
        rtb.SelectionLength = Len(rtb)
        'rtb.SelText = ""
        rtb.SelectedText = ""
    End Sub

    Public Sub GetCust()
        'FINDS OLD CUSTOMER & CONTINUES ON
        lblFirstName.Text = Mail.First
        lblLastName.Text = Mail.Last
        lblAddress.Text = Mail.Address
        lblAddAddress.Text = Mail.AddAddress
        lblCity.Text = Mail.City
        lblZip.Text = Mail.Zip
        lblTele1.Text = DressAni(CleanAni(Mail.Tele))
        lblTele2.Text = DressAni(CleanAni(Mail.Tele2))
        lblTele3.Text = DressAni(CleanAni(Mail2.Tele3))
        SetTelephoneCaptions(Mail.PhoneLabel1, Mail.PhoneLabel2, Mail2.PhoneLabel3)

        If ARPaySetUp.AccountFound <> "Y" Then
            lblAccount.Text = Trim(ArNo)
        End If
    End Sub

    Private Sub SetTelephoneCaptions(ByVal Lbl1 As String, ByVal Lbl2 As String, ByVal Lbl3 As String)
        Dim Longest As Integer
        If Trim(Lbl1) = "" Then Lbl1 = "Tele 1: "
        If Trim(Lbl2) = "" Then Lbl2 = "Tele 2: "
        If Trim(Lbl3) = "" Then Lbl3 = "Tele 3: "
        If Microsoft.VisualBasic.Right(Trim(Lbl1), 1) <> ":" Then Lbl1 = Lbl1 & ": "
        If Microsoft.VisualBasic.Right(Trim(Lbl2), 1) <> ":" Then Lbl2 = Lbl2 & ": "
        If Microsoft.VisualBasic.Right(Trim(Lbl3), 1) <> ":" Then Lbl3 = Lbl3 & ": "
        lblTele1Caption.Text = Lbl1
        lblTele2Caption.Text = Lbl2
        lblTele3Caption.Text = Lbl3
        Longest = Max(lblTele1Caption.Width, lblTele2Caption.Width, lblTele3Caption.Width)
        lblTele1.Left = lblTele1Caption.Left + Longest + 60
        lblTele2.Left = lblTele2Caption.Left + Longest + 60
        lblTele3.Left = lblTele3Caption.Left + Longest + 60
    End Sub

    Private ReadOnly Property PayoffAsOfDate() As Date
        Get
            PayoffAsOfDate = Today     ' DEFAULT VALUE

            If OrderMode("A") Then
                If IsFormLoaded("BillOSale") Then
                    If IsDate(BillOSale.dteSaleDate.Value) Then
                        PayoffAsOfDate = DateValue(BillOSale.dteSaleDate.Value)
                    End If
                End If
            End If

            PayoffAsOfDate = DateValue(PayoffAsOfDate)
        End Get
    End Property

    Private Sub GetPreviousPayoff(ByVal ArNo As String, ByRef ppoLife As Boolean, ByRef ppoAcc As Boolean, ByRef ppoProp As Boolean, ByRef ppoIUI As Boolean) ', byRef ppoInt as boolean)
        Dim RS As ADODB.Recordset
        Dim CL As Boolean, CA As Boolean, cP As Boolean, cU As Boolean ' , cI as boolean
        RS = GetRecordsetBySQL("SELECT * FROM [Transactions] WHERE ArNo='" & ArNo & "' AND LCase(Left(Type,4)) IN ('Life','Acc.','Prop','IUI ') ORDER BY [TransactionID] DESC")

        ppoLife = False
        ppoAcc = False
        ppoProp = False
        ppoIUI = False

        Do While Not RS.EOF
            Select Case IfNullThenNilString(RS("Type"))
                Case arPT_Lif : CL = True
                Case arPT_poLif : If Not CL Then ppoLife = True : CL = True

                Case arPT_Acc : CA = True
                Case arPT_poAcc : If Not CA Then ppoAcc = True : CA = True

                Case arPT_Pro : cP = True
                Case arPT_poPro : If Not cP Then ppoProp = True : cP = True

                Case arPT_IUI : cU = True
                Case arPT_poIUI : If Not cU Then ppoIUI = True : cU = True

                    '      Case arPT_Int:        CI = True
                    '      Case arPT_poInt:      If Not CI Then ppoInt = True: CI = True
            End Select
            RS.MoveNext
        Loop

        RS = Nothing
    End Sub

    Private Function AlreadyMadeSameAsCash() As Boolean
        Dim R As ADODB.Recordset
        R = GetRecordsetBySQL("SELECT * FROM [Transactions] WHERE [ArNo]=""" & ProtectSQL(ArNo) & """ ORDER BY [TransDate] ASC, [TransactionID] ASC", , GetDatabaseAtLocation())
        Do While Not R.EOF
            If LCase(Microsoft.VisualBasic.Left(R("Type").Value, 7)) = LCase(arPT_New) Then AlreadyMadeSameAsCash = False
            If LCase(Microsoft.VisualBasic.Left(R("Receipt").Value, 12)) = "same as cash" Then AlreadyMadeSameAsCash = True
            R.MoveNext
        Loop
        DisposeDA(R)
    End Function

    Public Function QueryPayLogSale(ByVal I As Integer) As String
        If I > PayCount Then Exit Function
        QueryPayLogSale = PayLog(I - 1).SaleNo
    End Function

    Public Function QueryPayLogAmount(ByVal I As Integer) As Decimal
        If I > PayCount Then Exit Function
        QueryPayLogAmount = PayLog(I - 1).Amount
    End Function

    Public Sub GetCustomerAccount()
        ArNo = -1
        Show()

        If ARPaySetUp.AccountFound = "" Then  'show entry form
TryAgain:
            ArCheck.Text = "Installment Customer"
            ArCheck.lblInstructions.Text = "Customer Account Number"
            'ArCheck.HelpContextID = HelpContextID

            ArCheck.ShowDialog(Me)
            ArNo = IIf(ArCheck.Customer = "", 0, ArCheck.Customer)
            mArNo = ArNo

            mDBAccess_Init(ArNo)
            mDBAccess.GetRecord()    ' this gets the record
            mDBAccess.dbClose()
            mDBAccess = Nothing

            If ArNo <> "-1" And ArNo <> "0" Then 'not found
                GetCustomer()
                mDBAccessTransactions_Init(ArNo)
                mDBAccessTransactions.GetRecord()    ' this gets the record
                mDBAccessTransactions.dbClose()
                mDBAccessTransactions = Nothing
                GetPayoff()
                GetAgeing()
                filFile_Click()
            ElseIf ArNo = 0 Then
                'Unload Me
                Me.Close()
            Else 'If ArNo = "-1" Then 'not found
                If MessageBox.Show("Incorrect Account Number.  Try again?", "", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) = DialogResult.Yes Then
                    GoTo TryAgain
                Else
                    'Unload Me
                    Me.Close()
                    MainMenu.Show()
                End If
            End If
        End If
    End Sub

    Public Sub VoidAccount()
        If ArNo <> "0" And Status <> arST_Void Then
            If MessageBox.Show("Are you sure you want to void this installment contract?", "", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning) = DialogResult.OK Then
                mDBAccess_Init(ArNo)
                mDBAccess.SetRecord()    ' this sets the record, it will set either
                mDBAccess.dbClose()
                mDBAccess = Nothing
            End If
        End If

        ' *** do something with money ***
        'Unload Me
        Me.Close()
        MainMenu.Show()
    End Sub

End Class