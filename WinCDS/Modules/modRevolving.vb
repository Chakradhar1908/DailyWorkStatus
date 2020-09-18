Imports VBRUN
Module modRevolving
    Private Structure ReportColumn
        Dim Title As String
        Dim Position As Integer
        Dim Alignment As AlignConstants
        Dim Type As String
    End Structure

    Private Const SkipRecord As String = "*SKIP*"
    Private Const SummaryColumn_Max As Integer = 9

    Private Enum eRevolvingSummaryColumns
        ' Account     Name     Balance     Interest     LateFees     Email     Phone1     Phone2
        eRSC_Account = 0
        eRSC_Status = 1
        eRSC_Name = 2
        eRSC_Balance = 3
        eRSC_Interest = 4
        eRSC_LateFees = 5
        eRSC_Email = 6
        eRSC_Phone1 = 7
        eRSC_Phone2 = 8
        eRSC_Phone3 = 9
    End Enum

    Private SummaryCols() As ReportColumn
    Private SummaryData(,) As Object
    Private CurrentRecord As Integer ' Painted myself into a corner and have to use a form variable...
    Private SummaryRecordCount As Integer
    Private SummarySkipCount As Integer

    Public CancelRevolvingProcessing As Boolean

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
        Dim I As Integer, J As Integer
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

    Public Function CalculateRevolvingPayment(ByRef NewBalance As Decimal, Optional ByRef RoundUp As Boolean = False, Optional ByRef Months As Integer = 0) As Decimal
        ' Future store settings may change this calculation.  Evridge's wants 1/3 of the balance due monthly.
        ' Balance due is complicated.  See RevolvingCurrentFinancedAmount for details.
        Dim Portion As Decimal
        If Months > 0 Then
            Portion = NewBalance / Months
        Else
            Portion = NewBalance * RevolvingMinimumPaymentPercent()
        End If
        If RoundUp Then
            CalculateRevolvingPayment = Math.Round(Portion + 0.49, 0)
        Else
            CalculateRevolvingPayment = Math.Round(Portion + 0.0049, 2)
        End If
    End Function

    Public Function RevolvingMinimumPaymentPercent() As Double
        'RevolvingMinimumPaymentPercent = 1 / 3
        RevolvingMinimumPaymentPercent = StoreSettings.ModifiedRevolvingMinPmt
        If RevolvingMinimumPaymentPercent > 1 Then RevolvingMinimumPaymentPercent = RevolvingMinimumPaymentPercent / 100
    End Function

    Public Function RevolvingSameAsCash() As Integer
        ' Months same as cash.  Encapsulated here in case we make this a store option later, only one place to change.
        'RevolvingSameAsCash = 3
        RevolvingSameAsCash = StoreSettings.ModifiedRevolvingSameAsCash
    End Function

    Public Sub DoRevolvingProcessAccount(ByVal DTE As Date, ByVal ArNo As String, Optional ByVal PrintStatements As Boolean = True)
        ' For each S account
        '   for each unpaid sale over (RevolvingSameAsCash() - saved in Installment?) days old
        '      charge interest (Rate in Installment) [as lump sum, or per sale?]

        Dim RS As ADODB.Recordset, SQL As String
        Dim Rate As Double, Balance As Decimal, NewInterest As Decimal, NewInterestTotal As Decimal, LastInterestDate As Date
        Dim AmountProcessed As Decimal
        Dim LateInterestOn As Date
        Dim RealLateDueOn As Integer

        On Error GoTo ProcessError

        Dim IA2 As New cInstallment ' One object to loop through, one to edit..

        If SummaryRecordCount = 0 Then  ' if called on a single account
            ReDim SummaryData(2, SummaryColumn_Max)
            CurrentRecord = 1
        End If


        RevolvingLog("Processing")

        NewInterestTotal = 0
        AmountProcessed = 0

        '  Balance = InstAcct.Balance
        IA2.Load(ArNo)
        Balance = IA2.Balance
        LastInterestDate = IA2.GetLastInterestDate
        LateInterestOn = DateAdd("d", 1, DateAdd("m", IA2.CashOpt, IA2.DeliveryDate))
        RealLateDueOn = DateAndTime.Day(LateInterestOn)

        DoLog("Processing Account " & IA2.ArNo, True)

        'If IsDevelopment And IA2.ArNo = "142 R" Then Stop
        'If IsDevelopment And IA2.ArNo = "L2126 R" Then Stop

        '  If RealLateDueOn <> Day(dte) And RevolvingStatementDay <> Day(dte) Then
        ''    If RealLateDueOn <> Day(dte) Then
        '    ' Wrong day for processing this account.
        '    SummaryData(CurrentRecord, eRSC_Account) = SkipRecord
        '    DoLog "Account " & AlignString(IA2.ArNo, 7) & " must be processed on the " & Ordinal(RealLateDueOn) & " or the " & Ordinal(RevolvingStatementDay) & ".", True
        ''      DoLog "Account " & AlignString(IA2.ArNo, 7) & " must be processed on the " & Ordinal(RealLateDueOn) & ".", True
        If DateAdd("m", 1, LastInterestDate) > DTE Then
            ' Interest has already been charged this month

            DoLog(" Account " & AlignString(IA2.ArNo, 7) & " owes a balance of " & AlignString(FormatCurrency(IA2.Balance), 12) & ".", True)
            DoLog("  Interest has already been charged this month [" & LastInterestDate & "].", True)
            ' Show how much?


            'BFH20150117 - Calculations are now defunct
            '    If Not RevolvingPrintStatementWithNoInterest Then
            '      SummaryData(CurrentRecord, eRSC_Account) = SkipRecord
            '    End If
        ElseIf IA2.FirstPayment > DTE Then
            DoLog(" Account " & AlignString(IA2.ArNo, 7) & " owes a balance of " & AlignString(FormatCurrency(IA2.Balance), 12) & ".", True)
            DoLog("  No interest because first payment isn't due [" & IA2.FirstPayment & "].", True)

            'BFH20150117 - Calculations are now defunct
            '    If Not RevolvingPrintStatementWithNoInterest Then
            '      SummaryData(CurrentRecord, eRSC_Account) = SkipRecord
            '    End If
        ElseIf IA2.Balance = 0 Then
            DoLog(" Account " & AlignString(IA2.ArNo, 7) & " has a $0.00 balance.", True)
        ElseIf IA2.Balance < 0 Then
            ' We owe them money.
            DoLog(" Account " & AlignString(IA2.ArNo, 7) & " is owed a refund of " & AlignString(FormatCurrency(IA2.Balance), 12) & ".", True)

            'BFH20150117 - Calculations are now defunct
            '    If Not RevolvingPrintStatementWithNoInterest Then
            '      SummaryData(CurrentRecord, eRSC_Account) = SkipRecord
            '    End If
        ElseIf IA2.Balance > 0 Then
            ' We are owed money and maybe interest.  Calculate and charge interest here.
            Dim InterestAmount As Decimal, MonthsInterest As Integer
            DoLog("Account " & AlignString(IA2.ArNo, 7) & " owes a balance of " & AlignString(FormatCurrency(IA2.Balance), 12) & ".")

            Application.DoEvents()
            AmountProcessed = IA2.INTEREST
            NewInterestTotal = Math.Round(IA2.INTEREST * IA2.Rate / 100 + 0.0049, 2) ' IA.Interest must be maintained as it gets paid off!
            If NewInterestTotal <> 0 Then DoLog("  Compounding " & AlignString(FormatCurrency(NewInterestTotal), 12) & " interest to account " & AlignString(IA2.ArNo, 7) & ". " & DescribeInterest(IA2.INTEREST, IA2.Rate, 1))
            Dim PIP As Decimal, IsOverdue As Boolean
            PIP = IA2.PaidInPeriod(DTE)
            IsOverdue = PIP < IA2.PerMonth ' This is the payment due per month, right?
            ' Sort of.  It may be wrong for the first statement after a purchase...
            ' unused value unless we do late fees

            Dim Holding As New cHolding
            Dim H2 As New cHolding
            Holding.Load(IA2.ArNo, "ArNo")
            Do Until Holding.DataAccess.Record_EOF
                If IsIn(Holding.Status, "S", "F") And Holding.Sale > Holding.Deposit Then
                    ' Sale is open and owes money
                    ' If the sale is older than SameAsCash or if IsOverdue, charge interest,
                    ' Unless interest has already been charged in this period.
                    ' Or unless it hasn't been on a statement yet.
                    Dim FDate As Date
                    '          FDate = Holding.FinanceDate
                    FDate = IA2.DeliveryDate
                    AmountProcessed = AmountProcessed + Holding.Sale - Holding.Deposit
                    MonthsInterest = RevolvingMonthsInterest(FDate, DTE, IA2.CashOpt)
                    If MonthsInterest > 0 And Holding.InterestChargedInPeriod(DTE, 1) = 0 Then
                        '          If DateAdd("d", RevolvingStatementDay - Day(FDate), FDate) < dte And DateAdd("m", IA2.CashOpt, FDate) <= dte And Holding.InterestChargedInPeriod(dte, 1) = 0 Then
                        ' Interest on (cashopt+1)th day is (cashopt) months worth of the remaining balance
                        '            If IA2.CashOpt > 0 And DateAdd("m", IA2.CashOpt, FDate) <= dte And DateAdd("m", IA2.CashOpt + 1, FDate) > dte Then
                        '              MonthsInterest = IA2.CashOpt
                        '            Else
                        '              MonthsInterest = 1
                        '            End If
                        InterestAmount = Math.Round((Holding.Sale - Holding.Deposit) * (IA2.Rate / 100) * MonthsInterest + 0.0049, 2)
                        '            If MonthsInterest = IA2.CashOpt Then
                        '              ' If we're doing the 91st day interest, we shouldn't charge the same interest we charged for the first 2 late payments
                        '              ' But we're doing sane interest, so none was charged for late payments.
                        '              InterestAmount = InterestAmount - Holding.InterestChargedInPeriod(dte, MonthsInterest)
                        '              If InterestAmount < 0 Then InterestAmount = 0
                        '            End If
                        'Dim H2 As New cHolding ' Because saving Holding loses the recordset we're looping, edit a new Holding object.
                        H2.Load(Holding.LeaseNo)
                        H2.AddInterest(InterestAmount, IA2.RevolvingInterestDate(DTE)) ' Updates Holding.Deposit and makes a GM row.
                        H2.Save() ' This changes the data source, doesn't it.. yeah, resets it to one row and breaks the loop.
                        NewInterestTotal = NewInterestTotal + InterestAmount
                        '            IA2.AddInterest InterestAmount, dte ' Updates InstallmentInfo fields and makes a Transactions row. ' But we want just one interest row in the account level
                        'BFH20150117 - Calculations are now defunct
                        '          SummaryData(CurrentRecord, eRSC_Interest) = SummaryData(CurrentRecord, eRSC_Interest) + InterestAmount
                        If InterestAmount <> 0 Then DoLog("  Adding " & AlignString(FormatCurrency(InterestAmount), 12) & " interest to account " & AlignString(IA2.ArNo, 7) & " for sale #" & AlignString(H2.LeaseNo, 12) & ". " & DescribeInterest(H2.Sale - H2.Deposit, IA2.Rate, MonthsInterest))
                    End If
                End If
                Holding.DataAccess.Records_MoveNext()
            Loop
            DisposeDA(Holding, H2)
            If AmountProcessed < Balance Then
                ' interest + sale totals < balance.  Charge interest on the difference respecting sale date and cashopt.
                ' IA2.DeliveryDate may not be right for this..
                MonthsInterest = RevolvingMonthsInterest(IA2.DeliveryDate, DTE, IA2.CashOpt)
                InterestAmount = Math.Round((Balance - AmountProcessed) * (IA2.Rate / 100) * MonthsInterest + 0.0049, 2)
                NewInterestTotal = NewInterestTotal + InterestAmount
                'BFH20150117 - Calculations are now defunct
                '      SummaryData(CurrentRecord, eRSC_Interest) = SummaryData(CurrentRecord, eRSC_Interest) + InterestAmount
                If InterestAmount <> 0 Then DoLog("  Adding " & AlignString(FormatCurrency(InterestAmount), 12) & " interest to account " & AlignString(IA2.ArNo, 7) & " for amount financed with no sale number. " & DescribeInterest(Balance - AmountProcessed, IA2.Rate, MonthsInterest))
            End If
            If NewInterestTotal = 0 And Not RevolvingVerbose Then
                ' Only one line got added to the listbox, remove it.
                If IsFormLoaded("frmRevolving") Then
                    If frmRevolving.lstLog.Items.Count > 0 Then frmRevolving.lstLog.Items.Remove(frmRevolving.lstLog.Items.Count - 1)
                End If
            End If
            '      If NewInterestTotal = 0 And Not RevolvingPrintStatementWithNoInterest Then
            '        SummaryData(CurrentRecord, eRSC_Account) = SkipRecord
            '      End If
        End If
        ' Recalculate next payment and maybe late charge.  If current payment ends in .00, assume it was set to round up on ArPaySetup.
        IA2.AddInterest(NewInterestTotal, IA2.RevolvingInterestDate(DTE))  ' Updates InstallmentInfo fields and makes a Transactions row.
        IA2.PerMonth = CalculateRevolvingPayment(IA2.Balance, Right(CurrencyFormat(IA2.PerMonth), 2) = ".00", IA2.Months)
        IA2.Save()

        ' record balance on account (adjust for any late fees)
        '  SummaryData(CurrentRecord, eRSC_Balance) = IA2.Balance ' balance after late fees
        ' BFH20150117 - So we can run, rerun, or be completely overrunnable,
        ' this was converted to a calculation, rather than a record operations.
        ' All other calculations of this are now defunct.
        CalculateAccountValues(ArNo, DTE, DateAndTime.Day(DTE) = RevolvingStatementDay())


        ' And print the statement
        If PrintStatements Then
            If RevolvingStatementDay() <> DateAndTime.Day(DTE) Then
                '      SummaryData(CurrentRecord, eRSC_Account) = SkipRecord
                '      DoLog "  Incorrect day for statement (" & Day(dte) & "<>" & RevolvingStatementDay & ")", True

            Else

                '      If SummaryData(CurrentRecord, eRSC_Account) <> SkipRecord Then
                If Not RevolvingPrintStatementWithNoInterest Then
                    If IA2.GetChargedInPeriod(RevolvingPreviousStatementDate(DTE), DTE) = 0 Then
                        DoLog("   Skipping Invoice Print for " & AlignString(IA2.ArNo, 7) & " because no interest was charged in period [" & RevolvingPreviousStatementDate(DTE) & " - " & DTE & "].")
                        GoTo SkipRecordAnyway
                    End If
                End If

                If IA2.Balance <> 0 Then
                    DoLog("   Printing Invoice Print for " & AlignString(IA2.ArNo, 7) & ".")
                    PrintMonthlyStatement(DTE, IA2)
                End If
                SummaryData(SummaryRecordCount + 1, eRevolvingSummaryColumns.eRSC_Balance) = SummaryData(SummaryRecordCount + 1, eRevolvingSummaryColumns.eRSC_Balance) + SummaryData(CurrentRecord, eRevolvingSummaryColumns.eRSC_Balance) ' running totals
                SummaryData(SummaryRecordCount + 1, eRevolvingSummaryColumns.eRSC_Interest) = SummaryData(SummaryRecordCount + 1, eRevolvingSummaryColumns.eRSC_Interest) + SummaryData(CurrentRecord, eRevolvingSummaryColumns.eRSC_Interest)
                SummaryData(SummaryRecordCount + 1, eRevolvingSummaryColumns.eRSC_LateFees) = SummaryData(SummaryRecordCount + 1, eRevolvingSummaryColumns.eRSC_LateFees) + SummaryData(CurrentRecord, eRevolvingSummaryColumns.eRSC_LateFees)
                '      Else
                '        SummarySkipCount = SummarySkipCount + 1
                '      End If


            End If
        End If
SkipRecordAnyway:

        Application.DoEvents()

        DisposeDA(IA2)
        Exit Sub

ProcessError:
        MessageBox.Show("Error processing statements: " & Err.Description, "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        Resume Next

    End Sub

    Public Function RevolvingLog(ByVal vMsg As String) As Boolean
        RevolvingLog = LogFile("Revolving.txt", vMsg)
    End Function

    Private Sub DoLog(ByVal vMsg As String, Optional ByVal vVerbose As Boolean = False)
        If vMsg = "CLEAR" Then
            If IsFormLoaded("frmRevolving") Then frmRevolving.lstLog.Items.Clear()
            Exit Sub
        End If
        If Not RevolvingVerbose And vVerbose Then Exit Sub

        If IsFormLoaded("frmRevolving") Then
            frmRevolving.lstLog.Items.Add(vMsg)
        End If
        RevolvingLog(vMsg)
    End Sub

    Private Function DescribeInterest(ByRef Balance As Decimal, ByRef Rate As Double, ByRef MonthsInterest As Integer) As String
        DescribeInterest = "(" & FormatCurrency(Balance) & " * " & CurrencyFormat(Rate) & "%" & IIf(MonthsInterest = 1, "", " * " & MonthsInterest & " months") & ")"
    End Function

    Public Function RevolvingMonthsInterest(ByVal SaleDate As Date, ByVal QueryDate As Date, ByVal CashOpt As Integer) As Decimal
        RevolvingMonthsInterest = 0 ' Default to no interest
        If CashOpt < 0 Then CashOpt = 0
        If DateAdd("m", CashOpt, SaleDate) >= QueryDate Then Exit Function ' No interest before the sale date or before cashopt expires

        ' Interest on (cashopt+1)th day is (cashopt) months worth of the remaining balance
        If CashOpt > 0 And DateAdd("m", CashOpt, SaleDate) <= QueryDate And DateAdd("m", CashOpt + 1, SaleDate) > QueryDate Then
            RevolvingMonthsInterest = CashOpt
        Else
            RevolvingMonthsInterest = 1
        End If
    End Function

    Private ReadOnly Property RevolvingVerbose() As Boolean
        Get
            If RevolvingFullDebug Or RevolvingSimpleDebug Then RevolvingVerbose = True : Exit Property
            If IsFormLoaded("frmRevolving") Then
                If frmRevolving.fraDev.Visible = True Then
                    If frmRevolving.chkDevVerbose.Checked <> False Then
                        RevolvingVerbose = True
                        Exit Property
                    End If
                End If
            End If
            RevolvingVerbose = False
        End Get
    End Property

    Private Sub CalculateAccountValues(ByVal ArNo As String, Optional ByVal DTE As Date = Nothing, Optional ByVal ForEntireMonth As Boolean = False)
        Dim CI As cInstallment
        CI = New cInstallment

        '  If IsDevelopment And ArNo = "4159 R" Then Stop
        CI.Load(ArNo, "ArNo")

        SummaryData(CurrentRecord, eRevolvingSummaryColumns.eRSC_Balance) = CI.Balance
        If ForEntireMonth Then
            SummaryData(CurrentRecord, eRevolvingSummaryColumns.eRSC_Interest) = CI.GetChargedInPeriod(DateAdd("m", -1, DTE), DTE)
            SummaryData(CurrentRecord, eRevolvingSummaryColumns.eRSC_LateFees) = CI.GetChargedInPeriod(DateAdd("m", -1, DTE), DTE, "Late Fees")
        Else
            SummaryData(CurrentRecord, eRevolvingSummaryColumns.eRSC_Interest) = CI.GetChargedInPeriod(DateAdd("d", -1, DTE), DTE)
            SummaryData(CurrentRecord, eRevolvingSummaryColumns.eRSC_LateFees) = CI.GetChargedInPeriod(DateAdd("d", -1, DTE), DTE, "Late Fees")
        End If

        DisposeDA(CI)
    End Sub

    Public Function RevolvingStatementDay() As Integer
        RevolvingStatementDay = 25
    End Function

    Private ReadOnly Property RevolvingPrintStatementWithNoInterest() As Boolean
        Get
            RevolvingPrintStatementWithNoInterest = True
        End Get
    End Property

    Public Function RevolvingPreviousStatementDate(ByVal D As Date) As Date
        RevolvingPreviousStatementDate = DateAdd("m", -1, D)
    End Function

    Public Sub PrintMonthlyStatement(ByRef DTE As Date, ByRef InstAcct As cInstallment)
        ' store and customer information
        On Error Resume Next
        Dim mR As clsMailRec, NextPaymentDate As Date

        'If IsDevelopment And InstAcct.ArNo = "4299 R" Then Stop


        mR = New clsMailRec
        mR.Load(InstAcct.MailIndex)
        NextPaymentDate = DateAdd("m", 1, DateAdd("d", -DateAndTime.Day(Today) + InstAcct.LateDueOn, Today)) ' 25th of next month, when LateDueOn=25
        Do While NextPaymentDate < DTE
            NextPaymentDate = DateAdd("m", 1, NextPaymentDate)
        Loop
        PrintMonthlyStatementHeadingSingleWindow(mR, InstAcct, DTE, NextPaymentDate)

        ' previous balance, recent payments, interest, late fees, current balance
        PrintTransactionsHeader()
        Dim Trans As New cTransaction, TotNewCharges As Decimal
        Dim TotCharges As Decimal, TotCredits As Decimal, TotNewSales As Decimal

        Trans.DataAccess.Records_OpenFieldIndexAt("ArNo", InstAcct.ArNo, "TransactionID")
        Do While Trans.DataAccess.Records_Available
            If Trans.TransDate > DateAdd("m", -1, DTE) And Trans.TransDate <= DTE Then
                '    If DateAdd("m", 1, Trans.TransDate) > dte Then
                If Printer.CurrentY + Printer.TextHeight("X") > Printer.ScaleHeight Then PrintMonthlyStatementHeading(mR, InstAcct) : PrintTransactionsHeader()
                PrintToPosition(Printer, Trans.TransDate, 1800, AlignConstants.vbAlignRight, False)
                PrintToPosition(Printer, Trans.TransType, 2100, AlignConstants.vbAlignLeft, False)
                PrintToPosition(Printer, CurrencyFormat(Trans.Charges), 4900, AlignConstants.vbAlignRight, False)
                If Not IsIn(Left(Trans.TransType, 7), "NewSale", "Doc Fee") Then ' don't show "credits" on a new sale, because it's the deposit...  hide it in the numbers
                    PrintToPosition(Printer, CurrencyFormat(Trans.Charges - Trans.Credits), 4900, AlignConstants.vbAlignRight, False)
                Else
                    PrintToPosition(Printer, CurrencyFormat(Trans.Charges), 4900, AlignConstants.vbAlignRight, False)
                    PrintToPosition(Printer, CurrencyFormat(Trans.Credits), 6000, AlignConstants.vbAlignRight, False)
                End If
                PrintToPosition(Printer, CurrencyFormat(Trans.Balance), 7200, AlignConstants.vbAlignRight, False)
                PrintToPosition(Printer, Trans.Receipt, 11500, AlignConstants.vbAlignRight, True)
                TotNewCharges = TotNewCharges - Trans.Credits + Trans.Charges
                If Not IsIn(Left(Trans.TransType, 7), "NewSale", "Doc Fee") Then ' don't show these (desposits for new sales)
                    TotCredits = TotCredits + Trans.Credits
                End If
                If IsIn(Left(Trans.TransType, 7), "NewSale", "Doc Fee") Then
                    TotNewSales = TotNewSales + Trans.Charges - Trans.Credits
                Else
                    TotCharges = TotCharges + Trans.Charges
                End If
            End If
        Loop
        If Printer.CurrentY + Printer.TextHeight("X") * 4 > Printer.ScaleHeight Then PrintMonthlyStatementHeading(mR, InstAcct)
        PrintTransactionsFooter(InstAcct.Balance, TotNewCharges, TotNewSales, TotCredits, TotCharges)
        PrintCentered("Interest assessed at " & Format(InstAcct.Rate, "0.00##") & "% per month (" & Format(InstAcct.Rate * 12, "0.00##") & "% APR).")
        '  PrintCentered "If sale is paid in full within 90 days of invoice date, any interest charges will be reversed upon receipt of your final payment."
        Printer.Print()

        ' Previous balance is computed from total due minus recent charges
        If Printer.CurrentY + Printer.TextHeight("X") > Printer.ScaleHeight Then PrintMonthlyStatementHeading(mR, InstAcct)
        Dim Str As String
        Str = "Your next payment of " & FormatCurrency(InstAcct.PerMonth) & " is due on " & NextPaymentDate & "."
        PrintCentered(Str, , True)
        Printer.Print()

        Printer.FontBold = False

        'If IsDevelopment And InstAcct.ArNo = "4299 R" Then Stop

        ' and for each open sale, list of items, amounts, totals
        Dim GM As New CGrossMargin
        Dim Holding As New cHolding
        Holding.Load(InstAcct.ArNo, "ArNo")
        Do Until Holding.DataAccess.Record_EOF
            ' If it's open, print the details.
            '    If IsIn(Holding.Status, "F", "S") And Holding.Sale > Holding.Deposit Then ' wasn't showing delivered, paid sales...
            If IsIn(Holding.Status, "F", "S", "D") Then
                ' Don't forget to check for end of page
                If Printer.CurrentY + Printer.TextHeight("X") * 5 > Printer.ScaleHeight Then PrintMonthlyStatementHeading(mR, InstAcct)
                PrintHoldingHeading(Holding)
                'Dim GM As New CGrossMargin -------------> Moved this line to the above Do Until Loop.If not, DisposeDA (botton code line) will not recognizing GM variable.
                GM.Load(Holding.LeaseNo, "SaleNo")
                Do Until GM.DataAccess.Record_EOF
                    If GM.SellPrice <> 0 And GM.Style <> "SUB" And GM.Style <> "INTEREST" Then
                        If Printer.CurrentY + Printer.TextHeight("X") > Printer.ScaleHeight Then PrintMonthlyStatementHeading(mR, InstAcct) : PrintHoldingHeading(Holding)
                        PrintToPosition(Printer, GM.SellDte, 1800, AlignConstants.vbAlignRight, False)
                        PrintToPosition(Printer, GM.Quantity, 2900, AlignConstants.vbAlignRight, False)
                        PrintToPosition(Printer, GM.Desc, 3100, AlignConstants.vbAlignLeft, False)
                        PrintToPosition(Printer, FormatCurrency(GM.SellPrice), 11500, AlignConstants.vbAlignRight, True)
                    End If
                    GM.DataAccess.Records_MoveNext()
                Loop
                If Printer.CurrentY + Printer.TextHeight("X") * 3 > Printer.ScaleHeight Then PrintMonthlyStatementHeading(mR, InstAcct) : PrintHoldingHeading(Holding)
                PrintHoldingFooting(Holding)
            End If
            Holding.DataAccess.Records_MoveNext()
        Loop
        Printer.EndDoc()

        ' Cleanup
        DisposeDA(mR, Holding, GM, Trans)
    End Sub

    Private ReadOnly Property RevolvingFullDebug() As Boolean
        Get
            RevolvingFullDebug = False
            If IsFormLoaded("frmRevolving") Then
                If frmRevolving.fraDev.Visible Then
                    If frmRevolving.chkDevDebugFull.Checked <> False Then
                        RevolvingFullDebug = True
                        Exit Property
                    End If
                End If
            End If
        End Get
    End Property

    Private ReadOnly Property RevolvingSimpleDebug() As Boolean
        Get
            RevolvingSimpleDebug = False
            If IsFormLoaded("frmRevolving") Then
                If frmRevolving.fraDev.Visible Then
                    If frmRevolving.chkDevDebugLite.checked <> False Then
                        RevolvingSimpleDebug = True
                        Exit Property
                    End If
                End If
            End If
        End Get
    End Property

    Private Sub PrintMonthlyStatementHeadingSingleWindow(ByRef mR As clsMailRec, ByRef InstAcct As cInstallment, ByRef StatementDate As Date, ByRef NextPaymentDate As Date)
        If Printer.CurrentY <> 0 Then Printer.NewPage()
        Printer.FontName = "Arial"
        Printer.FontSize = 18
        PrintCentered("Monthly Statement", , True)
        Printer.FontSize = 10

        Printer.CurrentX = 0
        Printer.CurrentY = 800 ' 1500 matches ArCard, but too low (1300 too low)
        Printer.Print(TAB(10), StoreSettings.Name) '; Tab(90); "Report Date: "; DateFormat(Now)
        Printer.Print(TAB(10), StoreSettings.Address)
        Printer.Print(TAB(10), StoreSettings.City)
        Printer.Print(TAB(10), StoreSettings.Phone)
        Printer.Print()

        Printer.FontBold = True
        Printer.FontSize = 13
        PrintToPosition(Printer, "Account:", 10300, AlignConstants.vbAlignRight, False)
        PrintToPosition(Printer, InstAcct.ArNo, 11500, AlignConstants.vbAlignRight, False)
        Printer.FontSize = 10
        Printer.FontBold = False
        PrintToPosition(Printer, "Statement Date:", 7300, AlignConstants.vbAlignRight, False)
        PrintToPosition(Printer, StatementDate, 8500, AlignConstants.vbAlignRight, True)
        PrintToPosition(Printer, "Total Balance:", 7300, AlignConstants.vbAlignRight, False)
        PrintToPosition(Printer, Format(InstAcct.Balance, "Currency"), 8500, AlignConstants.vbAlignRight, True)
        PrintToPosition(Printer, "Next Payment:", 7300, AlignConstants.vbAlignRight, False)
        PrintToPosition(Printer, Format(InstAcct.PerMonth, "Currency"), 8500, AlignConstants.vbAlignRight, True)
        PrintToPosition(Printer, "Payment Due:", 7300, AlignConstants.vbAlignRight, False)
        PrintToPosition(Printer, NextPaymentDate, 8500, AlignConstants.vbAlignRight, True)


        Printer.FontBold = False
        Printer.FontSize = 10
        Printer.CurrentX = 0
        Printer.CurrentY = 2830 ' 3530 matches ArCard, but too low (3300 too low)
        Printer.Print()
        Printer.Print(TAB(10), Trim(mR.First), " ", Trim(mR.Last), TAB(60))
        Printer.Print(TAB(10), mR.Address)
        If Trim(mR.AddAddress) <> "" Then Printer.Print(TAB(10), mR.AddAddress)
        Printer.Print(TAB(10), mR.City, " ", mR.Zip)
        Printer.Print()

        ' if we want telephone numbers, captions may need fixing (which should be in the class object)

        Printer.Print()
        Printer.Print()   'BFH20150211 2 extra lines (total three) added to clear the window.  Resulted in Currenty=4634
        Printer.Print()
        Printer.Print()   'BFH20150216 - one more
        Printer.Print()   'BFH20150217 - one more


        '  Printer.CurrentY = 4634
        Dim mR2 As clsMailShipTo
        mR2 = mR.ShipTo
        Printer.Print(TAB(10), DressTelephoneLabel(mR.PhoneLabel1, mR.Tele), DressAni(mR.Tele), "   ", DressTelephoneLabel(mR.PhoneLabel2, mR.Tele2), DressAni(mR.Tele2), "   ", DressTelephoneLabel(mR2.PhoneLabel3, mR2.Tele), DressAni(mR2.Tele))
        If mR.Email <> "" Then Printer.Print(TAB(10), "Email: ", mR.Email)

        DisposeDA(mR2)
    End Sub

    Public Sub PrintTransactionsHeader()
        ' Check for 1 line free before calling this.
        PrintCentered("Recent Transactions", , True)
        PrintLine()
        Printer.FontBold = True
        PrintToPosition(Printer, "Date", 1800, AlignConstants.vbAlignRight, False)
        PrintToPosition(Printer, "Type", 2100, AlignConstants.vbAlignLeft, False)
        PrintToPosition(Printer, "Charges", 4900, AlignConstants.vbAlignRight, False)
        PrintToPosition(Printer, "Credits", 6000, AlignConstants.vbAlignRight, False)
        PrintToPosition(Printer, "Balance", 7200, AlignConstants.vbAlignRight, False)
        PrintToPosition(Printer, "Receipt / Notes", 11500, AlignConstants.vbAlignRight, True)
        Printer.FontBold = False
        PrintLine()
    End Sub

    Private Sub PrintMonthlyStatementHeading(ByRef mR As clsMailRec, ByRef InstAcct As cInstallment)
        If Printer.CurrentY <> 0 Then Printer.NewPage()
        Printer.FontName = "Arial"
        Printer.FontSize = 10

        Printer.CurrentX = 0
        Printer.CurrentY = 300
        Printer.Print("Sold :", TAB(10), Trim(mR.First & " " & mR.Last))
        Printer.Print(" To", TAB(10), mR.Address)
        Printer.Print(TAB(10), mR.City)

        Printer.CurrentX = 0
        Printer.CurrentY = 300
        Printer.Print(TAB(38), "OPEN CHARGE")
        Printer.Print(TAB(38), "Acct # ", TAB(48), InstAcct.ArNo)
        Printer.Print(TAB(38), "Date : ", TAB(48), Today)
        Printer.Print(TAB(38), "Balance: ", TAB(48), FormatCurrency(InstAcct.Balance))

        Printer.CurrentX = 0
        Printer.CurrentY = 300

        Printer.Print(TAB(80), StoreSettings.Name, " Page: " & Printer.Page)
        Printer.Print(TAB(80), StoreSettings.Address)
        Printer.Print(TAB(80), StoreSettings.City)
        Printer.Print(TAB(80), StoreSettings.Phone)

        Printer.Print()
    End Sub

    Public Sub PrintTransactionsFooter(ByVal Balance As Decimal, ByVal TotNewCharges As Decimal, Optional ByVal TotNewSales As Decimal = 0, Optional ByVal TotNewPayments As Decimal = 0, Optional ByVal TotNewInterest As Decimal = 0)
        ' Check for 3 lines free before calling this.
        PrintLine()
        PrintToPosition(Printer, "Previous Balance", 6000, AlignConstants.vbAlignRight, False)
        PrintToPosition(Printer, FormatCurrency(Balance - TotNewCharges), 7200, AlignConstants.vbAlignRight, True)

        PrintToPosition(Printer, "+New Sales", 6000, AlignConstants.vbAlignRight, False)
        PrintToPosition(Printer, FormatCurrency(TotNewSales), 7200, AlignConstants.vbAlignRight, True)

        PrintToPosition(Printer, "-Payments", 6000, AlignConstants.vbAlignRight, False)
        PrintToPosition(Printer, FormatCurrency(TotNewPayments), 7200, AlignConstants.vbAlignRight, True)

        PrintToPosition(Printer, "+Interest", 6000, AlignConstants.vbAlignRight, False)
        PrintToPosition(Printer, FormatCurrency(TotNewInterest), 7200, AlignConstants.vbAlignRight, True)

        Printer.FontBold = True
        PrintToPosition(Printer, "---------------", 6000, AlignConstants.vbAlignRight, False)
        PrintToPosition(Printer, "---------------", 7200, AlignConstants.vbAlignRight, True)

        '  PrintToPosition Printer, "New Charges", 6000, vbAlignRight, False
        '  PrintToPosition Printer, FormatCurrency(TotNewCharges), 7200, vbAlignRight, False
        '  PrintToPosition Printer, "(Payments and Interest)", 7500, vbAlignLeft, True

        PrintToPosition(Printer, "Total Balance", 6000, AlignConstants.vbAlignRight, False)
        PrintToPosition(Printer, FormatCurrency(Balance), 7200, AlignConstants.vbAlignRight, True)
        Printer.FontBold = False
    End Sub

    Private Sub PrintHoldingHeading(ByRef H As cHolding)
        ' Check for 5 lines free before printing this.
        Printer.FontBold = True
        Printer.Print("Invoice #", H.LeaseNo)
        PrintLine()
        PrintToPosition(Printer, "DATE", 1800, AlignConstants.vbAlignRight, False)
        PrintToPosition(Printer, "QTY", 2900, AlignConstants.vbAlignRight, False)
        PrintToPosition(Printer, "DESCRIPTION", 3100, AlignConstants.vbAlignLeft, False)
        PrintToPosition(Printer, "AMOUNT", 11500, AlignConstants.vbAlignRight, True)
        PrintLine()
        Printer.FontBold = False
    End Sub

    Private Sub PrintHoldingFooting(ByRef H As cHolding)
        ' Check for 3 lines free before printing this.
        Printer.FontBold = True
        PrintLine()
        PrintToPosition(Printer, "TOTAL", 10500, AlignConstants.vbAlignRight, False)
        PrintToPosition(Printer, FormatCurrency(H.Sale), 11500, AlignConstants.vbAlignRight, True)
        PrintToPosition(Printer, "-AMT PAID", 10500, AlignConstants.vbAlignRight, False)
        PrintToPosition(Printer, FormatCurrency(H.Deposit), 11500, AlignConstants.vbAlignRight, True)
        PrintToPosition(Printer, "=UNPAID BAL", 10500, AlignConstants.vbAlignRight, False)
        PrintToPosition(Printer, FormatCurrency(H.Sale - H.Deposit), 11500, AlignConstants.vbAlignRight, True)
        Printer.FontBold = False
    End Sub

    Private Function DressTelephoneLabel(ByRef S As String, Optional ByRef IfNotBlank As String = "Not Blank") As String
        If Trim(IfNotBlank) = "" Then Exit Function
        DressTelephoneLabel = Trim(S)
        If DressTelephoneLabel = "" Then DressTelephoneLabel = "Telephone:"
        If Right(DressTelephoneLabel, 1) <> ":" Then DressTelephoneLabel = DressTelephoneLabel & ":"
        DressTelephoneLabel = DressTelephoneLabel & " "
    End Function

    Public Function ConvertToRevolvingCharge(ByRef ArNo As String) As Boolean
        ' This can never happen, says Jerry. MJK20140223
        Exit Function

        '  ' Already revolving? Already an R account with this prefix?
        '  ' Message should probably be handled in the converter for better detail.
        '  'MsgBox "Failed to convert to a revolving account.", vbCritical + vbOKOnly, "Conversion Failure"
        '
        '  ' This may be very wrong - is it better to close the old account and open a new?
        '
        '  ConvertToRevolvingCharge = False
        '
        '  Dim NewArno As String
        '  NewArno = ArNo & RevolvingSuffix()
        '  Dim SQL As String, RS As Recordset
        '  SQL = "select count(*) as cnt from InstallmentInfo where ArNo=""" & ProtectSQL(NewArno) & """"
        '  Set RS = GetRecordsetBySQL(SQL)
        '  If RS("cnt") > 0 Then
        '    ' Fail because there's already a revolving charge with this prefix.
        '    MsgBox "There's already a revolving charge starting with " & ArNo & ".", vbCritical + vbOKOnly, "Conversion Failure"
        '    DisposeDA RS
        '    Exit Function
        '  End If
        '  DisposeDA RS
        '
        '  ' ArNo is referenced in InstallmentInfo,ArApp, and Transactions.  Update all references.
        '  SQL = "update InstallmentInfo set ArNo=""" & ProtectSQL(NewArno) & """ where ArNo=""" & ProtectSQL(ArNo) & """"
        '  ExecuteRecordsetBySQL SQL
        '  SQL = "update ArApp set ArNo=""" & ProtectSQL(NewArno) & """ where ArNo=""" & ProtectSQL(ArNo) & """"
        '  ExecuteRecordsetBySQL SQL
        '  SQL = "update Transactions set ArNo=""" & ProtectSQL(NewArno) & """ where ArNo=""" & ProtectSQL(ArNo) & """"
        '  ExecuteRecordsetBySQL SQL
        '
        '  ConvertToRevolvingCharge = True
    End Function

End Module
