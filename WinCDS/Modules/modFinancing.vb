Module modFinancing
    Public Const ArNo_AddOnRecordSeparator As String = "-"
    Public Const ArNo_AddOnRecordToken As String = "AddOnRecord"
    Public Const ArNo_AddOnRecordIndicator As String = ArNo_AddOnRecordSeparator & ArNo_AddOnRecordToken & ArNo_AddOnRecordSeparator
    Public Const ArNo_AddOnRecordPattern_LIKE As String = "*" & ArNo_AddOnRecordToken & "*"
    Public Const ArNo_AddOnRecordPattern_SQL As String = "%" & ArNo_AddOnRecordToken & "%" ' MS Access requires % for wildcard

    Public Function ArNoIsAddOnRecord(ByVal ArNo As String) As Boolean
        ArNoIsAddOnRecord = ArNo Like ArNo_AddOnRecordPattern_LIKE
    End Function

    Public Function CalculateSIR(ByVal Balance As Decimal, ByVal TargetAPR As Double, ByVal Months As Integer) As Double
        Dim SIR As Double, FC As Decimal, Incr As Boolean, Delta As Double, APR As Double

        Dim Mx As Integer
        If Balance = 0 Then Exit Function
        SIR = 0.12
        FC = ((Balance * SIR) / 12 * Months)
        APR = CalculateAPR(Balance, FC, Months)
        Delta = 0.01
        If APR = TargetAPR Then   ' not likely..
            CalculateSIR = SIR
            Exit Function
        End If
        Incr = TargetAPR > APR

        Mx = 0
        Do While SIR <> APR And Delta > 0.00000001
            Mx = Mx + 1
            If TargetAPR > APR Then   ' target is higher than current
                If Incr Then            ' still going up
                    SIR = SIR + Delta
                Else                    ' we were going down, but went too far!
                    Delta = Delta / 10.0#
                    SIR = SIR + Delta
                    Incr = True           ' turn around and go back up (after changing delta)
                End If
            Else                      ' target is lower
                If Not Incr Then
                    SIR = SIR - Delta
                Else
                    Delta = Delta / 10.0#
                    SIR = SIR - Delta
                    Incr = False          ' turn around and go back down (after changing delta)
                End If
            End If
            FC = ((Balance * SIR) / 12 * Months)
            APR = CalculateAPR(Balance, FC, Months)
            If Mx > 1000000 Then Exit Do
        Loop
        If APR > TargetAPR Then
            SIR = SIR - 10 * Delta  'always keep below target, use 10* b/c it would have been shrunk
            FC = ((Balance * SIR) / 12 * Months)
            APR = CalculateAPR(Balance, FC, Months)
        End If
        CalculateSIR = SIR
        'Debug.Print "APR=" & APR
        'Debug.Print "SIR=" & SIR
    End Function

    Public Function CalculateAPR(ByVal Balance As Decimal, ByVal FinanceCharge As Decimal, ByVal Months As Integer, Optional ByVal DeferredMonths As Integer = 0) As Double
        Dim T As Double
        On Error GoTo BadVBARateFunction
        CalculateAPR = 1200 * Financial.Rate(Months, -(Balance + FinanceCharge) / Months, Balance)
        Exit Function
BadVBARateFunction:
        T = (3 * Balance * (Months + DeferredMonths + 1) + (FinanceCharge * (Months + DeferredMonths + 1)))
        If T <> 0 Then
            '  CalculateAPR = 100 * (6 * 12 * FinanceCharge) / (3 * Balance * (Months + DeferredMonths + 1) + (FinanceCharge * (Months + DeferredMonths + 1)))
            CalculateAPR = 100 * (6 * 12 * FinanceCharge) / T
        End If
    End Function

    Public Sub GetPreviousContractTerms(ByVal ArNo As String, Optional ByVal StoreNo As Integer = 0, Optional ByRef Prev As Decimal = 0, Optional ByRef Sale As Decimal = 0, Optional ByRef Deposit As Decimal = 0, Optional ByRef DocFee As Decimal = 0, Optional ByRef tLife As Decimal = 0, Optional ByRef tAcc As Decimal = 0, Optional ByRef tProp As Decimal = 0, Optional ByRef tIUI As Decimal = 0, Optional ByRef tInt As Decimal = 0, Optional ByRef tIntST As Decimal = 0)
        Dim RS As ADODB.Recordset
        Dim CL As Boolean, CA As Boolean, cP As Boolean, cU As Boolean ' , cI as boolean
        Dim IsSale As Boolean
        Dim F As String
        RS = GetRecordsetBySQL("SELECT * FROM [Transactions] WHERE ArNo='" & ArNo & "' ORDER BY [TransactionID]", , GetDatabaseAtLocation(StoreNo))

        Sale = 0
        Deposit = 0
        Prev = 0
        DocFee = 0
        tLife = 0
        tAcc = 0
        tProp = 0
        tIUI = 0
        tInt = 0
        tIntST = 0

        Do While Not RS.EOF
            F = IfNullThenNilString(RS("Type").Value)
            '    Debug.Print F
            If IsSale Or F Like "NewSal*" Then
                If F Like "NewSal*" Then
                    IsSale = True
                    Sale = IfNullThenZeroCurrency(RS("Charges").Value)
                    Deposit = IfNullThenZeroCurrency(RS("Credits").Value)
                    Prev = IfNullThenZeroCurrency(RS("Balance").Value) - IfNullThenZeroCurrency(RS("Charges").Value) + IfNullThenZeroCurrency(RS("Credits").Value)
                    DocFee = 0
                    tLife = 0
                    tAcc = 0
                    tProp = 0
                    tIUI = 0
                    tInt = 0
                    tIntST = 0
                Else
                    Select Case F
                        Case arPT_New
                        Case arPT_Doc : DocFee = IfNullThenZeroCurrency(RS("Charges").Value)
                        Case arPT_Lif : tLife = IfNullThenZeroCurrency(RS("Charges").Value)
                        Case arPT_Acc : tAcc = IfNullThenZeroCurrency(RS("Charges").Value)
                        Case arPT_Pro : tProp = IfNullThenZeroCurrency(RS("Charges").Value)
                        Case arPT_Int : tInt = IfNullThenZeroCurrency(RS("Charges").Value)
                        Case arPT_Tax : tIntST = IfNullThenZeroCurrency(RS("Charges").Value)
                    End Select
                End If
            End If
            RS.MoveNext()
        Loop

        RS = Nothing
    End Sub

    Public Function QueryLateDate(ByVal nDay As Integer, Optional ByVal nDate As Date = NullDate, Optional ByVal AsString As Boolean = True, Optional ByVal NoGraceAdj As Boolean = False) As Date
        Dim G As Integer
        QueryLateDate = QueryDueDate(nDay, nDate, , True)
        If Not NoGraceAdj Then
            G = AdjustedGracePeriod(nDay)
            QueryLateDate = DateAdd("d", G, QueryLateDate)
        End If
        If AsString Then QueryLateDate = DueDateFormat(QueryLateDate)
    End Function

    Public Function QueryDueDate(ByVal nDay As Integer, Optional ByVal nDate As Date = NullDate, Optional ByVal AsString As Boolean = True, Optional ByVal NoGraceAdj As Boolean = False) As Date
        Dim G As Integer
        If nDate = NullDate Then nDate = Today
        G = IIf(NoGraceAdj, 0, AdjustedGracePeriod(nDay))
        nDate = DateAdd("d", -G, nDate)
        QueryDueDate = DaySeek(nDate, nDay, False)
        If AsString Then QueryDueDate = DueDateFormat(QueryDueDate)
        'If AsString Then QueryDueDate = Date.Parse(DueDateFormat(QueryDueDate), Globalization.CultureInfo.InvariantCulture)
    End Function

    Public Function DueDateFormat(ByVal D As Date) As String
        DueDateFormat = Format(D, "MMM dd")
        'DueDateFormat = Format(D, "MMM dd yyyy")
    End Function

    Public Function GetArCreditHistory(ByVal ArNo As String, Optional ByRef RunAsDate As String = "", Optional ByVal nMonths As Integer = 24, Optional ByVal StoreNo As Integer = 0) As String
        '1 = 0 payments past due (current account)
        '2 = 30 - 59 days past due date
        '3 = 60 - 89 days past due date
        '4 = 90 - 119 days past due date
        '5 = 120+ days past due date
        'We use ' ' (space) to indicate no credit history.

        Dim S As String
        S = GetPaymentHistoryEquifax(ArNo, RunAsDate, nMonths, StoreNo)
        S = Replace(S, "B", " ")

        S = Replace(S, "6", "5")
        S = Replace(S, "4", "5")
        S = Replace(S, "3", "4")
        S = Replace(S, "2", "3")
        S = Replace(S, "1", "2")
        S = Replace(S, "0", "1")

        S = Replace(S, "D", " ")
        S = Replace(S, "E", " ")
        S = Replace(S, "G", " ")
        S = Replace(S, "H", " ")
        S = Replace(S, "J", " ")
        S = Replace(S, "K", " ")
        S = Replace(S, "L", " ")

        GetArCreditHistory = S
        Exit Function

    End Function

    Public Function ProRata(ByVal Value As Decimal, ByVal Months As Integer, ByVal LastPayDate As String, Optional ByVal WhenDate As Date = NullDate) As Decimal
        Dim N As Integer

        If Months <= 0 Then Exit Function
        If Not IsDate(LastPayDate) Then Exit Function
        CheckNullDate(WhenDate)
        N = CountMonths(WhenDate, DateValue(LastPayDate), , True)
        ProRata = Value / CDbl(Months) * N
        If ProRata < 0 Then ProRata = 0
        ProRata = Math.Round(ProRata, 2)
    End Function

    Public Function Rule78(ByVal InputDate As Date, ByVal InputValue As Decimal, ByVal MonthsToFinance As Integer, Optional ByVal DontRoundMonths As Boolean = False, Optional ByVal AsOfDate As Date = NullDate) As Decimal
        CheckNullDate(AsOfDate)
        If DateDiff("d", InputDate, AsOfDate) <= 1 Then ' fudge if close enough to contract start.
            Rule78 = InputValue
        Else
            Rule78 = QueryRule78(InputValue, CountMonths(InputDate, AsOfDate, DontRoundMonths), MonthsToFinance) ' dteDate1.value was FirstPayment.
        End If
    End Function

    Public Function QueryRule78(ByVal InputValue As Decimal, ByVal MonthsSinceDue As Integer, ByVal MonthsCalc As Integer) As Decimal
        ' This should always be based on Delivery Date, not Sale Date.
        ' C/C REBATE  RULE 78
        On Error Resume Next
        QueryRule78 = 0

        'DETERMINE IF BEYOND CONTRACT
        If MonthsSinceDue > MonthsCalc Then
            MonthsCalc = MonthsSinceDue
        End If

        If MonthsSinceDue < 1 Then
            'less than 1 month
            QueryRule78 = InputValue
            Exit Function
        End If


        QueryRule78 = CurrencyFormat(InputValue * (MonthsCalc - MonthsSinceDue) * (MonthsCalc - MonthsSinceDue + 1) / (MonthsCalc * (MonthsCalc + 1)))
        QueryRule78 = Math.Round(QueryRule78, 2)
    End Function

    Public Function CountMonths(ByVal InputDate As Date, Optional ByVal nDate As Date = NullDate, Optional ByVal DontRoundMonths As Boolean = False, Optional ByVal Fifteen As Boolean = True) As Integer
        Dim N As Integer, Ck As Date
        CheckNullDate(nDate)

        If DateBefore(nDate, InputDate, False) Then Exit Function
        If DontRoundMonths Then
            Do While True
                N = N + 1
                If DateAfter(DateAdd("m", N, InputDate), nDate, False) Then Exit Do
            Loop
        ElseIf Fifteen Then   ' we're guessing this is what people generally want...
            Do While True
                Ck = DateAdd("d", 14, DateAdd("m", N, InputDate))
                If DateAfter(Ck, nDate, False) Then Exit Do
                N = N + 1
                If DateAfter(DateAdd("m", N, InputDate), nDate, False) Then Exit Do
            Loop
        Else
            ' this is probably never correct.  It simply takes Month(B) - Month(A)....
            N = DateDiff("m", InputDate, nDate)
        End If
        CountMonths = N
    End Function

    Public Function RuleOfAnticipationForTreehouse(ByVal OrigLoan As Decimal, ByVal OrigPmt As Decimal, ByVal OrigPrem As Decimal, ByVal OrigMonths As Integer, ByVal RemainMonths As Integer, ByVal OrigRate As Double, ByVal RemainingRate As Double) As Decimal
        Dim RI As Decimal
        RI = OrigPmt * RemainMonths
        RuleOfAnticipationForTreehouse = ((RI / 100 * RemainingRate) / (OrigLoan / 100 * OrigRate)) * OrigPrem
        RuleOfAnticipationForTreehouse = Math.Round(RuleOfAnticipationForTreehouse, 2)
        If RuleOfAnticipationForTreehouse > OrigPrem Then RuleOfAnticipationForTreehouse = OrigPrem
    End Function

    Public Function ComputeAgeing(ByVal CurrentDate As Date,
      ByVal FirstPay As Date, ByVal Months As Integer, ByVal PaidBy As Integer, ByVal Weekly As Boolean,
      ByVal Payment As Decimal, ByVal TotPaid As Decimal, ByVal Financed As Decimal, ByVal Balance As Decimal,
      ByVal wGrace As Boolean, ByVal PlusOne As Boolean,
      Optional ByRef Arrears As Decimal = 0, Optional ByRef LateCurrent As Decimal = 0, Optional ByRef Late30 As Decimal = 0,
      Optional ByRef Late60 As Decimal = 0, Optional ByRef Late90 As Decimal = 0,
      Optional ByRef TotDue As Decimal = 0, Optional MonthsBehind As Single = 0,
      Optional ByRef PastDueDate As Date = Nothing,
      Optional ByRef Late90Past120 As Decimal = 0, Optional ByRef DaysLate As Integer = 0,
      Optional ByRef BillingDate As Date = Nothing,
      Optional ByRef NextDueDate As String = "") As Decimal

        Dim AdjustDays As Integer, CD As Date, TotalPeriods As Integer
        Dim PeriodsElapsed As Integer
        Dim ActualPeriodsElapsed As Integer
        Dim PeriodsOver As Integer
        Dim LastPaymentDate As Date
        Dim GraceDate As Date, Grace As Integer, Over As Boolean
        Dim PayAmount As Decimal, LateAmt As Decimal
        '###GRACE10 This is certainly broken for Grace > 10



        '  Debug.Print "ComputeAgeing Date=" & CurrentDate & ", FirstPay=" & FirstPay & ", Months=" & Months & ", PaidBy=" & PaidBy & ", Weekly=" & Weekly & ", Payment=" & Payment & ", TotPaid=" & TotPaid & ", Balance=" & Balance & ", wGrace=" & wGrace & " (" & StoreSettings.GracePeriod & ")"
        '  LogFile "ageing.txt", "ComputeAgeing Date=" & CurrentDate & ", FirstPay=" & FirstPay & ", Months=" & Months & ", PaidBy=" & PaidBy & ", Weekly=" & Weekly & ", Payment=" & Payment & ", TotPaid=" & TotPaid & ", Balance=" & Balance & ", wGrace=" & wGrace & " (" & StoreSettings.GracePeriod & ")"

        On Error GoTo ComputationError

        DaysLate = 0
        Arrears = 0
        LateCurrent = 0
        Late30 = 0
        Late60 = 0
        Late90 = 0
        Late90Past120 = 0

        Grace = IIf(wGrace, AdjustedGracePeriod(PaidBy), 0) ' For consistency, we modify grace for PaidBy=1.. e.g. For grace=10, Late date is 10th, 20th, 30th.. intead of 11th
        GraceDate = DateAdd("d", -Grace, CurrentDate)
        'Debug.Print "GraceDate=" & GraceDate

        ' convert due date to late date
        AdjustDays = 0 ' IIf(IsBoyd, 0, 1) ' 0 makes it show up late on the 10th.  1 makes it show up late on the 11th.
        BillingDate = Format(FirstPay, "MM/" & (PaidBy + AdjustDays) & "/YYYY")
        'BFH20170216 - This shouldn't happen...  PaidBy isn't == Day(FirstPay)...  But, it does.
        If DateBefore(BillingDate, FirstPay, False) Then BillingDate = DateAdd("m", 1, BillingDate)
        LastPaymentDate = DateAdd("m", Months, FirstPay)
        'Debug.Print "BillingDate=" & BillingDate

        ' The customers are billed monthly, but late payments are aged by 30 day increments.
        ' To properly calculate, then, we need to know how many months have passed since the
        ' first due date (txtFirstPay).

        If Not Weekly Then
            TotalPeriods = Months
            PeriodsElapsed = CountMonths(BillingDate, GraceDate, True)
            If PeriodsElapsed < 0 Then PeriodsElapsed = 0
            PastDueDate = DateAdd("d", Grace, DateAdd("m", PeriodsElapsed - 1, BillingDate))
        Else
            CD = DateAdd("m", Months - 1, FirstPay)
            TotalPeriods = DateDiff("ww", BillingDate, CD) + 1

            If DateDiff("d", GraceDate, BillingDate) <= 0 Then
                CD = DateValue(BillingDate)
                Do While DateDiff("d", CD, GraceDate) > 0
                    CD = DateAdd("ww", 1, CD)
                    PeriodsElapsed = PeriodsElapsed + 1
                Loop
            End If
            PastDueDate = DateAdd("d", Grace, DateAdd("ww", PeriodsElapsed - 1, BillingDate))
        End If

        '  If PlusOne Then PeriodsElapsed = PeriodsElapsed - 1

        ActualPeriodsElapsed = PeriodsElapsed
        If PeriodsElapsed > TotalPeriods Then PeriodsElapsed = TotalPeriods : Over = True
        TotDue = PeriodsElapsed * Payment ' This is correctly calculated by months.
        If TotDue > Financed Then TotDue = Financed

        Arrears = TotDue - TotPaid
        ' BFH20090404 - "There are possibilities of a negative Arrearages on the screen and that is when a customer overpays.  Check Treehouse Acc 2550.  For example Should show -100.00 arrearages."
        ' BFH20090326 - the lateamt < 0 check was commented out.. no idea why... it is put back in...
        '  If LateAmt < 0 Then LateAmt = 0 ' BFH20060814
        If Payment = 0 Then MonthsBehind = 0 Else MonthsBehind = CurrencyFormat(Arrears / Payment)

        If Arrears > Financed Then Arrears = Financed

        NextDueDate = DateAdd("m", 1, PastDueDate)
        If Arrears <> 0 Then
            If Payment >= 1 Then
                NextDueDate = DateAdd("m", RoundDn(-Arrears / Payment), NextDueDate)
            End If
        ElseIf Arrears = 0 Then
            NextDueDate = DateAdd("m", 1, PastDueDate)
            '  Else
            '    NextDueDate = "In Arrears"
        End If
        If DateAfter(NextDueDate, DateAdd("m", Months - 1, FirstPay), False) Then
            NextDueDate = ""
        ElseIf DateBefore(NextDueDate, FirstPay, False) Then
            NextDueDate = ""
        End If

        If Arrears <= 0 Then
            LateCurrent = 0
            Late30 = 0
            Late60 = 0
            Late90 = 0
        ElseIf PeriodsElapsed > 0 And Not Over And Payment > 0 Then
            ' We need a new method, since it's possible (at just the right date) to owe nothing
            ' in a 30 day period.  Here it is:
            ' We have the late amount.
            ' LateAmt/Payment, rounded up is how many months late the payment is.
            ' For each month backward, starting with the current date, add the late amount to the proper display box.
            ' The last amount may be a partial payment.

            ' Figure the last payment date.

            Dim DateDue As Date
            DateDue = Format(GraceDate, "mm/" & (PaidBy + AdjustDays) & "/yyyy")
            If DateAndTime.Day(GraceDate) < PaidBy + AdjustDays Then DateDue = DateAdd("m", -1, DateDue)

            LateAmt = Arrears
            Do While LateAmt > 0
                DaysLate = DateDiff("d", DateDue, GraceDate) ' This is how many days late the payment is.
                If LateAmt >= Payment Then
                    PayAmount = Payment
                    LateAmt = LateAmt - PayAmount
                Else
                    PayAmount = LateAmt
                    LateAmt = 0
                End If
                If DaysLate < 31 Then
                    LateCurrent = LateCurrent + PayAmount
                ElseIf DaysLate < 61 Then
                    Late30 = Late30 + PayAmount
                ElseIf DaysLate < 91 Then
                    Late60 = Late60 + PayAmount
                ElseIf DaysLate >= 91 Then
                    Late90 = Late90 + PayAmount
                    If DaysLate >= 120 Then Late90Past120 = Late90Past120 + PayAmount
                End If
                DateDue = DateAdd("m", -1, DateDue)
            Loop

            If MonthsBehind = 0 Then    'allows for late amount less than payment
                LateCurrent = 0
                Late30 = 0
            End If


        Else 'over the contract length

            PeriodsOver = ActualPeriodsElapsed - Months
            If CLng(Payment) = 0 Then
                MonthsBehind = PeriodsOver
                DaysLate = DateDiff("d", LastPaymentDate, CurrentDate)
            ElseIf CLng(Payment) = 0 Then
                MonthsBehind = PeriodsOver + ((TotDue * 100) \ (Payment * 100))
                DaysLate = DateDiff("d", LastPaymentDate, CurrentDate) + ((TotDue * 100) \ (Payment) * 100) * 30
            Else
                ' The above 2 cases are just when payment (or clng of it) is zero.. div by zero
                MonthsBehind = PeriodsOver + TotDue \ Payment
                DaysLate = DateDiff("d", LastPaymentDate, CurrentDate) + TotDue \ Payment * 30
            End If

            '    DateDue = Format(GraceDate, "mm/" & (PaidBy + AdjustDays) & "/yyyy")
            '    If Day(GraceDate) < PaidBy + AdjustDays Then DateDue = DateAdd("m", -1, DateDue)
            '    DaysLate = DateDiff("d", DateDue, GraceDate) ' This is how many days late the payment is.
            '
            LateCurrent = 0
            Late30 = 0
            Late60 = 0
            Late90 = 0 'Arrears
            Late90Past120 = 0 'Late90

            If PeriodsOver = 1 Then
                LateCurrent = Arrears
            ElseIf PeriodsOver = 2 Then
                Late30 = Arrears
            ElseIf PeriodsOver = 3 Then
                Late60 = Arrears
            ElseIf PeriodsOver = 4 Then
                Late90 = Arrears
            ElseIf PeriodsOver > 4 Then
                Late90 = Arrears
                Late90Past120 = Arrears
            End If

            '    If PeriodsElapsed >= 1 And PeriodsElapsed < 2 Then
            '      LateCurrent = Arrears
            '    ElseIf PeriodsElapsed >= 2 And PeriodsElapsed < 3 Then
            '      Late30 = Arrears
            '    ElseIf PeriodsElapsed >= 3 And PeriodsElapsed < 4 Then
            '      Late60 = Arrears
            '    ElseIf PeriodsElapsed >= 4 Then
            '      Late90 = Arrears
            '    End If
        End If

        ComputeAgeing = Arrears

        Exit Function

ComputationError:

        If IsDevelopment() Then
            MessageBox.Show("Computation error" & vbCrLf & Err.Description)
        End If
        Resume Next
    End Function

    Public Function GetPaymentHistoryEquifax(ByVal ArNo As String, Optional ByRef RunAsDate As String = "", Optional ByVal nMonths As Integer = 24, Optional ByVal StoreNo As Integer = 0) As String
        ' Pulled from Metro426.PDF
        ' Pages:  33 (PDF 38) - Field Definition, 69 (PDF 74) - Examples
        '
        ' ***** Payment History Profile *****
        'Contains up to 24 months of consecutive payment activity for the previous 24 calendar months prior to the
        'activity date being reported. Report one month’s payment record in each byte from the left to right in most
        'recent to least recent order. The first byte should represent the previous month’s status. Values available:
        '
        '0 = 0 payments past due (current account)
        '1 = 30 - 59 days past due date
        '2 = 60 - 89 days past due date
        '3 = 90 - 119 days past due date
        '4 = 120 - 149 days past due date
        '5 = 150 - 179 days past due date
        '6 = 180 or more days past due date
        'B = No payment history available prior to this time - may not be embedded within other values.
        'D = No payment history available this month. A "D" may be embedded in the payment pattern.
        'E = Zero balance and current account
        'G = Collection
        'H = Foreclosure
        'J = Voluntary Surrender
        'K = Repossession
        'L = Charge - Off
        'If a full 24 months of history are not available for reporting, the ending positions of this field should be filled with Bs.
        'No other values are acceptable in this field.
        'Reporting of the Payment History Profile provides a method for automated correction of erroneously reported history.
        'Exhibit 5 provides examples of reporting payment history.
        'Note: First-time reporters should refer to Commonly
        '
        '
        'The Account Status, which is reported in field 17A of the Base Segment, contains the status of the account as of the current month being reported.
        'Field 18 contains up to 24 months of consecutive history prior to the current month.
        '
        'Examples of Account Status and Payment History Profile:
        'A. Status Code = 11; Billing Date = 01/31/2000
        'Field 18 = 000011000000EEEE0000BBBB
        'In the above example, field 18 data represents, from left to right, 12/31/1999 through 01/31/1998.
        'The E’s indicate that the account was current with a zero balance in 12/1998, 11/1998, 10/1998
        'and 09/1998. The B’s indicate that no payment history was available prior to 05/1998, which was
        'most likely the date the account was opened.
        'B. Status Code = 80; Billing Date = 01/31/2000; Date of First Delinquency = 08/31/1999
        '
        'Field 18 = 2211100000DD000101000000
        'In the above example, field 18 data represents, from left to right, 12/31/1999 through 01/31/1998.
        'The D’s indicate that no payment history was available for 02/1999 or 01/1999.
        'C. Status Code = 11; Billing Date = 01/31/2000
        '
        'Field 18 = EEEEEEEEE000EEEE0000EE00
        'In the above example, field 18 data represents, from left to right, 12/31/1999 through 01/31/1998.
        'The E’s indicate that the account was current with a zero balance from 12/1999 through
        '04/1999, from 12/1998 through 09/1998, and from 04/1998 to 03/1998. The account was current
        '(and active) during the other months.
        'D. Status Code = 97; Billing Date = 01/31/2000; Date of First Delinquency = 06/30/1998
        '
        'Field 18 = LLGGGGGGGG66654332100010
        'In the above example, field 18 represents, from left to right, 12/31/1999 through 01/31/1998. The
        'L’s indicate that the account was a charge-off from 12/1999 through 11/1999, and the G’s
        'indicate that the account was a collection from 10/1999 through 03/1999.
        'Note: The Date of First Delinquency (06/30/1998)
        '

        Dim I As Integer, RD As Date, D As Date
        Dim FirstPay As Date, WriteOffDate As Date
        Dim ContractDate As Date, FirstContractDate As Date, IsAddOn As Boolean
        Dim R As ADODB.Recordset
        Dim R2 As ADODB.Recordset
        Dim Months As Integer
        Dim TotPaid As Decimal, PaidBy As Integer, Weekly As Boolean, Payment As Decimal
        Dim Financed As Decimal, Balance As Decimal, wGrace As Boolean, PlusOne As Boolean
        Dim AmtDue As Decimal, Cnt As Integer, C As String, T As String, DeferredLC As Decimal
        Dim InSale As Boolean
        Dim CurStatus As String, JustClosed As Boolean
        Dim LCTotal As Decimal
        Dim IsLastPayment As Boolean, LastPaymentAmt As Decimal, IsPostTerm As Boolean

        '  Dim Arrears as decimal, L0 as decimal, L30 as decimal, L60 as decimal, L90 as decimal
        '  Dim Totdue as decimal, MonthsBehind As Single, PastDueDate As Date
        Dim S As String
        Dim SS As String

        Dim Res As String
        Dim AddOnAfter As Integer

        ' The Credit Co didn't like us simply saying "Not Availble" for 24 months in every situation and said we had to calculate it.
        '  GetPaymentHistory = String(24, "D")
        '  Exit Function

        R = GetRecordsetBySQL("SELECT * FROM [InstallmentInfo] WHERE [ArNo]=""" & ProtectSQL(ArNo) & """", , GetDatabaseAtLocation(StoreNo))
        If R.EOF Then Exit Function

        ContractDate = GetArNoContractDate(ArNo, StoreNo)
        FirstContractDate = GetArNoFirstContractDate(ArNo, StoreNo)
        IsAddOn = Not DateEqual(ContractDate, FirstContractDate)

        If IsDate(RunAsDate) Then
            RD = DateValue(RunAsDate)
        Else
            RD = Today
        End If

        SS = ""
        SS = SS & "SELECT * FROM [Transactions] "
        SS = SS & "WHERE 1=1 "
        SS = SS & " AND ([TransDate] < #" & RD & "#) "
        '  SS = SS & "AND " & SQLDateRange("TransDate", DayAdd(ContractDate, 1), RD)
        SS = SS & "AND [ArNo]='" & ArNo & "' "
        SS = SS & "ORDER BY TransDate ASC, TransactionID ASC"
        R2 = GetRecordsetBySQL(SS, , GetDatabaseAtLocation(StoreNo))

        ' find last expected payment..
        D = DateValue(Month(RD) & "/" & Val(R("LateDueOn")) & "/" & Year(RD))
        'BFH20170522 - Grace Period ONLY applies to Late Charges
        '  D = DayAdd(D, StoreSettings.GracePeriod)
        Do While DateAfter(D, RD)
            D = MonthAdd(D, -1)
        Loop
        ' This one isn't calculabe, so we need to go back one more....
        D = MonthAdd(D, -1)
        RunAsDate = D

        ' now we get to 24 months ago and begin processing...  24, including the most recent expected payment
        D = MonthAdd(D, -(nMonths - 1))

        ' But, first, we also record a few details about the account
        FirstPay = DateValue(R("FirstPayment").Value)
        Months = IfNullThenZero(R("Months").Value)

        Weekly = IfNullThenNilString(R("Period").Value) <> "M"
        Payment = IfNullThenZeroCurrency(R("Permonth").Value)
        TotPaid = IfNullThenZeroCurrency(R("TotPaid").Value)
        PaidBy = IfNullThenZero(R("LateDueOn").Value)
        Financed = IfNullThenZeroCurrency(R("Financed").Value)
        Balance = IfNullThenZeroCurrency(R("balance").Value)
        WriteOffDate = IfNullThenZeroDate(R("WriteOffDate").Value)
        'If CLng(WriteOffDate) = 0 Then WriteOffDate = MonthAdd(DateValue(RunAsDate), nMonths * 2)
        If WriteOffDate = " 01/01/0001 0:00:00" Then WriteOffDate = MonthAdd(DateValue(RunAsDate), nMonths * 2)

        Res = ""
        AmtDue = 0

        AmtDue = Payment * DateDiff("m", FirstPay, D)
        If AmtDue < 0 Then AmtDue = 0

        CurStatus = "V" ' Default current status (AR status, not Equifax status)

        ' going from 24 months ago...
        For I = 1 To nMonths
            'If IsDevelopment And I = nMonths Then Stop
            If DateBefore(D, FirstPay, False) Then
                If IsAddOn And DateAfter(D, FirstContractDate) Then
                    'BFH20170605
                    ' Right now, we don't keep track of the previous contract in an Add-On.  The account could have been open
                    ' for several months, and then added onto, and if so, we have lost the ability to calculate any
                    ' of these prior months.  In this situation, we report a 'D' indicator, saying that payment
                    ' history is unavailable for these months.
                    ' Currently, the alternatives to this would be to SAVE previously exported payment history fields
                    ' (which fails for existing accounts) or save the [InstallmentInfo] fields (to where?) so that
                    ' we can reverse calculate the contracts before Add-On (which again fails for existing accounts).
                    'BFH20170614
                    ' Jerry has authorized the modification to store old contract info after Add On..
                    Res = "D" & Res
                Else
                    ' If the account isn't opened yet, mark it with a 'B' (blank).
                    Res = "B" & Res
                End If
            ElseIf DateAfter(D, WriteOffDate) Then
                Res = "L" & Res
            Else
                IsLastPayment = ((DateDiff("m", FirstPay, D) + 1) = Months)
                IsPostTerm = ((DateDiff("m", FirstPay, D) + 1) > Months)
                If IsLastPayment Then
                    LastPaymentAmt = Financed - (Payment * (Months - 1))
                    AmtDue = AmtDue + LastPaymentAmt + DeferredLC
                ElseIf IsPostTerm Then
                    AmtDue = AmtDue + DeferredLC
                Else
                    AmtDue = AmtDue + Payment + DeferredLC
                End If
                DeferredLC = 0
                If AmtDue > Financed Then AmtDue = Financed

                'If IsDevelopment And Month(D) = 7 And Year(D) >= 2017 And ArNo = "2000" Then Stop

                If Not R2.EOF Then
                    Do While DateBefore(DateValue(R2("TransDate").Value), MonthAfter(D), False)
                        T = IfNullThenNilString(R2("Type").Value)

                        If ArTypeIsStatusChange(T) Then
                            Select Case R2("Type").Value
                                Case arPT_stReO : CurStatus = arST_Open
                                Case arPT_stClo : CurStatus = arST_Clos : JustClosed = True
                                Case arPT_stVoi : CurStatus = arST_Void
                                Case arPT_stWtO : CurStatus = arST_Writ
                                Case arPT_stRep : CurStatus = arST_Repo
                                Case arPT_stLeg : CurStatus = arST_Lega
                                Case arPT_stBkr : CurStatus = arST_Bank
                                Case Else : CurStatus = arST_Open
                            End Select
                        End If

                        If ArTypeIsNewSale(T) Then
                            If Not InSale Then
                                InSale = True
                            Else
                                ' if this is an add-on, we don't have any way to track before the add-on,
                                ' so we will report no payment history before new sale date...
                                ' As if the account as new from the add on
                                ' We need to add all the B's we would have incurred.
                                'BFH20170614 -
                                '              Res = String(I - 1, "B") ' always -1 here, because this loop will add one in the end
                                LCTotal = 0
                                DeferredLC = 0
                                AddOnAfter = DateDiff("m", IfNullThenZeroDate(R2("TransDate")), D)
                                AmtDue = Payment * AddOnAfter
                            End If
                        End If

                        If Not ArTypeIsContract(T) Or ArTypeIsPayoff(T) Or T = arPT_Prv Then
                            If T = arPT_L_C And DateAfter(DateAdd("d", 5, R2("TransDate").Value), D) Then
                                DeferredLC = DeferredLC + IfNullThenZeroCurrency(R2("Charges").Value) - IfNullThenZeroCurrency(R2("Credits").Value)
                            Else
                                AmtDue = AmtDue + IfNullThenZeroCurrency(R2("Charges").Value) - IfNullThenZeroCurrency(R2("Credits").Value)
                            End If
                            If ArTypeIsLCAdjust(T) Then LCTotal = LCTotal + IfNullThenZeroCurrency(R2("Charges").Value) - IfNullThenZeroCurrency(R2("Credits").Value)
                        End If
                        R2.MoveNext()
                        If R2.EOF Then GoTo NoMoreTransactions
                    Loop
                End If
NoMoreTransactions:

                If CurStatus <> "O" Then
                    Select Case CurStatus
                        Case arST_Clos
                            If JustClosed Then
                                JustClosed = False
                                GoTo DoNumber ' Bad code?  Or convinient shortcut?
                            End If
                            Res = "E" & Res   ' Zero Balance, Up to date
                        Case arST_Void : Res = "D" & Res
                        Case arST_Writ : Res = "L" & Res   ' Charge-off
                        Case arST_Repo : Res = "K" & Res   ' Repo
                        Case arST_Lega : Res = "G" & Res   ' Collection
                        Case arST_Bank : Res = "G" & Res
                    End Select
                ElseIf AmtDue <= LateGraceAmt Then    ' $1 grace
                    Res = "0" & Res
                Else
DoNumber:
                    Cnt = FitRange(0, RoundUp((AmtDue - LCTotal) / Payment), 6)
                    C = IIf(Cnt > 6, "6", "" & Cnt)
                    Res = C & Res
                End If
            End If

            D = MonthAdd(D, 1)
        Next

        Res = ArCreditHistoryFillInAddOns(ArNo, StoreNo, Res, nMonths, RunAsDate)
        ExecuteRecordsetBySQL("UPDATE [InstallmentInfo] SET [PaymentHistoryProfile]='" & Res & "' WHERE [ArNo]=""" & ProtectSQL(ArNo) & """", , GetDatabaseAtLocation(StoreNo))
        GetPaymentHistoryEquifax = Res
    End Function

    Private Function ArCreditHistoryFillInAddOns(ByVal ArNo As String, ByVal StoreNo As Integer, ByVal PHP As String, Optional ByVal tLen As Integer = 24, Optional ByVal RunAsDate As Date = Nothing) As String
        Dim AddOnAcc As String, AddOnPHP As String
        Dim EffDate As String, D1 As Date, D2 As Date, N As Integer
        Dim C As String
        Dim I As Integer

        ArCreditHistoryFillInAddOns = PHP
        If Len(ArCreditHistoryFillInAddOns) <> tLen Then Exit Function

        AddOnAcc = GetArNoLastAddOnAccountNo(ArNo, StoreNo)
        If AddOnAcc = "" Then Exit Function

        AddOnPHP = GetArNoLastAddOnPHP(AddOnAcc, StoreNo)
        If AddOnPHP = "" Or Len(AddOnPHP) <> tLen Then Exit Function

        EffDate = Trim("" & GetValueBySQL("SELECT [WriteOffDate] FROM [InstallmentInfo] WHERE [ArNo]=""" & ProtectSQL(AddOnAcc) & """", , GetDatabaseAtLocation(StoreNo)))
        If Not IsDate(EffDate) Then Exit Function

        D1 = DateValue(EffDate)
        D2 = DateValue(RunAsDate)
        If DateBefore(D1, D2, False) Then
            N = 0
            Do Until Month(D1) = Month(D2) And Year(D1) = Year(D2)
                D1 = DateAdd("m", 1, D1)
                AddOnPHP = Left("D" & AddOnPHP, tLen)
                N = N + 1
                If N > tLen Then Exit Do
            Loop
        End If

        For I = 1 To tLen
            If Mid(ArCreditHistoryFillInAddOns, I, 1) = "D" Then
                C = Mid(AddOnPHP, I, 1)
                If C = "D" Then C = "B"
                If I = 1 Then
                    ArCreditHistoryFillInAddOns = C & Mid(ArCreditHistoryFillInAddOns, 2)
                ElseIf I = tLen Then
                    ArCreditHistoryFillInAddOns = Left(ArCreditHistoryFillInAddOns, tLen - 1) & C
                Else
                    ArCreditHistoryFillInAddOns = Left(ArCreditHistoryFillInAddOns, I - 1) & C & Mid(ArCreditHistoryFillInAddOns, I + 1)
                End If
            End If
            If IsDevelopment() And Len(ArCreditHistoryFillInAddOns) <> tLen Then Stop
        Next
    End Function

    Public Function ArAddOnCreateContractHistoryAccount(ByVal ArNo As String, Optional ByVal StoreNo As Integer = 0) As String
        Dim D As String
        ArAddOnCreateContractHistoryAccount = ArAddOnContractHistoryAccountNo(ArNo, StoreNo)                ' Get and return the Record Keeping ArNo
        D = Today                                                                                            ' We need to know the effective date of this PHP
        GetPaymentHistoryEquifax(ArNo, D)                                                                    ' Record the current Payment History Profile
        If Not DuplicateArInstallmentInfoRecord(StoreNo, ArNo, ArAddOnCreateContractHistoryAccount, arST_Void) Then   ' Duplicate the record to the new ArNo
            MessageBox.Show("Could not create AddOn Record Account: " & ArAddOnCreateContractHistoryAccount)            ' If failure, notify
        Else                                                                                                ' else, Make the new Record Keeping ArNo VOID and record PHP date
            ExecuteRecordsetBySQL("UPDATE [InstallmentInfo] SET [Status]='" & arST_Void & "', [WriteOffDate]=#" & D & "#, [PaymentHistoryProfile]='' WHERE [ArNo]=""" & ProtectSQL(ArAddOnCreateContractHistoryAccount) & """", , GetDatabaseAtLocation(StoreNo))
            AddNewARTransactionExisting(StoreNo, ArAddOnCreateContractHistoryAccount, Today, arPT_stVoi, 0, 0, "Account Created VOID for A/R Add-On, Account #" & ArNo)
        End If
    End Function

    Public Function ArAddOnToNewCloseOutAccount(ByVal StoreNo As Integer, ByVal ArNo As String, ByVal NewArNo As String, ByVal Balance As Decimal) As String
        Dim D As String
        ' Close out the old account, marking a balance credit and recording the transfer to another account, as well as changing status...
        AddNewARTransactionExisting(StoreNo, ArNo, Today, arPT_crPri, 0, Balance, "Added onto Account: " & NewArNo)
        ExecuteRecordsetBySQL("UPDATE [InstallmentInfo] Set [TotPaid]=[Financed], [Balance]=0, [LateChargeBal]=0, [Status]='" & arST_Clos & "' WHERE [ArNo]=""" & ProtectSQL(ArNo) & """", , GetDatabaseAtLocation(StoreNo))
        AddNewARTransactionExisting(StoreNo, ArNo, Today, arPT_stClo, 0, 0, "Account Closed via Add-On-To, Account: " & NewArNo)
    End Function

    Public Function ArAddOnContractHistoryAccountNo(ByVal ArNo As String, Optional ByVal StoreNo As Integer = 0) As String
        Dim N As Integer, X As String
        N = 1
        Do While True
            X = ArNo & ArNo_AddOnRecordIndicator & N
            If Not ArNoExists(X, StoreNo) Then
                ArAddOnContractHistoryAccountNo = X
                Exit Function
            End If
            N = N + 1
        Loop
    End Function

    Public Function DuplicateArInstallmentInfoRecord(ByVal StoreNo As Integer, ByVal ArNo As String, ByVal NewArNo As String, Optional ByVal NewStatus As String = "") As Boolean
        ' Primarily to be used for Add-On's.
        Dim R As ADODB.Recordset

        If Not ArNoExists(ArNo, StoreNo) Then Exit Function
        If ArNoExists(NewArNo, StoreNo) Then Exit Function

        On Error Resume Next
        R = GetRecordsetBySQL("SELECT * FROM [InstallmentInfo] WHERE [ArNo]=""" & ProtectSQL(ArNo) & """", , GetDatabaseAtLocation(StoreNo))

        If NewStatus = "" Then NewStatus = IfNullThenNilString(R("Status"))

        AddNewARInstallmentInfo(StoreNo, NewArNo,
  IfNullThenNilString(R("LastName").Value), IfNullThenNilString(R("Telephone").Value),
  IfNullThenZero(R("MailIndex").Value), IfNullThenZeroCurrency(R("Financed").Value), IfNullThenZeroCurrency(R("PerMonth").Value), IfNullThenZero(R("Months").Value),
  IfNullThenZeroDouble(R("Rate").Value), IfNullThenZero(R("LateDueOn").Value), IfNullThenZeroCurrency(R("LateCharge").Value),
  IfNullThenNullDate(R("DeliveryDate").Value), IfNullThenNullDate(R("FirstPayment").Value), IfNullThenZero(R("CashOpt").Value),
  IfNullThenZeroCurrency(R("TotPaid").Value), IfNullThenZeroCurrency(R("Balance").Value), IfNullThenZeroCurrency(R("LateChargeBal").Value),
  NewStatus,
  IfNullThenZeroCurrency(R("Interest").Value), IfNullThenZeroCurrency(R("Life").Value), IfNullThenZeroCurrency(R("Accident").Value), IfNullThenZeroCurrency(R("Prop").Value),
  IfNullThenNilString(R("WriteOffDate").Value), IfNullThenNilString(R("SendNotice").Value), IfNullThenNilString(R("LastMetro426Status").Value),
  IfNullThenZeroCurrency(R("InterestSalesTax").Value), IfNullThenZeroDouble(R("APR").Value), IfNullThenZero(R("Period").Value),
  IfNullThenZero(R("LifeType").Value), IfNullThenZero(R("Satisfied").Value), IfNullThenNilString(R("SatisfiedDate").Value),
  IfNullThenZeroCurrency(R("IUI").Value), IfNullThenZero(R("LastNotice").Value), IfNullThenNilString(R("LastLateCharge").Value), IfNullThenNilString(R("PaymentHistoryProfile").Value))

        If Not ArNoExists(NewArNo, StoreNo) Then Exit Function
        DuplicateArInstallmentInfoRecord = True
    End Function

    Public Function AddNewARTransactionExisting(ByVal StoreNo As Integer,
  ByVal ArNo As String, Optional ByVal TransDate As String = "",
  Optional ByVal Typee As String = "",
  Optional ByVal Charges As Decimal = 0, Optional ByVal Credits As Decimal = 0,
  Optional ByVal Receipt As String = "",
  Optional ByVal AdjustArBalance As Boolean = False,
  Optional ByVal AdjustTotPaid As Boolean = False) As Boolean
        ' Adds onto an existing account, automatically continuing mail & name, and calculating balance.

        Dim N As String, M As Integer, B As Integer, C As Decimal
        N = GetValueBySQL("SELECT [LastName] FROM [InstallmentInfo] WHERE [ArNo]=""" & ProtectSQL(ArNo) & """", , GetDatabaseAtLocation(StoreNo))
        M = Val(GetValueBySQL("SELECT [MailIndex] FROM [InstallmentInfo] WHERE [ArNo]=""" & ProtectSQL(ArNo) & """", , GetDatabaseAtLocation(StoreNo)))
        C = GetPrice(GetValueBySQL("SELECT TOP 1 [Balance] FROM [Transactions] WHERE [ArNo]='" & ArNo & "' ORDER BY [TransDate] DESC, [TransactionID] DESC", , GetDatabaseAtLocation(StoreNo)))
        If Not IsDate(TransDate) Then TransDate = Today
        AddNewARTransactionExisting = AddNewARTransaction(StoreNo, ArNo, N, TransDate, M, Typee, Charges, Credits, C + Charges - Credits, Receipt, AdjustArBalance, AdjustTotPaid)
    End Function

    Public Function AddNewARInstallmentInfo(ByVal StoreNo As Integer,
  ByVal ArNo As String, Optional ByVal LastName As String = "", Optional ByVal Telephone As String = "",
  Optional ByVal MailIndex As Integer = 0, Optional ByVal Financed As Decimal = 0, Optional ByVal PerMonth As Decimal = 0, Optional ByVal Months As Integer = 0,
  Optional ByVal Rate As Double = 0, Optional ByVal LateDueOn As Integer = 0, Optional ByVal LateCharge As Decimal = 0,
  Optional ByVal DeliveryDate As String = "", Optional ByVal FirstPayment As String = "", Optional ByVal CashOpt As Integer = 0,
  Optional ByVal TotPaid As Decimal = 0, Optional ByVal Balance As Decimal = 0, Optional ByVal LateChargeBal As Decimal = 0,
  Optional ByVal Status As String = "", Optional ByVal INTEREST As Decimal = 0, Optional ByVal Life As Decimal = 0, Optional ByVal Accident As Decimal = 0, Optional ByVal Prop As Decimal = 0,
  Optional ByVal WriteOffDate As String = "", Optional ByVal SendNotice As String = "", Optional ByVal LastMetro426Status As String = "",
  Optional ByVal InterestSalesTax As Decimal = 0, Optional ByVal APR As Double = 0, Optional ByVal Period As Integer = 0, Optional ByVal LifeType As Integer = 0,
  Optional ByVal Satisfied As Integer = 0, Optional ByVal SatisfiedDate As String = "",
  Optional ByVal IUI As Decimal = 0, Optional ByVal LastNotice As Integer = 0,
  Optional ByVal LastLateCharge As String = "", Optional ByVal PaymentHistoryProfile As String = "") As Boolean
        '' Adds a AR Transaction record regardless of any other records in the table.
        'ArNo  LastName  Telephone
        'MailIndex Financed  PerMonth  Months
        'Rate  LateDueOn LateCharge
        'DeliveryDate  FirstPayment  CashOpt
        'TotPaid Balance LateChargeBal
        'Status  Interest  Life  Accident  Prop
        'WriteOffDate  SendNotice  LastMetro426Status
        'InterestSalesTax  APR Period  LifeType
        'Satisfied SatisfiedDate IUI LastNotice
        'LastLateCharge
        Dim S As String
        S = ""
        S = S & "INSERT INTO [InstallmentInfo] "
        S = S & "(ArNo, LastName, Telephone, MailIndex, Financed, PerMonth, Months, Rate, LateDueOn, LateCharge, DeliveryDate, FirstPayment, CashOpt, TotPaid, Balance, LateChargeBal, Status, Interest, Life, Accident, Prop, WriteOffDate, SendNotice, LastMetro426Status, InterestSalesTax, APR, Period, LifeType, Satisfied, SatisfiedDate, IUI, LastNotice, LastLateCharge, PaymentHistoryProfile) "
        S = S & " VALUES "
        S = S & "("
        'ArNo  LastName  Telephone
        S = S & """" & ProtectSQL(ArNo) & """, "
        S = S & """" & ProtectSQL(LastName) & """, "
        S = S & """" & ProtectSQL(Telephone) & """, "
        'MailIndex Financed  PerMonth  Months
        S = S & MailIndex & ", "
        S = S & SQLCurrency(Financed) & ", "
        S = S & SQLCurrency(PerMonth) & ", "
        S = S & Months & ", "
        'Rate  LateDueOn LateCharge
        S = S & Rate & ", "
        S = S & LateDueOn & ", "
        S = S & SQLCurrency(LateCharge) & ", "
        'DeliveryDate  FirstPayment  CashOpt
        S = S & SQLDate(DeliveryDate, """") & ", "
        S = S & SQLDate(FirstPayment, """") & ", "
        S = S & CashOpt & ", "
        'TotPaid Balance LateChargeBal
        S = S & SQLCurrency(TotPaid) & ", "
        S = S & SQLCurrency(Balance) & ", "
        S = S & SQLCurrency(LateChargeBal) & ", "
        'Status  Interest  Life  Accident  Prop
        S = S & """" & Left(ProtectSQL(Status), 1) & """, "
        S = S & SQLCurrency(INTEREST) & ", "
        S = S & SQLCurrency(Life) & ", "
        S = S & SQLCurrency(Accident) & ", "
        S = S & SQLCurrency(Prop) & ", "
        'WriteOffDate  SendNotice  LastMetro426Status
        S = S & IIf(Trim(WriteOffDate) = "", "Null", """" & Trim(WriteOffDate) & """") & ", "
        S = S & IIf(SendNotice = "", "Null", """" & ProtectSQL(SendNotice) & """") & ", "
        S = S & """" & Left(ProtectSQL(LastMetro426Status), 2) & """, "
        'InterestSalesTax  APR Period  LifeType
        S = S & SQLCurrency(InterestSalesTax) & ", "
        S = S & APR & ", "
        S = S & Period & ", "
        S = S & LifeType & ", "
        'Satisfied SatisfiedDate IUI LastNotice
        S = S & Satisfied & ", "
        S = S & IIf(IsDate(SatisfiedDate), """" & Trim(SatisfiedDate) & """", "Null") & ", "
        S = S & SQLCurrency(IUI) & ", "
        S = S & LastNotice & ", "
        'LastLateCharge
        S = S & """" & ProtectSQL(LastLateCharge) & """, "
        S = S & """" & Left(ProtectSQL(PaymentHistoryProfile), 24) & """"

        S = S & ")"
        ExecuteRecordsetBySQL(S, , GetDatabaseAtLocation(StoreNo))
        AddNewARInstallmentInfo = True
    End Function

    Public Function AddNewARTransaction(ByVal StoreNo As Integer,
  ByVal ArNo As String, Optional ByVal Name As String = "", Optional ByVal TransDate As String = "",
  Optional ByVal MailIndex As Integer = 0, Optional ByVal Typee As String = "",
  Optional ByVal Charges As Decimal = 0, Optional ByVal Credits As Decimal = 0, Optional ByVal Balance As Decimal = 0,
  Optional ByVal Receipt As String = "",
  Optional ByVal AdjustArBalance As Boolean = False,
  Optional ByVal AdjustTotPaid As Boolean = False) As Decimal
        ' Adds a AR Transaction record regardless of any other records in the table.

        Dim S As String
        Dim tS As String

        S = ""
        S = S & "INSERT INTO [Transactions] "
        S = S & "(ArNo, [LastName], TransDate, MailIndex,Type,Charges,Credits,Balance" & IIf(Receipt = "", "", ", Receipt") & ") "
        S = S & " VALUES "
        S = S & "(""" & ProtectSQL(ArNo) & """, """ & ProtectSQL(Name) & """, #" & DateValue(TransDate) & "#,"
        S = S & MailIndex & ",""" & ProtectSQL(Typee) & """, "
        S = S & CurrencyFormat(Charges, , , True) & ", " & CurrencyFormat(Credits, , , True) & ", " & CurrencyFormat(Balance, , , True)
        If Receipt <> "" Then S = S & ", """ & ProtectSQL(Receipt) & """"
        S = S & ")"
        ExecuteRecordsetBySQL(S, , GetDatabaseAtLocation(StoreNo))

        If AdjustArBalance Then
            tS = "UPDATE [InstallmentInfo] SET [Balance]=" & SQLCurrency(Balance) & " WHERE [ArNo]='" & ArNo & "'"
            ExecuteRecordsetBySQL(tS, , GetDatabaseAtLocation(StoreNo))
        End If

        If AdjustTotPaid Then
            Dim Tp As Decimal
            Tp = GetValueBySQL("SELECT [TotPaid] FROM [InstallmentInfo] WHERE [ArNo]='" & ArNo & "'", , GetDatabaseAtLocation(StoreNo))
            Tp = Tp + Credits
            tS = "UPDATE [InstallmentInfo] SET [TotPaid]=" & SQLCurrency(Tp) & " WHERE [ArNo]='" & ArNo & "'"
            ExecuteRecordsetBySQL(tS, , GetDatabaseAtLocation(StoreNo))
        End If

        AddNewARTransaction = Balance
    End Function

End Module
