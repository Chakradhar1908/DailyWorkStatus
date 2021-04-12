Imports VBRUN
Public Class DateForm
    ' For printing sales tax
    Dim TotSales As Decimal
    Dim Counter As Integer
    Dim NonTaxable As String
    Dim Taxable As String
    Dim TotTaxable As String
    Dim TotNonTaxable As String
    Dim SaleTaxRate As String
    Dim TaxCode As Integer
    Dim TAXREC As Decimal
    Dim TotTaxRec As Decimal
    Dim NonTaxableSales As Decimal
    Dim TaxTotals As clsHashTable
    Dim TotalTaxRate As Double ' for Canadian sales tax
    Dim CurTaxRate As Double
    Private Const FRM_W1 As Integer = 3825
    Private Const FRM_W2 As Integer = 2020
    Private Const FRM_H1 As Integer = 2300
    Dim Cy3 As Integer

    Private Sub cmdCancel_Click(sender As Object, e As EventArgs) Handles cmdCancel.Click
        'Unload Me
        Me.Close()
    End Sub

    Private Sub DateForm_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        MainMenu.Show()
        modProgramState.Order = ""
        modProgramState.Reports = ""
    End Sub

    Private Sub cmdPrint_Click(sender As Object, e As EventArgs) Handles cmdPrint.Click
        If OrderMode("C") Then
            ' Void order
            Hide()
            Exit Sub
        End If

        If OrderMode("ATR") Then
            If cboStoreSelect.SelectedIndex < 0 Or cboStoreSelect.SelectedIndex > ssMaxStore - 1 Then
                MessageBox.Show("Please select a store.", "WinCDS", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                Exit Sub
            End If
        End If

        If POMode("REC") Then
            'Unload Me
            Me.Close()
        Else
            Ctrls(False)
            OutputObject = Printer
            OutputToPrinter = True
        End If

        If OrderMode("ST") Then
            SalesTax()
        ElseIf OrderMode("ATR") Then     ' Advertising types report.
            AdvertisingReport(dDate.Value, toDate.Value, 1, cboStoreSelect.SelectedIndex + 1, chkGroupByZip.Checked = True, chkSortByZip.Checked = True)
            'DateForm.HelpContextID = 49700
        End If

        Printer.EndDoc()
        Ctrls(True)
    End Sub

    Private Sub SalesTax()
        '----------
        ' if using variable rates do not use default below
        ' As I understand it, a store can only use TAX1 or TAX2, but not both.
        ' This is not consistent with the way we handle sale creation.
        '----------  BFH20060207 -  As of now, we are allowing TAX1 & TAX2
        '                           Why they were disallowed, we don't really know, but
        '                           they probably shouldn't have been
        '  Or, at least that's the plan... this is put on hold until someone needs it..
        '  Audit & Holding table only have 1 tax field... would have to rewrite to use GM Table

        ' Canadian taxes are combined into one field in the audit table.
        ' To separate them back out, we need to know the total rate
        ' as well as the rate currently being processed.
        If IsCanadian() Then
            TotalTaxRate = StoreSettings.SalesTax
            If QuerySalesTax2(1) <> "" Then
                For TaxCode = 1 To GetMaxTaxRate() - 1
                    Debug.Print("Tax Rate " & TaxCode & " = " & QuerySalesTax2Rate(TaxCode))
                    TotalTaxRate = TotalTaxRate + QuerySalesTax2Rate(TaxCode)
                Next
            End If
        End If

        ' Handle the default tax
        TaxTotals = New clsHashTable
        ClearVariables()
        SaleTaxRate = StoreSettings.SalesTax
        CurTaxRate = StoreSettings.SalesTax
        TaxCode = 0
        '  UndeliveredHeading
        ProcessTaxRates(0)
        TaxTotals.Add("00 TotSales", TotSales)
        TaxTotals.Add("00 NonTaxtable", TotNonTaxable)
        TaxTotals.Add("00 Taxable", TotTaxable)
        TaxTotals.Add("00 TotTaxRec", TotTaxRec)

        ' BFH20060207 - Now, we MUST process all tax rates represented in the table..
        ' NOTE: Realistically, the table SHOULD store the actual rate, not the index.
        ' If an entry gets deleted, the whole table gets messed up at this point.
        ' It can only be correct through editing or addition, but never through
        ' deleting.  This could potentially be a very large data-integrity bug, but
        ' adjusting it could be more headache..?
        If QuerySalesTax2(1) <> "" Then
            For TaxCode = 1 To GetMaxTaxRate() - 1
                '      If QuerySalesTax2(TaxCode) <> "" Then
                If OutputToPrinter Then OutputObject.NewPage Else frmPrintPreviewDocument.NewPage()
                ClearVariables()
                SaleTaxRate = QuerySalesTax2(TaxCode)
                CurTaxRate = QuerySalesTax2Rate(TaxCode)
                '        UndeliveredHeading
                ProcessTaxRates(TaxCode + 1)
                '      End If
                TaxTotals.Add(Format(TaxCode + 1, "00") & " TotSales", TotSales)
                TaxTotals.Add(Format(TaxCode + 1, "00") & " NonTaxtable", TotNonTaxable)
                TaxTotals.Add(Format(TaxCode + 1, "00") & " Taxable", TotTaxable)
                TaxTotals.Add(Format(TaxCode + 1, "00") & " TotTaxRec", TotTaxRec)
            Next
        End If

        Order = "STTOT"
        If OutputToPrinter Then
            Printer.NewPage()
        Else
            frmPrintPreviewDocument.NewPage()
        End If

        ClearVariables()
        UndeliveredHeading()
        Order = "ST"

        '<CT>
        Dim Cy As Integer
        Cy = 900
        OutputObject.FontBold = False
        '</CT>
        Dim effTaxCode As Integer
        For TaxCode = 0 To GetMaxTaxRate()
            effTaxCode = TaxCode
            If TaxCode = 0 Then
                'PrintTo(OutputObject, "Default Tax", 1, AlignConstants.vbAlignLeft)
                PrintTo(OutputObject, "Default Tax", 1, AlignConstants.vbAlignLeft,, Cy)
                'PrintTo(OutputObject, StoreSettings.SalesTax, 20, AlignConstants.vbAlignLeft)
                PrintTo(OutputObject, StoreSettings.SalesTax, 20, AlignConstants.vbAlignLeft,, Cy)
                If IsCanadian() Then effTaxCode = 1
            Else
                'PrintTo(OutputObject, "Tax Zone #" & TaxCode, 1, AlignConstants.vbAlignLeft)
                PrintTo(OutputObject, "Tax Zone #" & TaxCode, 1, AlignConstants.vbAlignLeft,, Cy)
                'PrintTo(OutputObject, QuerySalesTax2(TaxCode - 1), 20, AlignConstants.vbAlignLeft)
                PrintTo(OutputObject, QuerySalesTax2(TaxCode - 1), 20, AlignConstants.vbAlignLeft,, Cy)
                If IsCanadian() And TaxCode = 1 Then effTaxCode = 0
            End If

            'PrintTo(OutputObject, CurrencyFormat(TaxTotals.Item(Format(effTaxCode, "00") & " TotSales")), 75, AlignConstants.vbAlignRight)
            PrintTo(OutputObject, CurrencyFormat(TaxTotals.Item(Format(effTaxCode, "00") & " TotSales")), 75, AlignConstants.vbAlignRight,, Cy)
            TotSales = TotSales + GetPrice(TaxTotals.Item(Format(effTaxCode, "00") & " TotSales"))

            'PrintTo(OutputObject, CurrencyFormat(TaxTotals.Item(Format(effTaxCode, "00") & " NonTaxtable")), 91, AlignConstants.vbAlignRight)
            PrintTo(OutputObject, CurrencyFormat(TaxTotals.Item(Format(effTaxCode, "00") & " NonTaxtable")), 91, AlignConstants.vbAlignRight,, Cy)
            TotNonTaxable = TotNonTaxable + GetPrice(TaxTotals.Item(Format(effTaxCode, "00") & " NonTaxtable"))

            'PrintTo(OutputObject, CurrencyFormat(TaxTotals.Item(Format(effTaxCode, "00") & " Taxable")), 106, AlignConstants.vbAlignRight)
            PrintTo(OutputObject, CurrencyFormat(TaxTotals.Item(Format(effTaxCode, "00") & " Taxable")), 106, AlignConstants.vbAlignRight,, Cy)
            TotTaxable = TotTaxable + GetPrice(TaxTotals.Item(Format(effTaxCode, "00") & " Taxable"))

            'PrintTo(OutputObject, CurrencyFormat(TaxTotals.Item(Format(effTaxCode, "00") & " TotTaxRec")), 121, AlignConstants.vbAlignRight, True)
            PrintTo(OutputObject, CurrencyFormat(TaxTotals.Item(Format(effTaxCode, "00") & " TotTaxRec")), 121, AlignConstants.vbAlignRight, True, Cy)
            TotTaxRec = TotTaxRec + GetPrice(TaxTotals.Item(Format(effTaxCode, "00") & " TotTaxRec"))
            Cy = Cy + 200
        Next

        If Not IsDoddsLtd Then ' BFH20130423 they do multiple TAX2's on a sale.  Totals are not necessary.
            Cy = Cy + 200
            OutputObject.Print
            OutputObject.FontBold = True
            'PrintTo(OutputObject, "Totals:", 1, AlignConstants.vbAlignLeft)
            PrintTo(OutputObject, "Totals:", 1, AlignConstants.vbAlignLeft,, Cy)

            'PrintTo(OutputObject, CurrencyFormat(TotSales), 75, AlignConstants.vbAlignRight)
            PrintTo(OutputObject, CurrencyFormat(TotSales), 75, AlignConstants.vbAlignRight,, Cy)
            'PrintTo(OutputObject, CurrencyFormat(TotNonTaxable), 91, AlignConstants.vbAlignRight)
            PrintTo(OutputObject, CurrencyFormat(TotNonTaxable), 91, AlignConstants.vbAlignRight,, Cy)
            'PrintTo(OutputObject, CurrencyFormat(TotTaxable), 106, AlignConstants.vbAlignRight)
            PrintTo(OutputObject, CurrencyFormat(TotTaxable), 106, AlignConstants.vbAlignRight,, Cy)
            'PrintTo(OutputObject, CurrencyFormat(TotTaxRec), 121, AlignConstants.vbAlignRight, True)
            PrintTo(OutputObject, CurrencyFormat(TotTaxRec), 121, AlignConstants.vbAlignRight, True, Cy)
            'PrintTo(OutputObject, "Sale totals may be over-counted if sales have multiple TAX2 rates.", 10, , True)
            PrintTo(OutputObject, "Sale totals may be over-counted if sales have multiple TAX2 rates.", 10, , True, Cy + 200)

            OutputObject.FontBold = False
        End If
    End Sub

    Public Sub UndeliveredHeading()
        Dim Cy As Integer

        OutputObject.FontName = "Arial"
        OutputObject.FontSize = 18
        OutputObject.CurrentY = 100
        OutputObject.FontBold = True

        If OrderMode("ST") Then
            If optDelivered.Checked = True Then
                PrintCentered("Sales Tax Report:  Delivered Sales")
            ElseIf optWritten.Checked = True Then
                Cy = OutputObject.CurrentY
                PrintCentered("Sales Tax Report:  Written Sales")
            End If
        ElseIf OrderMode("NST") Then
            PrintCentered("Non Taxable Sales Report")
        End If

        PrintSet(, 10, 100, "Arial", 8, 0)
        '<CT>
        OutputObject.FontSize = 8
        OutputObject.FontBold = False
        'Cy = Printer.CurrentY
        '</CT>

        If optWritten.Checked = True Then
            OutputObject.CurrentY = Cy
            OutputObject.Print("From: " & " " & DateFormat(dDate.Value) & "To: " & " " & DateFormat(toDate.Value))
            OutputObject.Print("Time: " & " " & Format(Now, "h:mm:ss tt"))
        Else
            Cy = Printer.CurrentY
            OutputObject.Print("From: " & " " & DateFormat(dDate.Value) & "To: " & " " & DateFormat(toDate.Value))
            OutputObject.Print("Time: " & " " & Format(Now, "h:mm:ss tt"))
        End If


        PrintSet(, 10100, 100)
        If OutputToPrinter Then PageNumber = OutputObject.Page
        '<CT>
        'If optWritten.Checked = True Then
        'Else
        OutputObject.CurrentX = 10100
        OutputObject.CurrentY = Cy
        'End If

        '</CT>
        OutputObject.Print("Page: " & PageNumber)

        If OrderMode("ST") Then
            PrintToPosition(OutputObject, ShortenSalesTaxDescription(SaleTaxRate), 10100 + Printer.TextWidth("Page: " & PageNumber), AlignConstants.vbAlignRight, True)
        End If

        OutputObject.CurrentY = 500
        PrintCentered(StoreSettings.Name & "    " & StoreSettings.Address & "    " & StoreSettings.City)

        PrintSet(, 0, 700, , 9, 1)

        '<CT>
        Cy = 700
        OutputObject.FontBold = True
        OutputObject.FontSize = 9
        '</CT>
        If OrderMode("ST") Then 'Sales tax
            'PrintTo(OutputObject, "Last Name", 0, AlignConstants.vbAlignLeft, False)
            PrintTo(OutputObject, "Last Name", 0, AlignConstants.vbAlignLeft, False, Cy)
            'PrintTo(OutputObject, "Sale No.", 30, AlignConstants.vbAlignLeft, False)
            PrintTo(OutputObject, "Sale No.", 30, AlignConstants.vbAlignLeft, False, Cy)
            'PrintTo(OutputObject, "Trans Date", 59, AlignConstants.vbAlignRight, False)
            PrintTo(OutputObject, "Trans Date", 59, AlignConstants.vbAlignRight, False, Cy)
            'PrintTo(OutputObject, "Tot Sale", 75, AlignConstants.vbAlignRight, False)
            PrintTo(OutputObject, "Tot Sale", 75, AlignConstants.vbAlignRight, False, Cy)
            'PrintTo(OutputObject, "Non Taxable", 91, AlignConstants.vbAlignRight, False)
            PrintTo(OutputObject, "Non Taxable", 91, AlignConstants.vbAlignRight, False, Cy)
            'PrintTo(OutputObject, "Taxable", 106, AlignConstants.vbAlignRight, False)
            PrintTo(OutputObject, "Taxable", 106, AlignConstants.vbAlignRight, False, Cy)
            'PrintTo(OutputObject, IIf(optDelivered.Checked = True, "Tax Rec", "Tax Chg"), 121, AlignConstants.vbAlignRight, True)
            PrintTo(OutputObject, IIf(optDelivered.Checked = True, "Tax Rec", "Tax Chg"), 121, AlignConstants.vbAlignRight, True, Cy)
        ElseIf OrderMode("STTOT") Then 'Sales tax
            '<CT>
            'Cy = 700
            'OutputObject.FontBold = True
            'OutputObject.FontSize = 9
            '</CT>
            'PrintTo(OutputObject, "Tot Sale", 75, AlignConstants.vbAlignRight, False)
            PrintTo(OutputObject, "Tot Sale", 75, AlignConstants.vbAlignRight, False, Cy)
            'PrintTo(OutputObject, "Non Taxable", 91, AlignConstants.vbAlignRight, False)
            PrintTo(OutputObject, "Non Taxable", 91, AlignConstants.vbAlignRight, False, Cy)
            'PrintTo(OutputObject, "Taxable", 106, AlignConstants.vbAlignRight, False)
            PrintTo(OutputObject, "Taxable", 106, AlignConstants.vbAlignRight, False, Cy)
            'PrintTo(OutputObject, IIf(optDelivered.Checked = True, "Tax Rec", "Tax Chg"), 121, AlignConstants.vbAlignRight, True)
            PrintTo(OutputObject, IIf(optDelivered.Checked = True, "Tax Rec", "Tax Chg"), 121, AlignConstants.vbAlignRight, True, Cy)
        End If

        PrintSet(FontBold:=0)
    End Sub

    Private Sub ProcessTaxRates(ByVal TaxCode As Integer)
        Dim SQL As String, RS As ADODB.Recordset, NewAudit As SalesJournalNew
        Dim nT As Decimal
        Dim Titled As Boolean

        SQL = ""
        SQL = SQL & "SELECT Audit.*, Holding.NonTaxable "
        SQL = SQL & "FROM Audit LEFT JOIN Holding "
        SQL = SQL & "ON Trim(Audit.SaleNo)=Holding.LeaseNo "
        SQL = SQL & "WHERE TransDate BETWEEN #" & dDate.Value & "# AND #" & toDate.Value & "# "
        If IsCanadian() And TaxCode <> 0 Then
            ' This makes Dodd's adjustments show up in PST.
            SQL = SQL & "AND TaxCode in (1," & TaxCode & ") "
        Else
            If TaxCode <> 0 Then SQL = SQL & "AND TaxCode=" & TaxCode & " "
        End If

        If optDelivered.Checked = True Then      'delivered sales
            SQL = SQL & " AND (DelSls<>0 or TaxRec1<>0)"
        ElseIf optWritten.Checked = True Then    'written sales
            SQL = SQL & " AND (Written<>0 or TaxCharged1<>0)"
        End If

        SQL = SQL & " ORDER BY SaleNo, TransDate, AuditID"
        RS = GetRecordsetBySQL(SQL)

        ProgressForm(0, RS.RecordCount, "Processing Tax Zone #" & TaxCode & " (0/" & RS.RecordCount & ")")
        On Error Resume Next           'Needed for BO w/ no date

        Cy3 = 900
        Do Until RS.EOF
            If Not Titled Then UndeliveredHeading() : Titled = True
            ProgressForm(RS.AbsolutePosition, , "Processing " & IIf(TaxCode = 0, "Def Tax Zone", "Tax Zone #" & TaxCode) & " (" & RS.AbsolutePosition & " of " & RS.RecordCount & " Sales)")
            SalesJournalNew_RecordSet_Set(NewAudit, RS)
            nT = IIf(Microsoft.VisualBasic.Left(NewAudit.Name1, 1) = "V", -1, 1) * IfNullThenZeroCurrency(RS("Holding.NonTaxable").Value)

            '    If IsDevelopment And NewAudit.SaleNo = "504159" Then Stop
            '    If IsDevelopment And NewAudit.SaleNo = "16660" Then Stop
            '    If IsDevelopment And NewAudit.SaleNo = "16656" Then Stop
            ' BFH20060519
            ' patch for audit's nontaxable was inserted in May 2006
            ' we only want this to run on dates that the full month is okay, so we limit it to
            ' after this apply date
            ' the audit's nontaxable also negates voids automatically
            ' This was mainly for Ergonomic b/c we believe they are the only ones who use the
            ' sales tax report for written sales...
            If optWritten.Checked = True And DateAfter(dDate.Value, #6/1/2006#) And Not IsDoddsLtd Then
                nT = IfNullThenZeroCurrency(RS("Audit.NonTaxable").Value)
            End If
            SalesTaxLines(TaxCode, NewAudit, nT)
            RS.MoveNext()
            Cy3 = Cy3 + 200
        Loop

        If Titled Then SalesTaxtotals(TaxCode = 0)
        ProgressForm()
    End Sub

    Private Sub SalesTaxtotals(Optional ByVal InstallmentTax As Boolean = False)
        On Error Resume Next
        Cy3 = OutputObject.CurrentY
        OutputObject.Print
        Cy3 = Cy3 + 200
        'PrintTo(OutputObject, "Totals:", 3, AlignConstants.vbAlignLeft, False)
        PrintTo(OutputObject, "Totals:", 3, AlignConstants.vbAlignLeft, False, Cy3)
        'PrintTo(OutputObject, CurrencyFormat(TotSales), 75, AlignConstants.vbAlignRight, False)
        PrintTo(OutputObject, CurrencyFormat(TotSales), 75, AlignConstants.vbAlignRight, False, Cy3)
        'PrintTo(OutputObject, CurrencyFormat(TotNonTaxable), 91, AlignConstants.vbAlignRight, False)
        PrintTo(OutputObject, CurrencyFormat(TotNonTaxable), 91, AlignConstants.vbAlignRight, False, Cy3)
        'PrintTo(OutputObject, CurrencyFormat(TotTaxable), 106, AlignConstants.vbAlignRight, False)
        PrintTo(OutputObject, CurrencyFormat(TotTaxable), 106, AlignConstants.vbAlignRight, False, Cy3)
        'PrintTo(OutputObject, CurrencyFormat(TotTaxRec), 121, AlignConstants.vbAlignRight, True)
        PrintTo(OutputObject, CurrencyFormat(TotTaxRec), 121, AlignConstants.vbAlignRight, True, Cy3)

        If InstallmentTax Then
            If StoreSettings.bInstallmentInterestIsTaxable Then
                Dim T As Decimal, U As Decimal, V As Decimal
                T = CalculateInstallmentInterest()
                U = CalculateInstallmentInterestCredit()
                V = T - U
                PrintTo(OutputObject, "", , , True)
                If T <> 0 Then
                    PrintTo(OutputObject, "Period Installment Interest Sls Tax:   ", 121, AlignConstants.vbAlignRight, False)
                    PrintTo(OutputObject, CurrencyFormat(T), 131, AlignConstants.vbAlignRight, True)
                End If
                If U <> 0 Then
                    PrintTo(OutputObject, "Period Installment Interest Sls Tax Credits:   ", 121, AlignConstants.vbAlignRight, False)
                    PrintTo(OutputObject, CurrencyFormat(U), 131, AlignConstants.vbAlignRight, True)
                End If
                If T <> 0 Or U <> 0 Or V <> 0 Then
                    PrintTo(OutputObject, "Period Installment Interest Sls Tax (Total):   ", 121, AlignConstants.vbAlignRight, False)
                    PrintTo(OutputObject, CurrencyFormat(V), 131, AlignConstants.vbAlignRight, True)
                End If
            End If
        End If
    End Sub

    Private Function CalculateInstallmentInterestCredit() As Decimal
        Dim SQL As String, RS As ADODB.Recordset
        On Error Resume Next
        SQL = ""
        SQL = SQL & "SELECT Sum(Credits) As Tot FROM [Transactions] LEFT JOIN [InstallmentInfo] "
        SQL = SQL & "ON ([Transactions].[ArNo]=[InstallmentInfo].[ArNo]) "
        SQL = SQL & "WHERE (true=true) "
        SQL = SQL & "AND ([TransDate] BETWEEN #" & dDate.Value & "# AND #" & toDate.Value & "#) "
        SQL = SQL & "AND ([Type] = 'Sls. Tax Payoff') "
        SQL = SQL & "AND ([Status] <> 'V') " ' bfh20090415 removed VOIDS

        RS = GetRecordsetBySQL(SQL)
        If Not RS Is Nothing Then
            CalculateInstallmentInterestCredit = RS("Tot").Value
        End If
    End Function

    Private Function CalculateInstallmentInterest() As Decimal
        Dim SQL As String, RS As ADODB.Recordset
        On Error Resume Next
        SQL = ""
        SQL = SQL & "SELECT Sum(Charges) As Tot FROM [Transactions] LEFT JOIN [InstallmentInfo] "
        SQL = SQL & "ON ([Transactions].[ArNo]=[InstallmentInfo].[ArNo]) "
        SQL = SQL & "WHERE (true=true) "
        SQL = SQL & "AND ([TransDate] BETWEEN #" & dDate.Value & "# AND #" & toDate.Value & "#) "
        SQL = SQL & "AND ([Type] = 'Int. Sls Tax') "
        SQL = SQL & "AND (Not [Status] IN ('V','W')) "
        '  SQL = SQL & "AND (Left([Type],7)='NewSale') "

        RS = GetRecordsetBySQL(SQL)
        If Not RS Is Nothing Then
            CalculateInstallmentInterest = RS("Tot").Value
        End If
    End Function

    Private Sub SalesTaxLines(ByVal TaxCode As Integer, ByRef NewAudit As SalesJournalNew, ByVal NonTaxable As Decimal)
        Dim GrossSale As Decimal, Taxable As Decimal, TAXREC As Decimal, Amt As Decimal, Amt2 As Decimal
        Dim TaxTypeChargedOnSale As Decimal, IsTaxed As Boolean, IsTaxedAtAll As Boolean
        Dim IsAdjustment As Boolean
        Dim effTaxRec As Decimal

        '  If IsDevelopment And Trim(NewAudit.LeaseNo) = "3977" Then Stop    ' DEBUGING

        GetSaleInfo(TaxCode, NewAudit.SaleNo, TaxTypeChargedOnSale, IsTaxed, IsTaxedAtAll)
        ' BFH20060227 - For a VOID line, we must take off all added tax..
        If IsIn(Microsoft.VisualBasic.Left(NewAudit.Name1, 2), "VD", "VB") Then TaxTypeChargedOnSale = -TaxTypeChargedOnSale

        On Error GoTo HandleErr
        If optDelivered.Checked = True Then 'Tax on Delivered
            TAXREC = GetPrice(NewAudit.TaxRec1)
            Taxable = GetPrice(NewAudit.DelSls) - NonTaxable
            If Not IsTaxed Then TAXREC = 0
            If IsCanadian() Then effTaxRec = TAXREC * CurTaxRate / TotalTaxRate Else effTaxRec = TAXREC ' MJK 20130602
            GrossSale = Taxable + NonTaxable + TAXREC
        Else 'Written
            TAXREC = GetPrice(NewAudit.TaxCharged1)
            If Not IsTaxed Then TAXREC = 0
            If IsCanadian() Then effTaxRec = TAXREC * CurTaxRate / TotalTaxRate Else effTaxRec = TAXREC ' MJK 20130602

            Amt = GetPrice(NewAudit.Written)

            If Microsoft.VisualBasic.Left(NewAudit.Name1, 5) = "Adj. " Then
                If TAXREC <> 0 Then
                    Taxable = Amt
                Else
                    Taxable = 0
                End If
            Else
                Taxable = Amt - NonTaxable
                If Amt > 0 Then
                    If Amt < NonTaxable Then Taxable = 0
                Else
                    If Math.Abs(Amt) < NonTaxable Then Taxable = 0
                End If
            End If

            GrossSale = Amt + TAXREC
        End If

        ''bfh20060227 - this fails for adjustments and voids
        '###               as for right now, two tax types on a sale will probably fail
        '' BFH20060216
        '' The amount received cant be greater than the tax charged on this sale for this tax type.
        '' it could be an issue for multiple tax charged since Audit only stores 1 field for
        '' all tax received
        '  If TAXREC > 0 Then
        '    If TAXREC > TaxTypeChargedOnSale Then TAXREC = TaxTypeChargedOnSale
        '  ElseIf TAXREC < 0 Then   ' negative tax received for adjustment->return
        '    If TAXREC < TaxTypeChargedOnSale Then TAXREC = TaxTypeChargedOnSale
        '  End If

        '  If IsDevelopment And Trim(NewAudit.SaleNo) = "12239" Then Stop

        '<CT>
        'If TaxCode > 0 And Taxable = 0 Then Exit Sub   ' only print completely non-taxable sales on the default sales tax report
        If TaxCode > 0 And Taxable = 0 Then
            Cy3 = Cy3 - 200
            Exit Sub   ' only print completely non-taxable sales on the default sales tax report
        End If

        ' if there's none of this type of tax, skip it
        ' unless it's the first report and there is nontaxable present
        'If Not IsTaxed And (TaxCode <> 0 Or (TaxCode = 0 And (NonTaxable = 0 Or IsTaxedAtAll))) Then Exit Sub
        If Not IsTaxed And (TaxCode <> 0 Or (TaxCode = 0 And (NonTaxable = 0 Or IsTaxedAtAll))) Then
            Cy3 = Cy3 - 200
            Exit Sub
        End If
        '</CT>

        ' if this type of tax wasn't on the sale, but we're here, then we're showing
        ' it because of the nontaxable amount on the sale.  We just need to make sure
        ' that taxrec is 0 so the report and totals come out right

        'TotTaxRec = TotTaxRec + TaxRec
        TotTaxRec = TotTaxRec + effTaxRec ' MJK 20130602
        TotSales = TotSales + GrossSale
        TotTaxable = TotTaxable + Taxable
        TotNonTaxable = TotNonTaxable + NonTaxable
        If Taxable = GrossSale Then NonTaxableSales = NonTaxableSales + GrossSale

        '<CT>
        OutputObject.FontBold = False
        '</CT>
        'PrintTo(OutputObject, NewAudit.Name1, 0, AlignConstants.vbAlignLeft, False)
        PrintTo(OutputObject, NewAudit.Name1, 0, AlignConstants.vbAlignLeft, False, Cy3)
        'PrintTo(OutputObject, NewAudit.SaleNo, 30, AlignConstants.vbAlignLeft, False)
        PrintTo(OutputObject, NewAudit.SaleNo, 30, AlignConstants.vbAlignLeft, False, Cy3)
        'PrintTo(OutputObject, NewAudit.TransDate, 59, AlignConstants.vbAlignRight, False)
        PrintTo(OutputObject, NewAudit.TransDate, 59, AlignConstants.vbAlignRight, False, Cy3)
        'PrintTo(OutputObject, CurrencyFormat(GrossSale), 75, AlignConstants.vbAlignRight, False)
        PrintTo(OutputObject, CurrencyFormat(GrossSale), 75, AlignConstants.vbAlignRight, False, Cy3)
        'PrintTo(OutputObject, CurrencyFormat(NonTaxable), 91, AlignConstants.vbAlignRight, False)
        PrintTo(OutputObject, CurrencyFormat(NonTaxable), 91, AlignConstants.vbAlignRight, False, Cy3)
        'PrintTo(OutputObject, CurrencyFormat(Taxable), 106, AlignConstants.vbAlignRight, False)
        PrintTo(OutputObject, CurrencyFormat(Taxable), 106, AlignConstants.vbAlignRight, False, Cy3)
        'PrintTo(OutputObject, CurrencyFormat(effTaxRec), 121, AlignConstants.vbAlignRight, True)
        PrintTo(OutputObject, CurrencyFormat(effTaxRec), 121, AlignConstants.vbAlignRight, True, Cy3)
        '  PrintTo OutputObject, CurrencyFormat(TAXREC), 121, alignconstants.vbalignright, True

        'BFH20070419 used for debugging...  dumps all tax records to a CSV file
        '  WriteFile DevOutputFolder & "taxrep.csv", (Taxable) & "," & (TAXREC)

        Counter = Counter + 1
        '  If Counter = 41 And OutputToPrinter = False Then OutputObject.Print String(120, "_")

        If Counter >= 62 Then
            If OutputToPrinter Then OutputObject.NewPage Else frmPrintPreviewDocument.NewPage()
            UndeliveredHeading()
            Counter = 0
        End If
        Exit Sub

HandleErr:
        Select Case Err().ToString
            Case 13 'Type Mismatch
                Resume Next
            Case 75
                Resume
            Case Else
                Debug.Assert(False)
                Resume Next
        End Select
    End Sub

    Private Sub GetSaleInfo(ByVal TaxCode As Integer, ByVal SaleNo As String, ByRef TaxTypeCharged As Decimal, ByRef IsTaxed As Boolean, ByRef IsTaxedAtAll As Boolean)
        Dim SQL As String, RS As ADODB.Recordset
        On Error Resume Next
        SQL = ""
        SQL = SQL & "SELECT Sum(SellPrice) AS TaxR, Count(SellPrice) AS IsTaxed FROM GrossMargin"
        SQL = SQL & " WHERE SaleNo='" & Trim(SaleNo) & "'"
        If TaxCode = 0 Then
            SQL = SQL & " AND (Style='TAX1' OR (Style='TAX2' and Quantity=1))"
        Else
            SQL = SQL & " AND Style='TAX2' AND Quantity=" & TaxCode
        End If

        RS = GetRecordsetBySQL(SQL, , GetDatabaseAtLocation())
        TaxTypeCharged = IfNullThenZeroCurrency(RS("TaxR").Value)
        IsTaxed = (IfNullThenZero(RS("IsTaxed").Value) > 0)
        DisposeDA(RS)

        SQL = ""
        SQL = SQL & "SELECT Sum(SellPrice) AS TaxR, Count(SellPrice) AS IsTaxed FROM GrossMargin"
        SQL = SQL & " WHERE SaleNo='" & Trim(SaleNo) & "'"
        SQL = SQL & " AND Style IN ('TAX1','TAX2')"

        RS = GetRecordsetBySQL(SQL, , GetDatabaseAtLocation())
        IsTaxedAtAll = (IfNullThenZero(RS("IsTaxed").Value) > 0)
        DisposeDA(RS)
    End Sub

    Private Sub ClearVariables()
        TotSales = 0
        TotNonTaxable = 0
        TotTaxable = 0
        Taxable = 0
        Counter = 0
        TotTaxRec = 0
        NonTaxableSales = 0
        cboStoreSelect.SelectedIndex = -1
    End Sub

    Private Function GetMaxTaxRate() As Integer
        Dim SQL As String, RS As ADODB.Recordset
        On Error Resume Next
        GetMaxTaxRate = SalesTax2Count()
        SQL = "SELECT Max(TaxCode) As MaxTaxCode From Audit WHERE TransDate BETWEEN #" & dDate.Value & "# AND #" & toDate.Value & "#"
        RS = GetRecordsetBySQL(SQL)
        If RS.EOF Then
            RS = Nothing
            Exit Function
        End If
        If RS("MaxTaxCode").Value > GetMaxTaxRate Then GetMaxTaxRate = RS("MaxTaxCode").Value
        RS.Close()
        RS = Nothing
    End Function

    Private Sub Ctrls(ByVal Enabled As Boolean)
        'MousePointer = IIf(Enabled, vbDefault, vbHourglass)
        Me.Cursor = IIf(Enabled, Cursors.Default, Cursors.WaitCursor)
        On Error Resume Next
        cmdPrint.Enabled = Enabled
        cmdPrintPreview.Enabled = Enabled
        cmdCancel.Enabled = Enabled
        EnableFrame(Me, fraSaleType, Enabled)
        EnableFrame(Me, fraDates, Enabled)
        EnableFrame(Me, fraStoreSelect, Enabled)
        Application.DoEvents()
    End Sub

    Private Sub cmdPrintPreview_Click(sender As Object, e As EventArgs) Handles cmdPrintPreview.Click
        '<CT>
        Printer.PrintAction = Printing.PrintAction.PrintToPreview
        '</CT>

        If OrderMode("ATR") Then
            If cboStoreSelect.SelectedIndex < 0 Or cboStoreSelect.SelectedIndex > ssMaxStore - 1 Then
                MessageBox.Show("Please select a store.", "WinCDS", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                Exit Sub
            End If
        End If

        Ctrls(False)

        'Load frmPrintPreviewMain
        'OutputToPrinter = False
        OutputToPrinter = True
        OutputObject = Printer
        'OutputObject = frmPrintPreviewDocument.picPicture

        'frmPrintPreviewDocument.CallingForm = Me
        'frmPrintPreviewDocument.ReportName = Text
        If OrderMode("ST") Then
            SalesTax()
        ElseIf OrderMode("ATR") Then
            'AdvertisingReport(dDate.Value, toDate.Value, 0, cboStoreSelect.SelectedIndex + 1, chkGroupByZip.Checked = True, chkSortByZip.Checked = True)
            AdvertisingReport(dDate.Value, toDate.Value, 1, cboStoreSelect.SelectedIndex + 1, chkGroupByZip.Checked = True, chkSortByZip.Checked = True)
        Else
            'Unload frmPrintPreviewMain
            frmPrintPreviewMain.Close()
            cmdPrint.Enabled = True
            cmdPrintPreview.Enabled = True
            cmdCancel.Enabled = True
            'MousePointer = vbDefault
            Me.Cursor = Cursors.Default
            Exit Sub
        End If

        Ctrls(True)
        'Hide()

        '<CT>
        'frmPrintPreviewDocument.DataEnd()
        Printer.EndDoc()
        '</CT>
    End Sub

    Private Sub DateForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim HDiff As Integer, XI As Integer

        SetButtonImage(cmdCancel, 3)
        SetButtonImage(cmdPrint, 19)
        SetButtonImage(cmdPrintPreview, 20)

        '<CT>
        ButtonGroupbox.Location = New Point(8, 144)
        Me.Size = New Size(294, 257)
        '</CT>

        fraSaleType.Visible = OrderMode("ST")
        lblNote.Visible = False

        If OrderMode("T", "ST") Or ArMode("N", "O") Then
            'Both Dates
            cmdCancel.Visible = True
            lblDate1.Text = "&From:"
            lblDate2.Text = "&To:"
        ElseIf OrderMode("ATR") Then
            lblDate1.Text = "&From:"
            lblDate2.Text = "&To:"
            fraStoreSelect.Visible = True
            '<CT>
            fraStoreSelect.Location = New Point(8, 79)
            ButtonGroupbox.Location = New Point(17, 182)
            Me.Size = New Size(292, 315)
            '</CT>
            fraSaleType.Visible = False
            'HDiff = cmdPrint.Top
            'cmdPrint.Top = fraStoreSelect.Top + fraStoreSelect.Height + 60
            'cmdPrintPreview.Top = cmdPrint.Top
            'cmdCancel.Top = cmdPrint.Top
            'HDiff = cmdPrint.Top - HDiff
            'Height = Height + HDiff
            cmdCancel.Visible = True
            Text = "Advertising Types Report"
            'DateForm.HelpContextID = 49700

            LoadStoresIntoComboBox(cboStoreSelect, StoresSld, False, True)
            '    With cboStoreSelect
            '      .Clear
            '      For XI = 1 To LicensedNoOfStores ' bfh20051208
            '        .AddItem "Loc " & XI & ": " & frmSetup .QueryStoreLocAdd(XI)
            '      Next
            '      If .ListCount >= StoresSld Then .ListIndex = StoresSld - 1
            '    End With
        Else
            cmdCancel.Visible = False 'One date
            toDate.Enabled = False
            cmdPrintPreview.Visible = False

            If Not OrderMode("C") Then 'void
                Width = FRM_W2
                fraDates.Width = 1815
            End If
            'cmdPrint.Move(ScaleWidth - cmdPrint.Width) / 2, 1380
            cmdPrint.Location = New Point(((Me.ClientSize.Width - cmdPrint.Width) / 2), 1380)
            cmdPrint.Text = "&Apply"
            Height = FRM_H1 ' 2190
        End If

        If ArMode("T") Then                   ' Trial Balance
            dDate.Value = DateFormat(CurrentMonthStart)
            toDate.Value = DateFormat(Now) 'Month (Date)
        ElseIf OrderMode("ST") Then                 ' Sales Tax Report
            dDate.Value = MonthlyReportDefaultStart() '  LastFullMonthStart
            toDate.Value = MonthlyReportDefaultEnd()
        ElseIf OrderMode("ATR") Then
            dDate.Value = MonthlyReportDefaultStart()
            toDate.Value = MonthlyReportDefaultEnd()
        Else                                              ' Otherwise...
            dDate.Value = DateFormat(Now)
            toDate.Value = DateFormat(Now)
        End If

        If OrderMode("ST") Then
            Text = "Sales Tax Report"
            lblDate1.Text = "&From:"
        ElseIf OrderMode("C") Then 'void added 11-21-01
            lblNote.Visible = True
            toDate.Visible = False
            toDate.Value = Today '""
            Text = "Void on This Date"
            lblDate1.Text = "&Date:"
        End If
    End Sub
End Class
