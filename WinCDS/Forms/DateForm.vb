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
        Dim effTaxCode As Long
        For TaxCode = 0 To GetMaxTaxRate()
            effTaxCode = TaxCode
            If TaxCode = 0 Then
                PrintTo(OutputObject, "Default Tax", 1, AlignConstants.vbAlignLeft)
                PrintTo(OutputObject, StoreSettings.SalesTax, 20, AlignConstants.vbAlignLeft)
                If IsCanadian() Then effTaxCode = 1
            Else
                PrintTo(OutputObject, "Tax Zone #" & TaxCode, 1, AlignConstants.vbAlignLeft)
                PrintTo(OutputObject, QuerySalesTax2(TaxCode - 1), 20, AlignConstants.vbAlignLeft)
                If IsCanadian() And TaxCode = 1 Then effTaxCode = 0
            End If

            PrintTo(OutputObject, CurrencyFormat(TaxTotals.Item(Format(effTaxCode, "00") & " TotSales")), 75, AlignConstants.vbAlignRight)
            TotSales = TotSales + GetPrice(TaxTotals.Item(Format(effTaxCode, "00") & " TotSales"))

            PrintTo(OutputObject, CurrencyFormat(TaxTotals.Item(Format(effTaxCode, "00") & " NonTaxtable")), 91, AlignConstants.vbAlignRight)
            TotNonTaxable = TotNonTaxable + GetPrice(TaxTotals.Item(Format(effTaxCode, "00") & " NonTaxtable"))

            PrintTo(OutputObject, CurrencyFormat(TaxTotals.Item(Format(effTaxCode, "00") & " Taxable")), 106, AlignConstants.vbAlignRight)
            TotTaxable = TotTaxable + GetPrice(TaxTotals.Item(Format(effTaxCode, "00") & " Taxable"))

            PrintTo(OutputObject, CurrencyFormat(TaxTotals.Item(Format(effTaxCode, "00") & " TotTaxRec")), 121, AlignConstants.vbAlignRight, True)
            TotTaxRec = TotTaxRec + GetPrice(TaxTotals.Item(Format(effTaxCode, "00") & " TotTaxRec"))
        Next

        If Not IsDoddsLtd Then ' BFH20130423 they do multiple TAX2's on a sale.  Totals are not necessary.
            OutputObject.Print
            OutputObject.FontBold = True
            PrintTo(OutputObject, "Totals:", 1, AlignConstants.vbAlignLeft)

            PrintTo(OutputObject, CurrencyFormat(TotSales), 75, AlignConstants.vbAlignRight)
            PrintTo(OutputObject, CurrencyFormat(TotNonTaxable), 91, AlignConstants.vbAlignRight)
            PrintTo(OutputObject, CurrencyFormat(TotTaxable), 106, AlignConstants.vbAlignRight)
            PrintTo(OutputObject, CurrencyFormat(TotTaxRec), 121, AlignConstants.vbAlignRight, True)
            PrintTo(OutputObject, "Sale totals may be over-counted if sales have multiple TAX2 rates.", 10, , True)

            OutputObject.FontBold = False
        End If
    End Sub

    Public Sub UndeliveredHeading()
        OutputObject.FontName = "Arial"
        OutputObject.FontSize = 18
        OutputObject.CurrentY = 100
        OutputObject.FontBold = True

        If OrderMode("ST") Then
            If optDelivered.Checked = True Then
                PrintCentered("Sales Tax Report:  Delivered Sales")
            ElseIf optWritten.Checked = True Then
                PrintCentered("Sales Tax Report:  Written Sales")
            End If
        ElseIf OrderMode("NST") Then
            PrintCentered("Non Taxable Sales Report")
        End If

        PrintSet(, 10, 100, "Arial", 8, 0)
        OutputObject.Print("From: ", DateFormat(dDate), "  To: ", DateFormat(toDate))
        OutputObject.Print("Time: ", Format(Now, "h:mm:ss tt"))

        PrintSet(, 10100, 100)
        If OutputToPrinter Then PageNumber = OutputObject.Page

        OutputObject.Print("Page: " & PageNumber)
        If OrderMode("ST") Then
            PrintToPosition(OutputObject, ShortenSalesTaxDescription(SaleTaxRate), 10100 + Printer.TextWidth("Page: " & PageNumber), AlignConstants.vbAlignRight, True)
        End If

        OutputObject.CurrentY = 500
        PrintCentered(StoreSettings.Name & "    " & StoreSettings.Address & "    " & StoreSettings.City)

        PrintSet(, 0, 700, , 9, 1)

        If OrderMode("ST") Then 'Sales tax
            PrintTo(OutputObject, "Last Name", 0, AlignConstants.vbAlignLeft, False)
            PrintTo(OutputObject, "Sale No.", 30, AlignConstants.vbAlignLeft, False)
            PrintTo(OutputObject, "Trans Date", 59, AlignConstants.vbAlignRight, False)
            PrintTo(OutputObject, "Tot Sale", 75, AlignConstants.vbAlignRight, False)
            PrintTo(OutputObject, "Non Taxable", 91, AlignConstants.vbAlignRight, False)
            PrintTo(OutputObject, "Taxable", 106, AlignConstants.vbAlignRight, False)
            PrintTo(OutputObject, IIf(optDelivered.Checked = True, "Tax Rec", "Tax Chg"), 121, AlignConstants.vbAlignRight, True)
        ElseIf OrderMode("STTOT") Then 'Sales tax
            PrintTo(OutputObject, "Tot Sale", 75, AlignConstants.vbAlignRight, False)
            PrintTo(OutputObject, "Non Taxable", 91, AlignConstants.vbAlignRight, False)
            PrintTo(OutputObject, "Taxable", 106, AlignConstants.vbAlignRight, False)
            PrintTo(OutputObject, IIf(optDelivered.Checked = True, "Tax Rec", "Tax Chg"), 121, AlignConstants.vbAlignRight, True)
        End If

        PrintSet(FontBold:=0)
    End Sub

    Private Sub ProcessTaxRates(ByVal TaxCode As Long)
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
        Loop

        If Titled Then SalesTaxtotals(TaxCode = 0)
        ProgressForm()
    End Sub

    Private Sub SalesTaxtotals(Optional ByVal InstallmentTax As Boolean)
        On Error Resume Next
        OutputObject.Print
        PrintTo OutputObject, "Totals:", 3, vbAlignLeft, False
  PrintTo OutputObject, CurrencyFormat(TotSales), 75, vbAlignRight, False
  PrintTo OutputObject, CurrencyFormat(TotNonTaxable), 91, vbAlignRight, False
  PrintTo OutputObject, CurrencyFormat(TotTaxable), 106, vbAlignRight, False
  PrintTo OutputObject, CurrencyFormat(TotTaxRec), 121, vbAlignRight, True

  If InstallmentTax Then
            If StoreSettings.bInstallmentInterestIsTaxable Then
                Dim T As Currency, U As Currency, V As Currency
                T = CalculateInstallmentInterest
                U = CalculateInstallmentInterestCredit
                V = T - U
                PrintTo OutputObject, "", , , True
      If T <> 0 Then
                    PrintTo OutputObject, "Period Installment Interest Sls Tax:   ", 121, vbAlignRight, False
        PrintTo OutputObject, CurrencyFormat(T), 131, vbAlignRight, True
      End If
                If U <> 0 Then
                    PrintTo OutputObject, "Period Installment Interest Sls Tax Credits:   ", 121, vbAlignRight, False
        PrintTo OutputObject, CurrencyFormat(U), 131, vbAlignRight, True
      End If
                If T <> 0 Or U <> 0 Or V <> 0 Then
                    PrintTo OutputObject, "Period Installment Interest Sls Tax (Total):   ", 121, vbAlignRight, False
        PrintTo OutputObject, CurrencyFormat(V), 131, vbAlignRight, True
      End If
            End If
        End If
    End Sub

    Private Sub SalesTaxLines(ByVal TaxCode As Long, ByRef NewAudit As SalesJournalNew, ByVal NonTaxable As Currency)
        Dim GrossSale As Currency, Taxable As Currency, TAXREC As Currency, Amt As Currency, Amt2 As Currency
        Dim TaxTypeChargedOnSale As Currency, IsTaxed As Boolean, IsTaxedAtAll As Boolean
        Dim IsAdjustment As Boolean
        Dim effTaxRec As Currency

        '  If IsDevelopment And Trim(NewAudit.LeaseNo) = "3977" Then Stop    ' DEBUGING

        GetSaleInfo TaxCode, NewAudit.SaleNo, TaxTypeChargedOnSale, IsTaxed, IsTaxedAtAll
  ' BFH20060227 - For a VOID line, we must take off all added tax..
        If IsIn(Left(NewAudit.Name1, 2), "VD", "VB") Then TaxTypeChargedOnSale = -TaxTypeChargedOnSale

        On Error GoTo HandleErr
        If optDelivered Then 'Tax on Delivered
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

            If Left(NewAudit.Name1, 5) = "Adj. " Then
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
                    If Abs(Amt) < NonTaxable Then Taxable = 0
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
        If TaxCode > 0 And Taxable = 0 Then Exit Sub   ' only print completely non-taxable sales on the default sales tax report
        ' if there's none of this type of tax, skip it
        ' unless it's the first report and there is nontaxable present
        If Not IsTaxed And (TaxCode <> 0 Or (TaxCode = 0 And (NonTaxable = 0 Or IsTaxedAtAll))) Then Exit Sub
        ' if this type of tax wasn't on the sale, but we're here, then we're showing
        ' it because of the nontaxable amount on the sale.  We just need to make sure
        ' that taxrec is 0 so the report and totals come out right

        'TotTaxRec = TotTaxRec + TaxRec
        TotTaxRec = TotTaxRec + effTaxRec ' MJK 20130602
        TotSales = TotSales + GrossSale
        TotTaxable = TotTaxable + Taxable
        TotNonTaxable = TotNonTaxable + NonTaxable
        If Taxable = GrossSale Then NonTaxableSales = NonTaxableSales + GrossSale

        PrintTo OutputObject, NewAudit.Name1, 0, vbAlignLeft, False
  PrintTo OutputObject, NewAudit.SaleNo, 30, vbAlignLeft, False
  PrintTo OutputObject, NewAudit.TransDate, 59, vbAlignRight, False
  PrintTo OutputObject, CurrencyFormat(GrossSale), 75, vbAlignRight, False
  PrintTo OutputObject, CurrencyFormat(NonTaxable), 91, vbAlignRight, False
  PrintTo OutputObject, CurrencyFormat(Taxable), 106, vbAlignRight, False
  PrintTo OutputObject, CurrencyFormat(effTaxRec), 121, vbAlignRight, True
'  PrintTo OutputObject, CurrencyFormat(TAXREC), 121, vbAlignRight, True


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
        Select Case Err()
            Case 13 'Type Mismatch
                Resume Next
            Case 75
                Resume
            Case Else
                Debug.Assert False
      Resume Next
        End Select
    End Sub

    Private Sub ClearVariables()
        TotSales = 0
        TotNonTaxable = 0
        TotTaxable = 0
        Taxable = 0
        Counter = 0
        TotTaxRec = 0
        NonTaxableSales = 0
        cboStoreSelect.ListIndex = -1
    End Sub

    Private Function GetMaxTaxRate() As Long
        Dim SQL As String, RS As ADODB.Recordset
        On Error Resume Next
        GetMaxTaxRate = SalesTax2Count()
        SQL = "SELECT Max(TaxCode) As MaxTaxCode From Audit WHERE TransDate BETWEEN #" & dDate & "# AND #" & toDate & "#"
  Set RS = GetRecordsetBySQL(SQL)
  If RS.EOF Then
    Set RS = Nothing
    Exit Function
        End If
        If RS("MaxTaxCode") > GetMaxTaxRate Then GetMaxTaxRate = RS("MaxTaxCode")
        RS.Close()
          Set RS = Nothing
End Function

    Private Sub Ctrls(ByVal Enabled As Boolean)
        MousePointer = IIf(Enabled, vbDefault, vbHourglass)
        On Error Resume Next
        cmdPrint.Enabled = Enabled
        cmdPrintPreview.Enabled = Enabled
        cmdCancel.Enabled = Enabled
        EnableFrame Me, fraSaleType, Enabled
  EnableFrame Me, fraDates, Enabled
  EnableFrame Me, fraStoreSelect, Enabled
  DoEvents
    End Sub

End Class