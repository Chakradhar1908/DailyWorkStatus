Imports VBRUN
Public Class OrdAudit2
    Private CashJournalRecordSet As ADODB.Recordset
    Dim Written As Decimal
    Dim TaxCHARGED As Decimal
    Dim ARCASHSALES As Decimal
    Dim CUSTDEP As Decimal
    Dim UNDSALES As Decimal
    Dim DELSALES As Decimal
    Dim TAXREC As Decimal
    Dim PWRITTEN As Decimal
    Dim PTAXCHARGED As Decimal
    Dim PARCASHSALES As Decimal
    Dim PCUSTDEP As Decimal
    Dim PUNDSALES As Decimal
    Dim PTAXREC As Decimal
    Dim PDELSALES As Decimal
    Dim Index As Integer
    Dim Header As Integer
    Dim Start As Integer
    Dim SCASH As Decimal, PTAX As Decimal, STAX As Decimal, Tax As Decimal, Sales As Decimal, PSALES As Decimal, BEGCASH As Decimal, Pcash As Decimal, Cash As Decimal, ECHECKS As Decimal, PECHECKS As Decimal
    Dim RSALES As Decimal, PRSALES As Decimal, REFUNDTAX As Decimal, PREFUNDTAX As Decimal, BCASHSALES As Decimal, PBCASHSALES As Decimal

    Dim Counter As Integer
    Dim NETRECEIPTS As Decimal, CASHSALES As Decimal, Check As Decimal, VISA As Decimal, MASTER As Decimal, Disc As Decimal, AMCX As Decimal, FINC As Decimal, None As Decimal, STORECARD As Decimal
    Dim PNETRECEIPTS As Decimal, PCASHSALES As Decimal, PCheck As Decimal, PVISA As Decimal, PMASTER As Decimal, PDISC As Decimal, PAMCX As Decimal, PFINC As Decimal, PNONE As Decimal, PSTORECARD As Decimal

    Dim BVISA As Decimal, BDisc As Decimal, BAmcx As Decimal, BSTORECARD As Decimal
    Dim PBVISA As Decimal, PBDisc As Decimal, PBAmcx As Decimal, PBSTORECARD As Decimal, VISACHECK As Decimal, PTOTCASHIN As Decimal, TOTCASHIN As Decimal, BECHECKS As Decimal, PBECHECKS As Decimal
    Dim PCashIn As Decimal, CashIn As Decimal, PCashOut As Decimal, CashOut As Decimal, CARRYCHARGE As Decimal, PCARRYCHARGE As Decimal
    Dim POTHER As Decimal, PRTAX As Decimal, OTHER As Decimal, PBANK As Decimal, CASHONHAND As Decimal, SalesTax As String
    Dim PBEGCASH As Decimal, BMASTER As Decimal, BNETRECEIPTS As Decimal, COHAND As Decimal, PCOHAND As Decimal

    Dim PBMASTER As Decimal, PEROIDBANK As Decimal, MTDEBITS As Decimal, PMTDEBITS As Decimal, MTDCREDITS As Decimal, PMTDCREDITS As Decimal

    Dim PAR As Decimal, AR As Decimal, PINTEREST As Decimal, INTEREST As Decimal, Bank As Decimal, RECPAYMENTS As Decimal, PRECPAYMENTS As Decimal

    Private Sub txtPriorPeriodCash_TextChanged(sender As Object, e As EventArgs) Handles txtPriorPeriodCash.TextChanged

    End Sub

    Dim PCHECKREFUND As Decimal, CheckRefund As Decimal, PVISACHECK As Decimal, PBCPAY As Decimal, BCPAY As Decimal, PMISCCASHIN As Decimal, MISCCASHIN As Decimal
    Dim PMISCCASHOUT As Decimal, MiscCashOut As Decimal, PPURCHASES As Decimal, PURCHASES As Decimal
    Dim PRESALE As Decimal, RESALE As Decimal, PFINANCE As Decimal, FINANCE As Decimal
    Dim PVISADISC As Decimal, VISADISC As Decimal, PPETTYCASH As Decimal, PETTYCASH As Decimal, Total As Decimal, PTotal As Decimal
    Dim PFREIGHTOUT As Decimal, FREIGHTOUT As Decimal, PGAS As Decimal, GAS As Decimal, PCREDIT As Decimal, Credit As Decimal
    Dim PMAINTENANCE As Decimal, MAINTENANCE As Decimal, PREPAIR As Decimal, REPAIR As Decimal, PWHSE As Decimal, WHSE As Decimal
    Dim POFFICE As Decimal, OFFICE As Decimal, PCASUAL As Decimal, CASUAL As Decimal, PTRAVEL As Decimal, TRAVEL As Decimal

    Dim PFORFEIT As Decimal, Forfeit As Decimal

    Dim XCashIn As Decimal, XPCashIn As Decimal

    Dim Typee As String
    Dim TaxableSales As Decimal
    Dim TaxExemptSales As Decimal
    Dim TotBackOrders As Decimal
    Dim Backs As Decimal

    Dim NewAudit As SalesJournalNew
    Dim NewCash As CashJournalNew
    Dim theDate As String
    Dim ToTheDate As String
    Dim TotCash As Decimal
    Dim TotPCash As Decimal
    Dim Debit As Decimal, PDebit As Decimal

    Private Const EntireStore As String = "[Entire Store]"

    Private ReadOnly Property DateFilter(Optional ByVal Previous As Boolean = False) As String
        Get
            '"TransDate >= #" & Format(theDate, "mm/01/yyyy") & "# AND TransDate < #" & theDate & "# ORDER BY [TransDate], AuditID"
            If Previous Then
                DateFilter = "AND ([TransDate] >= #" & MonthStart(theDate) & "# AND [TransDate] < #" & theDate & "#) "
            Else
                DateFilter = "AND " & SQLDateRange("TransDate", theDate, ToTheDate) & " "
            End If
        End Get
    End Property

    Private ReadOnly Property TerminalFilter(Optional ByVal Fld As String = "TransDate") As String
        Get
            Dim C As String
            C = EntireStore
            If C = EntireStore Then
                TerminalFilter = ""
            Else
                TerminalFilter = "AND ([Terminal]=""" & ProtectSQL(C) & """) "
            End If
        End Get
    End Property

    Private ReadOnly Property CashierFilter(Optional ByVal Fld As String = "TransDate") As String
        Get
            Dim C As String
            C = cmbCashier.Text
            If C = EntireStore Then
                CashierFilter = ""
            Else
                CashierFilter = "AND ([Cashier]=""" & ProtectSQL(C) & """) "
            End If
            CashierFilter = CashierFilter & TerminalFilter
        End Get
    End Property

    Private Sub cmdCancel_Click(sender As Object, e As EventArgs) Handles cmdCancel.Click
        If OutputToPrinter Then OutputObject.KillDoc
        CashJournalRecordSet = Nothing
        'Unload frmPrintPreviewMain
        frmPrintPreviewMain.Close()
        'Unload Me
        Me.Dispose()
        Me.Close()
        MainMenu.Show()
        modProgramState.Reports = ""
    End Sub

    Private Sub cmdPrint_Click(sender As Object, e As EventArgs) Handles cmdPrint.Click
        Dim DBG As String
        On Error GoTo ErrorHandler
        DBG = "a"
        If Printer Is Nothing Then  ' bfh20050617 - hopefully this will make it not crash when no printer is installed
            MessageBox.Show("No default printer is set.  Use preview to view this report.", "Default Printer Not Found", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Exit Sub
        End If

        DBG = "b"
        If Not StoreSettings.bManualBillofSaleNo Then
            If Not frmEditSalesJournal.OutOfDateSalesReport(StoresSld, dteStartDate.Value, dteEndDate.Value) Then Exit Sub
            If Not frmEditCash.OutOfDateCashReport(StoresSld, dteStartDate.Value, dteEndDate.Value) Then Exit Sub
        End If

        DBG = "c"
        'MousePointer = vbHourglass
        Me.Cursor = Cursors.WaitCursor
        cmdPrint.Enabled = False
        cmdPrintPreview.Enabled = False
        cmdCancel.Enabled = False

        OutputObject = New cPrinter
        OutputObject.SetPrintToPDF("Daily Audit Report", "Daily Audit,Audit,Daily Audit Report")

        DBG = "d"
        PrintSub()

        DBG = "e"
        'MousePointer = vbDefault
        Me.Cursor = Cursors.Default
        cmdPrint.Enabled = True
        cmdPrintPreview.Enabled = True
        cmdCancel.Enabled = True 'Enable user to cancel generation of report.
        DBG = "f"
        TrackUsage("DailyAuditReport")
        Exit Sub

ErrorHandler:
        Select Case Err.Number
            Case 482 : ErrNoPrinter()
                Exit Sub
            Case Else
                MessageBox.Show("Error in previewing audit report (" & Err.Number & "):" & vbCrLf & Err.Description & vbCrLf & "Source: " & Err.Source & vbCrLf & "DBG=" & DBG)
        End Select
        Resume Next
    End Sub

    Private Sub PrintSub()
        TotCash = 0
        TotPCash = 0
        Sales = 0
        PSALES = 0
        ECHECKS = 0
        PECHECKS = 0
        BCASHSALES = 0
        PBCASHSALES = 0
        CheckRefund = 0
        PCHECKREFUND = 0
        BVISA = 0
        PBVISA = 0
        BECHECKS = 0
        PBECHECKS = 0
        BDisc = 0
        PBDisc = 0
        BAmcx = 0
        PBAmcx = 0
        BSTORECARD = 0
        PBSTORECARD = 0
        PNETRECEIPTS = 0
        CASHSALES = 0
        PCASHSALES = 0
        PCheck = 0
        PVISA = 0
        PMASTER = 0
        PDISC = 0
        PAMCX = 0
        PFINC = 0
        PNONE = 0
        PEROIDBANK = 0
        PBCASHSALES = 0
        PBMASTER = 0
        PBDisc = 0
        PBAmcx = 0
        PMTDEBITS = 0
        PNETRECEIPTS = 0
        PTOTCASHIN = 0
        PBEGCASH = 0
        PMTDCREDITS = 0
        PEROIDBANK = 0
        PCashOut = 0
        NETRECEIPTS = 0
        Check = 0
        VISA = 0
        MASTER = 0
        Disc = 0
        AMCX = 0
        FINC = 0
        None = 0
        STORECARD = 0
        PSTORECARD = 0
        BNETRECEIPTS = 0
        BCASHSALES = 0
        BVISA = 0
        BMASTER = 0
        BDisc = 0
        BAmcx = 0
        MTDEBITS = 0
        BEGCASH = 0
        NETRECEIPTS = 0
        TOTCASHIN = 0
        MTDCREDITS = 0
        Bank = 0
        PBANK = 0
        CashOut = 0
        PVISADISC = 0
        VISADISC = 0
        PPETTYCASH = 0
        PETTYCASH = 0
        Total = 0
        PTotal = 0
        PFREIGHTOUT = 0
        FREIGHTOUT = 0
        PGAS = 0
        GAS = 0
        PCREDIT = 0
        Credit = 0
        PMAINTENANCE = 0
        MAINTENANCE = 0
        PREPAIR = 0
        REPAIR = 0
        PWHSE = 0
        WHSE = 0
        POFFICE = 0
        OFFICE = 0
        PCASUAL = 0
        CASUAL = 0
        PTRAVEL = 0
        TRAVEL = 0
        RECPAYMENTS = 0
        PRECPAYMENTS = 0
        RSALES = 0
        PRSALES = 0
        PINTEREST = 0
        INTEREST = 0
        MiscCashOut = 0
        PMISCCASHOUT = 0
        MISCCASHIN = 0
        PMISCCASHIN = 0
        VISACHECK = 0
        PVISACHECK = 0
        BCPAY = 0
        PBCPAY = 0
        RESALE = 0
        PRESALE = 0
        FINANCE = 0
        PFINANCE = 0
        PURCHASES = 0
        PPURCHASES = 0
        FINANCE = 0
        PFINANCE = 0
        MISCCASHIN = 0
        PMISCCASHIN = 0
        RESALE = 0
        PRESALE = 0
        Debit = 0
        PDebit = 0
        AR = 0
        PAR = 0
        Forfeit = 0
        PFORFEIT = 0
        PURCHASES = 0
        PPURCHASES = 0
        PCashIn = 0
        CashIn = 0

        XCashIn = 0
        XPCashIn = 0

        On Error Resume Next
        AuditReport()

        CalculatePreviousTotals()

        LoadCash(DateFormat(theDate), DateFormat(ToTheDate))
        CashReport()

        '    Refund  ' Part of CashReport now.
        CashInReport()

        CashOutDwr()
        Banking()
        CheckForfeit()
        RecControl()

        If Installment Then
            ArPayments()
            ArLate()
        End If

        CashSummary()
        GenLedgerTrans()

        CashJournalRecordSet = Nothing

        OutputObject.EndDoc
    End Sub

    Private Sub Headings()
        ' Left Side
        OutputObject.FontName = "Arial"
        OutputObject.CurrentX = 0
        OutputObject.CurrentY = 100
        OutputObject.FontSize = 8
        'OutputObject.PrintNL("Time: ", Format(Now, "h:mm:ss am/pm")) 'More Conventional
        OutputObject.PrintNL("Time: ", Format(Now, "h:mm:ss tt")) 'More Conventional
        OutputObject.CurrentY = 280
        OutputObject.PrintNL("Beginning Date: ", DateFormat(theDate))


        ' Right Header
        OutputObject.CurrentX = 9500
        OutputObject.CurrentY = 100
        OutputObject.PrintNL("Report Date: ", DateFormat(Now))

        OutputObject.CurrentX = 8700
        OutputObject.CurrentY = 280
        If OutputToPrinter Then 'Does _not_ work with IIf()
            OutputObject.PrintNL(" Page: ", OutputObject.Page, "  Ending Date: ", DateFormat(ToTheDate))
        Else
            OutputObject.PrintNL(" Page: ", PageNumber, "  Ending Date: ", DateFormat(ToTheDate))
        End If

        OutputObject.FontSize = 8
        Counter = 0

        OutputObject.CurrentY = 500
        PrintCentered(StoreSettings.Name & "    " & StoreSettings.Address & "    " & StoreSettings.City)
        OutputObject.CurrentY = 800

        'Main Headings
        OutputObject.FontSize = 18
        OutputObject.FontBold = True
        OutputObject.CurrentY = 100

        Select Case Header
            Case 1 : PrintCentered("Sales Journal")
            Case 2 : PrintCentered("Customer Payments - Refunds")
            Case 3 : PrintCentered("Misc. Cash In")
            Case 4 : PrintCentered("Receivables - Back Orders")
            Case 5 : PrintCentered("Installment Payments")
            Case 6 : PrintCentered("Installment Late Charges")
            Case 7 : PrintCentered("Cash Management")
            Case 8 : PrintCentered("General Ledger Summary")
            Case 9 : PrintCentered("Out of Date Cash Transactions")
            Case 10 : PrintCentered("Out of Date Audit Records")
            Case 11 : PrintCentered("Check Refund/Forfeit")
            Case 31 : PrintCentered("Misc. Cash Out")
            Case 32 : PrintCentered("Bank Deposits")
            Case 33 : PrintCentered("Check Refund/Forfeit (21500/41500)")
        End Select

        OutputObject.FontBold = False
        OutputObject.FontSize = 8
        OutputObject.CurrentY = 800
    End Sub

    Private Sub GenLedgerTrans()
        Dim Mcol As Integer, Pcol As Integer
        Mcol = 93
        Pcol = 123

        'GenLedgerHeader
        Header = 8
        Headings()
        OutputObject.FontSize = 12
        OutputObject.FontBold = True

        OutputObject.PrintNL
        'PrintTo(OutputObject, "ACC", 1, AlignConstants.vbAlignLeft, False)
        PrintTo(OutputObject, "ACC", 1, AlignConstants.vbAlignLeft, False, 1200)
        'PrintTo(OutputObject, "DESC", 23, AlignConstants.vbAlignLeft, False)
        PrintTo(OutputObject, "DESC", 23, AlignConstants.vbAlignLeft, False, 1200)
        'PrintTo(OutputObject, "MONTH TO DATE", Mcol, AlignConstants.vbAlignRight, False)
        PrintTo(OutputObject, "MONTH TO DATE", Mcol, AlignConstants.vbAlignRight, False, 1200)
        'PrintTo(OutputObject, "PERIOD TO DATE", Pcol, AlignConstants.vbAlignRight, True)
        PrintTo(OutputObject, "PERIOD TO DATE", Pcol, AlignConstants.vbAlignRight, True, 1200)

        OutputObject.FontBold = False

        'PrintGLTrans("11000", "AR CASH SALES", ARCASHSALES, PARCASHSALES)
        PrintGLTrans("11000", "AR CASH SALES", ARCASHSALES, PARCASHSALES,, 1500)

        'PrintGLTrans("21400", "CUSTOMER DEP (Change)", CUSTDEP, PCUSTDEP)
        PrintGLTrans("21400", "CUSTOMER DEP (Change)", CUSTDEP, PCUSTDEP,, 1800)

        'PrintGLTrans(IIf(UNDSALES >= 0, "40600", "11100"), "UND. SALES       " & IIf(UNDSALES >= 0, "-", "+"), UNDSALES, PUNDSALES)
        PrintGLTrans(IIf(UNDSALES >= 0, "40600", "11100"), "UND. SALES       " & IIf(UNDSALES >= 0, "-", "+"), UNDSALES, PUNDSALES,, 2100)
        'PrintGLTrans("40200", "DEL SALES", DELSALES, PDELSALES)
        PrintGLTrans("40200", "DEL SALES", DELSALES, PDELSALES,, 2400)
        'PrintGLTrans("41700", "SALES TAX REC.", TAXREC, PTAXREC)
        PrintGLTrans("41700", "SALES TAX REC.", TAXREC, PTAXREC,, 2700)
        'PrintGLTrans("10100", "CUST PAYMENTS", TotCash - RSALES, TotPCash - PRSALES)
        PrintGLTrans("10100", "CUST PAYMENTS", TotCash - RSALES, TotPCash - PRSALES,, 3000)
        'PrintGLTrans("10100", "CUST REFUNDS", RSALES, PRSALES)
        PrintGLTrans("10100", "CUST REFUNDS", RSALES, PRSALES,, 3300)

        '    PrintGLTrans "SUB", "Subtotal", _
        'ARCASHSALES+CUSTDEP - Math.Abs(UNDSALES) - DELSALES - TAXREC + TotCash - RSALES + RSALES,
        'PARCASHSALES + PCUSTDEP - Math.Abs(PUNDSALES) - PDELSALES - PTAXREC + TotPCash - PRSALES + PRSALES, False)

        OutputObject.PrintNL
        'PrintGLTrans("11300", "A/R PRINCIPAL PAY", AR, PAR)
        PrintGLTrans("11300", "A/R PRINCIPAL PAY", AR, PAR,, 4000)
        'PrintGLTrans("40500", "A/R LATE CHARGES", INTEREST, PINTEREST)
        PrintGLTrans("40500", "A/R LATE CHARGES", INTEREST, PINTEREST,, 4300)
        'PrintGLTrans("10100", "A/R PAYMENTS", -(AR + INTEREST), -(PAR + PINTEREST))
        PrintGLTrans("10100", "A/R PAYMENTS", -(AR + INTEREST), -(PAR + PINTEREST),, 4600)

        'Cash in
        OutputObject.PrintNL
        'PrintGLTrans("69600", "OTHER INCOME", MISCCASHIN, PMISCCASHIN)
        PrintGLTrans("69600", "OTHER INCOME", MISCCASHIN, PMISCCASHIN,, 5300)
        ' PrintGLTrans "10700", "MASTER/VISA CHK", VISACHECK, PVISACHECK
        'PrintGLTrans("61600", "MEDICAL CO-PAY", BCPAY, PBCPAY)
        PrintGLTrans("61600", "MEDICAL CO-PAY", BCPAY, PBCPAY,, 5600)
        ' PrintGLTrans "50200", "PURCHASES COD", RESALE, PRESALE
        ' PrintGLTrans "70000/71000", "FINANCE EXCHANGE", FINANCE, PFINANCE
        'PrintGLTrans("41500", "FORFEIT DEPOSITS", Forfeit, PFORFEIT)   ' FOFFEIT in v8
        PrintGLTrans("41500", "FORFEIT DEPOSITS", Forfeit, PFORFEIT,, 5900)   ' FOFFEIT in v8
        'PrintGLTrans("10100", "MISC RECEIPTS TOTAL", -(TOTCASHIN), -(PTOTCASHIN))
        PrintGLTrans("10100", "MISC RECEIPTS TOTAL", -(TOTCASHIN), -(PTOTCASHIN),, 6200)
        '    PrintGLTrans "10100", "MISC RECEIPTS TOTAL", -(TOTCASHIN + FORFEIT), -(PTOTCASHIN + PFORFEIT)  ' original..??

        'Cash out
        OutputObject.PrintNL
        'PrintGLTrans("64200", "MISC EXPENSE", MiscCashOut, PMISCCASHOUT)
        PrintGLTrans("64200", "MISC EXPENSE", MiscCashOut, PMISCCASHOUT,, 6900)
        'PrintGLTrans("50200", "PURCHASES COD", PURCHASES, PPURCHASES)
        PrintGLTrans("50200", "PURCHASES COD", PURCHASES, PPURCHASES,, 7200)
        'PrintGLTrans("50500", "FREIGHT OUT", FREIGHTOUT, PFREIGHTOUT)
        PrintGLTrans("50500", "FREIGHT OUT", FREIGHTOUT, PFREIGHTOUT,, 7500)
        'PrintGLTrans("50600", "DISCOUNT/FINAN", Credit, PCREDIT)
        PrintGLTrans("50600", "DISCOUNT/FINAN", Credit, PCREDIT,, 7800)
        'PrintGLTrans("60100", "GAS & OIL", GAS, PGAS)
        PrintGLTrans("60100", "GAS & OIL", GAS, PGAS,, 8100)
        'PrintGLTrans("60500", "DISC VISA Etc.", VISADISC, PVISADISC)
        PrintGLTrans("60500", "DISC VISA Etc.", VISADISC, PVISADISC,, 8400)
        'PrintGLTrans("62300", "MAINTENANCE", MAINTENANCE, PMAINTENANCE)
        PrintGLTrans("62300", "MAINTENANCE", MAINTENANCE, PMAINTENANCE,, 8700)
        'PrintGLTrans("62400", "REPAIR & REFIN", REPAIR, PREPAIR)
        PrintGLTrans("62400", "REPAIR & REFIN", REPAIR, PREPAIR,, 9000)
        'PrintGLTrans("63500", "WHSE SUPPLY", WHSE, PWHSE)
        PrintGLTrans("63500", "WHSE SUPPLY", WHSE, PWHSE,, 9300)
        'PrintGLTrans("64100", "OFFICE SUPPLY", OFFICE, POFFICE)
        PrintGLTrans("64100", "OFFICE SUPPLY", OFFICE, POFFICE,, 9600)
        'PrintGLTrans("65200", "CASUAL LABOR", CASUAL, PCASUAL)
        PrintGLTrans("65200", "CASUAL LABOR", CASUAL, PCASUAL,, 9900)
        'PrintGLTrans("67500", "MEALS & ENTERTAIN", TRAVEL, PTRAVEL)
        PrintGLTrans("67500", "MEALS & ENTERTAIN", TRAVEL, PTRAVEL,, 10200)
        'PrintGLTrans("52000", "CASH OVER/SHORT", COHAND - GetPrice(txtCashInDrawer.Text), PCOHAND - GetPrice(txtCashInDrawer.Text))
        PrintGLTrans("52000", "CASH OVER/SHORT", COHAND - GetPrice(txtCashInDrawer.Text), PCOHAND - GetPrice(txtCashInDrawer.Text),, 10500)
        'PrintGLTrans("10100", "CASH OUT TOTAL", -CashOut, -PCashOut)
        PrintGLTrans("10100", "CASH OUT TOTAL", -CashOut, -PCashOut,, 10800)


        OutputObject.PrintNL
        'PrintGLTrans("10200", "CASH & CHECKS", BCASHSALES, PBCASHSALES)
        PrintGLTrans("10200", "CASH & CHECKS", BCASHSALES, PBCASHSALES,, 11500)
        'PrintGLTrans("10200", "ELECTRONIC CHECKS", BECHECKS, PBECHECKS)
        PrintGLTrans("10200", "ELECTRONIC CHECKS", BECHECKS, PBECHECKS,, 11800)
        'PrintGLTrans("10200", "VISA-MASTER", BVISA, PBVISA)
        PrintGLTrans("10200", "VISA-MASTER", BVISA, PBVISA,, 12100)
        'PrintGLTrans("10200", "DISCOVER", BDisc, PBDisc)
        PrintGLTrans("10200", "DISCOVER", BDisc, PBDisc,, 12400)
        'PrintGLTrans("10200", "AMEX", BAmcx, PBAmcx)
        PrintGLTrans("10200", "AMEX", BAmcx, PBAmcx,, 12700)
        'PrintGLTrans("10200", "DEBIT CARD", Convert.ToDecimal(Debit), Convert.ToDecimal(PDebit))
        PrintGLTrans("10200", "DEBIT CARD", Convert.ToDecimal(Debit), Convert.ToDecimal(PDebit),, 13000)
        'PrintGLTrans("10200", "STORE CREDIT CARD", BSTORECARD, PBSTORECARD)
        PrintGLTrans("10200", "STORE CREDIT CARD", BSTORECARD, PBSTORECARD,, 13300)
        'PrintGLTrans("10100", "BANK DEPOSITS TOTAL", -Bank, -PBANK)
        PrintGLTrans("10100", "BANK DEPOSITS TOTAL", -Bank, -PBANK,, 13600)

        OutputObject.PrintNL
        'PrintGLTrans("21500", "CHECK REFUND", CheckRefund, PCHECKREFUND)
        PrintGLTrans("21500", "CHECK REFUND", CheckRefund, PCHECKREFUND,, 14100)
        'PrintGLTrans("11200", "BACK ORDERS", RECPAYMENTS, PRECPAYMENTS)
        PrintGLTrans("11200", "BACK ORDERS", RECPAYMENTS, PRECPAYMENTS,, 14400)

        'PrintTo(OutputObject, "______________", Mcol, AlignConstants.vbAlignRight, False)
        PrintTo(OutputObject, "______________", Mcol, AlignConstants.vbAlignRight, False, 14700)
        'PrintTo(OutputObject, "______________", Pcol, AlignConstants.vbAlignRight, True)
        PrintTo(OutputObject, "______________", Pcol, AlignConstants.vbAlignRight, True, 14700)

        Total = RECPAYMENTS + ARCASHSALES + CashOut + Sales + CUSTDEP + MISCCASHIN + VISACHECK + BCPAY + Bank + CheckRefund + AR + INTEREST + Forfeit '+ RSALES '+ TotCash
        Total = Total + RESALE + FINANCE
        Total = Total - DELSALES - TAXREC - BCASHSALES - BECHECKS - BVISA - BMASTER - BDisc - BAmcx - Debit - BSTORECARD - TOTCASHIN - MiscCashOut - AR - INTEREST - CARRYCHARGE
        Total = Total - FREIGHTOUT - GAS - Credit - MAINTENANCE - REPAIR - WHSE - OFFICE - CASUAL - TRAVEL - PURCHASES - VISADISC - (COHAND - GetPrice(txtCashInDrawer.Text))

        PTotal = PRECPAYMENTS + PARCASHSALES + PCashOut + Pcash + PCUSTDEP + PMISCCASHIN + PVISACHECK + PBCPAY + PBANK + PCHECKREFUND + TotPCash + PAR + PINTEREST + PFORFEIT '+ PRSALES
        PTotal = PTotal + PRESALE + PFINANCE
        PTotal = PTotal - PDELSALES - PTAXREC - PBCASHSALES - PBECHECKS - PBVISA - PBMASTER - PBDisc - PBAmcx - PDebit - PBSTORECARD - PTOTCASHIN - PMISCCASHOUT - PAR - PINTEREST - PCARRYCHARGE
        PTotal = PTotal - PFREIGHTOUT - PGAS - PCREDIT - PMAINTENANCE - PREPAIR - PWHSE - POFFICE - PCASUAL - PTRAVEL - PPURCHASES - PVISADISC - (PCOHAND - GetPrice(txtCashInDrawer.Text))


        If UNDSALES > 1 Then
            Total = Total - Math.Abs(UNDSALES)
        Else
            Total = Total - UNDSALES
        End If

        If PUNDSALES > 1 Then
            PTotal = PTotal - Math.Abs(PUNDSALES)
        Else
            PTotal = PTotal - PUNDSALES
        End If

        'PrintTo(OutputObject, Format(Total, "$###,##0.00"), Mcol, AlignConstants.vbAlignRight, False)
        PrintTo(OutputObject, Format(Total, "$###,##0.00"), Mcol, AlignConstants.vbAlignRight, False, 15000)
        'PrintTo(OutputObject, Format(PTotal, "$###,##0.00"), Pcol, AlignConstants.vbAlignRight, True)
        PrintTo(OutputObject, Format(PTotal, "$###,##0.00"), Pcol, AlignConstants.vbAlignRight, True, 15000)
    End Sub

    Private Sub PrintGLTrans(ByVal Account As String, ByVal Desc As String, ByVal MTD As Decimal, ByVal PTD As Decimal, Optional ByVal Negative As Boolean = False, Optional ByVal CY As Integer = 0)
        Dim Mcol As Integer, Pcol As Integer
        Mcol = 93
        Pcol = 123
        'PrintTo(OutputObject, Account, 0, AlignConstants.vbAlignLeft, False)
        PrintTo(OutputObject, Account, 0, AlignConstants.vbAlignLeft, False, CY)
        'PrintTo(OutputObject, Desc, 23, AlignConstants.vbAlignLeft, False)
        PrintTo(OutputObject, Desc, 23, AlignConstants.vbAlignLeft, False, CY)

        Negative = False

        If Negative Or MTD < 0 Then
            'PrintTo(OutputObject, Format(Math.Abs(MTD), "($###,##0.00)"), Mcol + 1, AlignConstants.vbAlignRight, False)
            PrintTo(OutputObject, Format(Math.Abs(MTD), "($###,##0.00)"), Mcol + 1, AlignConstants.vbAlignRight, False, CY)
        Else
            'PrintTo(OutputObject, Format(MTD, "$###,##0.00"), Mcol, AlignConstants.vbAlignRight, False)
            PrintTo(OutputObject, Format(MTD, "$###,##0.00"), Mcol, AlignConstants.vbAlignRight, False, CY)
        End If

        If Negative Or PTD < 0 Then
            'PrintTo(OutputObject, Format(Math.Abs(PTD), "($###,##0.00)"), Pcol + 1, AlignConstants.vbAlignRight, True)
            PrintTo(OutputObject, Format(Math.Abs(PTD), "($###,##0.00)"), Pcol + 1, AlignConstants.vbAlignRight, True, CY)
        Else
            'PrintTo(OutputObject, Format(PTD, "$###,##0.00"), Pcol, AlignConstants.vbAlignRight, True)
            PrintTo(OutputObject, Format(PTD, "$###,##0.00"), Pcol, AlignConstants.vbAlignRight, True, CY)
        End If


        '  If Not Negative Then
        '    If MTD < 0 Then
        '      PrintTo OutputObject, Format(-MTD, "($###,##0.00)"), Mcol + 1, vbAlignRight, False
        '    Else
        '      PrintTo OutputObject, Format(MTD, "$###,##0.00"), Mcol, vbAlignRight, False
        '    End If
        '    If PTD < 0 Then
        '      PrintTo OutputObject, Format(-PTD, "($###,##0.00)"), Pcol + 1, vbAlignRight, True
        '    Else
        '      PrintTo OutputObject, Format(PTD, "$###,##0.00"), Pcol, vbAlignRight, True
        '    End If
        ''    PrintTo OutputObject, format(MTD, "$###,##0.00"), Mcol, vbAlignRight, False
        ''    PrintTo OutputObject, format(PTD, "$###,##0.00"), Pcol, vbAlignRight, True
        '  Else
        '    PrintTo OutputObject, Format(MTD, "($###,##0.00)"), Mcol + 1, vbAlignRight, False
        '    PrintTo OutputObject, Format(PTD, "($###,##0.00)"), Pcol + 1, vbAlignRight, True
        '  End If
    End Sub

    Private Sub CashSummary()
        Cash = 0
        Pcash = 0
        Counter = 0
        If OutputToPrinter Then
            If OutputObject.CurrentY <> 0 Then OutputObject.NewPage
        Else
            frmPrintPreviewDocument.NewPage()
        End If

        Header = 7
        Headings()
        OutputObject.FontSize = 10

        Dim CS1 As Integer, CS2 As Integer, CS3 As Integer, CS4 As Integer
        CS1 = 60
        CS2 = 80
        CS3 = 100
        CS4 = 120
        'PrintTo(OutputObject, "M/T/D DEBITS", CS1, AlignConstants.vbAlignRight, False)
        PrintTo(OutputObject, "M/T/D DEBITS", CS1, AlignConstants.vbAlignRight, False, 900)
        'PrintTo(OutputObject, "CREDITS", CS2, AlignConstants.vbAlignRight, False)
        PrintTo(OutputObject, "CREDITS", CS2, AlignConstants.vbAlignRight, False, 900)
        'PrintTo(OutputObject, "PERIOD DEBITS", CS3, AlignConstants.vbAlignRight, False)
        PrintTo(OutputObject, "PERIOD DEBITS", CS3, AlignConstants.vbAlignRight, False, 900)
        'PrintTo(OutputObject, "CREDITS", CS4, AlignConstants.vbAlignRight, True)
        PrintTo(OutputObject, "CREDITS", CS4, AlignConstants.vbAlignRight, True, 900)

        'If Start = 1 Then  ' commented out mjk 20030908
        If DateAndTime.Day(theDate) = 1 Then
            PBEGCASH = BEGCASH   ' ***  Is this right, or should it be added to the calculated value?
        Else
            PBEGCASH = GetPrice(txtPriorPeriodCash.Text)
        End If

        'PrintTo(OutputObject, vbCrLf & "Beginning Cash", 0, AlignConstants.vbAlignLeft, False)
        PrintTo(OutputObject, "Beginning Cash", 0, AlignConstants.vbAlignLeft, False, 1200)
        'PrintTo(OutputObject, Format(BEGCASH, "$###,##0.00"), CS1, AlignConstants.vbAlignRight, False)  ' was txtCashInDrawer
        PrintTo(OutputObject, Format(BEGCASH, "$###,##0.00"), CS1, AlignConstants.vbAlignRight, False, 1260)  ' was txtCashInDrawer
        'PrintTo(OutputObject, Format(PBEGCASH, "$###,##0.00"), CS3, AlignConstants.vbAlignRight, True)
        PrintTo(OutputObject, Format(PBEGCASH, "$###,##0.00"), CS3, AlignConstants.vbAlignRight, True, 1260)

        'PrintTo(OutputObject, vbCrLf & "Cash", 0, AlignConstants.vbAlignLeft, False)
        PrintTo(OutputObject, "Cash", 0, AlignConstants.vbAlignLeft, False, 1600)
        'PrintTo(OutputObject, Format(CASHSALES, "$###,##0.00"), CS1, AlignConstants.vbAlignRight, False)
        PrintTo(OutputObject, Format(CASHSALES, "$###,##0.00"), CS1, AlignConstants.vbAlignRight, False, 1660)
        'PrintTo(OutputObject, Format(PCASHSALES, "$###,##0.00"), CS3, AlignConstants.vbAlignRight, True)
        PrintTo(OutputObject, Format(PCASHSALES, "$###,##0.00"), CS3, AlignConstants.vbAlignRight, True, 1660)

        'PrintTo(OutputObject, "Check", 0, AlignConstants.vbAlignLeft, False)
        PrintTo(OutputObject, "Check", 0, AlignConstants.vbAlignLeft, False, 1850)
        'PrintTo(OutputObject, Format(Check, "$###,##0.00"), CS1, AlignConstants.vbAlignRight, False)
        PrintTo(OutputObject, Format(Check, "$###,##0.00"), CS1, AlignConstants.vbAlignRight, False, 1910)
        'PrintTo(OutputObject, Format(PCheck, "$###,##0.00"), CS3, AlignConstants.vbAlignRight, True)
        PrintTo(OutputObject, Format(PCheck, "$###,##0.00"), CS3, AlignConstants.vbAlignRight, True, 1910)

        'PrintTo(OutputObject, "Visa", 0, AlignConstants.vbAlignLeft, False)
        PrintTo(OutputObject, "Visa", 0, AlignConstants.vbAlignLeft, False, 2070)
        'PrintTo(OutputObject, Format(VISA, "$###,##0.00"), CS1, AlignConstants.vbAlignRight, False)
        PrintTo(OutputObject, Format(VISA, "$###,##0.00"), CS1, AlignConstants.vbAlignRight, False, 2130)
        'PrintTo(OutputObject, Format(PVISA, "$###,##0.00"), CS3, AlignConstants.vbAlignRight, True)
        PrintTo(OutputObject, Format(PVISA, "$###,##0.00"), CS3, AlignConstants.vbAlignRight, True, 2130)

        'PrintTo(OutputObject, "Master", 0, AlignConstants.vbAlignLeft, False)
        PrintTo(OutputObject, "Master", 0, AlignConstants.vbAlignLeft, False, 2300)
        'PrintTo(OutputObject, Format(MASTER, "$###,##0.00"), CS1, AlignConstants.vbAlignRight, False)
        PrintTo(OutputObject, Format(MASTER, "$###,##0.00"), CS1, AlignConstants.vbAlignRight, False, 2360)
        'PrintTo(OutputObject, Format(PMASTER, "$###,##0.00"), CS3, AlignConstants.vbAlignRight, True)
        PrintTo(OutputObject, Format(PMASTER, "$###,##0.00"), CS3, AlignConstants.vbAlignRight, True, 2360)

        'PrintTo(OutputObject, "Discover", 0, AlignConstants.vbAlignLeft, False)
        PrintTo(OutputObject, "Discover", 0, AlignConstants.vbAlignLeft, False, 2540)
        'PrintTo(OutputObject, Format(Disc, "$###,##0.00"), CS1, AlignConstants.vbAlignRight, False)
        PrintTo(OutputObject, Format(Disc, "$###,##0.00"), CS1, AlignConstants.vbAlignRight, False, 2600)
        'PrintTo(OutputObject, Format(PDISC, "$###,##0.00"), CS3, AlignConstants.vbAlignRight, True)
        PrintTo(OutputObject, Format(PDISC, "$###,##0.00"), CS3, AlignConstants.vbAlignRight, True, 2600)

        'PrintTo(OutputObject, "Amex", 0, AlignConstants.vbAlignLeft, False)
        PrintTo(OutputObject, "Amex", 0, AlignConstants.vbAlignLeft, False, 2770)
        'PrintTo(OutputObject, Format(AMCX, "$###,##0.00"), CS1, AlignConstants.vbAlignRight, False)
        PrintTo(OutputObject, Format(AMCX, "$###,##0.00"), CS1, AlignConstants.vbAlignRight, False, 2830)
        'PrintTo(OutputObject, Format(PAMCX, "$###,##0.00"), CS3, AlignConstants.vbAlignRight, True)
        PrintTo(OutputObject, Format(PAMCX, "$###,##0.00"), CS3, AlignConstants.vbAlignRight, True, 2830)

        'PrintTo(OutputObject, "Finance Co.", 0, AlignConstants.vbAlignLeft, False)
        PrintTo(OutputObject, "Finance Co.", 0, AlignConstants.vbAlignLeft, False, 2980)
        'PrintTo(OutputObject, Format(FINC, "$###,##0.00"), CS1, AlignConstants.vbAlignRight, False)
        PrintTo(OutputObject, Format(FINC, "$###,##0.00"), CS1, AlignConstants.vbAlignRight, False, 3070)
        'PrintTo(OutputObject, Format(PFINC, "$###,##0.00"), CS3, AlignConstants.vbAlignRight, True)
        PrintTo(OutputObject, Format(PFINC, "$###,##0.00"), CS3, AlignConstants.vbAlignRight, True, 3070)

        'PrintTo(OutputObject, "Debit Card", 0, AlignConstants.vbAlignLeft, False)
        PrintTo(OutputObject, "Debit Card", 0, AlignConstants.vbAlignLeft, False, 3240)
        'PrintTo(OutputObject, Format(None, "$###,##0.00"), CS1, AlignConstants.vbAlignRight, False)
        PrintTo(OutputObject, Format(None, "$###,##0.00"), CS1, AlignConstants.vbAlignRight, False, 3290)
        'PrintTo(OutputObject, Format(PNONE, "$###,##0.00"), CS3, AlignConstants.vbAlignRight, True)
        PrintTo(OutputObject, Format(PNONE, "$###,##0.00"), CS3, AlignConstants.vbAlignRight, True, 3290)

        'PrintTo(OutputObject, "Store Credit Card", 0, AlignConstants.vbAlignLeft, False)
        PrintTo(OutputObject, "Store Credit Card", 0, AlignConstants.vbAlignLeft, False, 3460)
        'PrintTo(OutputObject, Format(STORECARD, "$###,##0.00"), CS1, AlignConstants.vbAlignRight, False)
        PrintTo(OutputObject, Format(STORECARD, "$###,##0.00"), CS1, AlignConstants.vbAlignRight, False, 3520)
        'PrintTo(OutputObject, Format(PSTORECARD, "$###,##0.00"), CS3, AlignConstants.vbAlignRight, True)
        PrintTo(OutputObject, Format(PSTORECARD, "$###,##0.00"), CS3, AlignConstants.vbAlignRight, True, 3520)

        ' BFH20110609 - removed, cuz echeck shoudln't be ion cash management
        ' jk put back 6-16-2011  Wrong logic
        'PrintTo(OutputObject, "Electronic Checks", 0, AlignConstants.vbAlignLeft, False)
        PrintTo(OutputObject, "Electronic Checks", 0, AlignConstants.vbAlignLeft, False, 3660)
        'PrintTo(OutputObject, Format(ECHECKS, "$###,##0.00"), CS1, AlignConstants.vbAlignRight, False)
        PrintTo(OutputObject, Format(ECHECKS, "$###,##0.00"), CS1, AlignConstants.vbAlignRight, False, 3720)
        'PrintTo(OutputObject, Format(PECHECKS, "$###,##0.00"), CS3, AlignConstants.vbAlignRight, True)
        PrintTo(OutputObject, Format(PECHECKS, "$###,##0.00"), CS3, AlignConstants.vbAlignRight, True, 3720)

        'PrintTo(OutputObject, "Forfeit Deposit", 0, AlignConstants.vbAlignLeft, False)
        PrintTo(OutputObject, "Forfeit Deposit", 0, AlignConstants.vbAlignLeft, False, 3860)
        'PrintTo(OutputObject, Format(Forfeit, "$###,##0.00"), CS1, AlignConstants.vbAlignRight, False)
        PrintTo(OutputObject, Format(Forfeit, "$###,##0.00"), CS1, AlignConstants.vbAlignRight, False, 3920)
        'PrintTo(OutputObject, Format(PFORFEIT, "$###,##0.00"), CS3, AlignConstants.vbAlignRight, True)
        PrintTo(OutputObject, Format(PFORFEIT, "$###,##0.00"), CS3, AlignConstants.vbAlignRight, True, 3920)

        '    OutputObject.Print Tab(29); "______________"; Tab(69); "______________" & vbCrLf
        'PrintTo(OutputObject, "______________", CS1, AlignConstants.vbAlignRight, False)
        PrintTo(OutputObject, "______________", CS1, AlignConstants.vbAlignRight, False, 4340)
        'PrintTo(OutputObject, "______________", CS3, AlignConstants.vbAlignRight, True)
        PrintTo(OutputObject, "______________", CS3, AlignConstants.vbAlignRight, True, 4340)

        'OutputObject.Print
        ' BFH20110609 - Echecks removed
        PNETRECEIPTS = PCASHSALES + PFORFEIT + PCheck + PVISA + PMASTER + PDISC + PAMCX + PFINC + PNONE + PSTORECARD + PECHECKS
        NETRECEIPTS = CASHSALES + Forfeit + Check + VISA + MASTER + Disc + AMCX + FINC + None + STORECARD + ECHECKS

        'PrintTo(OutputObject, "Net Receipts", 0, AlignConstants.vbAlignLeft, False)
        PrintTo(OutputObject, "Net Receipts", 0, AlignConstants.vbAlignLeft, False, 4550)
        'PrintTo(OutputObject, Format(NETRECEIPTS, "$###,##0.00"), CS1, AlignConstants.vbAlignRight, False)
        PrintTo(OutputObject, Format(NETRECEIPTS, "$###,##0.00"), CS1, AlignConstants.vbAlignRight, False, 4550)
        'PrintTo(OutputObject, Format(PNETRECEIPTS, "$###,##0.00"), CS3, AlignConstants.vbAlignRight, True)
        PrintTo(OutputObject, Format(PNETRECEIPTS, "$###,##0.00"), CS3, AlignConstants.vbAlignRight, True, 4550)

        OutputObject.PrintNL
        'PrintTo(OutputObject, "Misc. Receipts", 8, AlignConstants.vbAlignLeft, False)
        PrintTo(OutputObject, "Misc. Receipts", 8, AlignConstants.vbAlignLeft, False, 4900)
        'PrintTo(OutputObject, Format(CashIn, "$###,##0.00"), CS1, AlignConstants.vbAlignRight, False)
        PrintTo(OutputObject, Format(CashIn, "$###,##0.00"), CS1, AlignConstants.vbAlignRight, False, 4900)
        'PrintTo(OutputObject, Format(PCashIn, "$###,##0.00"), CS3, AlignConstants.vbAlignRight, True)
        PrintTo(OutputObject, Format(PCashIn, "$###,##0.00"), CS3, AlignConstants.vbAlignRight, True, 4900)

        OutputObject.FontBold = True
        OutputObject.PrintNL(vbCrLf & " Bank Deposits:")
        OutputObject.FontBold = False

        'PrintTo(OutputObject, "Cash & Checks", 0, AlignConstants.vbAlignLeft, False)
        PrintTo(OutputObject, "Cash & Checks", 0, AlignConstants.vbAlignLeft, False, 5800)
        'PrintTo(OutputObject, Format(BCASHSALES, "$###,##0.00"), CS2, AlignConstants.vbAlignRight, False)
        PrintTo(OutputObject, Format(BCASHSALES, "$###,##0.00"), CS2, AlignConstants.vbAlignRight, False, 5800)
        'PrintTo(OutputObject, Format(PBCASHSALES, "$###,##0.00"), CS4, AlignConstants.vbAlignRight, True)
        PrintTo(OutputObject, Format(PBCASHSALES, "$###,##0.00"), CS4, AlignConstants.vbAlignRight, True, 5800)

        'PrintTo(OutputObject, "Visa Master", 0, AlignConstants.vbAlignLeft, False)
        PrintTo(OutputObject, "Visa Master", 0, AlignConstants.vbAlignLeft, False, 6000)
        'PrintTo(OutputObject, Format(BVISA, "$###,##0.00"), CS2, AlignConstants.vbAlignRight, False)
        PrintTo(OutputObject, Format(BVISA, "$###,##0.00"), CS2, AlignConstants.vbAlignRight, False, 6000)
        'PrintTo(OutputObject, Format(PBVISA, "$###,##0.00"), CS4, AlignConstants.vbAlignRight, True)
        PrintTo(OutputObject, Format(PBVISA, "$###,##0.00"), CS4, AlignConstants.vbAlignRight, True, 6000)

        'PrintTo(OutputObject, "Discover", 0, AlignConstants.vbAlignLeft, False)
        PrintTo(OutputObject, "Discover", 0, AlignConstants.vbAlignLeft, False, 6200)
        'PrintTo(OutputObject, Format(BDisc, "$###,##0.00"), CS2, AlignConstants.vbAlignRight, False)
        PrintTo(OutputObject, Format(BDisc, "$###,##0.00"), CS2, AlignConstants.vbAlignRight, False, 6200)
        'PrintTo(OutputObject, Format(PBDisc, "$###,##0.00"), CS4, AlignConstants.vbAlignRight, True)
        PrintTo(OutputObject, Format(PBDisc, "$###,##0.00"), CS4, AlignConstants.vbAlignRight, True, 6200)

        'PrintTo(OutputObject, "Amex", 0, AlignConstants.vbAlignLeft, False)
        PrintTo(OutputObject, "Amex", 0, AlignConstants.vbAlignLeft, False, 6400)
        'PrintTo(OutputObject, Format(BAmcx, "$###,##0.00"), CS2, AlignConstants.vbAlignRight, False)
        PrintTo(OutputObject, Format(BAmcx, "$###,##0.00"), CS2, AlignConstants.vbAlignRight, False, 6400)
        'PrintTo(OutputObject, Format(PBAmcx, "$###,##0.00"), CS4, AlignConstants.vbAlignRight, True)
        PrintTo(OutputObject, Format(PBAmcx, "$###,##0.00"), CS4, AlignConstants.vbAlignRight, True, 6400)

        'PrintTo(OutputObject, "Debit Deposit", 0, AlignConstants.vbAlignLeft, False)
        PrintTo(OutputObject, "Debit Deposit", 0, AlignConstants.vbAlignLeft, False, 6600)
        'PrintTo(OutputObject, Format(Debit, "$###,##0.00"), CS2, AlignConstants.vbAlignRight, False)
        PrintTo(OutputObject, Format(Debit, "$###,##0.00"), CS2, AlignConstants.vbAlignRight, False, 6600)
        'PrintTo(OutputObject, Format(PDebit, "$###,##0.00"), CS4, AlignConstants.vbAlignRight, True)
        PrintTo(OutputObject, Format(PDebit, "$###,##0.00"), CS4, AlignConstants.vbAlignRight, True, 6600)

        'PrintTo(OutputObject, "Store Credit Card", 0, AlignConstants.vbAlignLeft, False)
        PrintTo(OutputObject, "Store Credit Card", 0, AlignConstants.vbAlignLeft, False, 6800)
        'PrintTo(OutputObject, Format(BSTORECARD, "$###,##0.00"), CS2, AlignConstants.vbAlignRight, False)
        PrintTo(OutputObject, Format(BSTORECARD, "$###,##0.00"), CS2, AlignConstants.vbAlignRight, False, 6800)
        'PrintTo(OutputObject, Format(PBSTORECARD, "$###,##0.00"), CS4, AlignConstants.vbAlignRight, True)
        PrintTo(OutputObject, Format(PBSTORECARD, "$###,##0.00"), CS4, AlignConstants.vbAlignRight, True, 6800)

        'PrintTo(OutputObject, "Electronic Checks", 0, AlignConstants.vbAlignLeft, False)
        PrintTo(OutputObject, "Electronic Checks", 0, AlignConstants.vbAlignLeft, False, 7000)
        'PrintTo(OutputObject, Format(BECHECKS, "$###,##0.00"), CS2, AlignConstants.vbAlignRight, False)
        PrintTo(OutputObject, Format(BECHECKS, "$###,##0.00"), CS2, AlignConstants.vbAlignRight, False, 7000)
        'PrintTo(OutputObject, Format(PBECHECKS, "$###,##0.00"), CS4, AlignConstants.vbAlignRight, True)
        PrintTo(OutputObject, Format(PBECHECKS, "$###,##0.00"), CS4, AlignConstants.vbAlignRight, True, 7000)

        '    OutputObject.Print Tab(49); "______________"; Tab(88); "______________" & vbCrLf
        'PrintTo(OutputObject, "______________", CS2, AlignConstants.vbAlignRight, False)
        PrintTo(OutputObject, "______________", CS2, AlignConstants.vbAlignRight, False, 7400)
        'PrintTo(OutputObject, "______________", CS4, AlignConstants.vbAlignRight, True)
        PrintTo(OutputObject, "______________", CS4, AlignConstants.vbAlignRight, True, 7400)

        PEROIDBANK = PBCASHSALES + PBVISA + PBMASTER + PBDisc + PBAmcx + PDebit + PBSTORECARD + PBECHECKS
        BNETRECEIPTS = BCASHSALES + BVISA + BMASTER + BDisc + BAmcx + Debit + BSTORECARD + BECHECKS

        'PrintTo(OutputObject, "Total Bank Deposits", 3, AlignConstants.vbAlignLeft, False)
        PrintTo(OutputObject, "Total Bank Deposits", 3, AlignConstants.vbAlignLeft, False, 7600)
        'PrintTo(OutputObject, Format(Bank, "$###,##0.00"), CS2, AlignConstants.vbAlignRight, False)
        PrintTo(OutputObject, Format(Bank, "$###,##0.00"), CS2, AlignConstants.vbAlignRight, False, 7600)
        'PrintTo(OutputObject, Format(PEROIDBANK, "$###,##0.00"), CS4, AlignConstants.vbAlignRight, True)
        PrintTo(OutputObject, Format(PEROIDBANK, "$###,##0.00"), CS4, AlignConstants.vbAlignRight, True, 7600)

        'PrintTo(OutputObject, vbCrLf & "Less Petty Cash Out", 3, AlignConstants.vbAlignLeft, False)
        PrintTo(OutputObject, vbCrLf & "Less Petty Cash Out", 3, AlignConstants.vbAlignLeft, False, 7800)
        'PrintTo(OutputObject, Format(CashOut, "$###,##0.00"), CS2, AlignConstants.vbAlignRight, False)
        PrintTo(OutputObject, Format(CashOut, "$###,##0.00"), CS2, AlignConstants.vbAlignRight, False, 8000)
        'PrintTo(OutputObject, Format(PCashOut, "$###,##0.00"), CS4, AlignConstants.vbAlignRight, True)
        PrintTo(OutputObject, Format(PCashOut, "$###,##0.00"), CS4, AlignConstants.vbAlignRight, True, 8000)

        '    MTDEBITS = GetPrice(txtCashInDrawer) + NETRECEIPTS + TOTCASHIN
        MTDEBITS = NETRECEIPTS + CashIn + BEGCASH
        MTDCREDITS = Bank + CashOut
        PMTDEBITS = PNETRECEIPTS + PCashIn + PBEGCASH
        PMTDCREDITS = PEROIDBANK + PCashOut

        '    OutputObject.Print Tab(37); "_______________________________________________________"
        'PrintTo(OutputObject, New String("_", 65), CS4, AlignConstants.vbAlignRight, True)
        PrintTo(OutputObject, New String("_", 65), CS4, AlignConstants.vbAlignRight, True, 8400)

        'PrintTo(OutputObject, Format(MTDEBITS, "$###,##0.00"), CS1, AlignConstants.vbAlignRight, False)
        PrintTo(OutputObject, Format(MTDEBITS, "$###,##0.00"), CS1, AlignConstants.vbAlignRight, False, 8600)
        'PrintTo(OutputObject, Format(MTDCREDITS, "$###,##0.00"), CS2, AlignConstants.vbAlignRight, False)
        PrintTo(OutputObject, Format(MTDCREDITS, "$###,##0.00"), CS2, AlignConstants.vbAlignRight, False, 8600)

        'PrintTo(OutputObject, Format(PMTDEBITS, "$###,##0.00"), CS3, AlignConstants.vbAlignRight, False)
        PrintTo(OutputObject, Format(PMTDEBITS, "$###,##0.00"), CS3, AlignConstants.vbAlignRight, False, 8600)
        'PrintTo(OutputObject, Format(PMTDCREDITS, "$###,##0.00"), CS4, AlignConstants.vbAlignRight, True)
        PrintTo(OutputObject, Format(PMTDCREDITS, "$###,##0.00"), CS4, AlignConstants.vbAlignRight, True, 8600)

        OutputObject.PrintNL

        'OutputObject.Print
        ' Change Sign
        COHAND = MTDEBITS - MTDCREDITS
        PCOHAND = PMTDEBITS - PMTDCREDITS
        CashOut = CashOut + (COHAND - GetPrice(txtCashInDrawer.Text))
        PCashOut = PCashOut + (PCOHAND - GetPrice(txtCashInDrawer.Text))
        'PrintTo(OutputObject, "Calculated Cash On Hand", 3, AlignConstants.vbAlignLeft, False)
        PrintTo(OutputObject, "Calculated Cash On Hand", 3, AlignConstants.vbAlignLeft, False, 9000)
        'PrintTo(OutputObject, Format(COHAND, "$###,##0.00"), CS2, AlignConstants.vbAlignRight, False)
        PrintTo(OutputObject, Format(COHAND, "$###,##0.00"), CS2, AlignConstants.vbAlignRight, False, 9000)
        'PrintTo(OutputObject, Format(PCOHAND, "$###,##0.00"), CS4, AlignConstants.vbAlignRight, True)
        PrintTo(OutputObject, Format(PCOHAND, "$###,##0.00"), CS4, AlignConstants.vbAlignRight, True, 9000)

        'PrintTo(OutputObject, "Actual Cash On Hand", 3, AlignConstants.vbAlignLeft, False)
        PrintTo(OutputObject, "Actual Cash On Hand", 3, AlignConstants.vbAlignLeft, False, 9200)
        'PrintTo(OutputObject, Format(txtCashInDrawer, "-$###,##0.00"), CS2, AlignConstants.vbAlignRight, False)
        PrintTo(OutputObject, Format(txtCashInDrawer.Text, "-$###,##0.00"), CS2, AlignConstants.vbAlignRight, False, 9200)
        'PrintTo(OutputObject, Format(txtCashInDrawer, "-$###,##0.00"), CS4, AlignConstants.vbAlignRight, True)
        PrintTo(OutputObject, Format(txtCashInDrawer.Text, "-$###,##0.00"), CS4, AlignConstants.vbAlignRight, True, 9200)

        'PrintTo(OutputObject, New String("_", 65), CS4, AlignConstants.vbAlignRight, True)
        PrintTo(OutputObject, New String("_", 65), CS4, AlignConstants.vbAlignRight, True, 9600)

        'PrintTo(OutputObject, "Difference Over/Short   Acc: 52000", 2, AlignConstants.vbAlignLeft, False)
        PrintTo(OutputObject, "Difference Over/Short   Acc: 52000", 2, AlignConstants.vbAlignLeft, False, 9800)
        'PrintTo(OutputObject, Format(COHAND - GetPrice(txtCashInDrawer.Text), "$###,##0.00"), CS2, AlignConstants.vbAlignRight, False)
        PrintTo(OutputObject, Format(COHAND - GetPrice(txtCashInDrawer.Text), "$###,##0.00"), CS2, AlignConstants.vbAlignRight, False, 9800)
        'PrintTo(OutputObject, Format(PCOHAND - GetPrice(txtCashInDrawer.Text), "$###,##0.00"), CS4, AlignConstants.vbAlignRight, True)
        PrintTo(OutputObject, Format(PCOHAND - GetPrice(txtCashInDrawer.Text), "$###,##0.00"), CS4, AlignConstants.vbAlignRight, True, 9800)

        'PrintTo(OutputObject, New String("10", 3), 0, AlignConstants.vbAlignLeft, False)
        'PrintTo(OutputObject, "Gross Receipts (No Refunds)", 0, AlignConstants.vbAlignLeft, False)
        PrintTo(OutputObject, "Gross Receipts (No Refunds)", 0, AlignConstants.vbAlignLeft, False, 10400)
        'PrintTo(OutputObject, Format(NETRECEIPTS - RSALES, "$###,##0.00"), CS1, AlignConstants.vbAlignRight, False)
        PrintTo(OutputObject, Format(NETRECEIPTS - RSALES, "$###,##0.00"), CS1, AlignConstants.vbAlignRight, False, 10400)
        'PrintTo(OutputObject, Format(PNETRECEIPTS - PRSALES, "$###,##0.00"), CS3, AlignConstants.vbAlignRight, True)
        PrintTo(OutputObject, Format(PNETRECEIPTS - PRSALES, "$###,##0.00"), CS3, AlignConstants.vbAlignRight, True, 10400)

        'PrintTo(OutputObject, "Gross Refunds", 0, AlignConstants.vbAlignLeft, False)
        PrintTo(OutputObject, "Gross Refunds", 0, AlignConstants.vbAlignLeft, False, 10600)
        'PrintTo(OutputObject, Format(RSALES + CheckRefund, "$###,##0.00"), CS1, AlignConstants.vbAlignRight, False)
        PrintTo(OutputObject, Format(RSALES + CheckRefund, "$###,##0.00"), CS1, AlignConstants.vbAlignRight, False, 10600)
        'PrintTo(OutputObject, Format(PRSALES + CheckRefund, "$###,##0.00"), CS3, AlignConstants.vbAlignRight, True)
        PrintTo(OutputObject, Format(PRSALES + CheckRefund, "$###,##0.00"), CS3, AlignConstants.vbAlignRight, True, 10600)

        '    OutputObject.Print Tab(31); "____________________________"
        'PrintTo(OutputObject, "______________", CS1, AlignConstants.vbAlignRight, False)
        PrintTo(OutputObject, "______________", CS1, AlignConstants.vbAlignRight, False, 11000)
        'PrintTo(OutputObject, "______________", CS3, AlignConstants.vbAlignRight, True)
        PrintTo(OutputObject, "______________", CS3, AlignConstants.vbAlignRight, True, 11000)

        'PrintTo(OutputObject, "Net Receipts", 10, AlignConstants.vbAlignLeft, False)
        PrintTo(OutputObject, "Net Receipts", 10, AlignConstants.vbAlignLeft, False, 11200)
        'PrintTo(OutputObject, Format(NETRECEIPTS + CheckRefund, "$###,##0.00"), CS1, AlignConstants.vbAlignRight, False)
        PrintTo(OutputObject, Format(NETRECEIPTS + CheckRefund, "$###,##0.00"), CS1, AlignConstants.vbAlignRight, False, 11200)
        'PrintTo(OutputObject, Format(PNETRECEIPTS + CheckRefund, "$###,##0.00"), CS3, AlignConstants.vbAlignRight, True)
        PrintTo(OutputObject, Format(PNETRECEIPTS + CheckRefund, "$###,##0.00"), CS3, AlignConstants.vbAlignRight, True, 11200)

        If OutputToPrinter Then
            If OutputObject.CurrentY <> 0 Then OutputObject.NewPage
        Else
            frmPrintPreviewDocument.NewPage()
        End If
    End Sub

    Private Sub ArLate()
        PINTEREST = 0
        Counter = 0

        If optDetail.Checked = True Then 'detail
            If OutputToPrinter Then
                If OutputObject.CurrentY <> 0 Then OutputObject.NewPage
            Else
                frmPrintPreviewDocument.NewPage()
            End If

            Index = 2
            Printer.FontSize = 8
            Header = 6
            Headings()
            SubHeading()

            'PrintTo(OutputObject, CurrencyFormat(INTEREST), 32, AlignConstants.vbAlignRight, False)
            PrintTo(OutputObject, CurrencyFormat(INTEREST), 32, AlignConstants.vbAlignRight, False, 980)
            'PrintTo(OutputObject, "*** Previous Balance ***", 85, AlignConstants.vbAlignLeft, True)
            PrintTo(OutputObject, "*** Previous Balance ***", 85, AlignConstants.vbAlignLeft, True, 980)
        End If

        If Not (CashJournalRecordSet.EOF And CashJournalRecordSet.BOF) Then
            CashJournalRecordSet.MoveFirst()
            If optDetail.Checked = True Then OutputObject.PrintNL
        End If
        Do Until CashJournalRecordSet.EOF
            CashJournalNew_RecordSet_Set(NewCash, CashJournalRecordSet)
            Application.DoEvents()
            NewCash.Money = GetPrice(NewCash.Money)

            'Period To date
            Typee = Val(NewCash.Account)
            If Typee >= 560 And Typee <= 569 Then
                ConvertCode
                PSalesDistribution()
                If optDetail.Checked = True Then 'detail
                    PrintLines
                End If
                PINTEREST = PINTEREST + NewCash.Money
                NewCash.Account = Mid(Trim(NewCash.Account), 3, 1)  ' I don't see how this can possibly help anything.
            End If
            CashJournalRecordSet.MoveNext()

            If Counter >= 72 Then
                If OutputToPrinter Then
                    If OutputObject.CurrentY > 0 Then OutputObject.NewPage
                Else
                    frmPrintPreviewDocument.NewPage()
                End If
                Counter = 0
                Headings()
                SubHeading
            End If
        Loop

        INTEREST = INTEREST + PINTEREST

        If optDetail.Checked = True Then 'detail
            OutputObject.PrintNL
            'PrintTo(OutputObject, "Period To Date:", 0, AlignConstants.vbAlignLeft, False)
            PrintTo(OutputObject, "Period To Date:", 0, AlignConstants.vbAlignLeft, False, 1280)
            'PrintTo(OutputObject, Format(PINTEREST, "$###,##0.00"), 32, AlignConstants.vbAlignRight, True)
            PrintTo(OutputObject, Format(PINTEREST, "$###,##0.00"), 32, AlignConstants.vbAlignRight, True, 1280)

            'PrintTo(OutputObject, "Month To Date:", 0, AlignConstants.vbAlignLeft, False)
            PrintTo(OutputObject, "Month To Date:", 0, AlignConstants.vbAlignLeft, False, 1450)
            'PrintTo(OutputObject, Format(INTEREST, "$###,##0.00"), 32, AlignConstants.vbAlignRight, True)
            PrintTo(OutputObject, Format(INTEREST, "$###,##0.00"), 32, AlignConstants.vbAlignRight, True, 1450)
            If Counter > 56 Then
                If OutputToPrinter Then
                    If OutputObject.CurrentY > 0 Then OutputObject.NewPage
                Else
                    frmPrintPreviewDocument.NewPage()
                End If
                Counter = 0
                Headings()
            End If
        End If
    End Sub

    Private Sub PrintLines()
        'Big sections
        OutputObject.FontName = "Arial"
        OutputObject.FontBold = False
        OutputObject.FontSize = 8
        'If Index <> 4 Then
        PrintTo(OutputObject, Trim(NewCash.LeaseNo), 5, AlignConstants.vbAlignLeft, False)
        PrintTo(OutputObject, Format(Trim(NewCash.Money), "###,##0.00"), 37, AlignConstants.vbAlignRight, False)
        PrintTo(OutputObject, Typee, 40, AlignConstants.vbAlignLeft, False)
        PrintTo(OutputObject, Trim(NewCash.TransDate), 64, AlignConstants.vbAlignRight, False)
        PrintTo(OutputObject, Trim(NewCash.Note), 83, AlignConstants.vbAlignLeft, False)
        PrintTo(OutputObject, IIf(Trim(NewCash.Cashier) = "DEMO", "", Trim(NewCash.Cashier)), 125, AlignConstants.vbAlignLeft, True)
        Counter = Counter + 1
        PageCheck
    End Sub

    Private Sub PageCheck()
        If Counter >= 68 Or (OutputObject.CurrentY + OutputObject.TextHeight("X") > Printer.ScaleHeight) Then

            If OutputToPrinter Then
                If OutputObject.CurrentY <> 0 Then OutputObject.NewPage
            Else
                frmPrintPreviewDocument.NewPage()
            End If

            Counter = 0
            Headings()
            SubHeading()
        End If
    End Sub

    Private Sub PSalesDistribution()
        'Period To Date
        ' 1-9, 551-559, 561-569

        If Val(NewCash.Account) = 12 Then
            PSTORECARD = PSTORECARD + NewCash.Money
        ElseIf Val(NewCash.Account) = 13 Then
            PECHECKS = PECHECKS + NewCash.Money
        ElseIf Val(NewCash.Account) = 15 Then
            PCheck = PCheck + NewCash.Money
        Else
            Select Case Val(Microsoft.VisualBasic.Right(Trim(NewCash.Account), 1))
                Case 1
                    PCASHSALES = PCASHSALES + NewCash.Money
                Case 2
                    PCheck = PCheck + NewCash.Money
                Case 3
                    PVISA = PVISA + NewCash.Money
                Case 4
                    PMASTER = PMASTER + NewCash.Money
                Case 5
                    PDISC = PDISC + NewCash.Money
                Case 6
                    PAMCX = PAMCX + NewCash.Money
                Case 8
                    PFINC = PFINC + NewCash.Money
                Case 9
                    PNONE = PNONE + NewCash.Money 'debit
            End Select
        End If
        SalesDistribution  ' Also add to MTD.
    End Sub

    Private Sub SetWorking(ByVal Working As Boolean)
        cmdPrint.Enabled = Not Working
        cmdPrintPreview.Enabled = Not Working
        cmdCancel.Enabled = Not Working

        lblStartDateLabel.Enabled = Not Working
        lblEndDateLabel.Enabled = Not Working
        dteStartDate.Enabled = Not Working
        dteEndDate.Enabled = Not Working

        optDetail.Enabled = Not Working
        optSummary.Enabled = Not Working

        lblCashInDrawer.Enabled = Not Working
        lblPriorPeriodCash.Enabled = Not Working
        txtCashInDrawer.Enabled = Not Working
        txtPriorPeriodCash.Enabled = Not Working

        lblCashier.Enabled = Not Working
        cmbCashier.Enabled = Not Working

        'MousePointer = IIf(Working, vbHourglass, vbDefault)
        Me.Cursor = IIf(Working, Cursors.WaitCursor, Cursors.Default)
    End Sub

    Private Sub SalesDistribution()
        ' 1-9, 551-559, 561-569
        If Val(NewCash.Account) = 12 Then
            STORECARD = STORECARD + NewCash.Money
        ElseIf Val(NewCash.Account) = 13 Then
            ECHECKS = ECHECKS + NewCash.Money
        ElseIf Val(NewCash.Account) = 15 Then
            Check = Check + NewCash.Money
        Else
            Select Case Val(Microsoft.VisualBasic.Right(Trim(NewCash.Account), 1))
                Case 1
                    CASHSALES = CASHSALES + NewCash.Money
                Case 2
                    Check = Check + NewCash.Money
                Case 3
                    VISA = VISA + NewCash.Money
                Case 4
                    MASTER = MASTER + NewCash.Money
                Case 5
                    Disc = Disc + NewCash.Money
                Case 6
                    AMCX = AMCX + NewCash.Money
                Case 8
                    FINC = FINC + NewCash.Money
                Case 9
                    None = None + NewCash.Money
            End Select
        End If
    End Sub

    Private Sub ConvertCode()
        Typee = TranslateAccountCode(NewCash.Account, Typee)
    End Sub

    Private Sub SubHeading()
        Select Case Index
            Case 1
                'PrintTo(OutputObject, "SALE", 0, AlignConstants.vbAlignLeft, False)
                PrintTo(OutputObject, "SALE", 0, AlignConstants.vbAlignLeft, False, 800)
                'PrintTo(OutputObject, "TYPE NAME", 12, AlignConstants.vbAlignLeft, False)
                PrintTo(OutputObject, "TYPE NAME", 12, AlignConstants.vbAlignLeft, False, 800)
                'PrintTo(OutputObject, "ENTRY DATE", 33, AlignConstants.vbAlignLeft, False)
                PrintTo(OutputObject, "ENTRY DATE", 33, AlignConstants.vbAlignLeft, False, 800)
                'PrintTo(OutputObject, "WRITTEN", 57, AlignConstants.vbAlignRight, False)
                PrintTo(OutputObject, "WRITTEN", 57, AlignConstants.vbAlignRight, False, 800)
                'PrintTo(OutputObject, "TAX", 71, AlignConstants.vbAlignRight, False)
                PrintTo(OutputObject, "TAX", 71, AlignConstants.vbAlignRight, False, 800)
                '      PrintTo OutputObject, "AR CSH SLS", 86, alignconstants.vbAlignRight, False)
                'PrintTo(OutputObject, "CUST DEP", 86, AlignConstants.vbAlignRight, False)
                PrintTo(OutputObject, "CUST DEP", 86, AlignConstants.vbAlignRight, False, 800)
                'PrintTo(OutputObject, "UND. SLS", 99, AlignConstants.vbAlignRight, False)
                PrintTo(OutputObject, "UND. SLS", 99, AlignConstants.vbAlignRight, False, 800)
                'PrintTo(OutputObject, "DEL SALES", 112, AlignConstants.vbAlignRight, False)
                PrintTo(OutputObject, "DEL SALES", 112, AlignConstants.vbAlignRight, False, 800)
                'PrintTo(OutputObject, "DEL SALES", 112, AlignConstants.vbAlignRight, False)
                PrintTo(OutputObject, "TAX REC.", 125, AlignConstants.vbAlignRight, False, 800)
                'PrintTo(OutputObject, "CASHIER", 130, AlignConstants.vbAlignLeft, True)
                PrintTo(OutputObject, "CASHIER", 130, AlignConstants.vbAlignLeft, True, 800)
            Case 2
                'PrintTo(OutputObject, "SALE NO.", 5, AlignConstants.vbAlignLeft, False)
                PrintTo(OutputObject, "SALE NO.", 5, AlignConstants.vbAlignLeft, False, 800)
                'PrintTo(OutputObject, "CASH", 37, AlignConstants.vbAlignRight, False)
                PrintTo(OutputObject, "CASH", 37, AlignConstants.vbAlignRight, False, 800)
                'PrintTo(OutputObject, "TYPE", 40, AlignConstants.vbAlignLeft, False)
                PrintTo(OutputObject, "TYPE", 40, AlignConstants.vbAlignLeft, False, 800)
                'PrintTo(OutputObject, "TRANS DATE", 64, AlignConstants.vbAlignRight, False)
                PrintTo(OutputObject, "TRANS DATE", 64, AlignConstants.vbAlignRight, False, 800)
                'PrintTo(OutputObject, "NAME", 83, AlignConstants.vbAlignLeft, False)
                PrintTo(OutputObject, "NAME", 83, AlignConstants.vbAlignLeft, False, 800)
                'PrintTo(OutputObject, "CASHIER", 125, AlignConstants.vbAlignLeft, True)
                PrintTo(OutputObject, "CASHIER", 125, AlignConstants.vbAlignLeft, True, 800)
            Case 3
                PrintTo(OutputObject, "A/R NO.", 5, AlignConstants.vbAlignLeft, False)
                PrintTo(OutputObject, "CASH", 37, AlignConstants.vbAlignRight, False)
                PrintTo(OutputObject, "ACC", 40, AlignConstants.vbAlignLeft, False)
                PrintTo(OutputObject, "TRANS DATE", 53, AlignConstants.vbAlignLeft, False)
                PrintTo(OutputObject, "NAME", 85, AlignConstants.vbAlignLeft, True)
            Case 4
                'PrintTo(OutputObject, "CASH", 35, AlignConstants.vbAlignRight, False)
                PrintTo(OutputObject, "CASH", 35, AlignConstants.vbAlignRight, False, 800)
                'PrintTo(OutputObject, "ACC", 51, AlignConstants.vbAlignRight, False)
                PrintTo(OutputObject, "ACC", 51, AlignConstants.vbAlignRight, False, 800)
                'PrintTo(OutputObject, "TRANS DATE", 73, AlignConstants.vbAlignRight, False)
                PrintTo(OutputObject, "TRANS DATE", 73, AlignConstants.vbAlignRight, False, 800)
                'PrintTo(OutputObject, "COMMENTS", 85, AlignConstants.vbAlignLeft, False)
                PrintTo(OutputObject, "COMMENTS", 85, AlignConstants.vbAlignLeft, False, 800)
                'PrintTo(OutputObject, "CASHIER", 125, AlignConstants.vbAlignLeft, True)
                PrintTo(OutputObject, "CASHIER", 125, AlignConstants.vbAlignLeft, True, 800)
            Case 5
                PrintTo(OutputObject, "SALE NO.", 5, AlignConstants.vbAlignLeft, False)
                PrintTo(OutputObject, "ORIG DEP", 36, AlignConstants.vbAlignLeft, False)
                PrintTo(OutputObject, "ACC", 38, AlignConstants.vbAlignLeft, False)
                PrintTo(OutputObject, "TAX", 49, AlignConstants.vbAlignLeft, False)
                PrintTo(OutputObject, "NAME/REWRITE", 80, AlignConstants.vbAlignLeft, True)
            Case 6
                PrintTo(OutputObject, "INDEX", 10, AlignConstants.vbAlignRight, False)
                PrintTo(OutputObject, "SALE NO.", 13, AlignConstants.vbAlignLeft, False)
                PrintTo(OutputObject, "ACC", 25, AlignConstants.vbAlignLeft, False)
                PrintTo(OutputObject, "TRANS DATE", 55, AlignConstants.vbAlignRight, False)
                PrintTo(OutputObject, "COMMENTS", 57, AlignConstants.vbAlignLeft, True)
            Case 7 '<CT> Misc. Cash Out </CT>
                'PrintTo(OutputObject, "CASH", 35, AlignConstants.vbAlignRight, False)
                PrintTo(OutputObject, "CASH", 35, AlignConstants.vbAlignRight, False, 2650)
                'PrintTo(OutputObject, "ACC", 51, AlignConstants.vbAlignRight, False)
                PrintTo(OutputObject, "ACC", 51, AlignConstants.vbAlignRight, False, 2650)
                'PrintTo(OutputObject, "TRANS DATE", 73, AlignConstants.vbAlignRight, False)
                PrintTo(OutputObject, "TRANS DATE", 73, AlignConstants.vbAlignRight, False, 2650)
                'PrintTo(OutputObject, "COMMENTS", 85, AlignConstants.vbAlignLeft, False)
                PrintTo(OutputObject, "COMMENTS", 85, AlignConstants.vbAlignLeft, False, 2650)
                'PrintTo(OutputObject, "CASHIER", 125, AlignConstants.vbAlignLeft, True)
                PrintTo(OutputObject, "CASHIER", 125, AlignConstants.vbAlignLeft, True, 2650)
            Case 8 '<CT> Bank Deposits </CT>
                'PrintTo(OutputObject, "CASH", 35, AlignConstants.vbAlignRight, False)
                PrintTo(OutputObject, "CASH", 35, AlignConstants.vbAlignRight, False, 4770)
                'PrintTo(OutputObject, "ACC", 51, AlignConstants.vbAlignRight, False)
                PrintTo(OutputObject, "ACC", 51, AlignConstants.vbAlignRight, False, 4770)
                'PrintTo(OutputObject, "TRANS DATE", 73, AlignConstants.vbAlignRight, False)
                PrintTo(OutputObject, "TRANS DATE", 73, AlignConstants.vbAlignRight, False, 4770)
                'PrintTo(OutputObject, "COMMENTS", 85, AlignConstants.vbAlignLeft, False)
                PrintTo(OutputObject, "COMMENTS", 85, AlignConstants.vbAlignLeft, False, 4770)
                'PrintTo(OutputObject, "CASHIER", 125, AlignConstants.vbAlignLeft, True)
                PrintTo(OutputObject, "CASHIER", 125, AlignConstants.vbAlignLeft, True, 4770)
            Case 9 '<CT> Check Refund/Forfeit </CT>
                'PrintTo(OutputObject, "SALE NO.", 5, AlignConstants.vbAlignLeft, False)
                PrintTo(OutputObject, "SALE NO.", 5, AlignConstants.vbAlignLeft, False, 6900)
                'PrintTo(OutputObject, "CASH", 37, AlignConstants.vbAlignRight, False)
                PrintTo(OutputObject, "CASH", 37, AlignConstants.vbAlignRight, False, 6900)
                'PrintTo(OutputObject, "TYPE", 40, AlignConstants.vbAlignLeft, False)
                PrintTo(OutputObject, "TYPE", 40, AlignConstants.vbAlignLeft, False, 6900)
                'PrintTo(OutputObject, "TRANS DATE", 64, AlignConstants.vbAlignRight, False)
                PrintTo(OutputObject, "TRANS DATE", 64, AlignConstants.vbAlignRight, False, 6900)
                'PrintTo(OutputObject, "NAME", 83, AlignConstants.vbAlignLeft, False)
                PrintTo(OutputObject, "NAME", 83, AlignConstants.vbAlignLeft, False, 6900)
                'PrintTo(OutputObject, "CASHIER", 125, AlignConstants.vbAlignLeft, True)
                PrintTo(OutputObject, "CASHIER", 125, AlignConstants.vbAlignLeft, True, 6900)
        End Select
        OutputObject.FontBold = False
        '  OutputObject.CurrentY = 1000
    End Sub

    Private Sub ArPayments()
        PAR = 0
        Counter = 0

        If optDetail.Checked = True Then 'detail
            If OutputToPrinter Then
                If OutputObject.CurrentY <> 0 Then OutputObject.NewPage
            Else
                frmPrintPreviewDocument.NewPage()
            End If

            Header = 5
            Headings()
            SubHeading()

            'PrintTo(OutputObject, CurrencyFormat(AR), 32, AlignConstants.vbAlignRight, False)
            PrintTo(OutputObject, CurrencyFormat(AR), 32, AlignConstants.vbAlignRight, False, 980)
            'PrintTo(OutputObject, "*** Previous Balance ***", 85, AlignConstants.vbAlignLeft, True)
            PrintTo(OutputObject, "*** Previous Balance ***", 85, AlignConstants.vbAlignLeft, True, 980)
        End If

        If Not (CashJournalRecordSet.EOF And CashJournalRecordSet.BOF) Then
            CashJournalRecordSet.MoveFirst()
            If optDetail.Checked = True Then OutputObject.PrintNL
        End If
        Do Until CashJournalRecordSet.EOF
            CashJournalNew_RecordSet_Set(NewCash, CashJournalRecordSet)
            Application.DoEvents()
            NewCash.Money = GetPrice(NewCash.Money)

            'Period to date
            Typee = Val(NewCash.Account)
            If (Typee >= 550 And Typee <= 559) Or Typee = 15 Then
                ConvertCode                          ' Turn account # into a description
                PSalesDistribution()              ' This separates charges into categories, which we need..
                If optDetail.Checked = True Then PrintLinesCash  'detail, print the converted code instead of account #.
                PAR = PAR + NewCash.Money
                NewCash.Account = Mid(NewCash.Account, 3, 1)
            End If

            CashJournalRecordSet.MoveNext()

            If Counter >= 72 Then
                If OutputToPrinter Then
                    If OutputObject.CurrentY > 0 Then OutputObject.NewPage
                Else
                    frmPrintPreviewDocument.NewPage()
                End If
                Counter = 0
                Headings()
                SubHeading
            End If
        Loop

        AR = AR + PAR

        If optDetail.Checked = True Then 'detail
            OutputObject.PrintNL
            'PrintTo(OutputObject, "Period To Date:", 0, AlignConstants.vbAlignLeft, False)
            PrintTo(OutputObject, "Period To Date:", 0, AlignConstants.vbAlignLeft, False, 1280)
            'PrintTo(OutputObject, Format(PAR, "$###,##0.00"), 32, AlignConstants.vbAlignRight, True)
            PrintTo(OutputObject, Format(PAR, "$###,##0.00"), 32, AlignConstants.vbAlignRight, True, 1280)

            'PrintTo(OutputObject, "Month To Date:", 0, AlignConstants.vbAlignLeft, False)
            PrintTo(OutputObject, "Month To Date:", 0, AlignConstants.vbAlignLeft, False, 1450)
            'PrintTo(OutputObject, Format(AR, "$###,##0.00"), 32, AlignConstants.vbAlignRight, True)
            PrintTo(OutputObject, Format(AR, "$###,##0.00"), 32, AlignConstants.vbAlignRight, True, 1450)
        End If
    End Sub

    Private Sub PrintLinesCash()
        'Cash & A/R payments
        OutputObject.FontName = "Arial"
        OutputObject.FontBold = False
        OutputObject.FontSize = 8
        ConvertCode()

        PrintTo(OutputObject, Trim(NewCash.LeaseNo), 5, AlignConstants.vbAlignLeft, False)
        PrintTo(OutputObject, Format(NewCash.Money, "###,##0.00"), 37, AlignConstants.vbAlignRight, False)
        PrintTo(OutputObject, Trim(Typee), 40, AlignConstants.vbAlignLeft, False)
        PrintTo(OutputObject, Trim(NewCash.TransDate), 64, AlignConstants.vbAlignRight, False)
        PrintTo(OutputObject, Trim(NewCash.Note), 83, AlignConstants.vbAlignLeft, False)
        PrintTo(OutputObject, IIf(Trim(NewCash.Cashier) = "DEMO", "", Trim(NewCash.Cashier)), 125, AlignConstants.vbAlignLeft, True)

        Counter = Counter + 1
        PageCheck()
    End Sub

    Private Sub PrintLinesCashOutOfDate(ByVal Index As Integer)
        OutputObject.FontName = "Arial"
        OutputObject.FontBold = False
        OutputObject.FontSize = 8
        ConvertCode()

        PrintTo(OutputObject, Index, 10, AlignConstants.vbAlignRight, False)
        PrintTo(OutputObject, NewCash.LeaseNo, 13, AlignConstants.vbAlignLeft, False)
        PrintTo(OutputObject, Typee, 25, AlignConstants.vbAlignLeft, False)
        PrintTo(OutputObject, NewCash.TransDate, 55, AlignConstants.vbAlignRight, False)
        PrintTo(OutputObject, NewCash.Note, 57, AlignConstants.vbAlignLeft, True)
        Counter = Counter + 1
    End Sub

    Private Sub RecControl() '****************** Payments & Credits On Control ****************
        PRECPAYMENTS = 0
        Counter = 0

        If optDetail.Checked = True Then 'detail
            If OutputToPrinter Then
                If OutputObject.CurrentY <> 0 Then OutputObject.NewPage
            Else
                frmPrintPreviewDocument.NewPage()
            End If

            Index = 2
            Header = 4
            Headings()
            SubHeading()

            'PrintTo(OutputObject, CurrencyFormat(RECPAYMENTS), 32, AlignConstants.vbAlignRight, False)
            PrintTo(OutputObject, CurrencyFormat(RECPAYMENTS), 32, AlignConstants.vbAlignRight, False, 980)
            'PrintTo(OutputObject, "*** Previous Balance ***", 85, AlignConstants.vbAlignLeft, True)
            PrintTo(OutputObject, "*** Previous Balance ***", 85, AlignConstants.vbAlignLeft, True, 980)
        End If

        If Not (CashJournalRecordSet.EOF And CashJournalRecordSet.BOF) Then
            CashJournalRecordSet.MoveFirst()
            If optDetail.Checked = True Then OutputObject.PrintNL
        End If
        Do Until CashJournalRecordSet.EOF
            CashJournalNew_RecordSet_Set(NewCash, CashJournalRecordSet)
            Application.DoEvents()
            NewCash.Money = GetPrice(NewCash.Money)
            'Period to date
            Typee = Val(NewCash.Account)
            If Typee = 11200 Or Typee = 11300 Or Typee = 13200 Then     ' 11200 for B/O and 11300 for installment module
                If optDetail.Checked = True Then PrintLinesCash()                          ' 8/2008  13200 is wrong account for store finance
                PRECPAYMENTS = PRECPAYMENTS + NewCash.Money
            End If
            CashJournalRecordSet.MoveNext()
        Loop

        RECPAYMENTS = RECPAYMENTS + PRECPAYMENTS
        If optDetail.Checked = True Then 'detail
            OutputObject.PrintNL
            'PrintTo(OutputObject, "Period To Date:", 5, AlignConstants.vbAlignLeft, False)
            PrintTo(OutputObject, "Period To Date:", 5, AlignConstants.vbAlignLeft, False, 1200)
            'PrintTo(OutputObject, Format(PRECPAYMENTS, "$###,##0.00"), 32, AlignConstants.vbAlignRight, True)
            PrintTo(OutputObject, Format(PRECPAYMENTS, "$###,##0.00"), 32, AlignConstants.vbAlignRight, True, 1200)

            'PrintTo(OutputObject, "Month To Date:", 5, AlignConstants.vbAlignLeft, False)
            PrintTo(OutputObject, "Month To Date:", 5, AlignConstants.vbAlignLeft, False, 1450)
            'PrintTo(OutputObject, Format(RECPAYMENTS, "$###,##0.00"), 32, AlignConstants.vbAlignRight, True)
            PrintTo(OutputObject, Format(RECPAYMENTS, "$###,##0.00"), 32, AlignConstants.vbAlignRight, True, 1450)
        End If
    End Sub

    Private Sub CheckForfeit()
        '21500 check refunds
        '41500 Forfeit
        Header = 11
        '  Counter = 0  ' Don't reset line count.

        '  CheckRefund = 0
        PCHECKREFUND = 0
        PFORFEIT = 0

        If optDetail.Checked = True Then 'detail
            OutputObject.PrintNL(vbCrLf)
            OutputObject.FontName = "Arial"
            OutputObject.FontSize = 18
            OutputObject.FontBold = True
            PrintCentered("Check Refund/Forfeit (21500/41500)")
            OutputObject.FontBold = False
            OutputObject.FontSize = 8
            Header = 33
            '<CT>
            'Index = 2
            Index = 9
            '</CT>
            SubHeading()

            'PrintTo(OutputObject, CurrencyFormat(CheckRefund + Forfeit), 32, AlignConstants.vbAlignRight, False)
            PrintTo(OutputObject, CurrencyFormat(CheckRefund + Forfeit), 32, AlignConstants.vbAlignRight, False, 7100)
            'PrintTo(OutputObject, "*** Previous Balance ***", 85, AlignConstants.vbAlignLeft, True)
            PrintTo(OutputObject, "*** Previous Balance ***", 85, AlignConstants.vbAlignLeft, True, 7100)
        End If

        If Not (CashJournalRecordSet.EOF And CashJournalRecordSet.BOF) Then
            CashJournalRecordSet.MoveFirst()
            If optDetail.Checked = True Then OutputObject.PrintNL
        End If
        Do Until CashJournalRecordSet.EOF
            CashJournalNew_RecordSet_Set(NewCash, CashJournalRecordSet)
            Application.DoEvents()
            NewCash.Money = GetPrice(NewCash.Money)
            'Period to date
            Typee = Val(NewCash.Account)
            Select Case Val(NewCash.Account)
                Case 21500
                    If optDetail.Checked = True Then PrintLinesCash()          'detail
                    PCHECKREFUND = PCHECKREFUND + NewCash.Money
'        PFORFEIT = PFORFEIT + NewCash.Money

' 01-10-2003  return a refund check? Shows up 2 times
                Case 41500
                    If optDetail.Checked = True Then PrintLines()
                    PFORFEIT = PFORFEIT + NewCash.Money
                    '        Pcash = Pcash + NewCash.Money
            End Select
            CashJournalRecordSet.MoveNext()
        Loop

        Forfeit = Forfeit + PFORFEIT
        CheckRefund = CheckRefund + PCHECKREFUND

        '  TOTCASHIN = TOTCASHIN + FORFEIT
        '  PTOTCASHIN = PTOTCASHIN + PFORFEIT

        If optDetail.Checked = True Then 'detail
            OutputObject.PrintNL
            PageCheck()
            'PrintTo(OutputObject, "Period To Date:", 0, AlignConstants.vbAlignLeft, False)
            PrintTo(OutputObject, "Period To Date:", 0, AlignConstants.vbAlignLeft, False, 7440)
            'PrintTo(OutputObject, Format(PCHECKREFUND + PFORFEIT, "$###,##0.00"), 32, AlignConstants.vbAlignRight, True)
            PrintTo(OutputObject, Format(PCHECKREFUND + PFORFEIT, "$###,##0.00"), 32, AlignConstants.vbAlignRight, True, 7440)
            PageCheck()

            'PrintTo(OutputObject, "Month To Date:", 0, AlignConstants.vbAlignLeft, False)
            PrintTo(OutputObject, "Month To Date:", 0, AlignConstants.vbAlignLeft, False, 7640)
            'PrintTo(OutputObject, Format(CheckRefund + Forfeit, "$###,##0.00"), 32, AlignConstants.vbAlignRight, True)
            PrintTo(OutputObject, Format(CheckRefund + Forfeit, "$###,##0.00"), 32, AlignConstants.vbAlignRight, True, 7640)
        End If
    End Sub

    Private Sub Banking()
        PBANK = 0
        '  Counter = 0  ' Prevents page breaks.

        If optDetail.Checked Then 'detail
            OutputObject.PrintNL(vbCrLf)
            OutputObject.FontName = "Arial"
            OutputObject.FontSize = 18
            OutputObject.FontBold = True
            PrintCentered("Bank Deposits")
            OutputObject.FontSize = 8
            OutputObject.FontBold = False
            Header = 32
            '<CT>
            'Index = 4
            Index = 8
            '</CT>
            SubHeading()

            'PrintTo(OutputObject, CurrencyFormat(Bank), 32, AlignConstants.vbAlignRight, False)
            PrintTo(OutputObject, CurrencyFormat(Bank), 32, AlignConstants.vbAlignRight, False, 4970)
            'PrintTo(OutputObject, "*** Previous Balance ***", 85, AlignConstants.vbAlignLeft, True)
            PrintTo(OutputObject, "*** Previous Balance ***", 85, AlignConstants.vbAlignLeft, True, 4970)
        End If

        If Not (CashJournalRecordSet.EOF And CashJournalRecordSet.BOF) Then
            CashJournalRecordSet.MoveFirst()
            If optDetail.Checked = True Then OutputObject.PrintNL
        End If
        Do Until CashJournalRecordSet.EOF
            CashJournalNew_RecordSet_Set(NewCash, CashJournalRecordSet)
            Application.DoEvents()
            NewCash.Money = GetPrice(NewCash.Money)

            ' Period to date
            Typee = Val(NewCash.Account)
            If Val(NewCash.Account) >= 10200 And Val(NewCash.Account) <= 10650 Then
                PBANK = PBANK + NewCash.Money
                BankTypePeriod
                If optDetail.Checked = True Then PrintLinesMisc        'detail
            End If
            CashJournalRecordSet.MoveNext()
        Loop

        Bank = Bank + PBANK

        BCASHSALES = BCASHSALES + PBCASHSALES
        BECHECKS = BECHECKS + PBECHECKS
        BVISA = BVISA + PBVISA
        BDisc = BDisc + PBDisc
        BAmcx = BAmcx + PBAmcx
        Debit = Debit + PDebit
        BSTORECARD = BSTORECARD + PBSTORECARD

        If optDetail.Checked = True Then 'detail
            OutputObject.PrintNL
            'PrintTo(OutputObject, "Period To Date:", 0, AlignConstants.vbAlignLeft, False)
            PrintTo(OutputObject, "Period To Date:", 0, AlignConstants.vbAlignLeft, False, 5370)
            'PrintTo(OutputObject, Format(PBANK, "$###,##0.00"), 32, AlignConstants.vbAlignRight, True)
            PrintTo(OutputObject, Format(PBANK, "$###,##0.00"), 32, AlignConstants.vbAlignRight, True, 5370)
            PageCheck()

            'PrintTo(OutputObject, "Month To Date:", 0, AlignConstants.vbAlignLeft, False)
            PrintTo(OutputObject, "Month To Date:", 0, AlignConstants.vbAlignLeft, False, 5570)
            'PrintTo(OutputObject, Format(Bank, "$###,##0.00"), 32, AlignConstants.vbAlignRight, True)
            PrintTo(OutputObject, Format(Bank, "$###,##0.00"), 32, AlignConstants.vbAlignRight, True, 5570)
            PageCheck()
        End If
    End Sub

    Private Sub PrintLinesMisc()
        'Misc Cash and Bank reports
        OutputObject.FontName = "Arial"
        OutputObject.FontBold = False
        OutputObject.FontSize = 8

        PrintTo(OutputObject, Format(NewCash.Money, "###,##0.00"), 35, AlignConstants.vbAlignRight, False)
        PrintTo(OutputObject, Trim(NewCash.Account), 51, AlignConstants.vbAlignRight, False)
        PrintTo(OutputObject, Trim(NewCash.TransDate), 73, AlignConstants.vbAlignRight, False)
        PrintTo(OutputObject, Trim(NewCash.Note), 83, AlignConstants.vbAlignLeft, False)
        PrintTo(OutputObject, IIf(Trim(NewCash.Cashier) = "DEMO", "", Trim(NewCash.Cashier)), 125, AlignConstants.vbAlignLeft, True)

        '    PrintTo OutputObject, Trim(Typee), 40, vbAlignLeft, False
        Counter = Counter + 1
        PageCheck()
    End Sub

    Private Sub BankTypePeriod()
        ' PERIOD BANK TOTALS
        Select Case Val(NewCash.Account)
            Case 10200
                PBCASHSALES = PBCASHSALES + NewCash.Money
            Case 10250
                PBECHECKS = PBECHECKS + NewCash.Money
            Case 10300
                PBVISA = PBVISA + NewCash.Money
            Case 10400
                PBDisc = PBDisc + NewCash.Money
            Case 10500
                PBAmcx = PBAmcx + NewCash.Money
            Case 10600
                PDebit = PDebit + NewCash.Money
            Case 10650
                PBSTORECARD = PBSTORECARD + NewCash.Money
        End Select
    End Sub

    Private Sub CashOutDwr()
        ' CashOut = 0
        PCashOut = 0
        '  Counter = 0 'Prevents page breaks..
        Counter = Counter + 7 ' account for headings

        If optDetail.Checked = True Then 'detail
            OutputObject.PrintNL(vbCrLf)
            OutputObject.FontName = "Arial"
            OutputObject.FontSize = 18
            OutputObject.FontBold = True
            PrintCentered("Misc. Cash Out")
            OutputObject.FontBold = False
            OutputObject.FontSize = 8

            Header = 31
            '<CT>
            'Index = 4
            Index = 7
            '</CT>
            SubHeading()

            'PrintTo(OutputObject, CurrencyFormat(CashOut), 32, AlignConstants.vbAlignRight, False)
            PrintTo(OutputObject, CurrencyFormat(CashOut), 32, AlignConstants.vbAlignRight, False, 2850)
            'PrintTo(OutputObject, "*** Previous Balance ***", 85, AlignConstants.vbAlignLeft, True)
            PrintTo(OutputObject, "*** Previous Balance ***", 85, AlignConstants.vbAlignLeft, True, 2850)
        End If

        If Not (CashJournalRecordSet.EOF And CashJournalRecordSet.BOF) Then
            CashJournalRecordSet.MoveFirst()
            If optDetail.Checked = True Then OutputObject.PrintNL  ' Extra space after previous balance.
        End If
        Do Until CashJournalRecordSet.EOF
            CashJournalNew_RecordSet_Set(NewCash, CashJournalRecordSet)
            Application.DoEvents()
            NewCash.Money = GetPrice(NewCash.Money)
            ' Period to date
            Typee = Val(NewCash.Account)
            Select Case Val(NewCash.Account)
      'Case 10000
        'PETTYCASH = PETTYCASH + NewCash.Money
        'PrintLines
        'Pcash = Pcash + NewCash.Money
                Case 50500
                    PFREIGHTOUT = PFREIGHTOUT + NewCash.Money
                Case 50600
                    PCREDIT = PCREDIT + NewCash.Money
                Case 60100
                    PGAS = PGAS + NewCash.Money
                Case 60500
                    PVISADISC = PVISADISC + NewCash.Money
                Case 62300
                    PMAINTENANCE = PMAINTENANCE + NewCash.Money
                Case 62400
                    PREPAIR = PREPAIR + NewCash.Money
                Case 63500
                    PWHSE = PWHSE + NewCash.Money
                Case 64100
                    POFFICE = POFFICE + NewCash.Money
                Case 65200
                    PCASUAL = PCASUAL + NewCash.Money
                Case 67500
                    PTRAVEL = PTRAVEL + NewCash.Money
                Case 99800
                    PMISCCASHOUT = PMISCCASHOUT + NewCash.Money
                Case 10000
                    PMISCCASHOUT = PMISCCASHOUT + NewCash.Money
                Case 50200
                    ' bfh20051104 - only show in cashout rpt now
                    '        If NewCash.Money >= 0 Then GoTo SkipLine
                    PPURCHASES = PPURCHASES + NewCash.Money
                Case Else
                    GoTo SkipLine
            End Select

            PCashOut = PCashOut + NewCash.Money
            If optDetail.Checked = True Then
                PrintLinesMisc()      'detail
            End If
SkipLine:
            CashJournalRecordSet.MoveNext()
        Loop

        FREIGHTOUT = FREIGHTOUT + PFREIGHTOUT
        Credit = Credit + PCREDIT
        GAS = GAS + PGAS
        VISADISC = VISADISC + PVISADISC
        MAINTENANCE = MAINTENANCE + PMAINTENANCE
        REPAIR = REPAIR + PREPAIR
        WHSE = WHSE + PWHSE
        OFFICE = OFFICE + POFFICE
        CASUAL = CASUAL + PCASUAL
        TRAVEL = TRAVEL + PTRAVEL
        MiscCashOut = MiscCashOut + PMISCCASHOUT
        PURCHASES = PURCHASES + PPURCHASES

        CashOut = CashOut + PCashOut
        If optDetail.Checked = True Then 'detail
            OutputObject.PrintNL
            'PrintTo(OutputObject, "Period To Date:", 0, AlignConstants.vbAlignLeft, False)
            PrintTo(OutputObject, "Period To Date:", 0, AlignConstants.vbAlignLeft, False, 3250)
            'PrintTo(OutputObject, Format(PCashOut, "$###,##0.00"), 32, AlignConstants.vbAlignRight, True)
            PrintTo(OutputObject, Format(PCashOut, "$###,##0.00"), 32, AlignConstants.vbAlignRight, True, 3250)
            Counter = Counter + 1
            PageCheck()

            'PrintTo(OutputObject, "Month To Date:", 0, AlignConstants.vbAlignLeft, False)
            PrintTo(OutputObject, "Month To Date:", 0, AlignConstants.vbAlignLeft, False, 3450)
            'PrintTo(OutputObject, Format(CashOut, "$###,##0.00"), 32, AlignConstants.vbAlignRight, True)
            PrintTo(OutputObject, Format(CashOut, "$###,##0.00"), 32, AlignConstants.vbAlignRight, True, 3450)
            Counter = Counter + 1
            PageCheck()
        End If
    End Sub

    Private Sub CashInReport()
        PCashIn = 0
        XPCashIn = 0
        Counter = 7  ' This should reset the counter, but account for headings..

        If optDetail.Checked = True Then 'detail
            '    With OutputObject
            '      OutputObject.Print vbCrLf
            '
            '      .FontName = "Arial"
            '      .FontSize = 18
            '      .FontBold = True
            '
            '      PrintCentered "Misc. Cash In"
            '      .FontBold = False
            '      .FontSize = 8
            '    End With

            Header = 3
            Index = 4
            Headings()
            SubHeading()
            'OutputObject.FontSize = 11
            PrintTo(OutputObject, CurrencyFormat(CashIn), 32, AlignConstants.vbAlignRight, False, 980)
            PrintTo(OutputObject, "*** Previous Balance ***", 85, AlignConstants.vbAlignLeft, False, 980)
        End If

        If Not (CashJournalRecordSet.EOF And CashJournalRecordSet.BOF) Then
            CashJournalRecordSet.MoveFirst()
            If optDetail.Checked = True Then OutputObject.PrintNL  ' Extra space after previous balance.
        End If
        Do Until CashJournalRecordSet.EOF
            CashJournalNew_RecordSet_Set(NewCash, CashJournalRecordSet)
            Application.DoEvents()
            NewCash.Money = GetPrice(NewCash.Money)

            ' Period to date
            Typee = Val(NewCash.Account)

            Select Case Typee
                Case 10700
                    PVISACHECK = PVISACHECK + NewCash.Money
                Case 41500
                    PFORFEIT = PFORFEIT + NewCash.Money  ' See Forfeit report.
                    XPCashIn = XPCashIn + NewCash.Money
                    '        PCashIn = PCashIn + NewCash.Money
                    GoTo SkipLine
                Case 61600
                    PBCPAY = PBCPAY + NewCash.Money
                Case 70000
                    PFINANCE = PFINANCE + NewCash.Money
                Case 99900
                    PMISCCASHIN = PMISCCASHIN + NewCash.Money
                Case 50200
                    ' bfh20051104 - only show in cash out rpt now
                    '        If NewCash.Money < 0 Then
                    GoTo SkipLine
                    '        endif
                    '        PRESALE = PRESALE + NewCash.Money
                Case Else
                    GoTo SkipLine
            End Select

            XPCashIn = XPCashIn + NewCash.Money ' this is done for 41500
            PCashIn = PCashIn + NewCash.Money   ' this is not done for 41500

            If optDetail.Checked = True Then PrintLinesMisc()
SkipLine:
            CashJournalRecordSet.MoveNext()
        Loop

        CashIn = CashIn + PCashIn
        XCashIn = XCashIn + XPCashIn

        VISACHECK = PVISACHECK + VISACHECK
        '  FORFEIT = PFORFEIT + FORFEIT         'bfh20050412 this is recalculated in the forfeit report..  don't change the period total yet!
        BCPAY = PBCPAY + BCPAY
        FINANCE = PFINANCE + FINANCE
        MISCCASHIN = PMISCCASHIN + MISCCASHIN
        RESALE = PRESALE + RESALE

        If optDetail.Checked = True Then 'detail
            OutputObject.PrintNL
            'PrintTo(OutputObject, "Period To Date:", 0, AlignConstants.vbAlignLeft, False)
            PrintTo(OutputObject, "Period To Date:", 0, AlignConstants.vbAlignLeft, False, 1160)
            'PrintTo(OutputObject, CurrencyFormat(PCashIn), 32, AlignConstants.vbAlignRight, True)
            PrintTo(OutputObject, CurrencyFormat(PCashIn), 32, AlignConstants.vbAlignRight, True, 1160)
            Counter = Counter + 1
            PageCheck()

            'PrintTo(OutputObject, "Month To Date:", 0, AlignConstants.vbAlignLeft, False)
            PrintTo(OutputObject, "Month To Date:", 0, AlignConstants.vbAlignLeft, False, 1330)
            'PrintTo(OutputObject, CurrencyFormat(CashIn), 32, AlignConstants.vbAlignRight, True)
            PrintTo(OutputObject, CurrencyFormat(CashIn), 32, AlignConstants.vbAlignRight, True, 1330)
            Counter = Counter + 1
            PageCheck()
        End If

        TOTCASHIN = XCashIn
        PTOTCASHIN = XPCashIn
    End Sub

    Private Sub LoadCash(ByVal theDate As String, ByVal ToTheDate As String)
        Dim SS As String
        On Error GoTo HandleErr
        SS = ""
        SS = SS & "SELECT * FROM [" & CashJournal_TABLE & "] "
        SS = SS & "WHERE 1=1 "
        SS = SS & DateFilter
        SS = SS & CashierFilter
        SS = SS & "ORDER BY [TransDate], [CashID]"
        CashJournalRecordSet = GetRecordsetBySQL(SS, , GetDatabaseAtLocation())
        Exit Sub
HandleErr:
    End Sub

    Private Sub CashReport()
        Counter = 0

        ' Beginning cash is got from the data now.
        '  PBEGCASH = GetPrice(txtPriorPeriodCash)  ' ***  Where does this come from?  Textbox on screen.
        ' Text boxes are used only if the report starts on the first of a month.
        ' Jerry says enter $100 for each when running the report for real.
        ' Then he says to enter $100 no matter what.  I'm a little confused, but okay.

        If optDetail.Checked = True Then 'detail
            Index = 2
            Header = 2
            Headings()
            SubHeading()

            ' Print previous balances.
            'PrintTo(OutputObject, "Previous:", 0, AlignConstants.vbAlignLeft, False)
            PrintTo(OutputObject, "Previous:", 0, AlignConstants.vbAlignLeft, False, 980)
            'PrintTo(OutputObject, FormatCurrency(Cash), 33, AlignConstants.vbAlignRight, True)
            PrintTo(OutputObject, FormatCurrency(Cash), 33, AlignConstants.vbAlignRight, True, 980)
        End If

        If Not (CashJournalRecordSet.EOF And CashJournalRecordSet.BOF) Then CashJournalRecordSet.MoveFirst()

        Do Until CashJournalRecordSet.EOF
            CashJournalNew_RecordSet_Set(NewCash, CashJournalRecordSet)
            NewCash.Money = GetPrice(NewCash.Money)
            Typee = Val(NewCash.Account)

            If Typee = 90000 Then BEGCASH = NewCash.Money
            If NewCash.Money > 0 And (
      (Typee >= 1 And Typee <= 9) Or (Typee = 12) Or (Typee = 13)) Then
                '      Or _
                '      (Typee >= 550 And Typee <= 569))
                ' Account 15 (NSF) removed.
                PSalesDistribution()
                If optDetail.Checked = True Then PrintLinesCash()          ' This increments Counter and prints the record.
                Pcash = Pcash + NewCash.Money
                TotPCash = TotPCash + GetPrice(NewCash.Money)
            ElseIf (NewCash.Money < 0 Or Val(NewCash.Note) < 0) And ((Typee >= 1 And Typee <= 9) Or Typee = 12 Or Typee = 13) Then
                PSalesDistribution
                ConvertCode
                PRSALES = PRSALES + NewCash.Money
                Pcash = Pcash + NewCash.Money
                TotPCash = TotPCash + GetPrice(NewCash.Money)
                If optDetail.Checked = True Then PrintLinesCash()
            End If
            CashJournalRecordSet.MoveNext()
            Application.DoEvents()
        Loop

        TotCash = TotCash + TotPCash
        Sales = Cash + Pcash
        PSALES = Pcash

        Cash = Cash + Pcash
        '  CASHSALES = CASHSALES + PCASHSALES
        '  CHECK = CHECK + PCHECK
        '  VISA = VISA + PVISA
        '  MASTER = MASTER + PMASTER
        '  DISC = DISC + PDISC
        '  AMCX = AMCX + PAMCX
        '  FINC = FINC + PFINC
        '  NONE = NONE + PNONE
        '  STORECARD = STORECARD + PSTORECARD
        RSALES = RSALES + PRSALES

        Tax = 0
        STAX = 0

        If optDetail.Checked = True Then 'detail
            OutputObject.PrintNL
            'PrintTo(OutputObject, "Period To Date:", 0, AlignConstants.vbAlignLeft, False)
            PrintTo(OutputObject, "Period To Date:", 0, AlignConstants.vbAlignLeft, False, 1280)
            'PrintTo(OutputObject, FormatCurrency(Pcash), 33, AlignConstants.vbAlignRight, True)
            PrintTo(OutputObject, FormatCurrency(Pcash), 33, AlignConstants.vbAlignRight, True, 1280)

            'PrintTo(OutputObject, "Month To Date:", 0, AlignConstants.vbAlignLeft, False)
            PrintTo(OutputObject, "Month To Date:", 0, AlignConstants.vbAlignLeft, False, 1450)
            PrintTo(OutputObject, FormatCurrency(Cash), 33, AlignConstants.vbAlignRight, True, 1450)


            If OutputToPrinter Then
                If OutputObject.CurrentY <> 0 Then OutputObject.NewPage
            Else
                frmPrintPreviewDocument.NewPage()
            End If
        End If
    End Sub

    Private Sub AuditReport()
        Dim RS As ADODB.Recordset, SS As String

        On Error GoTo HandleErr

        Written = 0
        TaxCHARGED = 0
        CASHSALES = 0
        ARCASHSALES = 0
        CUSTDEP = 0
        UNDSALES = 0
        DELSALES = 0
        TAXREC = 0
        PWRITTEN = 0
        PTAXCHARGED = 0
        PARCASHSALES = 0
        PCUSTDEP = 0
        PUNDSALES = 0
        PDELSALES = 0
        PTAXREC = 0
        OTHER = 0
        POTHER = 0

        If optDetail.Checked = True Then 'detail
            Index = 1
            Header = 1
            Headings()
            SubHeading()
        End If

        ' Prior balance
        If DateAndTime.Day(theDate) <> 1 Then
            SS = ""
            SS = SS & "SELECT * FROM [" & SalesJournal_TABLE & "] "
            SS = SS & "WHERE 1=1 "
            SS = SS & DateFilter(True)
            SS = SS & CashierFilter
            SS = SS & "ORDER BY [TransDate], AuditID"
            RS = GetRecordsetBySQL(SS, , GetDatabaseAtLocation())
            Do Until RS.EOF
                SalesJournalNew_RecordSet_Set(NewAudit, RS)
                Written = Written + GetPrice(NewAudit.Written)
                TaxCHARGED = TaxCHARGED + GetPrice(NewAudit.TaxCharged1)
                ARCASHSALES = ARCASHSALES + GetPrice(NewAudit.ArCashSls)
                CUSTDEP = CUSTDEP + GetPrice(NewAudit.Control)
                UNDSALES = UNDSALES + GetPrice(NewAudit.UndSls)
                DELSALES = DELSALES + GetPrice(NewAudit.DelSls)
                TAXREC = TAXREC + GetPrice(NewAudit.TaxRec1)
                RS.MoveNext()
                Application.DoEvents()
            Loop
        End If
        If optDetail.Checked = True Then ' detail
            'PrintTo(OutputObject, "Previous Balance:", 0, AlignConstants.vbAlignLeft, False)
            PrintTo(OutputObject, "Previous Balance:", 0, AlignConstants.vbAlignLeft, False, 980)
            'PrintTo(OutputObject, CurrencyFormat(Written), 57, AlignConstants.vbAlignRight, False)
            PrintTo(OutputObject, CurrencyFormat(Written), 57, AlignConstants.vbAlignRight, False, 980)
            'PrintTo(OutputObject, CurrencyFormat(TaxCHARGED), 71, AlignConstants.vbAlignRight, False)
            PrintTo(OutputObject, CurrencyFormat(TaxCHARGED), 71, AlignConstants.vbAlignRight, False, 980)
            '    PrintTo OutputObject, CurrencyFormat(ARCASHSALES), 86, alignconstants.vbAlignRight, False ' was 14 wide
            'PrintTo(OutputObject, CurrencyFormat(CUSTDEP), 86, AlignConstants.vbAlignRight, False)
            PrintTo(OutputObject, CurrencyFormat(CUSTDEP), 86, AlignConstants.vbAlignRight, False, 980)
            'PrintTo(OutputObject, CurrencyFormat(UNDSALES), 99, AlignConstants.vbAlignRight, False)
            PrintTo(OutputObject, CurrencyFormat(UNDSALES), 99, AlignConstants.vbAlignRight, False, 980)
            'PrintTo(OutputObject, CurrencyFormat(DELSALES), 112, AlignConstants.vbAlignRight, False)
            PrintTo(OutputObject, CurrencyFormat(DELSALES), 112, AlignConstants.vbAlignRight, False, 980)
            'PrintTo(OutputObject, CurrencyFormat(TAXREC), 125, AlignConstants.vbAlignRight, True)
            PrintTo(OutputObject, CurrencyFormat(TAXREC), 125, AlignConstants.vbAlignRight, True, 980)
        End If

        SS = ""
        SS = SS & "SELECT * FROM [" & SalesJournal_TABLE & "] "
        SS = SS & "WHERE 1=1 "
        SS = SS & DateFilter()
        SS = SS & CashierFilter
        SS = SS & "ORDER BY [TransDate], SaleNo, AuditID"
        RS = GetRecordsetBySQL(SS, , GetDatabaseAtLocation())

        Dim Cy As Integer = 1160
        Do Until RS.EOF
            ' Period
            Application.DoEvents()
            SalesJournalNew_RecordSet_Set(NewAudit, RS)
            PWRITTEN = PWRITTEN + GetPrice(NewAudit.Written)
            PTAXCHARGED = PTAXCHARGED + GetPrice(NewAudit.TaxCharged1)
            PARCASHSALES = PARCASHSALES + GetPrice(NewAudit.ArCashSls)
            PCUSTDEP = PCUSTDEP + GetPrice(NewAudit.Control)
            PUNDSALES = PUNDSALES + GetPrice(NewAudit.UndSls)
            PDELSALES = PDELSALES + GetPrice(NewAudit.DelSls)
            PTAXREC = PTAXREC + GetPrice(NewAudit.TaxRec1)

            If optDetail.Checked = True Then 'detail
                OutputObject.FontSize = 8
                'PrintTo(OutputObject, Trim(NewAudit.SaleNo), 0, AlignConstants.vbAlignLeft, False)
                PrintTo(OutputObject, Trim(NewAudit.SaleNo), 0, AlignConstants.vbAlignLeft, False, Cy)
                'PrintTo(OutputObject, Trim(Microsoft.VisualBasic.Left(NewAudit.Name1, 17)), 12, AlignConstants.vbAlignLeft, False)
                PrintTo(OutputObject, Trim(Microsoft.VisualBasic.Left(NewAudit.Name1, 17)), 12, AlignConstants.vbAlignLeft, False, Cy)
                'PrintTo(OutputObject, Trim(NewAudit.TransDate), 33, AlignConstants.vbAlignLeft, False)
                PrintTo(OutputObject, Trim(NewAudit.TransDate), 33, AlignConstants.vbAlignLeft, False, Cy)
                'PrintTo(OutputObject, CurrencyFormat(NewAudit.Written), 57, AlignConstants.vbAlignRight, False)
                PrintTo(OutputObject, CurrencyFormat(NewAudit.Written), 57, AlignConstants.vbAlignRight, False, Cy)
                'PrintTo(OutputObject, CurrencyFormat(NewAudit.TaxCharged1), 71, AlignConstants.vbAlignRight, False)
                PrintTo(OutputObject, CurrencyFormat(NewAudit.TaxCharged1), 71, AlignConstants.vbAlignRight, False, Cy)
                '      PrintTo OutputObject, CurrencyFormat(NewAudit.ArCashSls), 86, alignconstants.vbAlignRight, False
                'PrintTo(OutputObject, CurrencyFormat(NewAudit.Control), 86, AlignConstants.vbAlignRight, False)
                PrintTo(OutputObject, CurrencyFormat(NewAudit.Control), 86, AlignConstants.vbAlignRight, False, Cy)
                'PrintTo(OutputObject, CurrencyFormat(NewAudit.UndSls), 99, AlignConstants.vbAlignRight, False)
                PrintTo(OutputObject, CurrencyFormat(NewAudit.UndSls), 99, AlignConstants.vbAlignRight, False, Cy)
                'PrintTo(OutputObject, CurrencyFormat(NewAudit.DelSls), 112, AlignConstants.vbAlignRight, False)
                PrintTo(OutputObject, CurrencyFormat(NewAudit.DelSls), 112, AlignConstants.vbAlignRight, False, Cy)
                'PrintTo(OutputObject, CurrencyFormat(NewAudit.TaxRec1), 125, AlignConstants.vbAlignRight, False)
                PrintTo(OutputObject, CurrencyFormat(NewAudit.TaxRec1), 125, AlignConstants.vbAlignRight, False, Cy)
                'PrintTo(OutputObject, IIf(Trim(NewAudit.Cashier) = "DEMO", "", Trim(NewAudit.Cashier)), 130, AlignConstants.vbAlignLeft, True) ' 139, alignconstants.vbAlignRight, True
                PrintTo(OutputObject, IIf(Trim(NewAudit.Cashier) = "DEMO", "", Trim(NewAudit.Cashier)), 130, AlignConstants.vbAlignLeft, True, Cy) ' 139, alignconstants.vbAlignRight, True
                Counter = Counter + 1
                Cy = Cy + 180
                If Counter >= 69 Then
                    If OutputToPrinter Then
                        If OutputObject.CurrentY <> 0 Then OutputObject.NewPage
                    Else
                        frmPrintPreviewDocument.NewPage()
                    End If
                    Counter = 0
                    Headings()
                    SubHeading()
                End If
            End If
            RS.MoveNext()
        Loop

        Written = Written + PWRITTEN
        TaxCHARGED = TaxCHARGED + PTAXCHARGED
        ARCASHSALES = ARCASHSALES + PARCASHSALES
        CUSTDEP = CUSTDEP + PCUSTDEP
        UNDSALES = UNDSALES + PUNDSALES
        DELSALES = DELSALES + PDELSALES
        TAXREC = TAXREC + PTAXREC

        Cy = Cy + 260
        If optDetail.Checked = True Then 'detail
            OutputObject.PrintNL
            'PrintTo(OutputObject, "Period To Date:", 0, AlignConstants.vbAlignLeft, False)
            PrintTo(OutputObject, "Period To Date:", 0, AlignConstants.vbAlignLeft, False, Cy)
            'PrintTo(OutputObject, CurrencyFormat(PWRITTEN), 57, AlignConstants.vbAlignRight, False)
            PrintTo(OutputObject, CurrencyFormat(PWRITTEN), 57, AlignConstants.vbAlignRight, False, Cy)
            'PrintTo(OutputObject, CurrencyFormat(PTAXCHARGED), 71, AlignConstants.vbAlignRight, False)
            PrintTo(OutputObject, CurrencyFormat(PTAXCHARGED), 71, AlignConstants.vbAlignRight, False, Cy)
            '    PrintTo OutputObject, CurrencyFormat(PARCASHSALES), 86, alignconstants.vbAlignRight, False
            'PrintTo(OutputObject, CurrencyFormat(PCUSTDEP), 86, AlignConstants.vbAlignRight, False)
            PrintTo(OutputObject, CurrencyFormat(PCUSTDEP), 86, AlignConstants.vbAlignRight, False, Cy)
            'PrintTo(OutputObject, CurrencyFormat(PUNDSALES), 99, AlignConstants.vbAlignRight, False)
            PrintTo(OutputObject, CurrencyFormat(PUNDSALES), 99, AlignConstants.vbAlignRight, False, Cy)
            'PrintTo(OutputObject, CurrencyFormat(PDELSALES), 112, AlignConstants.vbAlignRight, False)
            PrintTo(OutputObject, CurrencyFormat(PDELSALES), 112, AlignConstants.vbAlignRight, False, Cy)
            'PrintTo(OutputObject, CurrencyFormat(PTAXREC), 125, AlignConstants.vbAlignRight, True)
            PrintTo(OutputObject, CurrencyFormat(PTAXREC), 125, AlignConstants.vbAlignRight, True, Cy)

            Cy = Cy + 180
            'PrintTo(OutputObject, "Month to Date:", 0, AlignConstants.vbAlignLeft, False)
            PrintTo(OutputObject, "Month to Date:", 0, AlignConstants.vbAlignLeft, False, Cy)
            'PrintTo(OutputObject, CurrencyFormat(Written), 57, AlignConstants.vbAlignRight, False)
            PrintTo(OutputObject, CurrencyFormat(Written), 57, AlignConstants.vbAlignRight, False, Cy)
            'PrintTo(OutputObject, CurrencyFormat(TaxCHARGED), 71, AlignConstants.vbAlignRight, False)
            PrintTo(OutputObject, CurrencyFormat(TaxCHARGED), 71, AlignConstants.vbAlignRight, False, Cy)
            '    PrintTo OutputObject, CurrencyFormat(ARCASHSALES), 86, alignconstants.vbAlignRight, False
            'PrintTo(OutputObject, CurrencyFormat(CUSTDEP), 86, AlignConstants.vbAlignRight, False)
            PrintTo(OutputObject, CurrencyFormat(CUSTDEP), 86, AlignConstants.vbAlignRight, False, Cy)
            'PrintTo(OutputObject, CurrencyFormat(UNDSALES), 99, AlignConstants.vbAlignRight, False)
            PrintTo(OutputObject, CurrencyFormat(UNDSALES), 99, AlignConstants.vbAlignRight, False, Cy)
            'PrintTo(OutputObject, CurrencyFormat(DELSALES), 112, AlignConstants.vbAlignRight, False)
            PrintTo(OutputObject, CurrencyFormat(DELSALES), 112, AlignConstants.vbAlignRight, False, Cy)
            'PrintTo(OutputObject, CurrencyFormat(TAXREC), 125, AlignConstants.vbAlignRight, True)
            PrintTo(OutputObject, CurrencyFormat(TAXREC), 125, AlignConstants.vbAlignRight, True, Cy)

            If OutputToPrinter Then
                If OutputObject.CurrentY <> 0 Then
                    OutputObject.NewPage
                End If
            Else
                frmPrintPreviewDocument.NewPage()
            End If
        End If
        Exit Sub

HandleErr:
        Resume Next
    End Sub

    Private Sub LoadPriorCash(ByVal theDate As String)
        Dim SS As String
        On Error GoTo HandleErr
        SS = ""
        SS = SS & "SELECT * FROM [" & CashJournal_TABLE & "] "
        SS = SS & "WHERE 1=1 "
        SS = SS & DateFilter(True)
        SS = SS & CashierFilter
        SS = SS & "ORDER BY [TransDate], [CashID]"
        CashJournalRecordSet = GetRecordsetBySQL(SS, , GetDatabaseAtLocation())
        Exit Sub
HandleErr:
    End Sub

    Private Sub CalculatePreviousTotals()
        CashIn = 0
        XCashIn = 0
        LoadPriorCash(theDate)

        ' The cash/check/visa/etc subtotals are used in CashSummary (Cash Management report) later.
        ' We need two loops, for previous and period balances for each category.
        ' Can't do an aggregate query easily here.  Can we do all the previous balance loops together?
        If Not (CashJournalRecordSet.EOF And CashJournalRecordSet.BOF) Then CashJournalRecordSet.MoveFirst()

        Do Until CashJournalRecordSet.EOF
            CashJournalNew_RecordSet_Set(NewCash, CashJournalRecordSet)
            NewCash.Money = GetPrice(NewCash.Money)
            Typee = Val(NewCash.Account)

            ' CashReport
            If Typee = 90000 Then BEGCASH = NewCash.Money
            If NewCash.Money > 0 And (
      (Typee >= 1 And Typee <= 9) Or (Typee = 12) Or Typee = 13) Then
                SalesDistribution
                Cash = Cash + NewCash.Money
            End If
            If NewCash.Money < 0 Or Val(NewCash.Note) < 0 Then
                If (Typee >= 1 And Typee <= 9) Or Typee = 12 Or Typee = 13 Then 'both O/E & Install
                    SalesDistribution
                    Cash = Cash + NewCash.Money
                    RSALES = RSALES + NewCash.Money
                End If
            End If

            ' CashIn
            Select Case Typee
                Case 10700
                    VISACHECK = VISACHECK + NewCash.Money
                    CashIn = CashIn + NewCash.Money
                    XCashIn = XCashIn + NewCash.Money
                Case 41500
                    Forfeit = Forfeit + NewCash.Money  ' This goes into the Forfeit report with 12500.
                    XCashIn = XCashIn + NewCash.Money  ' xcashin is the one with 41500's in there for the 'MISC RECEIPTS TOTAL' line on the last page
        'CashIn = CashIn + NewCash.Money
                Case 61600
                    BCPAY = BCPAY + NewCash.Money
                    CashIn = CashIn + NewCash.Money
                    XCashIn = XCashIn + NewCash.Money
                Case 70000
                    FINANCE = FINANCE + NewCash.Money
                    CashIn = CashIn + NewCash.Money
                    XCashIn = XCashIn + NewCash.Money
                Case 99900
                    MISCCASHIN = MISCCASHIN + NewCash.Money
                    CashIn = CashIn + NewCash.Money
                    XCashIn = XCashIn + NewCash.Money
                Case 50200  '
                    If NewCash.Money >= 0 Then
                        RESALE = RESALE + NewCash.Money
                        CashIn = CashIn + NewCash.Money
                        XCashIn = XCashIn + NewCash.Money
                    End If
            End Select

            ' CashOutDwr
            Select Case Typee
      'Case 10000
        'PETTYCASH = PETTYCASH + NewCash.Money
        'Cashout = Cashout + NewCash.Money
                Case 50500
                    FREIGHTOUT = FREIGHTOUT + NewCash.Money
                    CashOut = CashOut + NewCash.Money
                Case 50600
                    Credit = Credit + NewCash.Money
                    CashOut = CashOut + NewCash.Money
                Case 60100
                    GAS = GAS + NewCash.Money
                    CashOut = CashOut + NewCash.Money
                Case 60500
                    VISADISC = VISADISC + NewCash.Money
                    CashOut = CashOut + NewCash.Money
                Case 62300
                    MAINTENANCE = MAINTENANCE + NewCash.Money
                    CashOut = CashOut + NewCash.Money
                Case 62400
                    REPAIR = REPAIR + NewCash.Money
                    CashOut = CashOut + NewCash.Money
                Case 63500
                    WHSE = WHSE + NewCash.Money
                    CashOut = CashOut + NewCash.Money
                Case 64100
                    OFFICE = OFFICE + NewCash.Money
                    CashOut = CashOut + NewCash.Money
                Case 65200
                    CASUAL = CASUAL + NewCash.Money
                    CashOut = CashOut + NewCash.Money
                Case 67500
                    TRAVEL = TRAVEL + NewCash.Money
                    CashOut = CashOut + NewCash.Money
                Case 99800
                    MiscCashOut = MiscCashOut + NewCash.Money
                    CashOut = CashOut + NewCash.Money
                Case 10000
                    MiscCashOut = MiscCashOut + NewCash.Money
                    CashOut = CashOut + NewCash.Money
                Case 50200
                    If NewCash.Money < 0 Then
                        PURCHASES = PURCHASES + NewCash.Money
                        CashOut = CashOut + NewCash.Money
                    End If
            End Select

            ' Banking
            If Typee >= 10200 And Typee <= 10650 Then
                Bank = Bank + NewCash.Money
                BankType
            End If

            ' check refund
            If Typee = 21500 Then
                CheckRefund = CheckRefund + NewCash.Money
                '      FORFEIT = FORFEIT + NewCash.Money  ' Forfeit is 41500.
            End If

            ' RecControl
            If Typee = 11200 Then
                RECPAYMENTS = RECPAYMENTS + NewCash.Money
            End If

            ' AR/Installment Payments
            If (Typee >= 550 And Typee <= 559) Or Typee = 15 Then
                SalesDistribution
                AR = AR + NewCash.Money
            End If

            ' ARLate/Interest
            If Typee >= 560 And Typee <= 569 Then
                SalesDistribution
                INTEREST = INTEREST + NewCash.Money
            End If

            TotCash = TotCash + NewCash.Money

            CashJournalRecordSet.MoveNext()
            Application.DoEvents()
        Loop

        TotCash = Cash
    End Sub

    Private Sub BankType()
        'BANK TOTALS
        Select Case Val(NewCash.Account)
            Case 10200
                BCASHSALES = BCASHSALES + NewCash.Money
            Case 10250
                BECHECKS = BECHECKS + NewCash.Money
            Case 10300
                BVISA = BVISA + NewCash.Money
            Case 10400
                BDisc = BDisc + NewCash.Money
            Case 10500
                BAmcx = BAmcx + NewCash.Money
            Case 10600
                Debit = Debit + NewCash.Money
            Case 10650
                BSTORECARD = BSTORECARD + NewCash.Money
        End Select
    End Sub

    Private Sub cmdPrintPreview_Click(sender As Object, e As EventArgs) Handles cmdPrintPreview.Click
        Dim DBG As String
        On Error GoTo ErrorHandler

        DBG = "a"

        If Not StoreSettings.bManualBillofSaleNo Then
            If Not frmEditSalesJournal.OutOfDateSalesReport(StoresSld, dteStartDate.Value, dteEndDate.Value) Then Exit Sub
            DBG = "aa"
            If Not frmEditCash.OutOfDateCashReport(StoresSld, dteStartDate.Value, dteEndDate.Value) Then Exit Sub
        End If

        DBG = "b"
        SetWorking(True)

        OutputObject = New cPrinter
        '<CT>
        'OutputObject.SetPreview("Daily Audit Report", "Daily Audit,Audit,Daily Audit Report", Me)
        OutputObject.SetPrintToPDF("Daily Audit Report", "Daily Audit,Audit,Daily Audit Report")
        '</CT>

        DBG = "cc"
        PrintSub()

        DBG = "dd"

        DBG = "e"
        SetWorking(False)
        DBG = "f"

        Exit Sub

ErrorHandler:
        Select Case Err.Number
            Case 482 : ErrNoPrinter()
                Exit Sub
            Case Else
                MessageBox.Show("Error in previewing audit report (" & Err.Number & "):" & vbCrLf & Err.Description & vbCrLf & "Source: " & Err.Source & vbCrLf & "DBG=" & DBG)
        End Select
        Resume Next
    End Sub

    Private Sub OrdAudit2_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        SetButtonImage(cmdPrint, 19) ' "print")
        SetButtonImage(cmdPrintPreview, 20) ' "preview")
        SetButtonImage(cmdCancel, 3) ' "cancel")
        theDate = CurrentMonthStart()
        dteStartDate.Value = Date.ParseExact(theDate, "MM-dd-yyyy", Globalization.CultureInfo.InvariantCulture)

        ToTheDate = Today
        dteEndDate.Value = ToTheDate

        LoadCashiers
    End Sub

    Private Sub LoadCashiers()
        Dim SS As String, RS As ADODB.Recordset
        cmbCashier.Items.Clear()
        cmbCashier.Items.Add(EntireStore)
        cmbCashier.Text = EntireStore

        If False Then ' Not IsDevelopment Then
            lblCashier.Visible = False
            cmbCashier.Visible = False
            Exit Sub
        End If

        lblCashier.Visible = True
        cmbCashier.Visible = True

        SS = ""
        SS = SS & "SELECT DISTINCT [Cashier] FROM [Audit] "
        SS = SS & "WHERE 1=1 "
        SS = SS & DateFilter
        SS = SS & "ORDER BY [Cashier]"
        RS = GetRecordsetBySQL(SS, , GetDatabaseAtLocation)
        Do While Not RS.EOF
            cmbCashier.Items.Add(IfNullThenNilString(RS("Cashier").Value))
            RS.MoveNext
        Loop
    End Sub

    Private Sub txtCashInDrawer_Enter(sender As Object, e As EventArgs) Handles txtCashInDrawer.Enter
        SelectContents(txtCashInDrawer)
    End Sub

    Private Sub txtCashInDrawer_Leave(sender As Object, e As EventArgs) Handles txtCashInDrawer.Leave
        'txtCashInDrawer.Text = Format(txtCashInDrawer.Text, "$###,##0.00")
        Try
            txtCashInDrawer.Text = Format(Convert.ToDecimal(txtCashInDrawer.Text), "$###,##0.00")
        Catch ex As FormatException
        End Try
    End Sub

    Private Sub txtPriorPeriodCash_Enter(sender As Object, e As EventArgs) Handles txtPriorPeriodCash.Enter
        SelectContents(txtPriorPeriodCash)
    End Sub

    Private Sub txtPriorPeriodCash_Leave(sender As Object, e As EventArgs) Handles txtPriorPeriodCash.Leave
        'txtPriorPeriodCash.Text = Format(txtPriorPeriodCash.Text, "$###,##0.00")
        Try
            txtPriorPeriodCash.Text = Format(Convert.ToDecimal(txtPriorPeriodCash.Text), "$###,##0.00")
        Catch ex As Exception
        End Try
    End Sub

    Private Sub dteStartDate_CloseUp(sender As Object, e As EventArgs) Handles dteStartDate.CloseUp
        theDate = dteStartDate.Value
        LoadCashiers()
    End Sub

    Private Sub dteEndDate_CloseUp(sender As Object, e As EventArgs) Handles dteEndDate.CloseUp
        ToTheDate = dteEndDate.Value
        LoadCashiers()
    End Sub

    Private Sub OrdAudit2_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        'If UnloadMode = vbFormControlMenu Then cmdCancel.Value = True
        cmdCancel_Click(cmdCancel, New EventArgs)
    End Sub
End Class

