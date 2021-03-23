Imports VBRUN
Public Class ReportPrint
    Dim Curr As Boolean
    Dim Days30 As Boolean
    Dim Days60 As Boolean
    Dim Days90 As Boolean

    Dim Nosale As Decimal, BNoSale As Decimal
    Dim Curr30 As Decimal, BCurr30 As Decimal
    Dim Sales30 As Decimal, BSales30 As Decimal
    Dim Sales60 As Decimal, BSales60 As Decimal
    Dim Sales90 As Decimal, BSales90 As Decimal
    'Dim Sales120 As decimal

    Dim SI As StoreInfo
    Dim StoreName As String
    Dim StoreAddress As String
    Dim StoreCity As String
    Dim StorePhone As String
    Dim StoreShipTo As String
    Dim StoreShipAdd As String
    Dim StoreShipCity As String
    Dim StoreShipPhone As String
    Dim PrintedCost As Decimal
    Dim TotCost As Decimal
    Dim Cost As Decimal
    Dim TotGross As Decimal, SubGross As Decimal
    Dim TotDep As Decimal, SubDep As Decimal
    Dim TotBalance As Decimal, SubBalance As Decimal
    Dim Counter As Long

    Private Sub ReportPrint_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        SetButtonImage(cmdCancel, 3)
        SetButtonImage(cmdPrint, 19)
        SetButtonImage(cmdPrintPreview, 20)
        ColorDatePicker(dteReportDate)
        dteReportDate.Value = Today

        chkLastPay.Visible = OrderMode("R") And ReportsMode("O")  ' Undelivered Sales Report

        If OrderMode("R", "L", "C", "B") Then
            Opt1.Visible = True
            Opt2.Visible = True
            'Opt2.Value = True
            Opt2.Checked = True
            Opt3.Visible = True
            Opt5.Visible = True
            fraOptions.Visible = True
            Opt1.Text = "Sale No"
            Opt2.Text = "Name"
            Opt3.Text = "Ageing"
        Else
            Opt1.Visible = False
            Opt2.Visible = False
            Opt3.Visible = False
            Opt5.Visible = False
            fraOptions.Visible = False
        End If

        If InvenMode("P") Then
            Text = "Print Out P/Os"
            lblSelectDate.Text = "Po Date"
        End If

        Select Case Order
            Case "R"
                Text = "Undelivered Sales"
                lblSelectDate.Text = "Report Date:"
                Opt2.Checked = True
                'HelpContextID = 49630
            Case "L"
                lblSelectDate.Text = "This report will use the current date."
                dteReportDate.Visible = False
                Text = "Lay-a-Way Report"
                'HelpContextID = 49640
            Case "B"
                lblSelectDate.Text = "This report will use the current date."
                dteReportDate.Visible = False
                Text = "Back Order Report"
                'HelpContextID = 49650
            Case "C"
                lblSelectDate.Text = "This report will use the current date."
                dteReportDate.Visible = False
                Text = "Credit Report"
                'HelpContextID = 49660
        End Select
    End Sub

    Private Sub cmdCancel_Click(sender As Object, e As EventArgs) Handles cmdCancel.Click
        If InvenMode("P") Or OrderMode("R", "L", "C", "B") Then
            'Unload Me
            Me.Close()
            MainMenu.Show()
            modProgramState.Reports = ""
            modProgramState.Order = ""
        End If
    End Sub

    Private Sub cmdPrint_Click(sender As Object, e As EventArgs) Handles cmdPrint.Click
        'MousePointer = vbHourglass
        Me.Cursor = Cursors.WaitCursor
        cmdPrint.Enabled = False
        cmdPrintPreview.Enabled = False
        cmdCancel.Enabled = False
        OutputObject = Printer
        OutputToPrinter = True

        If InvenMode("P") Then
            SelectPo()
        End If
        If OrderMode("R", "L", "C", "B") Then
            'End Of Month
            Counter = 0
            Undelivered()
        End If

        OutputObject.EndDoc
        cmdPrint.Enabled = True
        cmdPrintPreview.Enabled = True
        cmdCancel.Enabled = True
        'MousePointer = vbDefault
        Me.Cursor = Cursors.Default
        Exit Sub
HandleErr:
        MessageBox.Show("There are no P/Os to Print", "WinCDS", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        'Unload Me
        Me.Close()
        modProgramState.Order = ""
        MainMenu.Show()
        Exit Sub
    End Sub

    Private Sub cmdPrintPreview_Click(sender As Object, e As EventArgs) Handles cmdPrintPreview.Click
        'MousePointer = vbHourglass
        Me.Cursor = Cursors.WaitCursor
        cmdPrint.Enabled = False
        cmdPrintPreview.Enabled = False
        cmdCancel.Enabled = False

        'Load frmPrintPreviewMain
        OutputToPrinter = False
        OutputObject = frmPrintPreviewDocument.picPicture

        frmPrintPreviewDocument.CallingForm = Me
        frmPrintPreviewDocument.ReportName = Text
        If InvenMode("P") Then
            SelectPo()
        End If
        If OrderMode("R", "L", "C", "B") Then
            'End Of Month
            Counter = 0
            Undelivered()
        End If
        Hide()
        'MousePointer = vbDefault
        Me.Cursor = Cursors.Default
        frmPrintPreviewDocument.DataEnd()

        cmdPrint.Enabled = True
        cmdPrintPreview.Enabled = True
        cmdCancel.Enabled = True
        Exit Sub
HandleErr:
        MessageBox.Show("There are no P/Os to Preview", "WinCDS", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        'Unload Me
        Me.Close()
        modProgramState.Order = ""
        MainMenu.Show()
        Exit Sub
    End Sub

    Private Sub Undelivered()
        Dim OldSalesStaff As String

        On Error GoTo HandleErr

        Curr = False
        Days30 = False
        Days60 = False
        Days90 = False
        Nosale = 0 : BNoSale = 0
        Curr30 = 0 : BCurr30 = 0
        Sales30 = 0 : BSales30 = 0
        Sales60 = 0 : BSales60 = 0
        Sales90 = 0 : BSales90 = 0
        SubGross = 0 : SubDep = 0 : SubBalance = 0

        Counter = 0

        Dim RS As ADODB.Recordset, SQL As String, Eom As EomFile, EOMRecord As Long
        SQL = ""
        SQL = SQL & "SELECT Holding.Sale, Holding.Deposit, Holding.LeaseNo, Holding.Status, "
        SQL = SQL & "Holding.LastPay as DBLastPay, trim(Holding.Salesman) as FirstSalesman, Mail.Last, Mail.First, "
        If chkLastPay.Checked = True Then
            SQL = SQL & "(Select Max(SellDate) FROM GrossMargin WHERE SaleNo=Holding.LeaseNo AND Style='PAYMENT') as LastPay,  "
        Else
            SQL = SQL & "'' as LastPay,  "
        End If
        'BFH20150302 - Last Pay removed
        SQL = SQL & "(Select Min(SellDate) FROM GrossMargin WHERE SaleNo=Holding.LeaseNo) as FirstSaleDate "
        SQL = SQL & "FROM Holding LEFT JOIN mail ON Mail.Index=Holding.Index "

        Select Case Order
            Case "R"
                SQL = SQL & "WHERE Holding.Status IN ('O', 'L', '1', '2', '3', '4', 'S') "
            Case "L"
                SQL = SQL & "WHERE Holding.Status IN ('L', '1', '2', '3', '4') "
            Case "C"
                SQL = SQL & "WHERE Holding.Status IN ('E', 'C') "
            Case "B"
                SQL = SQL & "WHERE Holding.Status IN ('B') "
            Case Else
                MessageBox.Show("Invalid report type.", "WinCDS", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                Exit Sub
        End Select
        If Opt1.Checked = True Then
            SQL = SQL & "ORDER BY Holding.LeaseNo"
        ElseIf Opt2.Checked = True Then
            SQL = SQL & "ORDER BY Mail.Last, Mail.First"
        ElseIf Opt3.Checked = True Then
            SQL = "SELECT * FROM (" & SQL & ") as InnerQuery ORDER BY LastPay"
        ElseIf Opt5.Checked = True Then
            SQL = "SELECT * FROM (" & SQL & ") as InnerQuery ORDER BY FirstSaleDate"
        Else
            SQL = SQL & "ORDER BY trim(Holding.Salesman)"
        End If
        ProgressForm(0, 1, "Loading records (may take a while)...")
        RS = GetRecordsetBySQL(SQL, , GetDatabaseAtLocation())


        UndeliveredHeading()
        OutputObject.FontSize = 9

        ProgressForm(0, RS.RecordCount, "Printing...")

        '<CT>
        Dim Cy As Integer = 900
        '</CT>
        Do Until RS.EOF
            'ProgressForm(RS.AbsolutePosition)
            GetEOM(Eom, RS)
            EOMRecord = EOMRecord + 1

            If Opt3.Checked = True Then Ageing(Eom)

            'If IsDevelopment And RS("LeaseNo") = "47391" Then Stop

            SubGross = SubGross + Eom.GrossSale
            SubDep = SubDep + Eom.TotDeposit
            SubBalance = SubBalance + Eom.Balance

            If Opt4.Checked = True Then
                If Trim(OldSalesStaff) <> Trim(Eom.Salesman) And EOMRecord > 1 Then
                    PrintTotals()
                    If OutputToPrinter Then OutputObject.NewPage Else frmPrintPreviewDocument.NewPage()
                    UndeliveredHeading()
                    Counter = 0
                End If
            End If

            ' New formatting, allowing right-and-center-aligned fields.
            'PrintTo(OutputObject, Microsoft.VisualBasic.Left(Eom.LastName, 18), 0, AlignConstants.vbAlignLeft, False)
            PrintTo(OutputObject, Microsoft.VisualBasic.Left(Eom.LastName, 18), 0, AlignConstants.vbAlignLeft, False, Cy)
            'PrintTo(OutputObject, Trim(Eom.Status), 38, AlignConstants.vbAlignLeft, False)
            PrintTo(OutputObject, Trim(Eom.Status), 38, AlignConstants.vbAlignLeft, False, Cy)
            'PrintTo(OutputObject, Trim(Eom.LeaseNo), 48, AlignConstants.vbAlignLeft, False)
            PrintTo(OutputObject, Trim(Eom.LeaseNo), 48, AlignConstants.vbAlignLeft, False, Cy)
            'PrintTo(OutputObject, Format(Eom.GrossSale, "$###,##0.00"), 74, AlignConstants.vbAlignRight, False)
            PrintTo(OutputObject, Format(Convert.ToDecimal(Eom.GrossSale), "$###,##0.00"), 74, AlignConstants.vbAlignRight, False, Cy)
            'PrintTo(OutputObject, Format(Eom.TotDeposit, "$###,##0.00"), 90, AlignConstants.vbAlignRight, False)
            PrintTo(OutputObject, Format(Convert.ToDecimal(Eom.TotDeposit), "$###,##0.00"), 90, AlignConstants.vbAlignRight, False, Cy)
            'PrintTo(OutputObject, Format(Eom.Balance, "$###,##0.00"), 109, AlignConstants.vbAlignRight, False)
            PrintTo(OutputObject, Format(Convert.ToDecimal(Eom.Balance), "$###,##0.00"), 109, AlignConstants.vbAlignRight, False, Cy)
            'PrintTo(OutputObject, Trim(Eom.LastPay), 124, AlignConstants.vbAlignRight, False)
            PrintTo(OutputObject, Trim(Eom.LastPay), 124, AlignConstants.vbAlignRight, False, Cy)
            'PrintTo(OutputObject, Trim(Eom.Salesman), 128, AlignConstants.vbAlignLeft, True)
            PrintTo(OutputObject, Trim(Eom.Salesman), 128, AlignConstants.vbAlignLeft, True, Cy)

            TotGross = TotGross + Eom.GrossSale
            TotDep = TotDep + Eom.TotDeposit
            TotBalance = TotBalance + Eom.Balance
            OldSalesStaff = Eom.Salesman

            Counter = Counter + 1
            '<CT>
            Cy = Cy + 200
            '</CT>
            If Counter >= 66 Then
                If OutputToPrinter Then OutputObject.NewPage Else frmPrintPreviewDocument.NewPage()
                UndeliveredHeading()
                Counter = 0
                Cy = 900
            End If
            RS.MoveNext()
        Loop
        
        ProgressForm()

        If Opt3.Checked = True Then Ageing(Eom, True)

        PrintTotals()
        Exit Sub

HandleErr:
        Select Case Err()
            Case Else
                MessageBox.Show("Error in ReportPrint.Undelivered: " & Err.Description, "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Resume Next
        End Select
    End Sub

    Private Sub PrintTotals()
        OutputObject.Print
        On Error Resume Next
        OutputObject.Line(4800, OutputObject.CurrentY - 70, 8800, OutputObject.CurrentY - 70)
        OutputObject.Line(4800, OutputObject.CurrentY + 20, 8800, OutputObject.CurrentY + 20)
        On Error GoTo 0

        OutputObject.FontBold = True
        PrintTo(OutputObject, Format(TotGross, "$###,##0.00"), 74, AlignConstants.vbAlignRight, False)
        PrintTo(OutputObject, Format(TotDep, "$###,##0.00"), 90, AlignConstants.vbAlignRight, False)
        PrintTo(OutputObject, Format(TotGross - TotDep, "$###,##0.00"), 109, AlignConstants.vbAlignRight, False)
        OutputObject.FontBold = False

        If Opt3.Checked = True Then   'Last Pay
            OutputObject.Print
            OutputObject.Print
            OutputObject.FontSize = 16
            OutputObject.Print(TAB(25), "Receivables Ageing Report")
            OutputObject.FontSize = 12
            OutputObject.FontBold = True
            PrintToTab(OutputObject, "No Deposit", 20)
            PrintToTab(OutputObject, "Current", 40)
            PrintToTab(OutputObject, "Over 30", 60)
            PrintToTab(OutputObject, "Over 60", 80)
            PrintToTab(OutputObject, "Over 90", 100, , True)
            OutputObject.FontBold = False

            OutputObject.FontBold = True
            PrintToTab(OutputObject, "Gross Sale")
            OutputObject.FontBold = False
            PrintToTab(OutputObject, FormatCurrency(Nosale), 20)
            PrintToTab(OutputObject, FormatCurrency(Curr30), 40)
            PrintToTab(OutputObject, FormatCurrency(Sales30), 60)
            PrintToTab(OutputObject, FormatCurrency(Sales60), 80)
            PrintToTab(OutputObject, FormatCurrency(Sales90), 100, , True)

            OutputObject.FontBold = True
            PrintToTab(OutputObject, "Balance")
            OutputObject.FontBold = False
            PrintToTab(OutputObject, FormatCurrency(BNoSale), 20)
            PrintToTab(OutputObject, FormatCurrency(BCurr30), 40)
            PrintToTab(OutputObject, FormatCurrency(BSales30), 60)
            PrintToTab(OutputObject, FormatCurrency(BSales60), 80)
            PrintToTab(OutputObject, FormatCurrency(BSales90), 100, , True)
        End If

        TotGross = 0
        TotDep = 0
        Curr = False
        Days30 = False
        Days60 = False
        Days90 = False
        Nosale = 0 : BNoSale = 0
        Curr30 = 0 : BCurr30 = 0
        Sales30 = 0 : BSales30 = 0
        Sales60 = 0 : BSales60 = 0
        Sales90 = 0 : BSales90 = 0
        SubGross = 0
        SubDep = 0
        SubBalance = 0
    End Sub

    Private Sub Ageing(ByRef Eom As EomFile, Optional ByVal LastTotal As Boolean = False)
        Dim TodayDate As Date
        TodayDate = dteReportDate.Value

        If Not LastTotal And Val(Eom.LastPay) <> 0 Then 'For wrong date format or no date

            If DateDiff("d", Eom.LastPay, TodayDate) >= 91 Then
                Sales90 = Sales90 + Eom.GrossSale
                BSales90 = BSales90 + Eom.Balance
                If Not Days90 Then
                    OutputObject.FontBold = True

                    If SubGross <> 0 Or SubDep <> 0 Or SubBalance <> 0 Then
                        On Error Resume Next
                        OutputObject.Line(4800, OutputObject.CurrentY, 8800, OutputObject.CurrentY)
                        On Error GoTo 0
                        PrintTo(OutputObject, Format(SubGross, "$###,##0.00"), 74, AlignConstants.vbAlignRight, False)
                        PrintTo(OutputObject, Format(SubDep, "$###,##0.00"), 90, AlignConstants.vbAlignRight, False)
                        PrintTo(OutputObject, Format(SubBalance, "$###,##0.00"), 109, AlignConstants.vbAlignRight, True)
                    End If
                    SubGross = 0 : SubDep = 0 : SubBalance = 0

                    OutputObject.Print
                    OutputObject.Print(TAB(25), New String("*", 30), "   Over 90 Days   ", New String("*", 30))
                    OutputObject.Print
                    OutputObject.FontBold = False
                    Days90 = True
                    Counter = Counter + 3
                    Exit Sub
                End If
            End If

            If DateDiff("d", Eom.LastPay, TodayDate) >= 61 And DateDiff("d", Eom.LastPay, TodayDate) <= 90 Then
                Sales60 = Sales60 + Eom.GrossSale
                BSales60 = BSales60 + Eom.Balance
                If Not Days60 Then
                    OutputObject.FontBold = True

                    If SubGross <> 0 Or SubDep <> 0 Or SubBalance <> 0 Then
                        On Error Resume Next
                        OutputObject.Line(4800, OutputObject.CurrentY, 8800, OutputObject.CurrentY)
                        On Error GoTo 0
                        PrintTo(OutputObject, Format(SubGross, "$###,##0.00"), 74, AlignConstants.vbAlignRight, False)
                        PrintTo(OutputObject, Format(SubDep, "$###,##0.00"), 90, AlignConstants.vbAlignRight, False)
                        PrintTo(OutputObject, Format(SubBalance, "$###,##0.00"), 109, AlignConstants.vbAlignRight, True)
                    End If
                    SubGross = 0 : SubDep = 0 : SubBalance = 0

                    OutputObject.Print
                    OutputObject.Print(TAB(25), New String("*", 30), "   Over 60 Days   ", New String("*", 30))
                    OutputObject.Print
                    OutputObject.FontBold = False
                    Days60 = True
                    Counter = Counter + 3
                    Exit Sub
                End If
            End If

            If DateDiff("d", Eom.LastPay, TodayDate) >= 31 And DateDiff("d", Eom.LastPay, TodayDate) <= 60 Then
                Sales30 = Sales30 + Eom.GrossSale
                BSales30 = BSales30 + Eom.Balance
                If Not Days30 Then
                    OutputObject.FontBold = True

                    If SubGross <> 0 Or SubDep <> 0 Or SubBalance <> 0 Then
                        On Error Resume Next
                        OutputObject.Line(4800, OutputObject.CurrentY, 8800, OutputObject.CurrentY)
                        On Error GoTo 0
                        PrintTo(OutputObject, Format(SubGross, "$###,##0.00"), 74, AlignConstants.vbAlignRight, False)
                        PrintTo(OutputObject, Format(SubDep, "$###,##0.00"), 90, AlignConstants.vbAlignRight, False)
                        PrintTo(OutputObject, Format(SubBalance, "$###,##0.00"), 109, AlignConstants.vbAlignRight, True)
                    End If
                    SubGross = 0 : SubDep = 0 : SubBalance = 0

                    OutputObject.Print
                    OutputObject.Print(TAB(25), New String("*", 30), "   Over 30 Days   ", New String("*", 30))
                    OutputObject.Print
                    OutputObject.FontBold = False
                    Days30 = True
                    Counter = Counter + 3
                    Exit Sub
                End If
            End If

            If DateDiff("d", Eom.LastPay, TodayDate) >= 0 And DateDiff("d", Eom.LastPay, TodayDate) <= 30 Then
                Curr30 = Curr30 + Eom.GrossSale
                BCurr30 = BCurr30 + Eom.Balance
                If Not Curr Then
                    OutputObject.FontBold = True

                    If SubGross <> 0 Or SubDep <> 0 Or SubBalance <> 0 Then
                        On Error Resume Next
                        OutputObject.Line(4800, OutputObject.CurrentY, 8800, OutputObject.CurrentY)
                        On Error GoTo 0
                        PrintTo(OutputObject, Format(SubGross, "$###,##0.00"), 74, AlignConstants.vbAlignRight, False)
                        PrintTo(OutputObject, Format(SubDep, "$###,##0.00"), 90, AlignConstants.vbAlignRight, False)
                        PrintTo(OutputObject, Format(SubBalance, "$###,##0.00"), 109, AlignConstants.vbAlignRight, True)
                    End If
                    SubGross = 0 : SubDep = 0 : SubBalance = 0

                    OutputObject.Print
                    OutputObject.Print(TAB(25), New String("*", 30), "    Current    ", New String("*", 30))
                    OutputObject.FontBold = False
                    OutputObject.Print
                    Curr = True
                    Counter = Counter + 3
                    Exit Sub
                End If
            End If
        End If

        If LastTotal Then
            OutputObject.FontBold = True

            If SubGross <> 0 Or SubDep <> 0 Or SubBalance <> 0 Then
                On Error Resume Next
                OutputObject.Line(4800, OutputObject.CurrentY, 8800, OutputObject.CurrentY)
                On Error GoTo 0
                PrintTo(OutputObject, Format(SubGross, "$###,##0.00"), 74, AlignConstants.vbAlignRight, False)
                PrintTo(OutputObject, Format(SubDep, "$###,##0.00"), 90, AlignConstants.vbAlignRight, False)
                PrintTo(OutputObject, Format(SubBalance, "$###,##0.00"), 109, AlignConstants.vbAlignRight, True)
            End If
            SubGross = 0 : SubDep = 0 : SubBalance = 0
            OutputObject.FontBold = False
        End If

        If Trim(Eom.LastPay) = "" Then Nosale = Nosale + Eom.GrossSale : BNoSale = BNoSale + Eom.Balance
    End Sub

    Private Sub GetEOM(ByRef Eom As EomFile, ByRef RS As ADODB.Recordset)
        Eom.Balance = RS("Sale").Value - RS("Deposit").Value
        Eom.GrossSale = RS("Sale").Value
        If IsNothing(RS("Last").Value) Then
            Eom.LastName = "Cash & Carry"
        Else
            Eom.LastName = RS("Last").Value & IIf(IsNothing(RS("First").Value), "", ", " & RS("First").Value)
        End If
        Eom.LastPay = IIf(IsNothing(RS("LastPay").Value), "", RS("LastPay").Value)
        Eom.LeaseNo = RS("LeaseNo").Value
        Eom.Salesman = TranslateSalesmen(IfNullThenNilString(RS("FirstSalesman").Value))
        Eom.Status = RS("Status").Value
        Eom.TotDeposit = RS("Deposit").Value
    End Sub

    Public Sub UndeliveredHeading()
        OutputObject.FontName = "Arial"
        OutputObject.FontSize = 18
        OutputObject.CurrentX = 0
        OutputObject.CurrentY = 100
        OutputObject.FontBold = True

        Select Case Order
            Case "B"
                PrintCentered("Back Order/Receivables Report")
            Case "C"
                PrintCentered("Credit Sales/Receivables Report")
            Case "L"
                PrintCentered("Lay-A-Way Sales Report")
            Case "R"
                PrintCentered("Undelivered/Receivables Report")
        End Select

        If ReportsMode("ML") Then PrintCentered("Master List of Manufacturers")

        OutputObject.FontBold = False

        OutputObject.FontSize = 8
        OutputObject.CurrentX = 10
        OutputObject.CurrentY = 100
        OutputObject.Print("Date: ", DateFormat(dteReportDate.Value))
        OutputObject.Print("Time: ", Format(Now, "h:mm:ss am/pm"))

        OutputObject.CurrentX = 10100
        OutputObject.CurrentY = 100
        If OutputToPrinter Then
            OutputObject.Print("Page:" & OutputObject.Page)
        Else
            OutputObject.Print("Page:" & PageNumber)
        End If

        OutputObject.CurrentY = 500
        PrintCentered(StoreSettings.Name & "    " & StoreSettings.Address & "    " & StoreSettings.City)

        OutputObject.CurrentX = 0
        OutputObject.CurrentY = 700
        OutputObject.FontSize = 11
        OutputObject.FontBold = True

        If OrderMode("R", "B", "L", "C") Then
            OutputObject.FontSize = 9

            'PrintTo(OutputObject, "Last Name", 0, AlignConstants.vbAlignLeft, False)
            PrintTo(OutputObject, "Last Name", 0, AlignConstants.vbAlignLeft, False, 700)
            'PrintTo(OutputObject, "Status", 38, AlignConstants.vbAlignLeft, False)
            PrintTo(OutputObject, "Status", 38, AlignConstants.vbAlignLeft, False, 700)
            'PrintTo(OutputObject, "Sale No.", 48, AlignConstants.vbAlignLeft, False)
            PrintTo(OutputObject, "Sale No.", 48, AlignConstants.vbAlignLeft, False, 700)
            'PrintTo(OutputObject, "Gross Sale", 74, AlignConstants.vbAlignRight, False)
            PrintTo(OutputObject, "Gross Sale", 74, AlignConstants.vbAlignRight, False, 700)
            'PrintTo(OutputObject, "Total Deposit", 90, AlignConstants.vbAlignRight, False)
            PrintTo(OutputObject, "Total Deposit", 90, AlignConstants.vbAlignRight, False, 700)
            'PrintTo(OutputObject, "Balance", 109, AlignConstants.vbAlignRight, False)
            PrintTo(OutputObject, "Balance", 109, AlignConstants.vbAlignRight, False, 700)
            'If chkLastPay.Checked = True Then PrintTo(OutputObject, "Last Pay", 124, AlignConstants.vbAlignRight, False)
            If chkLastPay.Checked = True Then PrintTo(OutputObject, "Last Pay", 124, AlignConstants.vbAlignRight, False, 700)
            'PrintTo(OutputObject, "Salesmen", 128, AlignConstants.vbAlignLeft, True)
            PrintTo(OutputObject, "Salesmen", 128, AlignConstants.vbAlignLeft, True, 700)

        End If

        If Reports = "ML" Then
            OutputObject.CurrentX = 9
            OutputObject.CurrentY = 700
            OutputObject.Print("  Code  FT%  O/S  List  Vendor               Code  FT%  O/S  List  Vendor              CODE  FT%  O/S  List  Vendor")

            OutputObject.Line(3900, 1000, 3900, 15000)
            OutputObject.Line(7900, 1000, 7900, 15000)
            OutputObject.FontSize = 9
            OutputObject.CurrentX = 0
            OutputObject.CurrentY = 1000
        End If

        OutputObject.FontBold = False
    End Sub

    Private Sub SelectPo()
        Domain_init
        Dim SQL As String
        SQL = "SELECT * FROM [" & PODetail_TABLE & "] WHERE PrintPO NOT IN ('X', 'V') ORDER BY [" &
    PODetail_INDEX & "]"

        Dim PO As cPODetail
        PO = New cPODetail
        PO.DataAccess.Records_OpenSQL(SQL)
        Do While PO.DataAccess.Records_Available
            PrintPo(PO)
        Loop
        DisposeDA(PO)

        Domain_exit()
    End Sub

    Private Sub LineItems(ByRef PO As cPODetail)
        If PO.wCost <> "1" Or StoreSettings.bPrintPoNoCost Then
            PrintedCost = "0"
        Else
            PrintedCost = PO.Cost
        End If

        OutputObject.Print(TAB(4), PO.Quantity, TAB(12), PO.Style, TAB(34), PO.Desc)
        If PO.PrintPo = "v" Then
            OutputObject.CurrentX = 5000
            OutputObject.FontBold = True
            OutputObject.Print("VOID VOID VOID")
            OutputObject.FontBold = False
        End If

        'Allow over-write
        OutputObject.CurrentX = 9900
        OutputObject.Print(AlignString(Format(PrintedCost, "$###,##0.00"), 13, AlignConstants.vbAlignRight))

        '  ' EditPO doesn't call this!
        '  If Not EditPO.PrintReport Then
        'Only for first print
        If PO.PrintPo = "v" Then PO.PrintPo = "V" Else PO.PrintPo = "X"
        PO.Save()
        '  End If
        Cost = PO.Cost
        TotCost = TotCost + Cost
    End Sub

    Private Sub PrintPo(ByRef PO As cPODetail)
        'Won't print voided POs
        If Trim(PO.PrintPo) = "V" Then Exit Sub

        GetLocation(PO)
        Heading(PO)
        LineItems(PO)

        Dim LastPO As Long
        LastPO = PO.PoNo

        If PO.DataAccess.Record_EOF Then
            TotalPo
            Exit Sub
        End If

        Do While PO.PoNo = LastPO
            PO.DataAccess.Records_MoveNext()
            LineItems(PO)

            If PO.DataAccess.Record_EOF Then
                TotalPo
                Exit Sub
            End If
        Loop
        TotalPo
    End Sub

    Private Sub TotalPo()
        If PrintedCost <> "0" Then
            OutputObject.Print(TAB(95), "___________")
            OutputObject.Print(TAB(60), "TOTAL:", TAB(95), AlignString(Format(TotCost, "$###,##0.00"), 13, AlignConstants.vbAlignRight, False))
        End If

        OutputObject.CurrentX = 5000
        OutputObject.CurrentY = 14800
        OutputObject.Print("Authorized By: ________________________________")

        Cost = 0
        TotCost = 0
        OutputObject.EndDoc
    End Sub

    Private Sub Heading(ByRef PO As cPODetail)
        OutputObject.FontName = "Arial"
        OutputObject.CurrentX = 0
        OutputObject.CurrentY = 200
        OutputObject.FontSize = 13

        OutputObject.Print(vbCrLf, vbCrLf, TAB(8), "SOLD TO:", TAB(60))

        OutputObject.FontBold = True
        OutputObject.Print("SHIP TO:" & vbCrLf)
        OutputObject.FontBold = False

        Addresses(PO)

        OutputObject.Print(TAB(10), StoreName, TAB(65), StoreShipTo)
        OutputObject.Print(TAB(10), StoreAddress, TAB(65), StoreShipAdd)
        OutputObject.Print(TAB(10), StoreCity, TAB(65), StoreShipCity)
        OutputObject.Print(TAB(10), DressAni(CleanAni(StorePhone)), TAB(65), DressAni(CleanAni(StoreShipPhone)))
        OutputObject.Print(vbCrLf3)

        VendorAddress(PO)

        OutputObject.Print(vbCrLf3)

        If StoreSettings.bPOSpecialInstr Then
            OutputObject.FontSize = 12
            OutputObject.CurrentX = 5900
            OutputObject.CurrentY = 3000
            OutputObject.FontBold = True
            OutputObject.Print(" **** SPECIAL INSTRUCTIONS ****")
            OutputObject.FontBold = False

            OutputObject.CurrentX = 7000
            OutputObject.CurrentY = 3300
            OutputObject.FontUnderline = True
            If PO.Note1 = "1" Then
                OutputObject.Print(" X ")
                OutputObject.FontUnderline = False
            Else
                OutputObject.Print("   ")
            End If
            OutputObject.FontUnderline = False
            OutputObject.CurrentY = 3300


            OutputObject.CurrentX = 7500
            If IsParkPlace Then
                OutputObject.Print("If order is less than 90 lbs., HOLD")
            Else
                OutputObject.Print("If order is less than $300.00, HOLD")
            End If
            OutputObject.CurrentX = 7500
            OutputObject.Print("and SHIP with other goods.")

            OutputObject.CurrentX = 7000
            OutputObject.CurrentY = 3900
            OutputObject.FontUnderline = True
            If PO.Note2 = "1" Then
                OutputObject.Print(" X ")
            Else
                OutputObject.Print("   ")
            End If
            OutputObject.FontUnderline = False
            OutputObject.CurrentY = 3900

            OutputObject.CurrentX = 7500
            OutputObject.Print("Sold orders:  Ship Complete Only")

            OutputObject.CurrentX = 7000
            OutputObject.CurrentY = 4200
            OutputObject.FontUnderline = True
            If PO.Note3 = "1" Then
                OutputObject.Print(" X ")
            Else
                OutputObject.Print("   ")
            End If
            OutputObject.FontUnderline = False
            OutputObject.CurrentY = 4200

            OutputObject.CurrentX = 7500
            OutputObject.Print("Ship UPS, PP OR With Other Goods")

            OutputObject.CurrentX = 7000
            OutputObject.CurrentY = 4500
            OutputObject.FontUnderline = True
            If PO.Note4 = "1" Then
                OutputObject.Print(" X ")
            Else
                OutputObject.Print("   ")
            End If
            OutputObject.FontUnderline = False
            OutputObject.CurrentY = 4500

            OutputObject.CurrentX = 7500
            OutputObject.Print("_______________________________")
        End If

        OutputObject.CurrentX = 0
        OutputObject.CurrentY = 5000
        OutputObject.FontSize = 12

        OutputObject.Print(vbCrLf, TAB(10), "Please put our PO NUMBER, ORDER NUMBER & TAG NAME on all correspondence!" & vbCrLf)

        OutputObject.FontBold = True
        OutputObject.Print(TAB(7), "PO Number: ", PO.PoNo, TAB(30), "Order Number: ", PO.SaleNo, TAB(56), "Date ", DateFormat(dteReportDate.Value), TAB(78), "TAG: ", PO.Name)

        OutputObject.Print(vbCrLf2 & "QUAN.   STYLE NO.                   DESCRIPTION", TAB(93), " Cost")
        OutputObject.FontBold = False
    End Sub

    Private Sub VendorAddress(ByRef PO As cPODetail)
        'Go to AP to get vendor Physical address
        Dim TName As String
        Dim tAddress As String
        Dim tAddress2 As String
        Dim tAddress3 As String
        Dim tZip As String
        Dim tPhone As String
        Dim tFax As String

        If UseQB() Then
            QBGetVendorName(PO.Vendor, TName, tAddress, tAddress2, tAddress3, tZip, tPhone, tFax)
        Else
            GetVendorName(PO.Vendor, TName, tAddress, tAddress2, tAddress3, tZip, tPhone, tFax)
        End If

        If Trim(TName) = "" Then TName = PO.Vendor
        OutputObject.Print(TAB(10), TName)
        OutputObject.Print(TAB(10), tAddress)
        OutputObject.Print(TAB(10), tAddress2 & " " & tZip)
        OutputObject.Print(TAB(10), PhoneAndFax(tPhone, tFax))
    End Sub

    Private Sub GetLocation(ByRef PO As cPODetail)
        Dim OK As Boolean
        If Val(PO.Location) = 0 Then PO.Location = 1
        SI = StoreSettings(PO.Location)
        If Not OK Then Exit Sub

        StoreName = SI.Name
        StoreAddress = SI.Address
        StoreCity = SI.City
        StorePhone = SI.Phone

        StoreShipTo = SI.StoreShipToName
        StoreShipAdd = SI.StoreShipToAddr
        StoreShipCity = SI.StoreShipToCity
        StoreShipPhone = SI.StoreShipToTele
    End Sub

    Private Sub Addresses(ByRef PO As cPODetail)
        If PO.SoldTo = "1" Then
            StoreName = SI.Name
            StoreAddress = SI.Address
            StoreCity = SI.City
            StorePhone = CleanAni(SI.Phone)
        Else
            StoreName = SI.StoreShipToName
            StoreAddress = SI.StoreShipToAddr
            StoreCity = SI.StoreShipToCity
            StorePhone = CleanAni(SI.StoreShipToTele)
        End If

        If PO.ShipTo = "2" Then
            StoreShipTo = SI.StoreShipToName
            StoreShipAdd = SI.StoreShipToAddr
            StoreShipCity = SI.StoreShipToCity
            StoreShipPhone = CleanAni(SI.StoreShipToTele)
        Else
            StoreShipTo = SI.Name
            StoreShipAdd = SI.Address
            StoreShipCity = SI.City
            StoreShipPhone = CleanAni(SI.Phone)
        End If
    End Sub

End Class