Imports Microsoft.VisualBasic.Interaction
Public Class InvPull
    Private mPull As String   ' Transplanted from MainMenu.
    Dim StoreTrans
    Dim CC As Integer
    Dim NoCost As Boolean
    Dim Billing(0 To Setup_MaxStores - 1) As String
    Dim StoreRec(0 To Setup_MaxStores - 1) As String
    Private Const ExtraFieldCount As Integer = 3
    ' Matrix of store transfers.
    ' Values are: 0 = No transfers, 1 = Billing, 2 = Receiving, 3 = Both.
    'Dim TransferList(1 To Setup_MaxStores, 1 To Setup_MaxStores) as integer
    Dim TransferList(0 To Setup_MaxStores - 1, 0 To Setup_MaxStores - 1) As Integer

    Public Property Pull() As Integer
        Get
            Pull = Val(mPull)
        End Get
        Set(value As Integer)
            mPull = value
            arrange
        End Set
    End Property

    Private Sub Arrange(Optional ByVal Working As Boolean = False)
        'MousePointer = IIf(Working, vbHourglass, vbDefault)
        Me.Cursor = IIf(Working, Cursors.WaitCursor, Cursors.Default)
        cmdPrint0.Enabled = Not Working
        cmdPrint1.Enabled = Not Working
        cmdCancel.Enabled = Not Working
    End Sub

    Private Sub InvPull_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'SetButtonImage cmdCancel
        'SetButtonImage cmdPrint(0)
        'SetButtonImage cmdPrint(1), "preview"
        SetButtonImage(cmdCancel, 3)
        SetButtonImage(cmdPrint0, 19)
        'SetButtonImage(cmdprint2,)
        ColorDatePicker(dteFrom)
        ColorDatePicker(dteTo)

        Allow()

        Dim FRM_W2 As Integer, FRA_W2 As Integer
        'FRM_W2 = 4545
        'FRA_W2 = 4215
        FRM_W2 = 454
        FRA_W2 = 421

        'Width = 2565
        Width = 256
        'fraDel.Width = 2250
        fraDel.Width = 225

        lblFrom.Visible = False
        lblTo.Visible = False
        chkShowCost.Visible = False
        chkDriverCopy.Visible = False
        cboStore.Visible = False
        'cmdPrint(1).Visible = False
        cmdPrint1.Visible = False

        lblJuice.Visible = False
        Juice.Visible = False
        chkEmail.Visible = False

        LoadStoresIntoComboBox(cboStore, , , , True)

        Select Case Pull
            Case 1
                dteFrom.Value = Today
                dteTo.Value = Today
                Text = "Pull Load For:"

                optPrintAll0.Visible = True
                optPrintAll1.Visible = True
                optPrintAll2.Visible = False
                txtSaleNo.Visible = False

                'cboStore.Move dteFrom.Left, optPrintAll(1).Top + optPrintAll(1).Height + 60
                cboStore.Location = New Point(dteFrom.Left, optPrintAll1.Top + optPrintAll1.Height + 6)
                cboStore.Visible = True
                'cboStore.Enabled = cboStore.ListCount > 2
                cboStore.Enabled = cboStore.Items.Count > 2
            Case 2
                dteFrom.Value = Today
                dteTo.Value = Today
                Text = "Print Delivery Tickets:"

                optPrintAll0.Visible = True
                optPrintAll1.Visible = True
                optPrintAll2.Visible = True
                txtSaleNo.Visible = False

                'cboStore.Move dteFrom.Left, optPrintAll(2).Top + optPrintAll(2).Height + 60
                cboStore.Location = New Point(dteFrom.Left, optPrintAll2.Top + optPrintAll2.Height + 6)
                cboStore.Visible = True
                'cboStore.Enabled = cboStore.ListCount > 2
                cboStore.Enabled = cboStore.Items.Count > 2

                chkEmail.Visible = True
                chkSoldOrders.Visible = False
            Case 3
                dteFrom.Value = MonthlyReportDefaultStart()
                dteTo.Value = MonthlyReportDefaultEnd()

                lblJuice.Visible = True
                Juice.Visible = True

                Text = "Store Billing For:"
                Width = FRM_W2
                fraDel.Width = FRA_W2

                lblFrom.Visible = True
                lblTo.Visible = True

                optPrintAll0.Visible = False
                optPrintAll1.Visible = False
                optPrintAll2.Visible = False
                txtSaleNo.Visible = False
                cmdPrint1.Visible = True

                chkDriverCopy.Visible = True
                chkDriverCopy.Text = "Dri&ver Pickup Copy?"
                'chkDriverCopy.Value = 0
                chkDriverCopy.Checked = False
                'chkDriverCopy.Move lblFrom.Left, dteTo.Top + dteTo.Height + 60
                chkDriverCopy.Location = New Point(lblFrom.Left, dteTo.Top + dteTo.Height + 6)
                cboStore.Top = chkDriverCopy.Top

                'cboStore.Top = Juice.Top + 280      ' because this is what's used for form sizing
                cboStore.Top = Juice.Top + 28      ' because this is what's used for form sizing
            Case 4
                dteFrom.Value = MonthlyReportDefaultStart()
                dteTo.Value = MonthlyReportDefaultEnd()

                chkShowCost.Visible = True
                chkDriverCopy.Visible = True
                lblJuice.Visible = True
                Juice.Visible = True
                Text = "Store Transfer For:"
                Width = FRM_W2
                fraDel.Width = FRA_W2

                lblFrom.Visible = True
                lblTo.Visible = True

                optPrintAll0.Visible = False
                optPrintAll1.Visible = False
                optPrintAll2.Visible = False
                txtSaleNo.Visible = False

                chkTransferNo.Visible = True

                cmdPrint1.Visible = True

                'chkShowCost.Top = dteFrom.Top + dteFrom.Height + 60
                chkShowCost.Top = dteFrom.Top + dteFrom.Height + 6
                'chkDriverCopy.Top = chkShowCost.Top + chkShowCost.Height + 60
                chkDriverCopy.Top = chkShowCost.Top + chkShowCost.Height + 6

                'cboStore.Move chkDriverCopy.Left, chkDriverCopy.Top + chkDriverCopy.Height + 60
                cboStore.Location = New Point(chkDriverCopy.Left, chkDriverCopy.Top + chkDriverCopy.Height + 6)
                cboStore.Visible = True
                'cboStore.Enabled = cboStore.ListCount > 2
                cboStore.Enabled = cboStore.Items.Count > 2

            Case 5
                dteFrom.Value = MonthlyReportDefaultStart()
                dteTo.Value = MonthlyReportDefaultEnd()

                Text = "Store Transfer List:"
                Width = FRM_W2
                fraDel.Width = FRA_W2

                lblFrom.Visible = True
                lblTo.Visible = True

                optPrintAll0.Visible = False
                optPrintAll1.Visible = False
                optPrintAll2.Visible = False
                txtSaleNo.Visible = False
                cmdPrint1.Visible = True

                chkSoldOrders.Visible = True
                'chkSoldOrders.Move lblFrom.Left, dteTo.Top + dteTo.Height + 60
                chkSoldOrders.Location = New Point(lblFrom.Left, dteTo.Top + dteTo.Height + 6)
                'chkSoldOrders = 1
                chkSoldOrders.Checked = True

                cboStore.Top = chkSoldOrders.Top ' + chkSoldOrders.Height + 60 'dteTo.top + 280      ' because this is what's used for form sizing

            Case 6
                dteFrom.Value = MonthlyReportDefaultStart()
                dteTo.Value = MonthlyReportDefaultEnd()
                Text = "Past Deliveries:"
                Width = FRM_W2
                fraDel.Width = FRA_W2

                optPrintAll0.Visible = False
                optPrintAll1.Visible = False
                optPrintAll2.Visible = False
                txtSaleNo.Visible = False
                cmdPrint1.Visible = True

                chkDriverCopy.Visible = True
                'cboStore.Move dteFrom.Left, dteFrom.Top + dteFrom.Height + 60
                cboStore.Location = New Point(dteFrom.Left, dteFrom.Top + dteFrom.Height + 6)
                cboStore.Visible = True
        End Select

        fraDel.Height = Switch(Pull = 2, cboStore.Top + cboStore.Height + 6 + 30, Pull <= 6, cboStore.Top + cboStore.Height + 6, True, dteFrom.Top + dteFrom.Height + 6)
        fraControls.Top = IIf(Pull <> 7, fraDel.Top + fraDel.Height + 6, fraDel.Top)
        cmdCancel.Left = IIf(Pull >= 3, 216, 108)
        fraControls.Width = cmdCancel.Left + cmdCancel.Width
        'fraControls.Left = ScaleWidth / 2 - fraControls.Width / 2
        fraControls.Left = Me.ClientSize.Width / 2 - fraControls.Width / 2

        'Height = fraControls.Top + fraControls.Height + 60 + (Height - ScaleHeight)
        Height = fraControls.Top + fraControls.Height + 6 + (Height - Me.ClientSize.Height)

        txtFrom.Text = WeekdayName(Weekday(dteFrom.Value))
        txtTo.Text = WeekdayName(Weekday(dteTo.Value))

        optPrintAll0.Checked = True
    End Sub

    Private Sub Allow(Optional ByVal nCost As Boolean = True, Optional ByVal nJuice As Boolean = True, Optional ByVal nDriverCopy As Boolean = True, Optional ByVal nLoc As Boolean = True)
        'If Not nCost Then chkShowCost = 0
        If Not nCost Then chkShowCost.Checked = False
        chkShowCost.Enabled = nCost
        If Not nJuice Then Juice.Text = ""
        Juice.Enabled = nJuice
        lblJuice.Enabled = nJuice
        'If Not nDriverCopy Then chkDriverCopy = 0
        If Not nDriverCopy Then chkDriverCopy.Checked = False
        chkDriverCopy.Enabled = nDriverCopy
        cboStore.Enabled = nLoc
    End Sub

    Private Sub txtSaleNo_TextChanged(sender As Object, e As EventArgs) Handles txtSaleNo.TextChanged
        If Pull = 2 And txtSaleNo.Text <> "" Then
            tmrEmail.Enabled = True
            tmrEmail.Interval = 100
        End If
    End Sub

    Private ReadOnly Property Store() As Integer
        Get
            'Store = cboStore.itemData(cboStore.ListIndex)
            Store = CType(cboStore.Items(cboStore.SelectedIndex), ItemDataClass).ItemData
        End Get
    End Property

    Private Sub EmailDeliveryTicket(ByVal SaleNo As String)
        Dim HtmlText As String
        Dim C As clsMailRec

        If Trim(txtSaleNo.Text) = "" Then MessageBox.Show("Please enter Sale Number", "WinCDS") : Exit Sub
        C = GetMailByLeaseNo(txtSaleNo.Text)
        If Trim(C.Email) = "" Then
            chkEmail.Checked = False
            MessageBox.Show("Email details are not available for the Sale No: " & txtSaleNo.Text, "WinCDS")
            Exit Sub
        End If

        HtmlText = printDeliveryTicketHTML(SaleNo, StoresSld)
        C = GetMailByLeaseNo(SaleNo)
        SendSimpleEmail(StoreSettings.Email, StoreSettings.Name, C.Email, C.First & " " & C.Last, "Delivery Ticket (Sale No:" & SaleNo & ")", HtmlText)
        MessageBox.Show("Email delivered for Delivery Ticket Sale No: " & SaleNo)
    End Sub

    Private Sub cmdPrintClick(sender As Object, e As EventArgs) Handles cmdPrint0.Click, cmdPrint1.Click
        'If Index = 1 Then Set frmPrintPreviewDocument.CallingForm = Nothing
        Dim ButtonName As String
        ButtonName = CType(sender, Button).Name
        If ButtonName = "cmdPrint1" Then
            frmPrintPreviewDocument.CallingForm = Nothing
        End If

        Arrange(True)
        Select Case Pull
            Case 1 : printInvPull_PrintRecords(Store, "" & dteFrom.Value, optPrintAll1.Checked)
            Case 2
                If chkEmail.Checked = True Then
                    EmailDeliveryTicket(txtSaleNo.Text)
                Else
                    printDeliveryTickets_PrintRecords(Store, "" & dteFrom.Value, optPrintAll1.Checked, imgLogo, txtSaleNo.Text)
                End If
            'Case 3 : printCrossSellBilling_PrintRecords 1, LicensedNoOfStores, dteFrom, dteTo, Val(Juice) / 100, Index = 0, chkDriverCopy.Value = 1
            Case 3 : printCrossSellBilling_PrintRecords(1, LicensedNoOfStores, dteFrom.Value, dteTo.Value, Val(Juice) / 100, False, chkDriverCopy.Checked = True)
            Case 4                      ' store transfer billing
                ' Case 4: printListOfStoreTransfers_PrintRecords 1, LicensedNoOfStores, dteFrom, dteTo, Index = 0  ' store transfer billing
                If chkTransferNo.Checked = True Then
                    If txtTransferNo.Text = "" Then
                        MessageBox.Show("Please enter a Transfer Number.", "No Transfer Number")
                    Else
                        'OutputObject = IIf(Index = 1, frmPrintPreviewDocument.picPicture, Printer)
                        OutputObject = IIf(ButtonName = "cmdPrint1", frmPrintPreviewDocument.picPicture, Printer)
                        'OutputToPrinter = (Index = 0)
                        OutputToPrinter = (ButtonName = "cmdPrint0")
                        If Not PrintTransfer(txtTransferNo.Text, chkShowCost.Checked = True, Juice.Text) Then
                            MessageBox.Show("Transfer #" & txtTransferNo.Text & " not found.", "Invalid Transfer Number")
                        End If
                    End If
                Else
                    GetTransfers()
                    'CreateBills Index = 0
                    CreateBills(ButtonName = "cmdPrint0")
                    'If Not IsEmpty(StoreTrans) Then Erase StoreTrans
                    If Not IsNothing(StoreTrans) Then Erase StoreTrans
                End If
            Case 5                    ' List of store transfer
                printListOfStoreTransfers_PrintRecords(Store, dteFrom.Value, dteTo.Value, ButtonName = "cmdPrint0", chkSoldOrders.Checked = True)
            Case 6
                'printPastDeliveries_PrintRecords Store, dteFrom, dteTo, Index = 0
                printPastDeliveries_PrintRecords(Store, dteFrom.Value, dteTo.Value, ButtonName = "cmdPrint0")
        End Select

        If Pull = 2 Then
            'Unload Me
            Me.Close()
        Else
            Arrange(False)
        End If
    End Sub

    Private Sub GetTransfers()
        ' store transfer billing
        Dim InvDetail As New CInventoryDetail
        Dim I As Integer, X As Double

        'MousePointer = 11
        Me.Cursor = Cursors.WaitCursor

        ProgressForm(0, 100, "Initializing...")
        CC = 0
        On Error GoTo HandleErr

        For I = 1 To Setup_MaxStores
            Billing(I) = False
            StoreRec(I) = False
        Next

        Dim SQL As String
        SQL = "SELECT * FROM Detail WHERE "
        'If cboStore.Visible And cboStore.itemData(cboStore.ListIndex) > 0 Then
        If cboStore.Visible And CType(cboStore.Items(cboStore.SelectedIndex), ItemDataClass).ItemData > 0 Then
            'SQL = SQL & "Loc" & cboStore.itemData(cboStore.ListIndex) & "< 0 AND "
            SQL = SQL & "Loc" & CType(cboStore.Items(cboStore.SelectedIndex), ItemDataClass).ItemData & "< 0 AND "
        End If
        SQL = SQL & "DDate1 BETWEEN #" & dteFrom.Value & "# AND #" & dteTo.Value & "# AND (Trans='TR' OR Trans='TP') ORDER BY DetailID"
        InvDetail.DataAccess.Records_OpenSQL(SQL)

        If InvDetail.DataAccess.Record_Count > 0 Then
            ReDim StoreTrans(InvDetail.DataAccess.Record_Count, Setup_MaxStores + ExtraFieldCount)
        End If
        ProgressForm(1, InvDetail.DataAccess.Record_Count, "Getting Transfers...")


        Do While InvDetail.DataAccess.Records_Available
            'build an array for stores to hold transfers as it goes throught the file
            CC = CC + 1
            ProgressForm(CC)

            For I = 1 To Setup_MaxStores
                X = InvDetail.GetLocationQuantity(I)
                If X <> 0 Then
                    If X < 0 Then Billing(I) = True 'to see if we need to create a bill from
                    If X > 0 Then StoreRec(I) = True
                    StoreTrans(CC, I) = X
                End If
            Next

            StoreTrans(CC, Setup_MaxStores + 1) = Val(InvDetail.Misc) 'transfer no
            StoreTrans(CC, Setup_MaxStores + 2) = InvDetail.DDate1
            StoreTrans(CC, Setup_MaxStores + 3) = InvDetail.Style 'to get landed cost each
        Loop
        ProgressForm()

        DisposeDA(InvDetail)
        Exit Sub

HandleErr:
        Resume Next
    End Sub

    Private Sub CreateBills(ByVal PrintIt As Boolean)
        Dim X As Integer, R As Integer
        Dim CD As Integer, C As Integer, DD As Integer
        Dim TotTransfer As Decimal
        Dim Counter As Integer
        Dim InvData As New CInvRec
        Dim JuiceAmt As Double

        X = ssMaxStore

        NoCost = False
        If chkDriverCopy.Checked = False And chkShowCost.Checked = True Then
            If Not CheckAccess("View Cost and Gross Margin", True, True, True) Then
                NoCost = True
            End If
        Else
            NoCost = True
        End If

        'MousePointer = vbHourglass
        Me.Cursor = Cursors.WaitCursor
        R = X * X * CC
        ProgressForm(0, R, "Creating Bills...")

        If PrintIt Then
            OutputObject = Printer
            OutputToPrinter = True
            If IsUFO() Then Printer.Copies = 2
        Else
            'Load frmPrintPreviewMain
            OutputToPrinter = False
            OutputObject = frmPrintPreviewDocument.picPicture
            frmPrintPreviewDocument.CallingForm = Me
            frmPrintPreviewDocument.ReportName = "Multi-Store Transfer Report"
        End If

        TotTransfer = 0
        Counter = 0
        For CD = 1 To X               ' Cycles first:  Billing store one by one
            R = (CD - 1) * X * CC
            ProgressForm(R, , "Creating Bills: Store " & CD & "...")
            For DD = 1 To X             ' CD = Inside Loop  'Receiving store
                R = ((CD - 1) * X * CC) + (DD - 1) * CC
                ProgressForm(R)
                If DD <> CD Then          ' Skip same store comparisons..
                    R = (((CD - 1) * X * CC) + (DD - 1) * CC) + CC
                    ProgressForm(R)
                    If CBool(Billing(CD)) And CBool(StoreRec(DD)) Then
                        For C = 1 To CC
                            R = ((((CD - 1) * X * CC) + (DD - 1) * CC) + CC) + C
                            ProgressForm(R)
                            If StoreTrans(C, CD) < 0 And StoreTrans(C, DD) > 0 Then
                                If InvData.Load(Trim(StoreTrans(C, Setup_MaxStores + 3))) Then

                                    ' Only print newpage and header if there are records...
                                    If Counter >= 26 Or (Counter = 0 And OutputObject.CurrentY > 0) Then
                                        If OutputToPrinter Then Printer.NewPage() Else frmPrintPreviewDocument.NewPage()
                                        Counter = 0
                                    End If

                                    If Counter = 0 Then
                                        GetBillHeading(CD, True)
                                        GetBillTo(DD)
                                    End If

                                    PrintToPosition(OutputObject, StoreTrans(C, Setup_MaxStores + 1), 0, VBRUN.AlignConstants.vbAlignLeft, False)
                                    PrintToPosition(OutputObject, StoreTrans(C, Setup_MaxStores + 2), 3250, VBRUN.AlignConstants.vbAlignRight, False)
                                    PrintToPosition(OutputObject, Math.Abs(StoreTrans(C, DD)), 4500, VBRUN.AlignConstants.vbAlignRight, False)
                                    PrintToPosition(OutputObject, StoreTrans(C, Setup_MaxStores + 3), 5000, VBRUN.AlignConstants.vbAlignLeft, (NoCost))
                                    If Not NoCost Then PrintToPosition(OutputObject, Format(InvData.Landed * Math.Abs(StoreTrans(C, DD)), "###,###.00"), 11250, VBRUN.AlignConstants.vbAlignRight, True)

                                    TotTransfer = TotTransfer + Format(InvData.Landed * Math.Abs(StoreTrans(C, DD)), "###,###.00")
                                    Counter = Counter + 1
                                End If
                            End If
                        Next

                        ' Always show the total, if anything was printed.
                        If Counter > 0 Then
                            If Not NoCost Then
                                If Microsoft.VisualBasic.Right(Juice.Text, 1) = "%" Then
                                    JuiceAmt = TotTransfer * GetDouble(Microsoft.VisualBasic.Left(Juice.Text, Len(Juice.Text) - 1)) / 100
                                Else
                                    JuiceAmt = TotTransfer * GetDouble(Juice.Text) / 100
                                End If
                                '20080219 - Changed Juice Field Text
                                '                If JuiceAmt <> 0 Then PrintToPosition OutputObject, "Juice:               " & Format(JuiceAmt, "###,###.00"), 11250, vbAlignRight, True
                                If JuiceAmt <> 0 Then PrintToPosition(OutputObject, "Warehouse Charge:    " & Format(JuiceAmt, "###,###.00"), 11250, VBRUN.AlignConstants.vbAlignRight, True)
                                PrintToPosition(OutputObject, "_________", 11250, VBRUN.AlignConstants.vbAlignRight, True)
                                PrintToPosition(OutputObject, "Total Transfer:      " & Format(TotTransfer + JuiceAmt, "###,###.00"), 11250, VBRUN.AlignConstants.vbAlignRight, True)
                            End If

                            ' If this is a receiving report and something got printed, show a signature line.
                            If chkDriverCopy.Checked = True Then 'Pass = 1 Then 'StoreRec(DD) = True Then   ' if pass=1 then?
                                If OutputObject.CurrentY > 14000 Then
                                    If OutputToPrinter Then Printer.NewPage() Else frmPrintPreviewDocument.NewPage()
                                    GetBillHeading(CD, True)
                                    GetBillTo(DD)
                                End If
                                OutputObject.CurrentY = 13500
                                PrintToPosition(OutputObject, "Driver:__________________________", 7000, VBRUN.AlignConstants.vbAlignRight, True)
                                OutputObject.Print
                                PrintToPosition(OutputObject, "Shipped By:__________________________", 7000, VBRUN.AlignConstants.vbAlignRight, True)
                                OutputObject.Print
                                PrintToPosition(OutputObject, "Received By:__________________________", 7000, VBRUN.AlignConstants.vbAlignRight, False)
                                PrintToPosition(OutputObject, "Date:______________", 7400, VBRUN.AlignConstants.vbAlignLeft, True)
                            End If
                            Counter = 0 ' Force a new page on the next pass.
                        End If
                        TotTransfer = 0
                    End If

                End If
                '      Printer.EndDoc  ' Don't end it here.. just go to a new page.
            Next
            Counter = 0
            'Counter = 0
        Next

        DisposeDA(InvData)

        If PrintIt Then
            OutputObject.EndDoc
        Else
            If OutputToPrinter Then 'If error, do not show form
                'MsgBox "Unknown error", vbCritical, "Print Preview Setup","Inventory Reports"
                'Unload frmPrintPreviewMain
                frmPrintPreviewMain.Close()
            Else
                Hide()
                frmPrintPreviewDocument.DataEnd()
            End If
        End If

        ProgressForm()
        'MousePointer = vbDefault
        Me.Cursor = Cursors.Default
    End Sub

    Private Sub GetBillTo(ByVal Store As Integer)
        Dim SI As StoreInfo
        SI = StoreSettings(Store)

        OutputObject.CurrentY = 600
        OutputObject.Print(TAB(75), "To Loc: ", Store)
        OutputObject.CurrentY = 3600

        OutputObject.CurrentX = 200
        OutputObject.Print(SI.Name)
        OutputObject.CurrentX = 200
        OutputObject.Print(SI.Address)
        OutputObject.CurrentX = 200
        OutputObject.Print(SI.City)
        OutputObject.CurrentX = 200
        OutputObject.Print(SI.Phone)
        OutputObject.Print

        OutputObject.FontBold = True
        'Printer.Print "Transfer No: ", "Date:"; Tab(25); "Quantity", "Style"; Tab(75); "Landed"
        PrintToPosition(OutputObject, "Transfer No", 0, VBRUN.AlignConstants.vbAlignLeft, False)
        PrintToPosition(OutputObject, "Date", 3250, VBRUN.AlignConstants.vbAlignRight, False)
        PrintToPosition(OutputObject, "Quantity", 4500, VBRUN.AlignConstants.vbAlignRight, False)
        PrintToPosition(OutputObject, "Style", 5000, VBRUN.AlignConstants.vbAlignLeft, NoCost)
        If Not NoCost Then PrintToPosition(OutputObject, "Landed", 11250, VBRUN.AlignConstants.vbAlignRight, True)
        OutputObject.FontBold = False
    End Sub

    Private Sub GetBillHeading(ByVal Store As Integer, ByVal Rev As Boolean)

        Dim SI As StoreInfo
        SI = StoreSettings(Store)
        '  Store = CD
        '-NEW 2003-02-20AA: GetLocation
        '-NEW 2003-02-20AA:  Pickcompany

        OutputObject.FontName = "Arial"
        OutputObject.FontBold = True
        OutputObject.FontSize = 18
        OutputObject.CurrentY = 500

        PrintCentered(SI.Name)
        PrintCentered(SI.Address)
        PrintCentered(SI.City)
        PrintCentered(SI.Phone)


        ' Delivery day & Date
        OutputObject.FontBold = False
        OutputObject.FontSize = 14
        OutputObject.CurrentX = 100
        OutputObject.CurrentY = 300
        OutputObject.Print("    From: ", dteFrom.Value, TAB(75), "To: ", dteTo.Value)

        If Pull <> "5" Then ' List of store transfer
            OutputObject.CurrentX = 100
            OutputObject.Print("    From Loc: ", Store)
        End If
        OutputObject.CurrentX = 100
        OutputObject.CurrentY = 3000
        If Pull <> "5" Then ' List of store transfer
            OutputObject.Print("Bill" & IIf(Rev, "", "ed") & " To:")
        End If
        OutputObject.Print
    End Sub

    'Private Sub OptPrintCheckedChanged() Handles optPrintAll0.CheckedChanged, optPrintAll1.CheckedChanged, optPrintAll2.CheckedChanged
    '    'If Index = 2 Then
    '    If optPrintAll2.Checked = True Then
    '        txtSaleNo.Visible = True
    '        On Error Resume Next
    '        txtSaleNo.Select()
    '        chkEmail.Enabled = Pull = 2
    '    Else
    '        txtSaleNo.Visible = False
    '        txtSaleNo.Text = ""
    '        chkEmail.Enabled = False
    '    End If

    '    'chkEmail.Value = vbUnchecked
    '    chkEmail.Checked = False
    'End Sub
    Private Sub OptPrintCheckedChanged(sender As Object, e As EventArgs) Handles optPrintAll2.CheckedChanged, optPrintAll1.CheckedChanged, optPrintAll0.CheckedChanged
        Dim optSelected As String

        optSelected = CType(sender, RadioButton).Name
        If optSelected = "optPrintAll2" Then
            txtSaleNo.Visible = True
            On Error Resume Next
            txtSaleNo.Select()
            chkEmail.Enabled = Pull = 2
        Else
            txtSaleNo.Visible = False
            txtSaleNo.Text = ""
            chkEmail.Enabled = False
        End If

        'chkEmail.Value = vbUnchecked
        chkEmail.Checked = False
    End Sub
End Class