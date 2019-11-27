Imports Microsoft.VisualBasic.Compatibility.VB6
Public Class OrdPay
    Dim Status As String             ' Used by cmdOK, needs to be saved between clicks..
    Dim OrgHoldingStatus As String   ' Original Holding Status.
    Dim PayMethod As String
    Dim PriorBal As Decimal
    Dim Deposit As Decimal

    Dim DeliveredAuditRecord As Integer, DeliveredPayment As Decimal

    ' Cash and Audit variables.. global until we rewrite those calls.
    Dim LeaseNo As String
    Dim Note As String
    Dim Money As Decimal
    Dim Account As String
    Dim Cashier As String
    Dim Name1 As String
    Dim TransDate As String
    Dim Written As Decimal
    Dim TaxCharged1 As Decimal
    Dim ArCashSls As Decimal
    Dim Controll As Decimal
    Dim UndSls As Decimal
    Dim DelSls As Decimal
    Dim TaxRec1 As Decimal
    Dim SalesTax1 As Decimal

    Dim Approval As String

    Dim FinanceArNo As String

    Public X As Integer                  ' BillOSale.GrossMargin checks this.
    Public Sale As Decimal           ' Called by ArPaySetup
    Public TotDeposit As Decimal     ' Called by ArPaySetup

    Private LockOn As Boolean         ' Used to simulate Modal state

    Public Sub FinanceOnAccount(ByVal ArNo As String)
        Dim X As Integer, GM As CGrossMargin, objHolding As cHolding
        ' Prepare the Holding object to accept info.

        FinanceArNo = ArNo
        PaymentOnAccount()
        FinanceArNo = ""
        objHolding = New cHolding
        objHolding.Load(g_Holding.LeaseNo)  ' Load the most current info.
        PostPaymentOnAccount(objHolding)
        objHolding.ArNo = ArNo ' this would work if the code came this direction!
        objHolding.Save()
        FinishRoutine(False)

        DisposeDA(objHolding)
    End Sub

    Private Function PaymentOnAccount() As Boolean
        Dim RealSaleDate As String
        Dim Amount As Decimal, pType As String, Appr As String, CCPayment As Boolean, TypeCode As Integer
        Dim TransID As String, BalanceReturned As Decimal
        RealSaleDate = BillOSale.dteSaleDate.Value

        PriorBal = 0
        PriorBal = BillOSale.BalDue.Text
        Deposit = GetPrice(txtAmount.Text)

        Amount = GetPrice(txtAmount.Text)
        pType = PayMethod
        'TypeCode = cboAccount.itemData(cboAccount.SelectedIndex)
        TypeCode = CType(cboAccount.SelectedItem, ItemDataClass).ItemData
        CCPayment = False

        ' Don't post 0 value payments, unless they're credit vouchers.
        If Deposit = 0 And Not PayTypeIsFinance(pType) Then
            PaymentOnAccount = True
            Exit Function
        End If

        If SwipeCreditCards() And PayTypeIsCC(PayMethod) Then
            If chkPayAll.Checked = True Then Amount = BillOSale.BalDue.Text
            If Amount < 0 Then
                If Not ProcessCCReturn(-Amount, Appr, "#") Then Exit Function
            Else
                If Not ProcessCC(Amount, pType, Appr, TypeCode, TransID, , BalanceReturned) Then Exit Function
            End If
            Approval = Appr
            CCPayment = True
        ElseIf SwipeDebitCards() And PayTypeIsDebit(PayMethod) Then
            If chkPayAll.Checked = True Then Amount = BillOSale.BalDue.Text
            If Amount < 0 Then
                If Not ProcessDebitReturn(-Amount, Appr) Then Exit Function
            Else
                If Not ProcessDebit(Amount, pType, Appr, TypeCode) Then Exit Function
            End If
            Approval = Appr
            CCPayment = True
        ElseIf SwipeGiftCards() And PayTypeIsStoreCard(PayMethod) Then
            If chkPayAll.Checked = True Then Amount = BillOSale.BalDue.Text
            If Amount < 0 Then
                If Not ProcessGiftCardReturn(-Amount, Appr) Then Exit Function
            Else
                If Not ProcessGiftCard(Amount, pType, Appr, TypeCode) Then Exit Function
            End If
            Approval = Appr
            CCPayment = True
        End If

        If CCPayment And GetPrice(Amount) <> GetPrice(txtAmount.Text) Then
            txtAmount.Text = CurrencyFormat(Amount)
            Deposit = GetPrice(txtAmount.Text)
        End If

        'If cboAccount.List(cboAccount.ListIndex) = "" Then
        If cboAccount.SelectedText = "" Then
            cboAccount.SelectedIndex = 0
        End If

        X = MailCheck.X  ' We shouldn't have to get this from MailCheck, and shouldn't have to open BillOSale to add records.
        BillOSale.X = X

        ' sub total line, except when delivering which should show a 'delivered' subtotal line
        BillOSale.SetStyle(X, "SUB")
        If OrderMode("B") Then BillOSale.SetStatus(X, "DEL")
        BillOSale.SetDesc(X, "               Sub Total =")  '    BillOSale.Desc = "               Sub Total ="
        BillOSale.SetPrice(X, BillOSale.BalDue.Text) '.Price = BillOSale.BalDue
        '.Price = Format(BillOSale.Price, "###,###.00")

        frmSalesList.SalesCode = MailCheck.SalesPerson
        BillOSale.RN = 0
        BillOSale.dteSaleDate.Value = TransDate         ' Changes sale date on BillOSale.

        'jk this form may not be loaded under payment
        Dim L As String
        '    Dim Margin As CGrossMargin, L As String
        '    Set Margin = New CGrossMargin
        '    Margin.DDelDat = TransDate   ' BFH20050516:  this was putting incorrect delivery dates for orders

        '    LastPay = TransDate

        ' **** This was losing SalesSplit! Need to check more code for similar omissions. MJK20130914 ****
        SaveNewMarginRecord(BillOSale.BillOfSale.Text, BillOSale.QueryStyle(X), BillOSale.QueryDesc(X), Val(BillOSale.QueryQuan(X)), GetPrice(BillOSale.QueryPrice(X)),
      BillOSale.QueryMfg(X), "", BillOSale.QueryMfgNo(X), 0, 0, 0, BillOSale.PorD, "", IIf(OrderMode("B"), "DEL", ""), "", BillOSale.QueryLoc(X), Today, BillOSale.DelDate,
      StoresSld, BillOSale.CustomerLast.Text, "", BillOSale.CustomerPhone1.Text, BillOSale.MailIndex _
      , , , , , , , , BillOSale.vGetSalesSplit) ' this is the SUB line, before PAYMENT

        '    BillOSale.GrossMargin Margin           ' Save the new Subtotal.
        '    DisposeDA Margin
        '    Set Margin = New CGrossMargin

        X = X + 1
        BillOSale.X = X
        BillOSale.SetStyle(X, "PAYMENT")
        If OrderMode("B") Then BillOSale.SetStatus(X, "DEL")
        BillOSale.SetQuan(X, TypeCode)

        L = Trim(Microsoft.VisualBasic.Left(pType, 20) & " " & TransDate & "  " & Memo.Text & Replace(Mid(Approval, 11), "Approval", "Appr"))
        If FinanceArNo <> "" Then L = L & "  Account #" & FinanceArNo
        If CCPayment Then L = ArrangeString(L, 46) & ArrangeString("I Authorize The Above Transaction:", 46) & "X ______________________________"
        BillOSale.SetDesc(X, L)

        BillOSale.SetPrice(X, Amount)

        ' check for blank or no deposit amount
        TotDeposit = TotDeposit + Deposit
        BillOSale.BalDue.Text = BillOSale.BalDue.Text - Deposit

        '.Price = Format(BillOSale.Price, "###,###.00")
        BillOSale.BalDue.Text = Format(BillOSale.BalDue.Text, "###,###.00")

        frmSalesList.SalesCode = MailCheck.SalesPerson
        BillOSale.RN = 0
        BillOSale.dteSaleDate.Value = TransDate

        BillOSale.SetTransID(X, TransID)

        SaveNewMarginRecord(BillOSale.BillOfSale.Text, BillOSale.QueryStyle(X), BillOSale.QueryDesc(X), Val(BillOSale.QueryQuan(X)), GetPrice(BillOSale.QueryPrice(X)),
     BillOSale.QueryMfg(X), "", BillOSale.QueryMfgNo(X), 0, 0, 0, BillOSale.PorD, "", IIf(OrderMode("B"), "DEL", ""), "", BillOSale.QueryLoc(X), Today, BillOSale.DelDate,
     StoresSld, BillOSale.CustomerLast.Text, "", BillOSale.CustomerPhone1.Text, BillOSale.MailIndex, , , , , , , TransID, BillOSale.vGetSalesSplit)


        '    On Error GoTo HandleErr       ' MJK 20030430 to combat duplicate records.
        '    BillOSale.GrossMargin Margin           ' Save the new Payment.
        '    On Error GoTo 0


        If BalanceReturned <> 0 Then
            X = X + 1
            BillOSale.X = X
            BillOSale.SetStyle(X, "NOTES")
            BillOSale.SetDesc(X, "CARD BALANCE: " & CurrencyFormat(BalanceReturned))
            If OrderMode("B") Then BillOSale.SetStatus(X, "DEL")

            SaveNewMarginRecord(BillOSale.BillOfSale.Text, BillOSale.QueryStyle(X), BillOSale.QueryDesc(X), Val(BillOSale.QueryQuan(X)), GetPrice(BillOSale.QueryPrice(X)),
       BillOSale.QueryMfg(X), "", BillOSale.QueryMfgNo(X), 0, 0, 0, BillOSale.PorD, "", "", "", BillOSale.QueryLoc(X), Today, BillOSale.DelDate,
       StoresSld, BillOSale.CustomerLast.Text, "", BillOSale.CustomerPhone1.Text, BillOSale.MailIndex _
       , , , , , , , , BillOSale.vGetSalesSplit)
            '      On Error GoTo HandleErr
            '      BillOSale.GrossMargin Margin           ' Save the new record.
            '      On Error GoTo 0
        End If

        BillOSale.dteSaleDate.Value = RealSaleDate

        BillOSale.GridMove(X)

        PaymentOnAccount = True
        Exit Function

HandleErr:
        'MsgBox "Error in PaymentOnAccount: " & Err.Number & ", " & Err.Description, vbCritical, "Error!"
        MessageBox.Show("Error in PaymentOnAccount: " & Err.Number & ", " & Err.Description, "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        Resume Next
    End Function

    Public Sub FinishRoutine(ByVal StayOnOrder As Boolean)
        ' This function uses the global Holding object.
        ' It will be reworked once I've researched all the side effects.

        If Receipt Then MakeMyReceipt() : Receipt = False
        If Email Then MakeEmail() : Email = False

        If StayOnOrder Then
            ' Clear temporary stuff..
            DoControls(True)
            BillOSale.cmdMainMenu.Enabled = True
            cboAccount.SelectedIndex = 0
            txtAmount.Text = ""
            Exit Sub
        End If

        DeliveredAuditRecord = 0
        DeliveredPayment = 0

        If OrderMode("B") Then
            If MessageBox.Show("Any More To Deliver?", "", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
                'Unload OrdPay
                Me.Close()
                'Unload ArCard
                ArCard.Close()
                'Unload ARPaySetUp
                ARPaySetUp.Close()
                'Unload BillOSale
                BillOSale.Close()
                BillOSale.Show()
                BillOSale.BillOSale2_Show()

                ' bfh20051010 - These lines moved before the mailcheck.show, esp taxrec1 and taxrec2 b/c
                ' they are called after all the delivery lines are processed... this meant that
                ' they cleared taxrec1&2 AFTER they were calculated, leaving zero's in the sales tax
                ' report, among other things...
                g_Holding.Status = ""
                Status = ""
                TotDeposit = 0
                Deposit = 0
                InvDel.TaxRec2 = "0.00"
                InvDel.TaxRec1 = "0.00"
                TaxRec1 = "0.00"
                frmSalesList.SalesCode = ""
                '''''''''''''''

                X = 0
                MailCheck.FirstRec = 0
                MailCheck.optSaleNo.Checked = True
                'MailCheck.Show vbModal, BillOSale
                MailCheck.ShowDialog(BillOSale)
                Exit Sub
            End If

            ' No more to deliver
            'Unload BillOSale
            BillOSale.Close()
            'Unload ARPaySetUp
            ARPaySetUp.Close()
            'Unload ArCard ' need for add on
            ArCard.Close()
            'Unload AddOnAcc
            AddOnAcc.Close()
            AddOnAcc = Nothing
            'Unload ArCard ' Why twice?  Does unloading AddOnAcc reload this?
            ArCard.Close()

            MainMenu.Show()
            MailCheck.FirstRec = 0
            InvDel.TaxRec2 = "0.00"
            InvDel.TaxRec1 = "0.00"
            TaxRec1 = "0.00"
            TotDeposit = 0
            Deposit = 0
            Status = ""
            g_Holding.Status = ""
            frmSalesList.SalesCode = ""
            ARPaySetUp.AccountFound = ""
            modProgramState.Order = ""
            modProgramState.ArSelect = ""
            TransDate = ""
            'Unload OrdPay
            Me.Close()
            Exit Sub
        End If

        If OrderMode("D") Then
            BillOSale.UGridIO1.GetDBGrid.Refresh() 'bfh20060113 - refresh & doevents added for cosmetic fix
            Application.DoEvents()

            If MessageBox.Show("Any More To Pay On?", "", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
                PayMethod = ""
                BillOSale.cmdMainMenu.Enabled = True
                'Unload OrdPay
                Me.Close()
                BillOSale.UGridIO1.Clear()
                MailCheck.FirstRec = 0
                TotDeposit = 0
                Deposit = 0
                'frmSalesList.SafeSalesClear = True
                frmSalesList.SalesCode = ""
                ' This allows retention of current customer for consecutive payments.
                ' Some customers may not want this, so they should go in the else block, using the old code, upon request
                If True Then
                    MailCheck.optTelephone.Checked = True
                    MailCheck.InputBox.Text = DressAni(CleanAni(BillOSale.CustomerPhone1.Text))
                Else
                    MailCheck.optSaleNo.Checked = True
                End If
                'MailCheck.Show vbModal, BillOSale
                MailCheck.ShowDialog(BillOSale)
                Exit Sub
            Else
                PayMethod = ""
                'Unload OrdPay
                Me.Close()
                'Unload BillOSale
                BillOSale.Close()
                MainMenu.Show()
                MailCheck.FirstRec = 0
                TotDeposit = 0
                Deposit = 0
                frmSalesList.SalesCode = ""
                TransDate = ""
            End If
        End If

        g_Holding.Status = ""
    End Sub

    Private Sub PostPaymentOnAccount(ByRef Holding As cHolding)
        Dim tAuditRS As ADODB.Recordset, SFNeedNewAudit As Boolean
        Dim Tp As Decimal

        OpenCashDrawer()

        LeaseNo = BillOSale.BillOfSale.Text
        Money = Deposit
        'Account = cboAccount.itemData(cboAccount.ListIndex)
        Account = CType(cboAccount.SelectedItem, ItemDataClass).ItemData
        If Account = cdsPayTypes.cdsPT_CompanyCheck Then Account = "21500"
        Note = BillOSale.CustomerLast.Text
        Cashier = GetCashierName

        If Account <> 10 Then Cash() ' discount

        On Error GoTo HandleErr


        '  If Order = "D" And MailCheck.OrigStatus = "C" Or MailCheck.OrigStatus = "B" Or MailCheck.OrigStatus = "F" Then
        If OrderMode("D", "B") And IsIn(MailCheck.OrigStatus, "C", "B", "F") Then
            If Val(BillOSale.BalDue) = 0 Then
                Holding.Status = OrgHoldingStatus
            End If
        End If

        If OrderMode("D") And IsIn(MailCheck.OrigStatus, "O") And PayTypeIs(PayMethod) = cdsPayTypes.cdsPT_StoreFinance Then
            Holding.Status = "S" ' "F" 'BFH20060804
            Holding.Save()
        End If

        If OrderMode("D") And IsIn(MailCheck.OrigStatus, "O") And PayTypeIsOutsideFinance(PayMethod) And (GetPrice(txtAmount.Text) = 0 Or GetPrice(txtAmount.Text) = BillOSale.SaleTotal) Then
            Holding.Status = "E" ' "F" 'BFH20060804
            Holding.Save()
        End If

        '  If OrderMode("B") And IsIn(MailCheck.OrigStatus, "E") Then
        '    Holding.Status = "C"
        '    Holding.Save
        '  End If


        If OrderMode("B") And IsIn(MailCheck.OrigStatus, "S") Then
            Holding.Status = "F"
            'BFH20161030 - StoreFinanceAsDelivered flag added... SF sales are no longer "delivered" upon creation
            If Not StoreFinanceAsDelivered Then
                'AuditID  SaleNo  Name1     TransDate   Written TaxCharged1 ArCashSls Controll  UndSls    DelSls  TaxRec1 TaxCode Salesman
                '1855     10350   DS CARTER 10/26/2016  $0.00   $0.00       ($392.15) $0.00     ($392.15) $369.95 $22.20  1       99
                Dim sRR As String
                sRR = ""
                sRR = sRR & "SELECT * FROM [Audit]"
                sRR = sRR & " WHERE 1=1"
                sRR = sRR & " AND Left([Name1],3) IN ('SF ', 'NS ')"
                sRR = sRR & " AND SaleNo='" & Holding.LeaseNo & "'"
                tAuditRS = GetRecordsetBySQL(sRR, , GetDatabaseAtLocation)
                If tAuditRS.RecordCount <> 0 Then
                    If IfNullThenZeroCurrency(tAuditRS("DelSls")) = 0 Then SFNeedNewAudit = True
                End If
                DisposeDA(tAuditRS)
                If SFNeedNewAudit Then
                    Tp = BillOSale.SaleTotal("paid") - BillOSale.SaleTotal(PayListItem(cdsPayTypes.cdsPT_OutsideFinance)) -
                                           BillOSale.SaleTotal(PayListItem(cdsPayTypes.cdsPT_OutsideFinance2)) -
                                           BillOSale.SaleTotal(PayListItem(cdsPayTypes.cdsPT_OutsideFinance3)) -
                                           BillOSale.SaleTotal(PayListItem(cdsPayTypes.cdsPT_OutsideFinance4)) -
                                           BillOSale.SaleTotal(PayListItem(cdsPayTypes.cdsPT_OutsideFinance5))
                    AddNewAuditRecord(Holding.LeaseNo, "DS " & BillOSale.CustomerLast.Text, dtePayDate.Value, 0, 0, -BillOSale.SaleTotal("gross"), Tp, -BillOSale.SaleTotal("delivered"), BillOSale.SaleTotal("written"), BillOSale.SalesTax1 + BillOSale.SalesTax2, IIf(MailCheck.TaxCode = 0, 1, MailCheck.TaxCode), MailCheck.SalesPerson)
                    ' For sales created in the new way, with StoreFinanceAsDelivered == False, these sales also need the
                    ' Sale total into Receivables
                    AddNewCashJournalRecord("11300", BillOSale.SaleTotal(PayListItem(cdsPayTypes.cdsPT_OutsideFinance)), Holding.LeaseNo, "", dtePayDate.Value) ' We need to do this too, in case of partial payment.
                    AddNewCashJournalRecord("11300", BillOSale.SaleTotal(), Holding.LeaseNo, "", dtePayDate.Value)
                End If
            End If
            Holding.Save()
        End If

        If OrderMode("B") And (IsIn(MailCheck.OriginalPrint, "E") Or Holding.Status = "E") Then
            Holding.Status = "C"

            tAuditRS = GetRecordsetBySQL("SELECT * FROM [Audit] WHERE Left([Name1],3)='NS ' AND SaleNo='" & Holding.LeaseNo & "'", , GetDatabaseAtLocation)
            If tAuditRS.RecordCount <> 0 Then
                If IfNullThenZeroCurrency(tAuditRS("DelSls")) = 0 Then SFNeedNewAudit = True
            End If
            DisposeDA(tAuditRS)

            If SFNeedNewAudit Then
                Tp = BillOSale.SaleTotal("paid") - BillOSale.SaleTotal(PayListItem(cdsPayTypes.cdsPT_OutsideFinance))
                AddNewAuditRecord(Holding.LeaseNo, "DS " & BillOSale.CustomerLast.Text, dtePayDate.Value, 0, 0, -BillOSale.SaleTotal("gross"), Tp, -BillOSale.SaleTotal("delivered"), BillOSale.SaleTotal("written"), BillOSale.SalesTax1 + BillOSale.SalesTax2, IIf(MailCheck.TaxCode = 0, 1, MailCheck.TaxCode), MailCheck.SalesPerson)
                ' For sales created in the new way, with StoreFinanceAsDelivered == False, these sales also need the
                ' Sale total into Receivables
                ' We must do total of cdsPT_OutsideFinance and SaleTotal, because there can be a partial as well as a "rest of the sale" application..
                ' They shouldn't ever happen together, since we set the total to 0.00 if it is for the rest of the sale and set the sale type to Open Credit on payment creation...
                AddNewCashJournalRecord("11300", BillOSale.SaleTotal(PayListItem(cdsPayTypes.cdsPT_OutsideFinance)) + BillOSale.SaleTotal, Holding.LeaseNo, "", dtePayDate.Value)
            End If

            Holding.Save()
        End If

        If OrderMode("D") And MailCheck.OrigStatus = "D" And Holding.Status = "B" Then
            ' This just puts in a zero line...  Because they're changing a Delivered to Backordered with a payment
            ' There is no control value, it all goes into the backorder account.
            Holding.Deposit = GetPrice(Holding.Deposit) + Money
            LeaseNo = BillOSale.BillOfSale.Text
            Name1 = "PA " + BillOSale.CustomerLast.Text
            Written = "0.00"
            TaxCharged1 = "0.00"
            ArCashSls = "0.00"
            Controll = "0.00"
            UndSls = "0.00"
            DelSls = "0.00"
            TaxRec1 = "0.00"
            Audit() 'puts a o entry in Audit
            If Money > 0 Then Holding.LastPay = DateFormat(TransDate)
        ElseIf OrderMode("D", "B") And Not PayTypeIsIn(cboAccount.SelectedText, cdsPayTypes.cdsPT_MiscDiscount) Then
            Holding.Deposit = GetPrice(Holding.Deposit) + Money
            LeaseNo = BillOSale.BillOfSale.Text
            Name1 = "PA " + BillOSale.CustomerLast.Text
            Written = "0.00"
            TaxCharged1 = "0.00"
            ArCashSls = "0.00"
            Controll = -GetPrice(txtAmount.Text)

            If IsIn(Trim(BillOSale.SaleStatus.Text), "BACK ORD.", "CREDIT") Then   'added 7-23-01 ADDED CREDIT 8-7-01
                ' payment on backorder with a 0 balance left
                Controll = "0.00"
            End If

            UndSls = "0.00"
            DelSls = "0.00"
            TaxRec1 = "0.00"
            If Money <> 0 Then Audit() 'puts a o entry in Audit
            If Money > 0 Then Holding.LastPay = DateFormat(TransDate)
        End If

        If PayTypeIsIn(cboAccount.SelectedText, cdsPayTypes.cdsPT_MiscDiscount) Then  'Discount On Sale  07-10-01  This one is being used
            LeaseNo = BillOSale.BillOfSale.Text
            Name1 = "CA " + BillOSale.CustomerLast.Text
            Written = -txtAmount.Text
            TaxCharged1 = "0.00"
            ArCashSls = -txtAmount.Text
            Controll = "0.00"
            UndSls = -txtAmount.Text
            DelSls = "0.00"
            TaxRec1 = "0.00"
            If MailCheck.OrigStatus = "B" Then DelSls = -txtAmount.Text
            Audit()
            'BFH20080112 - Don't change deposit in holding table, change the sale (like the audit already does)
            '    Holding.Deposit = Holding.Deposit + Deposit
            Holding.Sale = Holding.Sale - Deposit
            If Deposit > 0 Then Holding.LastPay = DateFormat(TransDate)
        End If

        TotDeposit = Holding.Deposit

        Exit Sub

HandleErr:
        If Err.Number = 13 Then
            'MsgBox "You have entered an incorrect payment!", vbExclamation
            MessageBox.Show("You have entered an incorrect payment!")
            txtAmount.Text = ""
            BillOSale.X = X
            BillOSale.SetPrice(X, "")
            BillOSale.X = X - 1
            BillOSale.SetPrice(X - 1, "")

            On Error GoTo 0
            X = X - 2
            '    txtAmount.SetFocus
            Err.Clear()
            Exit Sub

            Resume
        End If
    End Sub

    Public Property Receipt() As Boolean
        Get
            Receipt = (chkReceipt.Checked = True)
        End Get
        Set(value As Boolean)
            chkReceipt.Checked = IIf(value, True, False)
        End Set
    End Property

    Private Sub MakeMyReceipt()
        MakeReceipt(
    TransDate, eReceiptTypes.ert_SaleNo, BillOSale.BillOfSale.Text,
    BillOSale.CustomerFirst.Text, BillOSale.CustomerLast.Text, BillOSale.CustomerAddress.Text, BillOSale.CustomerAddress2.Text, BillOSale.CustomerCity.Text, BillOSale.CustomerZip.Text,
    PriorBal, PayMethod, Deposit, BillOSale.BalDue.Text, Memo.Text, Approval)
    End Sub

    Public Property Email() As Boolean
        Get
            Email = (chkEmail.Checked = True)
        End Get
        Set(value As Boolean)
            chkEmail.Checked = IIf(value, True, False)
        End Set
    End Property

    Private Sub MakeEmail()
        BillOSale.Recalculate()
        'BillOSale.cmdEmail.Value = True
        BillOSale.cmdEmail.PerformClick()
    End Sub

    Private Sub DoControls(ByVal Enabled As Boolean)
        cmdOk.Enabled = Enabled
        cmdCancel.Enabled = Enabled
        chkEmail.Enabled = Enabled
        chkReceipt.Enabled = Enabled
    End Sub

    Public Sub Cash()
        Dim tD As String
        If OrderMode("B", "D") And IsDate(TransDate) Then  'deliver sale & Payments
            tD = DateFormat(TransDate)
        Else
            tD = DateFormat(Now)
        End If

        AddNewCashJournalRecord(Trim(Account), GetPrice(Money), Microsoft.VisualBasic.Left(LeaseNo, 8), Microsoft.VisualBasic.Left(Note, 24) & " " & Memo.Text, tD, Trim(Cashier))
        Exit Sub

HandleErr:
        MessageBox.Show("ERROR in OrderPay: " & Err.Description & ", " & Err.Source & ", " & Err.Number)
        Err.Clear()
        Resume Next
    End Sub

    Private Sub OrdPay_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'SetButtonImage(cmdOk)
        'SetButtonImage(cmdCancel)
        SetButtonImage(cmdOk, 2)
        SetButtonImage(cmdCancel, 3)
        'SetCustomFrame(Me, ncBasicDialog)  -> Not required. It is just for formatting like colors, font etc.
        ColorDatePicker(dtePayDate)
        Left = 2800

        If OrderMode("D") Then          ' D (payment) changed to only current date BFH20100129
            dtePayDate.Value = Today
        ElseIf OrderMode("B") Then      ' B (deliver) added, bfh20061120
            dtePayDate.Value = IIf(TransDate = "", DateFormat(GetLastDeliveryDate), TransDate)
        Else
            dtePayDate.Value = Today
        End If
        TransDate = dtePayDate.Value

        X = MailCheck.X
        Top = IIf(X > 8, 1000, 3500)

        ' These lines refer to the global Holding object.
        ' They'll be reworked after I've researched all the side effects.
        If OrderMode("B") And g_Holding.Status <> "D" Then
            cmdCancel.Text = "Cancel Delivery"
            X = InvDel.X + 1
        End If
        If OrderMode("B") And g_Holding.Status = "D" Then   'after payment is made
            If g_Holding.Sale > g_Holding.Deposit Then
                cmdCancel.Enabled = False                            'on a delivered sale
            End If
        End If

        Setup()
    End Sub

    Private Sub Setup()
        DeliveredAuditRecord = 0
        DeliveredPayment = 0

        If Order >= "A" Then
            lblSaleTitle.Text = "Bill Of Sale:"
            lblSaleNo.Text = BillOSale.BillOfSale.Text
            lblName.Text = Trim(BillOSale.CustomerFirst.Text) & " " & Trim(BillOSale.CustomerLast.Text)
            lblAddress.Text = BillOSale.CustomerAddress.Text
            lblCity.Text = BillOSale.CustomerCity.Text & " " & BillOSale.CustomerZip.Text

            If OrderMode("B") Then
                dtePayDate.Value = InvDel.TransDate
                TransDate = dtePayDate.Value
            End If

            txtAmount.SelectionStart = 0
            If DoPayType(cdsPayTypes.cdsPT_Cash) Then AddAccountCode(PayListItem(cdsPayTypes.cdsPT_Cash), cdsPayTypes.cdsPT_Cash)
            cboAccount.SelectedIndex = 0
            If DoPayType(cdsPayTypes.cdsPT_Check) Then AddAccountCode(PayListItem(cdsPayTypes.cdsPT_Check), cdsPayTypes.cdsPT_Check)

            If DoPayType(cdsPayTypes.cdsPT_Visa) Then AddAccountCode(PayListItem(cdsPayTypes.cdsPT_Visa), cdsPayTypes.cdsPT_Visa)
            If DoPayType(cdsPayTypes.cdsPT_MCard) Then AddAccountCode(PayListItem(cdsPayTypes.cdsPT_MCard), cdsPayTypes.cdsPT_MCard)
            If DoPayType(cdsPayTypes.cdsPT_Disc) Then AddAccountCode(PayListItem(cdsPayTypes.cdsPT_Disc), cdsPayTypes.cdsPT_Disc)
            If DoPayType(cdsPayTypes.cdsPT_amex) Then AddAccountCode(PayListItem(cdsPayTypes.cdsPT_amex), cdsPayTypes.cdsPT_amex)
            If DoPayType(cdsPayTypes.cdsPT_BackOrder) Then AddAccountCode(PayListItem(cdsPayTypes.cdsPT_BackOrder), cdsPayTypes.cdsPT_BackOrder)
            If DoPayType(cdsPayTypes.cdsPT_OutsideFinance) Then AddAccountCode(PayListItem(cdsPayTypes.cdsPT_OutsideFinance), cdsPayTypes.cdsPT_OutsideFinance)
            If DoPayType(cdsPayTypes.cdsPT_OutsideFinance2) Then AddAccountCode(PayListItem(cdsPayTypes.cdsPT_OutsideFinance2), cdsPayTypes.cdsPT_OutsideFinance2)
            If DoPayType(cdsPayTypes.cdsPT_OutsideFinance3) Then AddAccountCode(PayListItem(cdsPayTypes.cdsPT_OutsideFinance3), cdsPayTypes.cdsPT_OutsideFinance3)
            If DoPayType(cdsPayTypes.cdsPT_OutsideFinance4) Then AddAccountCode(PayListItem(cdsPayTypes.cdsPT_OutsideFinance4), cdsPayTypes.cdsPT_OutsideFinance4)
            If DoPayType(cdsPayTypes.cdsPT_OutsideFinance5) Then AddAccountCode(PayListItem(cdsPayTypes.cdsPT_OutsideFinance5), cdsPayTypes.cdsPT_OutsideFinance5)
            If DoPayType(cdsPayTypes.cdsPT_DebitCard) Then AddAccountCode(PayListItem(cdsPayTypes.cdsPT_DebitCard), cdsPayTypes.cdsPT_DebitCard)
            If DoPayType(cdsPayTypes.cdsPT_CompanyCheck) Then AddAccountCode(PayListItem(cdsPayTypes.cdsPT_CompanyCheck), cdsPayTypes.cdsPT_CompanyCheck)

            If Order <> "B" Then  'deliver sale audit doens't work
                If DoPayType(cdsPayTypes.cdsPT_MiscDiscount) Then AddAccountCode(PayListItem(cdsPayTypes.cdsPT_MiscDiscount), cdsPayTypes.cdsPT_MiscDiscount)
            End If

            If Installment And OrderMode("B") Then
                If MailCheck.OrigStatus <> "F" Then
                    If DoPayType(cdsPayTypes.cdsPT_StoreFinance) Then AddAccountCode(PayListItem(cdsPayTypes.cdsPT_StoreFinance), cdsPayTypes.cdsPT_StoreFinance)
                End If
            End If

            If Installment And OrderMode("D") Then
                If DoPayType(cdsPayTypes.cdsPT_StoreFinance) Then AddAccountCode(PayListItem(cdsPayTypes.cdsPT_StoreFinance), cdsPayTypes.cdsPT_StoreFinance)
            End If

            If DoPayType(cdsPayTypes.cdsPT_StoreCreditCard) Then AddAccountCode(PayListItem(cdsPayTypes.cdsPT_StoreCreditCard), cdsPayTypes.cdsPT_StoreCreditCard)
            If DoPayType(cdsPayTypes.cdsPT_ECheck) Then AddAccountCode(PayListItem(cdsPayTypes.cdsPT_ECheck), cdsPayTypes.cdsPT_ECheck)

        End If

        If BillOSale.IsGridFull Then
            MessageBox.Show("The bill of sale is near the maximum number of lines." & vbCrLf &
             "You need to either deliver or void this sale.", "Maximum Lines Approaching", MessageBoxButtons.OK, MessageBoxIcon.Information)
            DoControls(True)
            '      chkPayAll.value = 1
            '      chkPayAll.Enabled = False
            '      txtAmount.Enabled = False
        End If
    End Sub

    Public Function Audit() As Integer
        SalesJournal_AddRecordNew_Data(
    BillOSale.BillOfSale.Text, Name1, TransDate, Written, TaxCharged1,
    ArCashSls, Controll, UndSls, DelSls, TaxRec1,
    IIf(MailCheck.TaxCode = 0, 1, MailCheck.TaxCode), MailCheck.SalesPerson, 0, Cashier)
        Audit = LastAuditID()
    End Function

    Private Sub AddAccountCode(ByVal Code As String, ByVal AccountVal As cdsPayTypes)
        'cboAccount.AddItem Code
        'cboAccount.itemData(cboAccount.NewIndex) = AccountVal
        cboAccount.Items.Add(New ItemDataClass(Code, AccountVal))
    End Sub

End Class