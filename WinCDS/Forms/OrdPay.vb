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
    Dim OrdPayCancelButtonSelected As Boolean

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
        'Left = 2800

        If OrderMode("D") Then          ' D (payment) changed to only current date BFH20100129
            dtePayDate.Value = Today
        ElseIf OrderMode("B") Then      ' B (deliver) added, bfh20061120
            dtePayDate.Value = IIf(TransDate = "", DateFormat(GetLastDeliveryDate), TransDate)
        Else
            dtePayDate.Value = Today
        End If
        TransDate = dtePayDate.Value

        X = MailCheck.X
        'Top = IIf(X > 8, 1000, 3500)

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

    Private Sub chkPayAll_Click(sender As Object, e As EventArgs) Handles chkPayAll.Click
        If chkPayAll.Checked = True And PayMethod <> PayListItem(cdsPayTypes.cdsPT_StoreFinance) Then
            txtAmount.Text = BillOSale.BalDue.Text
        Else
            txtAmount.Text = ""
        End If
    End Sub

    'NOTE: THIS CODE IS COMMENTED, CAUSE CLICK EVENT OF COMBOBOX WILL NOT EXECUTE LIKE IN VB6.0. IN VB.NET, SELECTEDINDEXCHANGED WILL WORK.
    'Private Sub cboAccount_Click(sender As Object, e As EventArgs) Handles cboAccount.Click
    '    '  cboAccount.List(cboAccount.ListIndex) = left(cboAccount.List(cboAccount.ListIndex), 14)
    '    'PayMethod = Trim(Left(cboAccount.List(cboAccount.ListIndex), 14))
    '    PayMethod = Trim(Microsoft.VisualBasic.Left(cboAccount.Items(cboAccount.SelectedIndex).ToString, 14))
    '    If PayTypeIsFinance(PayMethod) And Not AllowPartialFinancing Then
    '        txtAmount.Text = ""
    '        txtAmount.Enabled = False
    '    Else
    '        txtAmount.Enabled = True
    '    End If
    'End Sub

    Private Sub cmdChangeDate_Click(sender As Object, e As EventArgs) Handles cmdChangeDate.Click
        If Not dtePayDate.Enabled Then
            If Not RequestManagerApproval("Change Payment Dates") Then
                MessageBox.Show("You do not have access to change the payment date.", "Permission Denied", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                Exit Sub ' Match passwords or fail.
            End If
        End If
        dtePayDate.Enabled = Not dtePayDate.Enabled
        If dtePayDate.Enabled Then
            cmdChangeDate.Text = "Lock"
        Else
            cmdChangeDate.Text = "Change"
        End If
    End Sub

    Private Sub cmdOk_Click(sender As Object, e As EventArgs) Handles cmdOk.Click
        LockOn = False
        ' OK button
        ' The functions under this repeatedly load and unload the Holding record by LeaseNo.
        ' I'd like to load it once and save it after.
        ' Lucky for us, the global Holding is already defined and opened to the current record.
        ' Unfortunately, we don't know what side effects will be caused by updating it.
        ' So we're going to create a new instance, load the same record, and work with that.
        ' It'll be saved and updated at the end of this function.

        ' This form is loaded by "Payment On Account" and "Deliver Sale" menus.
        ' The significant difference in processing is "Deliver Sale" requires the
        ' entire balance to be accounted for, resulting in the possibility of many
        ' payments per load of the form.

        Dim objHolding As New cHolding
        Dim StayOnOrder As Boolean, TPM As String
        TPM = Trim(PayMethod)

        If dtePayDate.Value = NullDate Then
            If MessageBox.Show("Did you mean to make this payment on " & NullDate & "?", "", MessageBoxButtons.YesNo) = DialogResult.No Then
                MessageBox.Show("It appears you have found a bug in the software!  Please contact " & AdminContactString(Format:=1, Phone:=False) & " to have this issue resolved.", "Ooops!")
                Exit Sub  ' Bad initialization check.
            End If
        End If

        ' Prepare the Holding object to accept info.
        objHolding.Load(g_Holding.LeaseNo)  ' Load the most current info.

        'added 03-20-2003 for pressing ok when balance due and no money entered
        If BillOSale.BalDue.Text <> 0 And Val(txtAmount.Text) = 0 And Not IsIn(objHolding.Status, "S", "F", "E") Then
            If Not PayTypeIsFinance(TPM) Then
                MessageBox.Show("There is still a balance due.  Please enter the payment amount or change payment type to Back Order, Outside Finance Company, or Store Finance.", "Cannot Finish -- Balance Still Due")
                DisposeDA(objHolding)
                Exit Sub
            End If

            'BFH20161026 - This may not be the place for this, but it works here.
            '              Added because we now allow partial payment by "Outside Finance", and not just
            '              the full payment amount.  This forces a zero amount to be read as "Pay All".
            '    If PayTypeIs(TPM) = cdsPT_OutsideFinance Then
            '      txtAmount = FormatCurrency(BillOSale.BalDue)
            '    End If
        End If



        ' BFH20070521
        If GetPrice(txtAmount.Text) > GetPrice(BillOSale.BalDue.Text) Then
            If MessageBox.Show("You entered a sale amount of " & txtAmount.Text & "." & vbCrLf & "This is more than the balance due of " & BillOSale.BalDue.Text & "." & vbCrLf2 & "Do you really want to OVERPAY the sale?", "Confirm Overpayment", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = DialogResult.No Then
                DisposeDA(objHolding)
                Exit Sub
            End If
        End If

        If objHolding.Status = "D" And objHolding.Sale - objHolding.Deposit = 0 And GetPrice(txtAmount.Text) <> 0 Then
            If MessageBox.Show("You are making a payment on a fully-paid, delivered sale." & vbCrLf & "This will change this sale to Back Ordered.", "Confirm Status Change", MessageBoxButtons.OKCancel) = DialogResult.Cancel Then
                Exit Sub
            Else
                objHolding.Status = "B"
                objHolding.Save()
            End If
        End If

        If chkPayAll.Checked = True And cboAccount.SelectedIndex < 0 Then
            MessageBox.Show("Please select a payment type.")
            DisposeDA(objHolding)
            Exit Sub
        End If

        If PayTypeIs(TPM) = cdsPayTypes.cdsPT_BackOrder Then txtAmount.Text = 0

        DoControls(False)
        frmCCAd.Advertize()
        BillOSale.cmdMainMenu.Enabled = False

        ' Major bug:
        '  When the order is Status F (Store Finance) and Order is B (Deliver Sale),
        '  we're not being allowed to close this form without paying it off in full. This is
        '  improper, the form should close without requiring payment.

        'xxxxxxxxxxxxx Payment On Sale xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
        If OrderMode("D") And IsIn(MailCheck.OrigStatus, "O", "L", "1", "2", "3", "4") Then
            'payment on account - open order
            If PayTypeIs(TPM) = cdsPayTypes.cdsPT_StoreFinance Then

                ' ************************************
                ' bfh20050815
                ' code snippet taken from ordselect's store finance payment option...
                If Val(BillOSale.Index) = 0 Then
                    MessageBox.Show("You cannot set up an Installment Contract without the Customer's Name & Address!", "No Name or Address")
                    DoControls(True)
                    Exit Sub
                End If
                If Len(Trim(BillOSale.CustomerPhone1.Text)) < 1 Then
                    MessageBox.Show("You cannot set up an Installment Contract without the Customer's Telephone Number.", "No Telephone Number")
                    DoControls(True)
                    Exit Sub
                End If

                '        ArStatus = "C"  'sets status to call lease no
                'Unload ARPaySetUp
                ARPaySetUp.Close()
                ARPaySetUp.Show()
                DisposeDA(objHolding)
                Exit Sub
                ' ************************************
            Else
                If Not PaymentOnAccount() Then DoControls(True) : Exit Sub
                PostPaymentOnAccount(objHolding)
            End If
            Application.DoEvents()
        ElseIf OrderMode("D") Or IsIn(MailCheck.OrigStatus, "E", "C", "B", "S", "F") Then
            Dim OrgDeposit As Decimal
            'changed 07-10-01 lets you deliver the balance of the items and pay off sale!
            If IsIn(MailCheck.OrigStatus, "E", "C", "B", "S", "F") Then
                If Val(BillOSale.BalDue) <> 0 Then
                    ' BFH20161101 - Change Orig Deposit, because we could have multiple things in here...
                    '        OrgDeposit = BillOSale.BalDue
                    OrgDeposit = BillOSale.SaleTotal("paid")
                End If
            End If

            'pay off back order, finance, store finance
            If GetPrice(txtAmount.Text) <> 0 Then
                If Not PaymentOnAccount() Then DoControls(True) : Exit Sub
            End If

            'BFH20150410 We were running into a situtation where the recalculated balance did not match the
            ' displayed balance (from Holding). We tried changing the sale, but that threw accounting off.
            ' Instead, we perform a check here.  If it is displaying a virtually 0.00 Balance Due, but
            ' the Recalculate has another figure (often hundreds of dollars), we "deliver" the sale, and
            ' let the system process it.  The numbers are technically wrong on the sale, as they are missing
            ' an adjustment somewhere, but this was deemed a better fix, simply by not having these errors
            ' show up.  I suppose, if someone were to do the adding up on the sale, and catch this, we
            ' may have to revisit this again, unless we find the problem in the adjustment process, as the
            ' balance error can be either in the customer's or store's benefit.  For now, this will
            ' have to do.
            Dim BalanceBefore As Decimal
            BalanceBefore = GetPrice(BillOSale.BalDue.Text)
            ' bfh20050929...   made recalculate here not refigure taxes
            BillOSale.Recalculate(True)
            If Math.Abs(BalanceBefore) <= 0.03 And Math.Abs(GetPrice(BillOSale.BalDue.Text)) > 0.03 Then
                BillOSale.BalDue.Text = CurrencyFormat(BalanceBefore)
            End If
            If True Then
                If Math.Abs(Val(BillOSale.BalDue)) <= 0.03 Then
                    OrgHoldingStatus = "D"
                    objHolding.Status = "D"
                End If
            Else
                If Val(BillOSale.BalDue) = 0# Then
                    OrgHoldingStatus = "D"
                    objHolding.Status = "D"
                End If
            End If

            PostPaymentOnAccount(objHolding)

            If IsIn(MailCheck.OrigStatus, "E", "C", "B", "S", "F") Then
                Account = IIf(MailCheck.OrigStatus = "B", "11200", "11300")
                If MailCheck.OrigStatus = "B" Then
                    Money = -GetPrice(txtAmount.Text)
                Else
                    If GetPrice(txtAmount.Text) <> 0 And Val(BillOSale.BalDue) = 0 Then 'pay off backorder
                        Money = -OrgDeposit
                    Else
                        Money = -GetPrice(txtAmount.Text) ' bfh20051117 - partial payment of back order...
                    End If
                End If
                Cash()
            ElseIf MailCheck.OrigStatus = "D" And objHolding.Status = "B" Then
                '' The odd case where they put a payment on a fully paid/delivered sale..
                '' Creates a back-order.  This adds the 11200 entry into [Cash]
                Account = "11200"
                Money = -GetPrice(txtAmount.Text)
                Cash()
            End If

            'xxxxxxxxxxxxx Deliver Sale xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
        ElseIf OrderMode("B") And IsIn(Status, "", "T", "TT", "TTT") Then
            'deliver sale pay off in full
            '    If PayTypeIsIn(TPM, cdsPT_StoreFinance, cdsPT_OutsideFinance) And GetPrice(txtAmount) = 0 Then txtAmount = BillOSale.BalDue
            If Status = "" Then
                If Not PaymentOnAccount() Then DoControls(True) : Exit Sub 'prepaid order before delivery
                DeliverSale(objHolding)
            Else
                If Not PaymentOnAccount() Then DoControls(True) : Exit Sub
                BackOrderBal(objHolding)
            End If
            objHolding.Save() ' we save later too, but there's a possible Exit Sub in the middle here (see below)
            Application.DoEvents()

            If PayTypeIsFinance(TPM) Then
                If Status = "" Then
                    'backorder the entire balance
                    LeaseNo = BillOSale.BillOfSale.Text
                    Money = BillOSale.BalDue.Text
                    Account = IIf(PayTypeIs(TPM) = cdsPayTypes.cdsPT_BackOrder, "11200", "11300") 'BFH20080710 - SF changed to 11300 from 11200
                    Note = BillOSale.CustomerLast.Text

                    If PayTypeIsOutsideFinance(cboAccount.Text) And GetPrice(txtAmount.Text) <> GetPrice(BillOSale.BalDue.Text) And GetPrice(txtAmount.Text) <> 0 Then
                        '
                    Else
                        Cash()
                    End If
                End If

                If Installment And OrderMode("B") And MailCheck.OrigStatus = "O" And TPM = PayListItem(cdsPayTypes.cdsPT_StoreFinance) Then
                    ARPaySetUp.Show() 'vbModal, BillOSale
                    'Unload Me
                    Me.Close()
                    BillOSale.cmdMainMenu.Enabled = True
                    Exit Sub
                    ' If ArPaySetUp is cancelled, unselect the combo and remove the GM line.
                End If
            End If

            If (Not PayTypeIsFinance(TPM)) Or (PayTypeIsOutsideFinance(TPM) And (GetPrice(txtAmount.Text) <> GetPrice(BillOSale.BalDue.Text)) And (GetPrice(txtAmount.Text) <> 0)) Then
                If Val(BillOSale.BalDue.Text) > 0 Then
                    ' This happens when the current payment is less than the total owed.
                    MessageBox.Show("Order Not Complete!  You must Pay Off, Backorder or Finance Balance!")
                    X = X + 1
                    If Status = "" Then Status = "T" Else Status = "TT" 'temp
                    ' Abnormal exit - force continuation of order.
                    StayOnOrder = True
                ElseIf Val(BillOSale.BalDue.Text) < 0 Then
                    ' This happens when the current payment is less than the total owed.
                    MessageBox.Show("NOTE: You have overpaid on this order." & vbCrLf & "You must leave a zero balance.", "Notice")
                    X = X + 1
                    Status = "TTT"
                    '        If Status = "" Then Status = "TTT" Else Status = "TT" 'temp
                    ' Abnormal exit - force continuation of order.
                    StayOnOrder = True
                End If
            End If
        End If
        objHolding.Save()
        FinishRoutine(StayOnOrder)
        DisposeDA(objHolding)
    End Sub

    Private Sub BackOrderBal(ByRef Holding As cHolding)
        'NOT backorder, finance, or store finance
        'If Not IsIn(CStr(cboAccount.itemData(cboAccount.ListIndex)), "7", "8", "11") Then
        If Not IsIn(CType(cboAccount.Items(cboAccount.SelectedIndex), ItemDataClass).ItemData, "7", "8", "11") Then
            '  If cboAccount.ItemData(cboAccount.ListIndex) <> 7 And cboAccount.ItemData(cboAccount.ListIndex) <> 8 And cboAccount.ItemData(cboAccount.ListIndex) <> 11 Then
            LeaseNo = BillOSale.BillOfSale.Text
            Money = Deposit
            'Account = cboAccount.itemData(cboAccount.ListIndex) ' Val(cboAccount.ListIndex) + 1 ' BFH20151014
            Account = CType(cboAccount.Items(cboAccount.SelectedIndex), ItemDataClass).ItemData  ' Val(cboAccount.ListIndex) + 1 ' BFH20151014
            Note = BillOSale.CustomerLast.Text
            LeaseNo = BillOSale.BillOfSale.Text
            Name1 = "PA " + BillOSale.CustomerLast.Text
            Written = "0.00"
            TaxCharged1 = "0.00"
            ArCashSls = "0.00"
            Controll = -txtAmount.Text
            UndSls = "0.00"
            DelSls = "0.00"
            TaxRec1 = "0.00"
        Else
            '  End If
            '
            '  If cboAccount.ItemData(cboAccount.ListIndex) = 7 Or cboAccount.ItemData(cboAccount.ListIndex) = 8 Or cboAccount.ItemData(cboAccount.ListIndex) = 11 Then
            LeaseNo = BillOSale.BillOfSale.Text
            Money = BillOSale.BalDue.Text
            Account = IIf(cboAccount.Text = "BACK ORDER", "11200", "11300")
            Note = BillOSale.CustomerLast.Text

            Name1 = "BO " + BillOSale.CustomerLast.Text
            Written = "0.00"
            TaxCharged1 = "0.00"
            ArCashSls = "0.00"

            If Status = "T" Then 'if second payment is made before backorder
                Controll = "0.00"
            ElseIf Status = "TT" Then
                Controll = -Controll 'if second payment is made before backorder
            End If

            UndSls = "0.00"
            DelSls = "0.00"
            InvDel.TaxRec2 = "0.00"
            InvDel.TaxRec1 = "0.00"
            TaxRec1 = "0.00"
        End If

        Cashier = GetCashierName
        Cash()

        If Not IsIn(Status, "T", "TT") Then Audit()

        'If Trim(Left(cboAccount.List(cboAccount.ListIndex), 14)) = "OUTSIDE FIN CO" Then
        If Trim(Microsoft.VisualBasic.Left(cboAccount.Items(cboAccount.SelectedIndex).ToString, 14)) = "OUTSIDE FIN CO" Then
            Status = "C" : Holding.Status = "C" : BillOSale.SaleStatus.Text = "Finance Co"
            'ElseIf Trim(Left(cboAccount.List(cboAccount.ListIndex), 10)) = "BACK ORDER" Then
        ElseIf Trim(Microsoft.VisualBasic.Left(cboAccount.Items(cboAccount.SelectedIndex).ToString, 10)) = "BACK ORDER" Then
            Status = "B" : Holding.Status = "B" : BillOSale.SaleStatus.Text = "Back Order"
            'ElseIf Trim(Left(cboAccount.List(cboAccount.ListIndex), 13)) = "STORE FINANCE" Then
        ElseIf Trim(Microsoft.VisualBasic.Left(cboAccount.Items(cboAccount.SelectedIndex).ToString, 13)) = "STORE FINANCE" Then
            Status = "F" : Holding.Status = "F" : BillOSale.SaleStatus.Text = "Store Finance"
        End If

        Holding.Deposit = Holding.Deposit + Deposit
        TotDeposit = Holding.Deposit
    End Sub

    Public ReadOnly Property AllowPartialFinancing() As Boolean
        Get
            AllowPartialFinancing = False
            AllowPartialFinancing = AllowPartialFinancing Or IsDevelopment()    ' DEV MODE
            AllowPartialFinancing = AllowPartialFinancing Or IsPitUSA
        End Get
    End Property

    Private Sub cmdCancel_Click(sender As Object, e As EventArgs) Handles cmdCancel.Click
        Dim objHolding As cHolding
        LockOn = False

        ' Cancel
        If OrderMode("B") Then
            If DeliveredAuditRecord <> 0 Then
                'BFH20150620 - This line was missing a "space" after the "PA" portion.
                ExecuteRecordsetBySQL("UPDATE Audit SET Name1='PA '+Mid(Name1,3), ArCashSls=0, UndSls=0, DelSls=0, TaxRec1=0, Controll=" & -DeliveredPayment & " WHERE AuditID=" & DeliveredAuditRecord, , GetDatabaseAtLocation)
                DeliveredAuditRecord = 0
                DeliveredPayment = 0
                objHolding = New cHolding
                objHolding.Load(g_Holding.LeaseNo)  ' Load the most current info.
                objHolding.Status = OrgHoldingStatus
                objHolding.Save()
                DisposeDA(objHolding)
            End If
            Receipt = False
            FinishRoutine(False)
            Exit Sub
        End If

        If MessageBox.Show("Any More To Pay On?", "WinCDS", MessageBoxButtons.YesNo) = DialogResult.Yes Then
            'Unload OrdPay
            Close()
            '    Unload BillOSale
            MailCheck.FirstRec = 0
            TotDeposit = 0
            Deposit = 0
            'frmSalesList.SafeSalesClear = True
            frmSalesList.SalesCode = ""
            MailCheck.optSaleNo.Checked = True
            'MailCheck.Show vbModal
            MailCheck.ShowDialog()
            Exit Sub

        Else
            TransDate = ""
            'Unload OrdPay
            Close()
            '    Unload BillOSale
            'Unload BillOSale
            BillOSale.Close()
            MainMenu.Show()
            MailCheck.FirstRec = 0
            TotDeposit = 0
            Deposit = 0
            frmSalesList.SalesCode = ""
            MailCheck.InputBox.Text = ""
            OrdPayCancelButtonSelected = True
            Exit Sub
        End If

        'Unload OrdPay
        Close()
        '  Unload BillOSale
        'Unload BillOSale
        BillOSale.Close()
        MainMenu.Show()
        MailCheck.FirstRec = 0
        TotDeposit = 0
        Deposit = 0
        frmSalesList.SalesCode = ""
    End Sub

    Private Sub OrdPay_Activated(sender As Object, e As EventArgs) Handles MyBase.Activated
        If Not OrderMode("B", "D") Then Exit Sub
        If Not IsFormLoaded("ARPaySetup") Then LockOn = True
        If IsIn(HoldingStatusRepresents(BillOSale.SaleStatus.Text), "E", "C", "S", "F") Then TakeNoPayment
    End Sub

    Public Function TakeNoPayment(Optional ByVal Prevent As Boolean = True) As Boolean
        Const H = 1455
        On Error Resume Next
        txtNoPay.Visible = Prevent
        'txtNoPay.ZOrder 0
        txtNoPay.BringToFront()
    End Function

    Private Sub OrdPay_Deactivate(sender As Object, e As EventArgs) Handles MyBase.Deactivate
        If Not IsFormLoaded("ARPaySetup") Then
            If LockOn Then
                tmrLockOn.Enabled = False
                tmrLockOn.Interval = 100
                tmrLockOn.Enabled = True
                'On Error Resume Next
                '      BillOSale.Show
                '      BillOSale.BillOSale2_Show
                '      Show
            End If
        End If
    End Sub

    Private Sub tmrLockOn_Tick(sender As Object, e As EventArgs) Handles tmrLockOn.Tick
        'Debug.Print "OrdPay::tmrLockOn_Timer"
        tmrLockOn.Enabled = False
        On Error Resume Next

        If OrdPayCancelButtonSelected = True Then Exit Sub
        If Not IsFormLoaded("BillOSale") Then
            BillOSale.Show()
            BillOSale.BillOSale2_Show()
        End If
        Show()
    End Sub

    Private Sub DeliverSale(ByRef Holding As cHolding)
        Dim MD As Decimal, TR As Decimal
        cmdCancel.Enabled = False   'delivery mode.  if you make first payment you cannot cancel

        LeaseNo = BillOSale.BillOfSale.Text
        Money = Deposit

        'If cboAccount.ListIndex < 0 Then cboAccount.ListIndex = 0
        If cboAccount.SelectedIndex < 0 Then cboAccount.SelectedIndex = 0
        'Account = cboAccount.itemData(cboAccount.ListIndex)
        Account = CType(cboAccount.Items(cboAccount.SelectedIndex), ItemDataClass).ItemData
        Note = BillOSale.CustomerLast.Text
        Cashier = GetCashierName

        '  If PayTypeIs(cboAccount.Text) = cdsPT_OutsideFinance And GetPrice(txtAmount) <> GetPrice(BillOSale.BalDue) Then
        '    '
        '  Else
        Cash()
        '  End If

        On Error GoTo HandleErr
        OrgHoldingStatus = Holding.Status
        'Used when creating a back order the first time!

        Sale = GetPrice(Holding.Sale)


        Holding.Status = "D"
        If PayTypeIs(PayMethod) = cdsPayTypes.cdsPT_OutsideFinance3 Then
            If BillOSale.BalDue.Text = 0 Or GetPrice(txtAmount.Text) = 0 Then   ' If the full price is not paid, do not change entire sale status!!
                Holding.Status = "C"
            End If
        ElseIf PayTypeIs(PayMethod) = cdsPayTypes.cdsPT_BackOrder Then
            Holding.Status = "B"
        ElseIf PayTypeIs(PayMethod) = cdsPayTypes.cdsPT_StoreFinance Then
            Holding.Status = "F"
        End If

        ' Balance due prior to delivery
        MD = InvDel.MiscDisc
        If MD <> 0 Then
            If InvDel.TaxRec1 <> 0 Then
                TR = MD * GetStoreTax1()
            Else
                TR = MD * Val(QuerySalesTax2(Val(InvDel.Tax2Zone) - 1))
            End If
        End If

        LeaseNo = BillOSale.BillOfSale.Text
        Name1 = "DS " & BillOSale.CustomerLast.Text
        Written = "0.00"
        TaxCharged1 = -TR
        ArCashSls = -Sale
        Controll = BillOSale.SaleTotal("paid") - GetPrice(txtAmount.Text) - BillOSale.SaleTotal(PayListItem(cdsPayTypes.cdsPT_OutsideFinance3)) ' Trim(Holding.Deposit)
        UndSls = -Sale
        DelSls = MailCheck.GrossSale - InvDel.TaxRec1 - InvDel.TaxRec2 + TR
        TaxRec1 = InvDel.TaxRec1 + InvDel.TaxRec2 - TR

        ' BFH20051013 - is this right??  it means no additional entries are added when
        ' doing a 'delivery' on a delivered sale..
        ' So, make a payment accidentily and use delivery to correct it leaves no record in
        ' audit of the 'fixing payment'...  but what are the side-effects of allowing it?
        ' --> probably, it should just make a new "PA" line
        If OrgHoldingStatus <> "D" Then 'already delivered  added 8-20-01
            '    If PayTypeIs(cboAccount.Text) = cdsPT_OutsideFinance And GetPrice(txtAmount) <> GetPrice(BillOSale.BalDue) Then
            '      '
            '    Else
            If MailCheck.OrigStatus <> "D" Then DeliveredAuditRecord = Audit()
            '    End If

            ' If it's a store finance and there is a partial outside finance, we need to create the back order here..
            ' (For the same reason we took it out of the Controll value above...)
            If OrderMode("B") And IsIn(Holding.Status, "F", "C", "D") And BillOSale.SaleTotal(PayListItem(cdsPayTypes.cdsPT_OutsideFinance3)) <> 0 Then
                AddNewCashJournalRecord("11300", BillOSale.SaleTotal(PayListItem(cdsPayTypes.cdsPT_OutsideFinance3)), BillOSale.BillOfSale.Text, "", dtePayDate.Value)
            End If

            DeliveredPayment = GetPrice(txtAmount.Text)

            ' bfh20051013 - note above, that the payment recorded was holding.deposit, not txtAmount...
            ' this means that if they paid ABOVE the amount, we should have an extra audit line to record
            ' this amount over what was still due
            '    If GetPrice(txtAmount) > Holding.Deposit Then

            If GetPrice(txtAmount.Text) <> (Holding.Sale - Holding.Deposit) Then
                Name1 = "PA " & BillOSale.CustomerLast.Text
                ArCashSls = 0
                UndSls = 0
                DelSls = 0
                TaxRec1 = 0
                Controll = 0
                If GetPrice(txtAmount.Text) > (Holding.Sale - Holding.Deposit) Then
                    Controll = -(GetPrice(txtAmount.Text) - (Holding.Sale - Holding.Deposit)) ' this should be the appropriate adjusted value
                    '      Else
                    '        Controll = -(GetPrice(Holding.Deposit) + GetPrice(txtAmount))
                End If

                If Controll <> 0 Then Audit()
            End If


            '    If GetPrice(txtAmount) > (Holding.Sale - Holding.Deposit) Then
            '      ' leaseno, written, taxcharged:  already set w/ appropriate values
            '      Name1 = "PA " & BillOSale.CustomerLast
            '      ArCashSls = 0
            '      UndSls = 0
            '      DelSls = 0
            '      TaxRec1 = 0
            '      Controll = -(GetPrice(txtAmount) - (Holding.Sale - Holding.Deposit)) ' this should be the appropriate adjusted value
            ''      Controll = -(GetPrice(txtAmount) - (Holding.Deposit)) ' this should be the appropriate adjusted value
            '      If Controll <> 0 Then
            '        Audit   ' create the second audit line for this overpayment on account... this should balance the audit report for overpayment and not change the DS line
            '      End If
            '    End If
        Else    ' second or subsequent time they've delivered this sale
            ' leaseno, written, taxcharged:  already set w/ appropriate values
            If GetPrice(txtAmount.Text) <> 0 Then
                Name1 = "PA " & BillOSale.CustomerLast.Text
                ArCashSls = 0
                UndSls = 0
                DelSls = 0
                TaxRec1 = 0
                Controll = -GetPrice(txtAmount.Text)    ' in the redeliver case, this should be the full amount paid
                Audit()   ' create the second audit line for this overpayment on account
            End If
        End If

        Holding.Deposit = Holding.Deposit + Deposit
        TotDeposit = Holding.Deposit

        MailCheck.Controll = "0"
        Exit Sub

HandleErr:

        ' not correct entry
        If Err.Number = 13 Then
            MessageBox.Show("You have entered an incorrect payment!")
            txtAmount.Text = ""
            BillOSale.X = X
            BillOSale.SetPrice(X, "")
            If BillOSale.X > 0 Then
                BillOSale.X = X - 1
                BillOSale.SetPrice(X - 1, "")
            End If

            Err.Clear()
            On Error GoTo 0
            X = X - 2
            If X < 0 Then X = 0
            txtAmount.Select()
            Exit Sub

            Resume
        End If
    End Sub

    Private Sub dtePayDate_CloseUp(sender As Object, e As EventArgs) Handles dtePayDate.CloseUp
        TransDate = dtePayDate.Value
    End Sub

    Private Sub txtAmount_Enter(sender As Object, e As EventArgs) Handles txtAmount.Enter
        SelectContents(txtAmount)
    End Sub

    Private Function AccountCode() As Integer
        'AccountCode = cboAccount.List(cboAccount.ListIndex)
        AccountCode = cboAccount.Items(cboAccount.SelectedIndex).ToString
    End Function

    Private Sub txtNoPay_DoubleClick(sender As Object, e As EventArgs) Handles txtNoPay.DoubleClick
        TakeNoPayment(False)
    End Sub

    Private Sub chkPayAll_CheckedChanged(sender As Object, e As EventArgs) Handles chkPayAll.CheckedChanged

    End Sub

    Private Sub cboAccount_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboAccount.SelectedIndexChanged
        '  cboAccount.List(cboAccount.ListIndex) = left(cboAccount.List(cboAccount.ListIndex), 14)
        'PayMethod = Trim(Left(cboAccount.List(cboAccount.ListIndex), 14))
        PayMethod = Trim(Microsoft.VisualBasic.Left(cboAccount.Items(cboAccount.SelectedIndex).ToString, 14))
        If PayTypeIsFinance(PayMethod) And Not AllowPartialFinancing Then
            txtAmount.Text = ""
            txtAmount.Enabled = False
        Else
            txtAmount.Enabled = True
        End If
    End Sub

    Private Sub txtAmount_Leave(sender As Object, e As EventArgs) Handles txtAmount.Leave
        txtAmount.Text = CurrencyFormat(GetPrice(txtAmount.Text))
    End Sub
End Class