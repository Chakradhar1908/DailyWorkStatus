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
        Dim X As Long, GM As CGrossMargin, objHolding As cHolding
        ' Prepare the Holding object to accept info.

        FinanceArNo = ArNo
        PaymentOnAccount
        FinanceArNo = ""
  Set objHolding = New cHolding
  objHolding.Load g_Holding.LeaseNo  ' Load the most current info.
        PostPaymentOnAccount objHolding
  objHolding.ArNo = ArNo ' this would work if the code came this direction!
        objHolding.Save()
        FinishRoutine False

  DisposeDA objHolding
End Sub

    Public Sub FinishRoutine(ByVal StayOnOrder As Boolean)
        ' This function uses the global Holding object.
        ' It will be reworked once I've researched all the side effects.

        If Receipt Then MakeMyReceipt : Receipt = False
        If Email Then MakeEmail : Email = False

        If StayOnOrder Then
            ' Clear temporary stuff..
            DoControls True
    BillOSale.cmdMainMenu.Enabled = True
            cboAccount.ListIndex = 0
            txtAmount.Text = ""
            Exit Sub
        End If

        DeliveredAuditRecord = 0
        DeliveredPayment = 0

        If OrderMode("B") Then
            If MsgBox("Any More To Deliver?", vbQuestion + vbYesNo) = vbYes Then
                Unload OrdPay
      Unload ArCard
      Unload ARPaySetUp
      Unload BillOSale
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
                MailCheck.optSaleNo.Value = True
                MailCheck.Show vbModal, BillOSale
      Exit Sub
            End If

            ' No more to deliver
            Unload BillOSale
    Unload ARPaySetUp
    Unload ArCard ' need for add on
            Unload AddOnAcc
    Set AddOnAcc = Nothing
    Unload ArCard ' Why twice?  Does unloading AddOnAcc reload this?

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
            Unload OrdPay
    Exit Sub
        End If

        If OrderMode("D") Then
            BillOSale.UGridIO1.GetDBGrid.Refresh() 'bfh20060113 - refresh & doevents added for cosmetic fix
            DoEvents
            If MsgBox("Any More To Pay On?", vbQuestion + vbYesNo) = vbYes Then
                PayMethod = ""
                BillOSale.cmdMainMenu.Enabled = True
                Unload OrdPay
      BillOSale.UGridIO1.Clear()
                MailCheck.FirstRec = 0
                TotDeposit = 0
                Deposit = 0
                'frmSalesList.SafeSalesClear = True
                frmSalesList.SalesCode = ""
                ' This allows retention of current customer for consecutive payments.
                ' Some customers may not want this, so they should go in the else block, using the old code, upon request
                If True Then
                    MailCheck.optTelephone.Value = True
                    MailCheck.InputBox.Text = DressAni(CleanAni(BillOSale.CustomerPhone1))
                Else
                    MailCheck.optSaleNo.Value = True
                End If
                MailCheck.Show vbModal, BillOSale
      Exit Sub
            Else
                PayMethod = ""
                Unload OrdPay
      Unload BillOSale
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
End Class