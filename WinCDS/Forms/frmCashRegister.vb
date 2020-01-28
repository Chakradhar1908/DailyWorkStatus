Public Class frmCashRegister
    Public ReadOnly Property MailZip() As String
        Get
            MailZip = ""
            If MailIndex <> 0 Then
                Dim M As clsMailRec
                M = New clsMailRec
                If M.Load(frmCashRegisterAddress.MailIndex, "#Index") Then
                    MailZip = M.Zip
                End If
                DisposeDA(M)
            End If
        End Get
    End Property

    Public ReadOnly Property MailIndex() As Integer
        Get
            MailIndex = Val(lblCust.Tag)
        End Get
    End Property

    Public Sub BeginSale()
        Const GD As Boolean = True
        Dim F As Long
        ' Start a new sale.
        ' Prepare the form, set focus to the SKU entry box.
        frmCashRegister.HelpContextID = 42500

        'If gD Then MsgBox "frmCashReg:  " & F: F = F + 1
        cmdComm.Tag = "1"
        cmdComm.Value = True
        cboSalesList.Tag = ""               ' this will force the commissions person not to be carried over from sale to sale


        'If gD Then MsgBox "frmCashReg:  " & F: F = F + 1
        lblTotal.Caption = "0.00"
        lblTax.Caption = "0.00"
        lblTendered.Caption = "0.00"
        lblDue.Caption = "0.00"
        txtSku.Text = ""                    ' Clear any lingering SKUs.
        RunningTotal = 0                    ' Clear any lingering total..
        TaxableAmt = 0                      ' Clear any lingering taxes..
        fraSaleButtons.Visible = True       ' Show the sale buttons.
        SetReturnMode False                 ' Enter normal scan mode.
        SetNonTaxable False                 ' Default to taxable.
        SaleComplete = False                ' Default to non-complete.
        ReceiptPrinted = False              ' Clear receipt-printed flag.
        ShowButtons 0                       ' Show the charge/management buttons.
        Erase SaleItems                     ' Erase any item history we might remember.
        'If gD Then MsgBox "frmCashReg:  " & F: F = F + 1
        cmdPayment.Enabled = True
        cmdReturn.Enabled = True
        cmdDiscount.Enabled = False         ' Don't allow discounts until an item has been sold.
        cmdPrint.Enabled = False            ' Don't allow reprints until the sale is complete.
        cmdDone.Enabled = False             ' Don't allow sale completion with no sale.
        cmdCancelSale.Caption = "Cancel Sale"
        cmdCancelSale.ToolTipText = "Click to cancel the sale, discarding all purchase information."
        vsbReceipt.SmallChange = 1
        vsbReceipt.LargeChange = picReceiptContainer.ScaleHeight / picReceipt.TextHeight("X")
        LoadStoreLogo imgLogo, StoresSld, True  ' Load the store logo.
        'If gD Then MsgBox "frmCashReg:  " & F: F = F + 1
        picReceipt.Cls
        PrintReceiptHeader picReceipt       ' Print the receipt header.
        MoveReceipt picReceipt.CurrentY
  Show()                             ' Show the form.
        On Error Resume Next
        SetFocus
        txtSku.SetFocus                     ' And give focus to the SKU entry box.
        On Error GoTo 0

        ' If there's no printer set up, get one.
        'If gD Then MsgBox "frmCashReg:  " & F: F = F + 1
        CashRegisterPrinterSelector.SetSelectedPrinter CashRegisterPrinter
  If CashRegisterPrinterSelector.GetSelectedPrinter Is Nothing Then
            imgLogo.Visible = False
            CashRegisterPrinterSelector.Visible = True
            chkSavePrinter.Visible = True
        End If

        'If gD Then MsgBox "frmCashReg:  " & F: F = F + 1
        SetCustomer 0
  GotCust = False
    End Sub

End Class