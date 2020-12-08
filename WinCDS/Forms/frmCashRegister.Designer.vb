<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class frmCashRegister
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.lblCust = New System.Windows.Forms.Label()
        Me.cmdComm = New System.Windows.Forms.Button()
        Me.cboSalesList = New System.Windows.Forms.ComboBox()
        Me.lblTotal = New System.Windows.Forms.Label()
        Me.lblTax = New System.Windows.Forms.Label()
        Me.lblTendered = New System.Windows.Forms.Label()
        Me.lblDue = New System.Windows.Forms.Label()
        Me.fraSaleButtons = New System.Windows.Forms.GroupBox()
        Me.cmdPurchaseGiftCard = New System.Windows.Forms.Button()
        Me.cmdMainMenu = New System.Windows.Forms.Button()
        Me.cmdPayment = New System.Windows.Forms.Button()
        Me.cmdReturn = New System.Windows.Forms.Button()
        Me.cmdDiscount = New System.Windows.Forms.Button()
        Me.cmdCancelSale = New System.Windows.Forms.Button()
        Me.cmdPrint = New System.Windows.Forms.Button()
        Me.cmdDone = New System.Windows.Forms.Button()
        Me.vsbReceipt = New System.Windows.Forms.VScrollBar()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.txtSku = New System.Windows.Forms.TextBox()
        Me.cmdPayReturnToSale = New System.Windows.Forms.Button()
        Me.cmdPayGiftCard = New System.Windows.Forms.Button()
        Me.cmdPayStoreCard = New System.Windows.Forms.Button()
        Me.cmdPayDebit = New System.Windows.Forms.Button()
        Me.cmdPayCredit = New System.Windows.Forms.Button()
        Me.cmdPayCheck = New System.Windows.Forms.Button()
        Me.cmdPayCash = New System.Windows.Forms.Button()
        Me.lblEnterStyle = New System.Windows.Forms.Label()
        Me.cmdTax = New System.Windows.Forms.Button()
        Me.fraPaymentButtons = New System.Windows.Forms.GroupBox()
        Me.fraCust = New System.Windows.Forms.GroupBox()
        Me.fraEnterSku = New System.Windows.Forms.GroupBox()
        Me.cmdFND = New System.Windows.Forms.Button()
        Me.fraSaleTotals = New System.Windows.Forms.GroupBox()
        Me.cmdDev = New System.Windows.Forms.Button()
        Me.lblDueCaption = New System.Windows.Forms.Label()
        Me.lblTenderedCaption = New System.Windows.Forms.Label()
        Me.lblTotalCaption = New System.Windows.Forms.Label()
        Me.chkSavePrinter = New System.Windows.Forms.CheckBox()
        Me.pnlPicReceipt = New System.Windows.Forms.Panel()
        Me.picReceipt = New System.Windows.Forms.PictureBox()
        Me.picReceiptContainer = New System.Windows.Forms.PictureBox()
        Me.imgLogo = New System.Windows.Forms.PictureBox()
        Me.CashRegisterPrinterSelector = New WinCDS.PrinterSelector()
        Me.fraSaleButtons.SuspendLayout()
        Me.fraPaymentButtons.SuspendLayout()
        Me.fraCust.SuspendLayout()
        Me.fraEnterSku.SuspendLayout()
        Me.fraSaleTotals.SuspendLayout()
        Me.pnlPicReceipt.SuspendLayout()
        CType(Me.picReceipt, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.picReceiptContainer, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.imgLogo, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'lblCust
        '
        Me.lblCust.Location = New System.Drawing.Point(7, 18)
        Me.lblCust.Name = "lblCust"
        Me.lblCust.Size = New System.Drawing.Size(301, 20)
        Me.lblCust.TabIndex = 0
        Me.lblCust.Text = "lblCust"
        '
        'cmdComm
        '
        Me.cmdComm.Location = New System.Drawing.Point(269, 45)
        Me.cmdComm.Name = "cmdComm"
        Me.cmdComm.Size = New System.Drawing.Size(39, 26)
        Me.cmdComm.TabIndex = 1
        Me.cmdComm.Text = "&C"
        Me.ToolTip1.SetToolTip(Me.cmdComm, "Select salesman for commission.")
        Me.cmdComm.UseVisualStyleBackColor = True
        '
        'cboSalesList
        '
        Me.cboSalesList.FormattingEnabled = True
        Me.cboSalesList.Location = New System.Drawing.Point(9, 18)
        Me.cboSalesList.Name = "cboSalesList"
        Me.cboSalesList.Size = New System.Drawing.Size(299, 21)
        Me.cboSalesList.TabIndex = 2
        '
        'lblTotal
        '
        Me.lblTotal.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblTotal.Location = New System.Drawing.Point(168, 11)
        Me.lblTotal.Name = "lblTotal"
        Me.lblTotal.Size = New System.Drawing.Size(90, 20)
        Me.lblTotal.TabIndex = 3
        Me.lblTotal.Text = "0.00"
        Me.lblTotal.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'lblTax
        '
        Me.lblTax.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblTax.Location = New System.Drawing.Point(172, 39)
        Me.lblTax.Name = "lblTax"
        Me.lblTax.Size = New System.Drawing.Size(86, 20)
        Me.lblTax.TabIndex = 4
        Me.lblTax.Text = "0.00"
        Me.lblTax.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'lblTendered
        '
        Me.lblTendered.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblTendered.Location = New System.Drawing.Point(172, 66)
        Me.lblTendered.Name = "lblTendered"
        Me.lblTendered.Size = New System.Drawing.Size(86, 20)
        Me.lblTendered.TabIndex = 5
        Me.lblTendered.Text = "0.00"
        Me.lblTendered.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'lblDue
        '
        Me.lblDue.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblDue.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.lblDue.Location = New System.Drawing.Point(147, 90)
        Me.lblDue.Name = "lblDue"
        Me.lblDue.Size = New System.Drawing.Size(111, 24)
        Me.lblDue.TabIndex = 6
        Me.lblDue.Text = "0.00"
        Me.lblDue.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'fraSaleButtons
        '
        Me.fraSaleButtons.Controls.Add(Me.cmdPurchaseGiftCard)
        Me.fraSaleButtons.Controls.Add(Me.cmdMainMenu)
        Me.fraSaleButtons.Controls.Add(Me.cmdPayment)
        Me.fraSaleButtons.Controls.Add(Me.cmdReturn)
        Me.fraSaleButtons.Controls.Add(Me.cmdDiscount)
        Me.fraSaleButtons.Controls.Add(Me.cmdCancelSale)
        Me.fraSaleButtons.Controls.Add(Me.cmdPrint)
        Me.fraSaleButtons.Controls.Add(Me.cmdDone)
        Me.fraSaleButtons.Location = New System.Drawing.Point(12, 397)
        Me.fraSaleButtons.Name = "fraSaleButtons"
        Me.fraSaleButtons.Size = New System.Drawing.Size(612, 84)
        Me.fraSaleButtons.TabIndex = 8
        Me.fraSaleButtons.TabStop = False
        '
        'cmdPurchaseGiftCard
        '
        Me.cmdPurchaseGiftCard.BackColor = System.Drawing.SystemColors.Window
        Me.cmdPurchaseGiftCard.Image = Global.WinCDS.My.Resources.Resources.Picture
        Me.cmdPurchaseGiftCard.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.cmdPurchaseGiftCard.Location = New System.Drawing.Point(531, 13)
        Me.cmdPurchaseGiftCard.Name = "cmdPurchaseGiftCard"
        Me.cmdPurchaseGiftCard.Size = New System.Drawing.Size(75, 67)
        Me.cmdPurchaseGiftCard.TabIndex = 17
        Me.cmdPurchaseGiftCard.Text = "Purchase &Gift Card"
        Me.cmdPurchaseGiftCard.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.ToolTip1.SetToolTip(Me.cmdPurchaseGiftCard, "Click this button to add a gift card to this order.")
        Me.cmdPurchaseGiftCard.UseVisualStyleBackColor = False
        Me.cmdPurchaseGiftCard.Visible = False
        '
        'cmdMainMenu
        '
        Me.cmdMainMenu.BackColor = System.Drawing.SystemColors.Window
        Me.cmdMainMenu.Image = Global.WinCDS.My.Resources.Resources.menu
        Me.cmdMainMenu.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.cmdMainMenu.Location = New System.Drawing.Point(456, 13)
        Me.cmdMainMenu.Name = "cmdMainMenu"
        Me.cmdMainMenu.Size = New System.Drawing.Size(75, 67)
        Me.cmdMainMenu.TabIndex = 16
        Me.cmdMainMenu.Text = "&Main Menu"
        Me.cmdMainMenu.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.ToolTip1.SetToolTip(Me.cmdMainMenu, "Click to return to the WinCDS menu.")
        Me.cmdMainMenu.UseVisualStyleBackColor = False
        '
        'cmdPayment
        '
        Me.cmdPayment.BackColor = System.Drawing.SystemColors.Window
        Me.cmdPayment.Image = Global.WinCDS.My.Resources.Resources.cash_register_icon
        Me.cmdPayment.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.cmdPayment.Location = New System.Drawing.Point(6, 13)
        Me.cmdPayment.Name = "cmdPayment"
        Me.cmdPayment.Size = New System.Drawing.Size(75, 67)
        Me.cmdPayment.TabIndex = 9
        Me.cmdPayment.Text = "&Payment"
        Me.cmdPayment.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.ToolTip1.SetToolTip(Me.cmdPayment, "Click to select a payment method.")
        Me.cmdPayment.UseVisualStyleBackColor = False
        '
        'cmdReturn
        '
        Me.cmdReturn.BackColor = System.Drawing.SystemColors.Window
        Me.cmdReturn.Image = Global.WinCDS.My.Resources.Resources.return4
        Me.cmdReturn.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.cmdReturn.Location = New System.Drawing.Point(81, 13)
        Me.cmdReturn.Name = "cmdReturn"
        Me.cmdReturn.Size = New System.Drawing.Size(75, 67)
        Me.cmdReturn.TabIndex = 10
        Me.cmdReturn.Text = "&Return"
        Me.cmdReturn.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.cmdReturn.UseVisualStyleBackColor = False
        '
        'cmdDiscount
        '
        Me.cmdDiscount.BackColor = System.Drawing.SystemColors.Window
        Me.cmdDiscount.Image = Global.WinCDS.My.Resources.Resources.Icon_Specials
        Me.cmdDiscount.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.cmdDiscount.Location = New System.Drawing.Point(156, 13)
        Me.cmdDiscount.Name = "cmdDiscount"
        Me.cmdDiscount.Size = New System.Drawing.Size(75, 67)
        Me.cmdDiscount.TabIndex = 11
        Me.cmdDiscount.Text = "&Special"
        Me.cmdDiscount.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.ToolTip1.SetToolTip(Me.cmdDiscount, "Click to enter a discount on the last item purchased.")
        Me.cmdDiscount.UseVisualStyleBackColor = False
        '
        'cmdCancelSale
        '
        Me.cmdCancelSale.BackColor = System.Drawing.SystemColors.Window
        Me.cmdCancelSale.Image = Global.WinCDS.My.Resources.Resources.cancel
        Me.cmdCancelSale.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.cmdCancelSale.Location = New System.Drawing.Point(231, 13)
        Me.cmdCancelSale.Name = "cmdCancelSale"
        Me.cmdCancelSale.Size = New System.Drawing.Size(75, 67)
        Me.cmdCancelSale.TabIndex = 15
        Me.cmdCancelSale.Text = "&Cancel Sale"
        Me.cmdCancelSale.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.cmdCancelSale.UseVisualStyleBackColor = False
        '
        'cmdPrint
        '
        Me.cmdPrint.BackColor = System.Drawing.SystemColors.Control
        Me.cmdPrint.Image = Global.WinCDS.My.Resources.Resources.reprint
        Me.cmdPrint.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.cmdPrint.Location = New System.Drawing.Point(306, 13)
        Me.cmdPrint.Name = "cmdPrint"
        Me.cmdPrint.Size = New System.Drawing.Size(75, 67)
        Me.cmdPrint.TabIndex = 12
        Me.cmdPrint.Text = "Reprint Receip&t"
        Me.cmdPrint.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.ToolTip1.SetToolTip(Me.cmdPrint, "Click to print another copy of the receipt.  This function is available only for " &
        "completed sales.")
        Me.cmdPrint.UseVisualStyleBackColor = False
        '
        'cmdDone
        '
        Me.cmdDone.BackColor = System.Drawing.SystemColors.Window
        Me.cmdDone.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdDone.Image = Global.WinCDS.My.Resources.Resources.dollarsign
        Me.cmdDone.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.cmdDone.Location = New System.Drawing.Point(381, 13)
        Me.cmdDone.Name = "cmdDone"
        Me.cmdDone.Size = New System.Drawing.Size(75, 67)
        Me.cmdDone.TabIndex = 13
        Me.cmdDone.Text = "&Finish Sale"
        Me.cmdDone.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.ToolTip1.SetToolTip(Me.cmdDone, "Click to complete the sale and print a receipt.")
        Me.cmdDone.UseVisualStyleBackColor = False
        '
        'vsbReceipt
        '
        Me.vsbReceipt.Location = New System.Drawing.Point(705, 45)
        Me.vsbReceipt.Name = "vsbReceipt"
        Me.vsbReceipt.Size = New System.Drawing.Size(13, 379)
        Me.vsbReceipt.TabIndex = 16
        Me.vsbReceipt.Visible = False
        '
        'txtSku
        '
        Me.txtSku.Location = New System.Drawing.Point(9, 18)
        Me.txtSku.Name = "txtSku"
        Me.txtSku.Size = New System.Drawing.Size(299, 20)
        Me.txtSku.TabIndex = 22
        Me.ToolTip1.SetToolTip(Me.txtSku, "Type or scan the item's style number here.  You will be prompted for quantity and" &
        " price changes.")
        '
        'cmdPayReturnToSale
        '
        Me.cmdPayReturnToSale.BackColor = System.Drawing.SystemColors.Window
        Me.cmdPayReturnToSale.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdPayReturnToSale.Image = Global.WinCDS.My.Resources.Resources.returntosale
        Me.cmdPayReturnToSale.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.cmdPayReturnToSale.Location = New System.Drawing.Point(482, 13)
        Me.cmdPayReturnToSale.Name = "cmdPayReturnToSale"
        Me.cmdPayReturnToSale.Size = New System.Drawing.Size(75, 69)
        Me.cmdPayReturnToSale.TabIndex = 36
        Me.cmdPayReturnToSale.Text = "&No Payment"
        Me.cmdPayReturnToSale.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.ToolTip1.SetToolTip(Me.cmdPayReturnToSale, "Cancel payment, and hide payment buttons.")
        Me.cmdPayReturnToSale.UseVisualStyleBackColor = False
        '
        'cmdPayGiftCard
        '
        Me.cmdPayGiftCard.BackColor = System.Drawing.SystemColors.Window
        Me.cmdPayGiftCard.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdPayGiftCard.Image = Global.WinCDS.My.Resources.Resources.Picture
        Me.cmdPayGiftCard.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.cmdPayGiftCard.Location = New System.Drawing.Point(403, 13)
        Me.cmdPayGiftCard.Name = "cmdPayGiftCard"
        Me.cmdPayGiftCard.Size = New System.Drawing.Size(75, 69)
        Me.cmdPayGiftCard.TabIndex = 35
        Me.cmdPayGiftCard.Text = "G&ift Card"
        Me.cmdPayGiftCard.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.ToolTip1.SetToolTip(Me.cmdPayGiftCard, "Payment is by Store Credit Card")
        Me.cmdPayGiftCard.UseVisualStyleBackColor = False
        Me.cmdPayGiftCard.Visible = False
        '
        'cmdPayStoreCard
        '
        Me.cmdPayStoreCard.BackColor = System.Drawing.SystemColors.Window
        Me.cmdPayStoreCard.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdPayStoreCard.Image = Global.WinCDS.My.Resources.Resources.storecard
        Me.cmdPayStoreCard.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.cmdPayStoreCard.Location = New System.Drawing.Point(324, 13)
        Me.cmdPayStoreCard.Name = "cmdPayStoreCard"
        Me.cmdPayStoreCard.Size = New System.Drawing.Size(75, 69)
        Me.cmdPayStoreCard.TabIndex = 34
        Me.cmdPayStoreCard.Text = "S&tore Card"
        Me.cmdPayStoreCard.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.ToolTip1.SetToolTip(Me.cmdPayStoreCard, "Payment is by Store Credit Card")
        Me.cmdPayStoreCard.UseVisualStyleBackColor = False
        '
        'cmdPayDebit
        '
        Me.cmdPayDebit.BackColor = System.Drawing.SystemColors.Window
        Me.cmdPayDebit.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdPayDebit.Image = Global.WinCDS.My.Resources.Resources.debit
        Me.cmdPayDebit.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.cmdPayDebit.Location = New System.Drawing.Point(244, 13)
        Me.cmdPayDebit.Name = "cmdPayDebit"
        Me.cmdPayDebit.Size = New System.Drawing.Size(75, 69)
        Me.cmdPayDebit.TabIndex = 33
        Me.cmdPayDebit.Text = "D&ebit"
        Me.cmdPayDebit.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.ToolTip1.SetToolTip(Me.cmdPayDebit, "Payment is by Debit Card.")
        Me.cmdPayDebit.UseVisualStyleBackColor = False
        '
        'cmdPayCredit
        '
        Me.cmdPayCredit.BackColor = System.Drawing.SystemColors.Window
        Me.cmdPayCredit.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdPayCredit.Image = Global.WinCDS.My.Resources.Resources.mastercard
        Me.cmdPayCredit.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.cmdPayCredit.Location = New System.Drawing.Point(165, 13)
        Me.cmdPayCredit.Name = "cmdPayCredit"
        Me.cmdPayCredit.Size = New System.Drawing.Size(75, 69)
        Me.cmdPayCredit.TabIndex = 32
        Me.cmdPayCredit.Text = "C&redit"
        Me.cmdPayCredit.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.ToolTip1.SetToolTip(Me.cmdPayCredit, "Take a credit card.  You will be prompted for the card type.")
        Me.cmdPayCredit.UseVisualStyleBackColor = False
        '
        'cmdPayCheck
        '
        Me.cmdPayCheck.BackColor = System.Drawing.SystemColors.Window
        Me.cmdPayCheck.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdPayCheck.Image = Global.WinCDS.My.Resources.Resources.check
        Me.cmdPayCheck.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.cmdPayCheck.Location = New System.Drawing.Point(86, 13)
        Me.cmdPayCheck.Name = "cmdPayCheck"
        Me.cmdPayCheck.Size = New System.Drawing.Size(75, 69)
        Me.cmdPayCheck.TabIndex = 31
        Me.cmdPayCheck.Text = "C&heck"
        Me.cmdPayCheck.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.ToolTip1.SetToolTip(Me.cmdPayCheck, "Take a check from the customer.  Remember to check ID!")
        Me.cmdPayCheck.UseVisualStyleBackColor = False
        '
        'cmdPayCash
        '
        Me.cmdPayCash.BackColor = System.Drawing.SystemColors.Window
        Me.cmdPayCash.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdPayCash.Image = Global.WinCDS.My.Resources.Resources.cash
        Me.cmdPayCash.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.cmdPayCash.Location = New System.Drawing.Point(7, 13)
        Me.cmdPayCash.Name = "cmdPayCash"
        Me.cmdPayCash.Size = New System.Drawing.Size(75, 69)
        Me.cmdPayCash.TabIndex = 30
        Me.cmdPayCash.Text = "C&ash"
        Me.cmdPayCash.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.ToolTip1.SetToolTip(Me.cmdPayCash, "Accept a Cash payment.")
        Me.cmdPayCash.UseVisualStyleBackColor = False
        '
        'lblEnterStyle
        '
        Me.lblEnterStyle.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblEnterStyle.Location = New System.Drawing.Point(6, 44)
        Me.lblEnterStyle.Name = "lblEnterStyle"
        Me.lblEnterStyle.Size = New System.Drawing.Size(149, 19)
        Me.lblEnterStyle.TabIndex = 20
        Me.lblEnterStyle.Text = "Enter Style Number:"
        '
        'cmdTax
        '
        Me.cmdTax.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdTax.Location = New System.Drawing.Point(63, 35)
        Me.cmdTax.Name = "cmdTax"
        Me.cmdTax.Size = New System.Drawing.Size(50, 30)
        Me.cmdTax.TabIndex = 21
        Me.cmdTax.Text = "Tax:"
        Me.cmdTax.TextAlign = System.Drawing.ContentAlignment.TopLeft
        Me.cmdTax.UseVisualStyleBackColor = True
        '
        'fraPaymentButtons
        '
        Me.fraPaymentButtons.Controls.Add(Me.cmdPayReturnToSale)
        Me.fraPaymentButtons.Controls.Add(Me.cmdPayGiftCard)
        Me.fraPaymentButtons.Controls.Add(Me.cmdPayStoreCard)
        Me.fraPaymentButtons.Controls.Add(Me.cmdPayDebit)
        Me.fraPaymentButtons.Controls.Add(Me.cmdPayCredit)
        Me.fraPaymentButtons.Controls.Add(Me.cmdPayCheck)
        Me.fraPaymentButtons.Controls.Add(Me.cmdPayCash)
        Me.fraPaymentButtons.Location = New System.Drawing.Point(12, 397)
        Me.fraPaymentButtons.Name = "fraPaymentButtons"
        Me.fraPaymentButtons.Size = New System.Drawing.Size(563, 88)
        Me.fraPaymentButtons.TabIndex = 23
        Me.fraPaymentButtons.TabStop = False
        Me.fraPaymentButtons.Visible = False
        '
        'fraCust
        '
        Me.fraCust.Controls.Add(Me.lblCust)
        Me.fraCust.Location = New System.Drawing.Point(337, 130)
        Me.fraCust.Name = "fraCust"
        Me.fraCust.Size = New System.Drawing.Size(316, 38)
        Me.fraCust.TabIndex = 26
        Me.fraCust.TabStop = False
        Me.fraCust.Text = "Customer:"
        Me.fraCust.Visible = False
        '
        'fraEnterSku
        '
        Me.fraEnterSku.Controls.Add(Me.txtSku)
        Me.fraEnterSku.Controls.Add(Me.cmdFND)
        Me.fraEnterSku.Controls.Add(Me.lblEnterStyle)
        Me.fraEnterSku.Controls.Add(Me.cmdComm)
        Me.fraEnterSku.Controls.Add(Me.cboSalesList)
        Me.fraEnterSku.Location = New System.Drawing.Point(337, 191)
        Me.fraEnterSku.Name = "fraEnterSku"
        Me.fraEnterSku.Size = New System.Drawing.Size(316, 76)
        Me.fraEnterSku.TabIndex = 27
        Me.fraEnterSku.TabStop = False
        '
        'cmdFND
        '
        Me.cmdFND.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom
        Me.cmdFND.ImageAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.cmdFND.Location = New System.Drawing.Point(205, 45)
        Me.cmdFND.Name = "cmdFND"
        Me.cmdFND.Size = New System.Drawing.Size(61, 26)
        Me.cmdFND.TabIndex = 21
        Me.cmdFND.Text = "&FND"
        Me.cmdFND.TextImageRelation = System.Windows.Forms.TextImageRelation.TextBeforeImage
        Me.cmdFND.UseVisualStyleBackColor = True
        Me.cmdFND.Visible = False
        '
        'fraSaleTotals
        '
        Me.fraSaleTotals.Controls.Add(Me.cmdDev)
        Me.fraSaleTotals.Controls.Add(Me.lblDueCaption)
        Me.fraSaleTotals.Controls.Add(Me.lblTenderedCaption)
        Me.fraSaleTotals.Controls.Add(Me.lblTotalCaption)
        Me.fraSaleTotals.Controls.Add(Me.lblTotal)
        Me.fraSaleTotals.Controls.Add(Me.cmdTax)
        Me.fraSaleTotals.Controls.Add(Me.lblTax)
        Me.fraSaleTotals.Controls.Add(Me.lblTendered)
        Me.fraSaleTotals.Controls.Add(Me.lblDue)
        Me.fraSaleTotals.Location = New System.Drawing.Point(340, 268)
        Me.fraSaleTotals.Name = "fraSaleTotals"
        Me.fraSaleTotals.Size = New System.Drawing.Size(316, 124)
        Me.fraSaleTotals.TabIndex = 28
        Me.fraSaleTotals.TabStop = False
        '
        'cmdDev
        '
        Me.cmdDev.Location = New System.Drawing.Point(265, 88)
        Me.cmdDev.Name = "cmdDev"
        Me.cmdDev.Size = New System.Drawing.Size(40, 23)
        Me.cmdDev.TabIndex = 24
        Me.cmdDev.Text = "&D"
        Me.cmdDev.UseVisualStyleBackColor = True
        '
        'lblDueCaption
        '
        Me.lblDueCaption.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblDueCaption.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.lblDueCaption.Location = New System.Drawing.Point(61, 90)
        Me.lblDueCaption.Name = "lblDueCaption"
        Me.lblDueCaption.Size = New System.Drawing.Size(54, 24)
        Me.lblDueCaption.TabIndex = 23
        Me.lblDueCaption.Text = "Due:"
        Me.lblDueCaption.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'lblTenderedCaption
        '
        Me.lblTenderedCaption.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblTenderedCaption.Location = New System.Drawing.Point(25, 66)
        Me.lblTenderedCaption.Name = "lblTenderedCaption"
        Me.lblTenderedCaption.Size = New System.Drawing.Size(90, 20)
        Me.lblTenderedCaption.TabIndex = 22
        Me.lblTenderedCaption.Text = "Tendered:"
        Me.lblTenderedCaption.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'lblTotalCaption
        '
        Me.lblTotalCaption.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblTotalCaption.Location = New System.Drawing.Point(61, 11)
        Me.lblTotalCaption.Name = "lblTotalCaption"
        Me.lblTotalCaption.Size = New System.Drawing.Size(54, 20)
        Me.lblTotalCaption.TabIndex = 0
        Me.lblTotalCaption.Text = "Total:"
        Me.lblTotalCaption.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'chkSavePrinter
        '
        Me.chkSavePrinter.AutoSize = True
        Me.chkSavePrinter.Location = New System.Drawing.Point(487, 175)
        Me.chkSavePrinter.Name = "chkSavePrinter"
        Me.chkSavePrinter.Size = New System.Drawing.Size(137, 17)
        Me.chkSavePrinter.TabIndex = 30
        Me.chkSavePrinter.Text = "Always Use This Printer"
        Me.chkSavePrinter.UseVisualStyleBackColor = True
        Me.chkSavePrinter.Visible = False
        '
        'pnlPicReceipt
        '
        Me.pnlPicReceipt.AutoScroll = True
        Me.pnlPicReceipt.AutoScrollMinSize = New System.Drawing.Size(200, 400)
        Me.pnlPicReceipt.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.pnlPicReceipt.Controls.Add(Me.picReceipt)
        Me.pnlPicReceipt.Controls.Add(Me.picReceiptContainer)
        Me.pnlPicReceipt.Location = New System.Drawing.Point(12, 7)
        Me.pnlPicReceipt.Name = "pnlPicReceipt"
        Me.pnlPicReceipt.Size = New System.Drawing.Size(298, 384)
        Me.pnlPicReceipt.TabIndex = 31
        '
        'picReceipt
        '
        Me.picReceipt.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.picReceipt.Location = New System.Drawing.Point(6, 15)
        Me.picReceipt.Name = "picReceipt"
        Me.picReceipt.Size = New System.Drawing.Size(267, 359)
        Me.picReceipt.TabIndex = 35
        Me.picReceipt.TabStop = False
        '
        'picReceiptContainer
        '
        Me.picReceiptContainer.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.picReceiptContainer.Location = New System.Drawing.Point(6, 15)
        Me.picReceiptContainer.Name = "picReceiptContainer"
        Me.picReceiptContainer.Size = New System.Drawing.Size(267, 359)
        Me.picReceiptContainer.TabIndex = 35
        Me.picReceiptContainer.TabStop = False
        '
        'imgLogo
        '
        Me.imgLogo.Location = New System.Drawing.Point(347, 12)
        Me.imgLogo.Name = "imgLogo"
        Me.imgLogo.Size = New System.Drawing.Size(316, 108)
        Me.imgLogo.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.imgLogo.TabIndex = 19
        Me.imgLogo.TabStop = False
        '
        'CashRegisterPrinterSelector
        '
        Me.CashRegisterPrinterSelector.AllowDYMO = True
        Me.CashRegisterPrinterSelector.AutoSelect = False
        Me.CashRegisterPrinterSelector.Location = New System.Drawing.Point(347, 12)
        Me.CashRegisterPrinterSelector.Name = "CashRegisterPrinterSelector"
        Me.CashRegisterPrinterSelector.Size = New System.Drawing.Size(316, 108)
        Me.CashRegisterPrinterSelector.TabIndex = 24
        Me.CashRegisterPrinterSelector.Visible = False
        '
        'frmCashRegister
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(659, 489)
        Me.Controls.Add(Me.pnlPicReceipt)
        Me.Controls.Add(Me.chkSavePrinter)
        Me.Controls.Add(Me.fraSaleButtons)
        Me.Controls.Add(Me.CashRegisterPrinterSelector)
        Me.Controls.Add(Me.fraSaleTotals)
        Me.Controls.Add(Me.fraEnterSku)
        Me.Controls.Add(Me.fraCust)
        Me.Controls.Add(Me.fraPaymentButtons)
        Me.Controls.Add(Me.imgLogo)
        Me.Controls.Add(Me.vsbReceipt)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.MaximizeBox = False
        Me.Name = "frmCashRegister"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "WinCDS Cash Register"
        Me.fraSaleButtons.ResumeLayout(False)
        Me.fraPaymentButtons.ResumeLayout(False)
        Me.fraCust.ResumeLayout(False)
        Me.fraEnterSku.ResumeLayout(False)
        Me.fraEnterSku.PerformLayout()
        Me.fraSaleTotals.ResumeLayout(False)
        Me.pnlPicReceipt.ResumeLayout(False)
        CType(Me.picReceipt, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.picReceiptContainer, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.imgLogo, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents lblCust As Label
    Friend WithEvents cmdComm As Button
    Friend WithEvents cboSalesList As ComboBox
    Friend WithEvents lblTotal As Label
    Friend WithEvents lblTax As Label
    Friend WithEvents lblTendered As Label
    Friend WithEvents lblDue As Label
    Friend WithEvents fraSaleButtons As GroupBox
    Friend WithEvents cmdPayment As Button
    Friend WithEvents cmdReturn As Button
    Friend WithEvents cmdDiscount As Button
    Friend WithEvents cmdPrint As Button
    Friend WithEvents cmdDone As Button
    Friend WithEvents cmdCancelSale As Button
    Friend WithEvents vsbReceipt As VScrollBar
    Friend WithEvents imgLogo As PictureBox
    Friend WithEvents ToolTip1 As ToolTip
    Friend WithEvents lblEnterStyle As Label
    Friend WithEvents cmdTax As Button
    Friend WithEvents fraPaymentButtons As GroupBox
    Friend WithEvents CashRegisterPrinterSelector As PrinterSelector
    Friend WithEvents fraCust As GroupBox
    Friend WithEvents fraEnterSku As GroupBox
    Friend WithEvents cmdFND As Button
    Friend WithEvents fraSaleTotals As GroupBox
    Friend WithEvents lblDueCaption As Label
    Friend WithEvents lblTenderedCaption As Label
    Friend WithEvents lblTotalCaption As Label
    Friend WithEvents cmdPurchaseGiftCard As Button
    Friend WithEvents cmdMainMenu As Button
    Friend WithEvents cmdPayReturnToSale As Button
    Friend WithEvents cmdPayGiftCard As Button
    Friend WithEvents cmdPayStoreCard As Button
    Friend WithEvents cmdPayDebit As Button
    Friend WithEvents cmdPayCredit As Button
    Friend WithEvents cmdPayCheck As Button
    Friend WithEvents cmdPayCash As Button
    Friend WithEvents txtSku As TextBox
    Friend WithEvents chkSavePrinter As CheckBox
    Friend WithEvents cmdDev As Button
    Friend WithEvents pnlPicReceipt As Panel
    Friend WithEvents picReceipt As PictureBox
    Friend WithEvents picReceiptContainer As PictureBox
End Class
