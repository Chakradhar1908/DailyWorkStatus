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
        Me.txtSku = New System.Windows.Forms.TextBox()
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
        Me.picReceiptContainer = New System.Windows.Forms.PictureBox()
        Me.picReceipt = New System.Windows.Forms.PictureBox()
        Me.imgLogo = New System.Windows.Forms.PictureBox()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.lblEnterStyle = New System.Windows.Forms.Label()
        Me.cmdTax = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.fraPaymentButtons = New System.Windows.Forms.GroupBox()
        Me.chkSavePrinter = New System.Windows.Forms.CheckBox()
        Me.fraCust = New System.Windows.Forms.GroupBox()
        Me.CashRegisterPrinterSelector = New WinCDS.PrinterSelector()
        Me.fraEnterSku = New System.Windows.Forms.GroupBox()
        Me.cmdFND = New System.Windows.Forms.Button()
        Me.fraSaleTotals = New System.Windows.Forms.GroupBox()
        Me.lblDueCaption = New System.Windows.Forms.Label()
        Me.lblTenderedCaption = New System.Windows.Forms.Label()
        Me.lblTotalCaption = New System.Windows.Forms.Label()
        Me.fraSaleButtons.SuspendLayout()
        CType(Me.picReceiptContainer, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.picReceipt, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.imgLogo, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.fraCust.SuspendLayout()
        Me.fraEnterSku.SuspendLayout()
        Me.fraSaleTotals.SuspendLayout()
        Me.SuspendLayout()
        '
        'lblCust
        '
        Me.lblCust.Location = New System.Drawing.Point(7, 18)
        Me.lblCust.Name = "lblCust"
        Me.lblCust.Size = New System.Drawing.Size(222, 27)
        Me.lblCust.TabIndex = 0
        Me.lblCust.Text = "lblCust"
        '
        'cmdComm
        '
        Me.cmdComm.Location = New System.Drawing.Point(220, 47)
        Me.cmdComm.Name = "cmdComm"
        Me.cmdComm.Size = New System.Drawing.Size(39, 23)
        Me.cmdComm.TabIndex = 1
        Me.cmdComm.Text = "&C"
        Me.cmdComm.UseVisualStyleBackColor = True
        '
        'cboSalesList
        '
        Me.cboSalesList.FormattingEnabled = True
        Me.cboSalesList.Location = New System.Drawing.Point(696, 358)
        Me.cboSalesList.Name = "cboSalesList"
        Me.cboSalesList.Size = New System.Drawing.Size(121, 21)
        Me.cboSalesList.TabIndex = 2
        '
        'lblTotal
        '
        Me.lblTotal.AutoSize = True
        Me.lblTotal.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblTotal.Location = New System.Drawing.Point(185, 11)
        Me.lblTotal.Name = "lblTotal"
        Me.lblTotal.Size = New System.Drawing.Size(44, 20)
        Me.lblTotal.TabIndex = 3
        Me.lblTotal.Text = "0.00"
        Me.lblTotal.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'lblTax
        '
        Me.lblTax.AutoSize = True
        Me.lblTax.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblTax.Location = New System.Drawing.Point(185, 35)
        Me.lblTax.Name = "lblTax"
        Me.lblTax.Size = New System.Drawing.Size(44, 20)
        Me.lblTax.TabIndex = 4
        Me.lblTax.Text = "0.00"
        Me.lblTax.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'lblTendered
        '
        Me.lblTendered.AutoSize = True
        Me.lblTendered.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblTendered.Location = New System.Drawing.Point(185, 66)
        Me.lblTendered.Name = "lblTendered"
        Me.lblTendered.Size = New System.Drawing.Size(44, 20)
        Me.lblTendered.TabIndex = 5
        Me.lblTendered.Text = "0.00"
        Me.lblTendered.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'lblDue
        '
        Me.lblDue.AutoSize = True
        Me.lblDue.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblDue.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.lblDue.Location = New System.Drawing.Point(180, 90)
        Me.lblDue.Name = "lblDue"
        Me.lblDue.Size = New System.Drawing.Size(49, 24)
        Me.lblDue.TabIndex = 6
        Me.lblDue.Text = "0.00"
        Me.lblDue.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'txtSku
        '
        Me.txtSku.Location = New System.Drawing.Point(6, 19)
        Me.txtSku.Name = "txtSku"
        Me.txtSku.Size = New System.Drawing.Size(231, 20)
        Me.txtSku.TabIndex = 7
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
        Me.fraSaleButtons.Location = New System.Drawing.Point(12, 410)
        Me.fraSaleButtons.Name = "fraSaleButtons"
        Me.fraSaleButtons.Size = New System.Drawing.Size(654, 92)
        Me.fraSaleButtons.TabIndex = 8
        Me.fraSaleButtons.TabStop = False
        '
        'cmdPurchaseGiftCard
        '
        Me.cmdPurchaseGiftCard.Image = Global.WinCDS.My.Resources.Resources.Picture
        Me.cmdPurchaseGiftCard.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.cmdPurchaseGiftCard.Location = New System.Drawing.Point(577, 19)
        Me.cmdPurchaseGiftCard.Name = "cmdPurchaseGiftCard"
        Me.cmdPurchaseGiftCard.Size = New System.Drawing.Size(75, 67)
        Me.cmdPurchaseGiftCard.TabIndex = 17
        Me.cmdPurchaseGiftCard.Text = "Purchase &Gift Card"
        Me.cmdPurchaseGiftCard.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.cmdPurchaseGiftCard.UseVisualStyleBackColor = True
        '
        'cmdMainMenu
        '
        Me.cmdMainMenu.Image = Global.WinCDS.My.Resources.Resources.menu
        Me.cmdMainMenu.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.cmdMainMenu.Location = New System.Drawing.Point(496, 19)
        Me.cmdMainMenu.Name = "cmdMainMenu"
        Me.cmdMainMenu.Size = New System.Drawing.Size(75, 67)
        Me.cmdMainMenu.TabIndex = 16
        Me.cmdMainMenu.Text = "&Main Menu"
        Me.cmdMainMenu.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.cmdMainMenu.UseVisualStyleBackColor = True
        '
        'cmdPayment
        '
        Me.cmdPayment.Image = Global.WinCDS.My.Resources.Resources.cash_register_icon
        Me.cmdPayment.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.cmdPayment.Location = New System.Drawing.Point(6, 19)
        Me.cmdPayment.Name = "cmdPayment"
        Me.cmdPayment.Size = New System.Drawing.Size(75, 67)
        Me.cmdPayment.TabIndex = 9
        Me.cmdPayment.Text = "Button2"
        Me.cmdPayment.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.cmdPayment.UseVisualStyleBackColor = True
        '
        'cmdReturn
        '
        Me.cmdReturn.BackColor = System.Drawing.SystemColors.ButtonFace
        Me.cmdReturn.Image = Global.WinCDS.My.Resources.Resources.return4
        Me.cmdReturn.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.cmdReturn.Location = New System.Drawing.Point(87, 19)
        Me.cmdReturn.Name = "cmdReturn"
        Me.cmdReturn.Size = New System.Drawing.Size(75, 67)
        Me.cmdReturn.TabIndex = 10
        Me.cmdReturn.Text = "Button3"
        Me.cmdReturn.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.cmdReturn.UseVisualStyleBackColor = False
        '
        'cmdDiscount
        '
        Me.cmdDiscount.Image = Global.WinCDS.My.Resources.Resources.Icon_Specials
        Me.cmdDiscount.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.cmdDiscount.Location = New System.Drawing.Point(168, 19)
        Me.cmdDiscount.Name = "cmdDiscount"
        Me.cmdDiscount.Size = New System.Drawing.Size(75, 67)
        Me.cmdDiscount.TabIndex = 11
        Me.cmdDiscount.Text = "Button4"
        Me.cmdDiscount.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.cmdDiscount.UseVisualStyleBackColor = True
        '
        'cmdCancelSale
        '
        Me.cmdCancelSale.Image = Global.WinCDS.My.Resources.Resources.cancel
        Me.cmdCancelSale.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.cmdCancelSale.Location = New System.Drawing.Point(249, 19)
        Me.cmdCancelSale.Name = "cmdCancelSale"
        Me.cmdCancelSale.Size = New System.Drawing.Size(75, 67)
        Me.cmdCancelSale.TabIndex = 15
        Me.cmdCancelSale.Text = "Button1"
        Me.cmdCancelSale.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.cmdCancelSale.UseVisualStyleBackColor = True
        '
        'cmdPrint
        '
        Me.cmdPrint.Image = Global.WinCDS.My.Resources.Resources.reprint
        Me.cmdPrint.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.cmdPrint.Location = New System.Drawing.Point(334, 19)
        Me.cmdPrint.Name = "cmdPrint"
        Me.cmdPrint.Size = New System.Drawing.Size(75, 67)
        Me.cmdPrint.TabIndex = 12
        Me.cmdPrint.Text = "Button5"
        Me.cmdPrint.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.cmdPrint.UseVisualStyleBackColor = True
        '
        'cmdDone
        '
        Me.cmdDone.BackColor = System.Drawing.SystemColors.Control
        Me.cmdDone.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdDone.Image = Global.WinCDS.My.Resources.Resources.dollarsign
        Me.cmdDone.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.cmdDone.Location = New System.Drawing.Point(415, 22)
        Me.cmdDone.Name = "cmdDone"
        Me.cmdDone.Size = New System.Drawing.Size(75, 67)
        Me.cmdDone.TabIndex = 13
        Me.cmdDone.Text = "Button6"
        Me.cmdDone.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.cmdDone.UseVisualStyleBackColor = False
        '
        'vsbReceipt
        '
        Me.vsbReceipt.Location = New System.Drawing.Point(282, 16)
        Me.vsbReceipt.Name = "vsbReceipt"
        Me.vsbReceipt.Size = New System.Drawing.Size(13, 379)
        Me.vsbReceipt.TabIndex = 16
        '
        'picReceiptContainer
        '
        Me.picReceiptContainer.Location = New System.Drawing.Point(696, 245)
        Me.picReceiptContainer.Name = "picReceiptContainer"
        Me.picReceiptContainer.Size = New System.Drawing.Size(100, 50)
        Me.picReceiptContainer.TabIndex = 17
        Me.picReceiptContainer.TabStop = False
        '
        'picReceipt
        '
        Me.picReceipt.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.picReceipt.Location = New System.Drawing.Point(12, 15)
        Me.picReceipt.Name = "picReceipt"
        Me.picReceipt.Size = New System.Drawing.Size(267, 380)
        Me.picReceipt.TabIndex = 18
        Me.picReceipt.TabStop = False
        '
        'imgLogo
        '
        Me.imgLogo.Location = New System.Drawing.Point(691, 385)
        Me.imgLogo.Name = "imgLogo"
        Me.imgLogo.Size = New System.Drawing.Size(71, 52)
        Me.imgLogo.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.imgLogo.TabIndex = 19
        Me.imgLogo.TabStop = False
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
        Me.cmdTax.Location = New System.Drawing.Point(65, 35)
        Me.cmdTax.Name = "cmdTax"
        Me.cmdTax.Size = New System.Drawing.Size(50, 27)
        Me.cmdTax.TabIndex = 21
        Me.cmdTax.Text = "Tax:"
        Me.cmdTax.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(620, 287)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(39, 13)
        Me.Label1.TabIndex = 22
        Me.Label1.Text = "Label1"
        '
        'fraPaymentButtons
        '
        Me.fraPaymentButtons.Location = New System.Drawing.Point(608, 180)
        Me.fraPaymentButtons.Name = "fraPaymentButtons"
        Me.fraPaymentButtons.Size = New System.Drawing.Size(83, 54)
        Me.fraPaymentButtons.TabIndex = 23
        Me.fraPaymentButtons.TabStop = False
        Me.fraPaymentButtons.Text = "GroupBox1"
        '
        'chkSavePrinter
        '
        Me.chkSavePrinter.AutoSize = True
        Me.chkSavePrinter.Location = New System.Drawing.Point(98, 45)
        Me.chkSavePrinter.Name = "chkSavePrinter"
        Me.chkSavePrinter.Size = New System.Drawing.Size(137, 17)
        Me.chkSavePrinter.TabIndex = 25
        Me.chkSavePrinter.Text = "Always Use This Printer"
        Me.chkSavePrinter.UseVisualStyleBackColor = True
        Me.chkSavePrinter.Visible = False
        '
        'fraCust
        '
        Me.fraCust.Controls.Add(Me.chkSavePrinter)
        Me.fraCust.Controls.Add(Me.lblCust)
        Me.fraCust.Location = New System.Drawing.Point(308, 130)
        Me.fraCust.Name = "fraCust"
        Me.fraCust.Size = New System.Drawing.Size(267, 60)
        Me.fraCust.TabIndex = 26
        Me.fraCust.TabStop = False
        Me.fraCust.Text = "Customer:"
        Me.fraCust.Visible = False
        '
        'CashRegisterPrinterSelector
        '
        Me.CashRegisterPrinterSelector.AllowDYMO = True
        Me.CashRegisterPrinterSelector.AutoSelect = False
        Me.CashRegisterPrinterSelector.Location = New System.Drawing.Point(308, 16)
        Me.CashRegisterPrinterSelector.Name = "CashRegisterPrinterSelector"
        Me.CashRegisterPrinterSelector.Size = New System.Drawing.Size(267, 108)
        Me.CashRegisterPrinterSelector.TabIndex = 24
        Me.CashRegisterPrinterSelector.Visible = False
        '
        'fraEnterSku
        '
        Me.fraEnterSku.Controls.Add(Me.cmdFND)
        Me.fraEnterSku.Controls.Add(Me.txtSku)
        Me.fraEnterSku.Controls.Add(Me.lblEnterStyle)
        Me.fraEnterSku.Controls.Add(Me.cmdComm)
        Me.fraEnterSku.Location = New System.Drawing.Point(310, 193)
        Me.fraEnterSku.Name = "fraEnterSku"
        Me.fraEnterSku.Size = New System.Drawing.Size(265, 76)
        Me.fraEnterSku.TabIndex = 27
        Me.fraEnterSku.TabStop = False
        '
        'cmdFND
        '
        Me.cmdFND.Location = New System.Drawing.Point(156, 45)
        Me.cmdFND.Name = "cmdFND"
        Me.cmdFND.Size = New System.Drawing.Size(61, 23)
        Me.cmdFND.TabIndex = 21
        Me.cmdFND.Text = "&FND"
        Me.cmdFND.UseVisualStyleBackColor = True
        '
        'fraSaleTotals
        '
        Me.fraSaleTotals.Controls.Add(Me.lblDueCaption)
        Me.fraSaleTotals.Controls.Add(Me.lblTenderedCaption)
        Me.fraSaleTotals.Controls.Add(Me.lblTotalCaption)
        Me.fraSaleTotals.Controls.Add(Me.lblTotal)
        Me.fraSaleTotals.Controls.Add(Me.cmdTax)
        Me.fraSaleTotals.Controls.Add(Me.lblTax)
        Me.fraSaleTotals.Controls.Add(Me.lblTendered)
        Me.fraSaleTotals.Controls.Add(Me.lblDue)
        Me.fraSaleTotals.Location = New System.Drawing.Point(306, 271)
        Me.fraSaleTotals.Name = "fraSaleTotals"
        Me.fraSaleTotals.Size = New System.Drawing.Size(269, 124)
        Me.fraSaleTotals.TabIndex = 28
        Me.fraSaleTotals.TabStop = False
        '
        'lblDueCaption
        '
        Me.lblDueCaption.AutoSize = True
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
        Me.lblTenderedCaption.AutoSize = True
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
        Me.lblTotalCaption.AutoSize = True
        Me.lblTotalCaption.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblTotalCaption.Location = New System.Drawing.Point(61, 11)
        Me.lblTotalCaption.Name = "lblTotalCaption"
        Me.lblTotalCaption.Size = New System.Drawing.Size(54, 20)
        Me.lblTotalCaption.TabIndex = 0
        Me.lblTotalCaption.Text = "Total:"
        Me.lblTotalCaption.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'frmCashRegister
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(800, 514)
        Me.Controls.Add(Me.fraSaleTotals)
        Me.Controls.Add(Me.fraEnterSku)
        Me.Controls.Add(Me.fraCust)
        Me.Controls.Add(Me.CashRegisterPrinterSelector)
        Me.Controls.Add(Me.fraPaymentButtons)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.imgLogo)
        Me.Controls.Add(Me.picReceipt)
        Me.Controls.Add(Me.picReceiptContainer)
        Me.Controls.Add(Me.vsbReceipt)
        Me.Controls.Add(Me.fraSaleButtons)
        Me.Controls.Add(Me.cboSalesList)
        Me.Name = "frmCashRegister"
        Me.Text = "frmCashRegister"
        Me.fraSaleButtons.ResumeLayout(False)
        CType(Me.picReceiptContainer, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.picReceipt, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.imgLogo, System.ComponentModel.ISupportInitialize).EndInit()
        Me.fraCust.ResumeLayout(False)
        Me.fraCust.PerformLayout()
        Me.fraEnterSku.ResumeLayout(False)
        Me.fraEnterSku.PerformLayout()
        Me.fraSaleTotals.ResumeLayout(False)
        Me.fraSaleTotals.PerformLayout()
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
    Friend WithEvents txtSku As TextBox
    Friend WithEvents fraSaleButtons As GroupBox
    Friend WithEvents cmdPayment As Button
    Friend WithEvents cmdReturn As Button
    Friend WithEvents cmdDiscount As Button
    Friend WithEvents cmdPrint As Button
    Friend WithEvents cmdDone As Button
    Friend WithEvents cmdCancelSale As Button
    Friend WithEvents vsbReceipt As VScrollBar
    Friend WithEvents picReceiptContainer As PictureBox
    Friend WithEvents picReceipt As PictureBox
    Friend WithEvents imgLogo As PictureBox
    Friend WithEvents ToolTip1 As ToolTip
    Friend WithEvents lblEnterStyle As Label
    Friend WithEvents cmdTax As Button
    Friend WithEvents Label1 As Label
    Friend WithEvents fraPaymentButtons As GroupBox
    Friend WithEvents CashRegisterPrinterSelector As PrinterSelector
    Friend WithEvents chkSavePrinter As CheckBox
    Friend WithEvents fraCust As GroupBox
    Friend WithEvents fraEnterSku As GroupBox
    Friend WithEvents cmdFND As Button
    Friend WithEvents fraSaleTotals As GroupBox
    Friend WithEvents lblDueCaption As Label
    Friend WithEvents lblTenderedCaption As Label
    Friend WithEvents lblTotalCaption As Label
    Friend WithEvents cmdPurchaseGiftCard As Button
    Friend WithEvents cmdMainMenu As Button
End Class
