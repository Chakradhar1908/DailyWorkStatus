<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class BillOSale
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
        Me.lblDateCaption = New System.Windows.Forms.Label()
        Me.dteSaleDate = New System.Windows.Forms.DateTimePicker()
        Me.lblDelWeekday = New System.Windows.Forms.Label()
        Me.chkDelivery = New System.Windows.Forms.CheckBox()
        Me.chkPickup = New System.Windows.Forms.CheckBox()
        Me.optIndividual = New System.Windows.Forms.RadioButton()
        Me.optBusiness = New System.Windows.Forms.RadioButton()
        Me.StoreName = New System.Windows.Forms.Label()
        Me.StoreAddress = New System.Windows.Forms.Label()
        Me.StoreCity = New System.Windows.Forms.Label()
        Me.StorePhone = New System.Windows.Forms.Label()
        Me.txtSaleNo = New System.Windows.Forms.TextBox()
        Me.lblSaleNoCaption = New System.Windows.Forms.Label()
        Me.lblStatusCaption = New System.Windows.Forms.Label()
        Me.lblTypeCaption = New System.Windows.Forms.Label()
        Me.lblAdvCaption = New System.Windows.Forms.Label()
        Me.lblTaxCaption = New System.Windows.Forms.Label()
        Me.SaleStatus = New System.Windows.Forms.TextBox()
        Me.cboCustType = New System.Windows.Forms.ComboBox()
        Me.cboAdvertisingType = New System.Windows.Forms.ComboBox()
        Me.cboTaxZone = New System.Windows.Forms.ComboBox()
        Me.dteDelivery = New System.Windows.Forms.DateTimePicker()
        Me.lblFirst = New System.Windows.Forms.Label()
        Me.lblLast = New System.Windows.Forms.Label()
        Me.lblEmail = New System.Windows.Forms.Label()
        Me.CustomerFirst = New System.Windows.Forms.TextBox()
        Me.CustomerLast = New System.Windows.Forms.TextBox()
        Me.Email = New System.Windows.Forms.TextBox()
        Me.CustomerAddress = New System.Windows.Forms.TextBox()
        Me.ShipToFirst = New System.Windows.Forms.TextBox()
        Me.ShipToLast = New System.Windows.Forms.TextBox()
        Me.lblAddr = New System.Windows.Forms.Label()
        Me.lblShipToAddressCaption = New System.Windows.Forms.Label()
        Me.lblShipFirst = New System.Windows.Forms.Label()
        Me.lblShipLast = New System.Windows.Forms.Label()
        Me.AddAddress = New System.Windows.Forms.TextBox()
        Me.CustomerAddress2 = New System.Windows.Forms.TextBox()
        Me.lblAddAddr = New System.Windows.Forms.Label()
        Me.lblShipAddr = New System.Windows.Forms.Label()
        Me.CustomerCity = New System.Windows.Forms.TextBox()
        Me.CustomerZip = New System.Windows.Forms.TextBox()
        Me.CustomerCity2 = New System.Windows.Forms.TextBox()
        Me.CustomerZip2 = New System.Windows.Forms.TextBox()
        Me.lblCity = New System.Windows.Forms.Label()
        Me.lblZip = New System.Windows.Forms.Label()
        Me.lblShipCity = New System.Windows.Forms.Label()
        Me.lblShipZip = New System.Windows.Forms.Label()
        Me.cboPhone1 = New System.Windows.Forms.ComboBox()
        Me.cboPhone2 = New System.Windows.Forms.ComboBox()
        Me.cboPhone3 = New System.Windows.Forms.ComboBox()
        Me.CustomerPhone1 = New System.Windows.Forms.TextBox()
        Me.CustomerPhone2 = New System.Windows.Forms.TextBox()
        Me.CustomerPhone3 = New System.Windows.Forms.TextBox()
        Me.txtSpecInst = New System.Windows.Forms.TextBox()
        Me.lblSpecInstr = New System.Windows.Forms.Label()
        Me.Sales1 = New System.Windows.Forms.TextBox()
        Me.Sales2 = New System.Windows.Forms.TextBox()
        Me.Sales3 = New System.Windows.Forms.TextBox()
        Me.SalesSplit1 = New System.Windows.Forms.ComboBox()
        Me.SalesSplit2 = New System.Windows.Forms.ComboBox()
        Me.SalesSplit3 = New System.Windows.Forms.ComboBox()
        Me.lblSales1 = New System.Windows.Forms.Label()
        Me.lblSales2 = New System.Windows.Forms.Label()
        Me.lblSales3 = New System.Windows.Forms.Label()
        Me.cmdShowBodyOfSale = New System.Windows.Forms.Button()
        Me.fraTimeWindow = New System.Windows.Forms.GroupBox()
        Me.dtpDelWindow2 = New System.Windows.Forms.DateTimePicker()
        Me.lblTimeWindow = New System.Windows.Forms.Label()
        Me.dtpDelWindow = New System.Windows.Forms.DateTimePicker()
        Me.lblDelDate = New System.Windows.Forms.Label()
        Me.ToolTipBillOSale = New System.Windows.Forms.ToolTip(Me.components)
        Me.opt30323 = New System.Windows.Forms.RadioButton()
        Me.opt30252 = New System.Windows.Forms.RadioButton()
        Me.cmdPrintLabel = New System.Windows.Forms.Button()
        Me.cmdCancel = New System.Windows.Forms.Button()
        Me.cmdApplyBillOSale = New System.Windows.Forms.Button()
        Me.ScanUp = New System.Windows.Forms.Button()
        Me.cmdProcessSale = New System.Windows.Forms.Button()
        Me.cmdNextSale = New System.Windows.Forms.Button()
        Me.cmdMainMenu = New System.Windows.Forms.Button()
        Me.cmdClear = New System.Windows.Forms.Button()
        Me.ScanDn = New System.Windows.Forms.Button()
        Me.ScanUp123 = New System.Windows.Forms.Button()
        Me.cmdSoldTags = New System.Windows.Forms.Button()
        Me.cmdEmail = New System.Windows.Forms.Button()
        Me.cmdPrint = New System.Windows.Forms.Button()
        Me.cmdChangePrice = New System.Windows.Forms.Button()
        Me.imgCalendar = New System.Windows.Forms.PictureBox()
        Me.cmdNoChangePrice = New System.Windows.Forms.Button()
        Me.fraBOS2 = New System.Windows.Forms.GroupBox()
        Me.UGridIO1 = New WinCDS.UGridIO()
        Me.rtbStorePolicy = New WinCDS.RichTextBoxNew()
        Me.rtb = New WinCDS.RichTextBoxNew()
        Me.fraBOS2Commands = New System.Windows.Forms.GroupBox()
        Me.Notes_Open = New System.Windows.Forms.Button()
        Me.fraHover = New System.Windows.Forms.GroupBox()
        Me.picHover = New System.Windows.Forms.PictureBox()
        Me.lblBalDueCaption = New System.Windows.Forms.Label()
        Me.BalDue = New System.Windows.Forms.TextBox()
        Me.BillOfSale = New System.Windows.Forms.Label()
        Me.tmrHover = New System.Windows.Forms.Timer(Me.components)
        Me.tmrFormat = New System.Windows.Forms.Timer(Me.components)
        Me.txtFormatHelper = New System.Windows.Forms.TextBox()
        Me.fraButtons = New System.Windows.Forms.GroupBox()
        Me.fraPrintType = New System.Windows.Forms.GroupBox()
        Me.lblPrintType = New System.Windows.Forms.Label()
        Me.lblGrossSalesCaption = New System.Windows.Forms.Label()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.lblGrossSales = New System.Windows.Forms.Label()
        Me.picFormatHelper = New System.Windows.Forms.PictureBox()
        Me.imgLogo = New System.Windows.Forms.PictureBox()
        Me.ugrFake = New WinCDS.UGridIO()
        Me.fraTimeWindow.SuspendLayout()
        CType(Me.imgCalendar, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.fraBOS2.SuspendLayout()
        Me.fraBOS2Commands.SuspendLayout()
        Me.fraHover.SuspendLayout()
        CType(Me.picHover, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.fraButtons.SuspendLayout()
        Me.fraPrintType.SuspendLayout()
        CType(Me.picFormatHelper, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.imgLogo, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'lblDateCaption
        '
        Me.lblDateCaption.AutoSize = True
        Me.lblDateCaption.Font = New System.Drawing.Font("Arial", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblDateCaption.Location = New System.Drawing.Point(26, 0)
        Me.lblDateCaption.Name = "lblDateCaption"
        Me.lblDateCaption.Size = New System.Drawing.Size(50, 22)
        Me.lblDateCaption.TabIndex = 0
        Me.lblDateCaption.Text = "Date"
        '
        'dteSaleDate
        '
        Me.dteSaleDate.CustomFormat = "MM/dd/yyyy"
        Me.dteSaleDate.Font = New System.Drawing.Font("Arial", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dteSaleDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dteSaleDate.Location = New System.Drawing.Point(78, 5)
        Me.dteSaleDate.Name = "dteSaleDate"
        Me.dteSaleDate.Size = New System.Drawing.Size(118, 29)
        Me.dteSaleDate.TabIndex = 36
        '
        'lblDelWeekday
        '
        Me.lblDelWeekday.BackColor = System.Drawing.SystemColors.HighlightText
        Me.lblDelWeekday.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblDelWeekday.Font = New System.Drawing.Font("Arial", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblDelWeekday.Location = New System.Drawing.Point(78, 39)
        Me.lblDelWeekday.Name = "lblDelWeekday"
        Me.lblDelWeekday.Size = New System.Drawing.Size(100, 27)
        Me.lblDelWeekday.TabIndex = 39
        '
        'chkDelivery
        '
        Me.chkDelivery.AutoSize = True
        Me.chkDelivery.CheckAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.chkDelivery.Location = New System.Drawing.Point(8, 33)
        Me.chkDelivery.Name = "chkDelivery"
        Me.chkDelivery.Size = New System.Drawing.Size(64, 17)
        Me.chkDelivery.TabIndex = 37
        Me.chkDelivery.Text = "Delivery"
        Me.chkDelivery.UseVisualStyleBackColor = True
        '
        'chkPickup
        '
        Me.chkPickup.AutoSize = True
        Me.chkPickup.CheckAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.chkPickup.Location = New System.Drawing.Point(8, 51)
        Me.chkPickup.Name = "chkPickup"
        Me.chkPickup.Size = New System.Drawing.Size(64, 17)
        Me.chkPickup.TabIndex = 38
        Me.chkPickup.Text = "Pick Up"
        Me.ToolTipBillOSale.SetToolTip(Me.chkPickup, " Check on to set up a pick up ")
        Me.chkPickup.UseVisualStyleBackColor = True
        '
        'optIndividual
        '
        Me.optIndividual.AutoSize = True
        Me.optIndividual.Checked = True
        Me.optIndividual.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.optIndividual.Location = New System.Drawing.Point(45, 127)
        Me.optIndividual.Name = "optIndividual"
        Me.optIndividual.Size = New System.Drawing.Size(71, 18)
        Me.optIndividual.TabIndex = 43
        Me.optIndividual.TabStop = True
        Me.optIndividual.Text = "Inidividual"
        Me.optIndividual.UseVisualStyleBackColor = True
        '
        'optBusiness
        '
        Me.optBusiness.AutoSize = True
        Me.optBusiness.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.optBusiness.Location = New System.Drawing.Point(141, 127)
        Me.optBusiness.Name = "optBusiness"
        Me.optBusiness.Size = New System.Drawing.Size(70, 18)
        Me.optBusiness.TabIndex = 44
        Me.optBusiness.Text = "Business"
        Me.optBusiness.UseVisualStyleBackColor = True
        '
        'StoreName
        '
        Me.StoreName.AutoSize = True
        Me.StoreName.Font = New System.Drawing.Font("Arial", 18.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.StoreName.Location = New System.Drawing.Point(330, 5)
        Me.StoreName.Name = "StoreName"
        Me.StoreName.Size = New System.Drawing.Size(143, 29)
        Me.StoreName.TabIndex = 12
        Me.StoreName.Text = "Store name"
        '
        'StoreAddress
        '
        Me.StoreAddress.AutoSize = True
        Me.StoreAddress.Font = New System.Drawing.Font("Arial", 18.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.StoreAddress.Location = New System.Drawing.Point(330, 29)
        Me.StoreAddress.Name = "StoreAddress"
        Me.StoreAddress.Size = New System.Drawing.Size(172, 29)
        Me.StoreAddress.TabIndex = 13
        Me.StoreAddress.Text = "Store address"
        '
        'StoreCity
        '
        Me.StoreCity.AutoSize = True
        Me.StoreCity.Font = New System.Drawing.Font("Arial", 18.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.StoreCity.Location = New System.Drawing.Point(341, 50)
        Me.StoreCity.Name = "StoreCity"
        Me.StoreCity.Size = New System.Drawing.Size(122, 29)
        Me.StoreCity.TabIndex = 14
        Me.StoreCity.Text = "Store city"
        '
        'StorePhone
        '
        Me.StorePhone.AutoSize = True
        Me.StorePhone.Font = New System.Drawing.Font("Arial", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.StorePhone.Location = New System.Drawing.Point(342, 79)
        Me.StorePhone.Name = "StorePhone"
        Me.StorePhone.Size = New System.Drawing.Size(103, 19)
        Me.StorePhone.TabIndex = 15
        Me.StorePhone.Text = "Store phone"
        '
        'txtSaleNo
        '
        Me.txtSaleNo.Font = New System.Drawing.Font("Arial", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtSaleNo.Location = New System.Drawing.Point(743, 7)
        Me.txtSaleNo.Name = "txtSaleNo"
        Me.txtSaleNo.Size = New System.Drawing.Size(120, 29)
        Me.txtSaleNo.TabIndex = 46
        '
        'lblSaleNoCaption
        '
        Me.lblSaleNoCaption.AutoSize = True
        Me.lblSaleNoCaption.Font = New System.Drawing.Font("Arial", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblSaleNoCaption.Location = New System.Drawing.Point(756, 41)
        Me.lblSaleNoCaption.Name = "lblSaleNoCaption"
        Me.lblSaleNoCaption.Size = New System.Drawing.Size(82, 22)
        Me.lblSaleNoCaption.TabIndex = 17
        Me.lblSaleNoCaption.Text = "Sale No:"
        '
        'lblStatusCaption
        '
        Me.lblStatusCaption.AutoSize = True
        Me.lblStatusCaption.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblStatusCaption.Location = New System.Drawing.Point(703, 76)
        Me.lblStatusCaption.Name = "lblStatusCaption"
        Me.lblStatusCaption.Size = New System.Drawing.Size(38, 14)
        Me.lblStatusCaption.TabIndex = 46
        Me.lblStatusCaption.Text = "Status"
        '
        'lblTypeCaption
        '
        Me.lblTypeCaption.AutoSize = True
        Me.lblTypeCaption.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblTypeCaption.Location = New System.Drawing.Point(711, 104)
        Me.lblTypeCaption.Name = "lblTypeCaption"
        Me.lblTypeCaption.Size = New System.Drawing.Size(30, 14)
        Me.lblTypeCaption.TabIndex = 47
        Me.lblTypeCaption.Text = "Type"
        '
        'lblAdvCaption
        '
        Me.lblAdvCaption.AutoSize = True
        Me.lblAdvCaption.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblAdvCaption.Location = New System.Drawing.Point(679, 131)
        Me.lblAdvCaption.Name = "lblAdvCaption"
        Me.lblAdvCaption.Size = New System.Drawing.Size(62, 14)
        Me.lblAdvCaption.TabIndex = 48
        Me.lblAdvCaption.Text = "Advertising"
        '
        'lblTaxCaption
        '
        Me.lblTaxCaption.AutoSize = True
        Me.lblTaxCaption.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblTaxCaption.Location = New System.Drawing.Point(717, 154)
        Me.lblTaxCaption.Name = "lblTaxCaption"
        Me.lblTaxCaption.Size = New System.Drawing.Size(24, 14)
        Me.lblTaxCaption.TabIndex = 49
        Me.lblTaxCaption.Text = "Tax"
        '
        'SaleStatus
        '
        Me.SaleStatus.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.SaleStatus.Location = New System.Drawing.Point(743, 70)
        Me.SaleStatus.Name = "SaleStatus"
        Me.SaleStatus.Size = New System.Drawing.Size(121, 20)
        Me.SaleStatus.TabIndex = 32
        '
        'cboCustType
        '
        Me.cboCustType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboCustType.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboCustType.FormattingEnabled = True
        Me.cboCustType.Location = New System.Drawing.Point(743, 95)
        Me.cboCustType.Name = "cboCustType"
        Me.cboCustType.Size = New System.Drawing.Size(121, 22)
        Me.cboCustType.TabIndex = 33
        Me.ToolTipBillOSale.SetToolTip(Me.cboCustType, " Sets Special Pricing ")
        '
        'cboAdvertisingType
        '
        Me.cboAdvertisingType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboAdvertisingType.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboAdvertisingType.FormattingEnabled = True
        Me.cboAdvertisingType.Location = New System.Drawing.Point(743, 122)
        Me.cboAdvertisingType.Name = "cboAdvertisingType"
        Me.cboAdvertisingType.Size = New System.Drawing.Size(121, 22)
        Me.cboAdvertisingType.TabIndex = 34
        Me.ToolTipBillOSale.SetToolTip(Me.cboAdvertisingType, " Sets Advertising Type In Customer Set Up ")
        '
        'cboTaxZone
        '
        Me.cboTaxZone.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboTaxZone.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboTaxZone.FormattingEnabled = True
        Me.cboTaxZone.Location = New System.Drawing.Point(743, 149)
        Me.cboTaxZone.Name = "cboTaxZone"
        Me.cboTaxZone.Size = New System.Drawing.Size(121, 22)
        Me.cboTaxZone.TabIndex = 35
        '
        'dteDelivery
        '
        Me.dteDelivery.CustomFormat = "MM/dd/yyyy"
        Me.dteDelivery.Font = New System.Drawing.Font("Arial", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dteDelivery.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dteDelivery.Location = New System.Drawing.Point(77, 69)
        Me.dteDelivery.Name = "dteDelivery"
        Me.dteDelivery.Size = New System.Drawing.Size(133, 29)
        Me.dteDelivery.TabIndex = 40
        '
        'lblFirst
        '
        Me.lblFirst.AutoSize = True
        Me.lblFirst.Location = New System.Drawing.Point(5, 138)
        Me.lblFirst.Name = "lblFirst"
        Me.lblFirst.Size = New System.Drawing.Size(26, 13)
        Me.lblFirst.TabIndex = 28
        Me.lblFirst.Text = "First"
        '
        'lblLast
        '
        Me.lblLast.AutoSize = True
        Me.lblLast.Location = New System.Drawing.Point(221, 138)
        Me.lblLast.Name = "lblLast"
        Me.lblLast.Size = New System.Drawing.Size(27, 13)
        Me.lblLast.TabIndex = 29
        Me.lblLast.Text = "Last"
        '
        'lblEmail
        '
        Me.lblEmail.AutoSize = True
        Me.lblEmail.Location = New System.Drawing.Point(444, 138)
        Me.lblEmail.Name = "lblEmail"
        Me.lblEmail.Size = New System.Drawing.Size(36, 13)
        Me.lblEmail.TabIndex = 30
        Me.lblEmail.Text = "E-Mail"
        '
        'CustomerFirst
        '
        Me.CustomerFirst.Font = New System.Drawing.Font("Arial", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CustomerFirst.Location = New System.Drawing.Point(5, 154)
        Me.CustomerFirst.Name = "CustomerFirst"
        Me.CustomerFirst.Size = New System.Drawing.Size(203, 25)
        Me.CustomerFirst.TabIndex = 0
        '
        'CustomerLast
        '
        Me.CustomerLast.Font = New System.Drawing.Font("Arial", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CustomerLast.Location = New System.Drawing.Point(222, 154)
        Me.CustomerLast.Name = "CustomerLast"
        Me.CustomerLast.Size = New System.Drawing.Size(203, 25)
        Me.CustomerLast.TabIndex = 1
        '
        'Email
        '
        Me.Email.Font = New System.Drawing.Font("Arial", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Email.ForeColor = System.Drawing.Color.Blue
        Me.Email.Location = New System.Drawing.Point(444, 154)
        Me.Email.Name = "Email"
        Me.Email.Size = New System.Drawing.Size(203, 25)
        Me.Email.TabIndex = 2
        '
        'CustomerAddress
        '
        Me.CustomerAddress.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CustomerAddress.Location = New System.Drawing.Point(5, 200)
        Me.CustomerAddress.Name = "CustomerAddress"
        Me.CustomerAddress.Size = New System.Drawing.Size(417, 22)
        Me.CustomerAddress.TabIndex = 3
        '
        'ShipToFirst
        '
        Me.ShipToFirst.Font = New System.Drawing.Font("Arial", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ShipToFirst.Location = New System.Drawing.Point(444, 200)
        Me.ShipToFirst.Name = "ShipToFirst"
        Me.ShipToFirst.Size = New System.Drawing.Size(203, 25)
        Me.ShipToFirst.TabIndex = 20
        '
        'ShipToLast
        '
        Me.ShipToLast.Font = New System.Drawing.Font("Arial", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ShipToLast.Location = New System.Drawing.Point(661, 200)
        Me.ShipToLast.Name = "ShipToLast"
        Me.ShipToLast.Size = New System.Drawing.Size(203, 25)
        Me.ShipToLast.TabIndex = 21
        '
        'lblAddr
        '
        Me.lblAddr.AutoSize = True
        Me.lblAddr.Location = New System.Drawing.Point(5, 185)
        Me.lblAddr.Name = "lblAddr"
        Me.lblAddr.Size = New System.Drawing.Size(45, 13)
        Me.lblAddr.TabIndex = 37
        Me.lblAddr.Text = "Address"
        '
        'lblShipToAddressCaption
        '
        Me.lblShipToAddressCaption.AutoSize = True
        Me.lblShipToAddressCaption.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblShipToAddressCaption.Location = New System.Drawing.Point(594, 185)
        Me.lblShipToAddressCaption.Name = "lblShipToAddressCaption"
        Me.lblShipToAddressCaption.Size = New System.Drawing.Size(101, 14)
        Me.lblShipToAddressCaption.TabIndex = 38
        Me.lblShipToAddressCaption.Text = "SHIP TO ADDRESS"
        '
        'lblShipFirst
        '
        Me.lblShipFirst.AutoSize = True
        Me.lblShipFirst.Location = New System.Drawing.Point(444, 185)
        Me.lblShipFirst.Name = "lblShipFirst"
        Me.lblShipFirst.Size = New System.Drawing.Size(26, 13)
        Me.lblShipFirst.TabIndex = 39
        Me.lblShipFirst.Text = "First"
        '
        'lblShipLast
        '
        Me.lblShipLast.AutoSize = True
        Me.lblShipLast.Location = New System.Drawing.Point(788, 185)
        Me.lblShipLast.Name = "lblShipLast"
        Me.lblShipLast.Size = New System.Drawing.Size(76, 13)
        Me.lblShipLast.TabIndex = 40
        Me.lblShipLast.Text = "Last/Company"
        '
        'AddAddress
        '
        Me.AddAddress.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.AddAddress.Location = New System.Drawing.Point(5, 245)
        Me.AddAddress.Name = "AddAddress"
        Me.AddAddress.Size = New System.Drawing.Size(417, 22)
        Me.AddAddress.TabIndex = 4
        '
        'CustomerAddress2
        '
        Me.CustomerAddress2.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CustomerAddress2.Location = New System.Drawing.Point(444, 245)
        Me.CustomerAddress2.Name = "CustomerAddress2"
        Me.CustomerAddress2.Size = New System.Drawing.Size(419, 22)
        Me.CustomerAddress2.TabIndex = 22
        '
        'lblAddAddr
        '
        Me.lblAddAddr.AutoSize = True
        Me.lblAddAddr.Location = New System.Drawing.Point(5, 229)
        Me.lblAddAddr.Name = "lblAddAddr"
        Me.lblAddAddr.Size = New System.Drawing.Size(94, 13)
        Me.lblAddAddr.TabIndex = 43
        Me.lblAddAddr.Text = "Additional Address"
        '
        'lblShipAddr
        '
        Me.lblShipAddr.AutoSize = True
        Me.lblShipAddr.Location = New System.Drawing.Point(444, 229)
        Me.lblShipAddr.Name = "lblShipAddr"
        Me.lblShipAddr.Size = New System.Drawing.Size(45, 13)
        Me.lblShipAddr.TabIndex = 44
        Me.lblShipAddr.Text = "Address"
        '
        'CustomerCity
        '
        Me.CustomerCity.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CustomerCity.Location = New System.Drawing.Point(5, 290)
        Me.CustomerCity.Name = "CustomerCity"
        Me.CustomerCity.Size = New System.Drawing.Size(294, 22)
        Me.CustomerCity.TabIndex = 5
        '
        'CustomerZip
        '
        Me.CustomerZip.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CustomerZip.Location = New System.Drawing.Point(308, 290)
        Me.CustomerZip.Name = "CustomerZip"
        Me.CustomerZip.Size = New System.Drawing.Size(117, 22)
        Me.CustomerZip.TabIndex = 6
        '
        'CustomerCity2
        '
        Me.CustomerCity2.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CustomerCity2.Location = New System.Drawing.Point(444, 290)
        Me.CustomerCity2.Name = "CustomerCity2"
        Me.CustomerCity2.Size = New System.Drawing.Size(279, 22)
        Me.CustomerCity2.TabIndex = 23
        '
        'CustomerZip2
        '
        Me.CustomerZip2.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CustomerZip2.Location = New System.Drawing.Point(747, 290)
        Me.CustomerZip2.Name = "CustomerZip2"
        Me.CustomerZip2.Size = New System.Drawing.Size(117, 22)
        Me.CustomerZip2.TabIndex = 24
        '
        'lblCity
        '
        Me.lblCity.AutoSize = True
        Me.lblCity.Location = New System.Drawing.Point(5, 274)
        Me.lblCity.Name = "lblCity"
        Me.lblCity.Size = New System.Drawing.Size(60, 13)
        Me.lblCity.TabIndex = 49
        Me.lblCity.Text = "City / State"
        '
        'lblZip
        '
        Me.lblZip.AutoSize = True
        Me.lblZip.Location = New System.Drawing.Point(308, 274)
        Me.lblZip.Name = "lblZip"
        Me.lblZip.Size = New System.Drawing.Size(22, 13)
        Me.lblZip.TabIndex = 50
        Me.lblZip.Text = "Zip"
        '
        'lblShipCity
        '
        Me.lblShipCity.AutoSize = True
        Me.lblShipCity.Location = New System.Drawing.Point(444, 274)
        Me.lblShipCity.Name = "lblShipCity"
        Me.lblShipCity.Size = New System.Drawing.Size(60, 13)
        Me.lblShipCity.TabIndex = 51
        Me.lblShipCity.Text = "City / State"
        '
        'lblShipZip
        '
        Me.lblShipZip.AutoSize = True
        Me.lblShipZip.Location = New System.Drawing.Point(744, 274)
        Me.lblShipZip.Name = "lblShipZip"
        Me.lblShipZip.Size = New System.Drawing.Size(22, 13)
        Me.lblShipZip.TabIndex = 52
        Me.lblShipZip.Text = "Zip"
        '
        'cboPhone1
        '
        Me.cboPhone1.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboPhone1.FormattingEnabled = True
        Me.cboPhone1.Items.AddRange(New Object() {"Home Phone", "Cell Phone", "Work Phone"})
        Me.cboPhone1.Location = New System.Drawing.Point(5, 327)
        Me.cboPhone1.Name = "cboPhone1"
        Me.cboPhone1.Size = New System.Drawing.Size(121, 22)
        Me.cboPhone1.TabIndex = 7
        Me.cboPhone1.Text = "Home Phone"
        '
        'cboPhone2
        '
        Me.cboPhone2.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboPhone2.FormattingEnabled = True
        Me.cboPhone2.Items.AddRange(New Object() {"Home Phone", "Cell Phone", "Work Phone"})
        Me.cboPhone2.Location = New System.Drawing.Point(222, 327)
        Me.cboPhone2.Name = "cboPhone2"
        Me.cboPhone2.Size = New System.Drawing.Size(121, 22)
        Me.cboPhone2.TabIndex = 9
        Me.cboPhone2.Text = "Telephone2"
        '
        'cboPhone3
        '
        Me.cboPhone3.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboPhone3.FormattingEnabled = True
        Me.cboPhone3.Items.AddRange(New Object() {"Home Phone", "Cell Phone", "Work Phone"})
        Me.cboPhone3.Location = New System.Drawing.Point(444, 327)
        Me.cboPhone3.Name = "cboPhone3"
        Me.cboPhone3.Size = New System.Drawing.Size(121, 22)
        Me.cboPhone3.TabIndex = 11
        Me.cboPhone3.Text = "Telephone3"
        '
        'CustomerPhone1
        '
        Me.CustomerPhone1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CustomerPhone1.Location = New System.Drawing.Point(5, 350)
        Me.CustomerPhone1.Name = "CustomerPhone1"
        Me.CustomerPhone1.Size = New System.Drawing.Size(203, 22)
        Me.CustomerPhone1.TabIndex = 8
        '
        'CustomerPhone2
        '
        Me.CustomerPhone2.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CustomerPhone2.Location = New System.Drawing.Point(222, 350)
        Me.CustomerPhone2.Name = "CustomerPhone2"
        Me.CustomerPhone2.Size = New System.Drawing.Size(203, 22)
        Me.CustomerPhone2.TabIndex = 10
        '
        'CustomerPhone3
        '
        Me.CustomerPhone3.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CustomerPhone3.Location = New System.Drawing.Point(444, 350)
        Me.CustomerPhone3.Name = "CustomerPhone3"
        Me.CustomerPhone3.Size = New System.Drawing.Size(188, 22)
        Me.CustomerPhone3.TabIndex = 12
        '
        'txtSpecInst
        '
        Me.txtSpecInst.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtSpecInst.Location = New System.Drawing.Point(5, 400)
        Me.txtSpecInst.Multiline = True
        Me.txtSpecInst.Name = "txtSpecInst"
        Me.txtSpecInst.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.txtSpecInst.Size = New System.Drawing.Size(639, 38)
        Me.txtSpecInst.TabIndex = 13
        '
        'lblSpecInstr
        '
        Me.lblSpecInstr.AutoSize = True
        Me.lblSpecInstr.Location = New System.Drawing.Point(5, 384)
        Me.lblSpecInstr.Name = "lblSpecInstr"
        Me.lblSpecInstr.Size = New System.Drawing.Size(99, 13)
        Me.lblSpecInstr.TabIndex = 60
        Me.lblSpecInstr.Text = "Special Instructions"
        '
        'Sales1
        '
        Me.Sales1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Sales1.Location = New System.Drawing.Point(5, 458)
        Me.Sales1.Name = "Sales1"
        Me.Sales1.Size = New System.Drawing.Size(127, 22)
        Me.Sales1.TabIndex = 14
        '
        'Sales2
        '
        Me.Sales2.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Sales2.Location = New System.Drawing.Point(219, 458)
        Me.Sales2.Name = "Sales2"
        Me.Sales2.Size = New System.Drawing.Size(127, 22)
        Me.Sales2.TabIndex = 16
        '
        'Sales3
        '
        Me.Sales3.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Sales3.Location = New System.Drawing.Point(436, 458)
        Me.Sales3.Name = "Sales3"
        Me.Sales3.Size = New System.Drawing.Size(127, 22)
        Me.Sales3.TabIndex = 18
        '
        'SalesSplit1
        '
        Me.SalesSplit1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.SalesSplit1.FormattingEnabled = True
        Me.SalesSplit1.Location = New System.Drawing.Point(145, 458)
        Me.SalesSplit1.Name = "SalesSplit1"
        Me.SalesSplit1.Size = New System.Drawing.Size(64, 21)
        Me.SalesSplit1.TabIndex = 15
        '
        'SalesSplit2
        '
        Me.SalesSplit2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.SalesSplit2.FormattingEnabled = True
        Me.SalesSplit2.Location = New System.Drawing.Point(356, 458)
        Me.SalesSplit2.Name = "SalesSplit2"
        Me.SalesSplit2.Size = New System.Drawing.Size(70, 21)
        Me.SalesSplit2.TabIndex = 17
        '
        'SalesSplit3
        '
        Me.SalesSplit3.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.SalesSplit3.FormattingEnabled = True
        Me.SalesSplit3.Location = New System.Drawing.Point(573, 458)
        Me.SalesSplit3.Name = "SalesSplit3"
        Me.SalesSplit3.Size = New System.Drawing.Size(70, 21)
        Me.SalesSplit3.TabIndex = 19
        '
        'lblSales1
        '
        Me.lblSales1.AutoSize = True
        Me.lblSales1.Location = New System.Drawing.Point(5, 443)
        Me.lblSales1.Name = "lblSales1"
        Me.lblSales1.Size = New System.Drawing.Size(42, 13)
        Me.lblSales1.TabIndex = 67
        Me.lblSales1.Text = "Sales 1"
        '
        'lblSales2
        '
        Me.lblSales2.AutoSize = True
        Me.lblSales2.Location = New System.Drawing.Point(219, 443)
        Me.lblSales2.Name = "lblSales2"
        Me.lblSales2.Size = New System.Drawing.Size(42, 13)
        Me.lblSales2.TabIndex = 68
        Me.lblSales2.Text = "Sales 2"
        '
        'lblSales3
        '
        Me.lblSales3.AutoSize = True
        Me.lblSales3.Location = New System.Drawing.Point(435, 443)
        Me.lblSales3.Name = "lblSales3"
        Me.lblSales3.Size = New System.Drawing.Size(42, 13)
        Me.lblSales3.TabIndex = 69
        Me.lblSales3.Text = "Sales 3"
        '
        'cmdShowBodyOfSale
        '
        Me.cmdShowBodyOfSale.Location = New System.Drawing.Point(585, 856)
        Me.cmdShowBodyOfSale.Name = "cmdShowBodyOfSale"
        Me.cmdShowBodyOfSale.Size = New System.Drawing.Size(110, 23)
        Me.cmdShowBodyOfSale.TabIndex = 72
        Me.cmdShowBodyOfSale.Text = "&Show Body Of Sale"
        Me.cmdShowBodyOfSale.UseVisualStyleBackColor = True
        Me.cmdShowBodyOfSale.Visible = False
        '
        'fraTimeWindow
        '
        Me.fraTimeWindow.Controls.Add(Me.dtpDelWindow2)
        Me.fraTimeWindow.Controls.Add(Me.lblTimeWindow)
        Me.fraTimeWindow.Controls.Add(Me.dtpDelWindow)
        Me.fraTimeWindow.Location = New System.Drawing.Point(8, 93)
        Me.fraTimeWindow.Name = "fraTimeWindow"
        Me.fraTimeWindow.Size = New System.Drawing.Size(213, 31)
        Me.fraTimeWindow.TabIndex = 74
        Me.fraTimeWindow.TabStop = False
        Me.fraTimeWindow.Visible = False
        '
        'dtpDelWindow2
        '
        Me.dtpDelWindow2.CustomFormat = ""
        Me.dtpDelWindow2.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dtpDelWindow2.Format = System.Windows.Forms.DateTimePickerFormat.Time
        Me.dtpDelWindow2.Location = New System.Drawing.Point(122, 7)
        Me.dtpDelWindow2.Name = "dtpDelWindow2"
        Me.dtpDelWindow2.ShowUpDown = True
        Me.dtpDelWindow2.Size = New System.Drawing.Size(86, 21)
        Me.dtpDelWindow2.TabIndex = 42
        '
        'lblTimeWindow
        '
        Me.lblTimeWindow.AutoSize = True
        Me.lblTimeWindow.Location = New System.Drawing.Point(96, 16)
        Me.lblTimeWindow.Name = "lblTimeWindow"
        Me.lblTimeWindow.Size = New System.Drawing.Size(20, 13)
        Me.lblTimeWindow.TabIndex = 28
        Me.lblTimeWindow.Text = "To"
        '
        'dtpDelWindow
        '
        Me.dtpDelWindow.CustomFormat = ""
        Me.dtpDelWindow.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dtpDelWindow.Format = System.Windows.Forms.DateTimePickerFormat.Time
        Me.dtpDelWindow.Location = New System.Drawing.Point(6, 7)
        Me.dtpDelWindow.Name = "dtpDelWindow"
        Me.dtpDelWindow.ShowUpDown = True
        Me.dtpDelWindow.Size = New System.Drawing.Size(86, 21)
        Me.dtpDelWindow.TabIndex = 41
        '
        'lblDelDate
        '
        Me.lblDelDate.BackColor = System.Drawing.Color.White
        Me.lblDelDate.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblDelDate.Location = New System.Drawing.Point(77, 69)
        Me.lblDelDate.Name = "lblDelDate"
        Me.lblDelDate.Size = New System.Drawing.Size(101, 20)
        Me.lblDelDate.TabIndex = 75
        Me.lblDelDate.Text = "lblDelDate"
        '
        'opt30323
        '
        Me.opt30323.AutoSize = True
        Me.opt30323.Checked = True
        Me.opt30323.Location = New System.Drawing.Point(63, 37)
        Me.opt30323.Name = "opt30323"
        Me.opt30323.Size = New System.Drawing.Size(55, 17)
        Me.opt30323.TabIndex = 30
        Me.opt30323.TabStop = True
        Me.opt30323.Text = "30323"
        Me.ToolTipBillOSale.SetToolTip(Me.opt30323, "Click this for the wider DYMO Shipping labels.")
        Me.opt30323.UseVisualStyleBackColor = True
        '
        'opt30252
        '
        Me.opt30252.AutoSize = True
        Me.opt30252.Location = New System.Drawing.Point(6, 37)
        Me.opt30252.Name = "opt30252"
        Me.opt30252.Size = New System.Drawing.Size(55, 17)
        Me.opt30252.TabIndex = 29
        Me.opt30252.Text = "30252"
        Me.ToolTipBillOSale.SetToolTip(Me.opt30252, "Select this option for narrow DYMO address labels.")
        Me.opt30252.UseVisualStyleBackColor = True
        '
        'cmdPrintLabel
        '
        Me.cmdPrintLabel.Location = New System.Drawing.Point(124, 88)
        Me.cmdPrintLabel.Name = "cmdPrintLabel"
        Me.cmdPrintLabel.Size = New System.Drawing.Size(66, 27)
        Me.cmdPrintLabel.TabIndex = 31
        Me.cmdPrintLabel.Text = "&Print Label"
        Me.ToolTipBillOSale.SetToolTip(Me.cmdPrintLabel, " Dymo 330 Turbo ")
        Me.cmdPrintLabel.UseVisualStyleBackColor = True
        '
        'cmdCancel
        '
        Me.cmdCancel.Location = New System.Drawing.Point(94, 11)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.Size = New System.Drawing.Size(65, 59)
        Me.cmdCancel.TabIndex = 27
        Me.cmdCancel.Text = "&Cancel"
        Me.ToolTipBillOSale.SetToolTip(Me.cmdCancel, " Cancels the sale ")
        Me.cmdCancel.UseVisualStyleBackColor = True
        '
        'cmdApplyBillOSale
        '
        Me.cmdApplyBillOSale.Location = New System.Drawing.Point(23, 11)
        Me.cmdApplyBillOSale.Name = "cmdApplyBillOSale"
        Me.cmdApplyBillOSale.Size = New System.Drawing.Size(65, 59)
        Me.cmdApplyBillOSale.TabIndex = 26
        Me.cmdApplyBillOSale.Text = "&OK"
        Me.ToolTipBillOSale.SetToolTip(Me.cmdApplyBillOSale, "Processes the mailing list section ")
        Me.cmdApplyBillOSale.UseVisualStyleBackColor = True
        '
        'ScanUp
        '
        Me.ScanUp.Location = New System.Drawing.Point(676, 656)
        Me.ScanUp.Name = "ScanUp"
        Me.ScanUp.Size = New System.Drawing.Size(75, 23)
        Me.ScanUp.TabIndex = 80
        Me.ScanUp.Text = "Scan up"
        Me.ToolTipBillOSale.SetToolTip(Me.ScanUp, " Scans up to next sale ")
        Me.ScanUp.UseVisualStyleBackColor = True
        Me.ScanUp.Visible = False
        '
        'cmdProcessSale
        '
        Me.cmdProcessSale.Location = New System.Drawing.Point(9, 13)
        Me.cmdProcessSale.Name = "cmdProcessSale"
        Me.cmdProcessSale.Size = New System.Drawing.Size(78, 55)
        Me.cmdProcessSale.TabIndex = 105
        Me.cmdProcessSale.Text = "&Process Sale"
        Me.ToolTipBillOSale.SetToolTip(Me.cmdProcessSale, " Prints sale and posts to all modules  ")
        Me.cmdProcessSale.UseVisualStyleBackColor = True
        '
        'cmdNextSale
        '
        Me.cmdNextSale.Location = New System.Drawing.Point(92, 13)
        Me.cmdNextSale.Name = "cmdNextSale"
        Me.cmdNextSale.Size = New System.Drawing.Size(75, 55)
        Me.cmdNextSale.TabIndex = 105
        Me.cmdNextSale.Text = "Ne&xt Sale"
        Me.ToolTipBillOSale.SetToolTip(Me.cmdNextSale, " Enter the next sale ")
        Me.cmdNextSale.UseVisualStyleBackColor = True
        '
        'cmdMainMenu
        '
        Me.cmdMainMenu.Location = New System.Drawing.Point(172, 13)
        Me.cmdMainMenu.Name = "cmdMainMenu"
        Me.cmdMainMenu.Size = New System.Drawing.Size(75, 55)
        Me.cmdMainMenu.TabIndex = 105
        Me.cmdMainMenu.Text = "&Main Menu"
        Me.ToolTipBillOSale.SetToolTip(Me.cmdMainMenu, "Return to Main Menu ")
        Me.cmdMainMenu.UseVisualStyleBackColor = True
        '
        'cmdClear
        '
        Me.cmdClear.Location = New System.Drawing.Point(252, 13)
        Me.cmdClear.Name = "cmdClear"
        Me.cmdClear.Size = New System.Drawing.Size(75, 55)
        Me.cmdClear.TabIndex = 105
        Me.cmdClear.Text = "&Clear"
        Me.ToolTipBillOSale.SetToolTip(Me.cmdClear, "Clears only the inventory entries ")
        Me.cmdClear.UseVisualStyleBackColor = True
        '
        'ScanDn
        '
        Me.ScanDn.Location = New System.Drawing.Point(826, 137)
        Me.ScanDn.Name = "ScanDn"
        Me.ScanDn.Size = New System.Drawing.Size(28, 23)
        Me.ScanDn.TabIndex = 79
        Me.ToolTipBillOSale.SetToolTip(Me.ScanDn, "Scans down to previous sale ")
        Me.ScanDn.UseVisualStyleBackColor = True
        '
        'ScanUp123
        '
        Me.ScanUp123.AutoSize = True
        Me.ScanUp123.Location = New System.Drawing.Point(826, 116)
        Me.ScanUp123.Name = "ScanUp123"
        Me.ScanUp123.Size = New System.Drawing.Size(28, 23)
        Me.ScanUp123.TabIndex = 80
        Me.ToolTipBillOSale.SetToolTip(Me.ScanUp123, " Scans up to next sale ")
        Me.ScanUp123.UseVisualStyleBackColor = True
        '
        'cmdSoldTags
        '
        Me.cmdSoldTags.Location = New System.Drawing.Point(826, 226)
        Me.cmdSoldTags.Name = "cmdSoldTags"
        Me.cmdSoldTags.Size = New System.Drawing.Size(28, 23)
        Me.cmdSoldTags.TabIndex = 85
        Me.ToolTipBillOSale.SetToolTip(Me.cmdSoldTags, "Email current copy of sale to customer's email address.")
        Me.cmdSoldTags.UseVisualStyleBackColor = True
        '
        'cmdEmail
        '
        Me.cmdEmail.Location = New System.Drawing.Point(826, 196)
        Me.cmdEmail.Name = "cmdEmail"
        Me.cmdEmail.Size = New System.Drawing.Size(28, 23)
        Me.cmdEmail.TabIndex = 87
        Me.ToolTipBillOSale.SetToolTip(Me.cmdEmail, "Email current copy of sale to customer's email address.")
        Me.cmdEmail.UseVisualStyleBackColor = True
        '
        'cmdPrint
        '
        Me.cmdPrint.Location = New System.Drawing.Point(826, 175)
        Me.cmdPrint.Name = "cmdPrint"
        Me.cmdPrint.Size = New System.Drawing.Size(28, 23)
        Me.cmdPrint.TabIndex = 86
        Me.ToolTipBillOSale.SetToolTip(Me.cmdPrint, " Print current copy of sale. ")
        Me.cmdPrint.UseVisualStyleBackColor = True
        '
        'cmdChangePrice
        '
        Me.cmdChangePrice.AutoSize = True
        Me.cmdChangePrice.Location = New System.Drawing.Point(826, 80)
        Me.cmdChangePrice.Name = "cmdChangePrice"
        Me.cmdChangePrice.Size = New System.Drawing.Size(28, 23)
        Me.cmdChangePrice.TabIndex = 89
        Me.ToolTipBillOSale.SetToolTip(Me.cmdChangePrice, "Apply a Discount.")
        Me.cmdChangePrice.UseVisualStyleBackColor = True
        Me.cmdChangePrice.Visible = False
        '
        'imgCalendar
        '
        Me.imgCalendar.BackColor = System.Drawing.SystemColors.ButtonShadow
        Me.imgCalendar.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.imgCalendar.Image = Global.WinCDS.My.Resources.Resources.calendar
        Me.imgCalendar.Location = New System.Drawing.Point(180, 39)
        Me.imgCalendar.Name = "imgCalendar"
        Me.imgCalendar.Size = New System.Drawing.Size(36, 27)
        Me.imgCalendar.TabIndex = 6
        Me.imgCalendar.TabStop = False
        Me.ToolTipBillOSale.SetToolTip(Me.imgCalendar, "View the Delivery Calendar")
        '
        'cmdNoChangePrice
        '
        Me.cmdNoChangePrice.Location = New System.Drawing.Point(826, 81)
        Me.cmdNoChangePrice.Name = "cmdNoChangePrice"
        Me.cmdNoChangePrice.Size = New System.Drawing.Size(28, 23)
        Me.cmdNoChangePrice.TabIndex = 98
        Me.ToolTipBillOSale.SetToolTip(Me.cmdNoChangePrice, "Disable Changing Prices")
        Me.cmdNoChangePrice.UseVisualStyleBackColor = True
        Me.cmdNoChangePrice.Visible = False
        '
        'fraBOS2
        '
        Me.fraBOS2.BackColor = System.Drawing.SystemColors.Control
        Me.fraBOS2.Controls.Add(Me.UGridIO1)
        Me.fraBOS2.Controls.Add(Me.rtbStorePolicy)
        Me.fraBOS2.Controls.Add(Me.rtb)
        Me.fraBOS2.Controls.Add(Me.fraBOS2Commands)
        Me.fraBOS2.Controls.Add(Me.ScanDn)
        Me.fraBOS2.Controls.Add(Me.fraHover)
        Me.fraBOS2.Controls.Add(Me.ScanUp123)
        Me.fraBOS2.Controls.Add(Me.cmdNoChangePrice)
        Me.fraBOS2.Controls.Add(Me.lblBalDueCaption)
        Me.fraBOS2.Controls.Add(Me.cmdSoldTags)
        Me.fraBOS2.Controls.Add(Me.cmdEmail)
        Me.fraBOS2.Controls.Add(Me.BalDue)
        Me.fraBOS2.Controls.Add(Me.cmdPrint)
        Me.fraBOS2.Controls.Add(Me.cmdChangePrice)
        Me.fraBOS2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.fraBOS2.Location = New System.Drawing.Point(0, 182)
        Me.fraBOS2.Name = "fraBOS2"
        Me.fraBOS2.Size = New System.Drawing.Size(864, 408)
        Me.fraBOS2.TabIndex = 88
        Me.fraBOS2.TabStop = False
        '
        'UGridIO1
        '
        Me.UGridIO1.Activated = False
        Me.UGridIO1.Col = 0
        Me.UGridIO1.firstrow = 1
        Me.UGridIO1.Loading = False
        Me.UGridIO1.Location = New System.Drawing.Point(6, 18)
        Me.UGridIO1.MaxCols = 2
        Me.UGridIO1.MaxRows = 10
        Me.UGridIO1.Name = "UGridIO1"
        Me.UGridIO1.Row = 0
        Me.UGridIO1.Size = New System.Drawing.Size(810, 306)
        Me.UGridIO1.TabIndex = 104
        '
        'rtbStorePolicy
        '
        Me.rtbStorePolicy.Location = New System.Drawing.Point(85, 347)
        Me.rtbStorePolicy.Name = "rtbStorePolicy"
        Me.rtbStorePolicy.Size = New System.Drawing.Size(57, 48)
        Me.rtbStorePolicy.TabIndex = 106
        Me.rtbStorePolicy.Visible = False
        '
        'rtb
        '
        Me.rtb.Location = New System.Drawing.Point(17, 347)
        Me.rtb.Name = "rtb"
        Me.rtb.Size = New System.Drawing.Size(53, 48)
        Me.rtb.TabIndex = 105
        Me.rtb.Visible = False
        '
        'fraBOS2Commands
        '
        Me.fraBOS2Commands.Controls.Add(Me.Notes_Open)
        Me.fraBOS2Commands.Controls.Add(Me.cmdClear)
        Me.fraBOS2Commands.Controls.Add(Me.cmdMainMenu)
        Me.fraBOS2Commands.Controls.Add(Me.cmdNextSale)
        Me.fraBOS2Commands.Controls.Add(Me.cmdProcessSale)
        Me.fraBOS2Commands.Location = New System.Drawing.Point(203, 327)
        Me.fraBOS2Commands.Name = "fraBOS2Commands"
        Me.fraBOS2Commands.Size = New System.Drawing.Size(414, 75)
        Me.fraBOS2Commands.TabIndex = 104
        Me.fraBOS2Commands.TabStop = False
        '
        'Notes_Open
        '
        Me.Notes_Open.Location = New System.Drawing.Point(332, 13)
        Me.Notes_Open.Name = "Notes_Open"
        Me.Notes_Open.Size = New System.Drawing.Size(75, 55)
        Me.Notes_Open.TabIndex = 105
        Me.Notes_Open.Text = "&Notes"
        Me.Notes_Open.UseVisualStyleBackColor = True
        '
        'fraHover
        '
        Me.fraHover.BackColor = System.Drawing.SystemColors.Info
        Me.fraHover.Controls.Add(Me.picHover)
        Me.fraHover.Location = New System.Drawing.Point(233, 63)
        Me.fraHover.Name = "fraHover"
        Me.fraHover.Size = New System.Drawing.Size(200, 77)
        Me.fraHover.TabIndex = 94
        Me.fraHover.TabStop = False
        Me.fraHover.Visible = False
        '
        'picHover
        '
        Me.picHover.Location = New System.Drawing.Point(45, 16)
        Me.picHover.Name = "picHover"
        Me.picHover.Size = New System.Drawing.Size(100, 50)
        Me.picHover.TabIndex = 95
        Me.picHover.TabStop = False
        Me.picHover.Visible = False
        '
        'lblBalDueCaption
        '
        Me.lblBalDueCaption.AutoSize = True
        Me.lblBalDueCaption.Location = New System.Drawing.Point(665, 339)
        Me.lblBalDueCaption.Name = "lblBalDueCaption"
        Me.lblBalDueCaption.Size = New System.Drawing.Size(69, 13)
        Me.lblBalDueCaption.TabIndex = 91
        Me.lblBalDueCaption.Text = "Balance Due"
        Me.lblBalDueCaption.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'BalDue
        '
        Me.BalDue.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.BalDue.Location = New System.Drawing.Point(645, 355)
        Me.BalDue.Multiline = True
        Me.BalDue.Name = "BalDue"
        Me.BalDue.Size = New System.Drawing.Size(103, 38)
        Me.BalDue.TabIndex = 92
        Me.BalDue.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'BillOfSale
        '
        Me.BillOfSale.AutoSize = True
        Me.BillOfSale.Font = New System.Drawing.Font("Arial", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.BillOfSale.Location = New System.Drawing.Point(760, 14)
        Me.BillOfSale.Name = "BillOfSale"
        Me.BillOfSale.Size = New System.Drawing.Size(0, 24)
        Me.BillOfSale.TabIndex = 93
        '
        'tmrHover
        '
        '
        'tmrFormat
        '
        '
        'txtFormatHelper
        '
        Me.txtFormatHelper.Location = New System.Drawing.Point(657, 907)
        Me.txtFormatHelper.Name = "txtFormatHelper"
        Me.txtFormatHelper.Size = New System.Drawing.Size(100, 20)
        Me.txtFormatHelper.TabIndex = 96
        Me.txtFormatHelper.Visible = False
        '
        'fraButtons
        '
        Me.fraButtons.Controls.Add(Me.fraPrintType)
        Me.fraButtons.Controls.Add(Me.cmdPrintLabel)
        Me.fraButtons.Controls.Add(Me.cmdCancel)
        Me.fraButtons.Controls.Add(Me.cmdApplyBillOSale)
        Me.fraButtons.Location = New System.Drawing.Point(665, 341)
        Me.fraButtons.Name = "fraButtons"
        Me.fraButtons.Size = New System.Drawing.Size(198, 126)
        Me.fraButtons.TabIndex = 25
        Me.fraButtons.TabStop = False
        '
        'fraPrintType
        '
        Me.fraPrintType.Controls.Add(Me.opt30323)
        Me.fraPrintType.Controls.Add(Me.opt30252)
        Me.fraPrintType.Controls.Add(Me.lblPrintType)
        Me.fraPrintType.Location = New System.Drawing.Point(6, 69)
        Me.fraPrintType.Name = "fraPrintType"
        Me.fraPrintType.Size = New System.Drawing.Size(112, 51)
        Me.fraPrintType.TabIndex = 28
        Me.fraPrintType.TabStop = False
        '
        'lblPrintType
        '
        Me.lblPrintType.AutoSize = True
        Me.lblPrintType.Location = New System.Drawing.Point(6, 16)
        Me.lblPrintType.Name = "lblPrintType"
        Me.lblPrintType.Size = New System.Drawing.Size(60, 13)
        Me.lblPrintType.TabIndex = 5
        Me.lblPrintType.Text = "Label Type"
        '
        'lblGrossSalesCaption
        '
        Me.lblGrossSalesCaption.AutoSize = True
        Me.lblGrossSalesCaption.Location = New System.Drawing.Point(686, 325)
        Me.lblGrossSalesCaption.Name = "lblGrossSalesCaption"
        Me.lblGrossSalesCaption.Size = New System.Drawing.Size(66, 13)
        Me.lblGrossSalesCaption.TabIndex = 103
        Me.lblGrossSalesCaption.Text = "Gross Sales:"
        '
        'GroupBox1
        '
        Me.GroupBox1.Location = New System.Drawing.Point(168, 201)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(414, 62)
        Me.GroupBox1.TabIndex = 104
        Me.GroupBox1.TabStop = False
        '
        'lblGrossSales
        '
        Me.lblGrossSales.AutoSize = True
        Me.lblGrossSales.Location = New System.Drawing.Point(748, 325)
        Me.lblGrossSales.Name = "lblGrossSales"
        Me.lblGrossSales.Size = New System.Drawing.Size(0, 13)
        Me.lblGrossSales.TabIndex = 108
        '
        'picFormatHelper
        '
        Me.picFormatHelper.Location = New System.Drawing.Point(464, 856)
        Me.picFormatHelper.Name = "picFormatHelper"
        Me.picFormatHelper.Size = New System.Drawing.Size(100, 23)
        Me.picFormatHelper.TabIndex = 97
        Me.picFormatHelper.TabStop = False
        Me.picFormatHelper.Visible = False
        '
        'imgLogo
        '
        Me.imgLogo.Location = New System.Drawing.Point(230, 3)
        Me.imgLogo.Name = "imgLogo"
        Me.imgLogo.Size = New System.Drawing.Size(407, 119)
        Me.imgLogo.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.imgLogo.TabIndex = 11
        Me.imgLogo.TabStop = False
        '
        'ugrFake
        '
        Me.ugrFake.Activated = False
        Me.ugrFake.Col = 1
        Me.ugrFake.firstrow = 1
        Me.ugrFake.Loading = False
        Me.ugrFake.Location = New System.Drawing.Point(5, 608)
        Me.ugrFake.MaxCols = 2
        Me.ugrFake.MaxRows = 10
        Me.ugrFake.Name = "ugrFake"
        Me.ugrFake.Row = 0
        Me.ugrFake.Size = New System.Drawing.Size(817, 139)
        Me.ugrFake.TabIndex = 45
        Me.ugrFake.TabStop = False
        '
        'BillOSale
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(874, 881)
        Me.Controls.Add(Me.ugrFake)
        Me.Controls.Add(Me.fraBOS2)
        Me.Controls.Add(Me.lblGrossSalesCaption)
        Me.Controls.Add(Me.dteDelivery)
        Me.Controls.Add(Me.fraButtons)
        Me.Controls.Add(Me.picFormatHelper)
        Me.Controls.Add(Me.txtFormatHelper)
        Me.Controls.Add(Me.BillOfSale)
        Me.Controls.Add(Me.lblDelDate)
        Me.Controls.Add(Me.fraTimeWindow)
        Me.Controls.Add(Me.cmdShowBodyOfSale)
        Me.Controls.Add(Me.lblSales3)
        Me.Controls.Add(Me.lblSales2)
        Me.Controls.Add(Me.lblSales1)
        Me.Controls.Add(Me.SalesSplit3)
        Me.Controls.Add(Me.SalesSplit2)
        Me.Controls.Add(Me.SalesSplit1)
        Me.Controls.Add(Me.Sales3)
        Me.Controls.Add(Me.Sales2)
        Me.Controls.Add(Me.Sales1)
        Me.Controls.Add(Me.lblSpecInstr)
        Me.Controls.Add(Me.txtSpecInst)
        Me.Controls.Add(Me.CustomerPhone3)
        Me.Controls.Add(Me.CustomerPhone2)
        Me.Controls.Add(Me.CustomerPhone1)
        Me.Controls.Add(Me.cboPhone3)
        Me.Controls.Add(Me.cboPhone2)
        Me.Controls.Add(Me.cboPhone1)
        Me.Controls.Add(Me.lblShipZip)
        Me.Controls.Add(Me.lblShipCity)
        Me.Controls.Add(Me.lblZip)
        Me.Controls.Add(Me.lblCity)
        Me.Controls.Add(Me.CustomerZip2)
        Me.Controls.Add(Me.CustomerCity2)
        Me.Controls.Add(Me.CustomerZip)
        Me.Controls.Add(Me.CustomerCity)
        Me.Controls.Add(Me.lblShipAddr)
        Me.Controls.Add(Me.lblAddAddr)
        Me.Controls.Add(Me.CustomerAddress2)
        Me.Controls.Add(Me.AddAddress)
        Me.Controls.Add(Me.lblShipLast)
        Me.Controls.Add(Me.lblShipFirst)
        Me.Controls.Add(Me.lblShipToAddressCaption)
        Me.Controls.Add(Me.lblAddr)
        Me.Controls.Add(Me.ShipToLast)
        Me.Controls.Add(Me.ShipToFirst)
        Me.Controls.Add(Me.CustomerAddress)
        Me.Controls.Add(Me.Email)
        Me.Controls.Add(Me.CustomerLast)
        Me.Controls.Add(Me.CustomerFirst)
        Me.Controls.Add(Me.lblEmail)
        Me.Controls.Add(Me.lblLast)
        Me.Controls.Add(Me.lblFirst)
        Me.Controls.Add(Me.cboTaxZone)
        Me.Controls.Add(Me.cboAdvertisingType)
        Me.Controls.Add(Me.cboCustType)
        Me.Controls.Add(Me.SaleStatus)
        Me.Controls.Add(Me.lblTaxCaption)
        Me.Controls.Add(Me.lblAdvCaption)
        Me.Controls.Add(Me.lblTypeCaption)
        Me.Controls.Add(Me.lblStatusCaption)
        Me.Controls.Add(Me.lblSaleNoCaption)
        Me.Controls.Add(Me.txtSaleNo)
        Me.Controls.Add(Me.StorePhone)
        Me.Controls.Add(Me.StoreCity)
        Me.Controls.Add(Me.StoreAddress)
        Me.Controls.Add(Me.StoreName)
        Me.Controls.Add(Me.imgLogo)
        Me.Controls.Add(Me.optBusiness)
        Me.Controls.Add(Me.optIndividual)
        Me.Controls.Add(Me.imgCalendar)
        Me.Controls.Add(Me.chkPickup)
        Me.Controls.Add(Me.chkDelivery)
        Me.Controls.Add(Me.lblDelWeekday)
        Me.Controls.Add(Me.dteSaleDate)
        Me.Controls.Add(Me.lblDateCaption)
        Me.Controls.Add(Me.lblGrossSales)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Location = New System.Drawing.Point(100, 0)
        Me.MaximizeBox = False
        Me.Name = "BillOSale"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "BILL OF SALE"
        Me.fraTimeWindow.ResumeLayout(False)
        Me.fraTimeWindow.PerformLayout()
        CType(Me.imgCalendar, System.ComponentModel.ISupportInitialize).EndInit()
        Me.fraBOS2.ResumeLayout(False)
        Me.fraBOS2.PerformLayout()
        Me.fraBOS2Commands.ResumeLayout(False)
        Me.fraHover.ResumeLayout(False)
        CType(Me.picHover, System.ComponentModel.ISupportInitialize).EndInit()
        Me.fraButtons.ResumeLayout(False)
        Me.fraPrintType.ResumeLayout(False)
        Me.fraPrintType.PerformLayout()
        CType(Me.picFormatHelper, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.imgLogo, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents lblDateCaption As Label
    Friend WithEvents dteSaleDate As DateTimePicker
    Friend WithEvents lblDelWeekday As Label
    Friend WithEvents chkDelivery As CheckBox
    Friend WithEvents chkPickup As CheckBox
    Friend WithEvents imgCalendar As PictureBox
    Friend WithEvents optIndividual As RadioButton
    Friend WithEvents optBusiness As RadioButton
    Friend WithEvents imgLogo As PictureBox
    Friend WithEvents StoreName As Label
    Friend WithEvents StoreAddress As Label
    Friend WithEvents StoreCity As Label
    Friend WithEvents StorePhone As Label
    Friend WithEvents txtSaleNo As TextBox
    Friend WithEvents lblSaleNoCaption As Label
    Friend WithEvents lblStatusCaption As Label
    Friend WithEvents lblTypeCaption As Label
    Friend WithEvents lblAdvCaption As Label
    Friend WithEvents lblTaxCaption As Label
    Friend WithEvents SaleStatus As TextBox
    Friend WithEvents cboCustType As ComboBox
    Friend WithEvents cboAdvertisingType As ComboBox
    Friend WithEvents cboTaxZone As ComboBox
    Friend WithEvents dteDelivery As DateTimePicker
    Friend WithEvents lblFirst As Label
    Friend WithEvents lblLast As Label
    Friend WithEvents lblEmail As Label
    Friend WithEvents CustomerFirst As TextBox
    Friend WithEvents CustomerLast As TextBox
    Friend WithEvents Email As TextBox
    Friend WithEvents CustomerAddress As TextBox
    Friend WithEvents ShipToFirst As TextBox
    Friend WithEvents ShipToLast As TextBox
    Friend WithEvents lblAddr As Label
    Friend WithEvents lblShipToAddressCaption As Label
    Friend WithEvents lblShipFirst As Label
    Friend WithEvents lblShipLast As Label
    Friend WithEvents AddAddress As TextBox
    Friend WithEvents CustomerAddress2 As TextBox
    Friend WithEvents lblAddAddr As Label
    Friend WithEvents lblShipAddr As Label
    Friend WithEvents CustomerCity As TextBox
    Friend WithEvents CustomerZip As TextBox
    Friend WithEvents CustomerCity2 As TextBox
    Friend WithEvents CustomerZip2 As TextBox
    Friend WithEvents lblCity As Label
    Friend WithEvents lblZip As Label
    Friend WithEvents lblShipCity As Label
    Friend WithEvents lblShipZip As Label
    Friend WithEvents cboPhone1 As ComboBox
    Friend WithEvents cboPhone2 As ComboBox
    Friend WithEvents cboPhone3 As ComboBox
    Friend WithEvents CustomerPhone1 As TextBox
    Friend WithEvents CustomerPhone2 As TextBox
    Friend WithEvents CustomerPhone3 As TextBox
    Friend WithEvents txtSpecInst As TextBox
    Friend WithEvents lblSpecInstr As Label
    Friend WithEvents Sales1 As TextBox
    Friend WithEvents Sales2 As TextBox
    Friend WithEvents Sales3 As TextBox
    Friend WithEvents SalesSplit1 As ComboBox
    Friend WithEvents SalesSplit2 As ComboBox
    Friend WithEvents SalesSplit3 As ComboBox
    Friend WithEvents lblSales1 As Label
    Friend WithEvents lblSales2 As Label
    Friend WithEvents lblSales3 As Label
    Friend WithEvents cmdShowBodyOfSale As Button
    Friend WithEvents fraTimeWindow As GroupBox
    Friend WithEvents dtpDelWindow2 As DateTimePicker
    Friend WithEvents lblTimeWindow As Label
    Friend WithEvents dtpDelWindow As DateTimePicker
    Friend WithEvents lblDelDate As Label
    Friend WithEvents ToolTipBillOSale As ToolTip
    Friend WithEvents ScanDn As Button
    Friend WithEvents ScanUp123 As Button
    Friend WithEvents cmdSoldTags As Button
    Friend WithEvents cmdPrint As Button
    Friend WithEvents cmdEmail As Button
    Friend WithEvents fraBOS2 As GroupBox
    Friend WithEvents cmdChangePrice As Button
    Friend WithEvents lblBalDueCaption As Label
    Friend WithEvents BalDue As TextBox
    Friend WithEvents BillOfSale As Label
    Friend WithEvents fraHover As GroupBox
    Friend WithEvents picHover As PictureBox
    Friend WithEvents tmrHover As Timer
    Friend WithEvents tmrFormat As Timer
    Friend WithEvents txtFormatHelper As TextBox
    Friend WithEvents picFormatHelper As PictureBox
    Friend WithEvents cmdNoChangePrice As Button
    'Friend WithEvents rtb As RichTextBoxNew
    'Friend WithEvents rtbStorePolicy As RichTextBoxNew
    Friend WithEvents fraButtons As GroupBox
    Friend WithEvents fraPrintType As GroupBox
    Friend WithEvents opt30323 As RadioButton
    Friend WithEvents opt30252 As RadioButton
    Friend WithEvents lblPrintType As Label
    Friend WithEvents cmdPrintLabel As Button
    Friend WithEvents cmdCancel As Button
    Friend WithEvents cmdApplyBillOSale As Button
    Friend WithEvents lblGrossSalesCaption As Label
    Friend WithEvents fraBOS2Commands As GroupBox
    Friend WithEvents ScanUp As Button
    Friend WithEvents GroupBox1 As GroupBox
    Friend WithEvents Notes_Open As Button
    Friend WithEvents cmdClear As Button
    Friend WithEvents cmdMainMenu As Button
    Friend WithEvents cmdNextSale As Button
    Friend WithEvents cmdProcessSale As Button
    Friend WithEvents UGridIO1 As UGridIO
    Friend WithEvents rtbStorePolicy As RichTextBoxNew
    Friend WithEvents rtb As RichTextBoxNew
    Friend WithEvents ugrFake As UGridIO
    Friend WithEvents lblGrossSales As Label
End Class
