<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class ServiceParts
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
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
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.optTagStock = New System.Windows.Forms.RadioButton()
        Me.optTagCustomer = New System.Windows.Forms.RadioButton()
        Me.fraCustomer = New System.Windows.Forms.GroupBox()
        Me.lblFirstName = New System.Windows.Forms.Label()
        Me.lblLastName = New System.Windows.Forms.Label()
        Me.lblAddress = New System.Windows.Forms.Label()
        Me.lblAddress2 = New System.Windows.Forms.Label()
        Me.lblCity = New System.Windows.Forms.Label()
        Me.lblZip = New System.Windows.Forms.Label()
        Me.lblTele1Caption = New System.Windows.Forms.Label()
        Me.lblTele3 = New System.Windows.Forms.Label()
        Me.lblTele2Caption = New System.Windows.Forms.Label()
        Me.lblTele2 = New System.Windows.Forms.Label()
        Me.lblTele3Caption = New System.Windows.Forms.Label()
        Me.lblTele = New System.Windows.Forms.Label()
        Me.lblInvoiceNo = New System.Windows.Forms.Label()
        Me.lblSaleNo = New System.Windows.Forms.Label()
        Me.lblServiceOrderNoCaption = New System.Windows.Forms.Label()
        Me.lblServiceOrderNo = New System.Windows.Forms.Label()
        Me.lblWhatToDoWStyle = New System.Windows.Forms.Label()
        Me.txtInvoiceNo = New System.Windows.Forms.TextBox()
        Me.txtSaleNo = New System.Windows.Forms.TextBox()
        Me.txtStoreName = New System.Windows.Forms.TextBox()
        Me.dteClaimDateCaption1 = New System.Windows.Forms.DateTimePicker()
        Me.dteClaimDate = New System.Windows.Forms.DateTimePicker()
        Me.cboStores = New System.Windows.Forms.ComboBox()
        Me.txtStoreAddress = New System.Windows.Forms.TextBox()
        Me.txtStoreCity = New System.Windows.Forms.TextBox()
        Me.txtStorePhone = New System.Windows.Forms.TextBox()
        Me.cmdMoveFirst = New System.Windows.Forms.Button()
        Me.cmdMovePrevious = New System.Windows.Forms.Button()
        Me.cmdMoveNext = New System.Windows.Forms.Button()
        Me.cmdMoveLast = New System.Windows.Forms.Button()
        Me.cmdMoveSearch = New System.Windows.Forms.Button()
        Me.lblMoveRecords = New System.Windows.Forms.Label()
        Me.txtStyleNo = New System.Windows.Forms.TextBox()
        Me.txtDescription = New System.Windows.Forms.TextBox()
        Me.lblMarginLine = New System.Windows.Forms.Label()
        Me.lblPartsOrderNo = New System.Windows.Forms.Label()
        Me.lblClaimDate = New System.Windows.Forms.Label()
        Me.txtVendorAddress = New System.Windows.Forms.TextBox()
        Me.txtVendorCity = New System.Windows.Forms.TextBox()
        Me.txtVendorTele = New System.Windows.Forms.TextBox()
        Me.txtRepairCost = New System.Windows.Forms.TextBox()
        Me.chkPaid = New System.Windows.Forms.CheckBox()
        Me.Notes_Text = New System.Windows.Forms.TextBox()
        Me.optCBCredit = New System.Windows.Forms.RadioButton()
        Me.optCBDeduct = New System.Windows.Forms.RadioButton()
        Me.optCBChargeBack = New System.Windows.Forms.RadioButton()
        Me.cboStatus = New System.Windows.Forms.ComboBox()
        Me.txtVendorEmail = New System.Windows.Forms.TextBox()
        Me.fraSoldTo = New System.Windows.Forms.GroupBox()
        Me.fraVendor = New System.Windows.Forms.GroupBox()
        Me.txtVendorName = New System.Windows.Forms.ComboBox()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.fraStatus = New System.Windows.Forms.GroupBox()
        Me.cmdPictures = New System.Windows.Forms.Button()
        Me.imgPicture = New System.Windows.Forms.PictureBox()
        Me.lblClaimDateCaption = New System.Windows.Forms.Label()
        Me.lblPartsOrderNoCaption = New System.Windows.Forms.Label()
        Me.lblRepairCost = New System.Windows.Forms.Label()
        Me.dteClaimDateCaption = New System.Windows.Forms.Label()
        Me.lblStatus = New System.Windows.Forms.Label()
        Me.fraNotes = New System.Windows.Forms.GroupBox()
        Me.cmdMenu = New System.Windows.Forms.Button()
        Me.cmdNext = New System.Windows.Forms.Button()
        Me.cmdEmail = New System.Windows.Forms.Button()
        Me.cmdPrint = New System.Windows.Forms.Button()
        Me.cmdSave = New System.Windows.Forms.Button()
        Me.lblDescription = New System.Windows.Forms.Label()
        Me.lblStyleNo = New System.Windows.Forms.Label()
        Me.lblMarginLinelbl = New System.Windows.Forms.Label()
        Me.cmdAddPart = New System.Windows.Forms.Button()
        Me.cmdPrintChargeBack = New System.Windows.Forms.Button()
        Me.fraCustomer.SuspendLayout()
        Me.fraSoldTo.SuspendLayout()
        Me.fraVendor.SuspendLayout()
        Me.fraStatus.SuspendLayout()
        CType(Me.imgPicture, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.fraNotes.SuspendLayout()
        Me.SuspendLayout()
        '
        'optTagStock
        '
        Me.optTagStock.AutoSize = True
        Me.optTagStock.Location = New System.Drawing.Point(332, 4)
        Me.optTagStock.Name = "optTagStock"
        Me.optTagStock.Size = New System.Drawing.Size(93, 17)
        Me.optTagStock.TabIndex = 0
        Me.optTagStock.Text = "Tag &For Stock"
        Me.optTagStock.UseVisualStyleBackColor = True
        '
        'optTagCustomer
        '
        Me.optTagCustomer.AutoSize = True
        Me.optTagCustomer.Checked = True
        Me.optTagCustomer.Location = New System.Drawing.Point(467, 4)
        Me.optTagCustomer.Name = "optTagCustomer"
        Me.optTagCustomer.Size = New System.Drawing.Size(112, 17)
        Me.optTagCustomer.TabIndex = 1
        Me.optTagCustomer.TabStop = True
        Me.optTagCustomer.Text = "&Tag For Customer:"
        Me.optTagCustomer.UseVisualStyleBackColor = True
        '
        'fraCustomer
        '
        Me.fraCustomer.Controls.Add(Me.lblFirstName)
        Me.fraCustomer.Controls.Add(Me.lblLastName)
        Me.fraCustomer.Controls.Add(Me.lblAddress)
        Me.fraCustomer.Controls.Add(Me.lblAddress2)
        Me.fraCustomer.Controls.Add(Me.lblCity)
        Me.fraCustomer.Controls.Add(Me.lblZip)
        Me.fraCustomer.Controls.Add(Me.lblTele1Caption)
        Me.fraCustomer.Controls.Add(Me.lblTele3)
        Me.fraCustomer.Controls.Add(Me.lblTele2Caption)
        Me.fraCustomer.Controls.Add(Me.lblTele2)
        Me.fraCustomer.Controls.Add(Me.lblTele3Caption)
        Me.fraCustomer.Controls.Add(Me.lblTele)
        Me.fraCustomer.Location = New System.Drawing.Point(277, 197)
        Me.fraCustomer.Name = "fraCustomer"
        Me.fraCustomer.Size = New System.Drawing.Size(436, 130)
        Me.fraCustomer.TabIndex = 2
        Me.fraCustomer.TabStop = False
        Me.fraCustomer.Text = "Customer Address:"
        '
        'lblFirstName
        '
        Me.lblFirstName.Location = New System.Drawing.Point(11, 16)
        Me.lblFirstName.Name = "lblFirstName"
        Me.lblFirstName.Size = New System.Drawing.Size(184, 19)
        Me.lblFirstName.TabIndex = 24
        Me.lblFirstName.Text = "1"
        '
        'lblLastName
        '
        Me.lblLastName.AutoSize = True
        Me.lblLastName.Location = New System.Drawing.Point(201, 13)
        Me.lblLastName.Name = "lblLastName"
        Me.lblLastName.Size = New System.Drawing.Size(65, 13)
        Me.lblLastName.TabIndex = 25
        Me.lblLastName.Text = "lblLastName"
        '
        'lblAddress
        '
        Me.lblAddress.AutoSize = True
        Me.lblAddress.Location = New System.Drawing.Point(11, 36)
        Me.lblAddress.Name = "lblAddress"
        Me.lblAddress.Size = New System.Drawing.Size(13, 13)
        Me.lblAddress.TabIndex = 26
        Me.lblAddress.Text = "2"
        '
        'lblAddress2
        '
        Me.lblAddress2.AutoSize = True
        Me.lblAddress2.Location = New System.Drawing.Point(11, 52)
        Me.lblAddress2.Name = "lblAddress2"
        Me.lblAddress2.Size = New System.Drawing.Size(13, 13)
        Me.lblAddress2.TabIndex = 27
        Me.lblAddress2.Text = "3"
        '
        'lblCity
        '
        Me.lblCity.AutoSize = True
        Me.lblCity.Location = New System.Drawing.Point(12, 70)
        Me.lblCity.Name = "lblCity"
        Me.lblCity.Size = New System.Drawing.Size(13, 13)
        Me.lblCity.TabIndex = 28
        Me.lblCity.Text = "4"
        '
        'lblZip
        '
        Me.lblZip.AutoSize = True
        Me.lblZip.Location = New System.Drawing.Point(201, 70)
        Me.lblZip.Name = "lblZip"
        Me.lblZip.Size = New System.Drawing.Size(32, 13)
        Me.lblZip.TabIndex = 29
        Me.lblZip.Text = "lblZip"
        '
        'lblTele1Caption
        '
        Me.lblTele1Caption.AutoSize = True
        Me.lblTele1Caption.Location = New System.Drawing.Point(42, 86)
        Me.lblTele1Caption.Name = "lblTele1Caption"
        Me.lblTele1Caption.Size = New System.Drawing.Size(37, 13)
        Me.lblTele1Caption.TabIndex = 30
        Me.lblTele1Caption.Text = "Tele1:"
        '
        'lblTele3
        '
        Me.lblTele3.AutoSize = True
        Me.lblTele3.Location = New System.Drawing.Point(77, 114)
        Me.lblTele3.Name = "lblTele3"
        Me.lblTele3.Size = New System.Drawing.Size(44, 13)
        Me.lblTele3.TabIndex = 35
        Me.lblTele3.Text = "lblTele3"
        '
        'lblTele2Caption
        '
        Me.lblTele2Caption.AutoSize = True
        Me.lblTele2Caption.Location = New System.Drawing.Point(42, 100)
        Me.lblTele2Caption.Name = "lblTele2Caption"
        Me.lblTele2Caption.Size = New System.Drawing.Size(37, 13)
        Me.lblTele2Caption.TabIndex = 31
        Me.lblTele2Caption.Text = "Tele2:"
        '
        'lblTele2
        '
        Me.lblTele2.AutoSize = True
        Me.lblTele2.Location = New System.Drawing.Point(77, 100)
        Me.lblTele2.Name = "lblTele2"
        Me.lblTele2.Size = New System.Drawing.Size(44, 13)
        Me.lblTele2.TabIndex = 34
        Me.lblTele2.Text = "lblTele2"
        '
        'lblTele3Caption
        '
        Me.lblTele3Caption.AutoSize = True
        Me.lblTele3Caption.Location = New System.Drawing.Point(42, 114)
        Me.lblTele3Caption.Name = "lblTele3Caption"
        Me.lblTele3Caption.Size = New System.Drawing.Size(37, 13)
        Me.lblTele3Caption.TabIndex = 32
        Me.lblTele3Caption.Text = "Tele3:"
        '
        'lblTele
        '
        Me.lblTele.AutoSize = True
        Me.lblTele.Location = New System.Drawing.Point(77, 86)
        Me.lblTele.Name = "lblTele"
        Me.lblTele.Size = New System.Drawing.Size(38, 13)
        Me.lblTele.TabIndex = 33
        Me.lblTele.Text = "lblTele"
        '
        'lblInvoiceNo
        '
        Me.lblInvoiceNo.AutoSize = True
        Me.lblInvoiceNo.Location = New System.Drawing.Point(4, 36)
        Me.lblInvoiceNo.Name = "lblInvoiceNo"
        Me.lblInvoiceNo.Size = New System.Drawing.Size(62, 13)
        Me.lblInvoiceNo.TabIndex = 3
        Me.lblInvoiceNo.Text = "Invoice No:"
        '
        'lblSaleNo
        '
        Me.lblSaleNo.AutoSize = True
        Me.lblSaleNo.Location = New System.Drawing.Point(18, 84)
        Me.lblSaleNo.Name = "lblSaleNo"
        Me.lblSaleNo.Size = New System.Drawing.Size(48, 13)
        Me.lblSaleNo.TabIndex = 4
        Me.lblSaleNo.Text = "Sale No:"
        '
        'lblServiceOrderNoCaption
        '
        Me.lblServiceOrderNoCaption.AutoSize = True
        Me.lblServiceOrderNoCaption.Location = New System.Drawing.Point(216, 12)
        Me.lblServiceOrderNoCaption.Name = "lblServiceOrderNoCaption"
        Me.lblServiceOrderNoCaption.Size = New System.Drawing.Size(92, 13)
        Me.lblServiceOrderNoCaption.TabIndex = 5
        Me.lblServiceOrderNoCaption.Text = "Service Order No:"
        '
        'lblServiceOrderNo
        '
        Me.lblServiceOrderNo.AutoSize = True
        Me.lblServiceOrderNo.Location = New System.Drawing.Point(308, 12)
        Me.lblServiceOrderNo.Name = "lblServiceOrderNo"
        Me.lblServiceOrderNo.Size = New System.Drawing.Size(92, 13)
        Me.lblServiceOrderNo.TabIndex = 6
        Me.lblServiceOrderNo.Text = "Service Order No:"
        '
        'lblWhatToDoWStyle
        '
        Me.lblWhatToDoWStyle.AutoSize = True
        Me.lblWhatToDoWStyle.Location = New System.Drawing.Point(57, 13)
        Me.lblWhatToDoWStyle.Name = "lblWhatToDoWStyle"
        Me.lblWhatToDoWStyle.Size = New System.Drawing.Size(133, 13)
        Me.lblWhatToDoWStyle.TabIndex = 7
        Me.lblWhatToDoWStyle.Text = "Click in Box to Select Style"
        '
        'txtInvoiceNo
        '
        Me.txtInvoiceNo.Location = New System.Drawing.Point(74, 35)
        Me.txtInvoiceNo.Name = "txtInvoiceNo"
        Me.txtInvoiceNo.Size = New System.Drawing.Size(121, 20)
        Me.txtInvoiceNo.TabIndex = 8
        '
        'txtSaleNo
        '
        Me.txtSaleNo.Location = New System.Drawing.Point(74, 82)
        Me.txtSaleNo.Name = "txtSaleNo"
        Me.txtSaleNo.Size = New System.Drawing.Size(100, 20)
        Me.txtSaleNo.TabIndex = 9
        '
        'txtStoreName
        '
        Me.txtStoreName.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtStoreName.Location = New System.Drawing.Point(8, 28)
        Me.txtStoreName.Name = "txtStoreName"
        Me.txtStoreName.Size = New System.Drawing.Size(246, 20)
        Me.txtStoreName.TabIndex = 10
        Me.ToolTip1.SetToolTip(Me.txtStoreName, "This box should contain the Store's Name.")
        '
        'dteClaimDateCaption1
        '
        Me.dteClaimDateCaption1.Location = New System.Drawing.Point(722, 325)
        Me.dteClaimDateCaption1.Name = "dteClaimDateCaption1"
        Me.dteClaimDateCaption1.Size = New System.Drawing.Size(200, 20)
        Me.dteClaimDateCaption1.TabIndex = 11
        '
        'dteClaimDate
        '
        Me.dteClaimDate.CustomFormat = "MM/dd/yyyy"
        Me.dteClaimDate.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dteClaimDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dteClaimDate.Location = New System.Drawing.Point(74, 59)
        Me.dteClaimDate.Name = "dteClaimDate"
        Me.dteClaimDate.Size = New System.Drawing.Size(121, 22)
        Me.dteClaimDate.TabIndex = 12
        '
        'cboStores
        '
        Me.cboStores.FormattingEnabled = True
        Me.cboStores.Location = New System.Drawing.Point(722, 377)
        Me.cboStores.Name = "cboStores"
        Me.cboStores.Size = New System.Drawing.Size(121, 21)
        Me.cboStores.TabIndex = 13
        '
        'txtStoreAddress
        '
        Me.txtStoreAddress.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtStoreAddress.Location = New System.Drawing.Point(8, 50)
        Me.txtStoreAddress.Name = "txtStoreAddress"
        Me.txtStoreAddress.Size = New System.Drawing.Size(246, 20)
        Me.txtStoreAddress.TabIndex = 15
        Me.ToolTip1.SetToolTip(Me.txtStoreAddress, "This box should contain the Store's Mailing Address.")
        '
        'txtStoreCity
        '
        Me.txtStoreCity.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtStoreCity.Location = New System.Drawing.Point(8, 73)
        Me.txtStoreCity.Name = "txtStoreCity"
        Me.txtStoreCity.Size = New System.Drawing.Size(246, 20)
        Me.txtStoreCity.TabIndex = 16
        Me.ToolTip1.SetToolTip(Me.txtStoreCity, "This box should contain the Store's City, State, and Zip-Code.")
        '
        'txtStorePhone
        '
        Me.txtStorePhone.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtStorePhone.Location = New System.Drawing.Point(8, 96)
        Me.txtStorePhone.Name = "txtStorePhone"
        Me.txtStorePhone.Size = New System.Drawing.Size(246, 20)
        Me.txtStorePhone.TabIndex = 17
        Me.ToolTip1.SetToolTip(Me.txtStorePhone, "This box should contain the Store's Telephone Number.")
        '
        'cmdMoveFirst
        '
        Me.cmdMoveFirst.Location = New System.Drawing.Point(4, 195)
        Me.cmdMoveFirst.Name = "cmdMoveFirst"
        Me.cmdMoveFirst.Size = New System.Drawing.Size(36, 32)
        Me.cmdMoveFirst.TabIndex = 18
        Me.cmdMoveFirst.Text = "Button1"
        Me.cmdMoveFirst.UseVisualStyleBackColor = True
        '
        'cmdMovePrevious
        '
        Me.cmdMovePrevious.Location = New System.Drawing.Point(46, 195)
        Me.cmdMovePrevious.Name = "cmdMovePrevious"
        Me.cmdMovePrevious.Size = New System.Drawing.Size(35, 32)
        Me.cmdMovePrevious.TabIndex = 19
        Me.cmdMovePrevious.Text = "Button1"
        Me.cmdMovePrevious.UseVisualStyleBackColor = True
        '
        'cmdMoveNext
        '
        Me.cmdMoveNext.Location = New System.Drawing.Point(87, 195)
        Me.cmdMoveNext.Name = "cmdMoveNext"
        Me.cmdMoveNext.Size = New System.Drawing.Size(34, 32)
        Me.cmdMoveNext.TabIndex = 20
        Me.cmdMoveNext.Text = "Button1"
        Me.cmdMoveNext.UseVisualStyleBackColor = True
        '
        'cmdMoveLast
        '
        Me.cmdMoveLast.Location = New System.Drawing.Point(127, 195)
        Me.cmdMoveLast.Name = "cmdMoveLast"
        Me.cmdMoveLast.Size = New System.Drawing.Size(47, 32)
        Me.cmdMoveLast.TabIndex = 21
        Me.cmdMoveLast.Text = "Button1"
        Me.cmdMoveLast.UseVisualStyleBackColor = True
        '
        'cmdMoveSearch
        '
        Me.cmdMoveSearch.Location = New System.Drawing.Point(182, 195)
        Me.cmdMoveSearch.Name = "cmdMoveSearch"
        Me.cmdMoveSearch.Size = New System.Drawing.Size(75, 35)
        Me.cmdMoveSearch.TabIndex = 22
        Me.cmdMoveSearch.Text = "Parts Order &Look-Up"
        Me.cmdMoveSearch.UseVisualStyleBackColor = True
        '
        'lblMoveRecords
        '
        Me.lblMoveRecords.AutoSize = True
        Me.lblMoveRecords.Location = New System.Drawing.Point(17, 179)
        Me.lblMoveRecords.Name = "lblMoveRecords"
        Me.lblMoveRecords.Size = New System.Drawing.Size(80, 13)
        Me.lblMoveRecords.TabIndex = 23
        Me.lblMoveRecords.Text = "Move Records:"
        '
        'txtStyleNo
        '
        Me.txtStyleNo.Location = New System.Drawing.Point(60, 26)
        Me.txtStyleNo.Name = "txtStyleNo"
        Me.txtStyleNo.Size = New System.Drawing.Size(130, 20)
        Me.txtStyleNo.TabIndex = 36
        '
        'txtDescription
        '
        Me.txtDescription.Location = New System.Drawing.Point(292, 26)
        Me.txtDescription.Name = "txtDescription"
        Me.txtDescription.Size = New System.Drawing.Size(401, 20)
        Me.txtDescription.TabIndex = 37
        '
        'lblMarginLine
        '
        Me.lblMarginLine.AutoSize = True
        Me.lblMarginLine.Location = New System.Drawing.Point(616, 179)
        Me.lblMarginLine.Name = "lblMarginLine"
        Me.lblMarginLine.Size = New System.Drawing.Size(65, 13)
        Me.lblMarginLine.TabIndex = 38
        Me.lblMarginLine.Text = "Margin Line:"
        '
        'lblPartsOrderNo
        '
        Me.lblPartsOrderNo.AutoSize = True
        Me.lblPartsOrderNo.Location = New System.Drawing.Point(308, 31)
        Me.lblPartsOrderNo.Name = "lblPartsOrderNo"
        Me.lblPartsOrderNo.Size = New System.Drawing.Size(81, 13)
        Me.lblPartsOrderNo.TabIndex = 39
        Me.lblPartsOrderNo.Text = "lblPartsOrderNo"
        '
        'lblClaimDate
        '
        Me.lblClaimDate.AutoSize = True
        Me.lblClaimDate.Location = New System.Drawing.Point(308, 50)
        Me.lblClaimDate.Name = "lblClaimDate"
        Me.lblClaimDate.Size = New System.Drawing.Size(65, 13)
        Me.lblClaimDate.TabIndex = 40
        Me.lblClaimDate.Text = "lblClaimDate"
        '
        'txtVendorAddress
        '
        Me.txtVendorAddress.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtVendorAddress.Location = New System.Drawing.Point(8, 46)
        Me.txtVendorAddress.Name = "txtVendorAddress"
        Me.txtVendorAddress.Size = New System.Drawing.Size(246, 20)
        Me.txtVendorAddress.TabIndex = 42
        Me.ToolTip1.SetToolTip(Me.txtVendorAddress, "This box should contain the Vendor's Mailing Address.")
        '
        'txtVendorCity
        '
        Me.txtVendorCity.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtVendorCity.Location = New System.Drawing.Point(8, 69)
        Me.txtVendorCity.Name = "txtVendorCity"
        Me.txtVendorCity.Size = New System.Drawing.Size(246, 20)
        Me.txtVendorCity.TabIndex = 43
        Me.ToolTip1.SetToolTip(Me.txtVendorCity, "This box should contain the Vendor's City, State, and Zip-Code.")
        '
        'txtVendorTele
        '
        Me.txtVendorTele.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtVendorTele.Location = New System.Drawing.Point(8, 92)
        Me.txtVendorTele.Name = "txtVendorTele"
        Me.txtVendorTele.Size = New System.Drawing.Size(246, 20)
        Me.txtVendorTele.TabIndex = 44
        Me.ToolTip1.SetToolTip(Me.txtVendorTele, "This box should contain the Vendor's Telephone Number.")
        '
        'txtRepairCost
        '
        Me.txtRepairCost.Location = New System.Drawing.Point(74, 105)
        Me.txtRepairCost.Name = "txtRepairCost"
        Me.txtRepairCost.Size = New System.Drawing.Size(100, 20)
        Me.txtRepairCost.TabIndex = 45
        '
        'chkPaid
        '
        Me.chkPaid.AutoSize = True
        Me.chkPaid.CheckAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.chkPaid.Location = New System.Drawing.Point(40, 132)
        Me.chkPaid.Name = "chkPaid"
        Me.chkPaid.Size = New System.Drawing.Size(47, 17)
        Me.chkPaid.TabIndex = 46
        Me.chkPaid.Text = "&Paid"
        Me.chkPaid.UseVisualStyleBackColor = True
        '
        'Notes_Text
        '
        Me.Notes_Text.Location = New System.Drawing.Point(9, 49)
        Me.Notes_Text.Multiline = True
        Me.Notes_Text.Name = "Notes_Text"
        Me.Notes_Text.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.Notes_Text.Size = New System.Drawing.Size(684, 127)
        Me.Notes_Text.TabIndex = 47
        '
        'optCBCredit
        '
        Me.optCBCredit.AutoSize = True
        Me.optCBCredit.Location = New System.Drawing.Point(221, 115)
        Me.optCBCredit.Name = "optCBCredit"
        Me.optCBCredit.Size = New System.Drawing.Size(95, 17)
        Me.optCBCredit.TabIndex = 48
        Me.optCBCredit.Text = "&Request Credit"
        Me.optCBCredit.UseVisualStyleBackColor = True
        Me.optCBCredit.Visible = False
        '
        'optCBDeduct
        '
        Me.optCBDeduct.AutoSize = True
        Me.optCBDeduct.Location = New System.Drawing.Point(221, 149)
        Me.optCBDeduct.Name = "optCBDeduct"
        Me.optCBDeduct.Size = New System.Drawing.Size(124, 17)
        Me.optCBDeduct.TabIndex = 49
        Me.optCBDeduct.Text = "&Deduct From Invoice"
        Me.optCBDeduct.UseVisualStyleBackColor = True
        '
        'optCBChargeBack
        '
        Me.optCBChargeBack.AutoSize = True
        Me.optCBChargeBack.Checked = True
        Me.optCBChargeBack.Location = New System.Drawing.Point(221, 132)
        Me.optCBChargeBack.Name = "optCBChargeBack"
        Me.optCBChargeBack.Size = New System.Drawing.Size(87, 17)
        Me.optCBChargeBack.TabIndex = 50
        Me.optCBChargeBack.TabStop = True
        Me.optCBChargeBack.Text = "&Charge Back"
        Me.optCBChargeBack.UseVisualStyleBackColor = True
        '
        'cboStatus
        '
        Me.cboStatus.FormattingEnabled = True
        Me.cboStatus.Location = New System.Drawing.Point(74, 12)
        Me.cboStatus.Name = "cboStatus"
        Me.cboStatus.Size = New System.Drawing.Size(121, 21)
        Me.cboStatus.TabIndex = 51
        '
        'txtVendorEmail
        '
        Me.txtVendorEmail.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtVendorEmail.Location = New System.Drawing.Point(8, 115)
        Me.txtVendorEmail.Name = "txtVendorEmail"
        Me.txtVendorEmail.Size = New System.Drawing.Size(246, 20)
        Me.txtVendorEmail.TabIndex = 52
        Me.ToolTip1.SetToolTip(Me.txtVendorEmail, "This box should contain the Vendor's Email Address.")
        '
        'fraSoldTo
        '
        Me.fraSoldTo.Controls.Add(Me.txtStoreName)
        Me.fraSoldTo.Controls.Add(Me.txtStoreAddress)
        Me.fraSoldTo.Controls.Add(Me.txtStoreCity)
        Me.fraSoldTo.Controls.Add(Me.txtStorePhone)
        Me.fraSoldTo.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.fraSoldTo.Location = New System.Drawing.Point(6, 7)
        Me.fraSoldTo.Name = "fraSoldTo"
        Me.fraSoldTo.Size = New System.Drawing.Size(265, 126)
        Me.fraSoldTo.TabIndex = 53
        Me.fraSoldTo.TabStop = False
        Me.fraSoldTo.Text = "Sold To:"
        '
        'fraVendor
        '
        Me.fraVendor.Controls.Add(Me.txtVendorName)
        Me.fraVendor.Controls.Add(Me.txtVendorAddress)
        Me.fraVendor.Controls.Add(Me.txtVendorEmail)
        Me.fraVendor.Controls.Add(Me.txtVendorCity)
        Me.fraVendor.Controls.Add(Me.txtVendorTele)
        Me.fraVendor.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.fraVendor.Location = New System.Drawing.Point(6, 138)
        Me.fraVendor.Name = "fraVendor"
        Me.fraVendor.Size = New System.Drawing.Size(265, 143)
        Me.fraVendor.TabIndex = 54
        Me.fraVendor.TabStop = False
        Me.fraVendor.Text = "Vendor:"
        '
        'txtVendorName
        '
        Me.txtVendorName.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtVendorName.FormattingEnabled = True
        Me.txtVendorName.Location = New System.Drawing.Point(8, 23)
        Me.txtVendorName.Name = "txtVendorName"
        Me.txtVendorName.Size = New System.Drawing.Size(246, 21)
        Me.txtVendorName.TabIndex = 52
        Me.ToolTip1.SetToolTip(Me.txtVendorName, "This box should contain the Vendor's Name.")
        '
        'fraStatus
        '
        Me.fraStatus.Controls.Add(Me.cmdPictures)
        Me.fraStatus.Controls.Add(Me.imgPicture)
        Me.fraStatus.Controls.Add(Me.lblClaimDateCaption)
        Me.fraStatus.Controls.Add(Me.lblPartsOrderNoCaption)
        Me.fraStatus.Controls.Add(Me.optCBDeduct)
        Me.fraStatus.Controls.Add(Me.optCBChargeBack)
        Me.fraStatus.Controls.Add(Me.lblRepairCost)
        Me.fraStatus.Controls.Add(Me.dteClaimDateCaption)
        Me.fraStatus.Controls.Add(Me.optCBCredit)
        Me.fraStatus.Controls.Add(Me.lblStatus)
        Me.fraStatus.Controls.Add(Me.cboStatus)
        Me.fraStatus.Controls.Add(Me.lblInvoiceNo)
        Me.fraStatus.Controls.Add(Me.txtInvoiceNo)
        Me.fraStatus.Controls.Add(Me.lblClaimDate)
        Me.fraStatus.Controls.Add(Me.dteClaimDate)
        Me.fraStatus.Controls.Add(Me.chkPaid)
        Me.fraStatus.Controls.Add(Me.lblSaleNo)
        Me.fraStatus.Controls.Add(Me.lblPartsOrderNo)
        Me.fraStatus.Controls.Add(Me.txtRepairCost)
        Me.fraStatus.Controls.Add(Me.txtSaleNo)
        Me.fraStatus.Controls.Add(Me.lblServiceOrderNo)
        Me.fraStatus.Controls.Add(Me.lblServiceOrderNoCaption)
        Me.fraStatus.Location = New System.Drawing.Point(277, 23)
        Me.fraStatus.Name = "fraStatus"
        Me.fraStatus.Size = New System.Drawing.Size(436, 172)
        Me.fraStatus.TabIndex = 55
        Me.fraStatus.TabStop = False
        '
        'cmdPictures
        '
        Me.cmdPictures.Location = New System.Drawing.Point(376, 128)
        Me.cmdPictures.Name = "cmdPictures"
        Me.cmdPictures.Size = New System.Drawing.Size(54, 38)
        Me.cmdPictures.TabIndex = 57
        Me.cmdPictures.Text = "Pictur&es"
        Me.cmdPictures.UseVisualStyleBackColor = True
        '
        'imgPicture
        '
        Me.imgPicture.Location = New System.Drawing.Point(311, 67)
        Me.imgPicture.Name = "imgPicture"
        Me.imgPicture.Size = New System.Drawing.Size(43, 35)
        Me.imgPicture.TabIndex = 56
        Me.imgPicture.TabStop = False
        Me.imgPicture.Visible = False
        '
        'lblClaimDateCaption
        '
        Me.lblClaimDateCaption.AutoSize = True
        Me.lblClaimDateCaption.Location = New System.Drawing.Point(235, 50)
        Me.lblClaimDateCaption.Name = "lblClaimDateCaption"
        Me.lblClaimDateCaption.Size = New System.Drawing.Size(73, 13)
        Me.lblClaimDateCaption.TabIndex = 55
        Me.lblClaimDateCaption.Text = "Date of Claim:"
        '
        'lblPartsOrderNoCaption
        '
        Me.lblPartsOrderNoCaption.AutoSize = True
        Me.lblPartsOrderNoCaption.Location = New System.Drawing.Point(228, 31)
        Me.lblPartsOrderNoCaption.Name = "lblPartsOrderNoCaption"
        Me.lblPartsOrderNoCaption.Size = New System.Drawing.Size(80, 13)
        Me.lblPartsOrderNoCaption.TabIndex = 54
        Me.lblPartsOrderNoCaption.Text = "Parts Order No:"
        '
        'lblRepairCost
        '
        Me.lblRepairCost.AutoSize = True
        Me.lblRepairCost.Location = New System.Drawing.Point(1, 108)
        Me.lblRepairCost.Name = "lblRepairCost"
        Me.lblRepairCost.Size = New System.Drawing.Size(65, 13)
        Me.lblRepairCost.TabIndex = 53
        Me.lblRepairCost.Text = "Repair Cost:"
        '
        'dteClaimDateCaption
        '
        Me.dteClaimDateCaption.AutoSize = True
        Me.dteClaimDateCaption.Location = New System.Drawing.Point(-5, 60)
        Me.dteClaimDateCaption.Name = "dteClaimDateCaption"
        Me.dteClaimDateCaption.Size = New System.Drawing.Size(71, 13)
        Me.dteClaimDateCaption.TabIndex = 52
        Me.dteClaimDateCaption.Text = "Invoice Date:"
        '
        'lblStatus
        '
        Me.lblStatus.AutoSize = True
        Me.lblStatus.Location = New System.Drawing.Point(26, 12)
        Me.lblStatus.Name = "lblStatus"
        Me.lblStatus.Size = New System.Drawing.Size(40, 13)
        Me.lblStatus.TabIndex = 0
        Me.lblStatus.Text = "Status:"
        '
        'fraNotes
        '
        Me.fraNotes.Controls.Add(Me.cmdPrintChargeBack)
        Me.fraNotes.Controls.Add(Me.cmdAddPart)
        Me.fraNotes.Controls.Add(Me.lblMarginLinelbl)
        Me.fraNotes.Controls.Add(Me.cmdMenu)
        Me.fraNotes.Controls.Add(Me.cmdNext)
        Me.fraNotes.Controls.Add(Me.cmdEmail)
        Me.fraNotes.Controls.Add(Me.cmdPrint)
        Me.fraNotes.Controls.Add(Me.lblMarginLine)
        Me.fraNotes.Controls.Add(Me.cmdSave)
        Me.fraNotes.Controls.Add(Me.lblDescription)
        Me.fraNotes.Controls.Add(Me.lblStyleNo)
        Me.fraNotes.Controls.Add(Me.lblWhatToDoWStyle)
        Me.fraNotes.Controls.Add(Me.txtStyleNo)
        Me.fraNotes.Controls.Add(Me.Notes_Text)
        Me.fraNotes.Controls.Add(Me.cmdMoveSearch)
        Me.fraNotes.Controls.Add(Me.lblMoveRecords)
        Me.fraNotes.Controls.Add(Me.cmdMoveLast)
        Me.fraNotes.Controls.Add(Me.txtDescription)
        Me.fraNotes.Controls.Add(Me.cmdMoveNext)
        Me.fraNotes.Controls.Add(Me.cmdMoveFirst)
        Me.fraNotes.Controls.Add(Me.cmdMovePrevious)
        Me.fraNotes.Location = New System.Drawing.Point(14, 332)
        Me.fraNotes.Name = "fraNotes"
        Me.fraNotes.Size = New System.Drawing.Size(699, 245)
        Me.fraNotes.TabIndex = 56
        Me.fraNotes.TabStop = False
        Me.fraNotes.Text = "Request:"
        '
        'cmdMenu
        '
        Me.cmdMenu.Location = New System.Drawing.Point(516, 192)
        Me.cmdMenu.Name = "cmdMenu"
        Me.cmdMenu.Size = New System.Drawing.Size(60, 41)
        Me.cmdMenu.TabIndex = 52
        Me.cmdMenu.Text = "&Menu"
        Me.cmdMenu.UseVisualStyleBackColor = True
        '
        'cmdNext
        '
        Me.cmdNext.Location = New System.Drawing.Point(453, 192)
        Me.cmdNext.Name = "cmdNext"
        Me.cmdNext.Size = New System.Drawing.Size(60, 41)
        Me.cmdNext.TabIndex = 51
        Me.cmdNext.Text = "&New"
        Me.cmdNext.UseVisualStyleBackColor = True
        '
        'cmdEmail
        '
        Me.cmdEmail.Location = New System.Drawing.Point(392, 192)
        Me.cmdEmail.Name = "cmdEmail"
        Me.cmdEmail.Size = New System.Drawing.Size(60, 41)
        Me.cmdEmail.TabIndex = 50
        Me.cmdEmail.Text = "&Email"
        Me.cmdEmail.UseVisualStyleBackColor = True
        '
        'cmdPrint
        '
        Me.cmdPrint.Location = New System.Drawing.Point(326, 192)
        Me.cmdPrint.Name = "cmdPrint"
        Me.cmdPrint.Size = New System.Drawing.Size(60, 41)
        Me.cmdPrint.TabIndex = 49
        Me.cmdPrint.Text = "&Print"
        Me.cmdPrint.UseVisualStyleBackColor = True
        '
        'cmdSave
        '
        Me.cmdSave.Location = New System.Drawing.Point(260, 192)
        Me.cmdSave.Name = "cmdSave"
        Me.cmdSave.Size = New System.Drawing.Size(60, 41)
        Me.cmdSave.TabIndex = 48
        Me.cmdSave.Text = "&Save"
        Me.cmdSave.UseVisualStyleBackColor = True
        '
        'lblDescription
        '
        Me.lblDescription.AutoSize = True
        Me.lblDescription.Location = New System.Drawing.Point(228, 26)
        Me.lblDescription.Name = "lblDescription"
        Me.lblDescription.Size = New System.Drawing.Size(63, 13)
        Me.lblDescription.TabIndex = 37
        Me.lblDescription.Text = "Description:"
        '
        'lblStyleNo
        '
        Me.lblStyleNo.AutoSize = True
        Me.lblStyleNo.Location = New System.Drawing.Point(6, 29)
        Me.lblStyleNo.Name = "lblStyleNo"
        Me.lblStyleNo.Size = New System.Drawing.Size(50, 13)
        Me.lblStyleNo.TabIndex = 8
        Me.lblStyleNo.Text = "Style No:"
        '
        'lblMarginLinelbl
        '
        Me.lblMarginLinelbl.AutoSize = True
        Me.lblMarginLinelbl.Location = New System.Drawing.Point(540, 179)
        Me.lblMarginLinelbl.Name = "lblMarginLinelbl"
        Me.lblMarginLinelbl.Size = New System.Drawing.Size(65, 13)
        Me.lblMarginLinelbl.TabIndex = 53
        Me.lblMarginLinelbl.Text = "Margin Line:"
        '
        'cmdAddPart
        '
        Me.cmdAddPart.Location = New System.Drawing.Point(619, 195)
        Me.cmdAddPart.Name = "cmdAddPart"
        Me.cmdAddPart.Size = New System.Drawing.Size(60, 18)
        Me.cmdAddPart.TabIndex = 54
        Me.cmdAddPart.Text = "&Add Part"
        Me.cmdAddPart.UseVisualStyleBackColor = True
        '
        'cmdPrintChargeBack
        '
        Me.cmdPrintChargeBack.Location = New System.Drawing.Point(582, 210)
        Me.cmdPrintChargeBack.Name = "cmdPrintChargeBack"
        Me.cmdPrintChargeBack.Size = New System.Drawing.Size(111, 23)
        Me.cmdPrintChargeBack.TabIndex = 55
        Me.cmdPrintChargeBack.Text = "Send Charge Back &Letter"
        Me.cmdPrintChargeBack.UseVisualStyleBackColor = True
        '
        'ServiceParts
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(800, 580)
        Me.Controls.Add(Me.fraNotes)
        Me.Controls.Add(Me.fraStatus)
        Me.Controls.Add(Me.fraVendor)
        Me.Controls.Add(Me.fraSoldTo)
        Me.Controls.Add(Me.cboStores)
        Me.Controls.Add(Me.dteClaimDateCaption1)
        Me.Controls.Add(Me.fraCustomer)
        Me.Controls.Add(Me.optTagCustomer)
        Me.Controls.Add(Me.optTagStock)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.MaximizeBox = False
        Me.Name = "ServiceParts"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Parts Order Form"
        Me.fraCustomer.ResumeLayout(False)
        Me.fraCustomer.PerformLayout()
        Me.fraSoldTo.ResumeLayout(False)
        Me.fraSoldTo.PerformLayout()
        Me.fraVendor.ResumeLayout(False)
        Me.fraVendor.PerformLayout()
        Me.fraStatus.ResumeLayout(False)
        Me.fraStatus.PerformLayout()
        CType(Me.imgPicture, System.ComponentModel.ISupportInitialize).EndInit()
        Me.fraNotes.ResumeLayout(False)
        Me.fraNotes.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents optTagStock As RadioButton
    Friend WithEvents optTagCustomer As RadioButton
    Friend WithEvents fraCustomer As GroupBox
    Friend WithEvents lblInvoiceNo As Label
    Friend WithEvents lblSaleNo As Label
    Friend WithEvents lblServiceOrderNoCaption As Label
    Friend WithEvents lblServiceOrderNo As Label
    Friend WithEvents lblWhatToDoWStyle As Label
    Friend WithEvents txtInvoiceNo As TextBox
    Friend WithEvents txtSaleNo As TextBox
    Friend WithEvents txtStoreName As TextBox
    Friend WithEvents dteClaimDateCaption1 As DateTimePicker
    Friend WithEvents dteClaimDate As DateTimePicker
    Friend WithEvents cboStores As ComboBox
    Friend WithEvents txtStoreAddress As TextBox
    Friend WithEvents txtStoreCity As TextBox
    Friend WithEvents txtStorePhone As TextBox
    Friend WithEvents cmdMoveFirst As Button
    Friend WithEvents cmdMovePrevious As Button
    Friend WithEvents cmdMoveNext As Button
    Friend WithEvents cmdMoveLast As Button
    Friend WithEvents cmdMoveSearch As Button
    Friend WithEvents lblMoveRecords As Label
    Friend WithEvents lblFirstName As Label
    Friend WithEvents lblLastName As Label
    Friend WithEvents lblAddress As Label
    Friend WithEvents lblAddress2 As Label
    Friend WithEvents lblCity As Label
    Friend WithEvents lblZip As Label
    Friend WithEvents lblTele1Caption As Label
    Friend WithEvents lblTele2Caption As Label
    Friend WithEvents lblTele3Caption As Label
    Friend WithEvents lblTele As Label
    Friend WithEvents lblTele2 As Label
    Friend WithEvents lblTele3 As Label
    Friend WithEvents txtStyleNo As TextBox
    Friend WithEvents txtDescription As TextBox
    Friend WithEvents lblMarginLine As Label
    Friend WithEvents lblPartsOrderNo As Label
    Friend WithEvents lblClaimDate As Label
    Friend WithEvents txtVendorAddress As TextBox
    Friend WithEvents txtVendorCity As TextBox
    Friend WithEvents txtVendorTele As TextBox
    Friend WithEvents txtRepairCost As TextBox
    Friend WithEvents chkPaid As CheckBox
    Friend WithEvents Notes_Text As TextBox
    Friend WithEvents optCBCredit As RadioButton
    Friend WithEvents optCBDeduct As RadioButton
    Friend WithEvents optCBChargeBack As RadioButton
    Friend WithEvents cboStatus As ComboBox
    Friend WithEvents txtVendorEmail As TextBox
    Friend WithEvents ToolTip1 As ToolTip
    Friend WithEvents fraSoldTo As GroupBox
    Friend WithEvents fraVendor As GroupBox
    Friend WithEvents txtVendorName As ComboBox
    Friend WithEvents fraStatus As GroupBox
    Friend WithEvents lblStatus As Label
    Friend WithEvents cmdPictures As Button
    Friend WithEvents imgPicture As PictureBox
    Friend WithEvents lblClaimDateCaption As Label
    Friend WithEvents lblPartsOrderNoCaption As Label
    Friend WithEvents lblRepairCost As Label
    Friend WithEvents dteClaimDateCaption As Label
    Friend WithEvents fraNotes As GroupBox
    Friend WithEvents cmdMenu As Button
    Friend WithEvents cmdNext As Button
    Friend WithEvents cmdEmail As Button
    Friend WithEvents cmdPrint As Button
    Friend WithEvents cmdSave As Button
    Friend WithEvents lblDescription As Label
    Friend WithEvents lblStyleNo As Label
    Friend WithEvents cmdPrintChargeBack As Button
    Friend WithEvents cmdAddPart As Button
    Friend WithEvents lblMarginLinelbl As Label
End Class
