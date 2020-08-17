<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Service
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
        Me.imgLogo = New System.Windows.Forms.PictureBox()
        Me.fraCustInfo = New System.Windows.Forms.GroupBox()
        Me.cmdAddItem = New System.Windows.Forms.Button()
        Me.cmdRepairTag = New System.Windows.Forms.Button()
        Me.cmdAddItemNote = New System.Windows.Forms.Button()
        Me.cmdTagForRepair = New System.Windows.Forms.Button()
        Me.chkOther = New System.Windows.Forms.CheckBox()
        Me.chkPickupExchange = New System.Windows.Forms.CheckBox()
        Me.chkOutsideService = New System.Windows.Forms.CheckBox()
        Me.chkStoreService = New System.Windows.Forms.CheckBox()
        Me.tvItemNotes = New System.Windows.Forms.TreeView()
        Me.lblTimeWindow = New System.Windows.Forms.Label()
        Me.dtpDelWindow2 = New System.Windows.Forms.DateTimePicker()
        Me.dtpDelWindow = New System.Windows.Forms.DateTimePicker()
        Me.lblSaleNo = New System.Windows.Forms.Label()
        Me.lblSaleNoCaption = New System.Windows.Forms.Label()
        Me.lblClaimDate = New System.Windows.Forms.Label()
        Me.lblCapClaimDate = New System.Windows.Forms.Label()
        Me.lblServiceOrderNo = New System.Windows.Forms.Label()
        Me.lblCapServiceOrderNo = New System.Windows.Forms.Label()
        Me.lblTele3 = New System.Windows.Forms.Label()
        Me.lblCapTele3 = New System.Windows.Forms.Label()
        Me.lblServiceOnDate = New System.Windows.Forms.Label()
        Me.dteServiceDate = New System.Windows.Forms.DateTimePicker()
        Me.cboStatus = New System.Windows.Forms.ComboBox()
        Me.lblStatus = New System.Windows.Forms.Label()
        Me.lblTele2 = New System.Windows.Forms.Label()
        Me.lblCapTele2 = New System.Windows.Forms.Label()
        Me.lblTele = New System.Windows.Forms.Label()
        Me.lblCapTele = New System.Windows.Forms.Label()
        Me.lblZip = New System.Windows.Forms.Label()
        Me.lblCity = New System.Windows.Forms.Label()
        Me.lblAddress2 = New System.Windows.Forms.Label()
        Me.lblAddress = New System.Windows.Forms.Label()
        Me.lblLastName = New System.Windows.Forms.Label()
        Me.lblFirstName = New System.Windows.Forms.Label()
        Me.txtItems = New System.Windows.Forms.TextBox()
        Me.lstPurchases = New System.Windows.Forms.CheckedListBox()
        Me.Notes_Frame = New System.Windows.Forms.GroupBox()
        Me.lblStoreResponse = New System.Windows.Forms.Label()
        Me.lblPartsOrd = New System.Windows.Forms.Label()
        Me.cmdOrderParts = New System.Windows.Forms.Button()
        Me.cmdMenu = New System.Windows.Forms.Button()
        Me.cmdNext = New System.Windows.Forms.Button()
        Me.cmdPrint = New System.Windows.Forms.Button()
        Me.cmdSave = New System.Windows.Forms.Button()
        Me.cmdMoveSearch = New System.Windows.Forms.Button()
        Me.lblMoveRecords = New System.Windows.Forms.Label()
        Me.cmdMoveLast = New System.Windows.Forms.Button()
        Me.cmdMoveNext = New System.Windows.Forms.Button()
        Me.cmdMovePrevious = New System.Windows.Forms.Button()
        Me.cmdMoveFirst = New System.Windows.Forms.Button()
        Me.Notes_New = New System.Windows.Forms.TextBox()
        Me.Notes_Text = New System.Windows.Forms.TextBox()
        Me.lblSpecial = New System.Windows.Forms.Label()
        CType(Me.imgLogo, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.fraCustInfo.SuspendLayout()
        Me.Notes_Frame.SuspendLayout()
        Me.SuspendLayout()
        '
        'imgLogo
        '
        Me.imgLogo.Location = New System.Drawing.Point(3, 3)
        Me.imgLogo.Name = "imgLogo"
        Me.imgLogo.Size = New System.Drawing.Size(348, 164)
        Me.imgLogo.TabIndex = 0
        Me.imgLogo.TabStop = False
        '
        'fraCustInfo
        '
        Me.fraCustInfo.Controls.Add(Me.lblSpecial)
        Me.fraCustInfo.Controls.Add(Me.cmdAddItem)
        Me.fraCustInfo.Controls.Add(Me.cmdRepairTag)
        Me.fraCustInfo.Controls.Add(Me.cmdAddItemNote)
        Me.fraCustInfo.Controls.Add(Me.cmdTagForRepair)
        Me.fraCustInfo.Controls.Add(Me.chkOther)
        Me.fraCustInfo.Controls.Add(Me.chkPickupExchange)
        Me.fraCustInfo.Controls.Add(Me.chkOutsideService)
        Me.fraCustInfo.Controls.Add(Me.chkStoreService)
        Me.fraCustInfo.Controls.Add(Me.tvItemNotes)
        Me.fraCustInfo.Controls.Add(Me.lblTimeWindow)
        Me.fraCustInfo.Controls.Add(Me.dtpDelWindow2)
        Me.fraCustInfo.Controls.Add(Me.dtpDelWindow)
        Me.fraCustInfo.Controls.Add(Me.lblSaleNo)
        Me.fraCustInfo.Controls.Add(Me.lblSaleNoCaption)
        Me.fraCustInfo.Controls.Add(Me.lblClaimDate)
        Me.fraCustInfo.Controls.Add(Me.lblCapClaimDate)
        Me.fraCustInfo.Controls.Add(Me.lblServiceOrderNo)
        Me.fraCustInfo.Controls.Add(Me.lblCapServiceOrderNo)
        Me.fraCustInfo.Controls.Add(Me.lblTele3)
        Me.fraCustInfo.Controls.Add(Me.lblCapTele3)
        Me.fraCustInfo.Controls.Add(Me.lblServiceOnDate)
        Me.fraCustInfo.Controls.Add(Me.dteServiceDate)
        Me.fraCustInfo.Controls.Add(Me.cboStatus)
        Me.fraCustInfo.Controls.Add(Me.lblStatus)
        Me.fraCustInfo.Controls.Add(Me.lblTele2)
        Me.fraCustInfo.Controls.Add(Me.lblCapTele2)
        Me.fraCustInfo.Controls.Add(Me.lblTele)
        Me.fraCustInfo.Controls.Add(Me.lblCapTele)
        Me.fraCustInfo.Controls.Add(Me.lblZip)
        Me.fraCustInfo.Controls.Add(Me.lblCity)
        Me.fraCustInfo.Controls.Add(Me.lblAddress2)
        Me.fraCustInfo.Controls.Add(Me.lblAddress)
        Me.fraCustInfo.Controls.Add(Me.lblLastName)
        Me.fraCustInfo.Controls.Add(Me.lblFirstName)
        Me.fraCustInfo.Controls.Add(Me.txtItems)
        Me.fraCustInfo.Controls.Add(Me.lstPurchases)
        Me.fraCustInfo.Location = New System.Drawing.Point(3, 4)
        Me.fraCustInfo.Name = "fraCustInfo"
        Me.fraCustInfo.Size = New System.Drawing.Size(680, 276)
        Me.fraCustInfo.TabIndex = 1
        Me.fraCustInfo.TabStop = False
        Me.fraCustInfo.Text = " Customer Information "
        '
        'cmdAddItem
        '
        Me.cmdAddItem.Location = New System.Drawing.Point(587, 118)
        Me.cmdAddItem.Name = "cmdAddItem"
        Me.cmdAddItem.Size = New System.Drawing.Size(87, 23)
        Me.cmdAddItem.TabIndex = 33
        Me.cmdAddItem.Text = "Ad&d Item"
        Me.cmdAddItem.UseVisualStyleBackColor = True
        '
        'cmdRepairTag
        '
        Me.cmdRepairTag.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.cmdRepairTag.Location = New System.Drawing.Point(599, 246)
        Me.cmdRepairTag.Name = "cmdRepairTag"
        Me.cmdRepairTag.Size = New System.Drawing.Size(75, 23)
        Me.cmdRepairTag.TabIndex = 32
        Me.cmdRepairTag.Text = "Repa&ir Tag"
        Me.cmdRepairTag.UseVisualStyleBackColor = False
        '
        'cmdAddItemNote
        '
        Me.cmdAddItemNote.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.cmdAddItemNote.Location = New System.Drawing.Point(489, 246)
        Me.cmdAddItemNote.Name = "cmdAddItemNote"
        Me.cmdAddItemNote.Size = New System.Drawing.Size(102, 23)
        Me.cmdAddItemNote.TabIndex = 31
        Me.cmdAddItemNote.Text = "Add Note to Item"
        Me.cmdAddItemNote.UseVisualStyleBackColor = False
        '
        'cmdTagForRepair
        '
        Me.cmdTagForRepair.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.cmdTagForRepair.Location = New System.Drawing.Point(389, 246)
        Me.cmdTagForRepair.Name = "cmdTagForRepair"
        Me.cmdTagForRepair.Size = New System.Drawing.Size(90, 23)
        Me.cmdTagForRepair.TabIndex = 30
        Me.cmdTagForRepair.Text = "Tag for Repair"
        Me.cmdTagForRepair.UseVisualStyleBackColor = False
        '
        'chkOther
        '
        Me.chkOther.AutoSize = True
        Me.chkOther.Location = New System.Drawing.Point(334, 252)
        Me.chkOther.Name = "chkOther"
        Me.chkOther.Size = New System.Drawing.Size(52, 17)
        Me.chkOther.TabIndex = 29
        Me.chkOther.Text = "Other"
        Me.chkOther.UseVisualStyleBackColor = True
        '
        'chkPickupExchange
        '
        Me.chkPickupExchange.AutoSize = True
        Me.chkPickupExchange.Location = New System.Drawing.Point(213, 252)
        Me.chkPickupExchange.Name = "chkPickupExchange"
        Me.chkPickupExchange.Size = New System.Drawing.Size(118, 17)
        Me.chkPickupExchange.TabIndex = 28
        Me.chkPickupExchange.Text = "&Pick Up  &Exchange"
        Me.chkPickupExchange.UseVisualStyleBackColor = True
        '
        'chkOutsideService
        '
        Me.chkOutsideService.AutoSize = True
        Me.chkOutsideService.Location = New System.Drawing.Point(109, 252)
        Me.chkOutsideService.Name = "chkOutsideService"
        Me.chkOutsideService.Size = New System.Drawing.Size(101, 17)
        Me.chkOutsideService.TabIndex = 27
        Me.chkOutsideService.Text = "&Outside Service"
        Me.chkOutsideService.UseVisualStyleBackColor = True
        '
        'chkStoreService
        '
        Me.chkStoreService.AutoSize = True
        Me.chkStoreService.Location = New System.Drawing.Point(16, 252)
        Me.chkStoreService.Name = "chkStoreService"
        Me.chkStoreService.Size = New System.Drawing.Size(90, 17)
        Me.chkStoreService.TabIndex = 26
        Me.chkStoreService.Text = "Sto&re Service"
        Me.chkStoreService.UseVisualStyleBackColor = True
        '
        'tvItemNotes
        '
        Me.tvItemNotes.Location = New System.Drawing.Point(10, 143)
        Me.tvItemNotes.Name = "tvItemNotes"
        Me.tvItemNotes.Size = New System.Drawing.Size(664, 97)
        Me.tvItemNotes.TabIndex = 25
        '
        'lblTimeWindow
        '
        Me.lblTimeWindow.AutoSize = True
        Me.lblTimeWindow.Location = New System.Drawing.Point(564, 86)
        Me.lblTimeWindow.Name = "lblTimeWindow"
        Me.lblTimeWindow.Size = New System.Drawing.Size(20, 13)
        Me.lblTimeWindow.TabIndex = 24
        Me.lblTimeWindow.Text = "To"
        '
        'dtpDelWindow2
        '
        Me.dtpDelWindow2.CustomFormat = "h:mm tt"
        Me.dtpDelWindow2.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpDelWindow2.Location = New System.Drawing.Point(584, 85)
        Me.dtpDelWindow2.Name = "dtpDelWindow2"
        Me.dtpDelWindow2.ShowCheckBox = True
        Me.dtpDelWindow2.Size = New System.Drawing.Size(90, 20)
        Me.dtpDelWindow2.TabIndex = 23
        '
        'dtpDelWindow
        '
        Me.dtpDelWindow.CustomFormat = "h:mm tt"
        Me.dtpDelWindow.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpDelWindow.Location = New System.Drawing.Point(477, 86)
        Me.dtpDelWindow.Name = "dtpDelWindow"
        Me.dtpDelWindow.ShowCheckBox = True
        Me.dtpDelWindow.Size = New System.Drawing.Size(88, 20)
        Me.dtpDelWindow.TabIndex = 22
        '
        'lblSaleNo
        '
        Me.lblSaleNo.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.lblSaleNo.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblSaleNo.Location = New System.Drawing.Point(584, 61)
        Me.lblSaleNo.Name = "lblSaleNo"
        Me.lblSaleNo.Size = New System.Drawing.Size(90, 23)
        Me.lblSaleNo.TabIndex = 21
        '
        'lblSaleNoCaption
        '
        Me.lblSaleNoCaption.AutoSize = True
        Me.lblSaleNoCaption.Location = New System.Drawing.Point(494, 62)
        Me.lblSaleNoCaption.Name = "lblSaleNoCaption"
        Me.lblSaleNoCaption.Size = New System.Drawing.Size(71, 13)
        Me.lblSaleNoCaption.TabIndex = 20
        Me.lblSaleNoCaption.Text = "Sale Number:"
        '
        'lblClaimDate
        '
        Me.lblClaimDate.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.lblClaimDate.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblClaimDate.Location = New System.Drawing.Point(584, 37)
        Me.lblClaimDate.Name = "lblClaimDate"
        Me.lblClaimDate.Size = New System.Drawing.Size(90, 23)
        Me.lblClaimDate.TabIndex = 19
        '
        'lblCapClaimDate
        '
        Me.lblCapClaimDate.AutoSize = True
        Me.lblCapClaimDate.Location = New System.Drawing.Point(494, 40)
        Me.lblCapClaimDate.Name = "lblCapClaimDate"
        Me.lblCapClaimDate.Size = New System.Drawing.Size(73, 13)
        Me.lblCapClaimDate.TabIndex = 18
        Me.lblCapClaimDate.Text = "Date of Claim:"
        '
        'lblServiceOrderNo
        '
        Me.lblServiceOrderNo.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.lblServiceOrderNo.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblServiceOrderNo.Location = New System.Drawing.Point(584, 11)
        Me.lblServiceOrderNo.Name = "lblServiceOrderNo"
        Me.lblServiceOrderNo.Size = New System.Drawing.Size(90, 25)
        Me.lblServiceOrderNo.TabIndex = 17
        '
        'lblCapServiceOrderNo
        '
        Me.lblCapServiceOrderNo.AutoSize = True
        Me.lblCapServiceOrderNo.Location = New System.Drawing.Point(494, 15)
        Me.lblCapServiceOrderNo.Name = "lblCapServiceOrderNo"
        Me.lblCapServiceOrderNo.Size = New System.Drawing.Size(92, 13)
        Me.lblCapServiceOrderNo.TabIndex = 16
        Me.lblCapServiceOrderNo.Text = "Service Order No:"
        '
        'lblTele3
        '
        Me.lblTele3.AutoSize = True
        Me.lblTele3.Location = New System.Drawing.Point(369, 113)
        Me.lblTele3.Name = "lblTele3"
        Me.lblTele3.Size = New System.Drawing.Size(44, 13)
        Me.lblTele3.TabIndex = 15
        Me.lblTele3.Text = "lblTele3"
        '
        'lblCapTele3
        '
        Me.lblCapTele3.AutoSize = True
        Me.lblCapTele3.Location = New System.Drawing.Point(326, 113)
        Me.lblCapTele3.Name = "lblCapTele3"
        Me.lblCapTele3.Size = New System.Drawing.Size(37, 13)
        Me.lblCapTele3.TabIndex = 14
        Me.lblCapTele3.Text = "Tele3:"
        '
        'lblServiceOnDate
        '
        Me.lblServiceOnDate.AutoSize = True
        Me.lblServiceOnDate.Location = New System.Drawing.Point(345, 68)
        Me.lblServiceOnDate.Name = "lblServiceOnDate"
        Me.lblServiceOnDate.Size = New System.Drawing.Size(87, 13)
        Me.lblServiceOnDate.TabIndex = 13
        Me.lblServiceOnDate.Text = "Service on Date:"
        '
        'dteServiceDate
        '
        Me.dteServiceDate.CustomFormat = "MM/dd/yyyy"
        Me.dteServiceDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dteServiceDate.Location = New System.Drawing.Point(345, 86)
        Me.dteServiceDate.Name = "dteServiceDate"
        Me.dteServiceDate.ShowCheckBox = True
        Me.dteServiceDate.Size = New System.Drawing.Size(117, 20)
        Me.dteServiceDate.TabIndex = 12
        '
        'cboStatus
        '
        Me.cboStatus.FormattingEnabled = True
        Me.cboStatus.Location = New System.Drawing.Point(345, 37)
        Me.cboStatus.Name = "cboStatus"
        Me.cboStatus.Size = New System.Drawing.Size(121, 21)
        Me.cboStatus.TabIndex = 11
        Me.cboStatus.Text = "cboStatus"
        '
        'lblStatus
        '
        Me.lblStatus.AutoSize = True
        Me.lblStatus.Location = New System.Drawing.Point(345, 21)
        Me.lblStatus.Name = "lblStatus"
        Me.lblStatus.Size = New System.Drawing.Size(40, 13)
        Me.lblStatus.TabIndex = 10
        Me.lblStatus.Text = "Status:"
        '
        'lblTele2
        '
        Me.lblTele2.AutoSize = True
        Me.lblTele2.Location = New System.Drawing.Point(211, 113)
        Me.lblTele2.Name = "lblTele2"
        Me.lblTele2.Size = New System.Drawing.Size(44, 13)
        Me.lblTele2.TabIndex = 9
        Me.lblTele2.Text = "lblTele2"
        '
        'lblCapTele2
        '
        Me.lblCapTele2.AutoSize = True
        Me.lblCapTele2.Location = New System.Drawing.Point(148, 113)
        Me.lblCapTele2.Name = "lblCapTele2"
        Me.lblCapTele2.Size = New System.Drawing.Size(37, 13)
        Me.lblCapTele2.TabIndex = 8
        Me.lblCapTele2.Text = "Tele2:"
        '
        'lblTele
        '
        Me.lblTele.AutoSize = True
        Me.lblTele.Location = New System.Drawing.Point(59, 113)
        Me.lblTele.Name = "lblTele"
        Me.lblTele.Size = New System.Drawing.Size(38, 13)
        Me.lblTele.TabIndex = 7
        Me.lblTele.Text = "lblTele"
        '
        'lblCapTele
        '
        Me.lblCapTele.AutoSize = True
        Me.lblCapTele.Location = New System.Drawing.Point(16, 113)
        Me.lblCapTele.Name = "lblCapTele"
        Me.lblCapTele.Size = New System.Drawing.Size(37, 13)
        Me.lblCapTele.TabIndex = 6
        Me.lblCapTele.Text = "Tele1:"
        '
        'lblZip
        '
        Me.lblZip.Location = New System.Drawing.Point(189, 85)
        Me.lblZip.Name = "lblZip"
        Me.lblZip.Size = New System.Drawing.Size(115, 23)
        Me.lblZip.TabIndex = 5
        Me.lblZip.Text = "lblZip"
        '
        'lblCity
        '
        Me.lblCity.Location = New System.Drawing.Point(13, 85)
        Me.lblCity.Name = "lblCity"
        Me.lblCity.Size = New System.Drawing.Size(131, 23)
        Me.lblCity.TabIndex = 4
        Me.lblCity.Text = "lblCity"
        '
        'lblAddress2
        '
        Me.lblAddress2.Location = New System.Drawing.Point(13, 63)
        Me.lblAddress2.Name = "lblAddress2"
        Me.lblAddress2.Size = New System.Drawing.Size(297, 23)
        Me.lblAddress2.TabIndex = 3
        Me.lblAddress2.Text = "lblAddress2"
        '
        'lblAddress
        '
        Me.lblAddress.Location = New System.Drawing.Point(13, 38)
        Me.lblAddress.Name = "lblAddress"
        Me.lblAddress.Size = New System.Drawing.Size(318, 23)
        Me.lblAddress.TabIndex = 2
        Me.lblAddress.Text = "lblAddress"
        '
        'lblLastName
        '
        Me.lblLastName.Location = New System.Drawing.Point(161, 21)
        Me.lblLastName.Name = "lblLastName"
        Me.lblLastName.Size = New System.Drawing.Size(152, 23)
        Me.lblLastName.TabIndex = 1
        Me.lblLastName.Text = "lblLastName"
        '
        'lblFirstName
        '
        Me.lblFirstName.Location = New System.Drawing.Point(13, 21)
        Me.lblFirstName.Name = "lblFirstName"
        Me.lblFirstName.Size = New System.Drawing.Size(120, 23)
        Me.lblFirstName.TabIndex = 0
        Me.lblFirstName.Text = "lblFirstName"
        '
        'txtItems
        '
        Me.txtItems.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.txtItems.Location = New System.Drawing.Point(16, 147)
        Me.txtItems.Multiline = True
        Me.txtItems.Name = "txtItems"
        Me.txtItems.ScrollBars = System.Windows.Forms.ScrollBars.Both
        Me.txtItems.Size = New System.Drawing.Size(575, 55)
        Me.txtItems.TabIndex = 34
        '
        'lstPurchases
        '
        Me.lstPurchases.FormattingEnabled = True
        Me.lstPurchases.Location = New System.Drawing.Point(19, 166)
        Me.lstPurchases.Name = "lstPurchases"
        Me.lstPurchases.Size = New System.Drawing.Size(562, 19)
        Me.lstPurchases.TabIndex = 35
        '
        'Notes_Frame
        '
        Me.Notes_Frame.Controls.Add(Me.lblStoreResponse)
        Me.Notes_Frame.Controls.Add(Me.lblPartsOrd)
        Me.Notes_Frame.Controls.Add(Me.cmdOrderParts)
        Me.Notes_Frame.Controls.Add(Me.cmdMenu)
        Me.Notes_Frame.Controls.Add(Me.cmdNext)
        Me.Notes_Frame.Controls.Add(Me.cmdPrint)
        Me.Notes_Frame.Controls.Add(Me.cmdSave)
        Me.Notes_Frame.Controls.Add(Me.cmdMoveSearch)
        Me.Notes_Frame.Controls.Add(Me.lblMoveRecords)
        Me.Notes_Frame.Controls.Add(Me.cmdMoveLast)
        Me.Notes_Frame.Controls.Add(Me.cmdMoveNext)
        Me.Notes_Frame.Controls.Add(Me.cmdMovePrevious)
        Me.Notes_Frame.Controls.Add(Me.cmdMoveFirst)
        Me.Notes_Frame.Controls.Add(Me.Notes_New)
        Me.Notes_Frame.Controls.Add(Me.Notes_Text)
        Me.Notes_Frame.Location = New System.Drawing.Point(3, 286)
        Me.Notes_Frame.Name = "Notes_Frame"
        Me.Notes_Frame.Size = New System.Drawing.Size(680, 279)
        Me.Notes_Frame.TabIndex = 2
        Me.Notes_Frame.TabStop = False
        Me.Notes_Frame.Text = "Customer Comp&laint:"
        '
        'lblStoreResponse
        '
        Me.lblStoreResponse.AutoSize = True
        Me.lblStoreResponse.Location = New System.Drawing.Point(36, 115)
        Me.lblStoreResponse.Name = "lblStoreResponse"
        Me.lblStoreResponse.Size = New System.Drawing.Size(95, 13)
        Me.lblStoreResponse.TabIndex = 14
        Me.lblStoreResponse.Text = "Store Res&ponse:   "
        '
        'lblPartsOrd
        '
        Me.lblPartsOrd.AutoSize = True
        Me.lblPartsOrd.BackColor = System.Drawing.Color.Yellow
        Me.lblPartsOrd.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblPartsOrd.Location = New System.Drawing.Point(586, 259)
        Me.lblPartsOrd.Name = "lblPartsOrd"
        Me.lblPartsOrd.Size = New System.Drawing.Size(90, 15)
        Me.lblPartsOrd.TabIndex = 13
        Me.lblPartsOrd.Text = "N Parts On Order"
        '
        'cmdOrderParts
        '
        Me.cmdOrderParts.Location = New System.Drawing.Point(599, 232)
        Me.cmdOrderParts.Name = "cmdOrderParts"
        Me.cmdOrderParts.Size = New System.Drawing.Size(75, 23)
        Me.cmdOrderParts.TabIndex = 12
        Me.cmdOrderParts.Text = "&Order Parts"
        Me.cmdOrderParts.UseVisualStyleBackColor = True
        '
        'cmdMenu
        '
        Me.cmdMenu.Location = New System.Drawing.Point(424, 229)
        Me.cmdMenu.Name = "cmdMenu"
        Me.cmdMenu.Size = New System.Drawing.Size(47, 29)
        Me.cmdMenu.TabIndex = 11
        Me.cmdMenu.Text = "&Menu"
        Me.cmdMenu.UseVisualStyleBackColor = True
        '
        'cmdNext
        '
        Me.cmdNext.Location = New System.Drawing.Point(384, 229)
        Me.cmdNext.Name = "cmdNext"
        Me.cmdNext.Size = New System.Drawing.Size(40, 29)
        Me.cmdNext.TabIndex = 10
        Me.cmdNext.Text = "&Next"
        Me.cmdNext.UseVisualStyleBackColor = True
        '
        'cmdPrint
        '
        Me.cmdPrint.Location = New System.Drawing.Point(344, 229)
        Me.cmdPrint.Name = "cmdPrint"
        Me.cmdPrint.Size = New System.Drawing.Size(40, 29)
        Me.cmdPrint.TabIndex = 9
        Me.cmdPrint.Text = "&Print"
        Me.cmdPrint.UseVisualStyleBackColor = True
        '
        'cmdSave
        '
        Me.cmdSave.Location = New System.Drawing.Point(304, 229)
        Me.cmdSave.Name = "cmdSave"
        Me.cmdSave.Size = New System.Drawing.Size(40, 29)
        Me.cmdSave.TabIndex = 8
        Me.cmdSave.Text = "&Save"
        Me.cmdSave.UseVisualStyleBackColor = True
        '
        'cmdMoveSearch
        '
        Me.cmdMoveSearch.Location = New System.Drawing.Point(178, 246)
        Me.cmdMoveSearch.Name = "cmdMoveSearch"
        Me.cmdMoveSearch.Size = New System.Drawing.Size(69, 23)
        Me.cmdMoveSearch.TabIndex = 7
        Me.cmdMoveSearch.Text = "&Look-Up"
        Me.cmdMoveSearch.UseVisualStyleBackColor = True
        '
        'lblMoveRecords
        '
        Me.lblMoveRecords.AutoSize = True
        Me.lblMoveRecords.Location = New System.Drawing.Point(6, 229)
        Me.lblMoveRecords.Name = "lblMoveRecords"
        Me.lblMoveRecords.Size = New System.Drawing.Size(80, 13)
        Me.lblMoveRecords.TabIndex = 6
        Me.lblMoveRecords.Text = "Move Records:"
        '
        'cmdMoveLast
        '
        Me.cmdMoveLast.Location = New System.Drawing.Point(123, 242)
        Me.cmdMoveLast.Name = "cmdMoveLast"
        Me.cmdMoveLast.Size = New System.Drawing.Size(38, 30)
        Me.cmdMoveLast.TabIndex = 5
        Me.cmdMoveLast.UseVisualStyleBackColor = True
        '
        'cmdMoveNext
        '
        Me.cmdMoveNext.Location = New System.Drawing.Point(85, 242)
        Me.cmdMoveNext.Name = "cmdMoveNext"
        Me.cmdMoveNext.Size = New System.Drawing.Size(38, 30)
        Me.cmdMoveNext.TabIndex = 4
        Me.cmdMoveNext.UseVisualStyleBackColor = True
        '
        'cmdMovePrevious
        '
        Me.cmdMovePrevious.Location = New System.Drawing.Point(47, 242)
        Me.cmdMovePrevious.Name = "cmdMovePrevious"
        Me.cmdMovePrevious.Size = New System.Drawing.Size(38, 30)
        Me.cmdMovePrevious.TabIndex = 3
        Me.cmdMovePrevious.UseVisualStyleBackColor = True
        '
        'cmdMoveFirst
        '
        Me.cmdMoveFirst.Location = New System.Drawing.Point(9, 242)
        Me.cmdMoveFirst.Name = "cmdMoveFirst"
        Me.cmdMoveFirst.Size = New System.Drawing.Size(38, 30)
        Me.cmdMoveFirst.TabIndex = 2
        Me.cmdMoveFirst.UseVisualStyleBackColor = True
        '
        'Notes_New
        '
        Me.Notes_New.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.Notes_New.Location = New System.Drawing.Point(10, 131)
        Me.Notes_New.Multiline = True
        Me.Notes_New.Name = "Notes_New"
        Me.Notes_New.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.Notes_New.Size = New System.Drawing.Size(665, 94)
        Me.Notes_New.TabIndex = 1
        '
        'Notes_Text
        '
        Me.Notes_Text.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.Notes_Text.Location = New System.Drawing.Point(10, 19)
        Me.Notes_Text.Multiline = True
        Me.Notes_Text.Name = "Notes_Text"
        Me.Notes_Text.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.Notes_Text.Size = New System.Drawing.Size(665, 93)
        Me.Notes_Text.TabIndex = 0
        '
        'lblSpecial
        '
        Me.lblSpecial.Location = New System.Drawing.Point(16, 132)
        Me.lblSpecial.Name = "lblSpecial"
        Me.lblSpecial.Size = New System.Drawing.Size(378, 8)
        Me.lblSpecial.TabIndex = 36
        Me.lblSpecial.Text = "lblSpecial"
        '
        'Service
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(689, 571)
        Me.Controls.Add(Me.Notes_Frame)
        Me.Controls.Add(Me.fraCustInfo)
        Me.Controls.Add(Me.imgLogo)
        Me.Name = "Service"
        Me.Text = "Service Module Intake"
        CType(Me.imgLogo, System.ComponentModel.ISupportInitialize).EndInit()
        Me.fraCustInfo.ResumeLayout(False)
        Me.fraCustInfo.PerformLayout()
        Me.Notes_Frame.ResumeLayout(False)
        Me.Notes_Frame.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents imgLogo As PictureBox
    Friend WithEvents fraCustInfo As GroupBox
    Friend WithEvents lblZip As Label
    Friend WithEvents lblCity As Label
    Friend WithEvents lblAddress2 As Label
    Friend WithEvents lblAddress As Label
    Friend WithEvents lblLastName As Label
    Friend WithEvents lblFirstName As Label
    Friend WithEvents lblSaleNo As Label
    Friend WithEvents lblSaleNoCaption As Label
    Friend WithEvents lblClaimDate As Label
    Friend WithEvents lblCapClaimDate As Label
    Friend WithEvents lblServiceOrderNo As Label
    Friend WithEvents lblCapServiceOrderNo As Label
    Friend WithEvents lblTele3 As Label
    Friend WithEvents lblCapTele3 As Label
    Friend WithEvents lblServiceOnDate As Label
    Friend WithEvents dteServiceDate As DateTimePicker
    Friend WithEvents cboStatus As ComboBox
    Friend WithEvents lblStatus As Label
    Friend WithEvents lblTele2 As Label
    Friend WithEvents lblCapTele2 As Label
    Friend WithEvents lblTele As Label
    Friend WithEvents lblCapTele As Label
    Friend WithEvents dtpDelWindow2 As DateTimePicker
    Friend WithEvents dtpDelWindow As DateTimePicker
    Friend WithEvents lblTimeWindow As Label
    Friend WithEvents cmdRepairTag As Button
    Friend WithEvents cmdAddItemNote As Button
    Friend WithEvents cmdTagForRepair As Button
    Friend WithEvents chkOther As CheckBox
    Friend WithEvents chkPickupExchange As CheckBox
    Friend WithEvents chkOutsideService As CheckBox
    Friend WithEvents chkStoreService As CheckBox
    Friend WithEvents tvItemNotes As TreeView
    Friend WithEvents cmdAddItem As Button
    Friend WithEvents Notes_Frame As GroupBox
    Friend WithEvents Notes_New As TextBox
    Friend WithEvents Notes_Text As TextBox
    Friend WithEvents lblMoveRecords As Label
    Friend WithEvents cmdMoveLast As Button
    Friend WithEvents cmdMoveNext As Button
    Friend WithEvents cmdMovePrevious As Button
    Friend WithEvents cmdMoveFirst As Button
    Friend WithEvents cmdMenu As Button
    Friend WithEvents cmdNext As Button
    Friend WithEvents cmdPrint As Button
    Friend WithEvents cmdSave As Button
    Friend WithEvents cmdMoveSearch As Button
    Friend WithEvents lblPartsOrd As Label
    Friend WithEvents cmdOrderParts As Button
    Friend WithEvents lblStoreResponse As Label
    Friend WithEvents txtItems As TextBox
    Friend WithEvents lstPurchases As CheckedListBox
    Friend WithEvents lblSpecial As Label
End Class
