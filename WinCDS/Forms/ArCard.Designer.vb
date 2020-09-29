<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class ArCard
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
        Me.lblTotalPayoff = New System.Windows.Forms.Label()
        Me.cmdApply = New System.Windows.Forms.Button()
        Me.cmdEdit = New System.Windows.Forms.Button()
        Me.cmdExport = New System.Windows.Forms.Button()
        Me.cmdFields = New System.Windows.Forms.Button()
        Me.cmdMakeSameAsCash = New System.Windows.Forms.Button()
        Me.cmdMoveFirst = New System.Windows.Forms.Button()
        Me.cmdMoveLast = New System.Windows.Forms.Button()
        Me.cmdMoveNext = New System.Windows.Forms.Button()
        Me.cmdMovePrevious = New System.Windows.Forms.Button()
        Me.cmdPayoff = New System.Windows.Forms.Button()
        Me.cmdReceipt = New System.Windows.Forms.Button()
        Me.cmdReprintContract = New System.Windows.Forms.Button()
        Me.cmdPrint = New System.Windows.Forms.Button()
        Me.cmdCancel = New System.Windows.Forms.Button()
        Me.chkSendAllMail = New System.Windows.Forms.CheckBox()
        Me.txtPaymentHistory = New System.Windows.Forms.TextBox()
        Me.lblBalance = New System.Windows.Forms.Label()
        Me.txtDelivery = New System.Windows.Forms.TextBox()
        Me.txtMonths = New System.Windows.Forms.TextBox()
        Me.txtLastPay = New System.Windows.Forms.TextBox()
        Me.txtFinanced = New System.Windows.Forms.TextBox()
        Me.txtMonthlyPayment = New System.Windows.Forms.TextBox()
        Me.txtPayMemo = New System.Windows.Forms.TextBox()
        Me.txtRate = New System.Windows.Forms.TextBox()
        Me.txtSameAsCash = New System.Windows.Forms.TextBox()
        Me.lblLateCharge = New System.Windows.Forms.Label()
        Me.txtFirstPay = New System.Windows.Forms.TextBox()
        Me.txtPaidBy = New System.Windows.Forms.TextBox()
        Me.txtPayPeriod = New System.Windows.Forms.TextBox()
        Me.txtNextDue = New System.Windows.Forms.TextBox()
        Me.lblTotDue = New System.Windows.Forms.Label()
        Me.lblArrearages = New System.Windows.Forms.Label()
        Me.lblLate0 = New System.Windows.Forms.Label()
        Me.lblLate31 = New System.Windows.Forms.Label()
        Me.lblLate61 = New System.Windows.Forms.Label()
        Me.lblLate91 = New System.Windows.Forms.Label()
        Me.lbl0030 = New System.Windows.Forms.Label()
        Me.lbl3160 = New System.Windows.Forms.Label()
        Me.lbl6190 = New System.Windows.Forms.Label()
        Me.lblOver91 = New System.Windows.Forms.Label()
        Me.filFile = New Microsoft.VisualBasic.Compatibility.VB6.FileListBox()
        Me.lblAccount = New System.Windows.Forms.Label()
        Me.cboStatus = New System.Windows.Forms.ComboBox()
        Me.lblFirstName = New System.Windows.Forms.Label()
        Me.lblLastName = New System.Windows.Forms.Label()
        Me.lblAddress = New System.Windows.Forms.Label()
        Me.lblCity = New System.Windows.Forms.Label()
        Me.lblZip = New System.Windows.Forms.Label()
        Me.lblTele1 = New System.Windows.Forms.Label()
        Me.lblTele2 = New System.Windows.Forms.Label()
        Me.lblSSN = New System.Windows.Forms.Label()
        Me.txtLateChargeAmount = New System.Windows.Forms.TextBox()
        Me.lblCreditLimit = New System.Windows.Forms.Label()
        Me.lblApprovalTerms = New System.Windows.Forms.Label()
        Me.lblAddAddress = New System.Windows.Forms.Label()
        Me.lblTele3 = New System.Windows.Forms.Label()
        Me.lblTele1Caption = New System.Windows.Forms.Label()
        Me.lblTele2Caption = New System.Windows.Forms.Label()
        Me.lblTele3Caption = New System.Windows.Forms.Label()
        Me.fraPrint = New System.Windows.Forms.GroupBox()
        Me.rtfFile = New System.Windows.Forms.RichTextBox()
        Me.fraNav = New System.Windows.Forms.GroupBox()
        Me.cmdPrintLabel = New System.Windows.Forms.Button()
        Me.fraTerms = New System.Windows.Forms.GroupBox()
        Me.cmdReprintCoupons = New System.Windows.Forms.Button()
        Me.lblNextDue = New System.Windows.Forms.Label()
        Me.lblLastPay = New System.Windows.Forms.Label()
        Me.lblSameAsCash = New System.Windows.Forms.Label()
        Me.lbl1stPay = New System.Windows.Forms.Label()
        Me.lblDelivery = New System.Windows.Forms.Label()
        Me.lblPayBy = New System.Windows.Forms.Label()
        Me.lblLateCh = New System.Windows.Forms.Label()
        Me.lblPayment = New System.Windows.Forms.Label()
        Me.lblRate = New System.Windows.Forms.Label()
        Me.lblMonths = New System.Windows.Forms.Label()
        Me.lblPaymentHistory = New System.Windows.Forms.Label()
        Me.lblAPR = New System.Windows.Forms.Label()
        Me.lblFinanced = New System.Windows.Forms.Label()
        Me.fraBalance = New System.Windows.Forms.GroupBox()
        Me.txtPayment = New System.Windows.Forms.TextBox()
        Me.lblPayDate = New System.Windows.Forms.Label()
        Me.txtLateCharge = New System.Windows.Forms.TextBox()
        Me.chkPayLateFee = New System.Windows.Forms.CheckBox()
        Me.lblTotalPayOffCaption = New System.Windows.Forms.Label()
        Me.lblLateChargeCaption = New System.Windows.Forms.Label()
        Me.lblBalanceCaption = New System.Windows.Forms.Label()
        Me.DDate = New System.Windows.Forms.DateTimePicker()
        Me.lblPayMemo = New System.Windows.Forms.Label()
        Me.cmdCreditApp = New System.Windows.Forms.Button()
        Me.cmdDetail = New System.Windows.Forms.Button()
        Me.Notes_Open = New System.Windows.Forms.Button()
        Me.cmdPrintCard = New System.Windows.Forms.Button()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.cmdSaleTotals = New System.Windows.Forms.Button()
        Me.optPayType9 = New System.Windows.Forms.RadioButton()
        Me.Notes_Save = New System.Windows.Forms.Button()
        Me.cmdNotesCancel = New System.Windows.Forms.Button()
        Me.cmdNotesPrint = New System.Windows.Forms.Button()
        Me.cmdHistory = New System.Windows.Forms.Button()
        Me.fraEditOptions = New System.Windows.Forms.GroupBox()
        Me.optEditType18 = New System.Windows.Forms.RadioButton()
        Me.optEditType17 = New System.Windows.Forms.RadioButton()
        Me.optEditType16 = New System.Windows.Forms.RadioButton()
        Me.optEditType15 = New System.Windows.Forms.RadioButton()
        Me.optEditType14 = New System.Windows.Forms.RadioButton()
        Me.optEditType13 = New System.Windows.Forms.RadioButton()
        Me.optEditType12 = New System.Windows.Forms.RadioButton()
        Me.optEditType11 = New System.Windows.Forms.RadioButton()
        Me.optEditType10 = New System.Windows.Forms.RadioButton()
        Me.optEditType9 = New System.Windows.Forms.RadioButton()
        Me.optEditType8 = New System.Windows.Forms.RadioButton()
        Me.optEditType7 = New System.Windows.Forms.RadioButton()
        Me.fraPaymentOptions = New System.Windows.Forms.GroupBox()
        Me.optPayType8 = New System.Windows.Forms.RadioButton()
        Me.optPayType7 = New System.Windows.Forms.RadioButton()
        Me.optPayType6 = New System.Windows.Forms.RadioButton()
        Me.optPayType5 = New System.Windows.Forms.RadioButton()
        Me.optPayType4 = New System.Windows.Forms.RadioButton()
        Me.optPayType3 = New System.Windows.Forms.RadioButton()
        Me.optPayType2 = New System.Windows.Forms.RadioButton()
        Me.optPayType1 = New System.Windows.Forms.RadioButton()
        Me.Notes_Frame = New System.Windows.Forms.GroupBox()
        Me.UGrSaleTotals = New WinCDS.UGridIO()
        Me.lblNewNotes = New System.Windows.Forms.Label()
        Me.lblOldNotes = New System.Windows.Forms.Label()
        Me.Notes_New = New System.Windows.Forms.TextBox()
        Me.Notes_Text = New System.Windows.Forms.TextBox()
        Me.UGridIO1 = New WinCDS.UGridIO()
        Me.fraCustomer = New System.Windows.Forms.GroupBox()
        Me.lblStatusCaption = New System.Windows.Forms.Label()
        Me.lblAccountCaption = New System.Windows.Forms.Label()
        Me.fraButtons = New System.Windows.Forms.GroupBox()
        Me.fraPayoffInfo = New System.Windows.Forms.GroupBox()
        Me.txtPayoffInfo = New System.Windows.Forms.TextBox()
        Me.fraArrearControl = New System.Windows.Forms.GroupBox()
        Me.lblArrearControlDisplay = New System.Windows.Forms.Label()
        Me.dtpArrearControlDate = New System.Windows.Forms.DateTimePicker()
        Me.lblArrearControlDate = New System.Windows.Forms.Label()
        Me.lblArrearControlGrace = New System.Windows.Forms.Label()
        Me.chkArrearControlGrace = New System.Windows.Forms.CheckBox()
        Me.fraPrintType = New System.Windows.Forms.GroupBox()
        Me.lblPrintType = New System.Windows.Forms.Label()
        Me.opt30323 = New System.Windows.Forms.RadioButton()
        Me.opt30252 = New System.Windows.Forms.RadioButton()
        Me.fraPrint.SuspendLayout()
        Me.fraNav.SuspendLayout()
        Me.fraTerms.SuspendLayout()
        Me.fraBalance.SuspendLayout()
        Me.fraEditOptions.SuspendLayout()
        Me.fraPaymentOptions.SuspendLayout()
        Me.Notes_Frame.SuspendLayout()
        Me.fraCustomer.SuspendLayout()
        Me.fraButtons.SuspendLayout()
        Me.fraPayoffInfo.SuspendLayout()
        Me.fraArrearControl.SuspendLayout()
        Me.fraPrintType.SuspendLayout()
        Me.SuspendLayout()
        '
        'lblTotalPayoff
        '
        Me.lblTotalPayoff.AutoSize = True
        Me.lblTotalPayoff.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblTotalPayoff.Location = New System.Drawing.Point(219, 28)
        Me.lblTotalPayoff.Name = "lblTotalPayoff"
        Me.lblTotalPayoff.Size = New System.Drawing.Size(91, 16)
        Me.lblTotalPayoff.TabIndex = 0
        Me.lblTotalPayoff.Text = "lblTotalPayoff"
        '
        'cmdApply
        '
        Me.cmdApply.Location = New System.Drawing.Point(70, 75)
        Me.cmdApply.Name = "cmdApply"
        Me.cmdApply.Size = New System.Drawing.Size(51, 23)
        Me.cmdApply.TabIndex = 1
        Me.cmdApply.Text = "&Apply"
        Me.ToolTip1.SetToolTip(Me.cmdApply, " Post The Entry ")
        Me.cmdApply.UseVisualStyleBackColor = True
        '
        'cmdEdit
        '
        Me.cmdEdit.Location = New System.Drawing.Point(223, 40)
        Me.cmdEdit.Name = "cmdEdit"
        Me.cmdEdit.Size = New System.Drawing.Size(54, 23)
        Me.cmdEdit.TabIndex = 2
        Me.cmdEdit.Text = "&Edit"
        Me.ToolTip1.SetToolTip(Me.cmdEdit, " Allows You To Edit A Custom Letter Selected in Microsoft WordPad ")
        Me.cmdEdit.UseVisualStyleBackColor = True
        '
        'cmdExport
        '
        Me.cmdExport.Location = New System.Drawing.Point(276, 19)
        Me.cmdExport.Name = "cmdExport"
        Me.cmdExport.Size = New System.Drawing.Size(54, 23)
        Me.cmdExport.TabIndex = 3
        Me.cmdExport.Text = "E&xport"
        Me.ToolTip1.SetToolTip(Me.cmdExport, " Prints The Custom Letter Selected To Customer ")
        Me.cmdExport.UseVisualStyleBackColor = True
        '
        'cmdFields
        '
        Me.cmdFields.Location = New System.Drawing.Point(276, 40)
        Me.cmdFields.Name = "cmdFields"
        Me.cmdFields.Size = New System.Drawing.Size(54, 23)
        Me.cmdFields.TabIndex = 4
        Me.cmdFields.Text = "&Fields"
        Me.ToolTip1.SetToolTip(Me.cmdFields, " Allows You To Edit A Custom Letter Selected in Microsoft WordPad ")
        Me.cmdFields.UseVisualStyleBackColor = True
        '
        'cmdMakeSameAsCash
        '
        Me.cmdMakeSameAsCash.Location = New System.Drawing.Point(462, 97)
        Me.cmdMakeSameAsCash.Name = "cmdMakeSameAsCash"
        Me.cmdMakeSameAsCash.Size = New System.Drawing.Size(40, 23)
        Me.cmdMakeSameAsCash.TabIndex = 5
        Me.cmdMakeSameAsCash.Text = "SaC"
        Me.cmdMakeSameAsCash.UseVisualStyleBackColor = True
        '
        'cmdMoveFirst
        '
        Me.cmdMoveFirst.Location = New System.Drawing.Point(9, 35)
        Me.cmdMoveFirst.Name = "cmdMoveFirst"
        Me.cmdMoveFirst.Size = New System.Drawing.Size(27, 25)
        Me.cmdMoveFirst.TabIndex = 6
        Me.cmdMoveFirst.Text = "<<"
        Me.ToolTip1.SetToolTip(Me.cmdMoveFirst, " Move To The First Record ")
        Me.cmdMoveFirst.UseVisualStyleBackColor = True
        '
        'cmdMoveLast
        '
        Me.cmdMoveLast.Location = New System.Drawing.Point(85, 35)
        Me.cmdMoveLast.Name = "cmdMoveLast"
        Me.cmdMoveLast.Size = New System.Drawing.Size(27, 25)
        Me.cmdMoveLast.TabIndex = 7
        Me.cmdMoveLast.Text = ">>"
        Me.ToolTip1.SetToolTip(Me.cmdMoveLast, " Move To The Last Record ")
        Me.cmdMoveLast.UseVisualStyleBackColor = True
        '
        'cmdMoveNext
        '
        Me.cmdMoveNext.Location = New System.Drawing.Point(60, 35)
        Me.cmdMoveNext.Name = "cmdMoveNext"
        Me.cmdMoveNext.Size = New System.Drawing.Size(27, 25)
        Me.cmdMoveNext.TabIndex = 8
        Me.cmdMoveNext.Text = ">"
        Me.ToolTip1.SetToolTip(Me.cmdMoveNext, " Move Forward 1 Record ")
        Me.cmdMoveNext.UseVisualStyleBackColor = True
        '
        'cmdMovePrevious
        '
        Me.cmdMovePrevious.Location = New System.Drawing.Point(34, 35)
        Me.cmdMovePrevious.Name = "cmdMovePrevious"
        Me.cmdMovePrevious.Size = New System.Drawing.Size(27, 25)
        Me.cmdMovePrevious.TabIndex = 9
        Me.cmdMovePrevious.Text = "<"
        Me.ToolTip1.SetToolTip(Me.cmdMovePrevious, " Move Back 1 Record ")
        Me.cmdMovePrevious.UseVisualStyleBackColor = True
        '
        'cmdPayoff
        '
        Me.cmdPayoff.Location = New System.Drawing.Point(229, 75)
        Me.cmdPayoff.Name = "cmdPayoff"
        Me.cmdPayoff.Size = New System.Drawing.Size(45, 43)
        Me.cmdPayoff.TabIndex = 10
        Me.cmdPayoff.Text = "&Early Payoff"
        Me.ToolTip1.SetToolTip(Me.cmdPayoff, " This Will Automatically Pay Off An Account.  Enter Final Payment ")
        Me.cmdPayoff.UseVisualStyleBackColor = True
        '
        'cmdReceipt
        '
        Me.cmdReceipt.Location = New System.Drawing.Point(120, 75)
        Me.cmdReceipt.Name = "cmdReceipt"
        Me.cmdReceipt.Size = New System.Drawing.Size(57, 23)
        Me.cmdReceipt.TabIndex = 11
        Me.cmdReceipt.Text = "&Receipt"
        Me.ToolTip1.SetToolTip(Me.cmdReceipt, "  Post The Entry And Print Receipt ")
        Me.cmdReceipt.UseVisualStyleBackColor = True
        '
        'cmdReprintContract
        '
        Me.cmdReprintContract.Location = New System.Drawing.Point(21, 171)
        Me.cmdReprintContract.Name = "cmdReprintContract"
        Me.cmdReprintContract.Size = New System.Drawing.Size(119, 22)
        Me.cmdReprintContract.TabIndex = 12
        Me.cmdReprintContract.Text = "Reprint Contract"
        Me.cmdReprintContract.UseVisualStyleBackColor = True
        '
        'cmdPrint
        '
        Me.cmdPrint.Location = New System.Drawing.Point(223, 19)
        Me.cmdPrint.Name = "cmdPrint"
        Me.cmdPrint.Size = New System.Drawing.Size(54, 23)
        Me.cmdPrint.TabIndex = 13
        Me.cmdPrint.Text = "&Print"
        Me.ToolTip1.SetToolTip(Me.cmdPrint, " Prints The Custom Letter Selected To Customer ")
        Me.cmdPrint.UseVisualStyleBackColor = True
        '
        'cmdCancel
        '
        Me.cmdCancel.Location = New System.Drawing.Point(176, 74)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.Size = New System.Drawing.Size(53, 23)
        Me.cmdCancel.TabIndex = 14
        Me.cmdCancel.Text = "&Cancel"
        Me.ToolTip1.SetToolTip(Me.cmdCancel, " Return To Main Menu ")
        Me.cmdCancel.UseVisualStyleBackColor = True
        '
        'chkSendAllMail
        '
        Me.chkSendAllMail.AutoSize = True
        Me.chkSendAllMail.Checked = True
        Me.chkSendAllMail.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkSendAllMail.Location = New System.Drawing.Point(154, 19)
        Me.chkSendAllMail.Name = "chkSendAllMail"
        Me.chkSendAllMail.Size = New System.Drawing.Size(87, 17)
        Me.chkSendAllMail.TabIndex = 71
        Me.chkSendAllMail.Text = "Send All Mail"
        Me.chkSendAllMail.UseVisualStyleBackColor = True
        '
        'txtPaymentHistory
        '
        Me.txtPaymentHistory.Location = New System.Drawing.Point(203, 169)
        Me.txtPaymentHistory.Name = "txtPaymentHistory"
        Me.txtPaymentHistory.Size = New System.Drawing.Size(91, 20)
        Me.txtPaymentHistory.TabIndex = 0
        Me.txtPaymentHistory.Text = "123456789ACB123456789ACB"
        Me.ToolTip1.SetToolTip(Me.txtPaymentHistory, "Customer Payment History on Account -- SEE HELP FILE FOR MORE INFO (Press F1).")
        Me.txtPaymentHistory.Visible = False
        '
        'lblBalance
        '
        Me.lblBalance.AutoSize = True
        Me.lblBalance.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblBalance.Location = New System.Drawing.Point(11, 28)
        Me.lblBalance.Name = "lblBalance"
        Me.lblBalance.Size = New System.Drawing.Size(82, 16)
        Me.lblBalance.TabIndex = 17
        Me.lblBalance.Text = "lblBalance"
        '
        'txtDelivery
        '
        Me.txtDelivery.Location = New System.Drawing.Point(235, 14)
        Me.txtDelivery.Name = "txtDelivery"
        Me.txtDelivery.Size = New System.Drawing.Size(100, 20)
        Me.txtDelivery.TabIndex = 18
        Me.ToolTip1.SetToolTip(Me.txtDelivery, "Sale Delivery Date:")
        '
        'txtMonths
        '
        Me.txtMonths.Location = New System.Drawing.Point(65, 40)
        Me.txtMonths.Name = "txtMonths"
        Me.txtMonths.Size = New System.Drawing.Size(92, 20)
        Me.txtMonths.TabIndex = 19
        '
        'txtLastPay
        '
        Me.txtLastPay.Location = New System.Drawing.Point(235, 86)
        Me.txtLastPay.Name = "txtLastPay"
        Me.txtLastPay.Size = New System.Drawing.Size(100, 20)
        Me.txtLastPay.TabIndex = 20
        Me.ToolTip1.SetToolTip(Me.txtLastPay, "Last Payment Date:")
        '
        'txtFinanced
        '
        Me.txtFinanced.Location = New System.Drawing.Point(65, 20)
        Me.txtFinanced.Name = "txtFinanced"
        Me.txtFinanced.Size = New System.Drawing.Size(92, 20)
        Me.txtFinanced.TabIndex = 21
        '
        'txtMonthlyPayment
        '
        Me.txtMonthlyPayment.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtMonthlyPayment.Location = New System.Drawing.Point(64, 80)
        Me.txtMonthlyPayment.Name = "txtMonthlyPayment"
        Me.txtMonthlyPayment.Size = New System.Drawing.Size(68, 20)
        Me.txtMonthlyPayment.TabIndex = 22
        '
        'txtPayMemo
        '
        Me.txtPayMemo.Location = New System.Drawing.Point(50, 97)
        Me.txtPayMemo.Name = "txtPayMemo"
        Me.txtPayMemo.Size = New System.Drawing.Size(191, 20)
        Me.txtPayMemo.TabIndex = 23
        '
        'txtRate
        '
        Me.txtRate.Location = New System.Drawing.Point(65, 60)
        Me.txtRate.Name = "txtRate"
        Me.txtRate.Size = New System.Drawing.Size(35, 20)
        Me.txtRate.TabIndex = 24
        '
        'txtSameAsCash
        '
        Me.txtSameAsCash.Location = New System.Drawing.Point(286, 58)
        Me.txtSameAsCash.Name = "txtSameAsCash"
        Me.txtSameAsCash.Size = New System.Drawing.Size(49, 20)
        Me.txtSameAsCash.TabIndex = 25
        '
        'lblLateCharge
        '
        Me.lblLateCharge.AutoSize = True
        Me.lblLateCharge.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblLateCharge.Location = New System.Drawing.Point(108, 28)
        Me.lblLateCharge.Name = "lblLateCharge"
        Me.lblLateCharge.Size = New System.Drawing.Size(105, 16)
        Me.lblLateCharge.TabIndex = 26
        Me.lblLateCharge.Text = "lblLateCharge"
        '
        'txtFirstPay
        '
        Me.txtFirstPay.Location = New System.Drawing.Point(235, 34)
        Me.txtFirstPay.Name = "txtFirstPay"
        Me.txtFirstPay.Size = New System.Drawing.Size(100, 20)
        Me.txtFirstPay.TabIndex = 27
        Me.ToolTip1.SetToolTip(Me.txtFirstPay, "First Payment Date:")
        '
        'txtPaidBy
        '
        Me.txtPaidBy.Location = New System.Drawing.Point(63, 120)
        Me.txtPaidBy.Name = "txtPaidBy"
        Me.txtPaidBy.Size = New System.Drawing.Size(92, 20)
        Me.txtPaidBy.TabIndex = 28
        '
        'txtPayPeriod
        '
        Me.txtPayPeriod.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtPayPeriod.Location = New System.Drawing.Point(131, 79)
        Me.txtPayPeriod.Name = "txtPayPeriod"
        Me.txtPayPeriod.Size = New System.Drawing.Size(26, 20)
        Me.txtPayPeriod.TabIndex = 29
        '
        'txtNextDue
        '
        Me.txtNextDue.Location = New System.Drawing.Point(235, 119)
        Me.txtNextDue.Name = "txtNextDue"
        Me.txtNextDue.Size = New System.Drawing.Size(100, 20)
        Me.txtNextDue.TabIndex = 30
        Me.ToolTip1.SetToolTip(Me.txtNextDue, "Next Payment Due On:")
        '
        'lblTotDue
        '
        Me.lblTotDue.BackColor = System.Drawing.SystemColors.Window
        Me.lblTotDue.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblTotDue.Location = New System.Drawing.Point(172, 55)
        Me.lblTotDue.Name = "lblTotDue"
        Me.lblTotDue.Size = New System.Drawing.Size(142, 16)
        Me.lblTotDue.TabIndex = 31
        '
        'lblArrearages
        '
        Me.lblArrearages.BackColor = System.Drawing.SystemColors.Window
        Me.lblArrearages.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblArrearages.Location = New System.Drawing.Point(13, 55)
        Me.lblArrearages.Name = "lblArrearages"
        Me.lblArrearages.Size = New System.Drawing.Size(147, 16)
        Me.lblArrearages.TabIndex = 32
        '
        'lblLate0
        '
        Me.lblLate0.BackColor = System.Drawing.SystemColors.Window
        Me.lblLate0.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblLate0.Location = New System.Drawing.Point(13, 34)
        Me.lblLate0.Name = "lblLate0"
        Me.lblLate0.Size = New System.Drawing.Size(58, 16)
        Me.lblLate0.TabIndex = 33
        '
        'lblLate31
        '
        Me.lblLate31.BackColor = System.Drawing.SystemColors.Window
        Me.lblLate31.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblLate31.Location = New System.Drawing.Point(98, 34)
        Me.lblLate31.Name = "lblLate31"
        Me.lblLate31.Size = New System.Drawing.Size(58, 16)
        Me.lblLate31.TabIndex = 34
        '
        'lblLate61
        '
        Me.lblLate61.BackColor = System.Drawing.SystemColors.Window
        Me.lblLate61.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblLate61.Location = New System.Drawing.Point(172, 34)
        Me.lblLate61.Name = "lblLate61"
        Me.lblLate61.Size = New System.Drawing.Size(58, 16)
        Me.lblLate61.TabIndex = 35
        '
        'lblLate91
        '
        Me.lblLate91.BackColor = System.Drawing.SystemColors.Window
        Me.lblLate91.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblLate91.Location = New System.Drawing.Point(256, 34)
        Me.lblLate91.Name = "lblLate91"
        Me.lblLate91.Size = New System.Drawing.Size(58, 16)
        Me.lblLate91.TabIndex = 36
        '
        'lbl0030
        '
        Me.lbl0030.AutoSize = True
        Me.lbl0030.Location = New System.Drawing.Point(10, 17)
        Me.lbl0030.Name = "lbl0030"
        Me.lbl0030.Size = New System.Drawing.Size(61, 13)
        Me.lbl0030.TabIndex = 37
        Me.lbl0030.Text = "Late: 0 - 30"
        '
        'lbl3160
        '
        Me.lbl3160.AutoSize = True
        Me.lbl3160.Location = New System.Drawing.Point(98, 17)
        Me.lbl3160.Name = "lbl3160"
        Me.lbl3160.Size = New System.Drawing.Size(40, 13)
        Me.lbl3160.TabIndex = 38
        Me.lbl3160.Text = "31 - 60"
        '
        'lbl6190
        '
        Me.lbl6190.AutoSize = True
        Me.lbl6190.Location = New System.Drawing.Point(169, 16)
        Me.lbl6190.Name = "lbl6190"
        Me.lbl6190.Size = New System.Drawing.Size(40, 13)
        Me.lbl6190.TabIndex = 39
        Me.lbl6190.Text = "61 - 90"
        '
        'lblOver91
        '
        Me.lblOver91.AutoSize = True
        Me.lblOver91.Location = New System.Drawing.Point(254, 17)
        Me.lblOver91.Name = "lblOver91"
        Me.lblOver91.Size = New System.Drawing.Size(45, 13)
        Me.lblOver91.TabIndex = 40
        Me.lblOver91.Text = "Over 91"
        '
        'filFile
        '
        Me.filFile.FormattingEnabled = True
        Me.filFile.Location = New System.Drawing.Point(6, 19)
        Me.filFile.Name = "filFile"
        Me.filFile.Pattern = "*.txt;*.rtf"
        Me.filFile.Size = New System.Drawing.Size(203, 43)
        Me.filFile.TabIndex = 42
        '
        'lblAccount
        '
        Me.lblAccount.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblAccount.Location = New System.Drawing.Point(63, 20)
        Me.lblAccount.Name = "lblAccount"
        Me.lblAccount.Size = New System.Drawing.Size(93, 16)
        Me.lblAccount.TabIndex = 43
        Me.lblAccount.Text = "lblAccount"
        '
        'cboStatus
        '
        Me.cboStatus.FormattingEnabled = True
        Me.cboStatus.Location = New System.Drawing.Point(203, 22)
        Me.cboStatus.Name = "cboStatus"
        Me.cboStatus.Size = New System.Drawing.Size(121, 21)
        Me.cboStatus.TabIndex = 44
        '
        'lblFirstName
        '
        Me.lblFirstName.Location = New System.Drawing.Point(9, 48)
        Me.lblFirstName.Name = "lblFirstName"
        Me.lblFirstName.Size = New System.Drawing.Size(144, 22)
        Me.lblFirstName.TabIndex = 45
        Me.lblFirstName.Text = "lblFirstName"
        '
        'lblLastName
        '
        Me.lblLastName.Location = New System.Drawing.Point(176, 48)
        Me.lblLastName.Name = "lblLastName"
        Me.lblLastName.Size = New System.Drawing.Size(148, 29)
        Me.lblLastName.TabIndex = 46
        Me.lblLastName.Text = "lblLastName"
        '
        'lblAddress
        '
        Me.lblAddress.Location = New System.Drawing.Point(9, 68)
        Me.lblAddress.Name = "lblAddress"
        Me.lblAddress.Size = New System.Drawing.Size(310, 30)
        Me.lblAddress.TabIndex = 47
        Me.lblAddress.Text = "lblAddress"
        '
        'lblCity
        '
        Me.lblCity.Location = New System.Drawing.Point(9, 129)
        Me.lblCity.Name = "lblCity"
        Me.lblCity.Size = New System.Drawing.Size(181, 15)
        Me.lblCity.TabIndex = 48
        Me.lblCity.Text = "lblCity"
        '
        'lblZip
        '
        Me.lblZip.AutoSize = True
        Me.lblZip.Location = New System.Drawing.Point(213, 130)
        Me.lblZip.Name = "lblZip"
        Me.lblZip.Size = New System.Drawing.Size(32, 13)
        Me.lblZip.TabIndex = 49
        Me.lblZip.Text = "lblZip"
        '
        'lblTele1
        '
        Me.lblTele1.AutoSize = True
        Me.lblTele1.Location = New System.Drawing.Point(47, 148)
        Me.lblTele1.Name = "lblTele1"
        Me.lblTele1.Size = New System.Drawing.Size(44, 13)
        Me.lblTele1.TabIndex = 50
        Me.lblTele1.Text = "lblTele1"
        '
        'lblTele2
        '
        Me.lblTele2.AutoSize = True
        Me.lblTele2.Location = New System.Drawing.Point(47, 168)
        Me.lblTele2.Name = "lblTele2"
        Me.lblTele2.Size = New System.Drawing.Size(44, 13)
        Me.lblTele2.TabIndex = 51
        Me.lblTele2.Text = "lblTele2"
        '
        'lblSSN
        '
        Me.lblSSN.Location = New System.Drawing.Point(168, 9)
        Me.lblSSN.Name = "lblSSN"
        Me.lblSSN.Size = New System.Drawing.Size(24, 15)
        Me.lblSSN.TabIndex = 52
        Me.lblSSN.Text = "lblSSN"
        Me.lblSSN.Visible = False
        '
        'txtLateChargeAmount
        '
        Me.txtLateChargeAmount.Location = New System.Drawing.Point(64, 100)
        Me.txtLateChargeAmount.Name = "txtLateChargeAmount"
        Me.txtLateChargeAmount.Size = New System.Drawing.Size(92, 20)
        Me.txtLateChargeAmount.TabIndex = 53
        '
        'lblCreditLimit
        '
        Me.lblCreditLimit.Location = New System.Drawing.Point(202, 9)
        Me.lblCreditLimit.Name = "lblCreditLimit"
        Me.lblCreditLimit.Size = New System.Drawing.Size(22, 13)
        Me.lblCreditLimit.TabIndex = 54
        Me.lblCreditLimit.Text = "lblCreditLimit"
        Me.lblCreditLimit.Visible = False
        '
        'lblApprovalTerms
        '
        Me.lblApprovalTerms.Location = New System.Drawing.Point(251, 9)
        Me.lblApprovalTerms.Name = "lblApprovalTerms"
        Me.lblApprovalTerms.Size = New System.Drawing.Size(28, 10)
        Me.lblApprovalTerms.TabIndex = 55
        Me.lblApprovalTerms.Text = "lblApprovalTerms"
        Me.lblApprovalTerms.Visible = False
        '
        'lblAddAddress
        '
        Me.lblAddAddress.Location = New System.Drawing.Point(9, 102)
        Me.lblAddAddress.Name = "lblAddAddress"
        Me.lblAddAddress.Size = New System.Drawing.Size(314, 24)
        Me.lblAddAddress.TabIndex = 56
        Me.lblAddAddress.Text = "lblAddAddress"
        '
        'lblTele3
        '
        Me.lblTele3.AutoSize = True
        Me.lblTele3.Location = New System.Drawing.Point(47, 187)
        Me.lblTele3.Name = "lblTele3"
        Me.lblTele3.Size = New System.Drawing.Size(44, 13)
        Me.lblTele3.TabIndex = 57
        Me.lblTele3.Text = "lblTele3"
        '
        'lblTele1Caption
        '
        Me.lblTele1Caption.AutoSize = True
        Me.lblTele1Caption.Location = New System.Drawing.Point(9, 148)
        Me.lblTele1Caption.Name = "lblTele1Caption"
        Me.lblTele1Caption.Size = New System.Drawing.Size(40, 13)
        Me.lblTele1Caption.TabIndex = 58
        Me.lblTele1Caption.Text = "Tele 1:"
        '
        'lblTele2Caption
        '
        Me.lblTele2Caption.AutoSize = True
        Me.lblTele2Caption.Location = New System.Drawing.Point(9, 168)
        Me.lblTele2Caption.Name = "lblTele2Caption"
        Me.lblTele2Caption.Size = New System.Drawing.Size(40, 13)
        Me.lblTele2Caption.TabIndex = 59
        Me.lblTele2Caption.Text = "Tele 2:"
        '
        'lblTele3Caption
        '
        Me.lblTele3Caption.AutoSize = True
        Me.lblTele3Caption.Location = New System.Drawing.Point(9, 187)
        Me.lblTele3Caption.Name = "lblTele3Caption"
        Me.lblTele3Caption.Size = New System.Drawing.Size(37, 13)
        Me.lblTele3Caption.TabIndex = 60
        Me.lblTele3Caption.Text = "Tele 3"
        '
        'fraPrint
        '
        Me.fraPrint.Controls.Add(Me.cmdExport)
        Me.fraPrint.Controls.Add(Me.cmdPrint)
        Me.fraPrint.Controls.Add(Me.rtfFile)
        Me.fraPrint.Controls.Add(Me.filFile)
        Me.fraPrint.Controls.Add(Me.cmdEdit)
        Me.fraPrint.Controls.Add(Me.cmdFields)
        Me.fraPrint.Location = New System.Drawing.Point(356, 408)
        Me.fraPrint.Name = "fraPrint"
        Me.fraPrint.Size = New System.Drawing.Size(342, 68)
        Me.fraPrint.TabIndex = 61
        Me.fraPrint.TabStop = False
        Me.fraPrint.Text = "Letter to Customer:"
        '
        'rtfFile
        '
        Me.rtfFile.Location = New System.Drawing.Point(236, 9)
        Me.rtfFile.Name = "rtfFile"
        Me.rtfFile.Size = New System.Drawing.Size(49, 16)
        Me.rtfFile.TabIndex = 43
        Me.rtfFile.Text = ""
        Me.rtfFile.Visible = False
        '
        'fraNav
        '
        Me.fraNav.Controls.Add(Me.fraPrintType)
        Me.fraNav.Controls.Add(Me.cmdPrintLabel)
        Me.fraNav.Controls.Add(Me.chkSendAllMail)
        Me.fraNav.Controls.Add(Me.cmdMoveFirst)
        Me.fraNav.Controls.Add(Me.cmdMovePrevious)
        Me.fraNav.Controls.Add(Me.cmdMoveNext)
        Me.fraNav.Controls.Add(Me.cmdMoveLast)
        Me.fraNav.Location = New System.Drawing.Point(12, 403)
        Me.fraNav.Name = "fraNav"
        Me.fraNav.Size = New System.Drawing.Size(335, 73)
        Me.fraNav.TabIndex = 62
        Me.fraNav.TabStop = False
        Me.fraNav.Text = "Move Records:"
        '
        'cmdPrintLabel
        '
        Me.cmdPrintLabel.Location = New System.Drawing.Point(147, 44)
        Me.cmdPrintLabel.Name = "cmdPrintLabel"
        Me.cmdPrintLabel.Size = New System.Drawing.Size(110, 23)
        Me.cmdPrintLabel.TabIndex = 72
        Me.cmdPrintLabel.Text = "&Print Address Label"
        Me.ToolTip1.SetToolTip(Me.cmdPrintLabel, " Dymo 330 Turbo Required  ")
        Me.cmdPrintLabel.UseVisualStyleBackColor = True
        '
        'fraTerms
        '
        Me.fraTerms.Controls.Add(Me.cmdReprintCoupons)
        Me.fraTerms.Controls.Add(Me.lblNextDue)
        Me.fraTerms.Controls.Add(Me.lblLastPay)
        Me.fraTerms.Controls.Add(Me.lblSameAsCash)
        Me.fraTerms.Controls.Add(Me.lbl1stPay)
        Me.fraTerms.Controls.Add(Me.lblDelivery)
        Me.fraTerms.Controls.Add(Me.lblPayBy)
        Me.fraTerms.Controls.Add(Me.lblLateCh)
        Me.fraTerms.Controls.Add(Me.lblPayment)
        Me.fraTerms.Controls.Add(Me.lblRate)
        Me.fraTerms.Controls.Add(Me.lblMonths)
        Me.fraTerms.Controls.Add(Me.lblPaymentHistory)
        Me.fraTerms.Controls.Add(Me.lblAPR)
        Me.fraTerms.Controls.Add(Me.lblFinanced)
        Me.fraTerms.Controls.Add(Me.txtFinanced)
        Me.fraTerms.Controls.Add(Me.txtMonths)
        Me.fraTerms.Controls.Add(Me.txtRate)
        Me.fraTerms.Controls.Add(Me.txtMonthlyPayment)
        Me.fraTerms.Controls.Add(Me.txtLateChargeAmount)
        Me.fraTerms.Controls.Add(Me.txtPayPeriod)
        Me.fraTerms.Controls.Add(Me.txtPaidBy)
        Me.fraTerms.Controls.Add(Me.txtDelivery)
        Me.fraTerms.Controls.Add(Me.txtFirstPay)
        Me.fraTerms.Controls.Add(Me.txtSameAsCash)
        Me.fraTerms.Controls.Add(Me.txtLastPay)
        Me.fraTerms.Controls.Add(Me.txtNextDue)
        Me.fraTerms.Controls.Add(Me.cmdReprintContract)
        Me.fraTerms.Controls.Add(Me.txtPaymentHistory)
        Me.fraTerms.Location = New System.Drawing.Point(357, 4)
        Me.fraTerms.Name = "fraTerms"
        Me.fraTerms.Size = New System.Drawing.Size(341, 199)
        Me.fraTerms.TabIndex = 63
        Me.fraTerms.TabStop = False
        Me.fraTerms.Text = " Terms And Conditions "
        '
        'cmdReprintCoupons
        '
        Me.cmdReprintCoupons.Location = New System.Drawing.Point(21, 147)
        Me.cmdReprintCoupons.Name = "cmdReprintCoupons"
        Me.cmdReprintCoupons.Size = New System.Drawing.Size(119, 23)
        Me.cmdReprintCoupons.TabIndex = 59
        Me.cmdReprintCoupons.Text = "Reprint Coupons"
        Me.cmdReprintCoupons.UseVisualStyleBackColor = True
        '
        'lblNextDue
        '
        Me.lblNextDue.AutoSize = True
        Me.lblNextDue.Location = New System.Drawing.Point(180, 119)
        Me.lblNextDue.Name = "lblNextDue"
        Me.lblNextDue.Size = New System.Drawing.Size(55, 13)
        Me.lblNextDue.TabIndex = 58
        Me.lblNextDue.Text = "Next Due:"
        '
        'lblLastPay
        '
        Me.lblLastPay.AutoSize = True
        Me.lblLastPay.Location = New System.Drawing.Point(184, 92)
        Me.lblLastPay.Name = "lblLastPay"
        Me.lblLastPay.Size = New System.Drawing.Size(51, 13)
        Me.lblLastPay.TabIndex = 57
        Me.lblLastPay.Text = "Last Pay:"
        '
        'lblSameAsCash
        '
        Me.lblSameAsCash.AutoSize = True
        Me.lblSameAsCash.Location = New System.Drawing.Point(169, 63)
        Me.lblSameAsCash.Name = "lblSameAsCash"
        Me.lblSameAsCash.Size = New System.Drawing.Size(79, 13)
        Me.lblSameAsCash.TabIndex = 56
        Me.lblSameAsCash.Text = "Same As Cash:"
        '
        'lbl1stPay
        '
        Me.lbl1stPay.AutoSize = True
        Me.lbl1stPay.Location = New System.Drawing.Point(185, 36)
        Me.lbl1stPay.Name = "lbl1stPay"
        Me.lbl1stPay.Size = New System.Drawing.Size(50, 13)
        Me.lbl1stPay.TabIndex = 55
        Me.lbl1stPay.Text = "1St. Pay:"
        '
        'lblDelivery
        '
        Me.lblDelivery.AutoSize = True
        Me.lblDelivery.Location = New System.Drawing.Point(187, 16)
        Me.lblDelivery.Name = "lblDelivery"
        Me.lblDelivery.Size = New System.Drawing.Size(48, 13)
        Me.lblDelivery.TabIndex = 54
        Me.lblDelivery.Text = "Delivery:"
        '
        'lblPayBy
        '
        Me.lblPayBy.AutoSize = True
        Me.lblPayBy.Location = New System.Drawing.Point(18, 124)
        Me.lblPayBy.Name = "lblPayBy"
        Me.lblPayBy.Size = New System.Drawing.Size(43, 13)
        Me.lblPayBy.TabIndex = 5
        Me.lblPayBy.Text = "Pay By:"
        '
        'lblLateCh
        '
        Me.lblLateCh.AutoSize = True
        Me.lblLateCh.Location = New System.Drawing.Point(2, 104)
        Me.lblLateCh.Name = "lblLateCh"
        Me.lblLateCh.Size = New System.Drawing.Size(59, 13)
        Me.lblLateCh.TabIndex = 4
        Me.lblLateCh.Text = "Late Chge:"
        '
        'lblPayment
        '
        Me.lblPayment.AutoSize = True
        Me.lblPayment.Location = New System.Drawing.Point(10, 83)
        Me.lblPayment.Name = "lblPayment"
        Me.lblPayment.Size = New System.Drawing.Size(51, 13)
        Me.lblPayment.TabIndex = 3
        Me.lblPayment.Text = "Payment:"
        '
        'lblRate
        '
        Me.lblRate.AutoSize = True
        Me.lblRate.Location = New System.Drawing.Point(28, 63)
        Me.lblRate.Name = "lblRate"
        Me.lblRate.Size = New System.Drawing.Size(33, 13)
        Me.lblRate.TabIndex = 2
        Me.lblRate.Text = "Rate:"
        '
        'lblMonths
        '
        Me.lblMonths.AutoSize = True
        Me.lblMonths.Location = New System.Drawing.Point(16, 43)
        Me.lblMonths.Name = "lblMonths"
        Me.lblMonths.Size = New System.Drawing.Size(45, 13)
        Me.lblMonths.TabIndex = 1
        Me.lblMonths.Text = "Months:"
        '
        'lblPaymentHistory
        '
        Me.lblPaymentHistory.AutoSize = True
        Me.lblPaymentHistory.Location = New System.Drawing.Point(200, 155)
        Me.lblPaymentHistory.Name = "lblPaymentHistory"
        Me.lblPaymentHistory.Size = New System.Drawing.Size(94, 13)
        Me.lblPaymentHistory.TabIndex = 72
        Me.lblPaymentHistory.Text = "Date: 00/00/0000"
        Me.ToolTip1.SetToolTip(Me.lblPaymentHistory, "Customer Payment History on Account -- SEE HELP FILE FOR MORE INFO (Press F1).")
        Me.lblPaymentHistory.Visible = False
        '
        'lblAPR
        '
        Me.lblAPR.AutoSize = True
        Me.lblAPR.Location = New System.Drawing.Point(102, 63)
        Me.lblAPR.Name = "lblAPR"
        Me.lblAPR.Size = New System.Drawing.Size(63, 13)
        Me.lblAPR.TabIndex = 70
        Me.lblAPR.Text = "##.## APR"
        '
        'lblFinanced
        '
        Me.lblFinanced.AutoSize = True
        Me.lblFinanced.Location = New System.Drawing.Point(7, 25)
        Me.lblFinanced.Name = "lblFinanced"
        Me.lblFinanced.Size = New System.Drawing.Size(54, 13)
        Me.lblFinanced.TabIndex = 0
        Me.lblFinanced.Text = "Financed:"
        '
        'fraBalance
        '
        Me.fraBalance.Controls.Add(Me.txtPayment)
        Me.fraBalance.Controls.Add(Me.lblPayDate)
        Me.fraBalance.Controls.Add(Me.txtLateCharge)
        Me.fraBalance.Controls.Add(Me.lblTotalPayoff)
        Me.fraBalance.Controls.Add(Me.chkPayLateFee)
        Me.fraBalance.Controls.Add(Me.lblTotalPayOffCaption)
        Me.fraBalance.Controls.Add(Me.lblLateChargeCaption)
        Me.fraBalance.Controls.Add(Me.lblBalanceCaption)
        Me.fraBalance.Controls.Add(Me.DDate)
        Me.fraBalance.Controls.Add(Me.lblPayMemo)
        Me.fraBalance.Controls.Add(Me.txtPayMemo)
        Me.fraBalance.Controls.Add(Me.lblBalance)
        Me.fraBalance.Controls.Add(Me.lblLateCharge)
        Me.fraBalance.Location = New System.Drawing.Point(12, 202)
        Me.fraBalance.Name = "fraBalance"
        Me.fraBalance.Size = New System.Drawing.Size(335, 122)
        Me.fraBalance.TabIndex = 64
        Me.fraBalance.TabStop = False
        '
        'txtPayment
        '
        Me.txtPayment.Location = New System.Drawing.Point(14, 69)
        Me.txtPayment.Name = "txtPayment"
        Me.txtPayment.Size = New System.Drawing.Size(96, 20)
        Me.txtPayment.TabIndex = 87
        '
        'lblPayDate
        '
        Me.lblPayDate.AutoSize = True
        Me.lblPayDate.Location = New System.Drawing.Point(251, 100)
        Me.lblPayDate.Name = "lblPayDate"
        Me.lblPayDate.Size = New System.Drawing.Size(54, 13)
        Me.lblPayDate.TabIndex = 86
        Me.lblPayDate.Text = "Pay Date:"
        '
        'txtLateCharge
        '
        Me.txtLateCharge.Location = New System.Drawing.Point(136, 69)
        Me.txtLateCharge.Name = "txtLateCharge"
        Me.txtLateCharge.Size = New System.Drawing.Size(76, 20)
        Me.txtLateCharge.TabIndex = 85
        '
        'chkPayLateFee
        '
        Me.chkPayLateFee.AutoSize = True
        Me.chkPayLateFee.Checked = True
        Me.chkPayLateFee.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkPayLateFee.Location = New System.Drawing.Point(14, 50)
        Me.chkPayLateFee.Name = "chkPayLateFee"
        Me.chkPayLateFee.Size = New System.Drawing.Size(105, 17)
        Me.chkPayLateFee.TabIndex = 72
        Me.chkPayLateFee.Text = "Pay Late Charge"
        Me.chkPayLateFee.UseVisualStyleBackColor = True
        '
        'lblTotalPayOffCaption
        '
        Me.lblTotalPayOffCaption.AutoSize = True
        Me.lblTotalPayOffCaption.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblTotalPayOffCaption.Location = New System.Drawing.Point(217, 11)
        Me.lblTotalPayOffCaption.Name = "lblTotalPayOffCaption"
        Me.lblTotalPayOffCaption.Size = New System.Drawing.Size(88, 16)
        Me.lblTotalPayOffCaption.TabIndex = 28
        Me.lblTotalPayOffCaption.Text = "Total Pay Off:"
        Me.ToolTip1.SetToolTip(Me.lblTotalPayOffCaption, " Balance Due After Early Payoff ")
        '
        'lblLateChargeCaption
        '
        Me.lblLateChargeCaption.AutoSize = True
        Me.lblLateChargeCaption.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblLateChargeCaption.Location = New System.Drawing.Point(108, 11)
        Me.lblLateChargeCaption.Name = "lblLateChargeCaption"
        Me.lblLateChargeCaption.Size = New System.Drawing.Size(84, 16)
        Me.lblLateChargeCaption.TabIndex = 27
        Me.lblLateChargeCaption.Text = "Late Charge:"
        Me.ToolTip1.SetToolTip(Me.lblLateChargeCaption, " Late Charge Balance Only ")
        '
        'lblBalanceCaption
        '
        Me.lblBalanceCaption.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblBalanceCaption.Location = New System.Drawing.Point(6, 9)
        Me.lblBalanceCaption.Name = "lblBalanceCaption"
        Me.lblBalanceCaption.Size = New System.Drawing.Size(77, 18)
        Me.lblBalanceCaption.TabIndex = 18
        Me.lblBalanceCaption.Text = " Balance Including Late Charges "
        '
        'DDate
        '
        Me.DDate.CustomFormat = "MM/dd/yyyy"
        Me.DDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DDate.Location = New System.Drawing.Point(238, 69)
        Me.DDate.Name = "DDate"
        Me.DDate.Size = New System.Drawing.Size(91, 20)
        Me.DDate.TabIndex = 80
        '
        'lblPayMemo
        '
        Me.lblPayMemo.AutoSize = True
        Me.lblPayMemo.Location = New System.Drawing.Point(10, 100)
        Me.lblPayMemo.Name = "lblPayMemo"
        Me.lblPayMemo.Size = New System.Drawing.Size(39, 13)
        Me.lblPayMemo.TabIndex = 75
        Me.lblPayMemo.Text = "Memo:"
        '
        'cmdCreditApp
        '
        Me.cmdCreditApp.Location = New System.Drawing.Point(4, 74)
        Me.cmdCreditApp.Name = "cmdCreditApp"
        Me.cmdCreditApp.Size = New System.Drawing.Size(67, 44)
        Me.cmdCreditApp.TabIndex = 66
        Me.cmdCreditApp.Text = "Credi&t App"
        Me.ToolTip1.SetToolTip(Me.cmdCreditApp, " Fill Out A Credit App For Customer On Screen ")
        Me.cmdCreditApp.UseVisualStyleBackColor = True
        '
        'cmdDetail
        '
        Me.cmdDetail.Location = New System.Drawing.Point(70, 96)
        Me.cmdDetail.Name = "cmdDetail"
        Me.cmdDetail.Size = New System.Drawing.Size(51, 23)
        Me.cmdDetail.TabIndex = 67
        Me.cmdDetail.Text = "&Detail"
        Me.ToolTip1.SetToolTip(Me.cmdDetail, " Open Up Account Transactions ")
        Me.cmdDetail.UseVisualStyleBackColor = True
        '
        'Notes_Open
        '
        Me.Notes_Open.Location = New System.Drawing.Point(120, 96)
        Me.Notes_Open.Name = "Notes_Open"
        Me.Notes_Open.Size = New System.Drawing.Size(57, 23)
        Me.Notes_Open.TabIndex = 68
        Me.Notes_Open.Text = "&Notes"
        Me.ToolTip1.SetToolTip(Me.Notes_Open, " Add Hidden Notes To This Account ")
        Me.Notes_Open.UseVisualStyleBackColor = True
        '
        'cmdPrintCard
        '
        Me.cmdPrintCard.Location = New System.Drawing.Point(176, 95)
        Me.cmdPrintCard.Name = "cmdPrintCard"
        Me.cmdPrintCard.Size = New System.Drawing.Size(53, 23)
        Me.cmdPrintCard.TabIndex = 69
        Me.cmdPrintCard.Text = "&Print"
        Me.ToolTip1.SetToolTip(Me.cmdPrintCard, " Print Account Card ")
        Me.cmdPrintCard.UseVisualStyleBackColor = True
        '
        'cmdSaleTotals
        '
        Me.cmdSaleTotals.Location = New System.Drawing.Point(273, 75)
        Me.cmdSaleTotals.Name = "cmdSaleTotals"
        Me.cmdSaleTotals.Size = New System.Drawing.Size(49, 43)
        Me.cmdSaleTotals.TabIndex = 73
        Me.cmdSaleTotals.Text = "Sale Totals"
        Me.ToolTip1.SetToolTip(Me.cmdSaleTotals, " This Will Automatically Pay Off An Account.  Enter Final Payment ")
        Me.cmdSaleTotals.UseVisualStyleBackColor = True
        Me.cmdSaleTotals.Visible = False
        '
        'optPayType9
        '
        Me.optPayType9.AutoSize = True
        Me.optPayType9.Location = New System.Drawing.Point(308, 44)
        Me.optPayType9.Name = "optPayType9"
        Me.optPayType9.Size = New System.Drawing.Size(92, 17)
        Me.optPayType9.TabIndex = 8
        Me.optPayType9.TabStop = True
        Me.optPayType9.Text = "Co Ck Refund"
        Me.ToolTip1.SetToolTip(Me.optPayType9, "Company Check Refund")
        Me.optPayType9.UseVisualStyleBackColor = True
        '
        'Notes_Save
        '
        Me.Notes_Save.Location = New System.Drawing.Point(615, 37)
        Me.Notes_Save.Name = "Notes_Save"
        Me.Notes_Save.Size = New System.Drawing.Size(59, 23)
        Me.Notes_Save.TabIndex = 45
        Me.Notes_Save.Text = "&Save"
        Me.ToolTip1.SetToolTip(Me.Notes_Save, " Saves The Note Created ")
        Me.Notes_Save.UseVisualStyleBackColor = True
        '
        'cmdNotesCancel
        '
        Me.cmdNotesCancel.Location = New System.Drawing.Point(615, 58)
        Me.cmdNotesCancel.Name = "cmdNotesCancel"
        Me.cmdNotesCancel.Size = New System.Drawing.Size(59, 23)
        Me.cmdNotesCancel.TabIndex = 46
        Me.cmdNotesCancel.Text = "&Cancel"
        Me.ToolTip1.SetToolTip(Me.cmdNotesCancel, " Cancel Notes ")
        Me.cmdNotesCancel.UseVisualStyleBackColor = True
        '
        'cmdNotesPrint
        '
        Me.cmdNotesPrint.Location = New System.Drawing.Point(615, 79)
        Me.cmdNotesPrint.Name = "cmdNotesPrint"
        Me.cmdNotesPrint.Size = New System.Drawing.Size(59, 23)
        Me.cmdNotesPrint.TabIndex = 47
        Me.cmdNotesPrint.Text = "&Print"
        Me.ToolTip1.SetToolTip(Me.cmdNotesPrint, " Prints Out The Notes For This Customer ")
        Me.cmdNotesPrint.UseVisualStyleBackColor = True
        '
        'cmdHistory
        '
        Me.cmdHistory.Location = New System.Drawing.Point(263, 175)
        Me.cmdHistory.Name = "cmdHistory"
        Me.cmdHistory.Size = New System.Drawing.Size(61, 23)
        Me.cmdHistory.TabIndex = 76
        Me.cmdHistory.Text = "&Cust Hist"
        Me.cmdHistory.UseVisualStyleBackColor = True
        '
        'fraEditOptions
        '
        Me.fraEditOptions.Controls.Add(Me.optEditType18)
        Me.fraEditOptions.Controls.Add(Me.optEditType17)
        Me.fraEditOptions.Controls.Add(Me.optEditType16)
        Me.fraEditOptions.Controls.Add(Me.optEditType15)
        Me.fraEditOptions.Controls.Add(Me.optEditType14)
        Me.fraEditOptions.Controls.Add(Me.optEditType13)
        Me.fraEditOptions.Controls.Add(Me.optEditType12)
        Me.fraEditOptions.Controls.Add(Me.optEditType11)
        Me.fraEditOptions.Controls.Add(Me.optEditType10)
        Me.fraEditOptions.Controls.Add(Me.optEditType9)
        Me.fraEditOptions.Controls.Add(Me.optEditType8)
        Me.fraEditOptions.Controls.Add(Me.optEditType7)
        Me.fraEditOptions.Location = New System.Drawing.Point(12, 329)
        Me.fraEditOptions.Name = "fraEditOptions"
        Me.fraEditOptions.Size = New System.Drawing.Size(686, 70)
        Me.fraEditOptions.TabIndex = 77
        Me.fraEditOptions.TabStop = False
        Me.fraEditOptions.Text = "mo"
        '
        'optEditType18
        '
        Me.optEditType18.AutoSize = True
        Me.optEditType18.Location = New System.Drawing.Point(403, 47)
        Me.optEditType18.Name = "optEditType18"
        Me.optEditType18.Size = New System.Drawing.Size(69, 17)
        Me.optEditType18.TabIndex = 11
        Me.optEditType18.TabStop = True
        Me.optEditType18.Text = "Credit IUI"
        Me.optEditType18.UseVisualStyleBackColor = True
        '
        'optEditType17
        '
        Me.optEditType17.AutoSize = True
        Me.optEditType17.Location = New System.Drawing.Point(213, 47)
        Me.optEditType17.Name = "optEditType17"
        Me.optEditType17.Size = New System.Drawing.Size(73, 17)
        Me.optEditType17.TabIndex = 10
        Me.optEditType17.TabStop = True
        Me.optEditType17.Text = "Credit Tax"
        Me.optEditType17.UseVisualStyleBackColor = True
        '
        'optEditType16
        '
        Me.optEditType16.AutoSize = True
        Me.optEditType16.Location = New System.Drawing.Point(17, 47)
        Me.optEditType16.Name = "optEditType16"
        Me.optEditType16.Size = New System.Drawing.Size(76, 17)
        Me.optEditType16.TabIndex = 9
        Me.optEditType16.TabStop = True
        Me.optEditType16.Text = "Credit Prin."
        Me.optEditType16.UseVisualStyleBackColor = True
        '
        'optEditType15
        '
        Me.optEditType15.AutoSize = True
        Me.optEditType15.Location = New System.Drawing.Point(501, 47)
        Me.optEditType15.Name = "optEditType15"
        Me.optEditType15.Size = New System.Drawing.Size(68, 17)
        Me.optEditType15.TabIndex = 8
        Me.optEditType15.TabStop = True
        Me.optEditType15.Text = "Debit Int."
        Me.optEditType15.UseVisualStyleBackColor = True
        '
        'optEditType14
        '
        Me.optEditType14.AutoSize = True
        Me.optEditType14.Location = New System.Drawing.Point(599, 19)
        Me.optEditType14.Name = "optEditType14"
        Me.optEditType14.Size = New System.Drawing.Size(71, 17)
        Me.optEditType14.TabIndex = 7
        Me.optEditType14.TabStop = True
        Me.optEditType14.Text = "Debit L/C"
        Me.optEditType14.UseVisualStyleBackColor = True
        '
        'optEditType13
        '
        Me.optEditType13.AutoSize = True
        Me.optEditType13.Location = New System.Drawing.Point(501, 19)
        Me.optEditType13.Name = "optEditType13"
        Me.optEditType13.Size = New System.Drawing.Size(74, 17)
        Me.optEditType13.TabIndex = 6
        Me.optEditType13.TabStop = True
        Me.optEditType13.Text = "Debit Prin."
        Me.optEditType13.UseVisualStyleBackColor = True
        '
        'optEditType12
        '
        Me.optEditType12.AutoSize = True
        Me.optEditType12.Location = New System.Drawing.Point(307, 47)
        Me.optEditType12.Name = "optEditType12"
        Me.optEditType12.Size = New System.Drawing.Size(77, 17)
        Me.optEditType12.TabIndex = 5
        Me.optEditType12.TabStop = True
        Me.optEditType12.Text = "Credit Prop"
        Me.optEditType12.UseVisualStyleBackColor = True
        '
        'optEditType11
        '
        Me.optEditType11.AutoSize = True
        Me.optEditType11.Location = New System.Drawing.Point(403, 19)
        Me.optEditType11.Name = "optEditType11"
        Me.optEditType11.Size = New System.Drawing.Size(74, 17)
        Me.optEditType11.TabIndex = 4
        Me.optEditType11.TabStop = True
        Me.optEditType11.Text = "Credit Acc"
        Me.optEditType11.UseVisualStyleBackColor = True
        '
        'optEditType10
        '
        Me.optEditType10.AutoSize = True
        Me.optEditType10.Location = New System.Drawing.Point(307, 19)
        Me.optEditType10.Name = "optEditType10"
        Me.optEditType10.Size = New System.Drawing.Size(72, 17)
        Me.optEditType10.TabIndex = 3
        Me.optEditType10.TabStop = True
        Me.optEditType10.Text = "Credit Life"
        Me.optEditType10.UseVisualStyleBackColor = True
        '
        'optEditType9
        '
        Me.optEditType9.AutoSize = True
        Me.optEditType9.Location = New System.Drawing.Point(213, 19)
        Me.optEditType9.Name = "optEditType9"
        Me.optEditType9.Size = New System.Drawing.Size(70, 17)
        Me.optEditType9.TabIndex = 2
        Me.optEditType9.TabStop = True
        Me.optEditType9.Text = "Credit Int."
        Me.optEditType9.UseVisualStyleBackColor = True
        '
        'optEditType8
        '
        Me.optEditType8.AutoSize = True
        Me.optEditType8.Location = New System.Drawing.Point(116, 19)
        Me.optEditType8.Name = "optEditType8"
        Me.optEditType8.Size = New System.Drawing.Size(73, 17)
        Me.optEditType8.TabIndex = 1
        Me.optEditType8.TabStop = True
        Me.optEditType8.Text = "Credit L/C"
        Me.optEditType8.UseVisualStyleBackColor = True
        '
        'optEditType7
        '
        Me.optEditType7.AutoSize = True
        Me.optEditType7.Location = New System.Drawing.Point(17, 19)
        Me.optEditType7.Name = "optEditType7"
        Me.optEditType7.Size = New System.Drawing.Size(75, 17)
        Me.optEditType7.TabIndex = 0
        Me.optEditType7.TabStop = True
        Me.optEditType7.Text = "Credit Doc"
        Me.optEditType7.UseVisualStyleBackColor = True
        '
        'fraPaymentOptions
        '
        Me.fraPaymentOptions.Controls.Add(Me.optPayType8)
        Me.fraPaymentOptions.Controls.Add(Me.optPayType9)
        Me.fraPaymentOptions.Controls.Add(Me.optPayType7)
        Me.fraPaymentOptions.Controls.Add(Me.optPayType6)
        Me.fraPaymentOptions.Controls.Add(Me.optPayType5)
        Me.fraPaymentOptions.Controls.Add(Me.optPayType4)
        Me.fraPaymentOptions.Controls.Add(Me.optPayType3)
        Me.fraPaymentOptions.Controls.Add(Me.optPayType2)
        Me.fraPaymentOptions.Controls.Add(Me.optPayType1)
        Me.fraPaymentOptions.Location = New System.Drawing.Point(12, 329)
        Me.fraPaymentOptions.Name = "fraPaymentOptions"
        Me.fraPaymentOptions.Size = New System.Drawing.Size(686, 70)
        Me.fraPaymentOptions.TabIndex = 78
        Me.fraPaymentOptions.TabStop = False
        Me.fraPaymentOptions.Text = "Method of Payment:"
        '
        'optPayType8
        '
        Me.optPayType8.AutoSize = True
        Me.optPayType8.Location = New System.Drawing.Point(216, 44)
        Me.optPayType8.Name = "optPayType8"
        Me.optPayType8.Size = New System.Drawing.Size(66, 17)
        Me.optPayType8.TabIndex = 9
        Me.optPayType8.TabStop = True
        Me.optPayType8.Text = "Gift Card"
        Me.optPayType8.UseVisualStyleBackColor = True
        '
        'optPayType7
        '
        Me.optPayType7.AutoSize = True
        Me.optPayType7.Location = New System.Drawing.Point(111, 44)
        Me.optPayType7.Name = "optPayType7"
        Me.optPayType7.Size = New System.Drawing.Size(75, 17)
        Me.optPayType7.TabIndex = 7
        Me.optPayType7.TabStop = True
        Me.optPayType7.Text = "Debit Card"
        Me.optPayType7.UseVisualStyleBackColor = True
        '
        'optPayType6
        '
        Me.optPayType6.AutoSize = True
        Me.optPayType6.Location = New System.Drawing.Point(576, 18)
        Me.optPayType6.Name = "optPayType6"
        Me.optPayType6.Size = New System.Drawing.Size(93, 17)
        Me.optPayType6.TabIndex = 6
        Me.optPayType6.TabStop = True
        Me.optPayType6.Text = "American Exp."
        Me.optPayType6.UseVisualStyleBackColor = True
        '
        'optPayType5
        '
        Me.optPayType5.AutoSize = True
        Me.optPayType5.Location = New System.Drawing.Point(437, 18)
        Me.optPayType5.Name = "optPayType5"
        Me.optPayType5.Size = New System.Drawing.Size(92, 17)
        Me.optPayType5.TabIndex = 5
        Me.optPayType5.TabStop = True
        Me.optPayType5.Text = "Discover Card"
        Me.optPayType5.UseVisualStyleBackColor = True
        '
        'optPayType4
        '
        Me.optPayType4.AutoSize = True
        Me.optPayType4.Location = New System.Drawing.Point(308, 18)
        Me.optPayType4.Name = "optPayType4"
        Me.optPayType4.Size = New System.Drawing.Size(82, 17)
        Me.optPayType4.TabIndex = 4
        Me.optPayType4.TabStop = True
        Me.optPayType4.Text = "Master Card"
        Me.optPayType4.UseVisualStyleBackColor = True
        '
        'optPayType3
        '
        Me.optPayType3.AutoSize = True
        Me.optPayType3.Location = New System.Drawing.Point(216, 18)
        Me.optPayType3.Name = "optPayType3"
        Me.optPayType3.Size = New System.Drawing.Size(45, 17)
        Me.optPayType3.TabIndex = 3
        Me.optPayType3.TabStop = True
        Me.optPayType3.Text = "Visa"
        Me.optPayType3.UseVisualStyleBackColor = True
        '
        'optPayType2
        '
        Me.optPayType2.AutoSize = True
        Me.optPayType2.Location = New System.Drawing.Point(113, 18)
        Me.optPayType2.Name = "optPayType2"
        Me.optPayType2.Size = New System.Drawing.Size(56, 17)
        Me.optPayType2.TabIndex = 2
        Me.optPayType2.TabStop = True
        Me.optPayType2.Text = "Check"
        Me.optPayType2.UseVisualStyleBackColor = True
        '
        'optPayType1
        '
        Me.optPayType1.AutoSize = True
        Me.optPayType1.Location = New System.Drawing.Point(17, 18)
        Me.optPayType1.Name = "optPayType1"
        Me.optPayType1.Size = New System.Drawing.Size(49, 17)
        Me.optPayType1.TabIndex = 1
        Me.optPayType1.TabStop = True
        Me.optPayType1.Text = "Cash"
        Me.optPayType1.UseVisualStyleBackColor = True
        '
        'Notes_Frame
        '
        Me.Notes_Frame.Controls.Add(Me.UGrSaleTotals)
        Me.Notes_Frame.Controls.Add(Me.lblNewNotes)
        Me.Notes_Frame.Controls.Add(Me.cmdNotesPrint)
        Me.Notes_Frame.Controls.Add(Me.cmdNotesCancel)
        Me.Notes_Frame.Controls.Add(Me.Notes_Save)
        Me.Notes_Frame.Controls.Add(Me.lblOldNotes)
        Me.Notes_Frame.Controls.Add(Me.Notes_New)
        Me.Notes_Frame.Controls.Add(Me.Notes_Text)
        Me.Notes_Frame.Location = New System.Drawing.Point(12, 482)
        Me.Notes_Frame.Name = "Notes_Frame"
        Me.Notes_Frame.Size = New System.Drawing.Size(686, 142)
        Me.Notes_Frame.TabIndex = 83
        Me.Notes_Frame.TabStop = False
        Me.Notes_Frame.Text = "Notes "
        Me.Notes_Frame.Visible = False
        '
        'UGrSaleTotals
        '
        Me.UGrSaleTotals.Activated = False
        Me.UGrSaleTotals.Col = 1
        Me.UGrSaleTotals.firstrow = 1
        Me.UGrSaleTotals.Loading = False
        Me.UGrSaleTotals.Location = New System.Drawing.Point(366, 22)
        Me.UGrSaleTotals.MaxCols = 2
        Me.UGrSaleTotals.MaxRows = 10
        Me.UGrSaleTotals.Name = "UGrSaleTotals"
        Me.UGrSaleTotals.Row = 0
        Me.UGrSaleTotals.Size = New System.Drawing.Size(230, 85)
        Me.UGrSaleTotals.TabIndex = 74
        Me.UGrSaleTotals.Visible = False
        '
        'lblNewNotes
        '
        Me.lblNewNotes.AutoSize = True
        Me.lblNewNotes.Location = New System.Drawing.Point(611, 115)
        Me.lblNewNotes.Name = "lblNewNotes"
        Me.lblNewNotes.Size = New System.Drawing.Size(60, 13)
        Me.lblNewNotes.TabIndex = 48
        Me.lblNewNotes.Text = "New Notes"
        '
        'lblOldNotes
        '
        Me.lblOldNotes.AutoSize = True
        Me.lblOldNotes.Location = New System.Drawing.Point(612, 17)
        Me.lblOldNotes.Name = "lblOldNotes"
        Me.lblOldNotes.Size = New System.Drawing.Size(54, 13)
        Me.lblOldNotes.TabIndex = 44
        Me.lblOldNotes.Text = "Old Notes"
        '
        'Notes_New
        '
        Me.Notes_New.Location = New System.Drawing.Point(9, 73)
        Me.Notes_New.Multiline = True
        Me.Notes_New.Name = "Notes_New"
        Me.Notes_New.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.Notes_New.Size = New System.Drawing.Size(600, 63)
        Me.Notes_New.TabIndex = 1
        '
        'Notes_Text
        '
        Me.Notes_Text.Location = New System.Drawing.Point(9, 19)
        Me.Notes_Text.Multiline = True
        Me.Notes_Text.Name = "Notes_Text"
        Me.Notes_Text.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.Notes_Text.Size = New System.Drawing.Size(600, 63)
        Me.Notes_Text.TabIndex = 0
        '
        'UGridIO1
        '
        Me.UGridIO1.Activated = False
        Me.UGridIO1.Col = 1
        Me.UGridIO1.firstrow = 1
        Me.UGridIO1.Loading = False
        Me.UGridIO1.Location = New System.Drawing.Point(12, 482)
        Me.UGridIO1.MaxCols = 2
        Me.UGridIO1.MaxRows = 10
        Me.UGridIO1.Name = "UGridIO1"
        Me.UGridIO1.Row = 0
        Me.UGridIO1.Size = New System.Drawing.Size(686, 142)
        Me.UGridIO1.TabIndex = 79
        '
        'fraCustomer
        '
        Me.fraCustomer.Controls.Add(Me.lblStatusCaption)
        Me.fraCustomer.Controls.Add(Me.lblAccountCaption)
        Me.fraCustomer.Controls.Add(Me.cboStatus)
        Me.fraCustomer.Controls.Add(Me.lblFirstName)
        Me.fraCustomer.Controls.Add(Me.lblAddAddress)
        Me.fraCustomer.Controls.Add(Me.lblCity)
        Me.fraCustomer.Controls.Add(Me.lblTele1Caption)
        Me.fraCustomer.Controls.Add(Me.cmdHistory)
        Me.fraCustomer.Controls.Add(Me.lblTele2Caption)
        Me.fraCustomer.Controls.Add(Me.lblCreditLimit)
        Me.fraCustomer.Controls.Add(Me.lblApprovalTerms)
        Me.fraCustomer.Controls.Add(Me.lblTele3)
        Me.fraCustomer.Controls.Add(Me.lblTele3Caption)
        Me.fraCustomer.Controls.Add(Me.lblSSN)
        Me.fraCustomer.Controls.Add(Me.lblAccount)
        Me.fraCustomer.Controls.Add(Me.lblZip)
        Me.fraCustomer.Controls.Add(Me.lblLastName)
        Me.fraCustomer.Controls.Add(Me.lblAddress)
        Me.fraCustomer.Controls.Add(Me.lblTele2)
        Me.fraCustomer.Controls.Add(Me.lblTele1)
        Me.fraCustomer.Location = New System.Drawing.Point(12, -1)
        Me.fraCustomer.Name = "fraCustomer"
        Me.fraCustomer.Size = New System.Drawing.Size(335, 204)
        Me.fraCustomer.TabIndex = 84
        Me.fraCustomer.TabStop = False
        '
        'lblStatusCaption
        '
        Me.lblStatusCaption.AutoSize = True
        Me.lblStatusCaption.Location = New System.Drawing.Point(160, 22)
        Me.lblStatusCaption.Name = "lblStatusCaption"
        Me.lblStatusCaption.Size = New System.Drawing.Size(40, 13)
        Me.lblStatusCaption.TabIndex = 1
        Me.lblStatusCaption.Tag = "Account: "
        Me.lblStatusCaption.Text = "Status:"
        '
        'lblAccountCaption
        '
        Me.lblAccountCaption.AutoSize = True
        Me.lblAccountCaption.Location = New System.Drawing.Point(9, 20)
        Me.lblAccountCaption.Name = "lblAccountCaption"
        Me.lblAccountCaption.Size = New System.Drawing.Size(53, 13)
        Me.lblAccountCaption.TabIndex = 0
        Me.lblAccountCaption.Tag = "Account: "
        Me.lblAccountCaption.Text = "Account: "
        '
        'fraButtons
        '
        Me.fraButtons.Controls.Add(Me.lbl0030)
        Me.fraButtons.Controls.Add(Me.lbl3160)
        Me.fraButtons.Controls.Add(Me.lbl6190)
        Me.fraButtons.Controls.Add(Me.lblOver91)
        Me.fraButtons.Controls.Add(Me.lblLate0)
        Me.fraButtons.Controls.Add(Me.lblLate31)
        Me.fraButtons.Controls.Add(Me.lblLate61)
        Me.fraButtons.Controls.Add(Me.lblLate91)
        Me.fraButtons.Controls.Add(Me.lblArrearages)
        Me.fraButtons.Controls.Add(Me.cmdSaleTotals)
        Me.fraButtons.Controls.Add(Me.lblTotDue)
        Me.fraButtons.Controls.Add(Me.cmdPrintCard)
        Me.fraButtons.Controls.Add(Me.cmdCreditApp)
        Me.fraButtons.Controls.Add(Me.Notes_Open)
        Me.fraButtons.Controls.Add(Me.cmdApply)
        Me.fraButtons.Controls.Add(Me.cmdDetail)
        Me.fraButtons.Controls.Add(Me.cmdReceipt)
        Me.fraButtons.Controls.Add(Me.cmdCancel)
        Me.fraButtons.Controls.Add(Me.cmdPayoff)
        Me.fraButtons.Location = New System.Drawing.Point(356, 202)
        Me.fraButtons.Name = "fraButtons"
        Me.fraButtons.Size = New System.Drawing.Size(341, 124)
        Me.fraButtons.TabIndex = 85
        Me.fraButtons.TabStop = False
        '
        'fraPayoffInfo
        '
        Me.fraPayoffInfo.Controls.Add(Me.txtPayoffInfo)
        Me.fraPayoffInfo.Location = New System.Drawing.Point(355, 224)
        Me.fraPayoffInfo.Name = "fraPayoffInfo"
        Me.fraPayoffInfo.Size = New System.Drawing.Size(341, 102)
        Me.fraPayoffInfo.TabIndex = 86
        Me.fraPayoffInfo.TabStop = False
        Me.fraPayoffInfo.Text = "Payoff Information"
        Me.fraPayoffInfo.Visible = False
        '
        'txtPayoffInfo
        '
        Me.txtPayoffInfo.BackColor = System.Drawing.SystemColors.ButtonFace
        Me.txtPayoffInfo.Location = New System.Drawing.Point(10, 19)
        Me.txtPayoffInfo.Multiline = True
        Me.txtPayoffInfo.Name = "txtPayoffInfo"
        Me.txtPayoffInfo.ScrollBars = System.Windows.Forms.ScrollBars.Both
        Me.txtPayoffInfo.Size = New System.Drawing.Size(325, 85)
        Me.txtPayoffInfo.TabIndex = 0
        '
        'fraArrearControl
        '
        Me.fraArrearControl.Controls.Add(Me.lblArrearControlDisplay)
        Me.fraArrearControl.Controls.Add(Me.dtpArrearControlDate)
        Me.fraArrearControl.Controls.Add(Me.lblArrearControlDate)
        Me.fraArrearControl.Controls.Add(Me.lblArrearControlGrace)
        Me.fraArrearControl.Controls.Add(Me.chkArrearControlGrace)
        Me.fraArrearControl.Location = New System.Drawing.Point(356, 202)
        Me.fraArrearControl.Name = "fraArrearControl"
        Me.fraArrearControl.Size = New System.Drawing.Size(341, 115)
        Me.fraArrearControl.TabIndex = 87
        Me.fraArrearControl.TabStop = False
        Me.fraArrearControl.Text = "Arrearages Control"
        Me.fraArrearControl.Visible = False
        '
        'lblArrearControlDisplay
        '
        Me.lblArrearControlDisplay.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblArrearControlDisplay.Location = New System.Drawing.Point(8, 45)
        Me.lblArrearControlDisplay.Name = "lblArrearControlDisplay"
        Me.lblArrearControlDisplay.Size = New System.Drawing.Size(316, 62)
        Me.lblArrearControlDisplay.TabIndex = 4
        '
        'dtpArrearControlDate
        '
        Me.dtpArrearControlDate.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtpArrearControlDate.Location = New System.Drawing.Point(248, 20)
        Me.dtpArrearControlDate.Name = "dtpArrearControlDate"
        Me.dtpArrearControlDate.Size = New System.Drawing.Size(76, 20)
        Me.dtpArrearControlDate.TabIndex = 3
        '
        'lblArrearControlDate
        '
        Me.lblArrearControlDate.AutoSize = True
        Me.lblArrearControlDate.Location = New System.Drawing.Point(220, 23)
        Me.lblArrearControlDate.Name = "lblArrearControlDate"
        Me.lblArrearControlDate.Size = New System.Drawing.Size(22, 13)
        Me.lblArrearControlDate.TabIndex = 2
        Me.lblArrearControlDate.Text = "As:"
        '
        'lblArrearControlGrace
        '
        Me.lblArrearControlGrace.AutoSize = True
        Me.lblArrearControlGrace.BackColor = System.Drawing.Color.FromArgb(CType(CType(128, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.lblArrearControlGrace.Location = New System.Drawing.Point(106, 24)
        Me.lblArrearControlGrace.Name = "lblArrearControlGrace"
        Me.lblArrearControlGrace.Size = New System.Drawing.Size(21, 13)
        Me.lblArrearControlGrace.TabIndex = 1
        Me.lblArrearControlGrace.Text = "##"
        '
        'chkArrearControlGrace
        '
        Me.chkArrearControlGrace.AutoSize = True
        Me.chkArrearControlGrace.CheckAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.chkArrearControlGrace.Location = New System.Drawing.Point(6, 23)
        Me.chkArrearControlGrace.Name = "chkArrearControlGrace"
        Me.chkArrearControlGrace.Size = New System.Drawing.Size(83, 17)
        Me.chkArrearControlGrace.TabIndex = 0
        Me.chkArrearControlGrace.Text = "Use Grace?"
        Me.chkArrearControlGrace.UseVisualStyleBackColor = True
        '
        'fraPrintType
        '
        Me.fraPrintType.Controls.Add(Me.opt30252)
        Me.fraPrintType.Controls.Add(Me.opt30323)
        Me.fraPrintType.Controls.Add(Me.lblPrintType)
        Me.fraPrintType.Location = New System.Drawing.Point(263, 11)
        Me.fraPrintType.Name = "fraPrintType"
        Me.fraPrintType.Size = New System.Drawing.Size(72, 57)
        Me.fraPrintType.TabIndex = 89
        Me.fraPrintType.TabStop = False
        '
        'lblPrintType
        '
        Me.lblPrintType.AutoSize = True
        Me.lblPrintType.Location = New System.Drawing.Point(6, 8)
        Me.lblPrintType.Name = "lblPrintType"
        Me.lblPrintType.Size = New System.Drawing.Size(60, 13)
        Me.lblPrintType.TabIndex = 87
        Me.lblPrintType.Text = "Label Type"
        '
        'opt30323
        '
        Me.opt30323.AutoSize = True
        Me.opt30323.Checked = True
        Me.opt30323.Location = New System.Drawing.Point(9, 20)
        Me.opt30323.Name = "opt30323"
        Me.opt30323.Size = New System.Drawing.Size(55, 17)
        Me.opt30323.TabIndex = 89
        Me.opt30323.TabStop = True
        Me.opt30323.Text = "30323"
        Me.ToolTip1.SetToolTip(Me.opt30323, "Click this for the wider DYMO Shipping labels.")
        Me.opt30323.UseVisualStyleBackColor = True
        '
        'opt30252
        '
        Me.opt30252.AutoSize = True
        Me.opt30252.Location = New System.Drawing.Point(9, 38)
        Me.opt30252.Name = "opt30252"
        Me.opt30252.Size = New System.Drawing.Size(55, 17)
        Me.opt30252.TabIndex = 90
        Me.opt30252.Text = "30252"
        Me.ToolTip1.SetToolTip(Me.opt30252, "Select this option for narrow DYMO address labels.")
        Me.opt30252.UseVisualStyleBackColor = True
        '
        'ArCard
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(702, 640)
        Me.Controls.Add(Me.fraButtons)
        Me.Controls.Add(Me.fraPayoffInfo)
        Me.Controls.Add(Me.fraCustomer)
        Me.Controls.Add(Me.fraPaymentOptions)
        Me.Controls.Add(Me.fraEditOptions)
        Me.Controls.Add(Me.fraBalance)
        Me.Controls.Add(Me.fraTerms)
        Me.Controls.Add(Me.fraNav)
        Me.Controls.Add(Me.fraPrint)
        Me.Controls.Add(Me.cmdMakeSameAsCash)
        Me.Controls.Add(Me.fraArrearControl)
        Me.Controls.Add(Me.UGridIO1)
        Me.Controls.Add(Me.Notes_Frame)
        Me.Name = "ArCard"
        Me.fraPrint.ResumeLayout(False)
        Me.fraNav.ResumeLayout(False)
        Me.fraNav.PerformLayout()
        Me.fraTerms.ResumeLayout(False)
        Me.fraTerms.PerformLayout()
        Me.fraBalance.ResumeLayout(False)
        Me.fraBalance.PerformLayout()
        Me.fraEditOptions.ResumeLayout(False)
        Me.fraEditOptions.PerformLayout()
        Me.fraPaymentOptions.ResumeLayout(False)
        Me.fraPaymentOptions.PerformLayout()
        Me.Notes_Frame.ResumeLayout(False)
        Me.Notes_Frame.PerformLayout()
        Me.fraCustomer.ResumeLayout(False)
        Me.fraCustomer.PerformLayout()
        Me.fraButtons.ResumeLayout(False)
        Me.fraButtons.PerformLayout()
        Me.fraPayoffInfo.ResumeLayout(False)
        Me.fraPayoffInfo.PerformLayout()
        Me.fraArrearControl.ResumeLayout(False)
        Me.fraArrearControl.PerformLayout()
        Me.fraPrintType.ResumeLayout(False)
        Me.fraPrintType.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents lblTotalPayoff As Label
    Friend WithEvents cmdApply As Button
    Friend WithEvents cmdEdit As Button
    Friend WithEvents cmdExport As Button
    Friend WithEvents cmdFields As Button
    Friend WithEvents cmdMakeSameAsCash As Button
    Friend WithEvents cmdMoveFirst As Button
    Friend WithEvents cmdMoveLast As Button
    Friend WithEvents cmdMoveNext As Button
    Friend WithEvents cmdMovePrevious As Button
    Friend WithEvents cmdPayoff As Button
    Friend WithEvents cmdReceipt As Button
    Friend WithEvents cmdReprintContract As Button
    Friend WithEvents cmdPrint As Button
    Friend WithEvents cmdCancel As Button
    Friend WithEvents txtPaymentHistory As TextBox
    Friend WithEvents lblBalance As Label
    Friend WithEvents txtDelivery As TextBox
    Friend WithEvents txtMonths As TextBox
    Friend WithEvents txtLastPay As TextBox
    Friend WithEvents txtFinanced As TextBox
    Friend WithEvents txtMonthlyPayment As TextBox
    Friend WithEvents txtPayMemo As TextBox
    Friend WithEvents txtRate As TextBox
    Friend WithEvents txtSameAsCash As TextBox
    Friend WithEvents lblLateCharge As Label
    Friend WithEvents txtFirstPay As TextBox
    Friend WithEvents txtPaidBy As TextBox
    Friend WithEvents txtPayPeriod As TextBox
    Friend WithEvents txtNextDue As TextBox
    Friend WithEvents lblTotDue As Label
    Friend WithEvents lblArrearages As Label
    Friend WithEvents lblLate0 As Label
    Friend WithEvents lblLate31 As Label
    Friend WithEvents lblLate61 As Label
    Friend WithEvents lblLate91 As Label
    Friend WithEvents lbl0030 As Label
    Friend WithEvents lbl3160 As Label
    Friend WithEvents lbl6190 As Label
    Friend WithEvents lblOver91 As Label
    Friend WithEvents filFile As Compatibility.VB6.FileListBox
    Friend WithEvents lblAccount As Label
    Friend WithEvents cboStatus As ComboBox
    Friend WithEvents lblFirstName As Label
    Friend WithEvents lblLastName As Label
    Friend WithEvents lblAddress As Label
    Friend WithEvents lblCity As Label
    Friend WithEvents lblZip As Label
    Friend WithEvents lblTele1 As Label
    Friend WithEvents lblTele2 As Label
    Friend WithEvents lblSSN As Label
    Friend WithEvents txtLateChargeAmount As TextBox
    Friend WithEvents lblCreditLimit As Label
    Friend WithEvents lblApprovalTerms As Label
    Friend WithEvents lblAddAddress As Label
    Friend WithEvents lblTele3 As Label
    Friend WithEvents lblTele1Caption As Label
    Friend WithEvents lblTele2Caption As Label
    Friend WithEvents lblTele3Caption As Label
    Friend WithEvents fraPrint As GroupBox
    Friend WithEvents fraNav As GroupBox
    Friend WithEvents fraTerms As GroupBox
    Friend WithEvents fraBalance As GroupBox
    Friend WithEvents cmdCreditApp As Button
    Friend WithEvents cmdDetail As Button
    Friend WithEvents Notes_Open As Button
    Friend WithEvents cmdPrintCard As Button
    Friend WithEvents lblAPR As Label
    Friend WithEvents chkSendAllMail As CheckBox
    Friend WithEvents lblPaymentHistory As Label
    Friend WithEvents ToolTip1 As ToolTip
    Friend WithEvents cmdSaleTotals As Button
    Friend WithEvents UGrSaleTotals As UGridIO
    Friend WithEvents lblPayMemo As Label
    Friend WithEvents cmdHistory As Button
    Friend WithEvents fraEditOptions As GroupBox
    Friend WithEvents optEditType18 As RadioButton
    Friend WithEvents optEditType17 As RadioButton
    Friend WithEvents optEditType16 As RadioButton
    Friend WithEvents optEditType15 As RadioButton
    Friend WithEvents optEditType14 As RadioButton
    Friend WithEvents optEditType13 As RadioButton
    Friend WithEvents optEditType12 As RadioButton
    Friend WithEvents optEditType11 As RadioButton
    Friend WithEvents optEditType10 As RadioButton
    Friend WithEvents optEditType9 As RadioButton
    Friend WithEvents optEditType8 As RadioButton
    Friend WithEvents optEditType7 As RadioButton
    Friend WithEvents fraPaymentOptions As GroupBox
    Friend WithEvents UGridIO1 As UGridIO
    Friend WithEvents optPayType8 As RadioButton
    Friend WithEvents optPayType9 As RadioButton
    Friend WithEvents optPayType7 As RadioButton
    Friend WithEvents optPayType6 As RadioButton
    Friend WithEvents optPayType5 As RadioButton
    Friend WithEvents optPayType4 As RadioButton
    Friend WithEvents optPayType3 As RadioButton
    Friend WithEvents optPayType2 As RadioButton
    Friend WithEvents optPayType1 As RadioButton
    Friend WithEvents DDate As DateTimePicker
    Friend WithEvents Notes_Frame As GroupBox
    Friend WithEvents fraCustomer As GroupBox
    Friend WithEvents lblStatusCaption As Label
    Friend WithEvents lblAccountCaption As Label
    Friend WithEvents cmdReprintCoupons As Button
    Friend WithEvents lblNextDue As Label
    Friend WithEvents lblLastPay As Label
    Friend WithEvents lblSameAsCash As Label
    Friend WithEvents lbl1stPay As Label
    Friend WithEvents lblDelivery As Label
    Friend WithEvents lblPayBy As Label
    Friend WithEvents lblLateCh As Label
    Friend WithEvents lblPayment As Label
    Friend WithEvents lblRate As Label
    Friend WithEvents lblMonths As Label
    Friend WithEvents lblFinanced As Label
    Friend WithEvents lblLateChargeCaption As Label
    Friend WithEvents lblBalanceCaption As Label
    Friend WithEvents lblTotalPayOffCaption As Label
    Friend WithEvents chkPayLateFee As CheckBox
    Friend WithEvents txtPayment As TextBox
    Friend WithEvents lblPayDate As Label
    Friend WithEvents txtLateCharge As TextBox
    Friend WithEvents fraButtons As GroupBox
    Friend WithEvents fraPayoffInfo As GroupBox
    Friend WithEvents txtPayoffInfo As TextBox
    Friend WithEvents fraArrearControl As GroupBox
    Friend WithEvents lblArrearControlDisplay As Label
    Friend WithEvents dtpArrearControlDate As DateTimePicker
    Friend WithEvents lblArrearControlDate As Label
    Friend WithEvents lblArrearControlGrace As Label
    Friend WithEvents chkArrearControlGrace As CheckBox
    Friend WithEvents cmdPrintLabel As Button
    Friend WithEvents rtfFile As RichTextBox
    Friend WithEvents lblNewNotes As Label
    Friend WithEvents cmdNotesPrint As Button
    Friend WithEvents cmdNotesCancel As Button
    Friend WithEvents Notes_Save As Button
    Friend WithEvents lblOldNotes As Label
    Friend WithEvents Notes_New As TextBox
    Friend WithEvents Notes_Text As TextBox
    Friend WithEvents fraPrintType As GroupBox
    Friend WithEvents opt30252 As RadioButton
    Friend WithEvents opt30323 As RadioButton
    Friend WithEvents lblPrintType As Label
End Class
