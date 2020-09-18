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
        Me.fraPaymentOptions = New System.Windows.Forms.GroupBox()
        Me.fraEditOptions = New System.Windows.Forms.GroupBox()
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
        Me.rtfFile = New System.Windows.Forms.RichTextBox()
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
        Me.fraNav = New System.Windows.Forms.GroupBox()
        Me.fraTerms = New System.Windows.Forms.GroupBox()
        Me.fraBalance = New System.Windows.Forms.GroupBox()
        Me.fraPrintType = New System.Windows.Forms.GroupBox()
        Me.cmdCreditApp = New System.Windows.Forms.Button()
        Me.cmdDetail = New System.Windows.Forms.Button()
        Me.Notes_Open = New System.Windows.Forms.Button()
        Me.cmdPrintCard = New System.Windows.Forms.Button()
        Me.lblAPR = New System.Windows.Forms.Label()
        Me.chkSendAllMail = New System.Windows.Forms.CheckBox()
        Me.lblPaymentHistory = New System.Windows.Forms.Label()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.cmdSaleTotals = New System.Windows.Forms.Button()
        Me.UGrSaleTotals = New WinCDS.UGridIO()
        Me.SuspendLayout()
        '
        'lblTotalPayoff
        '
        Me.lblTotalPayoff.AutoSize = True
        Me.lblTotalPayoff.Location = New System.Drawing.Point(0, 0)
        Me.lblTotalPayoff.Name = "lblTotalPayoff"
        Me.lblTotalPayoff.Size = New System.Drawing.Size(71, 13)
        Me.lblTotalPayoff.TabIndex = 0
        Me.lblTotalPayoff.Text = "lblTotalPayoff"
        '
        'cmdApply
        '
        Me.cmdApply.Location = New System.Drawing.Point(-4, 42)
        Me.cmdApply.Name = "cmdApply"
        Me.cmdApply.Size = New System.Drawing.Size(75, 23)
        Me.cmdApply.TabIndex = 1
        Me.cmdApply.Text = "&Apply"
        Me.cmdApply.UseVisualStyleBackColor = True
        '
        'cmdEdit
        '
        Me.cmdEdit.Location = New System.Drawing.Point(-4, 89)
        Me.cmdEdit.Name = "cmdEdit"
        Me.cmdEdit.Size = New System.Drawing.Size(75, 23)
        Me.cmdEdit.TabIndex = 2
        Me.cmdEdit.Text = "&Edit"
        Me.cmdEdit.UseVisualStyleBackColor = True
        '
        'cmdExport
        '
        Me.cmdExport.Location = New System.Drawing.Point(3, 151)
        Me.cmdExport.Name = "cmdExport"
        Me.cmdExport.Size = New System.Drawing.Size(75, 23)
        Me.cmdExport.TabIndex = 3
        Me.cmdExport.Text = "E&xport"
        Me.cmdExport.UseVisualStyleBackColor = True
        '
        'cmdFields
        '
        Me.cmdFields.Location = New System.Drawing.Point(3, 191)
        Me.cmdFields.Name = "cmdFields"
        Me.cmdFields.Size = New System.Drawing.Size(75, 23)
        Me.cmdFields.TabIndex = 4
        Me.cmdFields.Text = "&Fields"
        Me.cmdFields.UseVisualStyleBackColor = True
        '
        'cmdMakeSameAsCash
        '
        Me.cmdMakeSameAsCash.Location = New System.Drawing.Point(3, 243)
        Me.cmdMakeSameAsCash.Name = "cmdMakeSameAsCash"
        Me.cmdMakeSameAsCash.Size = New System.Drawing.Size(75, 23)
        Me.cmdMakeSameAsCash.TabIndex = 5
        Me.cmdMakeSameAsCash.Text = "SaC"
        Me.cmdMakeSameAsCash.UseVisualStyleBackColor = True
        '
        'cmdMoveFirst
        '
        Me.cmdMoveFirst.Location = New System.Drawing.Point(3, 283)
        Me.cmdMoveFirst.Name = "cmdMoveFirst"
        Me.cmdMoveFirst.Size = New System.Drawing.Size(75, 23)
        Me.cmdMoveFirst.TabIndex = 6
        Me.cmdMoveFirst.Text = "<<"
        Me.cmdMoveFirst.UseVisualStyleBackColor = True
        '
        'cmdMoveLast
        '
        Me.cmdMoveLast.Location = New System.Drawing.Point(3, 312)
        Me.cmdMoveLast.Name = "cmdMoveLast"
        Me.cmdMoveLast.Size = New System.Drawing.Size(75, 23)
        Me.cmdMoveLast.TabIndex = 7
        Me.cmdMoveLast.Text = ">>"
        Me.cmdMoveLast.UseVisualStyleBackColor = True
        '
        'cmdMoveNext
        '
        Me.cmdMoveNext.Location = New System.Drawing.Point(12, 341)
        Me.cmdMoveNext.Name = "cmdMoveNext"
        Me.cmdMoveNext.Size = New System.Drawing.Size(75, 23)
        Me.cmdMoveNext.TabIndex = 8
        Me.cmdMoveNext.Text = ">"
        Me.cmdMoveNext.UseVisualStyleBackColor = True
        '
        'cmdMovePrevious
        '
        Me.cmdMovePrevious.Location = New System.Drawing.Point(12, 370)
        Me.cmdMovePrevious.Name = "cmdMovePrevious"
        Me.cmdMovePrevious.Size = New System.Drawing.Size(75, 23)
        Me.cmdMovePrevious.TabIndex = 9
        Me.cmdMovePrevious.Text = "<"
        Me.cmdMovePrevious.UseVisualStyleBackColor = True
        '
        'cmdPayoff
        '
        Me.cmdPayoff.Location = New System.Drawing.Point(12, 399)
        Me.cmdPayoff.Name = "cmdPayoff"
        Me.cmdPayoff.Size = New System.Drawing.Size(75, 23)
        Me.cmdPayoff.TabIndex = 10
        Me.cmdPayoff.Text = "&Early Payoff"
        Me.cmdPayoff.UseVisualStyleBackColor = True
        '
        'cmdReceipt
        '
        Me.cmdReceipt.Location = New System.Drawing.Point(116, 33)
        Me.cmdReceipt.Name = "cmdReceipt"
        Me.cmdReceipt.Size = New System.Drawing.Size(75, 23)
        Me.cmdReceipt.TabIndex = 11
        Me.cmdReceipt.Text = "&Receipt"
        Me.cmdReceipt.UseVisualStyleBackColor = True
        '
        'cmdReprintContract
        '
        Me.cmdReprintContract.Location = New System.Drawing.Point(116, 62)
        Me.cmdReprintContract.Name = "cmdReprintContract"
        Me.cmdReprintContract.Size = New System.Drawing.Size(75, 23)
        Me.cmdReprintContract.TabIndex = 12
        Me.cmdReprintContract.Text = "Reprint Contract"
        Me.cmdReprintContract.UseVisualStyleBackColor = True
        '
        'cmdPrint
        '
        Me.cmdPrint.Location = New System.Drawing.Point(116, 100)
        Me.cmdPrint.Name = "cmdPrint"
        Me.cmdPrint.Size = New System.Drawing.Size(75, 23)
        Me.cmdPrint.TabIndex = 13
        Me.cmdPrint.Text = "&Print"
        Me.cmdPrint.UseVisualStyleBackColor = True
        '
        'cmdCancel
        '
        Me.cmdCancel.Location = New System.Drawing.Point(116, 129)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.Size = New System.Drawing.Size(75, 23)
        Me.cmdCancel.TabIndex = 14
        Me.cmdCancel.Text = "&Cancel"
        Me.cmdCancel.UseVisualStyleBackColor = True
        '
        'fraPaymentOptions
        '
        Me.fraPaymentOptions.Location = New System.Drawing.Point(116, 191)
        Me.fraPaymentOptions.Name = "fraPaymentOptions"
        Me.fraPaymentOptions.Size = New System.Drawing.Size(200, 100)
        Me.fraPaymentOptions.TabIndex = 15
        Me.fraPaymentOptions.TabStop = False
        Me.fraPaymentOptions.Text = "Method of Payment:"
        '
        'fraEditOptions
        '
        Me.fraEditOptions.Location = New System.Drawing.Point(116, 297)
        Me.fraEditOptions.Name = "fraEditOptions"
        Me.fraEditOptions.Size = New System.Drawing.Size(200, 100)
        Me.fraEditOptions.TabIndex = 16
        Me.fraEditOptions.TabStop = False
        Me.fraEditOptions.Text = "mo"
        '
        'txtPaymentHistory
        '
        Me.txtPaymentHistory.Location = New System.Drawing.Point(126, 418)
        Me.txtPaymentHistory.Name = "txtPaymentHistory"
        Me.txtPaymentHistory.Size = New System.Drawing.Size(100, 20)
        Me.txtPaymentHistory.TabIndex = 0
        '
        'lblBalance
        '
        Me.lblBalance.AutoSize = True
        Me.lblBalance.Location = New System.Drawing.Point(360, 33)
        Me.lblBalance.Name = "lblBalance"
        Me.lblBalance.Size = New System.Drawing.Size(56, 13)
        Me.lblBalance.TabIndex = 17
        Me.lblBalance.Text = "lblBalance"
        '
        'txtDelivery
        '
        Me.txtDelivery.Location = New System.Drawing.Point(363, 49)
        Me.txtDelivery.Name = "txtDelivery"
        Me.txtDelivery.Size = New System.Drawing.Size(100, 20)
        Me.txtDelivery.TabIndex = 18
        '
        'txtMonths
        '
        Me.txtMonths.Location = New System.Drawing.Point(363, 75)
        Me.txtMonths.Name = "txtMonths"
        Me.txtMonths.Size = New System.Drawing.Size(100, 20)
        Me.txtMonths.TabIndex = 19
        '
        'txtLastPay
        '
        Me.txtLastPay.Location = New System.Drawing.Point(363, 103)
        Me.txtLastPay.Name = "txtLastPay"
        Me.txtLastPay.Size = New System.Drawing.Size(100, 20)
        Me.txtLastPay.TabIndex = 20
        '
        'txtFinanced
        '
        Me.txtFinanced.Location = New System.Drawing.Point(353, 129)
        Me.txtFinanced.Name = "txtFinanced"
        Me.txtFinanced.Size = New System.Drawing.Size(100, 20)
        Me.txtFinanced.TabIndex = 21
        '
        'txtMonthlyPayment
        '
        Me.txtMonthlyPayment.Location = New System.Drawing.Point(363, 166)
        Me.txtMonthlyPayment.Name = "txtMonthlyPayment"
        Me.txtMonthlyPayment.Size = New System.Drawing.Size(100, 20)
        Me.txtMonthlyPayment.TabIndex = 22
        '
        'txtPayMemo
        '
        Me.txtPayMemo.Location = New System.Drawing.Point(363, 192)
        Me.txtPayMemo.Name = "txtPayMemo"
        Me.txtPayMemo.Size = New System.Drawing.Size(100, 20)
        Me.txtPayMemo.TabIndex = 23
        '
        'txtRate
        '
        Me.txtRate.Location = New System.Drawing.Point(353, 218)
        Me.txtRate.Name = "txtRate"
        Me.txtRate.Size = New System.Drawing.Size(100, 20)
        Me.txtRate.TabIndex = 24
        '
        'txtSameAsCash
        '
        Me.txtSameAsCash.Location = New System.Drawing.Point(353, 246)
        Me.txtSameAsCash.Name = "txtSameAsCash"
        Me.txtSameAsCash.Size = New System.Drawing.Size(100, 20)
        Me.txtSameAsCash.TabIndex = 25
        '
        'lblLateCharge
        '
        Me.lblLateCharge.AutoSize = True
        Me.lblLateCharge.Location = New System.Drawing.Point(360, 278)
        Me.lblLateCharge.Name = "lblLateCharge"
        Me.lblLateCharge.Size = New System.Drawing.Size(72, 13)
        Me.lblLateCharge.TabIndex = 26
        Me.lblLateCharge.Text = "lblLateCharge"
        '
        'txtFirstPay
        '
        Me.txtFirstPay.Location = New System.Drawing.Point(363, 297)
        Me.txtFirstPay.Name = "txtFirstPay"
        Me.txtFirstPay.Size = New System.Drawing.Size(100, 20)
        Me.txtFirstPay.TabIndex = 27
        '
        'txtPaidBy
        '
        Me.txtPaidBy.Location = New System.Drawing.Point(363, 323)
        Me.txtPaidBy.Name = "txtPaidBy"
        Me.txtPaidBy.Size = New System.Drawing.Size(100, 20)
        Me.txtPaidBy.TabIndex = 28
        '
        'txtPayPeriod
        '
        Me.txtPayPeriod.Location = New System.Drawing.Point(363, 349)
        Me.txtPayPeriod.Name = "txtPayPeriod"
        Me.txtPayPeriod.Size = New System.Drawing.Size(100, 20)
        Me.txtPayPeriod.TabIndex = 29
        '
        'txtNextDue
        '
        Me.txtNextDue.Location = New System.Drawing.Point(332, 377)
        Me.txtNextDue.Name = "txtNextDue"
        Me.txtNextDue.Size = New System.Drawing.Size(100, 20)
        Me.txtNextDue.TabIndex = 30
        '
        'lblTotDue
        '
        Me.lblTotDue.AutoSize = True
        Me.lblTotDue.Location = New System.Drawing.Point(340, 400)
        Me.lblTotDue.Name = "lblTotDue"
        Me.lblTotDue.Size = New System.Drawing.Size(53, 13)
        Me.lblTotDue.TabIndex = 31
        Me.lblTotDue.Text = "lblTotDue"
        '
        'lblArrearages
        '
        Me.lblArrearages.AutoSize = True
        Me.lblArrearages.Location = New System.Drawing.Point(340, 421)
        Me.lblArrearages.Name = "lblArrearages"
        Me.lblArrearages.Size = New System.Drawing.Size(68, 13)
        Me.lblArrearages.TabIndex = 32
        Me.lblArrearages.Text = "lblArrearages"
        '
        'lblLate0
        '
        Me.lblLate0.AutoSize = True
        Me.lblLate0.Location = New System.Drawing.Point(496, 47)
        Me.lblLate0.Name = "lblLate0"
        Me.lblLate0.Size = New System.Drawing.Size(44, 13)
        Me.lblLate0.TabIndex = 33
        Me.lblLate0.Text = "lblLate0"
        '
        'lblLate31
        '
        Me.lblLate31.AutoSize = True
        Me.lblLate31.Location = New System.Drawing.Point(496, 72)
        Me.lblLate31.Name = "lblLate31"
        Me.lblLate31.Size = New System.Drawing.Size(50, 13)
        Me.lblLate31.TabIndex = 34
        Me.lblLate31.Text = "lblLate31"
        '
        'lblLate61
        '
        Me.lblLate61.AutoSize = True
        Me.lblLate61.Location = New System.Drawing.Point(496, 89)
        Me.lblLate61.Name = "lblLate61"
        Me.lblLate61.Size = New System.Drawing.Size(50, 13)
        Me.lblLate61.TabIndex = 35
        Me.lblLate61.Text = "lblLate61"
        '
        'lblLate91
        '
        Me.lblLate91.AutoSize = True
        Me.lblLate91.Location = New System.Drawing.Point(496, 110)
        Me.lblLate91.Name = "lblLate91"
        Me.lblLate91.Size = New System.Drawing.Size(50, 13)
        Me.lblLate91.TabIndex = 36
        Me.lblLate91.Text = "lblLate91"
        '
        'lbl0030
        '
        Me.lbl0030.AutoSize = True
        Me.lbl0030.Location = New System.Drawing.Point(496, 134)
        Me.lbl0030.Name = "lbl0030"
        Me.lbl0030.Size = New System.Drawing.Size(41, 13)
        Me.lbl0030.TabIndex = 37
        Me.lbl0030.Text = "lbl0030"
        '
        'lbl3160
        '
        Me.lbl3160.AutoSize = True
        Me.lbl3160.Location = New System.Drawing.Point(496, 151)
        Me.lbl3160.Name = "lbl3160"
        Me.lbl3160.Size = New System.Drawing.Size(41, 13)
        Me.lbl3160.TabIndex = 38
        Me.lbl3160.Text = "lbl3160"
        '
        'lbl6190
        '
        Me.lbl6190.AutoSize = True
        Me.lbl6190.Location = New System.Drawing.Point(496, 169)
        Me.lbl6190.Name = "lbl6190"
        Me.lbl6190.Size = New System.Drawing.Size(41, 13)
        Me.lbl6190.TabIndex = 39
        Me.lbl6190.Text = "lbl6190"
        '
        'lblOver91
        '
        Me.lblOver91.AutoSize = True
        Me.lblOver91.Location = New System.Drawing.Point(496, 192)
        Me.lblOver91.Name = "lblOver91"
        Me.lblOver91.Size = New System.Drawing.Size(52, 13)
        Me.lblOver91.TabIndex = 40
        Me.lblOver91.Text = "lblOver91"
        '
        'rtfFile
        '
        Me.rtfFile.Location = New System.Drawing.Point(499, 218)
        Me.rtfFile.Name = "rtfFile"
        Me.rtfFile.Size = New System.Drawing.Size(64, 38)
        Me.rtfFile.TabIndex = 41
        Me.rtfFile.Text = ""
        '
        'filFile
        '
        Me.filFile.FormattingEnabled = True
        Me.filFile.Location = New System.Drawing.Point(499, 263)
        Me.filFile.Name = "filFile"
        Me.filFile.Pattern = "*.*"
        Me.filFile.Size = New System.Drawing.Size(64, 43)
        Me.filFile.TabIndex = 42
        '
        'lblAccount
        '
        Me.lblAccount.AutoSize = True
        Me.lblAccount.Location = New System.Drawing.Point(498, 317)
        Me.lblAccount.Name = "lblAccount"
        Me.lblAccount.Size = New System.Drawing.Size(57, 13)
        Me.lblAccount.TabIndex = 43
        Me.lblAccount.Text = "lblAccount"
        '
        'cboStatus
        '
        Me.cboStatus.FormattingEnabled = True
        Me.cboStatus.Location = New System.Drawing.Point(499, 341)
        Me.cboStatus.Name = "cboStatus"
        Me.cboStatus.Size = New System.Drawing.Size(121, 21)
        Me.cboStatus.TabIndex = 44
        '
        'lblFirstName
        '
        Me.lblFirstName.AutoSize = True
        Me.lblFirstName.Location = New System.Drawing.Point(498, 365)
        Me.lblFirstName.Name = "lblFirstName"
        Me.lblFirstName.Size = New System.Drawing.Size(64, 13)
        Me.lblFirstName.TabIndex = 45
        Me.lblFirstName.Text = "lblFirstName"
        '
        'lblLastName
        '
        Me.lblLastName.AutoSize = True
        Me.lblLastName.Location = New System.Drawing.Point(501, 384)
        Me.lblLastName.Name = "lblLastName"
        Me.lblLastName.Size = New System.Drawing.Size(65, 13)
        Me.lblLastName.TabIndex = 46
        Me.lblLastName.Text = "lblLastName"
        '
        'lblAddress
        '
        Me.lblAddress.AutoSize = True
        Me.lblAddress.Location = New System.Drawing.Point(496, 404)
        Me.lblAddress.Name = "lblAddress"
        Me.lblAddress.Size = New System.Drawing.Size(55, 13)
        Me.lblAddress.TabIndex = 47
        Me.lblAddress.Text = "lblAddress"
        '
        'lblCity
        '
        Me.lblCity.AutoSize = True
        Me.lblCity.Location = New System.Drawing.Point(498, 421)
        Me.lblCity.Name = "lblCity"
        Me.lblCity.Size = New System.Drawing.Size(34, 13)
        Me.lblCity.TabIndex = 48
        Me.lblCity.Text = "lblCity"
        '
        'lblZip
        '
        Me.lblZip.AutoSize = True
        Me.lblZip.Location = New System.Drawing.Point(601, 42)
        Me.lblZip.Name = "lblZip"
        Me.lblZip.Size = New System.Drawing.Size(32, 13)
        Me.lblZip.TabIndex = 49
        Me.lblZip.Text = "lblZip"
        '
        'lblTele1
        '
        Me.lblTele1.AutoSize = True
        Me.lblTele1.Location = New System.Drawing.Point(601, 56)
        Me.lblTele1.Name = "lblTele1"
        Me.lblTele1.Size = New System.Drawing.Size(39, 13)
        Me.lblTele1.TabIndex = 50
        Me.lblTele1.Text = "Label1"
        '
        'lblTele2
        '
        Me.lblTele2.AutoSize = True
        Me.lblTele2.Location = New System.Drawing.Point(601, 72)
        Me.lblTele2.Name = "lblTele2"
        Me.lblTele2.Size = New System.Drawing.Size(44, 13)
        Me.lblTele2.TabIndex = 51
        Me.lblTele2.Text = "lblTele2"
        '
        'lblSSN
        '
        Me.lblSSN.AutoSize = True
        Me.lblSSN.Location = New System.Drawing.Point(601, 94)
        Me.lblSSN.Name = "lblSSN"
        Me.lblSSN.Size = New System.Drawing.Size(39, 13)
        Me.lblSSN.TabIndex = 52
        Me.lblSSN.Text = "lblSSN"
        '
        'txtLateChargeAmount
        '
        Me.txtLateChargeAmount.Location = New System.Drawing.Point(604, 110)
        Me.txtLateChargeAmount.Name = "txtLateChargeAmount"
        Me.txtLateChargeAmount.Size = New System.Drawing.Size(100, 20)
        Me.txtLateChargeAmount.TabIndex = 53
        '
        'lblCreditLimit
        '
        Me.lblCreditLimit.AutoSize = True
        Me.lblCreditLimit.Location = New System.Drawing.Point(594, 134)
        Me.lblCreditLimit.Name = "lblCreditLimit"
        Me.lblCreditLimit.Size = New System.Drawing.Size(39, 13)
        Me.lblCreditLimit.TabIndex = 54
        Me.lblCreditLimit.Text = "Label1"
        '
        'lblApprovalTerms
        '
        Me.lblApprovalTerms.AutoSize = True
        Me.lblApprovalTerms.Location = New System.Drawing.Point(594, 156)
        Me.lblApprovalTerms.Name = "lblApprovalTerms"
        Me.lblApprovalTerms.Size = New System.Drawing.Size(39, 13)
        Me.lblApprovalTerms.TabIndex = 55
        Me.lblApprovalTerms.Text = "Label1"
        '
        'lblAddAddress
        '
        Me.lblAddAddress.AutoSize = True
        Me.lblAddAddress.Location = New System.Drawing.Point(594, 173)
        Me.lblAddAddress.Name = "lblAddAddress"
        Me.lblAddAddress.Size = New System.Drawing.Size(74, 13)
        Me.lblAddAddress.TabIndex = 56
        Me.lblAddAddress.Text = "lblAddAddress"
        '
        'lblTele3
        '
        Me.lblTele3.AutoSize = True
        Me.lblTele3.Location = New System.Drawing.Point(594, 199)
        Me.lblTele3.Name = "lblTele3"
        Me.lblTele3.Size = New System.Drawing.Size(44, 13)
        Me.lblTele3.TabIndex = 57
        Me.lblTele3.Text = "lblTele3"
        '
        'lblTele1Caption
        '
        Me.lblTele1Caption.AutoSize = True
        Me.lblTele1Caption.Location = New System.Drawing.Point(601, 218)
        Me.lblTele1Caption.Name = "lblTele1Caption"
        Me.lblTele1Caption.Size = New System.Drawing.Size(80, 13)
        Me.lblTele1Caption.TabIndex = 58
        Me.lblTele1Caption.Text = "lblTele1Caption"
        '
        'lblTele2Caption
        '
        Me.lblTele2Caption.AutoSize = True
        Me.lblTele2Caption.Location = New System.Drawing.Point(594, 243)
        Me.lblTele2Caption.Name = "lblTele2Caption"
        Me.lblTele2Caption.Size = New System.Drawing.Size(80, 13)
        Me.lblTele2Caption.TabIndex = 59
        Me.lblTele2Caption.Text = "lblTele2Caption"
        '
        'lblTele3Caption
        '
        Me.lblTele3Caption.AutoSize = True
        Me.lblTele3Caption.Location = New System.Drawing.Point(588, 263)
        Me.lblTele3Caption.Name = "lblTele3Caption"
        Me.lblTele3Caption.Size = New System.Drawing.Size(80, 13)
        Me.lblTele3Caption.TabIndex = 60
        Me.lblTele3Caption.Text = "lblTele3Caption"
        '
        'fraPrint
        '
        Me.fraPrint.Location = New System.Drawing.Point(55, 486)
        Me.fraPrint.Name = "fraPrint"
        Me.fraPrint.Size = New System.Drawing.Size(145, 55)
        Me.fraPrint.TabIndex = 61
        Me.fraPrint.TabStop = False
        Me.fraPrint.Text = "fraPrint"
        '
        'fraNav
        '
        Me.fraNav.Location = New System.Drawing.Point(219, 486)
        Me.fraNav.Name = "fraNav"
        Me.fraNav.Size = New System.Drawing.Size(145, 55)
        Me.fraNav.TabIndex = 62
        Me.fraNav.TabStop = False
        Me.fraNav.Text = "fraNav"
        '
        'fraTerms
        '
        Me.fraTerms.Location = New System.Drawing.Point(387, 486)
        Me.fraTerms.Name = "fraTerms"
        Me.fraTerms.Size = New System.Drawing.Size(145, 55)
        Me.fraTerms.TabIndex = 63
        Me.fraTerms.TabStop = False
        Me.fraTerms.Text = "GroupBox1"
        '
        'fraBalance
        '
        Me.fraBalance.Location = New System.Drawing.Point(559, 486)
        Me.fraBalance.Name = "fraBalance"
        Me.fraBalance.Size = New System.Drawing.Size(145, 55)
        Me.fraBalance.TabIndex = 64
        Me.fraBalance.TabStop = False
        Me.fraBalance.Text = "fraBalance"
        '
        'fraPrintType
        '
        Me.fraPrintType.Location = New System.Drawing.Point(724, 498)
        Me.fraPrintType.Name = "fraPrintType"
        Me.fraPrintType.Size = New System.Drawing.Size(145, 55)
        Me.fraPrintType.TabIndex = 65
        Me.fraPrintType.TabStop = False
        Me.fraPrintType.Text = "GroupBox1"
        '
        'cmdCreditApp
        '
        Me.cmdCreditApp.Location = New System.Drawing.Point(219, 28)
        Me.cmdCreditApp.Name = "cmdCreditApp"
        Me.cmdCreditApp.Size = New System.Drawing.Size(75, 23)
        Me.cmdCreditApp.TabIndex = 66
        Me.cmdCreditApp.Text = "Credi&t App"
        Me.cmdCreditApp.UseVisualStyleBackColor = True
        '
        'cmdDetail
        '
        Me.cmdDetail.Location = New System.Drawing.Point(219, 62)
        Me.cmdDetail.Name = "cmdDetail"
        Me.cmdDetail.Size = New System.Drawing.Size(75, 23)
        Me.cmdDetail.TabIndex = 67
        Me.cmdDetail.Text = "&Detail"
        Me.cmdDetail.UseVisualStyleBackColor = True
        '
        'Notes_Open
        '
        Me.Notes_Open.Location = New System.Drawing.Point(219, 100)
        Me.Notes_Open.Name = "Notes_Open"
        Me.Notes_Open.Size = New System.Drawing.Size(75, 23)
        Me.Notes_Open.TabIndex = 68
        Me.Notes_Open.Text = "&Notes"
        Me.Notes_Open.UseVisualStyleBackColor = True
        '
        'cmdPrintCard
        '
        Me.cmdPrintCard.Location = New System.Drawing.Point(219, 134)
        Me.cmdPrintCard.Name = "cmdPrintCard"
        Me.cmdPrintCard.Size = New System.Drawing.Size(75, 23)
        Me.cmdPrintCard.TabIndex = 69
        Me.cmdPrintCard.Text = "&Print"
        Me.cmdPrintCard.UseVisualStyleBackColor = True
        '
        'lblAPR
        '
        Me.lblAPR.AutoSize = True
        Me.lblAPR.Location = New System.Drawing.Point(606, 293)
        Me.lblAPR.Name = "lblAPR"
        Me.lblAPR.Size = New System.Drawing.Size(63, 13)
        Me.lblAPR.TabIndex = 70
        Me.lblAPR.Text = "##.## APR"
        '
        'chkSendAllMail
        '
        Me.chkSendAllMail.AutoSize = True
        Me.chkSendAllMail.Location = New System.Drawing.Point(650, 344)
        Me.chkSendAllMail.Name = "chkSendAllMail"
        Me.chkSendAllMail.Size = New System.Drawing.Size(87, 17)
        Me.chkSendAllMail.TabIndex = 71
        Me.chkSendAllMail.Text = "Send All Mail"
        Me.chkSendAllMail.UseVisualStyleBackColor = True
        '
        'lblPaymentHistory
        '
        Me.lblPaymentHistory.AutoSize = True
        Me.lblPaymentHistory.Location = New System.Drawing.Point(721, 182)
        Me.lblPaymentHistory.Name = "lblPaymentHistory"
        Me.lblPaymentHistory.Size = New System.Drawing.Size(94, 13)
        Me.lblPaymentHistory.TabIndex = 72
        Me.lblPaymentHistory.Text = "Date: 00/00/0000"
        '
        'cmdSaleTotals
        '
        Me.cmdSaleTotals.Location = New System.Drawing.Point(219, 168)
        Me.cmdSaleTotals.Name = "cmdSaleTotals"
        Me.cmdSaleTotals.Size = New System.Drawing.Size(75, 23)
        Me.cmdSaleTotals.TabIndex = 73
        Me.cmdSaleTotals.Text = "Sale Totals"
        Me.cmdSaleTotals.UseVisualStyleBackColor = True
        '
        'UGrSaleTotals
        '
        Me.UGrSaleTotals.Activated = False
        Me.UGrSaleTotals.Col = 1
        Me.UGrSaleTotals.firstrow = 1
        Me.UGrSaleTotals.Loading = False
        Me.UGrSaleTotals.Location = New System.Drawing.Point(650, 365)
        Me.UGrSaleTotals.MaxCols = 2
        Me.UGrSaleTotals.MaxRows = 10
        Me.UGrSaleTotals.Name = "UGrSaleTotals"
        Me.UGrSaleTotals.Row = 0
        Me.UGrSaleTotals.Size = New System.Drawing.Size(230, 85)
        Me.UGrSaleTotals.TabIndex = 74
        '
        'ArCard
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(988, 604)
        Me.Controls.Add(Me.UGrSaleTotals)
        Me.Controls.Add(Me.cmdSaleTotals)
        Me.Controls.Add(Me.lblPaymentHistory)
        Me.Controls.Add(Me.chkSendAllMail)
        Me.Controls.Add(Me.lblAPR)
        Me.Controls.Add(Me.cmdPrintCard)
        Me.Controls.Add(Me.Notes_Open)
        Me.Controls.Add(Me.cmdDetail)
        Me.Controls.Add(Me.cmdCreditApp)
        Me.Controls.Add(Me.fraPrintType)
        Me.Controls.Add(Me.fraBalance)
        Me.Controls.Add(Me.fraTerms)
        Me.Controls.Add(Me.fraNav)
        Me.Controls.Add(Me.fraPrint)
        Me.Controls.Add(Me.lblTele3Caption)
        Me.Controls.Add(Me.lblTele2Caption)
        Me.Controls.Add(Me.lblTele1Caption)
        Me.Controls.Add(Me.lblTele3)
        Me.Controls.Add(Me.lblAddAddress)
        Me.Controls.Add(Me.lblApprovalTerms)
        Me.Controls.Add(Me.lblCreditLimit)
        Me.Controls.Add(Me.txtLateChargeAmount)
        Me.Controls.Add(Me.lblSSN)
        Me.Controls.Add(Me.lblTele2)
        Me.Controls.Add(Me.lblTele1)
        Me.Controls.Add(Me.lblZip)
        Me.Controls.Add(Me.lblCity)
        Me.Controls.Add(Me.lblAddress)
        Me.Controls.Add(Me.lblLastName)
        Me.Controls.Add(Me.lblFirstName)
        Me.Controls.Add(Me.cboStatus)
        Me.Controls.Add(Me.lblAccount)
        Me.Controls.Add(Me.filFile)
        Me.Controls.Add(Me.rtfFile)
        Me.Controls.Add(Me.lblOver91)
        Me.Controls.Add(Me.lbl6190)
        Me.Controls.Add(Me.lbl3160)
        Me.Controls.Add(Me.lbl0030)
        Me.Controls.Add(Me.lblLate91)
        Me.Controls.Add(Me.lblLate61)
        Me.Controls.Add(Me.lblLate31)
        Me.Controls.Add(Me.lblLate0)
        Me.Controls.Add(Me.lblArrearages)
        Me.Controls.Add(Me.lblTotDue)
        Me.Controls.Add(Me.txtNextDue)
        Me.Controls.Add(Me.txtPayPeriod)
        Me.Controls.Add(Me.txtPaidBy)
        Me.Controls.Add(Me.txtFirstPay)
        Me.Controls.Add(Me.lblLateCharge)
        Me.Controls.Add(Me.txtSameAsCash)
        Me.Controls.Add(Me.txtRate)
        Me.Controls.Add(Me.txtPayMemo)
        Me.Controls.Add(Me.txtMonthlyPayment)
        Me.Controls.Add(Me.txtFinanced)
        Me.Controls.Add(Me.txtLastPay)
        Me.Controls.Add(Me.txtMonths)
        Me.Controls.Add(Me.txtDelivery)
        Me.Controls.Add(Me.lblBalance)
        Me.Controls.Add(Me.txtPaymentHistory)
        Me.Controls.Add(Me.fraEditOptions)
        Me.Controls.Add(Me.fraPaymentOptions)
        Me.Controls.Add(Me.cmdCancel)
        Me.Controls.Add(Me.cmdPrint)
        Me.Controls.Add(Me.cmdReprintContract)
        Me.Controls.Add(Me.cmdReceipt)
        Me.Controls.Add(Me.cmdPayoff)
        Me.Controls.Add(Me.cmdMovePrevious)
        Me.Controls.Add(Me.cmdMoveNext)
        Me.Controls.Add(Me.cmdMoveLast)
        Me.Controls.Add(Me.cmdMoveFirst)
        Me.Controls.Add(Me.cmdMakeSameAsCash)
        Me.Controls.Add(Me.cmdFields)
        Me.Controls.Add(Me.cmdExport)
        Me.Controls.Add(Me.cmdEdit)
        Me.Controls.Add(Me.cmdApply)
        Me.Controls.Add(Me.lblTotalPayoff)
        Me.Name = "ArCard"
        Me.Text = "ArCardvb"
        Me.ResumeLayout(False)
        Me.PerformLayout()

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
    Friend WithEvents fraPaymentOptions As GroupBox
    Friend WithEvents fraEditOptions As GroupBox
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
    Friend WithEvents rtfFile As RichTextBox
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
    Friend WithEvents fraPrintType As GroupBox
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
End Class
