<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class OrdVoid
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
        Me.lblVoidDate = New System.Windows.Forms.Label()
        Me.dteVoidDate = New System.Windows.Forms.DateTimePicker()
        Me.fraReturn = New System.Windows.Forms.GroupBox()
        Me.txtApplyToSaleNo = New System.Windows.Forms.TextBox()
        Me.optRefundType4 = New System.Windows.Forms.RadioButton()
        Me.optRefundType3 = New System.Windows.Forms.RadioButton()
        Me.optRefundType2 = New System.Windows.Forms.RadioButton()
        Me.optRefundType1 = New System.Windows.Forms.RadioButton()
        Me.optRefundType0 = New System.Windows.Forms.RadioButton()
        Me.lblGeneral5 = New System.Windows.Forms.Label()
        Me.txtVoidNote = New System.Windows.Forms.TextBox()
        Me.fraControls = New System.Windows.Forms.GroupBox()
        Me.cmdCancel = New System.Windows.Forms.Button()
        Me.cmdOk = New System.Windows.Forms.Button()
        Me.fraPaymentSummary = New System.Windows.Forms.GroupBox()
        Me.lblRefundTotal = New System.Windows.Forms.Label()
        Me.lblTotalPaid = New System.Windows.Forms.Label()
        Me.lblTotalPaidLabel = New System.Windows.Forms.Label()
        Me.txtForfeit = New System.Windows.Forms.Label()
        Me.lblForfeit = New System.Windows.Forms.Label()
        Me.txtRefundSpecial = New System.Windows.Forms.TextBox()
        Me.lblSpecialPaymentType = New System.Windows.Forms.Label()
        Me.txtRefundAmount = New System.Windows.Forms.TextBox()
        Me.lblAmountPaid = New System.Windows.Forms.Label()
        Me.lblPaymentType = New System.Windows.Forms.Label()
        Me.lblGeneral2 = New System.Windows.Forms.Label()
        Me.lblGeneral1 = New System.Windows.Forms.Label()
        Me.lblGeneral0 = New System.Windows.Forms.Label()
        Me.fraReturn.SuspendLayout()
        Me.fraControls.SuspendLayout()
        Me.fraPaymentSummary.SuspendLayout()
        Me.SuspendLayout()
        '
        'lblVoidDate
        '
        Me.lblVoidDate.AutoSize = True
        Me.lblVoidDate.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblVoidDate.Location = New System.Drawing.Point(64, 14)
        Me.lblVoidDate.Name = "lblVoidDate"
        Me.lblVoidDate.Size = New System.Drawing.Size(94, 20)
        Me.lblVoidDate.TabIndex = 0
        Me.lblVoidDate.Text = "Void Date:"
        '
        'dteVoidDate
        '
        Me.dteVoidDate.CustomFormat = "MM/dd/yyyy"
        Me.dteVoidDate.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dteVoidDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dteVoidDate.Location = New System.Drawing.Point(161, 9)
        Me.dteVoidDate.Name = "dteVoidDate"
        Me.dteVoidDate.Size = New System.Drawing.Size(101, 26)
        Me.dteVoidDate.TabIndex = 1
        '
        'fraReturn
        '
        Me.fraReturn.Controls.Add(Me.txtApplyToSaleNo)
        Me.fraReturn.Controls.Add(Me.optRefundType4)
        Me.fraReturn.Controls.Add(Me.optRefundType3)
        Me.fraReturn.Controls.Add(Me.optRefundType2)
        Me.fraReturn.Controls.Add(Me.optRefundType1)
        Me.fraReturn.Controls.Add(Me.optRefundType0)
        Me.fraReturn.Location = New System.Drawing.Point(12, 39)
        Me.fraReturn.Name = "fraReturn"
        Me.fraReturn.Size = New System.Drawing.Size(146, 130)
        Me.fraReturn.TabIndex = 2
        Me.fraReturn.TabStop = False
        '
        'txtApplyToSaleNo
        '
        Me.txtApplyToSaleNo.Location = New System.Drawing.Point(96, 106)
        Me.txtApplyToSaleNo.Name = "txtApplyToSaleNo"
        Me.txtApplyToSaleNo.Size = New System.Drawing.Size(48, 20)
        Me.txtApplyToSaleNo.TabIndex = 5
        Me.txtApplyToSaleNo.Visible = False
        '
        'optRefundType4
        '
        Me.optRefundType4.AutoSize = True
        Me.optRefundType4.Location = New System.Drawing.Point(8, 107)
        Me.optRefundType4.Name = "optRefundType4"
        Me.optRefundType4.Size = New System.Drawing.Size(90, 17)
        Me.optRefundType4.TabIndex = 4
        Me.optRefundType4.Text = "Apply to Sale:"
        Me.optRefundType4.UseVisualStyleBackColor = True
        '
        'optRefundType3
        '
        Me.optRefundType3.AutoSize = True
        Me.optRefundType3.Location = New System.Drawing.Point(8, 84)
        Me.optRefundType3.Name = "optRefundType3"
        Me.optRefundType3.Size = New System.Drawing.Size(93, 17)
        Me.optRefundType3.TabIndex = 3
        Me.optRefundType3.Text = "&Forfeit Deposit"
        Me.optRefundType3.UseVisualStyleBackColor = True
        '
        'optRefundType2
        '
        Me.optRefundType2.AutoSize = True
        Me.optRefundType2.Location = New System.Drawing.Point(8, 61)
        Me.optRefundType2.Name = "optRefundType2"
        Me.optRefundType2.Size = New System.Drawing.Size(108, 17)
        Me.optRefundType2.TabIndex = 2
        Me.optRefundType2.Text = "&Issue Store Credit"
        Me.optRefundType2.UseVisualStyleBackColor = True
        '
        'optRefundType1
        '
        Me.optRefundType1.AutoSize = True
        Me.optRefundType1.Location = New System.Drawing.Point(8, 38)
        Me.optRefundType1.Name = "optRefundType1"
        Me.optRefundType1.Size = New System.Drawing.Size(131, 17)
        Me.optRefundType1.TabIndex = 1
        Me.optRefundType1.Text = "&Write Company Check"
        Me.optRefundType1.UseVisualStyleBackColor = True
        '
        'optRefundType0
        '
        Me.optRefundType0.AutoSize = True
        Me.optRefundType0.Checked = True
        Me.optRefundType0.Location = New System.Drawing.Point(8, 15)
        Me.optRefundType0.Name = "optRefundType0"
        Me.optRefundType0.Size = New System.Drawing.Size(130, 17)
        Me.optRefundType0.TabIndex = 0
        Me.optRefundType0.TabStop = True
        Me.optRefundType0.Text = "&Return Money as Paid"
        Me.optRefundType0.UseVisualStyleBackColor = True
        '
        'lblGeneral5
        '
        Me.lblGeneral5.AutoSize = True
        Me.lblGeneral5.Location = New System.Drawing.Point(211, 49)
        Me.lblGeneral5.Name = "lblGeneral5"
        Me.lblGeneral5.Size = New System.Drawing.Size(69, 13)
        Me.lblGeneral5.TabIndex = 3
        Me.lblGeneral5.Text = "Note on Void"
        '
        'txtVoidNote
        '
        Me.txtVoidNote.Location = New System.Drawing.Point(164, 67)
        Me.txtVoidNote.Name = "txtVoidNote"
        Me.txtVoidNote.Size = New System.Drawing.Size(173, 20)
        Me.txtVoidNote.TabIndex = 4
        '
        'fraControls
        '
        Me.fraControls.Controls.Add(Me.cmdCancel)
        Me.fraControls.Controls.Add(Me.cmdOk)
        Me.fraControls.Location = New System.Drawing.Point(164, 93)
        Me.fraControls.Name = "fraControls"
        Me.fraControls.Size = New System.Drawing.Size(173, 76)
        Me.fraControls.TabIndex = 5
        Me.fraControls.TabStop = False
        '
        'cmdCancel
        '
        Me.cmdCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cmdCancel.Location = New System.Drawing.Point(92, 16)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.Size = New System.Drawing.Size(71, 49)
        Me.cmdCancel.TabIndex = 1
        Me.cmdCancel.Text = "&Cancel"
        Me.cmdCancel.UseVisualStyleBackColor = True
        '
        'cmdOk
        '
        Me.cmdOk.Location = New System.Drawing.Point(15, 16)
        Me.cmdOk.Name = "cmdOk"
        Me.cmdOk.Size = New System.Drawing.Size(71, 49)
        Me.cmdOk.TabIndex = 0
        Me.cmdOk.Text = "&OK"
        Me.cmdOk.UseVisualStyleBackColor = True
        '
        'fraPaymentSummary
        '
        Me.fraPaymentSummary.Controls.Add(Me.lblRefundTotal)
        Me.fraPaymentSummary.Controls.Add(Me.lblTotalPaid)
        Me.fraPaymentSummary.Controls.Add(Me.lblTotalPaidLabel)
        Me.fraPaymentSummary.Controls.Add(Me.txtForfeit)
        Me.fraPaymentSummary.Controls.Add(Me.lblForfeit)
        Me.fraPaymentSummary.Controls.Add(Me.txtRefundSpecial)
        Me.fraPaymentSummary.Controls.Add(Me.lblSpecialPaymentType)
        Me.fraPaymentSummary.Controls.Add(Me.txtRefundAmount)
        Me.fraPaymentSummary.Controls.Add(Me.lblAmountPaid)
        Me.fraPaymentSummary.Controls.Add(Me.lblPaymentType)
        Me.fraPaymentSummary.Controls.Add(Me.lblGeneral2)
        Me.fraPaymentSummary.Controls.Add(Me.lblGeneral1)
        Me.fraPaymentSummary.Controls.Add(Me.lblGeneral0)
        Me.fraPaymentSummary.Location = New System.Drawing.Point(12, 174)
        Me.fraPaymentSummary.Name = "fraPaymentSummary"
        Me.fraPaymentSummary.Size = New System.Drawing.Size(334, 158)
        Me.fraPaymentSummary.TabIndex = 6
        Me.fraPaymentSummary.TabStop = False
        Me.fraPaymentSummary.Text = " Payment Summary "
        '
        'lblRefundTotal
        '
        Me.lblRefundTotal.AutoSize = True
        Me.lblRefundTotal.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblRefundTotal.Location = New System.Drawing.Point(282, 135)
        Me.lblRefundTotal.Name = "lblRefundTotal"
        Me.lblRefundTotal.Size = New System.Drawing.Size(39, 13)
        Me.lblRefundTotal.TabIndex = 12
        Me.lblRefundTotal.Text = "$0.00"
        Me.lblRefundTotal.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'lblTotalPaid
        '
        Me.lblTotalPaid.AutoSize = True
        Me.lblTotalPaid.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblTotalPaid.Location = New System.Drawing.Point(163, 135)
        Me.lblTotalPaid.Name = "lblTotalPaid"
        Me.lblTotalPaid.Size = New System.Drawing.Size(39, 13)
        Me.lblTotalPaid.TabIndex = 11
        Me.lblTotalPaid.Text = "$0.00"
        Me.lblTotalPaid.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'lblTotalPaidLabel
        '
        Me.lblTotalPaidLabel.AutoSize = True
        Me.lblTotalPaidLabel.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblTotalPaidLabel.Location = New System.Drawing.Point(8, 135)
        Me.lblTotalPaidLabel.Name = "lblTotalPaidLabel"
        Me.lblTotalPaidLabel.Size = New System.Drawing.Size(69, 13)
        Me.lblTotalPaidLabel.TabIndex = 10
        Me.lblTotalPaidLabel.Text = "Total Paid:"
        '
        'txtForfeit
        '
        Me.txtForfeit.AutoSize = True
        Me.txtForfeit.Location = New System.Drawing.Point(287, 110)
        Me.txtForfeit.Name = "txtForfeit"
        Me.txtForfeit.Size = New System.Drawing.Size(34, 13)
        Me.txtForfeit.TabIndex = 9
        Me.txtForfeit.Text = "$0.00"
        Me.txtForfeit.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'lblForfeit
        '
        Me.lblForfeit.AutoSize = True
        Me.lblForfeit.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblForfeit.Location = New System.Drawing.Point(8, 110)
        Me.lblForfeit.Name = "lblForfeit"
        Me.lblForfeit.Size = New System.Drawing.Size(94, 13)
        Me.lblForfeit.TabIndex = 8
        Me.lblForfeit.Text = "Forfeit Deposit:"
        '
        'txtRefundSpecial
        '
        Me.txtRefundSpecial.Location = New System.Drawing.Point(230, 80)
        Me.txtRefundSpecial.Name = "txtRefundSpecial"
        Me.txtRefundSpecial.Size = New System.Drawing.Size(91, 20)
        Me.txtRefundSpecial.TabIndex = 7
        Me.txtRefundSpecial.Text = "0.00"
        Me.txtRefundSpecial.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'lblSpecialPaymentType
        '
        Me.lblSpecialPaymentType.AutoSize = True
        Me.lblSpecialPaymentType.Location = New System.Drawing.Point(8, 80)
        Me.lblSpecialPaymentType.Name = "lblSpecialPaymentType"
        Me.lblSpecialPaymentType.Size = New System.Drawing.Size(99, 13)
        Me.lblSpecialPaymentType.TabIndex = 6
        Me.lblSpecialPaymentType.Text = "COMPANY CHECK"
        '
        'txtRefundAmount
        '
        Me.txtRefundAmount.Location = New System.Drawing.Point(230, 54)
        Me.txtRefundAmount.Name = "txtRefundAmount"
        Me.txtRefundAmount.Size = New System.Drawing.Size(91, 20)
        Me.txtRefundAmount.TabIndex = 5
        Me.txtRefundAmount.Text = "0.00"
        Me.txtRefundAmount.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'lblAmountPaid
        '
        Me.lblAmountPaid.AutoSize = True
        Me.lblAmountPaid.Location = New System.Drawing.Point(174, 54)
        Me.lblAmountPaid.Name = "lblAmountPaid"
        Me.lblAmountPaid.Size = New System.Drawing.Size(28, 13)
        Me.lblAmountPaid.TabIndex = 4
        Me.lblAmountPaid.Text = "0.00"
        Me.lblAmountPaid.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'lblPaymentType
        '
        Me.lblPaymentType.AutoSize = True
        Me.lblPaymentType.Location = New System.Drawing.Point(8, 54)
        Me.lblPaymentType.Name = "lblPaymentType"
        Me.lblPaymentType.Size = New System.Drawing.Size(36, 13)
        Me.lblPaymentType.TabIndex = 3
        Me.lblPaymentType.Text = "CASH"
        '
        'lblGeneral2
        '
        Me.lblGeneral2.AutoSize = True
        Me.lblGeneral2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblGeneral2.Location = New System.Drawing.Point(227, 30)
        Me.lblGeneral2.Name = "lblGeneral2"
        Me.lblGeneral2.Size = New System.Drawing.Size(94, 13)
        Me.lblGeneral2.TabIndex = 2
        Me.lblGeneral2.Text = "Refund Amount"
        '
        'lblGeneral1
        '
        Me.lblGeneral1.AutoSize = True
        Me.lblGeneral1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblGeneral1.Location = New System.Drawing.Point(136, 30)
        Me.lblGeneral1.Name = "lblGeneral1"
        Me.lblGeneral1.Size = New System.Drawing.Size(78, 13)
        Me.lblGeneral1.TabIndex = 1
        Me.lblGeneral1.Text = "Amount Paid"
        '
        'lblGeneral0
        '
        Me.lblGeneral0.AutoSize = True
        Me.lblGeneral0.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblGeneral0.Location = New System.Drawing.Point(8, 30)
        Me.lblGeneral0.Name = "lblGeneral0"
        Me.lblGeneral0.Size = New System.Drawing.Size(87, 13)
        Me.lblGeneral0.TabIndex = 0
        Me.lblGeneral0.Text = "Payment Type"
        '
        'OrdVoid
        '
        Me.AcceptButton = Me.cmdOk
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.cmdCancel
        Me.ClientSize = New System.Drawing.Size(351, 336)
        Me.Controls.Add(Me.fraPaymentSummary)
        Me.Controls.Add(Me.fraControls)
        Me.Controls.Add(Me.txtVoidNote)
        Me.Controls.Add(Me.lblGeneral5)
        Me.Controls.Add(Me.fraReturn)
        Me.Controls.Add(Me.dteVoidDate)
        Me.Controls.Add(Me.lblVoidDate)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.MaximizeBox = False
        Me.Name = "OrdVoid"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Void Menu"
        Me.fraReturn.ResumeLayout(False)
        Me.fraReturn.PerformLayout()
        Me.fraControls.ResumeLayout(False)
        Me.fraPaymentSummary.ResumeLayout(False)
        Me.fraPaymentSummary.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents lblVoidDate As Label
    Friend WithEvents dteVoidDate As DateTimePicker
    Friend WithEvents fraReturn As GroupBox
    Friend WithEvents txtApplyToSaleNo As TextBox
    Friend WithEvents optRefundType4 As RadioButton
    Friend WithEvents optRefundType3 As RadioButton
    Friend WithEvents optRefundType2 As RadioButton
    Friend WithEvents optRefundType1 As RadioButton
    Friend WithEvents optRefundType0 As RadioButton
    Friend WithEvents lblGeneral5 As Label
    Friend WithEvents txtVoidNote As TextBox
    Friend WithEvents fraControls As GroupBox
    Friend WithEvents cmdCancel As Button
    Friend WithEvents cmdOk As Button
    Friend WithEvents fraPaymentSummary As GroupBox
    Friend WithEvents lblRefundTotal As Label
    Friend WithEvents lblTotalPaid As Label
    Friend WithEvents lblTotalPaidLabel As Label
    Friend WithEvents txtForfeit As Label
    Friend WithEvents lblForfeit As Label
    Friend WithEvents txtRefundSpecial As TextBox
    Friend WithEvents lblSpecialPaymentType As Label
    Friend WithEvents txtRefundAmount As TextBox
    Friend WithEvents lblAmountPaid As Label
    Friend WithEvents lblPaymentType As Label
    Friend WithEvents lblGeneral2 As Label
    Friend WithEvents lblGeneral1 As Label
    Friend WithEvents lblGeneral0 As Label
End Class
