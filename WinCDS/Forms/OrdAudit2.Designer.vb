<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class OrdAudit2
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
        Me.lblStartDateLabel = New System.Windows.Forms.Label()
        Me.dteStartDate = New System.Windows.Forms.DateTimePicker()
        Me.lblEndDateLabel = New System.Windows.Forms.Label()
        Me.dteEndDate = New System.Windows.Forms.DateTimePicker()
        Me.lblCashInDrawer = New System.Windows.Forms.Label()
        Me.txtCashInDrawer = New System.Windows.Forms.TextBox()
        Me.lblPriorPeriodCash = New System.Windows.Forms.Label()
        Me.txtPriorPeriodCash = New System.Windows.Forms.TextBox()
        Me.optDetail = New System.Windows.Forms.RadioButton()
        Me.optSummary = New System.Windows.Forms.RadioButton()
        Me.lblCashier = New System.Windows.Forms.Label()
        Me.cmbCashier = New System.Windows.Forms.ComboBox()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.cmdPrint = New System.Windows.Forms.Button()
        Me.cmdPrintPreview = New System.Windows.Forms.Button()
        Me.cmdCancel = New System.Windows.Forms.Button()
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'lblStartDateLabel
        '
        Me.lblStartDateLabel.AutoSize = True
        Me.lblStartDateLabel.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblStartDateLabel.Location = New System.Drawing.Point(8, 14)
        Me.lblStartDateLabel.Name = "lblStartDateLabel"
        Me.lblStartDateLabel.Size = New System.Drawing.Size(88, 16)
        Me.lblStartDateLabel.TabIndex = 0
        Me.lblStartDateLabel.Text = "&Starting Date:"
        '
        'dteStartDate
        '
        Me.dteStartDate.CustomFormat = "MM/dd/yyyy"
        Me.dteStartDate.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dteStartDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dteStartDate.Location = New System.Drawing.Point(11, 32)
        Me.dteStartDate.Name = "dteStartDate"
        Me.dteStartDate.Size = New System.Drawing.Size(100, 22)
        Me.dteStartDate.TabIndex = 1
        '
        'lblEndDateLabel
        '
        Me.lblEndDateLabel.AutoSize = True
        Me.lblEndDateLabel.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblEndDateLabel.Location = New System.Drawing.Point(153, 13)
        Me.lblEndDateLabel.Name = "lblEndDateLabel"
        Me.lblEndDateLabel.Size = New System.Drawing.Size(85, 16)
        Me.lblEndDateLabel.TabIndex = 2
        Me.lblEndDateLabel.Text = "&Ending Date:"
        '
        'dteEndDate
        '
        Me.dteEndDate.CustomFormat = "MM/dd/yyyy"
        Me.dteEndDate.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dteEndDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dteEndDate.Location = New System.Drawing.Point(152, 32)
        Me.dteEndDate.Name = "dteEndDate"
        Me.dteEndDate.Size = New System.Drawing.Size(93, 22)
        Me.dteEndDate.TabIndex = 3
        '
        'lblCashInDrawer
        '
        Me.lblCashInDrawer.AutoSize = True
        Me.lblCashInDrawer.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblCashInDrawer.Location = New System.Drawing.Point(19, 87)
        Me.lblCashInDrawer.Name = "lblCashInDrawer"
        Me.lblCashInDrawer.Size = New System.Drawing.Size(101, 16)
        Me.lblCashInDrawer.TabIndex = 4
        Me.lblCashInDrawer.Text = "C&ash In Drawer:"
        '
        'txtCashInDrawer
        '
        Me.txtCashInDrawer.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtCashInDrawer.Location = New System.Drawing.Point(124, 77)
        Me.txtCashInDrawer.Name = "txtCashInDrawer"
        Me.txtCashInDrawer.Size = New System.Drawing.Size(121, 26)
        Me.txtCashInDrawer.TabIndex = 5
        Me.txtCashInDrawer.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'lblPriorPeriodCash
        '
        Me.lblPriorPeriodCash.AutoSize = True
        Me.lblPriorPeriodCash.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblPriorPeriodCash.Location = New System.Drawing.Point(4, 116)
        Me.lblPriorPeriodCash.Name = "lblPriorPeriodCash"
        Me.lblPriorPeriodCash.Size = New System.Drawing.Size(116, 16)
        Me.lblPriorPeriodCash.TabIndex = 6
        Me.lblPriorPeriodCash.Text = "Pr&ior Period Cash:"
        '
        'txtPriorPeriodCash
        '
        Me.txtPriorPeriodCash.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtPriorPeriodCash.Location = New System.Drawing.Point(124, 109)
        Me.txtPriorPeriodCash.Name = "txtPriorPeriodCash"
        Me.txtPriorPeriodCash.Size = New System.Drawing.Size(121, 26)
        Me.txtPriorPeriodCash.TabIndex = 7
        Me.txtPriorPeriodCash.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'optDetail
        '
        Me.optDetail.AutoSize = True
        Me.optDetail.Checked = True
        Me.optDetail.Location = New System.Drawing.Point(56, 152)
        Me.optDetail.Name = "optDetail"
        Me.optDetail.Size = New System.Drawing.Size(55, 17)
        Me.optDetail.TabIndex = 8
        Me.optDetail.TabStop = True
        Me.optDetail.Text = "&Detail "
        Me.optDetail.UseVisualStyleBackColor = True
        '
        'optSummary
        '
        Me.optSummary.AutoSize = True
        Me.optSummary.Location = New System.Drawing.Point(124, 152)
        Me.optSummary.Name = "optSummary"
        Me.optSummary.Size = New System.Drawing.Size(68, 17)
        Me.optSummary.TabIndex = 9
        Me.optSummary.Text = "Su&mmary"
        Me.optSummary.UseVisualStyleBackColor = True
        '
        'lblCashier
        '
        Me.lblCashier.AutoSize = True
        Me.lblCashier.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblCashier.Location = New System.Drawing.Point(35, 180)
        Me.lblCashier.Name = "lblCashier"
        Me.lblCashier.Size = New System.Drawing.Size(76, 16)
        Me.lblCashier.TabIndex = 10
        Me.lblCashier.Text = "By Cashier:"
        '
        'cmbCashier
        '
        Me.cmbCashier.FormattingEnabled = True
        Me.cmbCashier.Location = New System.Drawing.Point(124, 175)
        Me.cmbCashier.Name = "cmbCashier"
        Me.cmbCashier.Size = New System.Drawing.Size(121, 21)
        Me.cmbCashier.TabIndex = 11
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.txtPriorPeriodCash)
        Me.GroupBox1.Controls.Add(Me.cmbCashier)
        Me.GroupBox1.Controls.Add(Me.lblStartDateLabel)
        Me.GroupBox1.Controls.Add(Me.lblCashier)
        Me.GroupBox1.Controls.Add(Me.dteStartDate)
        Me.GroupBox1.Controls.Add(Me.optSummary)
        Me.GroupBox1.Controls.Add(Me.lblEndDateLabel)
        Me.GroupBox1.Controls.Add(Me.optDetail)
        Me.GroupBox1.Controls.Add(Me.dteEndDate)
        Me.GroupBox1.Controls.Add(Me.lblCashInDrawer)
        Me.GroupBox1.Controls.Add(Me.lblPriorPeriodCash)
        Me.GroupBox1.Controls.Add(Me.txtCashInDrawer)
        Me.GroupBox1.Location = New System.Drawing.Point(3, 2)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(260, 210)
        Me.GroupBox1.TabIndex = 12
        Me.GroupBox1.TabStop = False
        '
        'cmdPrint
        '
        Me.cmdPrint.Location = New System.Drawing.Point(14, 216)
        Me.cmdPrint.Name = "cmdPrint"
        Me.cmdPrint.Size = New System.Drawing.Size(75, 62)
        Me.cmdPrint.TabIndex = 13
        Me.cmdPrint.Text = "&Print"
        Me.cmdPrint.UseVisualStyleBackColor = True
        '
        'cmdPrintPreview
        '
        Me.cmdPrintPreview.Location = New System.Drawing.Point(95, 216)
        Me.cmdPrintPreview.Name = "cmdPrintPreview"
        Me.cmdPrintPreview.Size = New System.Drawing.Size(78, 62)
        Me.cmdPrintPreview.TabIndex = 14
        Me.cmdPrintPreview.Text = "P&rint Preview"
        Me.cmdPrintPreview.UseVisualStyleBackColor = True
        '
        'cmdCancel
        '
        Me.cmdCancel.Location = New System.Drawing.Point(179, 216)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.Size = New System.Drawing.Size(75, 62)
        Me.cmdCancel.TabIndex = 15
        Me.cmdCancel.Text = "&Cancel"
        Me.cmdCancel.UseVisualStyleBackColor = True
        '
        'OrdAudit2
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(271, 281)
        Me.Controls.Add(Me.cmdCancel)
        Me.Controls.Add(Me.cmdPrintPreview)
        Me.Controls.Add(Me.cmdPrint)
        Me.Controls.Add(Me.GroupBox1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "OrdAudit2"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Daily Audit Report"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents lblStartDateLabel As Label
    Friend WithEvents dteStartDate As DateTimePicker
    Friend WithEvents lblEndDateLabel As Label
    Friend WithEvents dteEndDate As DateTimePicker
    Friend WithEvents lblCashInDrawer As Label
    Friend WithEvents txtCashInDrawer As TextBox
    Friend WithEvents lblPriorPeriodCash As Label
    Friend WithEvents txtPriorPeriodCash As TextBox
    Friend WithEvents optDetail As RadioButton
    Friend WithEvents optSummary As RadioButton
    Friend WithEvents lblCashier As Label
    Friend WithEvents cmbCashier As ComboBox
    Friend WithEvents GroupBox1 As GroupBox
    Friend WithEvents cmdPrint As Button
    Friend WithEvents cmdPrintPreview As Button
    Friend WithEvents cmdCancel As Button
End Class
