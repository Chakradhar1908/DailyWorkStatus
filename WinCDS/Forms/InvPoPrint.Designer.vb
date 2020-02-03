<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class InvPoPrint
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
        Me.fra = New System.Windows.Forms.GroupBox()
        Me.chkShowCost = New System.Windows.Forms.CheckBox()
        Me.chkVendorSort = New System.Windows.Forms.CheckBox()
        Me.cboLoc = New System.Windows.Forms.ComboBox()
        Me.lblLabel2 = New System.Windows.Forms.Label()
        Me.txtDate = New System.Windows.Forms.DateTimePicker()
        Me.lblLabel = New System.Windows.Forms.Label()
        Me.cmdPrint = New System.Windows.Forms.Button()
        Me.cmdPrintPreview = New System.Windows.Forms.Button()
        Me.cmdCancel = New System.Windows.Forms.Button()
        Me.cboSortOrder = New System.Windows.Forms.ComboBox()
        Me.fra.SuspendLayout()
        Me.SuspendLayout()
        '
        'fra
        '
        Me.fra.Controls.Add(Me.chkShowCost)
        Me.fra.Controls.Add(Me.chkVendorSort)
        Me.fra.Controls.Add(Me.cboLoc)
        Me.fra.Controls.Add(Me.lblLabel2)
        Me.fra.Controls.Add(Me.txtDate)
        Me.fra.Controls.Add(Me.lblLabel)
        Me.fra.Location = New System.Drawing.Point(10, 6)
        Me.fra.Name = "fra"
        Me.fra.Size = New System.Drawing.Size(161, 164)
        Me.fra.TabIndex = 0
        Me.fra.TabStop = False
        '
        'chkShowCost
        '
        Me.chkShowCost.AutoSize = True
        Me.chkShowCost.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.chkShowCost.Location = New System.Drawing.Point(20, 139)
        Me.chkShowCost.Name = "chkShowCost"
        Me.chkShowCost.Size = New System.Drawing.Size(77, 17)
        Me.chkShowCost.TabIndex = 5
        Me.chkShowCost.Text = "Show Cost"
        Me.chkShowCost.UseVisualStyleBackColor = False
        '
        'chkVendorSort
        '
        Me.chkVendorSort.AutoSize = True
        Me.chkVendorSort.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.chkVendorSort.Location = New System.Drawing.Point(20, 116)
        Me.chkVendorSort.Name = "chkVendorSort"
        Me.chkVendorSort.Size = New System.Drawing.Size(97, 17)
        Me.chkVendorSort.TabIndex = 4
        Me.chkVendorSort.Text = "Sort By Vendor"
        Me.chkVendorSort.UseVisualStyleBackColor = False
        '
        'cboLoc
        '
        Me.cboLoc.FormattingEnabled = True
        Me.cboLoc.Location = New System.Drawing.Point(20, 80)
        Me.cboLoc.Name = "cboLoc"
        Me.cboLoc.Size = New System.Drawing.Size(121, 21)
        Me.cboLoc.TabIndex = 3
        '
        'lblLabel2
        '
        Me.lblLabel2.AutoSize = True
        Me.lblLabel2.Location = New System.Drawing.Point(21, 64)
        Me.lblLabel2.Name = "lblLabel2"
        Me.lblLabel2.Size = New System.Drawing.Size(49, 13)
        Me.lblLabel2.TabIndex = 2
        Me.lblLabel2.Text = "lblLabel2"
        '
        'txtDate
        '
        Me.txtDate.CustomFormat = "MM/dd/yyyy"
        Me.txtDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.txtDate.Location = New System.Drawing.Point(20, 33)
        Me.txtDate.Name = "txtDate"
        Me.txtDate.Size = New System.Drawing.Size(121, 20)
        Me.txtDate.TabIndex = 1
        '
        'lblLabel
        '
        Me.lblLabel.AutoSize = True
        Me.lblLabel.Location = New System.Drawing.Point(17, 17)
        Me.lblLabel.Name = "lblLabel"
        Me.lblLabel.Size = New System.Drawing.Size(43, 13)
        Me.lblLabel.TabIndex = 0
        Me.lblLabel.Text = "lblLabel"
        '
        'cmdPrint
        '
        Me.cmdPrint.Location = New System.Drawing.Point(9, 176)
        Me.cmdPrint.Name = "cmdPrint"
        Me.cmdPrint.Size = New System.Drawing.Size(48, 47)
        Me.cmdPrint.TabIndex = 1
        Me.cmdPrint.Text = "&Print"
        Me.cmdPrint.UseVisualStyleBackColor = True
        '
        'cmdPrintPreview
        '
        Me.cmdPrintPreview.Location = New System.Drawing.Point(63, 176)
        Me.cmdPrintPreview.Name = "cmdPrintPreview"
        Me.cmdPrintPreview.Size = New System.Drawing.Size(48, 47)
        Me.cmdPrintPreview.TabIndex = 2
        Me.cmdPrintPreview.Text = "P&review"
        Me.cmdPrintPreview.UseVisualStyleBackColor = True
        '
        'cmdCancel
        '
        Me.cmdCancel.Location = New System.Drawing.Point(117, 176)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.Size = New System.Drawing.Size(48, 47)
        Me.cmdCancel.TabIndex = 3
        Me.cmdCancel.Text = "&Cancel"
        Me.cmdCancel.UseVisualStyleBackColor = True
        '
        'cboSortOrder
        '
        Me.cboSortOrder.FormattingEnabled = True
        Me.cboSortOrder.Location = New System.Drawing.Point(66, 233)
        Me.cboSortOrder.Name = "cboSortOrder"
        Me.cboSortOrder.Size = New System.Drawing.Size(121, 21)
        Me.cboSortOrder.TabIndex = 4
        '
        'InvPoPrint
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(188, 247)
        Me.Controls.Add(Me.cboSortOrder)
        Me.Controls.Add(Me.cmdCancel)
        Me.Controls.Add(Me.cmdPrintPreview)
        Me.Controls.Add(Me.cmdPrint)
        Me.Controls.Add(Me.fra)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "InvPoPrint"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Print POs"
        Me.fra.ResumeLayout(False)
        Me.fra.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents fra As GroupBox
    Friend WithEvents chkShowCost As CheckBox
    Friend WithEvents chkVendorSort As CheckBox
    Friend WithEvents cboLoc As ComboBox
    Friend WithEvents lblLabel2 As Label
    Friend WithEvents txtDate As DateTimePicker
    Friend WithEvents lblLabel As Label
    Friend WithEvents cmdPrint As Button
    Friend WithEvents cmdPrintPreview As Button
    Friend WithEvents cmdCancel As Button
    Friend WithEvents cboSortOrder As ComboBox
End Class
