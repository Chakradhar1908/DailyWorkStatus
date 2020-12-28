<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class ServiceManualItem
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
        Me.txtSaleNo = New System.Windows.Forms.TextBox()
        Me.txtStyle = New System.Windows.Forms.TextBox()
        Me.txtDesc = New System.Windows.Forms.TextBox()
        Me.txtQuantity = New System.Windows.Forms.TextBox()
        Me.cboVendor = New System.Windows.Forms.ComboBox()
        Me.dtpDelDate = New System.Windows.Forms.DateTimePicker()
        Me.lblSaleNo = New System.Windows.Forms.Label()
        Me.lblStyle = New System.Windows.Forms.Label()
        Me.lblQuantity = New System.Windows.Forms.Label()
        Me.lblDesc = New System.Windows.Forms.Label()
        Me.lblVendor = New System.Windows.Forms.Label()
        Me.lblDelDate = New System.Windows.Forms.Label()
        Me.fra = New System.Windows.Forms.GroupBox()
        Me.cmdCancel = New System.Windows.Forms.Button()
        Me.cmdOK = New System.Windows.Forms.Button()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.fra.SuspendLayout()
        Me.SuspendLayout()
        '
        'txtSaleNo
        '
        Me.txtSaleNo.Location = New System.Drawing.Point(63, 14)
        Me.txtSaleNo.Name = "txtSaleNo"
        Me.txtSaleNo.Size = New System.Drawing.Size(100, 20)
        Me.txtSaleNo.TabIndex = 0
        '
        'txtStyle
        '
        Me.txtStyle.Location = New System.Drawing.Point(63, 40)
        Me.txtStyle.Name = "txtStyle"
        Me.txtStyle.Size = New System.Drawing.Size(100, 20)
        Me.txtStyle.TabIndex = 1
        '
        'txtDesc
        '
        Me.txtDesc.Location = New System.Drawing.Point(63, 92)
        Me.txtDesc.Name = "txtDesc"
        Me.txtDesc.Size = New System.Drawing.Size(200, 20)
        Me.txtDesc.TabIndex = 2
        '
        'txtQuantity
        '
        Me.txtQuantity.Location = New System.Drawing.Point(63, 66)
        Me.txtQuantity.Name = "txtQuantity"
        Me.txtQuantity.Size = New System.Drawing.Size(37, 20)
        Me.txtQuantity.TabIndex = 3
        '
        'cboVendor
        '
        Me.cboVendor.FormattingEnabled = True
        Me.cboVendor.Location = New System.Drawing.Point(63, 119)
        Me.cboVendor.Name = "cboVendor"
        Me.cboVendor.Size = New System.Drawing.Size(121, 21)
        Me.cboVendor.TabIndex = 5
        '
        'dtpDelDate
        '
        Me.dtpDelDate.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dtpDelDate.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtpDelDate.Location = New System.Drawing.Point(63, 146)
        Me.dtpDelDate.Name = "dtpDelDate"
        Me.dtpDelDate.Size = New System.Drawing.Size(121, 22)
        Me.dtpDelDate.TabIndex = 6
        '
        'lblSaleNo
        '
        Me.lblSaleNo.AutoSize = True
        Me.lblSaleNo.Location = New System.Drawing.Point(9, 14)
        Me.lblSaleNo.Name = "lblSaleNo"
        Me.lblSaleNo.Size = New System.Drawing.Size(48, 13)
        Me.lblSaleNo.TabIndex = 7
        Me.lblSaleNo.Text = "Sale N&o:"
        '
        'lblStyle
        '
        Me.lblStyle.AutoSize = True
        Me.lblStyle.Location = New System.Drawing.Point(9, 40)
        Me.lblStyle.Name = "lblStyle"
        Me.lblStyle.Size = New System.Drawing.Size(33, 13)
        Me.lblStyle.TabIndex = 8
        Me.lblStyle.Text = "St&yle:"
        '
        'lblQuantity
        '
        Me.lblQuantity.AutoSize = True
        Me.lblQuantity.Location = New System.Drawing.Point(9, 69)
        Me.lblQuantity.Name = "lblQuantity"
        Me.lblQuantity.Size = New System.Drawing.Size(49, 13)
        Me.lblQuantity.TabIndex = 9
        Me.lblQuantity.Text = "Qu&antity:"
        '
        'lblDesc
        '
        Me.lblDesc.AutoSize = True
        Me.lblDesc.Location = New System.Drawing.Point(9, 92)
        Me.lblDesc.Name = "lblDesc"
        Me.lblDesc.Size = New System.Drawing.Size(35, 13)
        Me.lblDesc.TabIndex = 10
        Me.lblDesc.Text = "D&esc;"
        '
        'lblVendor
        '
        Me.lblVendor.AutoSize = True
        Me.lblVendor.Location = New System.Drawing.Point(9, 119)
        Me.lblVendor.Name = "lblVendor"
        Me.lblVendor.Size = New System.Drawing.Size(44, 13)
        Me.lblVendor.TabIndex = 11
        Me.lblVendor.Text = "&Vendor:"
        '
        'lblDelDate
        '
        Me.lblDelDate.AutoSize = True
        Me.lblDelDate.Location = New System.Drawing.Point(9, 146)
        Me.lblDelDate.Name = "lblDelDate"
        Me.lblDelDate.Size = New System.Drawing.Size(52, 13)
        Me.lblDelDate.TabIndex = 12
        Me.lblDelDate.Text = "Del &Date:"
        '
        'fra
        '
        Me.fra.Controls.Add(Me.txtQuantity)
        Me.fra.Controls.Add(Me.lblDelDate)
        Me.fra.Controls.Add(Me.txtSaleNo)
        Me.fra.Controls.Add(Me.lblVendor)
        Me.fra.Controls.Add(Me.txtStyle)
        Me.fra.Controls.Add(Me.lblDesc)
        Me.fra.Controls.Add(Me.txtDesc)
        Me.fra.Controls.Add(Me.lblQuantity)
        Me.fra.Controls.Add(Me.cboVendor)
        Me.fra.Controls.Add(Me.lblStyle)
        Me.fra.Controls.Add(Me.dtpDelDate)
        Me.fra.Controls.Add(Me.lblSaleNo)
        Me.fra.Location = New System.Drawing.Point(7, 4)
        Me.fra.Name = "fra"
        Me.fra.Size = New System.Drawing.Size(277, 174)
        Me.fra.TabIndex = 13
        Me.fra.TabStop = False
        '
        'cmdCancel
        '
        Me.cmdCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cmdCancel.Location = New System.Drawing.Point(119, 184)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.Size = New System.Drawing.Size(51, 50)
        Me.cmdCancel.TabIndex = 16
        Me.cmdCancel.Text = "&Cancel"
        Me.ToolTip1.SetToolTip(Me.cmdCancel, "Return from this screen.")
        Me.cmdCancel.UseVisualStyleBackColor = True
        '
        'cmdOK
        '
        Me.cmdOK.Location = New System.Drawing.Point(62, 184)
        Me.cmdOK.Name = "cmdOK"
        Me.cmdOK.Size = New System.Drawing.Size(51, 50)
        Me.cmdOK.TabIndex = 15
        Me.cmdOK.Text = "&OK"
        Me.ToolTip1.SetToolTip(Me.cmdOK, "Save the details of this form.")
        Me.cmdOK.UseVisualStyleBackColor = True
        '
        'ServiceManualItem
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.cmdCancel
        Me.ClientSize = New System.Drawing.Size(291, 240)
        Me.Controls.Add(Me.cmdCancel)
        Me.Controls.Add(Me.cmdOK)
        Me.Controls.Add(Me.fra)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "ServiceManualItem"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Manually Entered Item"
        Me.fra.ResumeLayout(False)
        Me.fra.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents txtSaleNo As TextBox
    Friend WithEvents txtStyle As TextBox
    Friend WithEvents txtDesc As TextBox
    Friend WithEvents txtQuantity As TextBox
    Friend WithEvents cboVendor As ComboBox
    Friend WithEvents dtpDelDate As DateTimePicker
    Friend WithEvents lblSaleNo As Label
    Friend WithEvents lblStyle As Label
    Friend WithEvents lblQuantity As Label
    Friend WithEvents lblDesc As Label
    Friend WithEvents lblVendor As Label
    Friend WithEvents lblDelDate As Label
    Friend WithEvents fra As GroupBox
    Friend WithEvents cmdCancel As Button
    Friend WithEvents ToolTip1 As ToolTip
    Friend WithEvents cmdOK As Button
End Class
