<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmBOSDiscount
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
        Me.txtDiscountType = New System.Windows.Forms.TextBox()
        Me.lblDiscountAmount = New System.Windows.Forms.Label()
        Me.txtDiscountAmount = New System.Windows.Forms.TextBox()
        Me.lblPercentSign = New System.Windows.Forms.Label()
        Me.fraDiscntTo = New System.Windows.Forms.GroupBox()
        Me.txtFlatRate = New System.Windows.Forms.TextBox()
        Me.txtLastNItems = New System.Windows.Forms.TextBox()
        Me.optFlatRate = New System.Windows.Forms.RadioButton()
        Me.optLastNItems = New System.Windows.Forms.RadioButton()
        Me.optCurrentItem = New System.Windows.Forms.RadioButton()
        Me.optAllItems = New System.Windows.Forms.RadioButton()
        Me.cmdOK = New System.Windows.Forms.Button()
        Me.cmdCancel = New System.Windows.Forms.Button()
        Me.cmbDiscountType = New System.Windows.Forms.ComboBox()
        Me.fraDiscntTo.SuspendLayout()
        Me.SuspendLayout()
        '
        'txtDiscountType
        '
        Me.txtDiscountType.Location = New System.Drawing.Point(10, 21)
        Me.txtDiscountType.Name = "txtDiscountType"
        Me.txtDiscountType.Size = New System.Drawing.Size(188, 20)
        Me.txtDiscountType.TabIndex = 0
        '
        'lblDiscountAmount
        '
        Me.lblDiscountAmount.AutoSize = True
        Me.lblDiscountAmount.Location = New System.Drawing.Point(14, 47)
        Me.lblDiscountAmount.Name = "lblDiscountAmount"
        Me.lblDiscountAmount.Size = New System.Drawing.Size(91, 13)
        Me.lblDiscountAmount.TabIndex = 1
        Me.lblDiscountAmount.Text = "Discount Amount:"
        '
        'txtDiscountAmount
        '
        Me.txtDiscountAmount.Location = New System.Drawing.Point(108, 47)
        Me.txtDiscountAmount.Name = "txtDiscountAmount"
        Me.txtDiscountAmount.Size = New System.Drawing.Size(61, 20)
        Me.txtDiscountAmount.TabIndex = 2
        Me.txtDiscountAmount.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'lblPercentSign
        '
        Me.lblPercentSign.AutoSize = True
        Me.lblPercentSign.Location = New System.Drawing.Point(169, 50)
        Me.lblPercentSign.Name = "lblPercentSign"
        Me.lblPercentSign.Size = New System.Drawing.Size(15, 13)
        Me.lblPercentSign.TabIndex = 3
        Me.lblPercentSign.Text = "%"
        '
        'fraDiscntTo
        '
        Me.fraDiscntTo.Controls.Add(Me.txtFlatRate)
        Me.fraDiscntTo.Controls.Add(Me.txtLastNItems)
        Me.fraDiscntTo.Controls.Add(Me.optFlatRate)
        Me.fraDiscntTo.Controls.Add(Me.optLastNItems)
        Me.fraDiscntTo.Controls.Add(Me.optCurrentItem)
        Me.fraDiscntTo.Controls.Add(Me.optAllItems)
        Me.fraDiscntTo.Location = New System.Drawing.Point(6, 74)
        Me.fraDiscntTo.Name = "fraDiscntTo"
        Me.fraDiscntTo.Size = New System.Drawing.Size(200, 116)
        Me.fraDiscntTo.TabIndex = 4
        Me.fraDiscntTo.TabStop = False
        Me.fraDiscntTo.Text = "Apply Discount To:"
        '
        'txtFlatRate
        '
        Me.txtFlatRate.Location = New System.Drawing.Point(123, 91)
        Me.txtFlatRate.Name = "txtFlatRate"
        Me.txtFlatRate.Size = New System.Drawing.Size(63, 20)
        Me.txtFlatRate.TabIndex = 5
        '
        'txtLastNItems
        '
        Me.txtLastNItems.Location = New System.Drawing.Point(88, 67)
        Me.txtLastNItems.Name = "txtLastNItems"
        Me.txtLastNItems.Size = New System.Drawing.Size(32, 20)
        Me.txtLastNItems.TabIndex = 4
        Me.txtLastNItems.Text = "1"
        '
        'optFlatRate
        '
        Me.optFlatRate.AutoSize = True
        Me.optFlatRate.Location = New System.Drawing.Point(45, 93)
        Me.optFlatRate.Name = "optFlatRate"
        Me.optFlatRate.Size = New System.Drawing.Size(71, 17)
        Me.optFlatRate.TabIndex = 3
        Me.optFlatRate.Text = "&Flat Rate:"
        Me.optFlatRate.UseVisualStyleBackColor = True
        '
        'optLastNItems
        '
        Me.optLastNItems.AutoSize = True
        Me.optLastNItems.Location = New System.Drawing.Point(45, 66)
        Me.optLastNItems.Name = "optLastNItems"
        Me.optLastNItems.Size = New System.Drawing.Size(118, 17)
        Me.optLastNItems.TabIndex = 2
        Me.optLastNItems.Text = "&Last              Item(s)"
        Me.optLastNItems.UseVisualStyleBackColor = True
        '
        'optCurrentItem
        '
        Me.optCurrentItem.AutoSize = True
        Me.optCurrentItem.Location = New System.Drawing.Point(45, 42)
        Me.optCurrentItem.Name = "optCurrentItem"
        Me.optCurrentItem.Size = New System.Drawing.Size(68, 17)
        Me.optCurrentItem.TabIndex = 1
        Me.optCurrentItem.Text = "Last &Item"
        Me.optCurrentItem.UseVisualStyleBackColor = True
        '
        'optAllItems
        '
        Me.optAllItems.AutoSize = True
        Me.optAllItems.Checked = True
        Me.optAllItems.Location = New System.Drawing.Point(45, 19)
        Me.optAllItems.Name = "optAllItems"
        Me.optAllItems.Size = New System.Drawing.Size(64, 17)
        Me.optAllItems.TabIndex = 0
        Me.optAllItems.TabStop = True
        Me.optAllItems.Text = "&All Items"
        Me.optAllItems.UseVisualStyleBackColor = True
        '
        'cmdOK
        '
        Me.cmdOK.Location = New System.Drawing.Point(26, 196)
        Me.cmdOK.Name = "cmdOK"
        Me.cmdOK.Size = New System.Drawing.Size(75, 56)
        Me.cmdOK.TabIndex = 5
        Me.cmdOK.Text = "&OK"
        Me.cmdOK.UseVisualStyleBackColor = True
        '
        'cmdCancel
        '
        Me.cmdCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cmdCancel.Location = New System.Drawing.Point(115, 196)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.Size = New System.Drawing.Size(75, 56)
        Me.cmdCancel.TabIndex = 6
        Me.cmdCancel.Text = "&Cancel"
        Me.cmdCancel.UseVisualStyleBackColor = True
        '
        'cmbDiscountType
        '
        Me.cmbDiscountType.FormattingEnabled = True
        Me.cmbDiscountType.Location = New System.Drawing.Point(12, 7)
        Me.cmbDiscountType.Name = "cmbDiscountType"
        Me.cmbDiscountType.Size = New System.Drawing.Size(188, 21)
        Me.cmbDiscountType.TabIndex = 7
        Me.cmbDiscountType.Text = "cmbDiscountType"
        '
        'frmBOSDiscount
        '
        Me.AcceptButton = Me.cmdOK
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.cmdCancel
        Me.ClientSize = New System.Drawing.Size(217, 254)
        Me.Controls.Add(Me.cmdCancel)
        Me.Controls.Add(Me.cmdOK)
        Me.Controls.Add(Me.fraDiscntTo)
        Me.Controls.Add(Me.lblPercentSign)
        Me.Controls.Add(Me.txtDiscountAmount)
        Me.Controls.Add(Me.lblDiscountAmount)
        Me.Controls.Add(Me.txtDiscountType)
        Me.Controls.Add(Me.cmbDiscountType)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmBOSDiscount"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Select Item Discount"
        Me.fraDiscntTo.ResumeLayout(False)
        Me.fraDiscntTo.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents txtDiscountType As TextBox
    Friend WithEvents lblDiscountAmount As Label
    Friend WithEvents txtDiscountAmount As TextBox
    Friend WithEvents lblPercentSign As Label
    Friend WithEvents fraDiscntTo As GroupBox
    Friend WithEvents txtFlatRate As TextBox
    Friend WithEvents txtLastNItems As TextBox
    Friend WithEvents optFlatRate As RadioButton
    Friend WithEvents optLastNItems As RadioButton
    Friend WithEvents optCurrentItem As RadioButton
    Friend WithEvents optAllItems As RadioButton
    Friend WithEvents cmdOK As Button
    Friend WithEvents cmdCancel As Button
    Friend WithEvents cmbDiscountType As ComboBox
End Class
