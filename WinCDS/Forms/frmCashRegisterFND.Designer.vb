<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmCashRegisterFND
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
        Me.lblStyle = New System.Windows.Forms.Label()
        Me.txtPrice = New System.Windows.Forms.TextBox()
        Me.cmbVendor = New System.Windows.Forms.ComboBox()
        Me.txtDesc = New System.Windows.Forms.TextBox()
        Me.fraFND = New System.Windows.Forms.GroupBox()
        Me.lblPrice = New System.Windows.Forms.Label()
        Me.lblVendor = New System.Windows.Forms.Label()
        Me.lblDesc = New System.Windows.Forms.Label()
        Me.cmdOK = New System.Windows.Forms.Button()
        Me.cmdCancel = New System.Windows.Forms.Button()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.fraFND.SuspendLayout()
        Me.SuspendLayout()
        '
        'lblStyle
        '
        Me.lblStyle.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblStyle.Font = New System.Drawing.Font("Arial Black", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblStyle.Location = New System.Drawing.Point(6, 19)
        Me.lblStyle.Name = "lblStyle"
        Me.lblStyle.Size = New System.Drawing.Size(233, 23)
        Me.lblStyle.TabIndex = 0
        Me.lblStyle.Text = "#################"
        Me.lblStyle.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'txtPrice
        '
        Me.txtPrice.Location = New System.Drawing.Point(79, 56)
        Me.txtPrice.Name = "txtPrice"
        Me.txtPrice.Size = New System.Drawing.Size(83, 20)
        Me.txtPrice.TabIndex = 1
        '
        'cmbVendor
        '
        Me.cmbVendor.FormattingEnabled = True
        Me.cmbVendor.Location = New System.Drawing.Point(79, 82)
        Me.cmbVendor.Name = "cmbVendor"
        Me.cmbVendor.Size = New System.Drawing.Size(151, 21)
        Me.cmbVendor.TabIndex = 2
        '
        'txtDesc
        '
        Me.txtDesc.Location = New System.Drawing.Point(79, 107)
        Me.txtDesc.Name = "txtDesc"
        Me.txtDesc.Size = New System.Drawing.Size(151, 20)
        Me.txtDesc.TabIndex = 3
        '
        'fraFND
        '
        Me.fraFND.Controls.Add(Me.cmdCancel)
        Me.fraFND.Controls.Add(Me.cmdOK)
        Me.fraFND.Controls.Add(Me.lblDesc)
        Me.fraFND.Controls.Add(Me.cmbVendor)
        Me.fraFND.Controls.Add(Me.txtDesc)
        Me.fraFND.Controls.Add(Me.lblVendor)
        Me.fraFND.Controls.Add(Me.lblPrice)
        Me.fraFND.Controls.Add(Me.txtPrice)
        Me.fraFND.Controls.Add(Me.lblStyle)
        Me.fraFND.Location = New System.Drawing.Point(6, 6)
        Me.fraFND.Name = "fraFND"
        Me.fraFND.Size = New System.Drawing.Size(245, 190)
        Me.fraFND.TabIndex = 4
        Me.fraFND.TabStop = False
        Me.fraFND.Text = "Enter Found Item Information:"
        '
        'lblPrice
        '
        Me.lblPrice.AutoSize = True
        Me.lblPrice.Location = New System.Drawing.Point(39, 56)
        Me.lblPrice.Name = "lblPrice"
        Me.lblPrice.Size = New System.Drawing.Size(34, 13)
        Me.lblPrice.TabIndex = 1
        Me.lblPrice.Text = "&Price:"
        '
        'lblVendor
        '
        Me.lblVendor.AutoSize = True
        Me.lblVendor.Location = New System.Drawing.Point(31, 84)
        Me.lblVendor.Name = "lblVendor"
        Me.lblVendor.Size = New System.Drawing.Size(44, 13)
        Me.lblVendor.TabIndex = 2
        Me.lblVendor.Text = "&Vendor:"
        '
        'lblDesc
        '
        Me.lblDesc.AutoSize = True
        Me.lblDesc.Location = New System.Drawing.Point(38, 109)
        Me.lblDesc.Name = "lblDesc"
        Me.lblDesc.Size = New System.Drawing.Size(35, 13)
        Me.lblDesc.TabIndex = 3
        Me.lblDesc.Text = "&Desc:"
        '
        'cmdOK
        '
        Me.cmdOK.Location = New System.Drawing.Point(41, 138)
        Me.cmdOK.Name = "cmdOK"
        Me.cmdOK.Size = New System.Drawing.Size(75, 46)
        Me.cmdOK.TabIndex = 4
        Me.cmdOK.Text = "&OK"
        Me.ToolTip1.SetToolTip(Me.cmdOK, "Click here to add this item to the sale.")
        Me.cmdOK.UseVisualStyleBackColor = True
        '
        'cmdCancel
        '
        Me.cmdCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cmdCancel.Location = New System.Drawing.Point(134, 138)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.Size = New System.Drawing.Size(75, 46)
        Me.cmdCancel.TabIndex = 5
        Me.cmdCancel.Text = "&Cancel"
        Me.ToolTip1.SetToolTip(Me.cmdCancel, "Click to cancel this item.  It will not be added to the sale.")
        Me.cmdCancel.UseVisualStyleBackColor = True
        '
        'frmCashRegisterFND
        '
        Me.AcceptButton = Me.cmdOK
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.cmdCancel
        Me.ClientSize = New System.Drawing.Size(254, 200)
        Me.Controls.Add(Me.fraFND)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow
        Me.Name = "frmCashRegisterFND"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "FND Item - Cash Register"
        Me.fraFND.ResumeLayout(False)
        Me.fraFND.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents lblStyle As Label
    Friend WithEvents txtPrice As TextBox
    Friend WithEvents cmbVendor As ComboBox
    Friend WithEvents txtDesc As TextBox
    Friend WithEvents fraFND As GroupBox
    Friend WithEvents cmdCancel As Button
    Friend WithEvents ToolTip1 As ToolTip
    Friend WithEvents cmdOK As Button
    Friend WithEvents lblDesc As Label
    Friend WithEvents lblVendor As Label
    Friend WithEvents lblPrice As Label
End Class
