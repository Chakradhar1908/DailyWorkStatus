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
        Me.txtSaleNo = New System.Windows.Forms.TextBox()
        Me.txtStyle = New System.Windows.Forms.TextBox()
        Me.txtDesc = New System.Windows.Forms.TextBox()
        Me.txtQuantity = New System.Windows.Forms.TextBox()
        Me.TextBox5 = New System.Windows.Forms.TextBox()
        Me.cboVendor = New System.Windows.Forms.ComboBox()
        Me.dtpDelDate = New System.Windows.Forms.DateTimePicker()
        Me.SuspendLayout()
        '
        'txtSaleNo
        '
        Me.txtSaleNo.Location = New System.Drawing.Point(0, 0)
        Me.txtSaleNo.Name = "txtSaleNo"
        Me.txtSaleNo.Size = New System.Drawing.Size(100, 20)
        Me.txtSaleNo.TabIndex = 0
        '
        'txtStyle
        '
        Me.txtStyle.Location = New System.Drawing.Point(12, 38)
        Me.txtStyle.Name = "txtStyle"
        Me.txtStyle.Size = New System.Drawing.Size(100, 20)
        Me.txtStyle.TabIndex = 1
        '
        'txtDesc
        '
        Me.txtDesc.Location = New System.Drawing.Point(12, 77)
        Me.txtDesc.Name = "txtDesc"
        Me.txtDesc.Size = New System.Drawing.Size(100, 20)
        Me.txtDesc.TabIndex = 2
        '
        'txtQuantity
        '
        Me.txtQuantity.Location = New System.Drawing.Point(12, 117)
        Me.txtQuantity.Name = "txtQuantity"
        Me.txtQuantity.Size = New System.Drawing.Size(100, 20)
        Me.txtQuantity.TabIndex = 3
        '
        'TextBox5
        '
        Me.TextBox5.Location = New System.Drawing.Point(380, 301)
        Me.TextBox5.Name = "TextBox5"
        Me.TextBox5.Size = New System.Drawing.Size(100, 20)
        Me.TextBox5.TabIndex = 4
        '
        'cboVendor
        '
        Me.cboVendor.FormattingEnabled = True
        Me.cboVendor.Location = New System.Drawing.Point(12, 169)
        Me.cboVendor.Name = "cboVendor"
        Me.cboVendor.Size = New System.Drawing.Size(121, 21)
        Me.cboVendor.TabIndex = 5
        '
        'dtpDelDate
        '
        Me.dtpDelDate.Location = New System.Drawing.Point(12, 215)
        Me.dtpDelDate.Name = "dtpDelDate"
        Me.dtpDelDate.Size = New System.Drawing.Size(200, 20)
        Me.dtpDelDate.TabIndex = 6
        '
        'ServiceManualItem
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(800, 450)
        Me.Controls.Add(Me.dtpDelDate)
        Me.Controls.Add(Me.cboVendor)
        Me.Controls.Add(Me.TextBox5)
        Me.Controls.Add(Me.txtQuantity)
        Me.Controls.Add(Me.txtDesc)
        Me.Controls.Add(Me.txtStyle)
        Me.Controls.Add(Me.txtSaleNo)
        Me.Name = "ServiceManualItem"
        Me.Text = "ServiceManualItem"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents txtSaleNo As TextBox
    Friend WithEvents txtStyle As TextBox
    Friend WithEvents txtDesc As TextBox
    Friend WithEvents txtQuantity As TextBox
    Friend WithEvents TextBox5 As TextBox
    Friend WithEvents cboVendor As ComboBox
    Friend WithEvents dtpDelDate As DateTimePicker
End Class
