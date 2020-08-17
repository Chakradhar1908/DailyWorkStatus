<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmAdjustTax
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
        Me.txtGrossSale = New System.Windows.Forms.TextBox()
        Me.txtTaxable = New System.Windows.Forms.TextBox()
        Me.lstTaxes = New System.Windows.Forms.ListBox()
        Me.SuspendLayout()
        '
        'txtSaleNo
        '
        Me.txtSaleNo.Location = New System.Drawing.Point(0, 0)
        Me.txtSaleNo.Name = "txtSaleNo"
        Me.txtSaleNo.Size = New System.Drawing.Size(100, 20)
        Me.txtSaleNo.TabIndex = 0
        '
        'txtGrossSale
        '
        Me.txtGrossSale.Location = New System.Drawing.Point(350, 215)
        Me.txtGrossSale.Name = "txtGrossSale"
        Me.txtGrossSale.Size = New System.Drawing.Size(100, 20)
        Me.txtGrossSale.TabIndex = 1
        '
        'txtTaxable
        '
        Me.txtTaxable.Location = New System.Drawing.Point(337, 274)
        Me.txtTaxable.Name = "txtTaxable"
        Me.txtTaxable.Size = New System.Drawing.Size(100, 20)
        Me.txtTaxable.TabIndex = 2
        '
        'lstTaxes
        '
        Me.lstTaxes.FormattingEnabled = True
        Me.lstTaxes.Location = New System.Drawing.Point(360, 337)
        Me.lstTaxes.Name = "lstTaxes"
        Me.lstTaxes.Size = New System.Drawing.Size(120, 95)
        Me.lstTaxes.TabIndex = 3
        '
        'frmAdjustTax
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(800, 450)
        Me.Controls.Add(Me.lstTaxes)
        Me.Controls.Add(Me.txtTaxable)
        Me.Controls.Add(Me.txtGrossSale)
        Me.Controls.Add(Me.txtSaleNo)
        Me.Name = "frmAdjustTax"
        Me.Text = "frmAdjustTax"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents txtSaleNo As TextBox
    Friend WithEvents txtGrossSale As TextBox
    Friend WithEvents txtTaxable As TextBox
    Friend WithEvents lstTaxes As ListBox
End Class
