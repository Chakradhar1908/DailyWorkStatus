<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmSalesList
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
        Me.lstSalesmen = New System.Windows.Forms.ListBox()
        Me.fraButtons = New System.Windows.Forms.GroupBox()
        Me.cmdClear = New System.Windows.Forms.Button()
        Me.cmdApply = New System.Windows.Forms.Button()
        Me.fraButtons.SuspendLayout()
        Me.SuspendLayout()
        '
        'lstSalesmen
        '
        Me.lstSalesmen.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lstSalesmen.FormattingEnabled = True
        Me.lstSalesmen.ItemHeight = 20
        Me.lstSalesmen.Location = New System.Drawing.Point(8, 2)
        Me.lstSalesmen.Name = "lstSalesmen"
        Me.lstSalesmen.Size = New System.Drawing.Size(175, 124)
        Me.lstSalesmen.TabIndex = 0
        '
        'fraButtons
        '
        Me.fraButtons.Controls.Add(Me.cmdClear)
        Me.fraButtons.Controls.Add(Me.cmdApply)
        Me.fraButtons.Location = New System.Drawing.Point(8, 144)
        Me.fraButtons.Name = "fraButtons"
        Me.fraButtons.Size = New System.Drawing.Size(174, 54)
        Me.fraButtons.TabIndex = 1
        Me.fraButtons.TabStop = False
        '
        'cmdClear
        '
        Me.cmdClear.Location = New System.Drawing.Point(87, 15)
        Me.cmdClear.Name = "cmdClear"
        Me.cmdClear.Size = New System.Drawing.Size(75, 34)
        Me.cmdClear.TabIndex = 1
        Me.cmdClear.Text = "&Clear"
        Me.cmdClear.UseVisualStyleBackColor = True
        '
        'cmdApply
        '
        Me.cmdApply.Location = New System.Drawing.Point(6, 15)
        Me.cmdApply.Name = "cmdApply"
        Me.cmdApply.Size = New System.Drawing.Size(75, 34)
        Me.cmdApply.TabIndex = 0
        Me.cmdApply.Text = "&Apply"
        Me.cmdApply.UseVisualStyleBackColor = True
        '
        'frmSalesList
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(189, 204)
        Me.Controls.Add(Me.fraButtons)
        Me.Controls.Add(Me.lstSalesmen)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmSalesList"
        Me.Text = "Sales Staff"
        Me.fraButtons.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents lstSalesmen As ListBox
    Friend WithEvents fraButtons As GroupBox
    Friend WithEvents cmdClear As Button
    Friend WithEvents cmdApply As Button
End Class
