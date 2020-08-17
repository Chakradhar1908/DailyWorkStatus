<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class InvOrdStatus
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
        Me.cmdNextSale = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'cmdNextSale
        '
        Me.cmdNextSale.Location = New System.Drawing.Point(399, 104)
        Me.cmdNextSale.Name = "cmdNextSale"
        Me.cmdNextSale.Size = New System.Drawing.Size(75, 23)
        Me.cmdNextSale.TabIndex = 0
        Me.cmdNextSale.Text = "Button1"
        Me.cmdNextSale.UseVisualStyleBackColor = True
        '
        'InvOrdStatus
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(800, 450)
        Me.Controls.Add(Me.cmdNextSale)
        Me.Name = "InvOrdStatus"
        Me.Text = "InvOrdStatus"
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents cmdNextSale As Button
End Class
