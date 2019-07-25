<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class frmCashRegister
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
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
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.lblCust = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'lblCust
        '
        Me.lblCust.AutoSize = True
        Me.lblCust.Location = New System.Drawing.Point(17, 9)
        Me.lblCust.Name = "lblCust"
        Me.lblCust.Size = New System.Drawing.Size(39, 13)
        Me.lblCust.TabIndex = 0
        Me.lblCust.Text = "Label1"
        '
        'frmCashRegister
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(800, 450)
        Me.Controls.Add(Me.lblCust)
        Me.Name = "frmCashRegister"
        Me.Text = "frmCashRegister"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents lblCust As Label
End Class
