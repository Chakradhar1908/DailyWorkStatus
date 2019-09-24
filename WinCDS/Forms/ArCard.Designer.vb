<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class ArCard
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
        Me.lblTotalPayoff = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'lblTotalPayoff
        '
        Me.lblTotalPayoff.AutoSize = True
        Me.lblTotalPayoff.Location = New System.Drawing.Point(0, 0)
        Me.lblTotalPayoff.Name = "lblTotalPayoff"
        Me.lblTotalPayoff.Size = New System.Drawing.Size(71, 13)
        Me.lblTotalPayoff.TabIndex = 0
        Me.lblTotalPayoff.Text = "lblTotalPayoff"
        '
        'ArCard
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(800, 450)
        Me.Controls.Add(Me.lblTotalPayoff)
        Me.Name = "ArCard"
        Me.Text = "ArCardvb"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents lblTotalPayoff As Label
End Class
