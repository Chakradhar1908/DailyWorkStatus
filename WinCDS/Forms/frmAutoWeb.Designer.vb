<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmAutoWeb
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
        Me.txtSiteAddr = New System.Windows.Forms.TextBox()
        Me.SuspendLayout()
        '
        'txtSiteAddr
        '
        Me.txtSiteAddr.Location = New System.Drawing.Point(514, 141)
        Me.txtSiteAddr.Name = "txtSiteAddr"
        Me.txtSiteAddr.Size = New System.Drawing.Size(100, 20)
        Me.txtSiteAddr.TabIndex = 0
        '
        'frmAutoWeb
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(800, 450)
        Me.Controls.Add(Me.txtSiteAddr)
        Me.Name = "frmAutoWeb"
        Me.Text = "frmAutoWeb"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents txtSiteAddr As TextBox
End Class
