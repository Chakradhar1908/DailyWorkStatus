<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmDemoNotify
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
        Me.txtEnterLicense = New System.Windows.Forms.TextBox()
        Me.SuspendLayout()
        '
        'txtEnterLicense
        '
        Me.txtEnterLicense.Location = New System.Drawing.Point(348, 43)
        Me.txtEnterLicense.Name = "txtEnterLicense"
        Me.txtEnterLicense.Size = New System.Drawing.Size(100, 20)
        Me.txtEnterLicense.TabIndex = 0
        '
        'frmDemoNotify
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(800, 450)
        Me.Controls.Add(Me.txtEnterLicense)
        Me.Name = "frmDemoNotify"
        Me.Text = "frmDemoNotify"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents txtEnterLicense As TextBox
End Class
