<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class MailBookEmail
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
        Me.txtBodyFile = New System.Windows.Forms.TextBox()
        Me.SuspendLayout()
        '
        'txtBodyFile
        '
        Me.txtBodyFile.Location = New System.Drawing.Point(0, 0)
        Me.txtBodyFile.Name = "txtBodyFile"
        Me.txtBodyFile.Size = New System.Drawing.Size(100, 20)
        Me.txtBodyFile.TabIndex = 0
        '
        'MailBookEmail
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(800, 450)
        Me.Controls.Add(Me.txtBodyFile)
        Me.Name = "MailBookEmail"
        Me.Text = "MailBookEmail"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents txtBodyFile As TextBox
End Class
