<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class OnScreenReport
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
        Me.cmdNext = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'cmdNext
        '
        Me.cmdNext.Location = New System.Drawing.Point(346, 27)
        Me.cmdNext.Name = "cmdNext"
        Me.cmdNext.Size = New System.Drawing.Size(75, 23)
        Me.cmdNext.TabIndex = 0
        Me.cmdNext.Text = "Button1"
        Me.cmdNext.UseVisualStyleBackColor = True
        '
        'OnScreenReport
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(800, 450)
        Me.Controls.Add(Me.cmdNext)
        Me.Name = "OnScreenReport"
        Me.Text = "OnScreenReport"
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents cmdNext As Button
End Class
