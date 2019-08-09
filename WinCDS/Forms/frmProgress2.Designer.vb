<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmProgress2
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
        Me.fra = New System.Windows.Forms.GroupBox()
        Me.SuspendLayout()
        '
        'fra
        '
        Me.fra.Location = New System.Drawing.Point(0, 0)
        Me.fra.Name = "fra"
        Me.fra.Size = New System.Drawing.Size(200, 100)
        Me.fra.TabIndex = 0
        Me.fra.TabStop = False
        Me.fra.Text = "GroupBox1"
        '
        'frmProgress2
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(800, 450)
        Me.Controls.Add(Me.fra)
        Me.Name = "frmProgress2"
        Me.Text = "frmProgress2"
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents fra As GroupBox
End Class
