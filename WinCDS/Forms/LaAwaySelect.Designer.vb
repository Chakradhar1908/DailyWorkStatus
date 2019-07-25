<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class LaAwaySelect
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
        Me.opt30 = New System.Windows.Forms.RadioButton()
        Me.opt60 = New System.Windows.Forms.RadioButton()
        Me.opt90 = New System.Windows.Forms.RadioButton()
        Me.opt120 = New System.Windows.Forms.RadioButton()
        Me.cmdApply = New System.Windows.Forms.Button()
        Me.fra.SuspendLayout()
        Me.SuspendLayout()
        '
        'fra
        '
        Me.fra.Controls.Add(Me.opt120)
        Me.fra.Controls.Add(Me.opt90)
        Me.fra.Controls.Add(Me.opt60)
        Me.fra.Controls.Add(Me.opt30)
        Me.fra.Location = New System.Drawing.Point(12, 12)
        Me.fra.Name = "fra"
        Me.fra.Size = New System.Drawing.Size(164, 119)
        Me.fra.TabIndex = 0
        Me.fra.TabStop = False
        Me.fra.Text = "Pick Option "
        '
        'opt30
        '
        Me.opt30.AutoSize = True
        Me.opt30.Location = New System.Drawing.Point(11, 20)
        Me.opt30.Name = "opt30"
        Me.opt30.Size = New System.Drawing.Size(104, 17)
        Me.opt30.TabIndex = 0
        Me.opt30.TabStop = True
        Me.opt30.Text = "&30 Day Layaway"
        Me.opt30.UseVisualStyleBackColor = True
        '
        'opt60
        '
        Me.opt60.AutoSize = True
        Me.opt60.Location = New System.Drawing.Point(11, 43)
        Me.opt60.Name = "opt60"
        Me.opt60.Size = New System.Drawing.Size(104, 17)
        Me.opt60.TabIndex = 1
        Me.opt60.TabStop = True
        Me.opt60.Text = "&60 Day Layaway"
        Me.opt60.UseVisualStyleBackColor = True
        '
        'opt90
        '
        Me.opt90.AutoSize = True
        Me.opt90.Location = New System.Drawing.Point(11, 66)
        Me.opt90.Name = "opt90"
        Me.opt90.Size = New System.Drawing.Size(104, 17)
        Me.opt90.TabIndex = 2
        Me.opt90.TabStop = True
        Me.opt90.Text = "&90 Day Layaway"
        Me.opt90.UseVisualStyleBackColor = True
        '
        'opt120
        '
        Me.opt120.AutoSize = True
        Me.opt120.Location = New System.Drawing.Point(11, 89)
        Me.opt120.Name = "opt120"
        Me.opt120.Size = New System.Drawing.Size(110, 17)
        Me.opt120.TabIndex = 3
        Me.opt120.TabStop = True
        Me.opt120.Text = "&120 Day Layaway"
        Me.opt120.UseVisualStyleBackColor = True
        '
        'cmdApply
        '
        Me.cmdApply.Location = New System.Drawing.Point(39, 137)
        Me.cmdApply.Name = "cmdApply"
        Me.cmdApply.Size = New System.Drawing.Size(94, 23)
        Me.cmdApply.TabIndex = 1
        Me.cmdApply.Text = "&Post Time Limit"
        Me.cmdApply.UseVisualStyleBackColor = True
        '
        'LaAwaySelect
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(185, 163)
        Me.Controls.Add(Me.cmdApply)
        Me.Controls.Add(Me.fra)
        Me.Name = "LaAwaySelect"
        Me.Text = "Select Time Frame"
        Me.fra.ResumeLayout(False)
        Me.fra.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents fra As GroupBox
    Friend WithEvents opt120 As RadioButton
    Friend WithEvents opt90 As RadioButton
    Friend WithEvents opt60 As RadioButton
    Friend WithEvents opt30 As RadioButton
    Friend WithEvents cmdApply As Button
End Class
