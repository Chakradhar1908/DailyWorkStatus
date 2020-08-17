<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmAdjustAdd
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
        Me.fraSelect = New System.Windows.Forms.GroupBox()
        Me.optCarpet = New System.Windows.Forms.RadioButton()
        Me.optNotes = New System.Windows.Forms.RadioButton()
        Me.optLabor = New System.Windows.Forms.RadioButton()
        Me.optDelivery = New System.Windows.Forms.RadioButton()
        Me.optStain = New System.Windows.Forms.RadioButton()
        Me.optEnterStyle = New System.Windows.Forms.RadioButton()
        Me.cmdOK = New System.Windows.Forms.Button()
        Me.cmdCancel = New System.Windows.Forms.Button()
        Me.fraSelect.SuspendLayout()
        Me.SuspendLayout()
        '
        'fraSelect
        '
        Me.fraSelect.Controls.Add(Me.optCarpet)
        Me.fraSelect.Controls.Add(Me.optNotes)
        Me.fraSelect.Controls.Add(Me.optLabor)
        Me.fraSelect.Controls.Add(Me.optDelivery)
        Me.fraSelect.Controls.Add(Me.optStain)
        Me.fraSelect.Controls.Add(Me.optEnterStyle)
        Me.fraSelect.Location = New System.Drawing.Point(6, 8)
        Me.fraSelect.Name = "fraSelect"
        Me.fraSelect.Size = New System.Drawing.Size(150, 129)
        Me.fraSelect.TabIndex = 0
        Me.fraSelect.TabStop = False
        Me.fraSelect.Text = "Select Item To Add:"
        '
        'optCarpet
        '
        Me.optCarpet.AutoSize = True
        Me.optCarpet.Location = New System.Drawing.Point(15, 107)
        Me.optCarpet.Name = "optCarpet"
        Me.optCarpet.Size = New System.Drawing.Size(56, 17)
        Me.optCarpet.TabIndex = 5
        Me.optCarpet.Text = "&Carpet"
        Me.optCarpet.UseVisualStyleBackColor = True
        '
        'optNotes
        '
        Me.optNotes.AutoSize = True
        Me.optNotes.Location = New System.Drawing.Point(15, 90)
        Me.optNotes.Name = "optNotes"
        Me.optNotes.Size = New System.Drawing.Size(53, 17)
        Me.optNotes.TabIndex = 4
        Me.optNotes.Text = "&Notes"
        Me.optNotes.UseVisualStyleBackColor = True
        '
        'optLabor
        '
        Me.optLabor.AutoSize = True
        Me.optLabor.Location = New System.Drawing.Point(15, 73)
        Me.optLabor.Name = "optLabor"
        Me.optLabor.Size = New System.Drawing.Size(92, 17)
        Me.optLabor.TabIndex = 3
        Me.optLabor.Text = "&Labor  Charge"
        Me.optLabor.UseVisualStyleBackColor = True
        '
        'optDelivery
        '
        Me.optDelivery.AutoSize = True
        Me.optDelivery.Location = New System.Drawing.Point(15, 56)
        Me.optDelivery.Name = "optDelivery"
        Me.optDelivery.Size = New System.Drawing.Size(100, 17)
        Me.optDelivery.TabIndex = 2
        Me.optDelivery.Text = "&Delivery Charge"
        Me.optDelivery.UseVisualStyleBackColor = True
        '
        'optStain
        '
        Me.optStain.AutoSize = True
        Me.optStain.Location = New System.Drawing.Point(15, 39)
        Me.optStain.Name = "optStain"
        Me.optStain.Size = New System.Drawing.Size(100, 17)
        Me.optStain.TabIndex = 1
        Me.optStain.Text = "&Stain Protection"
        Me.optStain.UseVisualStyleBackColor = True
        '
        'optEnterStyle
        '
        Me.optEnterStyle.AutoSize = True
        Me.optEnterStyle.Checked = True
        Me.optEnterStyle.Location = New System.Drawing.Point(15, 22)
        Me.optEnterStyle.Name = "optEnterStyle"
        Me.optEnterStyle.Size = New System.Drawing.Size(116, 17)
        Me.optEnterStyle.TabIndex = 0
        Me.optEnterStyle.TabStop = True
        Me.optEnterStyle.Text = "&Enter Style Number"
        Me.optEnterStyle.UseVisualStyleBackColor = True
        '
        'cmdOK
        '
        Me.cmdOK.Location = New System.Drawing.Point(6, 143)
        Me.cmdOK.Name = "cmdOK"
        Me.cmdOK.Size = New System.Drawing.Size(150, 23)
        Me.cmdOK.TabIndex = 1
        Me.cmdOK.Text = "&OK"
        Me.cmdOK.UseVisualStyleBackColor = True
        '
        'cmdCancel
        '
        Me.cmdCancel.Location = New System.Drawing.Point(46, 143)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.Size = New System.Drawing.Size(56, 23)
        Me.cmdCancel.TabIndex = 1
        Me.cmdCancel.Text = "Cancel"
        Me.cmdCancel.UseVisualStyleBackColor = True
        '
        'frmAdjustAdd
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(162, 172)
        Me.Controls.Add(Me.cmdOK)
        Me.Controls.Add(Me.cmdCancel)
        Me.Controls.Add(Me.fraSelect)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmAdjustAdd"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Add Item"
        Me.fraSelect.ResumeLayout(False)
        Me.fraSelect.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents fraSelect As GroupBox
    Friend WithEvents optCarpet As RadioButton
    Friend WithEvents optNotes As RadioButton
    Friend WithEvents optLabor As RadioButton
    Friend WithEvents optDelivery As RadioButton
    Friend WithEvents optStain As RadioButton
    Friend WithEvents optEnterStyle As RadioButton
    Friend WithEvents cmdOK As Button
    Friend WithEvents cmdCancel As Button
End Class
