<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmSelectOption
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
        Me.lstSelection = New System.Windows.Forms.ListBox()
        Me.optSelection = New System.Windows.Forms.RadioButton()
        Me.lstSelectionCheck = New System.Windows.Forms.CheckedListBox()
        Me.cmdOk = New System.Windows.Forms.Button()
        Me.cmdCancel = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'lstSelection
        '
        Me.lstSelection.Font = New System.Drawing.Font("Lucida Console", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lstSelection.FormattingEnabled = True
        Me.lstSelection.ItemHeight = 11
        Me.lstSelection.Location = New System.Drawing.Point(72, 10)
        Me.lstSelection.Name = "lstSelection"
        Me.lstSelection.Size = New System.Drawing.Size(219, 15)
        Me.lstSelection.TabIndex = 0
        Me.lstSelection.Visible = False
        '
        'optSelection
        '
        Me.optSelection.AutoSize = True
        Me.optSelection.Font = New System.Drawing.Font("Lucida Console", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.optSelection.Location = New System.Drawing.Point(72, 67)
        Me.optSelection.Name = "optSelection"
        Me.optSelection.Size = New System.Drawing.Size(121, 15)
        Me.optSelection.TabIndex = 1
        Me.optSelection.TabStop = True
        Me.optSelection.Text = "&Generic Option"
        Me.optSelection.UseVisualStyleBackColor = True
        Me.optSelection.Visible = False
        '
        'lstSelectionCheck
        '
        Me.lstSelectionCheck.FormattingEnabled = True
        Me.lstSelectionCheck.Location = New System.Drawing.Point(72, 109)
        Me.lstSelectionCheck.Name = "lstSelectionCheck"
        Me.lstSelectionCheck.Size = New System.Drawing.Size(219, 19)
        Me.lstSelectionCheck.TabIndex = 2
        Me.lstSelectionCheck.Visible = False
        '
        'cmdOk
        '
        Me.cmdOk.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdOk.Location = New System.Drawing.Point(12, 173)
        Me.cmdOk.Name = "cmdOk"
        Me.cmdOk.Size = New System.Drawing.Size(75, 65)
        Me.cmdOk.TabIndex = 3
        Me.cmdOk.Text = "&Print"
        Me.cmdOk.UseVisualStyleBackColor = True
        '
        'cmdCancel
        '
        Me.cmdCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cmdCancel.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdCancel.Location = New System.Drawing.Point(95, 173)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.Size = New System.Drawing.Size(75, 65)
        Me.cmdCancel.TabIndex = 4
        Me.cmdCancel.Text = "&Cancel"
        Me.cmdCancel.UseVisualStyleBackColor = True
        '
        'frmSelectOption
        '
        Me.AcceptButton = Me.cmdOk
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.cmdCancel
        Me.ClientSize = New System.Drawing.Size(374, 392)
        Me.Controls.Add(Me.cmdCancel)
        Me.Controls.Add(Me.cmdOk)
        Me.Controls.Add(Me.lstSelectionCheck)
        Me.Controls.Add(Me.optSelection)
        Me.Controls.Add(Me.lstSelection)
        Me.Name = "frmSelectOption"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Select Option"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents lstSelection As ListBox
    Friend WithEvents optSelection As RadioButton
    Friend WithEvents lstSelectionCheck As CheckedListBox
    Friend WithEvents cmdOk As Button
    Friend WithEvents cmdCancel As Button
End Class
