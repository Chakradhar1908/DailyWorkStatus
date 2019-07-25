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
        Me.cmdOk = New System.Windows.Forms.Button()
        Me.lstSelection = New System.Windows.Forms.ListBox()
        Me.lstSelectionCheck = New System.Windows.Forms.CheckedListBox()
        Me.SuspendLayout()
        '
        'cmdOk
        '
        Me.cmdOk.Location = New System.Drawing.Point(0, 0)
        Me.cmdOk.Name = "cmdOk"
        Me.cmdOk.Size = New System.Drawing.Size(75, 23)
        Me.cmdOk.TabIndex = 0
        Me.cmdOk.Text = "&Print"
        Me.cmdOk.UseVisualStyleBackColor = True
        '
        'lstSelection
        '
        Me.lstSelection.FormattingEnabled = True
        Me.lstSelection.Location = New System.Drawing.Point(12, 99)
        Me.lstSelection.Name = "lstSelection"
        Me.lstSelection.Size = New System.Drawing.Size(120, 95)
        Me.lstSelection.TabIndex = 1
        '
        'lstSelectionCheck
        '
        Me.lstSelectionCheck.FormattingEnabled = True
        Me.lstSelectionCheck.Location = New System.Drawing.Point(142, 282)
        Me.lstSelectionCheck.Name = "lstSelectionCheck"
        Me.lstSelectionCheck.Size = New System.Drawing.Size(120, 94)
        Me.lstSelectionCheck.TabIndex = 2
        '
        'frmSelectOption
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(800, 450)
        Me.Controls.Add(Me.lstSelectionCheck)
        Me.Controls.Add(Me.lstSelection)
        Me.Controls.Add(Me.cmdOk)
        Me.Name = "frmSelectOption"
        Me.Text = "frmSelectOption"
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents cmdOk As Button
    Friend WithEvents lstSelection As ListBox
    Friend WithEvents lstSelectionCheck As CheckedListBox
End Class
