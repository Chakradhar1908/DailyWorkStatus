<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class RichTextBoxNew
    Inherits System.Windows.Forms.UserControl

    'UserControl overrides dispose to clean up the component list.
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
        Me.mRichTextBox = New System.Windows.Forms.RichTextBox()
        Me.SuspendLayout()
        '
        'mRichTextBox
        '
        Me.mRichTextBox.Location = New System.Drawing.Point(0, 0)
        Me.mRichTextBox.Name = "mRichTextBox"
        Me.mRichTextBox.Size = New System.Drawing.Size(147, 137)
        Me.mRichTextBox.TabIndex = 0
        Me.mRichTextBox.Text = ""
        '
        'RichTextBoxNew
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.Controls.Add(Me.mRichTextBox)
        Me.Name = "RichTextBoxNew"
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents mRichTextBox As RichTextBox
End Class
