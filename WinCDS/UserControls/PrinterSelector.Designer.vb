<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class PrinterSelector
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
        Me.Scroller = New System.Windows.Forms.VScrollBar()
        Me.optPrinter0 = New System.Windows.Forms.RadioButton()
        Me.SuspendLayout()
        '
        'Scroller
        '
        Me.Scroller.Location = New System.Drawing.Point(150, 26)
        Me.Scroller.Name = "Scroller"
        Me.Scroller.Size = New System.Drawing.Size(17, 95)
        Me.Scroller.TabIndex = 0
        '
        'optPrinter0
        '
        Me.optPrinter0.Location = New System.Drawing.Point(21, 6)
        Me.optPrinter0.Name = "optPrinter0"
        Me.optPrinter0.Size = New System.Drawing.Size(129, 27)
        Me.optPrinter0.TabIndex = 1
        Me.optPrinter0.TabStop = True
        Me.optPrinter0.Text = "(Printer Device Name)"
        Me.optPrinter0.UseVisualStyleBackColor = True
        Me.optPrinter0.Visible = False
        '
        'PrinterSelector
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.Controls.Add(Me.optPrinter0)
        Me.Controls.Add(Me.Scroller)
        Me.Name = "PrinterSelector"
        Me.Size = New System.Drawing.Size(182, 150)
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents Scroller As VScrollBar
    Friend WithEvents optPrinter0 As RadioButton
End Class
