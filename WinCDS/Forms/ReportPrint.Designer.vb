<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class ReportPrint
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
        Me.dteReportDate = New System.Windows.Forms.DateTimePicker()
        Me.fraOptions = New System.Windows.Forms.GroupBox()
        Me.chkLastPay = New System.Windows.Forms.CheckBox()
        Me.Opt5 = New System.Windows.Forms.RadioButton()
        Me.Opt3 = New System.Windows.Forms.RadioButton()
        Me.Opt4 = New System.Windows.Forms.RadioButton()
        Me.Opt1 = New System.Windows.Forms.RadioButton()
        Me.Opt2 = New System.Windows.Forms.RadioButton()
        Me.cmdPrint = New System.Windows.Forms.Button()
        Me.cmdPrintPreview = New System.Windows.Forms.Button()
        Me.cmdCancel = New System.Windows.Forms.Button()
        Me.lblSelectDate = New System.Windows.Forms.Label()
        Me.fra.SuspendLayout()
        Me.fraOptions.SuspendLayout()
        Me.SuspendLayout()
        '
        'fra
        '
        Me.fra.Controls.Add(Me.dteReportDate)
        Me.fra.Controls.Add(Me.lblSelectDate)
        Me.fra.Location = New System.Drawing.Point(8, 8)
        Me.fra.Name = "fra"
        Me.fra.Size = New System.Drawing.Size(234, 51)
        Me.fra.TabIndex = 0
        Me.fra.TabStop = False
        '
        'dteReportDate
        '
        Me.dteReportDate.CustomFormat = "MM/dd/yyyy"
        Me.dteReportDate.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dteReportDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dteReportDate.Location = New System.Drawing.Point(98, 13)
        Me.dteReportDate.Name = "dteReportDate"
        Me.dteReportDate.Size = New System.Drawing.Size(129, 26)
        Me.dteReportDate.TabIndex = 0
        '
        'fraOptions
        '
        Me.fraOptions.Controls.Add(Me.chkLastPay)
        Me.fraOptions.Controls.Add(Me.Opt5)
        Me.fraOptions.Controls.Add(Me.Opt3)
        Me.fraOptions.Controls.Add(Me.Opt4)
        Me.fraOptions.Controls.Add(Me.Opt1)
        Me.fraOptions.Controls.Add(Me.Opt2)
        Me.fraOptions.Location = New System.Drawing.Point(8, 91)
        Me.fraOptions.Name = "fraOptions"
        Me.fraOptions.Size = New System.Drawing.Size(234, 91)
        Me.fraOptions.TabIndex = 1
        Me.fraOptions.TabStop = False
        Me.fraOptions.Text = "Sort By"
        '
        'chkLastPay
        '
        Me.chkLastPay.AutoSize = True
        Me.chkLastPay.Location = New System.Drawing.Point(6, 65)
        Me.chkLastPay.Name = "chkLastPay"
        Me.chkLastPay.Size = New System.Drawing.Size(97, 17)
        Me.chkLastPay.TabIndex = 5
        Me.chkLastPay.Text = "Show Last Pa&y"
        Me.chkLastPay.UseVisualStyleBackColor = True
        '
        'Opt5
        '
        Me.Opt5.AutoSize = True
        Me.Opt5.Location = New System.Drawing.Point(129, 65)
        Me.Opt5.Name = "Opt5"
        Me.Opt5.Size = New System.Drawing.Size(72, 17)
        Me.Opt5.TabIndex = 4
        Me.Opt5.TabStop = True
        Me.Opt5.Text = "Sale &Date"
        Me.Opt5.UseVisualStyleBackColor = True
        '
        'Opt3
        '
        Me.Opt3.AutoSize = True
        Me.Opt3.Location = New System.Drawing.Point(129, 42)
        Me.Opt3.Name = "Opt3"
        Me.Opt3.Size = New System.Drawing.Size(58, 17)
        Me.Opt3.TabIndex = 3
        Me.Opt3.TabStop = True
        Me.Opt3.Text = "&Ageing"
        Me.Opt3.UseVisualStyleBackColor = True
        '
        'Opt4
        '
        Me.Opt4.AutoSize = True
        Me.Opt4.Location = New System.Drawing.Point(129, 19)
        Me.Opt4.Name = "Opt4"
        Me.Opt4.Size = New System.Drawing.Size(87, 17)
        Me.Opt4.TabIndex = 2
        Me.Opt4.TabStop = True
        Me.Opt4.Text = "Sa&les Person"
        Me.Opt4.UseVisualStyleBackColor = True
        '
        'Opt1
        '
        Me.Opt1.AutoSize = True
        Me.Opt1.Location = New System.Drawing.Point(6, 42)
        Me.Opt1.Name = "Opt1"
        Me.Opt1.Size = New System.Drawing.Size(63, 17)
        Me.Opt1.TabIndex = 1
        Me.Opt1.TabStop = True
        Me.Opt1.Text = "&Sale No"
        Me.Opt1.UseVisualStyleBackColor = True
        '
        'Opt2
        '
        Me.Opt2.AutoSize = True
        Me.Opt2.Location = New System.Drawing.Point(6, 19)
        Me.Opt2.Name = "Opt2"
        Me.Opt2.Size = New System.Drawing.Size(53, 17)
        Me.Opt2.TabIndex = 0
        Me.Opt2.TabStop = True
        Me.Opt2.Text = "&Name"
        Me.Opt2.UseVisualStyleBackColor = True
        '
        'cmdPrint
        '
        Me.cmdPrint.Location = New System.Drawing.Point(9, 188)
        Me.cmdPrint.Name = "cmdPrint"
        Me.cmdPrint.Size = New System.Drawing.Size(75, 60)
        Me.cmdPrint.TabIndex = 2
        Me.cmdPrint.Text = "&Print"
        Me.cmdPrint.UseVisualStyleBackColor = True
        '
        'cmdPrintPreview
        '
        Me.cmdPrintPreview.Location = New System.Drawing.Point(84, 188)
        Me.cmdPrintPreview.Name = "cmdPrintPreview"
        Me.cmdPrintPreview.Size = New System.Drawing.Size(77, 60)
        Me.cmdPrintPreview.TabIndex = 3
        Me.cmdPrintPreview.Text = "P&rint Preview"
        Me.cmdPrintPreview.UseVisualStyleBackColor = True
        '
        'cmdCancel
        '
        Me.cmdCancel.Location = New System.Drawing.Point(161, 188)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.Size = New System.Drawing.Size(75, 60)
        Me.cmdCancel.TabIndex = 4
        Me.cmdCancel.Text = "&Cancel"
        Me.cmdCancel.UseVisualStyleBackColor = True
        '
        'lblSelectDate
        '
        Me.lblSelectDate.Location = New System.Drawing.Point(6, 18)
        Me.lblSelectDate.Name = "lblSelectDate"
        Me.lblSelectDate.Size = New System.Drawing.Size(210, 23)
        Me.lblSelectDate.TabIndex = 1
        Me.lblSelectDate.Text = "Label1"
        '
        'ReportPrint
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(247, 251)
        Me.Controls.Add(Me.cmdCancel)
        Me.Controls.Add(Me.cmdPrintPreview)
        Me.Controls.Add(Me.cmdPrint)
        Me.Controls.Add(Me.fraOptions)
        Me.Controls.Add(Me.fra)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.MaximizeBox = False
        Me.Name = "ReportPrint"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "ReportPrint"
        Me.fra.ResumeLayout(False)
        Me.fraOptions.ResumeLayout(False)
        Me.fraOptions.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents fra As GroupBox
    Friend WithEvents dteReportDate As DateTimePicker
    Friend WithEvents fraOptions As GroupBox
    Friend WithEvents chkLastPay As CheckBox
    Friend WithEvents Opt5 As RadioButton
    Friend WithEvents Opt3 As RadioButton
    Friend WithEvents Opt4 As RadioButton
    Friend WithEvents Opt1 As RadioButton
    Friend WithEvents Opt2 As RadioButton
    Friend WithEvents cmdPrint As Button
    Friend WithEvents cmdPrintPreview As Button
    Friend WithEvents cmdCancel As Button
    Friend WithEvents lblSelectDate As Label
End Class
