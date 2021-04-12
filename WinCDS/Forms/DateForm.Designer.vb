<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class DateForm
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
        Me.fraDates = New System.Windows.Forms.GroupBox()
        Me.lblDate1 = New System.Windows.Forms.Label()
        Me.lblNote = New System.Windows.Forms.Label()
        Me.toDate = New System.Windows.Forms.DateTimePicker()
        Me.dDate = New System.Windows.Forms.DateTimePicker()
        Me.lblDate2 = New System.Windows.Forms.Label()
        Me.fraSaleType = New System.Windows.Forms.GroupBox()
        Me.optWritten = New System.Windows.Forms.RadioButton()
        Me.optDelivered = New System.Windows.Forms.RadioButton()
        Me.fraStoreSelect = New System.Windows.Forms.GroupBox()
        Me.cboStoreSelect = New System.Windows.Forms.ComboBox()
        Me.chkSortByZip = New System.Windows.Forms.CheckBox()
        Me.chkGroupByZip = New System.Windows.Forms.CheckBox()
        Me.cmdCancel = New System.Windows.Forms.Button()
        Me.cmdPrintPreview = New System.Windows.Forms.Button()
        Me.cmdPrint = New System.Windows.Forms.Button()
        Me.ButtonGroupbox = New System.Windows.Forms.GroupBox()
        Me.fraDates.SuspendLayout()
        Me.fraSaleType.SuspendLayout()
        Me.fraStoreSelect.SuspendLayout()
        Me.ButtonGroupbox.SuspendLayout()
        Me.SuspendLayout()
        '
        'fraDates
        '
        Me.fraDates.Controls.Add(Me.lblDate1)
        Me.fraDates.Controls.Add(Me.lblNote)
        Me.fraDates.Controls.Add(Me.toDate)
        Me.fraDates.Controls.Add(Me.dDate)
        Me.fraDates.Controls.Add(Me.lblDate2)
        Me.fraDates.Location = New System.Drawing.Point(7, 7)
        Me.fraDates.Name = "fraDates"
        Me.fraDates.Size = New System.Drawing.Size(261, 66)
        Me.fraDates.TabIndex = 0
        Me.fraDates.TabStop = False
        Me.fraDates.Text = "Date"
        '
        'lblDate1
        '
        Me.lblDate1.Location = New System.Drawing.Point(7, 13)
        Me.lblDate1.Name = "lblDate1"
        Me.lblDate1.Size = New System.Drawing.Size(100, 21)
        Me.lblDate1.TabIndex = 3
        '
        'lblNote
        '
        Me.lblNote.AutoSize = True
        Me.lblNote.Location = New System.Drawing.Point(133, 16)
        Me.lblNote.Name = "lblNote"
        Me.lblNote.Size = New System.Drawing.Size(69, 13)
        Me.lblNote.TabIndex = 2
        Me.lblNote.Text = "Note on Void"
        '
        'toDate
        '
        Me.toDate.CustomFormat = "MM/dd/yyyy"
        Me.toDate.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.toDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.toDate.Location = New System.Drawing.Point(136, 35)
        Me.toDate.Name = "toDate"
        Me.toDate.Size = New System.Drawing.Size(107, 26)
        Me.toDate.TabIndex = 1
        '
        'dDate
        '
        Me.dDate.CustomFormat = "MM/dd/yyyy"
        Me.dDate.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dDate.Location = New System.Drawing.Point(9, 35)
        Me.dDate.Name = "dDate"
        Me.dDate.Size = New System.Drawing.Size(107, 26)
        Me.dDate.TabIndex = 0
        '
        'lblDate2
        '
        Me.lblDate2.Location = New System.Drawing.Point(133, 8)
        Me.lblDate2.Name = "lblDate2"
        Me.lblDate2.Size = New System.Drawing.Size(100, 21)
        Me.lblDate2.TabIndex = 4
        '
        'fraSaleType
        '
        Me.fraSaleType.Controls.Add(Me.optWritten)
        Me.fraSaleType.Controls.Add(Me.optDelivered)
        Me.fraSaleType.Location = New System.Drawing.Point(8, 79)
        Me.fraSaleType.Name = "fraSaleType"
        Me.fraSaleType.Size = New System.Drawing.Size(260, 45)
        Me.fraSaleType.TabIndex = 1
        Me.fraSaleType.TabStop = False
        Me.fraSaleType.Text = "Type of Sale"
        '
        'optWritten
        '
        Me.optWritten.AutoSize = True
        Me.optWritten.Location = New System.Drawing.Point(133, 23)
        Me.optWritten.Name = "optWritten"
        Me.optWritten.Size = New System.Drawing.Size(59, 17)
        Me.optWritten.TabIndex = 1
        Me.optWritten.Text = "&Written"
        Me.optWritten.UseVisualStyleBackColor = True
        '
        'optDelivered
        '
        Me.optDelivered.AutoSize = True
        Me.optDelivered.Checked = True
        Me.optDelivered.Location = New System.Drawing.Point(31, 23)
        Me.optDelivered.Name = "optDelivered"
        Me.optDelivered.Size = New System.Drawing.Size(70, 17)
        Me.optDelivered.TabIndex = 0
        Me.optDelivered.TabStop = True
        Me.optDelivered.Text = "&Delivered"
        Me.optDelivered.UseVisualStyleBackColor = True
        '
        'fraStoreSelect
        '
        Me.fraStoreSelect.Controls.Add(Me.cboStoreSelect)
        Me.fraStoreSelect.Controls.Add(Me.chkSortByZip)
        Me.fraStoreSelect.Controls.Add(Me.chkGroupByZip)
        Me.fraStoreSelect.Location = New System.Drawing.Point(8, 211)
        Me.fraStoreSelect.Name = "fraStoreSelect"
        Me.fraStoreSelect.Size = New System.Drawing.Size(260, 97)
        Me.fraStoreSelect.TabIndex = 2
        Me.fraStoreSelect.TabStop = False
        Me.fraStoreSelect.Text = "Select a Store"
        Me.fraStoreSelect.Visible = False
        '
        'cboStoreSelect
        '
        Me.cboStoreSelect.FormattingEnabled = True
        Me.cboStoreSelect.Location = New System.Drawing.Point(12, 19)
        Me.cboStoreSelect.Name = "cboStoreSelect"
        Me.cboStoreSelect.Size = New System.Drawing.Size(239, 21)
        Me.cboStoreSelect.TabIndex = 6
        '
        'chkSortByZip
        '
        Me.chkSortByZip.AutoSize = True
        Me.chkSortByZip.Location = New System.Drawing.Point(38, 76)
        Me.chkSortByZip.Name = "chkSortByZip"
        Me.chkSortByZip.Size = New System.Drawing.Size(140, 17)
        Me.chkSortByZip.TabIndex = 2
        Me.chkSortByZip.Text = "&Sort Report by Zip Code"
        Me.chkSortByZip.UseVisualStyleBackColor = True
        '
        'chkGroupByZip
        '
        Me.chkGroupByZip.AutoSize = True
        Me.chkGroupByZip.Location = New System.Drawing.Point(38, 53)
        Me.chkGroupByZip.Name = "chkGroupByZip"
        Me.chkGroupByZip.Size = New System.Drawing.Size(155, 17)
        Me.chkGroupByZip.TabIndex = 1
        Me.chkGroupByZip.Text = "&Combine Cities by Zip Code"
        Me.chkGroupByZip.UseVisualStyleBackColor = True
        '
        'cmdCancel
        '
        Me.cmdCancel.Location = New System.Drawing.Point(168, 13)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.Size = New System.Drawing.Size(75, 56)
        Me.cmdCancel.TabIndex = 5
        Me.cmdCancel.Text = "&Cancel"
        Me.cmdCancel.UseVisualStyleBackColor = True
        '
        'cmdPrintPreview
        '
        Me.cmdPrintPreview.Location = New System.Drawing.Point(78, 13)
        Me.cmdPrintPreview.Name = "cmdPrintPreview"
        Me.cmdPrintPreview.Size = New System.Drawing.Size(90, 56)
        Me.cmdPrintPreview.TabIndex = 4
        Me.cmdPrintPreview.Text = "P&rint Preview"
        Me.cmdPrintPreview.UseVisualStyleBackColor = True
        '
        'cmdPrint
        '
        Me.cmdPrint.Location = New System.Drawing.Point(4, 13)
        Me.cmdPrint.Name = "cmdPrint"
        Me.cmdPrint.Size = New System.Drawing.Size(75, 56)
        Me.cmdPrint.TabIndex = 3
        Me.cmdPrint.Text = "&Print"
        Me.cmdPrint.UseVisualStyleBackColor = True
        '
        'ButtonGroupbox
        '
        Me.ButtonGroupbox.Controls.Add(Me.cmdPrintPreview)
        Me.ButtonGroupbox.Controls.Add(Me.cmdPrint)
        Me.ButtonGroupbox.Controls.Add(Me.cmdCancel)
        Me.ButtonGroupbox.Location = New System.Drawing.Point(8, 149)
        Me.ButtonGroupbox.Name = "ButtonGroupbox"
        Me.ButtonGroupbox.Size = New System.Drawing.Size(262, 75)
        Me.ButtonGroupbox.TabIndex = 6
        Me.ButtonGroupbox.TabStop = False
        '
        'DateForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(278, 311)
        Me.Controls.Add(Me.ButtonGroupbox)
        Me.Controls.Add(Me.fraStoreSelect)
        Me.Controls.Add(Me.fraSaleType)
        Me.Controls.Add(Me.fraDates)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Name = "DateForm"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Date"
        Me.fraDates.ResumeLayout(False)
        Me.fraDates.PerformLayout()
        Me.fraSaleType.ResumeLayout(False)
        Me.fraSaleType.PerformLayout()
        Me.fraStoreSelect.ResumeLayout(False)
        Me.fraStoreSelect.PerformLayout()
        Me.ButtonGroupbox.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents fraDates As GroupBox
    Friend WithEvents toDate As DateTimePicker
    Friend WithEvents dDate As DateTimePicker
    Friend WithEvents fraSaleType As GroupBox
    Friend WithEvents optWritten As RadioButton
    Friend WithEvents optDelivered As RadioButton
    Friend WithEvents fraStoreSelect As GroupBox
    Friend WithEvents cmdCancel As Button
    Friend WithEvents cmdPrintPreview As Button
    Friend WithEvents cmdPrint As Button
    Friend WithEvents chkSortByZip As CheckBox
    Friend WithEvents chkGroupByZip As CheckBox
    Friend WithEvents lblNote As Label
    Friend WithEvents cboStoreSelect As ComboBox
    Friend WithEvents lblDate1 As Label
    Friend WithEvents lblDate2 As Label
    Friend WithEvents ButtonGroupbox As GroupBox
End Class
