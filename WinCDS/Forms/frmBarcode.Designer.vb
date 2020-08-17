<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmBarcode
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
        Me.lblTitle = New System.Windows.Forms.Label()
        Me.lstBarcodes = New System.Windows.Forms.ListBox()
        Me.cmdTransfer = New System.Windows.Forms.Button()
        Me.cmdCancel = New System.Windows.Forms.Button()
        Me.cmdDelete = New System.Windows.Forms.Button()
        Me.cmdClear = New System.Windows.Forms.Button()
        Me.cmdGetBarcodes = New System.Windows.Forms.Button()
        Me.cmdClearCS1504 = New System.Windows.Forms.Button()
        Me.cmdOptions = New System.Windows.Forms.Button()
        Me.cmdHelp = New System.Windows.Forms.Button()
        Me.chkShowCost = New System.Windows.Forms.CheckBox()
        Me.lblStatus = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'lblTitle
        '
        Me.lblTitle.Location = New System.Drawing.Point(4, -1)
        Me.lblTitle.Name = "lblTitle"
        Me.lblTitle.Size = New System.Drawing.Size(100, 16)
        Me.lblTitle.TabIndex = 0
        Me.lblTitle.Text = "&Barcodes:"
        '
        'lstBarcodes
        '
        Me.lstBarcodes.FormattingEnabled = True
        Me.lstBarcodes.Location = New System.Drawing.Point(4, 16)
        Me.lstBarcodes.Name = "lstBarcodes"
        Me.lstBarcodes.Size = New System.Drawing.Size(120, 173)
        Me.lstBarcodes.TabIndex = 1
        '
        'cmdTransfer
        '
        Me.cmdTransfer.Enabled = False
        Me.cmdTransfer.Location = New System.Drawing.Point(138, 8)
        Me.cmdTransfer.Name = "cmdTransfer"
        Me.cmdTransfer.Size = New System.Drawing.Size(83, 23)
        Me.cmdTransfer.TabIndex = 2
        Me.cmdTransfer.Text = "&Transfer"
        Me.cmdTransfer.UseVisualStyleBackColor = True
        '
        'cmdCancel
        '
        Me.cmdCancel.Location = New System.Drawing.Point(138, 31)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.Size = New System.Drawing.Size(83, 23)
        Me.cmdCancel.TabIndex = 3
        Me.cmdCancel.Text = "&Cancel"
        Me.cmdCancel.UseVisualStyleBackColor = True
        '
        'cmdDelete
        '
        Me.cmdDelete.Enabled = False
        Me.cmdDelete.Location = New System.Drawing.Point(138, 58)
        Me.cmdDelete.Name = "cmdDelete"
        Me.cmdDelete.Size = New System.Drawing.Size(83, 23)
        Me.cmdDelete.TabIndex = 4
        Me.cmdDelete.Text = "&Delete"
        Me.cmdDelete.UseVisualStyleBackColor = True
        '
        'cmdClear
        '
        Me.cmdClear.Enabled = False
        Me.cmdClear.Location = New System.Drawing.Point(138, 81)
        Me.cmdClear.Name = "cmdClear"
        Me.cmdClear.Size = New System.Drawing.Size(83, 23)
        Me.cmdClear.TabIndex = 5
        Me.cmdClear.Text = "&Clear"
        Me.cmdClear.UseVisualStyleBackColor = True
        '
        'cmdGetBarcodes
        '
        Me.cmdGetBarcodes.Location = New System.Drawing.Point(138, 109)
        Me.cmdGetBarcodes.Name = "cmdGetBarcodes"
        Me.cmdGetBarcodes.Size = New System.Drawing.Size(83, 23)
        Me.cmdGetBarcodes.TabIndex = 6
        Me.cmdGetBarcodes.Text = "&Acquire"
        Me.cmdGetBarcodes.UseVisualStyleBackColor = True
        '
        'cmdClearCS1504
        '
        Me.cmdClearCS1504.Location = New System.Drawing.Point(138, 132)
        Me.cmdClearCS1504.Name = "cmdClearCS1504"
        Me.cmdClearCS1504.Size = New System.Drawing.Size(83, 23)
        Me.cmdClearCS1504.TabIndex = 7
        Me.cmdClearCS1504.Text = "C&lear CS1504"
        Me.cmdClearCS1504.UseVisualStyleBackColor = True
        '
        'cmdOptions
        '
        Me.cmdOptions.Location = New System.Drawing.Point(138, 161)
        Me.cmdOptions.Name = "cmdOptions"
        Me.cmdOptions.Size = New System.Drawing.Size(83, 23)
        Me.cmdOptions.TabIndex = 8
        Me.cmdOptions.Text = "&Options"
        Me.cmdOptions.UseVisualStyleBackColor = True
        '
        'cmdHelp
        '
        Me.cmdHelp.Enabled = False
        Me.cmdHelp.Location = New System.Drawing.Point(138, 192)
        Me.cmdHelp.Name = "cmdHelp"
        Me.cmdHelp.Size = New System.Drawing.Size(83, 23)
        Me.cmdHelp.TabIndex = 9
        Me.cmdHelp.Text = "&Help"
        Me.cmdHelp.UseVisualStyleBackColor = True
        '
        'chkShowCost
        '
        Me.chkShowCost.AutoSize = True
        Me.chkShowCost.Location = New System.Drawing.Point(138, 221)
        Me.chkShowCost.Name = "chkShowCost"
        Me.chkShowCost.Size = New System.Drawing.Size(77, 17)
        Me.chkShowCost.TabIndex = 10
        Me.chkShowCost.Text = "Show Co&st"
        Me.chkShowCost.UseVisualStyleBackColor = True
        '
        'lblStatus
        '
        Me.lblStatus.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblStatus.Location = New System.Drawing.Point(4, 240)
        Me.lblStatus.Name = "lblStatus"
        Me.lblStatus.Size = New System.Drawing.Size(217, 19)
        Me.lblStatus.TabIndex = 11
        Me.lblStatus.Text = "Port Closed"
        '
        'frmBarcode
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(226, 259)
        Me.Controls.Add(Me.lblStatus)
        Me.Controls.Add(Me.chkShowCost)
        Me.Controls.Add(Me.cmdHelp)
        Me.Controls.Add(Me.cmdOptions)
        Me.Controls.Add(Me.cmdClearCS1504)
        Me.Controls.Add(Me.cmdGetBarcodes)
        Me.Controls.Add(Me.cmdClear)
        Me.Controls.Add(Me.cmdDelete)
        Me.Controls.Add(Me.cmdCancel)
        Me.Controls.Add(Me.cmdTransfer)
        Me.Controls.Add(Me.lstBarcodes)
        Me.Controls.Add(Me.lblTitle)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.MaximizeBox = False
        Me.Name = "frmBarcode"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Acquire Barcodes"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents lblTitle As Label
    Friend WithEvents lstBarcodes As ListBox
    Friend WithEvents cmdTransfer As Button
    Friend WithEvents cmdCancel As Button
    Friend WithEvents cmdDelete As Button
    Friend WithEvents cmdClear As Button
    Friend WithEvents cmdGetBarcodes As Button
    Friend WithEvents cmdClearCS1504 As Button
    Friend WithEvents cmdOptions As Button
    Friend WithEvents cmdHelp As Button
    Friend WithEvents chkShowCost As CheckBox
    Friend WithEvents lblStatus As Label
End Class
