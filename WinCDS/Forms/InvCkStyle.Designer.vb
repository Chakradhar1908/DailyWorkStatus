<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class InvCkStyle
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
        Me.components = New System.ComponentModel.Container()
        Me.fraSearch = New System.Windows.Forms.GroupBox()
        Me.Style = New System.Windows.Forms.TextBox()
        Me.optSearchByStyle = New System.Windows.Forms.RadioButton()
        Me.optSearchByVendor = New System.Windows.Forms.RadioButton()
        Me.optSearchByDesc = New System.Windows.Forms.RadioButton()
        Me.optKitVendors = New System.Windows.Forms.RadioButton()
        Me.chkStkOnly = New System.Windows.Forms.CheckBox()
        Me.cmdApply = New System.Windows.Forms.Button()
        Me.cmdDesc = New System.Windows.Forms.Button()
        Me.cmdCancel = New System.Windows.Forms.Button()
        Me.cmdBarcode = New System.Windows.Forms.Button()
        Me.lblCaptions = New System.Windows.Forms.Label()
        Me.lstStyles = New System.Windows.Forms.ListBox()
        Me.tmrItemPreview = New System.Windows.Forms.Timer(Me.components)
        Me.tmrType = New System.Windows.Forms.Timer(Me.components)
        Me.fraSearch.SuspendLayout()
        Me.SuspendLayout()
        '
        'fraSearch
        '
        Me.fraSearch.Controls.Add(Me.Style)
        Me.fraSearch.Location = New System.Drawing.Point(8, 9)
        Me.fraSearch.Name = "fraSearch"
        Me.fraSearch.Size = New System.Drawing.Size(190, 50)
        Me.fraSearch.TabIndex = 0
        Me.fraSearch.TabStop = False
        Me.fraSearch.Text = "&Style Number:"
        '
        'Style
        '
        Me.Style.Font = New System.Drawing.Font("Lucida Console", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Style.Location = New System.Drawing.Point(12, 19)
        Me.Style.Name = "Style"
        Me.Style.Size = New System.Drawing.Size(167, 22)
        Me.Style.TabIndex = 0
        Me.Style.Text = "1234567890123456"
        '
        'optSearchByStyle
        '
        Me.optSearchByStyle.AutoSize = True
        Me.optSearchByStyle.Checked = True
        Me.optSearchByStyle.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.optSearchByStyle.Location = New System.Drawing.Point(8, 64)
        Me.optSearchByStyle.Name = "optSearchByStyle"
        Me.optSearchByStyle.Size = New System.Drawing.Size(80, 20)
        Me.optSearchByStyle.TabIndex = 1
        Me.optSearchByStyle.TabStop = True
        Me.optSearchByStyle.Text = "St&yle No."
        Me.optSearchByStyle.UseVisualStyleBackColor = True
        '
        'optSearchByVendor
        '
        Me.optSearchByVendor.AutoSize = True
        Me.optSearchByVendor.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.optSearchByVendor.Location = New System.Drawing.Point(8, 87)
        Me.optSearchByVendor.Name = "optSearchByVendor"
        Me.optSearchByVendor.Size = New System.Drawing.Size(110, 20)
        Me.optSearchByVendor.TabIndex = 2
        Me.optSearchByVendor.Text = "&Vendor Name"
        Me.optSearchByVendor.UseVisualStyleBackColor = True
        '
        'optSearchByDesc
        '
        Me.optSearchByDesc.AutoSize = True
        Me.optSearchByDesc.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.optSearchByDesc.Location = New System.Drawing.Point(8, 110)
        Me.optSearchByDesc.Name = "optSearchByDesc"
        Me.optSearchByDesc.Size = New System.Drawing.Size(94, 20)
        Me.optSearchByDesc.TabIndex = 3
        Me.optSearchByDesc.Text = "Desc&ription"
        Me.optSearchByDesc.UseVisualStyleBackColor = True
        '
        'optKitVendors
        '
        Me.optKitVendors.AutoSize = True
        Me.optKitVendors.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.optKitVendors.Location = New System.Drawing.Point(8, 133)
        Me.optKitVendors.Name = "optKitVendors"
        Me.optKitVendors.Size = New System.Drawing.Size(94, 20)
        Me.optKitVendors.TabIndex = 4
        Me.optKitVendors.Text = "&Kit Vendors"
        Me.optKitVendors.UseVisualStyleBackColor = True
        '
        'chkStkOnly
        '
        Me.chkStkOnly.AutoSize = True
        Me.chkStkOnly.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkStkOnly.Location = New System.Drawing.Point(126, 113)
        Me.chkStkOnly.Name = "chkStkOnly"
        Me.chkStkOnly.Size = New System.Drawing.Size(74, 20)
        Me.chkStkOnly.TabIndex = 6
        Me.chkStkOnly.Text = "In Stoc&k"
        Me.chkStkOnly.UseVisualStyleBackColor = True
        '
        'cmdApply
        '
        Me.cmdApply.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdApply.Location = New System.Drawing.Point(8, 168)
        Me.cmdApply.Name = "cmdApply"
        Me.cmdApply.Size = New System.Drawing.Size(54, 23)
        Me.cmdApply.TabIndex = 7
        Me.cmdApply.Text = "&OK"
        Me.cmdApply.UseVisualStyleBackColor = True
        '
        'cmdDesc
        '
        Me.cmdDesc.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdDesc.Location = New System.Drawing.Point(76, 168)
        Me.cmdDesc.Name = "cmdDesc"
        Me.cmdDesc.Size = New System.Drawing.Size(54, 23)
        Me.cmdDesc.TabIndex = 8
        Me.cmdDesc.Text = "&Desc."
        Me.cmdDesc.UseVisualStyleBackColor = True
        '
        'cmdCancel
        '
        Me.cmdCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cmdCancel.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdCancel.Location = New System.Drawing.Point(142, 168)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.Size = New System.Drawing.Size(60, 23)
        Me.cmdCancel.TabIndex = 9
        Me.cmdCancel.Text = "&Cancel"
        Me.cmdCancel.UseVisualStyleBackColor = True
        '
        'cmdBarcode
        '
        Me.cmdBarcode.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdBarcode.Location = New System.Drawing.Point(8, 197)
        Me.cmdBarcode.Name = "cmdBarcode"
        Me.cmdBarcode.Size = New System.Drawing.Size(194, 23)
        Me.cmdBarcode.TabIndex = 10
        Me.cmdBarcode.Text = "Get &Barcode (F2)"
        Me.cmdBarcode.UseVisualStyleBackColor = True
        '
        'lblCaptions
        '
        Me.lblCaptions.BackColor = System.Drawing.SystemColors.Control
        Me.lblCaptions.Font = New System.Drawing.Font("Lucida Console", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblCaptions.Location = New System.Drawing.Point(210, 6)
        Me.lblCaptions.Name = "lblCaptions"
        Me.lblCaptions.Size = New System.Drawing.Size(31, 13)
        Me.lblCaptions.TabIndex = 11
        Me.lblCaptions.Text = "###"
        '
        'lstStyles
        '
        Me.lstStyles.FormattingEnabled = True
        Me.lstStyles.Items.AddRange(New Object() {"1", "2", "3", "4"})
        Me.lstStyles.Location = New System.Drawing.Point(210, 23)
        Me.lstStyles.Name = "lstStyles"
        Me.lstStyles.Size = New System.Drawing.Size(220, 199)
        Me.lstStyles.TabIndex = 12
        '
        'tmrItemPreview
        '
        Me.tmrItemPreview.Interval = 200
        '
        'tmrType
        '
        '
        'InvCkStyle
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.cmdCancel
        Me.ClientSize = New System.Drawing.Size(432, 233)
        Me.ControlBox = False
        Me.Controls.Add(Me.lstStyles)
        Me.Controls.Add(Me.lblCaptions)
        Me.Controls.Add(Me.cmdBarcode)
        Me.Controls.Add(Me.cmdCancel)
        Me.Controls.Add(Me.cmdDesc)
        Me.Controls.Add(Me.cmdApply)
        Me.Controls.Add(Me.chkStkOnly)
        Me.Controls.Add(Me.optKitVendors)
        Me.Controls.Add(Me.optSearchByDesc)
        Me.Controls.Add(Me.optSearchByVendor)
        Me.Controls.Add(Me.optSearchByStyle)
        Me.Controls.Add(Me.fraSearch)
        Me.Name = "InvCkStyle"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Enter Style Number"
        Me.fraSearch.ResumeLayout(False)
        Me.fraSearch.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents fraSearch As GroupBox
    Friend WithEvents Style As TextBox
    Friend WithEvents optSearchByStyle As RadioButton
    Friend WithEvents optSearchByVendor As RadioButton
    Friend WithEvents optSearchByDesc As RadioButton
    Friend WithEvents optKitVendors As RadioButton
    Friend WithEvents chkStkOnly As CheckBox
    Friend WithEvents cmdApply As Button
    Friend WithEvents cmdDesc As Button
    Friend WithEvents cmdCancel As Button
    Friend WithEvents cmdBarcode As Button
    Friend WithEvents lblCaptions As Label
    Friend WithEvents lstStyles As ListBox
    Friend WithEvents tmrItemPreview As Timer
    Friend WithEvents tmrType As Timer
End Class
