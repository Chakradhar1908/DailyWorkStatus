<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class InvDel
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
        Me.cboDept = New System.Windows.Forms.ComboBox()
        Me.cboVendor = New System.Windows.Forms.ComboBox()
        Me.lblDept = New System.Windows.Forms.Label()
        Me.lblVendor = New System.Windows.Forms.Label()
        Me.Style = New System.Windows.Forms.Label()
        Me.DDate = New System.Windows.Forms.DateTimePicker()
        Me.fraControls = New System.Windows.Forms.GroupBox()
        Me.cmdDeliverAll = New System.Windows.Forms.Button()
        Me.cmdPrint = New System.Windows.Forms.Button()
        Me.cmdNotes = New System.Windows.Forms.Button()
        Me.cmdSkip = New System.Windows.Forms.Button()
        Me.cmdDeliver = New System.Windows.Forms.Button()
        Me.lblCost = New System.Windows.Forms.Label()
        Me.lblFreight = New System.Windows.Forms.Label()
        Me.Cost = New System.Windows.Forms.TextBox()
        Me.Freight = New System.Windows.Forms.TextBox()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.Button1 = New System.Windows.Forms.Button()
        Me.fraControls.SuspendLayout()
        Me.SuspendLayout()
        '
        'cboDept
        '
        Me.cboDept.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboDept.FormattingEnabled = True
        Me.cboDept.Location = New System.Drawing.Point(204, 11)
        Me.cboDept.Name = "cboDept"
        Me.cboDept.Size = New System.Drawing.Size(188, 24)
        Me.cboDept.TabIndex = 5
        Me.cboDept.Text = "cboDept"
        '
        'cboVendor
        '
        Me.cboVendor.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboVendor.FormattingEnabled = True
        Me.cboVendor.Location = New System.Drawing.Point(204, 34)
        Me.cboVendor.Name = "cboVendor"
        Me.cboVendor.Size = New System.Drawing.Size(188, 24)
        Me.cboVendor.TabIndex = 6
        Me.cboVendor.Text = "cboVendor"
        '
        'lblDept
        '
        Me.lblDept.AutoSize = True
        Me.lblDept.Location = New System.Drawing.Point(170, 11)
        Me.lblDept.Name = "lblDept"
        Me.lblDept.Size = New System.Drawing.Size(33, 13)
        Me.lblDept.TabIndex = 3
        Me.lblDept.Text = "De&pt:"
        '
        'lblVendor
        '
        Me.lblVendor.AutoSize = True
        Me.lblVendor.Location = New System.Drawing.Point(159, 37)
        Me.lblVendor.Name = "lblVendor"
        Me.lblVendor.Size = New System.Drawing.Size(44, 13)
        Me.lblVendor.TabIndex = 4
        Me.lblVendor.Text = "&Vendor:"
        '
        'Style
        '
        Me.Style.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Style.Font = New System.Drawing.Font("Lucida Console", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Style.Location = New System.Drawing.Point(2, 59)
        Me.Style.Name = "Style"
        Me.Style.Size = New System.Drawing.Size(166, 23)
        Me.Style.TabIndex = 0
        Me.Style.Text = "Style"
        Me.Style.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.ToolTip1.SetToolTip(Me.Style, "The style of your item.")
        '
        'DDate
        '
        Me.DDate.CustomFormat = "MM/dd/yyyy"
        Me.DDate.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.DDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DDate.Location = New System.Drawing.Point(32, 22)
        Me.DDate.Name = "DDate"
        Me.DDate.Size = New System.Drawing.Size(101, 26)
        Me.DDate.TabIndex = 9
        Me.ToolTip1.SetToolTip(Me.DDate, "The date to mark this item as delivered.")
        '
        'fraControls
        '
        Me.fraControls.Controls.Add(Me.cmdDeliverAll)
        Me.fraControls.Controls.Add(Me.cmdPrint)
        Me.fraControls.Controls.Add(Me.cmdNotes)
        Me.fraControls.Controls.Add(Me.cmdSkip)
        Me.fraControls.Controls.Add(Me.cmdDeliver)
        Me.fraControls.Location = New System.Drawing.Point(2, 85)
        Me.fraControls.Name = "fraControls"
        Me.fraControls.Size = New System.Drawing.Size(166, 93)
        Me.fraControls.TabIndex = 7
        Me.fraControls.TabStop = False
        '
        'cmdDeliverAll
        '
        Me.cmdDeliverAll.Location = New System.Drawing.Point(8, 60)
        Me.cmdDeliverAll.Name = "cmdDeliverAll"
        Me.cmdDeliverAll.Size = New System.Drawing.Size(149, 23)
        Me.cmdDeliverAll.TabIndex = 4
        Me.cmdDeliverAll.Text = "Deliver &All"
        Me.ToolTip1.SetToolTip(Me.cmdDeliverAll, "Click here to view notes for this item.")
        Me.cmdDeliverAll.UseVisualStyleBackColor = True
        '
        'cmdPrint
        '
        Me.cmdPrint.Location = New System.Drawing.Point(82, 38)
        Me.cmdPrint.Name = "cmdPrint"
        Me.cmdPrint.Size = New System.Drawing.Size(75, 23)
        Me.cmdPrint.TabIndex = 3
        Me.cmdPrint.Text = "&Print Bill"
        Me.ToolTip1.SetToolTip(Me.cmdPrint, "Click to print the bill.")
        Me.cmdPrint.UseVisualStyleBackColor = True
        '
        'cmdNotes
        '
        Me.cmdNotes.Location = New System.Drawing.Point(8, 38)
        Me.cmdNotes.Name = "cmdNotes"
        Me.cmdNotes.Size = New System.Drawing.Size(75, 23)
        Me.cmdNotes.TabIndex = 2
        Me.cmdNotes.Text = "&Notes"
        Me.ToolTip1.SetToolTip(Me.cmdNotes, "Click here to view notes for this item.")
        Me.cmdNotes.UseVisualStyleBackColor = True
        '
        'cmdSkip
        '
        Me.cmdSkip.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cmdSkip.Location = New System.Drawing.Point(82, 15)
        Me.cmdSkip.Name = "cmdSkip"
        Me.cmdSkip.Size = New System.Drawing.Size(75, 23)
        Me.cmdSkip.TabIndex = 1
        Me.cmdSkip.Text = "&Skip Item"
        Me.ToolTip1.SetToolTip(Me.cmdSkip, "Click here to skip this item.")
        Me.cmdSkip.UseVisualStyleBackColor = True
        '
        'cmdDeliver
        '
        Me.cmdDeliver.Location = New System.Drawing.Point(8, 15)
        Me.cmdDeliver.Name = "cmdDeliver"
        Me.cmdDeliver.Size = New System.Drawing.Size(75, 23)
        Me.cmdDeliver.TabIndex = 0
        Me.cmdDeliver.Text = "&Deliver"
        Me.ToolTip1.SetToolTip(Me.cmdDeliver, "Click here to deliver this item.")
        Me.cmdDeliver.UseVisualStyleBackColor = True
        '
        'lblCost
        '
        Me.lblCost.AutoSize = True
        Me.lblCost.Location = New System.Drawing.Point(255, 100)
        Me.lblCost.Name = "lblCost"
        Me.lblCost.Size = New System.Drawing.Size(34, 13)
        Me.lblCost.TabIndex = 8
        Me.lblCost.Text = "&Cost: "
        '
        'lblFreight
        '
        Me.lblFreight.Location = New System.Drawing.Point(249, 133)
        Me.lblFreight.Name = "lblFreight"
        Me.lblFreight.Size = New System.Drawing.Size(39, 13)
        Me.lblFreight.TabIndex = 9
        Me.lblFreight.Text = "&Freight:"
        '
        'Cost
        '
        Me.Cost.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Cost.Location = New System.Drawing.Point(292, 99)
        Me.Cost.Name = "Cost"
        Me.Cost.Size = New System.Drawing.Size(100, 26)
        Me.Cost.TabIndex = 7
        '
        'Freight
        '
        Me.Freight.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Freight.Location = New System.Drawing.Point(292, 129)
        Me.Freight.Name = "Freight"
        Me.Freight.Size = New System.Drawing.Size(100, 26)
        Me.Freight.TabIndex = 8
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(32, 1)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(75, 23)
        Me.Button1.TabIndex = 10
        Me.Button1.Text = "Button1"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'InvDel
        '
        Me.AcceptButton = Me.cmdDeliver
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.cmdSkip
        Me.ClientSize = New System.Drawing.Size(396, 187)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.Freight)
        Me.Controls.Add(Me.Cost)
        Me.Controls.Add(Me.lblFreight)
        Me.Controls.Add(Me.lblCost)
        Me.Controls.Add(Me.fraControls)
        Me.Controls.Add(Me.DDate)
        Me.Controls.Add(Me.Style)
        Me.Controls.Add(Me.lblVendor)
        Me.Controls.Add(Me.lblDept)
        Me.Controls.Add(Me.cboVendor)
        Me.Controls.Add(Me.cboDept)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.MaximizeBox = False
        Me.Name = "InvDel"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.fraControls.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents cboDept As ComboBox
    Friend WithEvents cboVendor As ComboBox
    Friend WithEvents lblDept As Label
    Friend WithEvents lblVendor As Label
    Friend WithEvents Style As Label
    Friend WithEvents DDate As DateTimePicker
    Friend WithEvents fraControls As GroupBox
    Friend WithEvents cmdDeliverAll As Button
    Friend WithEvents cmdPrint As Button
    Friend WithEvents cmdNotes As Button
    Friend WithEvents cmdSkip As Button
    Friend WithEvents cmdDeliver As Button
    Friend WithEvents lblCost As Label
    Friend WithEvents lblFreight As Label
    Friend WithEvents Cost As TextBox
    Friend WithEvents Freight As TextBox
    Friend WithEvents ToolTip1 As ToolTip
    Friend WithEvents Button1 As Button
End Class
