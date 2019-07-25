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
        Me.cboDept = New System.Windows.Forms.ComboBox()
        Me.cboVendor = New System.Windows.Forms.ComboBox()
        Me.lblDept = New System.Windows.Forms.Label()
        Me.lblVendor = New System.Windows.Forms.Label()
        Me.Style = New System.Windows.Forms.Label()
        Me.DDate = New System.Windows.Forms.DateTimePicker()
        Me.fraControls = New System.Windows.Forms.GroupBox()
        Me.cmdDeliver = New System.Windows.Forms.Button()
        Me.cmdSkip = New System.Windows.Forms.Button()
        Me.cmdNotes = New System.Windows.Forms.Button()
        Me.cmdPrint = New System.Windows.Forms.Button()
        Me.cmdDeliverAll = New System.Windows.Forms.Button()
        Me.lblCost = New System.Windows.Forms.Label()
        Me.lblFreight = New System.Windows.Forms.Label()
        Me.Cost = New System.Windows.Forms.TextBox()
        Me.Freight = New System.Windows.Forms.TextBox()
        Me.fraControls.SuspendLayout()
        Me.SuspendLayout()
        '
        'cboDept
        '
        Me.cboDept.FormattingEnabled = True
        Me.cboDept.Location = New System.Drawing.Point(214, 11)
        Me.cboDept.Name = "cboDept"
        Me.cboDept.Size = New System.Drawing.Size(188, 21)
        Me.cboDept.TabIndex = 1
        Me.cboDept.Text = "cboDept"
        '
        'cboVendor
        '
        Me.cboVendor.FormattingEnabled = True
        Me.cboVendor.Location = New System.Drawing.Point(214, 34)
        Me.cboVendor.Name = "cboVendor"
        Me.cboVendor.Size = New System.Drawing.Size(188, 21)
        Me.cboVendor.TabIndex = 2
        Me.cboVendor.Text = "cboVendor"
        '
        'lblDept
        '
        Me.lblDept.AutoSize = True
        Me.lblDept.Location = New System.Drawing.Point(180, 11)
        Me.lblDept.Name = "lblDept"
        Me.lblDept.Size = New System.Drawing.Size(33, 13)
        Me.lblDept.TabIndex = 3
        Me.lblDept.Text = "De&pt:"
        '
        'lblVendor
        '
        Me.lblVendor.AutoSize = True
        Me.lblVendor.Location = New System.Drawing.Point(169, 37)
        Me.lblVendor.Name = "lblVendor"
        Me.lblVendor.Size = New System.Drawing.Size(44, 13)
        Me.lblVendor.TabIndex = 4
        Me.lblVendor.Text = "&Vendor:"
        '
        'Style
        '
        Me.Style.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Style.Location = New System.Drawing.Point(12, 61)
        Me.Style.Name = "Style"
        Me.Style.Size = New System.Drawing.Size(166, 23)
        Me.Style.TabIndex = 5
        Me.Style.Text = "Style"
        '
        'DDate
        '
        Me.DDate.CustomFormat = "MM/dd/yyyy"
        Me.DDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DDate.Location = New System.Drawing.Point(42, 22)
        Me.DDate.Name = "DDate"
        Me.DDate.Size = New System.Drawing.Size(101, 20)
        Me.DDate.TabIndex = 6
        '
        'fraControls
        '
        Me.fraControls.Controls.Add(Me.cmdDeliverAll)
        Me.fraControls.Controls.Add(Me.cmdPrint)
        Me.fraControls.Controls.Add(Me.cmdNotes)
        Me.fraControls.Controls.Add(Me.cmdSkip)
        Me.fraControls.Controls.Add(Me.cmdDeliver)
        Me.fraControls.Location = New System.Drawing.Point(12, 85)
        Me.fraControls.Name = "fraControls"
        Me.fraControls.Size = New System.Drawing.Size(166, 93)
        Me.fraControls.TabIndex = 7
        Me.fraControls.TabStop = False
        '
        'cmdDeliver
        '
        Me.cmdDeliver.Location = New System.Drawing.Point(8, 15)
        Me.cmdDeliver.Name = "cmdDeliver"
        Me.cmdDeliver.Size = New System.Drawing.Size(75, 23)
        Me.cmdDeliver.TabIndex = 0
        Me.cmdDeliver.Text = "&Deliver"
        Me.cmdDeliver.UseVisualStyleBackColor = True
        '
        'cmdSkip
        '
        Me.cmdSkip.Location = New System.Drawing.Point(82, 15)
        Me.cmdSkip.Name = "cmdSkip"
        Me.cmdSkip.Size = New System.Drawing.Size(75, 23)
        Me.cmdSkip.TabIndex = 1
        Me.cmdSkip.Text = "&Skip Item"
        Me.cmdSkip.UseVisualStyleBackColor = True
        '
        'cmdNotes
        '
        Me.cmdNotes.Location = New System.Drawing.Point(8, 38)
        Me.cmdNotes.Name = "cmdNotes"
        Me.cmdNotes.Size = New System.Drawing.Size(75, 23)
        Me.cmdNotes.TabIndex = 2
        Me.cmdNotes.Text = "&Notes"
        Me.cmdNotes.UseVisualStyleBackColor = True
        '
        'cmdPrint
        '
        Me.cmdPrint.Location = New System.Drawing.Point(82, 38)
        Me.cmdPrint.Name = "cmdPrint"
        Me.cmdPrint.Size = New System.Drawing.Size(75, 23)
        Me.cmdPrint.TabIndex = 3
        Me.cmdPrint.Text = "&Print Bill"
        Me.cmdPrint.UseVisualStyleBackColor = True
        '
        'cmdDeliverAll
        '
        Me.cmdDeliverAll.Location = New System.Drawing.Point(8, 60)
        Me.cmdDeliverAll.Name = "cmdDeliverAll"
        Me.cmdDeliverAll.Size = New System.Drawing.Size(149, 23)
        Me.cmdDeliverAll.TabIndex = 4
        Me.cmdDeliverAll.Text = "Deliver &All"
        Me.cmdDeliverAll.UseVisualStyleBackColor = True
        '
        'lblCost
        '
        Me.lblCost.AutoSize = True
        Me.lblCost.Location = New System.Drawing.Point(262, 105)
        Me.lblCost.Name = "lblCost"
        Me.lblCost.Size = New System.Drawing.Size(34, 13)
        Me.lblCost.TabIndex = 8
        Me.lblCost.Text = "&Cost: "
        '
        'lblFreight
        '
        Me.lblFreight.AutoSize = True
        Me.lblFreight.Location = New System.Drawing.Point(254, 128)
        Me.lblFreight.Name = "lblFreight"
        Me.lblFreight.Size = New System.Drawing.Size(42, 13)
        Me.lblFreight.TabIndex = 9
        Me.lblFreight.Text = "&Freight:"
        '
        'Cost
        '
        Me.Cost.Location = New System.Drawing.Point(302, 102)
        Me.Cost.Name = "Cost"
        Me.Cost.Size = New System.Drawing.Size(100, 20)
        Me.Cost.TabIndex = 10
        '
        'Freight
        '
        Me.Freight.Location = New System.Drawing.Point(302, 126)
        Me.Freight.Name = "Freight"
        Me.Freight.Size = New System.Drawing.Size(100, 20)
        Me.Freight.TabIndex = 11
        '
        'InvDel
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(410, 187)
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
        Me.Name = "InvDel"
        Me.Text = "InvDel"
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
End Class
