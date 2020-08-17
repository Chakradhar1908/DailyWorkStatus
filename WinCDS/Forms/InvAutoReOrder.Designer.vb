<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class InvAutoReOrder
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
        Me.lblTotal = New System.Windows.Forms.Label()
        Me.lblTotalOrdered = New System.Windows.Forms.Label()
        Me.dteStartDate = New System.Windows.Forms.DateTimePicker()
        Me.dteEndDate = New System.Windows.Forms.DateTimePicker()
        Me.cboStoreSelect = New System.Windows.Forms.ComboBox()
        Me.cmdReset = New System.Windows.Forms.Button()
        Me.lblVendor = New System.Windows.Forms.Label()
        Me.lblDept = New System.Windows.Forms.Label()
        Me.cboVendors = New System.Windows.Forms.ComboBox()
        Me.cboDept = New System.Windows.Forms.ComboBox()
        Me.fraDemand = New System.Windows.Forms.GroupBox()
        Me.SuspendLayout()
        '
        'lblTotal
        '
        Me.lblTotal.AutoSize = True
        Me.lblTotal.Location = New System.Drawing.Point(330, 57)
        Me.lblTotal.Name = "lblTotal"
        Me.lblTotal.Size = New System.Drawing.Size(39, 13)
        Me.lblTotal.TabIndex = 0
        Me.lblTotal.Text = "Label1"
        '
        'lblTotalOrdered
        '
        Me.lblTotalOrdered.AutoSize = True
        Me.lblTotalOrdered.Location = New System.Drawing.Point(334, 91)
        Me.lblTotalOrdered.Name = "lblTotalOrdered"
        Me.lblTotalOrdered.Size = New System.Drawing.Size(39, 13)
        Me.lblTotalOrdered.TabIndex = 1
        Me.lblTotalOrdered.Text = "Label2"
        '
        'dteStartDate
        '
        Me.dteStartDate.Location = New System.Drawing.Point(322, 147)
        Me.dteStartDate.Name = "dteStartDate"
        Me.dteStartDate.Size = New System.Drawing.Size(200, 20)
        Me.dteStartDate.TabIndex = 2
        '
        'dteEndDate
        '
        Me.dteEndDate.Location = New System.Drawing.Point(333, 197)
        Me.dteEndDate.Name = "dteEndDate"
        Me.dteEndDate.Size = New System.Drawing.Size(200, 20)
        Me.dteEndDate.TabIndex = 3
        '
        'cboStoreSelect
        '
        Me.cboStoreSelect.FormattingEnabled = True
        Me.cboStoreSelect.Location = New System.Drawing.Point(358, 252)
        Me.cboStoreSelect.Name = "cboStoreSelect"
        Me.cboStoreSelect.Size = New System.Drawing.Size(121, 21)
        Me.cboStoreSelect.TabIndex = 4
        '
        'cmdReset
        '
        Me.cmdReset.Location = New System.Drawing.Point(329, 310)
        Me.cmdReset.Name = "cmdReset"
        Me.cmdReset.Size = New System.Drawing.Size(75, 23)
        Me.cmdReset.TabIndex = 5
        Me.cmdReset.Text = "Button1"
        Me.cmdReset.UseVisualStyleBackColor = True
        '
        'lblVendor
        '
        Me.lblVendor.AutoSize = True
        Me.lblVendor.Location = New System.Drawing.Point(384, 364)
        Me.lblVendor.Name = "lblVendor"
        Me.lblVendor.Size = New System.Drawing.Size(39, 13)
        Me.lblVendor.TabIndex = 6
        Me.lblVendor.Text = "Label1"
        '
        'lblDept
        '
        Me.lblDept.AutoSize = True
        Me.lblDept.Location = New System.Drawing.Point(400, 395)
        Me.lblDept.Name = "lblDept"
        Me.lblDept.Size = New System.Drawing.Size(39, 13)
        Me.lblDept.TabIndex = 7
        Me.lblDept.Text = "Label2"
        '
        'cboVendors
        '
        Me.cboVendors.FormattingEnabled = True
        Me.cboVendors.Location = New System.Drawing.Point(500, 326)
        Me.cboVendors.Name = "cboVendors"
        Me.cboVendors.Size = New System.Drawing.Size(121, 21)
        Me.cboVendors.TabIndex = 8
        '
        'cboDept
        '
        Me.cboDept.FormattingEnabled = True
        Me.cboDept.Location = New System.Drawing.Point(500, 365)
        Me.cboDept.Name = "cboDept"
        Me.cboDept.Size = New System.Drawing.Size(121, 21)
        Me.cboDept.TabIndex = 9
        '
        'fraDemand
        '
        Me.fraDemand.Location = New System.Drawing.Point(496, 404)
        Me.fraDemand.Name = "fraDemand"
        Me.fraDemand.Size = New System.Drawing.Size(84, 35)
        Me.fraDemand.TabIndex = 10
        Me.fraDemand.TabStop = False
        Me.fraDemand.Text = "GroupBox1"
        '
        'InvAutoReOrder
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(800, 450)
        Me.Controls.Add(Me.fraDemand)
        Me.Controls.Add(Me.cboDept)
        Me.Controls.Add(Me.cboVendors)
        Me.Controls.Add(Me.lblDept)
        Me.Controls.Add(Me.lblVendor)
        Me.Controls.Add(Me.cmdReset)
        Me.Controls.Add(Me.cboStoreSelect)
        Me.Controls.Add(Me.dteEndDate)
        Me.Controls.Add(Me.dteStartDate)
        Me.Controls.Add(Me.lblTotalOrdered)
        Me.Controls.Add(Me.lblTotal)
        Me.Name = "InvAutoReOrder"
        Me.Text = "InvAutoReOrder"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents lblTotal As Label
    Friend WithEvents lblTotalOrdered As Label
    Friend WithEvents dteStartDate As DateTimePicker
    Friend WithEvents dteEndDate As DateTimePicker
    Friend WithEvents cboStoreSelect As ComboBox
    Friend WithEvents cmdReset As Button
    Friend WithEvents lblVendor As Label
    Friend WithEvents lblDept As Label
    Friend WithEvents cboVendors As ComboBox
    Friend WithEvents cboDept As ComboBox
    Friend WithEvents fraDemand As GroupBox
End Class
