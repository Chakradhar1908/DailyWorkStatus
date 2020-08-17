<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmCashRegisterQuantity
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
        Me.cboChargeType = New System.Windows.Forms.ComboBox()
        Me.lblItem = New System.Windows.Forms.Label()
        Me.lblDescCap = New System.Windows.Forms.Label()
        Me.lblDesc = New System.Windows.Forms.Label()
        Me.lblSwipe = New System.Windows.Forms.Label()
        Me.cmdApply = New System.Windows.Forms.Button()
        Me.cmdSwipe = New System.Windows.Forms.Button()
        Me.txtSwipe = New System.Windows.Forms.TextBox()
        Me.txtPrice = New System.Windows.Forms.TextBox()
        Me.fraDiscType = New System.Windows.Forms.GroupBox()
        Me.lblStyle = New System.Windows.Forms.Label()
        Me.lblPrice = New System.Windows.Forms.Label()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.ShapeContainer1 = New Microsoft.VisualBasic.PowerPacks.ShapeContainer()
        Me.shpSwipe = New Microsoft.VisualBasic.PowerPacks.RectangleShape()
        Me.cmdTax = New System.Windows.Forms.Button()
        Me.lblQuan = New System.Windows.Forms.Label()
        Me.txtQuantity = New System.Windows.Forms.TextBox()
        Me.cmdCancel = New System.Windows.Forms.Button()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.optDiscType0 = New System.Windows.Forms.RadioButton()
        Me.optDiscType1 = New System.Windows.Forms.RadioButton()
        Me.optDiscType2 = New System.Windows.Forms.RadioButton()
        Me.tmrLocate = New System.Windows.Forms.Timer(Me.components)
        Me.tmrSwipe = New System.Windows.Forms.Timer(Me.components)
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'cboChargeType
        '
        Me.cboChargeType.FormattingEnabled = True
        Me.cboChargeType.Location = New System.Drawing.Point(0, 0)
        Me.cboChargeType.Name = "cboChargeType"
        Me.cboChargeType.Size = New System.Drawing.Size(121, 21)
        Me.cboChargeType.TabIndex = 0
        '
        'lblItem
        '
        Me.lblItem.AutoSize = True
        Me.lblItem.Location = New System.Drawing.Point(22, 73)
        Me.lblItem.Name = "lblItem"
        Me.lblItem.Size = New System.Drawing.Size(39, 13)
        Me.lblItem.TabIndex = 1
        Me.lblItem.Text = "Label1"
        '
        'lblDescCap
        '
        Me.lblDescCap.AutoSize = True
        Me.lblDescCap.Location = New System.Drawing.Point(22, 98)
        Me.lblDescCap.Name = "lblDescCap"
        Me.lblDescCap.Size = New System.Drawing.Size(39, 13)
        Me.lblDescCap.TabIndex = 2
        Me.lblDescCap.Text = "Label2"
        '
        'lblDesc
        '
        Me.lblDesc.AutoSize = True
        Me.lblDesc.Location = New System.Drawing.Point(22, 122)
        Me.lblDesc.Name = "lblDesc"
        Me.lblDesc.Size = New System.Drawing.Size(39, 13)
        Me.lblDesc.TabIndex = 3
        Me.lblDesc.Text = "Label3"
        '
        'lblSwipe
        '
        Me.lblSwipe.AutoSize = True
        Me.lblSwipe.Location = New System.Drawing.Point(22, 147)
        Me.lblSwipe.Name = "lblSwipe"
        Me.lblSwipe.Size = New System.Drawing.Size(39, 13)
        Me.lblSwipe.TabIndex = 4
        Me.lblSwipe.Text = "Label4"
        '
        'cmdApply
        '
        Me.cmdApply.Location = New System.Drawing.Point(25, 208)
        Me.cmdApply.Name = "cmdApply"
        Me.cmdApply.Size = New System.Drawing.Size(75, 23)
        Me.cmdApply.TabIndex = 5
        Me.cmdApply.Text = "Button1"
        Me.cmdApply.UseVisualStyleBackColor = True
        '
        'cmdSwipe
        '
        Me.cmdSwipe.Location = New System.Drawing.Point(25, 237)
        Me.cmdSwipe.Name = "cmdSwipe"
        Me.cmdSwipe.Size = New System.Drawing.Size(75, 23)
        Me.cmdSwipe.TabIndex = 6
        Me.cmdSwipe.Text = "Button1"
        Me.cmdSwipe.UseVisualStyleBackColor = True
        '
        'txtSwipe
        '
        Me.txtSwipe.Location = New System.Drawing.Point(25, 281)
        Me.txtSwipe.Name = "txtSwipe"
        Me.txtSwipe.Size = New System.Drawing.Size(100, 20)
        Me.txtSwipe.TabIndex = 7
        '
        'txtPrice
        '
        Me.txtPrice.Location = New System.Drawing.Point(25, 307)
        Me.txtPrice.Name = "txtPrice"
        Me.txtPrice.Size = New System.Drawing.Size(100, 20)
        Me.txtPrice.TabIndex = 8
        '
        'fraDiscType
        '
        Me.fraDiscType.Location = New System.Drawing.Point(25, 338)
        Me.fraDiscType.Name = "fraDiscType"
        Me.fraDiscType.Size = New System.Drawing.Size(200, 100)
        Me.fraDiscType.TabIndex = 9
        Me.fraDiscType.TabStop = False
        Me.fraDiscType.Text = "GroupBox1"
        '
        'lblStyle
        '
        Me.lblStyle.AutoSize = True
        Me.lblStyle.Location = New System.Drawing.Point(22, 170)
        Me.lblStyle.Name = "lblStyle"
        Me.lblStyle.Size = New System.Drawing.Size(39, 13)
        Me.lblStyle.TabIndex = 0
        Me.lblStyle.Text = "Label1"
        '
        'lblPrice
        '
        Me.lblPrice.AutoSize = True
        Me.lblPrice.Location = New System.Drawing.Point(82, 73)
        Me.lblPrice.Name = "lblPrice"
        Me.lblPrice.Size = New System.Drawing.Size(39, 13)
        Me.lblPrice.TabIndex = 10
        Me.lblPrice.Text = "Label1"
        '
        'ShapeContainer1
        '
        Me.ShapeContainer1.Location = New System.Drawing.Point(0, 0)
        Me.ShapeContainer1.Margin = New System.Windows.Forms.Padding(0)
        Me.ShapeContainer1.Name = "ShapeContainer1"
        Me.ShapeContainer1.Shapes.AddRange(New Microsoft.VisualBasic.PowerPacks.Shape() {Me.shpSwipe})
        Me.ShapeContainer1.Size = New System.Drawing.Size(800, 450)
        Me.ShapeContainer1.TabIndex = 11
        Me.ShapeContainer1.TabStop = False
        '
        'shpSwipe
        '
        Me.shpSwipe.Location = New System.Drawing.Point(246, 342)
        Me.shpSwipe.Name = "shpSwipe"
        Me.shpSwipe.Size = New System.Drawing.Size(117, 74)
        '
        'cmdTax
        '
        Me.cmdTax.Location = New System.Drawing.Point(386, 342)
        Me.cmdTax.Name = "cmdTax"
        Me.cmdTax.Size = New System.Drawing.Size(75, 23)
        Me.cmdTax.TabIndex = 12
        Me.cmdTax.Text = "Button1"
        Me.cmdTax.UseVisualStyleBackColor = True
        '
        'lblQuan
        '
        Me.lblQuan.AutoSize = True
        Me.lblQuan.Location = New System.Drawing.Point(22, 192)
        Me.lblQuan.Name = "lblQuan"
        Me.lblQuan.Size = New System.Drawing.Size(39, 13)
        Me.lblQuan.TabIndex = 13
        Me.lblQuan.Text = "Label1"
        '
        'txtQuantity
        '
        Me.txtQuantity.Location = New System.Drawing.Point(131, 281)
        Me.txtQuantity.Name = "txtQuantity"
        Me.txtQuantity.Size = New System.Drawing.Size(100, 20)
        Me.txtQuantity.TabIndex = 14
        '
        'cmdCancel
        '
        Me.cmdCancel.Location = New System.Drawing.Point(386, 371)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.Size = New System.Drawing.Size(75, 23)
        Me.cmdCancel.TabIndex = 15
        Me.cmdCancel.Text = "Button1"
        Me.cmdCancel.UseVisualStyleBackColor = True
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.optDiscType2)
        Me.GroupBox1.Controls.Add(Me.optDiscType1)
        Me.GroupBox1.Controls.Add(Me.optDiscType0)
        Me.GroupBox1.Location = New System.Drawing.Point(260, 69)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(148, 161)
        Me.GroupBox1.TabIndex = 16
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "GroupBox1"
        '
        'optDiscType0
        '
        Me.optDiscType0.AutoSize = True
        Me.optDiscType0.Location = New System.Drawing.Point(24, 33)
        Me.optDiscType0.Name = "optDiscType0"
        Me.optDiscType0.Size = New System.Drawing.Size(73, 17)
        Me.optDiscType0.TabIndex = 0
        Me.optDiscType0.TabStop = True
        Me.optDiscType0.Text = "Prev. Item"
        Me.optDiscType0.UseVisualStyleBackColor = True
        '
        'optDiscType1
        '
        Me.optDiscType1.AutoSize = True
        Me.optDiscType1.Location = New System.Drawing.Point(31, 71)
        Me.optDiscType1.Name = "optDiscType1"
        Me.optDiscType1.Size = New System.Drawing.Size(76, 17)
        Me.optDiscType1.TabIndex = 1
        Me.optDiscType1.TabStop = True
        Me.optDiscType1.Text = "Entire Sale"
        Me.optDiscType1.UseVisualStyleBackColor = True
        '
        'optDiscType2
        '
        Me.optDiscType2.AutoSize = True
        Me.optDiscType2.Location = New System.Drawing.Point(31, 110)
        Me.optDiscType2.Name = "optDiscType2"
        Me.optDiscType2.Size = New System.Drawing.Size(61, 17)
        Me.optDiscType2.TabIndex = 2
        Me.optDiscType2.TabStop = True
        Me.optDiscType2.Text = "Amount"
        Me.optDiscType2.UseVisualStyleBackColor = True
        '
        'tmrLocate
        '
        Me.tmrLocate.Interval = 1000
        '
        'tmrSwipe
        '
        Me.tmrSwipe.Interval = 300
        '
        'frmCashRegisterQuantity
        '
        Me.AcceptButton = Me.cmdApply
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(800, 450)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.cmdCancel)
        Me.Controls.Add(Me.txtQuantity)
        Me.Controls.Add(Me.lblQuan)
        Me.Controls.Add(Me.cmdTax)
        Me.Controls.Add(Me.lblPrice)
        Me.Controls.Add(Me.lblStyle)
        Me.Controls.Add(Me.fraDiscType)
        Me.Controls.Add(Me.txtPrice)
        Me.Controls.Add(Me.txtSwipe)
        Me.Controls.Add(Me.cmdSwipe)
        Me.Controls.Add(Me.cmdApply)
        Me.Controls.Add(Me.lblSwipe)
        Me.Controls.Add(Me.lblDesc)
        Me.Controls.Add(Me.lblDescCap)
        Me.Controls.Add(Me.lblItem)
        Me.Controls.Add(Me.cboChargeType)
        Me.Controls.Add(Me.ShapeContainer1)
        Me.Name = "frmCashRegisterQuantity"
        Me.Text = "frmCashRegisterQuantity"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents cboChargeType As ComboBox
    Friend WithEvents lblItem As Label
    Friend WithEvents lblDescCap As Label
    Friend WithEvents lblDesc As Label
    Friend WithEvents lblSwipe As Label
    Friend WithEvents cmdApply As Button
    Friend WithEvents cmdSwipe As Button
    Friend WithEvents txtSwipe As TextBox
    Friend WithEvents txtPrice As TextBox
    Friend WithEvents fraDiscType As GroupBox
    Friend WithEvents lblStyle As Label
    Friend WithEvents lblPrice As Label
    Friend WithEvents ToolTip1 As ToolTip
    Friend WithEvents ShapeContainer1 As PowerPacks.ShapeContainer
    Friend WithEvents shpSwipe As PowerPacks.RectangleShape
    Friend WithEvents cmdTax As Button
    Friend WithEvents lblQuan As Label
    Friend WithEvents txtQuantity As TextBox
    Friend WithEvents cmdCancel As Button
    Friend WithEvents GroupBox1 As GroupBox
    Friend WithEvents optDiscType2 As RadioButton
    Friend WithEvents optDiscType1 As RadioButton
    Friend WithEvents optDiscType0 As RadioButton
    Friend WithEvents tmrLocate As Timer
    Friend WithEvents tmrSwipe As Timer
End Class
