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
        Me.lblStyle = New System.Windows.Forms.Label()
        Me.lblPrice = New System.Windows.Forms.Label()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.txtQuantity = New System.Windows.Forms.TextBox()
        Me.cmdCancel = New System.Windows.Forms.Button()
        Me.ShapeContainer1 = New Microsoft.VisualBasic.PowerPacks.ShapeContainer()
        Me.shpSwipe = New Microsoft.VisualBasic.PowerPacks.RectangleShape()
        Me.cmdTax = New System.Windows.Forms.Button()
        Me.lblQuan = New System.Windows.Forms.Label()
        Me.fraDiscType = New System.Windows.Forms.GroupBox()
        Me.optDiscType2 = New System.Windows.Forms.RadioButton()
        Me.optDiscType1 = New System.Windows.Forms.RadioButton()
        Me.optDiscType0 = New System.Windows.Forms.RadioButton()
        Me.tmrLocate = New System.Windows.Forms.Timer(Me.components)
        Me.tmrSwipe = New System.Windows.Forms.Timer(Me.components)
        Me.imgCheat = New System.Windows.Forms.PictureBox()
        Me.lblNext = New System.Windows.Forms.Label()
        Me.fraDiscType.SuspendLayout()
        CType(Me.imgCheat, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'cboChargeType
        '
        Me.cboChargeType.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboChargeType.FormattingEnabled = True
        Me.cboChargeType.Location = New System.Drawing.Point(125, 26)
        Me.cboChargeType.Name = "cboChargeType"
        Me.cboChargeType.Size = New System.Drawing.Size(203, 28)
        Me.cboChargeType.TabIndex = 0
        Me.ToolTip1.SetToolTip(Me.cboChargeType, "Select the type of credit card.")
        Me.cboChargeType.Visible = False
        '
        'lblItem
        '
        Me.lblItem.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblItem.Location = New System.Drawing.Point(13, 28)
        Me.lblItem.Name = "lblItem"
        Me.lblItem.Size = New System.Drawing.Size(106, 26)
        Me.lblItem.TabIndex = 1
        Me.lblItem.Text = "Item:"
        '
        'lblDescCap
        '
        Me.lblDescCap.AutoSize = True
        Me.lblDescCap.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblDescCap.Location = New System.Drawing.Point(13, 60)
        Me.lblDescCap.Name = "lblDescCap"
        Me.lblDescCap.Size = New System.Drawing.Size(105, 20)
        Me.lblDescCap.TabIndex = 2
        Me.lblDescCap.Text = "Description:"
        '
        'lblDesc
        '
        Me.lblDesc.Location = New System.Drawing.Point(128, 63)
        Me.lblDesc.Name = "lblDesc"
        Me.lblDesc.Size = New System.Drawing.Size(200, 48)
        Me.lblDesc.TabIndex = 3
        Me.lblDesc.Text = "lblDesc"
        '
        'lblSwipe
        '
        Me.lblSwipe.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblSwipe.Location = New System.Drawing.Point(5, 26)
        Me.lblSwipe.Name = "lblSwipe"
        Me.lblSwipe.Size = New System.Drawing.Size(306, 22)
        Me.lblSwipe.TabIndex = 4
        Me.lblSwipe.Text = "### SWIPE ### SWIPE ### SWIPE ###"
        Me.lblSwipe.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.lblSwipe.Visible = False
        '
        'cmdApply
        '
        Me.cmdApply.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cmdApply.Location = New System.Drawing.Point(85, 181)
        Me.cmdApply.Name = "cmdApply"
        Me.cmdApply.Size = New System.Drawing.Size(75, 47)
        Me.cmdApply.TabIndex = 5
        Me.cmdApply.Text = "&Apply"
        Me.cmdApply.UseVisualStyleBackColor = True
        '
        'cmdSwipe
        '
        Me.cmdSwipe.Location = New System.Drawing.Point(85, 181)
        Me.cmdSwipe.Name = "cmdSwipe"
        Me.cmdSwipe.Size = New System.Drawing.Size(75, 47)
        Me.cmdSwipe.TabIndex = 6
        Me.cmdSwipe.Text = "S&wipe"
        Me.ToolTip1.SetToolTip(Me.cmdSwipe, "Click here to add this item to the sale.")
        Me.cmdSwipe.UseVisualStyleBackColor = True
        '
        'txtSwipe
        '
        Me.txtSwipe.Location = New System.Drawing.Point(3, 3)
        Me.txtSwipe.Name = "txtSwipe"
        Me.txtSwipe.Size = New System.Drawing.Size(41, 20)
        Me.txtSwipe.TabIndex = 7
        Me.txtSwipe.Visible = False
        '
        'txtPrice
        '
        Me.txtPrice.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtPrice.Location = New System.Drawing.Point(128, 146)
        Me.txtPrice.Name = "txtPrice"
        Me.txtPrice.Size = New System.Drawing.Size(77, 26)
        Me.txtPrice.TabIndex = 8
        Me.txtPrice.Text = "0.00"
        Me.txtPrice.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'lblStyle
        '
        Me.lblStyle.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblStyle.Location = New System.Drawing.Point(125, 26)
        Me.lblStyle.Name = "lblStyle"
        Me.lblStyle.Size = New System.Drawing.Size(203, 28)
        Me.lblStyle.TabIndex = 0
        Me.lblStyle.Text = "lblStyle"
        '
        'lblPrice
        '
        Me.lblPrice.AutoSize = True
        Me.lblPrice.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblPrice.Location = New System.Drawing.Point(13, 152)
        Me.lblPrice.Name = "lblPrice"
        Me.lblPrice.Size = New System.Drawing.Size(92, 20)
        Me.lblPrice.TabIndex = 10
        Me.lblPrice.Text = "Unit Price:"
        '
        'txtQuantity
        '
        Me.txtQuantity.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtQuantity.Location = New System.Drawing.Point(128, 114)
        Me.txtQuantity.Name = "txtQuantity"
        Me.txtQuantity.Size = New System.Drawing.Size(77, 26)
        Me.txtQuantity.TabIndex = 14
        Me.txtQuantity.Text = "0"
        Me.txtQuantity.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.ToolTip1.SetToolTip(Me.txtQuantity, "Enter a quantity.  This value must be greater than zero.")
        '
        'cmdCancel
        '
        Me.cmdCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cmdCancel.Location = New System.Drawing.Point(166, 181)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.Size = New System.Drawing.Size(75, 47)
        Me.cmdCancel.TabIndex = 15
        Me.cmdCancel.Text = "&Cancel"
        Me.ToolTip1.SetToolTip(Me.cmdCancel, "Click to cancel this item.  It will not be added to the sale.")
        Me.cmdCancel.UseVisualStyleBackColor = True
        '
        'ShapeContainer1
        '
        Me.ShapeContainer1.Location = New System.Drawing.Point(0, 0)
        Me.ShapeContainer1.Margin = New System.Windows.Forms.Padding(0)
        Me.ShapeContainer1.Name = "ShapeContainer1"
        Me.ShapeContainer1.Shapes.AddRange(New Microsoft.VisualBasic.PowerPacks.Shape() {Me.shpSwipe})
        Me.ShapeContainer1.Size = New System.Drawing.Size(341, 256)
        Me.ShapeContainer1.TabIndex = 11
        Me.ShapeContainer1.TabStop = False
        '
        'shpSwipe
        '
        Me.shpSwipe.BackColor = System.Drawing.Color.Lime
        Me.shpSwipe.FillColor = System.Drawing.Color.Yellow
        Me.shpSwipe.FillStyle = Microsoft.VisualBasic.PowerPacks.FillStyle.Solid
        Me.shpSwipe.Location = New System.Drawing.Point(226, 111)
        Me.shpSwipe.Name = "shpSwipe"
        Me.shpSwipe.Size = New System.Drawing.Size(11, 12)
        Me.shpSwipe.Visible = False
        '
        'cmdTax
        '
        Me.cmdTax.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdTax.Location = New System.Drawing.Point(211, 146)
        Me.cmdTax.Name = "cmdTax"
        Me.cmdTax.Size = New System.Drawing.Size(80, 26)
        Me.cmdTax.TabIndex = 12
        Me.cmdTax.Text = "&Taxable"
        Me.cmdTax.UseVisualStyleBackColor = True
        '
        'lblQuan
        '
        Me.lblQuan.AutoSize = True
        Me.lblQuan.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblQuan.Location = New System.Drawing.Point(13, 117)
        Me.lblQuan.Name = "lblQuan"
        Me.lblQuan.Size = New System.Drawing.Size(81, 20)
        Me.lblQuan.TabIndex = 13
        Me.lblQuan.Text = "Quantity:"
        '
        'fraDiscType
        '
        Me.fraDiscType.Controls.Add(Me.optDiscType2)
        Me.fraDiscType.Controls.Add(Me.optDiscType1)
        Me.fraDiscType.Controls.Add(Me.optDiscType0)
        Me.fraDiscType.Location = New System.Drawing.Point(239, 57)
        Me.fraDiscType.Name = "fraDiscType"
        Me.fraDiscType.Size = New System.Drawing.Size(95, 88)
        Me.fraDiscType.TabIndex = 16
        Me.fraDiscType.TabStop = False
        Me.fraDiscType.Text = "Discount Type:"
        Me.fraDiscType.Visible = False
        '
        'optDiscType2
        '
        Me.optDiscType2.AutoSize = True
        Me.optDiscType2.Location = New System.Drawing.Point(6, 61)
        Me.optDiscType2.Name = "optDiscType2"
        Me.optDiscType2.Size = New System.Drawing.Size(61, 17)
        Me.optDiscType2.TabIndex = 2
        Me.optDiscType2.TabStop = True
        Me.optDiscType2.Text = "Amount"
        Me.optDiscType2.UseVisualStyleBackColor = True
        '
        'optDiscType1
        '
        Me.optDiscType1.AutoSize = True
        Me.optDiscType1.Location = New System.Drawing.Point(6, 40)
        Me.optDiscType1.Name = "optDiscType1"
        Me.optDiscType1.Size = New System.Drawing.Size(76, 17)
        Me.optDiscType1.TabIndex = 1
        Me.optDiscType1.TabStop = True
        Me.optDiscType1.Text = "Entire Sale"
        Me.optDiscType1.UseVisualStyleBackColor = True
        '
        'optDiscType0
        '
        Me.optDiscType0.AutoSize = True
        Me.optDiscType0.Location = New System.Drawing.Point(6, 19)
        Me.optDiscType0.Name = "optDiscType0"
        Me.optDiscType0.Size = New System.Drawing.Size(73, 17)
        Me.optDiscType0.TabIndex = 0
        Me.optDiscType0.TabStop = True
        Me.optDiscType0.Text = "Prev. Item"
        Me.optDiscType0.UseVisualStyleBackColor = True
        '
        'tmrLocate
        '
        Me.tmrLocate.Interval = 1000
        '
        'tmrSwipe
        '
        Me.tmrSwipe.Interval = 300
        '
        'imgCheat
        '
        Me.imgCheat.Location = New System.Drawing.Point(277, 3)
        Me.imgCheat.Name = "imgCheat"
        Me.imgCheat.Size = New System.Drawing.Size(41, 39)
        Me.imgCheat.TabIndex = 17
        Me.imgCheat.TabStop = False
        '
        'lblNext
        '
        Me.lblNext.AutoSize = True
        Me.lblNext.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblNext.Location = New System.Drawing.Point(52, 231)
        Me.lblNext.Name = "lblNext"
        Me.lblNext.Size = New System.Drawing.Size(243, 24)
        Me.lblNext.TabIndex = 18
        Me.lblNext.Text = "NEXT form swipes card! "
        Me.lblNext.Visible = False
        '
        'frmCashRegisterQuantity
        '
        Me.AcceptButton = Me.cmdSwipe
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.cmdCancel
        Me.ClientSize = New System.Drawing.Size(341, 256)
        Me.Controls.Add(Me.fraDiscType)
        Me.Controls.Add(Me.lblDesc)
        Me.Controls.Add(Me.cboChargeType)
        Me.Controls.Add(Me.lblNext)
        Me.Controls.Add(Me.imgCheat)
        Me.Controls.Add(Me.cmdCancel)
        Me.Controls.Add(Me.txtQuantity)
        Me.Controls.Add(Me.lblQuan)
        Me.Controls.Add(Me.cmdTax)
        Me.Controls.Add(Me.lblPrice)
        Me.Controls.Add(Me.txtPrice)
        Me.Controls.Add(Me.txtSwipe)
        Me.Controls.Add(Me.cmdSwipe)
        Me.Controls.Add(Me.cmdApply)
        Me.Controls.Add(Me.lblDescCap)
        Me.Controls.Add(Me.lblSwipe)
        Me.Controls.Add(Me.lblItem)
        Me.Controls.Add(Me.lblStyle)
        Me.Controls.Add(Me.ShapeContainer1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmCashRegisterQuantity"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Enter Quantity"
        Me.fraDiscType.ResumeLayout(False)
        Me.fraDiscType.PerformLayout()
        CType(Me.imgCheat, System.ComponentModel.ISupportInitialize).EndInit()
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
    Friend WithEvents lblStyle As Label
    Friend WithEvents lblPrice As Label
    Friend WithEvents ToolTip1 As ToolTip
    Friend WithEvents ShapeContainer1 As PowerPacks.ShapeContainer
    Friend WithEvents shpSwipe As PowerPacks.RectangleShape
    Friend WithEvents cmdTax As Button
    Friend WithEvents lblQuan As Label
    Friend WithEvents txtQuantity As TextBox
    Friend WithEvents cmdCancel As Button
    Friend WithEvents fraDiscType As GroupBox
    Friend WithEvents optDiscType2 As RadioButton
    Friend WithEvents optDiscType1 As RadioButton
    Friend WithEvents optDiscType0 As RadioButton
    Friend WithEvents tmrLocate As Timer
    Friend WithEvents tmrSwipe As Timer
    Friend WithEvents imgCheat As PictureBox
    Friend WithEvents lblNext As Label
End Class
