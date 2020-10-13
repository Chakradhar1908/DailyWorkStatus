<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class InvPull
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
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
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.fraDel = New System.Windows.Forms.GroupBox()
        Me.chkSoldOrders = New System.Windows.Forms.CheckBox()
        Me.chkEmail = New System.Windows.Forms.CheckBox()
        Me.chkDriverCopy = New System.Windows.Forms.CheckBox()
        Me.chkShowCost = New System.Windows.Forms.CheckBox()
        Me.cboStore = New System.Windows.Forms.ComboBox()
        Me.txtTransferNo = New System.Windows.Forms.TextBox()
        Me.chkTransferNo = New System.Windows.Forms.CheckBox()
        Me.Juice = New System.Windows.Forms.TextBox()
        Me.lblJuice = New System.Windows.Forms.Label()
        Me.txtSaleNo = New System.Windows.Forms.TextBox()
        Me.optPrintAll2 = New System.Windows.Forms.RadioButton()
        Me.optPrintAll1 = New System.Windows.Forms.RadioButton()
        Me.optPrintAll0 = New System.Windows.Forms.RadioButton()
        Me.dteTo = New System.Windows.Forms.DateTimePicker()
        Me.dteFrom = New System.Windows.Forms.DateTimePicker()
        Me.imgLogo = New System.Windows.Forms.PictureBox()
        Me.txtTo = New System.Windows.Forms.Label()
        Me.txtFrom = New System.Windows.Forms.Label()
        Me.lblTo = New System.Windows.Forms.Label()
        Me.lblFrom = New System.Windows.Forms.Label()
        Me.fraControls = New System.Windows.Forms.GroupBox()
        Me.cmdCancel = New System.Windows.Forms.Button()
        Me.cmdPrint1 = New System.Windows.Forms.Button()
        Me.cmdPrint0 = New System.Windows.Forms.Button()
        Me.tmrEmail = New System.Windows.Forms.Timer(Me.components)
        Me.fraDel.SuspendLayout()
        CType(Me.imgLogo, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.fraControls.SuspendLayout()
        Me.SuspendLayout()
        '
        'fraDel
        '
        Me.fraDel.Controls.Add(Me.chkSoldOrders)
        Me.fraDel.Controls.Add(Me.chkEmail)
        Me.fraDel.Controls.Add(Me.chkDriverCopy)
        Me.fraDel.Controls.Add(Me.chkShowCost)
        Me.fraDel.Controls.Add(Me.cboStore)
        Me.fraDel.Controls.Add(Me.txtTransferNo)
        Me.fraDel.Controls.Add(Me.chkTransferNo)
        Me.fraDel.Controls.Add(Me.Juice)
        Me.fraDel.Controls.Add(Me.lblJuice)
        Me.fraDel.Controls.Add(Me.txtSaleNo)
        Me.fraDel.Controls.Add(Me.optPrintAll2)
        Me.fraDel.Controls.Add(Me.optPrintAll1)
        Me.fraDel.Controls.Add(Me.optPrintAll0)
        Me.fraDel.Controls.Add(Me.dteTo)
        Me.fraDel.Controls.Add(Me.dteFrom)
        Me.fraDel.Controls.Add(Me.imgLogo)
        Me.fraDel.Controls.Add(Me.txtTo)
        Me.fraDel.Controls.Add(Me.txtFrom)
        Me.fraDel.Controls.Add(Me.lblTo)
        Me.fraDel.Controls.Add(Me.lblFrom)
        Me.fraDel.Location = New System.Drawing.Point(6, 1)
        Me.fraDel.Name = "fraDel"
        Me.fraDel.Size = New System.Drawing.Size(324, 220)
        Me.fraDel.TabIndex = 4
        Me.fraDel.TabStop = False
        '
        'chkSoldOrders
        '
        Me.chkSoldOrders.AutoSize = True
        Me.chkSoldOrders.Location = New System.Drawing.Point(18, 199)
        Me.chkSoldOrders.Name = "chkSoldOrders"
        Me.chkSoldOrders.Size = New System.Drawing.Size(119, 17)
        Me.chkSoldOrders.TabIndex = 23
        Me.chkSoldOrders.Text = "Include So&ld Orders"
        Me.chkSoldOrders.UseVisualStyleBackColor = True
        '
        'chkEmail
        '
        Me.chkEmail.AutoSize = True
        Me.chkEmail.Location = New System.Drawing.Point(18, 188)
        Me.chkEmail.Name = "chkEmail"
        Me.chkEmail.Size = New System.Drawing.Size(51, 17)
        Me.chkEmail.TabIndex = 22
        Me.chkEmail.Text = "&Email"
        Me.chkEmail.UseVisualStyleBackColor = True
        '
        'chkDriverCopy
        '
        Me.chkDriverCopy.AutoSize = True
        Me.chkDriverCopy.Location = New System.Drawing.Point(18, 180)
        Me.chkDriverCopy.Name = "chkDriverCopy"
        Me.chkDriverCopy.Size = New System.Drawing.Size(112, 17)
        Me.chkDriverCopy.TabIndex = 21
        Me.chkDriverCopy.Text = "Print Dr&iver's Copy"
        Me.chkDriverCopy.UseVisualStyleBackColor = True
        '
        'chkShowCost
        '
        Me.chkShowCost.AutoSize = True
        Me.chkShowCost.Location = New System.Drawing.Point(18, 163)
        Me.chkShowCost.Name = "chkShowCost"
        Me.chkShowCost.Size = New System.Drawing.Size(77, 17)
        Me.chkShowCost.TabIndex = 20
        Me.chkShowCost.Text = "Show C&ost"
        Me.chkShowCost.UseVisualStyleBackColor = True
        '
        'cboStore
        '
        Me.cboStore.FormattingEnabled = True
        Me.cboStore.Location = New System.Drawing.Point(172, 136)
        Me.cboStore.Name = "cboStore"
        Me.cboStore.Size = New System.Drawing.Size(121, 21)
        Me.cboStore.TabIndex = 19
        Me.cboStore.Visible = False
        '
        'txtTransferNo
        '
        Me.txtTransferNo.Location = New System.Drawing.Point(186, 155)
        Me.txtTransferNo.Name = "txtTransferNo"
        Me.txtTransferNo.Size = New System.Drawing.Size(100, 20)
        Me.txtTransferNo.TabIndex = 18
        Me.txtTransferNo.Visible = False
        '
        'chkTransferNo
        '
        Me.chkTransferNo.AutoSize = True
        Me.chkTransferNo.Location = New System.Drawing.Point(172, 140)
        Me.chkTransferNo.Name = "chkTransferNo"
        Me.chkTransferNo.Size = New System.Drawing.Size(102, 17)
        Me.chkTransferNo.TabIndex = 17
        Me.chkTransferNo.Text = "Print Transfer #:"
        Me.chkTransferNo.UseVisualStyleBackColor = True
        '
        'Juice
        '
        Me.Juice.Location = New System.Drawing.Point(216, 104)
        Me.Juice.Name = "Juice"
        Me.Juice.Size = New System.Drawing.Size(70, 20)
        Me.Juice.TabIndex = 16
        '
        'lblJuice
        '
        Me.lblJuice.AutoSize = True
        Me.lblJuice.Location = New System.Drawing.Point(169, 104)
        Me.lblJuice.Name = "lblJuice"
        Me.lblJuice.Size = New System.Drawing.Size(46, 13)
        Me.lblJuice.TabIndex = 15
        Me.lblJuice.Text = "Juice %:"
        '
        'txtSaleNo
        '
        Me.txtSaleNo.Location = New System.Drawing.Point(63, 137)
        Me.txtSaleNo.Name = "txtSaleNo"
        Me.txtSaleNo.Size = New System.Drawing.Size(70, 20)
        Me.txtSaleNo.TabIndex = 14
        '
        'optPrintAll2
        '
        Me.optPrintAll2.AutoSize = True
        Me.optPrintAll2.Location = New System.Drawing.Point(18, 138)
        Me.optPrintAll2.Name = "optPrintAll2"
        Me.optPrintAll2.Size = New System.Drawing.Size(49, 17)
        Me.optPrintAll2.TabIndex = 13
        Me.optPrintAll2.TabStop = True
        Me.optPrintAll2.Text = "&Sale:"
        Me.optPrintAll2.UseVisualStyleBackColor = True
        '
        'optPrintAll1
        '
        Me.optPrintAll1.AutoSize = True
        Me.optPrintAll1.Location = New System.Drawing.Point(18, 118)
        Me.optPrintAll1.Name = "optPrintAll1"
        Me.optPrintAll1.Size = New System.Drawing.Size(73, 17)
        Me.optPrintAll1.TabIndex = 12
        Me.optPrintAll1.TabStop = True
        Me.optPrintAll1.Text = "Reprint &All"
        Me.optPrintAll1.UseVisualStyleBackColor = True
        '
        'optPrintAll0
        '
        Me.optPrintAll0.AutoSize = True
        Me.optPrintAll0.Location = New System.Drawing.Point(18, 100)
        Me.optPrintAll0.Name = "optPrintAll0"
        Me.optPrintAll0.Size = New System.Drawing.Size(95, 17)
        Me.optPrintAll0.TabIndex = 11
        Me.optPrintAll0.TabStop = True
        Me.optPrintAll0.Text = "Print &Unprinted"
        Me.optPrintAll0.UseVisualStyleBackColor = True
        '
        'dteTo
        '
        Me.dteTo.CustomFormat = "MM/dd/yyyy"
        Me.dteTo.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dteTo.Location = New System.Drawing.Point(172, 74)
        Me.dteTo.Name = "dteTo"
        Me.dteTo.Size = New System.Drawing.Size(114, 26)
        Me.dteTo.TabIndex = 10
        '
        'dteFrom
        '
        Me.dteFrom.CustomFormat = "MM/dd/yyyy"
        Me.dteFrom.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dteFrom.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dteFrom.Location = New System.Drawing.Point(18, 74)
        Me.dteFrom.Name = "dteFrom"
        Me.dteFrom.Size = New System.Drawing.Size(115, 26)
        Me.dteFrom.TabIndex = 9
        '
        'imgLogo
        '
        Me.imgLogo.Location = New System.Drawing.Point(128, 30)
        Me.imgLogo.Name = "imgLogo"
        Me.imgLogo.Size = New System.Drawing.Size(35, 33)
        Me.imgLogo.TabIndex = 8
        Me.imgLogo.TabStop = False
        Me.imgLogo.Visible = False
        '
        'txtTo
        '
        Me.txtTo.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.txtTo.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtTo.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtTo.Location = New System.Drawing.Point(172, 30)
        Me.txtTo.Name = "txtTo"
        Me.txtTo.Size = New System.Drawing.Size(100, 40)
        Me.txtTo.TabIndex = 7
        Me.txtTo.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'txtFrom
        '
        Me.txtFrom.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.txtFrom.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtFrom.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtFrom.Location = New System.Drawing.Point(18, 30)
        Me.txtFrom.Name = "txtFrom"
        Me.txtFrom.Size = New System.Drawing.Size(100, 40)
        Me.txtFrom.TabIndex = 6
        Me.txtFrom.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'lblTo
        '
        Me.lblTo.AutoSize = True
        Me.lblTo.Location = New System.Drawing.Point(169, 16)
        Me.lblTo.Name = "lblTo"
        Me.lblTo.Size = New System.Drawing.Size(23, 13)
        Me.lblTo.TabIndex = 5
        Me.lblTo.Text = "&To:"
        '
        'lblFrom
        '
        Me.lblFrom.AutoSize = True
        Me.lblFrom.Location = New System.Drawing.Point(15, 16)
        Me.lblFrom.Name = "lblFrom"
        Me.lblFrom.Size = New System.Drawing.Size(33, 13)
        Me.lblFrom.TabIndex = 4
        Me.lblFrom.Text = "&From:"
        '
        'fraControls
        '
        Me.fraControls.Controls.Add(Me.cmdCancel)
        Me.fraControls.Controls.Add(Me.cmdPrint1)
        Me.fraControls.Controls.Add(Me.cmdPrint0)
        Me.fraControls.Location = New System.Drawing.Point(6, 226)
        Me.fraControls.Name = "fraControls"
        Me.fraControls.Size = New System.Drawing.Size(255, 68)
        Me.fraControls.TabIndex = 5
        Me.fraControls.TabStop = False
        '
        'cmdCancel
        '
        Me.cmdCancel.Location = New System.Drawing.Point(172, 19)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.Size = New System.Drawing.Size(75, 41)
        Me.cmdCancel.TabIndex = 2
        Me.cmdCancel.Text = "&Cancel"
        Me.cmdCancel.UseVisualStyleBackColor = True
        '
        'cmdPrint1
        '
        Me.cmdPrint1.Location = New System.Drawing.Point(88, 19)
        Me.cmdPrint1.Name = "cmdPrint1"
        Me.cmdPrint1.Size = New System.Drawing.Size(75, 41)
        Me.cmdPrint1.TabIndex = 1
        Me.cmdPrint1.Text = "P&review"
        Me.cmdPrint1.UseVisualStyleBackColor = True
        '
        'cmdPrint0
        '
        Me.cmdPrint0.Location = New System.Drawing.Point(6, 19)
        Me.cmdPrint0.Name = "cmdPrint0"
        Me.cmdPrint0.Size = New System.Drawing.Size(75, 41)
        Me.cmdPrint0.TabIndex = 0
        Me.cmdPrint0.Text = "&Print"
        Me.cmdPrint0.UseVisualStyleBackColor = True
        '
        'InvPull
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(335, 301)
        Me.Controls.Add(Me.fraControls)
        Me.Controls.Add(Me.fraDel)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.MaximizeBox = False
        Me.Name = "InvPull"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Customer Deliveries"
        Me.fraDel.ResumeLayout(False)
        Me.fraDel.PerformLayout()
        CType(Me.imgLogo, System.ComponentModel.ISupportInitialize).EndInit()
        Me.fraControls.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents fraDel As GroupBox
    Friend WithEvents chkSoldOrders As CheckBox
    Friend WithEvents chkEmail As CheckBox
    Friend WithEvents chkDriverCopy As CheckBox
    Friend WithEvents chkShowCost As CheckBox
    Friend WithEvents cboStore As ComboBox
    Friend WithEvents txtTransferNo As TextBox
    Friend WithEvents chkTransferNo As CheckBox
    Friend WithEvents Juice As TextBox
    Friend WithEvents lblJuice As Label
    Friend WithEvents txtSaleNo As TextBox
    Friend WithEvents optPrintAll2 As RadioButton
    Friend WithEvents optPrintAll1 As RadioButton
    Friend WithEvents optPrintAll0 As RadioButton
    Friend WithEvents dteTo As DateTimePicker
    Friend WithEvents dteFrom As DateTimePicker
    Friend WithEvents imgLogo As PictureBox
    Friend WithEvents txtTo As Label
    Friend WithEvents txtFrom As Label
    Friend WithEvents lblTo As Label
    Friend WithEvents lblFrom As Label
    Friend WithEvents fraControls As GroupBox
    Friend WithEvents cmdCancel As Button
    Friend WithEvents cmdPrint1 As Button
    Friend WithEvents cmdPrint0 As Button
    Friend WithEvents tmrEmail As Timer
End Class
