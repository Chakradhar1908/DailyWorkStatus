<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class frmCashRegister
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
        Me.lblCust = New System.Windows.Forms.Label()
        Me.cmdComm = New System.Windows.Forms.Button()
        Me.cboSalesList = New System.Windows.Forms.ComboBox()
        Me.lblTotal = New System.Windows.Forms.Label()
        Me.lblTax = New System.Windows.Forms.Label()
        Me.lblTendered = New System.Windows.Forms.Label()
        Me.lblDue = New System.Windows.Forms.Label()
        Me.txtSku = New System.Windows.Forms.TextBox()
        Me.fraSaleButtons = New System.Windows.Forms.GroupBox()
        Me.cmdPayment = New System.Windows.Forms.Button()
        Me.cmdReturn = New System.Windows.Forms.Button()
        Me.cmdDiscount = New System.Windows.Forms.Button()
        Me.cmdPrint = New System.Windows.Forms.Button()
        Me.cmdDone = New System.Windows.Forms.Button()
        Me.cmdCancelSale = New System.Windows.Forms.Button()
        Me.vsbReceipt = New System.Windows.Forms.VScrollBar()
        Me.picReceiptContainer = New System.Windows.Forms.PictureBox()
        Me.picReceipt = New System.Windows.Forms.PictureBox()
        Me.imgLogo = New System.Windows.Forms.PictureBox()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.lblEnterStyle = New System.Windows.Forms.Label()
        Me.cmdTax = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.fraPaymentButtons = New System.Windows.Forms.GroupBox()
        CType(Me.picReceiptContainer, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.picReceipt, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.imgLogo, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'lblCust
        '
        Me.lblCust.AutoSize = True
        Me.lblCust.Location = New System.Drawing.Point(17, 9)
        Me.lblCust.Name = "lblCust"
        Me.lblCust.Size = New System.Drawing.Size(39, 13)
        Me.lblCust.TabIndex = 0
        Me.lblCust.Text = "Label1"
        '
        'cmdComm
        '
        Me.cmdComm.Location = New System.Drawing.Point(351, 34)
        Me.cmdComm.Name = "cmdComm"
        Me.cmdComm.Size = New System.Drawing.Size(75, 23)
        Me.cmdComm.TabIndex = 1
        Me.cmdComm.Text = "Button1"
        Me.cmdComm.UseVisualStyleBackColor = True
        '
        'cboSalesList
        '
        Me.cboSalesList.FormattingEnabled = True
        Me.cboSalesList.Location = New System.Drawing.Point(350, 75)
        Me.cboSalesList.Name = "cboSalesList"
        Me.cboSalesList.Size = New System.Drawing.Size(121, 21)
        Me.cboSalesList.TabIndex = 2
        '
        'lblTotal
        '
        Me.lblTotal.AutoSize = True
        Me.lblTotal.Location = New System.Drawing.Point(354, 110)
        Me.lblTotal.Name = "lblTotal"
        Me.lblTotal.Size = New System.Drawing.Size(39, 13)
        Me.lblTotal.TabIndex = 3
        Me.lblTotal.Text = "Label1"
        '
        'lblTax
        '
        Me.lblTax.AutoSize = True
        Me.lblTax.Location = New System.Drawing.Point(360, 143)
        Me.lblTax.Name = "lblTax"
        Me.lblTax.Size = New System.Drawing.Size(39, 13)
        Me.lblTax.TabIndex = 4
        Me.lblTax.Text = "Label2"
        '
        'lblTendered
        '
        Me.lblTendered.AutoSize = True
        Me.lblTendered.Location = New System.Drawing.Point(362, 181)
        Me.lblTendered.Name = "lblTendered"
        Me.lblTendered.Size = New System.Drawing.Size(39, 13)
        Me.lblTendered.TabIndex = 5
        Me.lblTendered.Text = "Label3"
        '
        'lblDue
        '
        Me.lblDue.AutoSize = True
        Me.lblDue.Location = New System.Drawing.Point(349, 220)
        Me.lblDue.Name = "lblDue"
        Me.lblDue.Size = New System.Drawing.Size(39, 13)
        Me.lblDue.TabIndex = 6
        Me.lblDue.Text = "Label4"
        '
        'txtSku
        '
        Me.txtSku.Location = New System.Drawing.Point(351, 247)
        Me.txtSku.Name = "txtSku"
        Me.txtSku.Size = New System.Drawing.Size(93, 20)
        Me.txtSku.TabIndex = 7
        '
        'fraSaleButtons
        '
        Me.fraSaleButtons.Location = New System.Drawing.Point(351, 295)
        Me.fraSaleButtons.Name = "fraSaleButtons"
        Me.fraSaleButtons.Size = New System.Drawing.Size(200, 100)
        Me.fraSaleButtons.TabIndex = 8
        Me.fraSaleButtons.TabStop = False
        Me.fraSaleButtons.Text = "GroupBox1"
        '
        'cmdPayment
        '
        Me.cmdPayment.Location = New System.Drawing.Point(568, 35)
        Me.cmdPayment.Name = "cmdPayment"
        Me.cmdPayment.Size = New System.Drawing.Size(75, 23)
        Me.cmdPayment.TabIndex = 9
        Me.cmdPayment.Text = "Button2"
        Me.cmdPayment.UseVisualStyleBackColor = True
        '
        'cmdReturn
        '
        Me.cmdReturn.Location = New System.Drawing.Point(572, 84)
        Me.cmdReturn.Name = "cmdReturn"
        Me.cmdReturn.Size = New System.Drawing.Size(75, 23)
        Me.cmdReturn.TabIndex = 10
        Me.cmdReturn.Text = "Button3"
        Me.cmdReturn.UseVisualStyleBackColor = True
        '
        'cmdDiscount
        '
        Me.cmdDiscount.Location = New System.Drawing.Point(570, 135)
        Me.cmdDiscount.Name = "cmdDiscount"
        Me.cmdDiscount.Size = New System.Drawing.Size(75, 23)
        Me.cmdDiscount.TabIndex = 11
        Me.cmdDiscount.Text = "Button4"
        Me.cmdDiscount.UseVisualStyleBackColor = True
        '
        'cmdPrint
        '
        Me.cmdPrint.Location = New System.Drawing.Point(572, 176)
        Me.cmdPrint.Name = "cmdPrint"
        Me.cmdPrint.Size = New System.Drawing.Size(75, 29)
        Me.cmdPrint.TabIndex = 12
        Me.cmdPrint.Text = "Button5"
        Me.cmdPrint.UseVisualStyleBackColor = True
        '
        'cmdDone
        '
        Me.cmdDone.Location = New System.Drawing.Point(581, 241)
        Me.cmdDone.Name = "cmdDone"
        Me.cmdDone.Size = New System.Drawing.Size(75, 23)
        Me.cmdDone.TabIndex = 13
        Me.cmdDone.Text = "Button6"
        Me.cmdDone.UseVisualStyleBackColor = True
        '
        'cmdCancelSale
        '
        Me.cmdCancelSale.Location = New System.Drawing.Point(572, 295)
        Me.cmdCancelSale.Name = "cmdCancelSale"
        Me.cmdCancelSale.Size = New System.Drawing.Size(75, 23)
        Me.cmdCancelSale.TabIndex = 15
        Me.cmdCancelSale.Text = "Button1"
        Me.cmdCancelSale.UseVisualStyleBackColor = True
        '
        'vsbReceipt
        '
        Me.vsbReceipt.Location = New System.Drawing.Point(709, 89)
        Me.vsbReceipt.Name = "vsbReceipt"
        Me.vsbReceipt.Size = New System.Drawing.Size(17, 80)
        Me.vsbReceipt.TabIndex = 16
        '
        'picReceiptContainer
        '
        Me.picReceiptContainer.Location = New System.Drawing.Point(696, 245)
        Me.picReceiptContainer.Name = "picReceiptContainer"
        Me.picReceiptContainer.Size = New System.Drawing.Size(100, 50)
        Me.picReceiptContainer.TabIndex = 17
        Me.picReceiptContainer.TabStop = False
        '
        'picReceipt
        '
        Me.picReceipt.Location = New System.Drawing.Point(704, 318)
        Me.picReceipt.Name = "picReceipt"
        Me.picReceipt.Size = New System.Drawing.Size(59, 36)
        Me.picReceipt.TabIndex = 18
        Me.picReceipt.TabStop = False
        '
        'imgLogo
        '
        Me.imgLogo.Location = New System.Drawing.Point(691, 385)
        Me.imgLogo.Name = "imgLogo"
        Me.imgLogo.Size = New System.Drawing.Size(71, 52)
        Me.imgLogo.TabIndex = 19
        Me.imgLogo.TabStop = False
        '
        'lblEnterStyle
        '
        Me.lblEnterStyle.AutoSize = True
        Me.lblEnterStyle.Location = New System.Drawing.Point(617, 424)
        Me.lblEnterStyle.Name = "lblEnterStyle"
        Me.lblEnterStyle.Size = New System.Drawing.Size(39, 13)
        Me.lblEnterStyle.TabIndex = 20
        Me.lblEnterStyle.Text = "Label1"
        '
        'cmdTax
        '
        Me.cmdTax.Location = New System.Drawing.Point(393, 427)
        Me.cmdTax.Name = "cmdTax"
        Me.cmdTax.Size = New System.Drawing.Size(75, 23)
        Me.cmdTax.TabIndex = 21
        Me.cmdTax.Text = "Button1"
        Me.cmdTax.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(307, 403)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(39, 13)
        Me.Label1.TabIndex = 22
        Me.Label1.Text = "Label1"
        '
        'fraPaymentButtons
        '
        Me.fraPaymentButtons.Location = New System.Drawing.Point(581, 340)
        Me.fraPaymentButtons.Name = "fraPaymentButtons"
        Me.fraPaymentButtons.Size = New System.Drawing.Size(83, 54)
        Me.fraPaymentButtons.TabIndex = 23
        Me.fraPaymentButtons.TabStop = False
        Me.fraPaymentButtons.Text = "GroupBox1"
        '
        'frmCashRegister
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(800, 450)
        Me.Controls.Add(Me.fraPaymentButtons)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.cmdTax)
        Me.Controls.Add(Me.lblEnterStyle)
        Me.Controls.Add(Me.imgLogo)
        Me.Controls.Add(Me.picReceipt)
        Me.Controls.Add(Me.picReceiptContainer)
        Me.Controls.Add(Me.vsbReceipt)
        Me.Controls.Add(Me.cmdCancelSale)
        Me.Controls.Add(Me.cmdDone)
        Me.Controls.Add(Me.cmdPrint)
        Me.Controls.Add(Me.cmdDiscount)
        Me.Controls.Add(Me.cmdReturn)
        Me.Controls.Add(Me.cmdPayment)
        Me.Controls.Add(Me.fraSaleButtons)
        Me.Controls.Add(Me.txtSku)
        Me.Controls.Add(Me.lblDue)
        Me.Controls.Add(Me.lblTendered)
        Me.Controls.Add(Me.lblTax)
        Me.Controls.Add(Me.lblTotal)
        Me.Controls.Add(Me.cboSalesList)
        Me.Controls.Add(Me.cmdComm)
        Me.Controls.Add(Me.lblCust)
        Me.Name = "frmCashRegister"
        Me.Text = "frmCashRegister"
        CType(Me.picReceiptContainer, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.picReceipt, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.imgLogo, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents lblCust As Label
    Friend WithEvents cmdComm As Button
    Friend WithEvents cboSalesList As ComboBox
    Friend WithEvents lblTotal As Label
    Friend WithEvents lblTax As Label
    Friend WithEvents lblTendered As Label
    Friend WithEvents lblDue As Label
    Friend WithEvents txtSku As TextBox
    Friend WithEvents fraSaleButtons As GroupBox
    Friend WithEvents cmdPayment As Button
    Friend WithEvents cmdReturn As Button
    Friend WithEvents cmdDiscount As Button
    Friend WithEvents cmdPrint As Button
    Friend WithEvents cmdDone As Button
    Friend WithEvents cmdCancelSale As Button
    Friend WithEvents vsbReceipt As VScrollBar
    Friend WithEvents picReceiptContainer As PictureBox
    Friend WithEvents picReceipt As PictureBox
    Friend WithEvents imgLogo As PictureBox
    Friend WithEvents ToolTip1 As ToolTip
    Friend WithEvents lblEnterStyle As Label
    Friend WithEvents cmdTax As Button
    Friend WithEvents Label1 As Label
    Friend WithEvents fraPaymentButtons As GroupBox
End Class
