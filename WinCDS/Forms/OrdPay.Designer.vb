<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class OrdPay
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
        Me.lblSaleTitle = New System.Windows.Forms.Label()
        Me.lblNameCaption = New System.Windows.Forms.Label()
        Me.lblAddressCaption = New System.Windows.Forms.Label()
        Me.lblCityCaption = New System.Windows.Forms.Label()
        Me.lblName = New System.Windows.Forms.Label()
        Me.lblAddress = New System.Windows.Forms.Label()
        Me.lblCity = New System.Windows.Forms.Label()
        Me.lblSaleNo = New System.Windows.Forms.Label()
        Me.lblPayDate = New System.Windows.Forms.Label()
        Me.dtePayDate = New System.Windows.Forms.DateTimePicker()
        Me.cmdChangeDate = New System.Windows.Forms.Button()
        Me.txtNoPay = New System.Windows.Forms.TextBox()
        Me.lblAmount = New System.Windows.Forms.Label()
        Me.lblAccount = New System.Windows.Forms.Label()
        Me.lblMemo = New System.Windows.Forms.Label()
        Me.txtAmount = New System.Windows.Forms.TextBox()
        Me.chkPayAll = New System.Windows.Forms.CheckBox()
        Me.cboAccount = New System.Windows.Forms.ComboBox()
        Me.Memo = New System.Windows.Forms.TextBox()
        Me.fraControl = New System.Windows.Forms.GroupBox()
        Me.cmdCancel = New System.Windows.Forms.Button()
        Me.cmdOk = New System.Windows.Forms.Button()
        Me.chkEmail = New System.Windows.Forms.CheckBox()
        Me.chkReceipt = New System.Windows.Forms.CheckBox()
        Me.tmrLockOn = New System.Windows.Forms.Timer(Me.components)
        Me.fraControl.SuspendLayout()
        Me.SuspendLayout()
        '
        'lblSaleTitle
        '
        Me.lblSaleTitle.AutoSize = True
        Me.lblSaleTitle.Location = New System.Drawing.Point(8, 9)
        Me.lblSaleTitle.Name = "lblSaleTitle"
        Me.lblSaleTitle.Size = New System.Drawing.Size(61, 13)
        Me.lblSaleTitle.TabIndex = 0
        Me.lblSaleTitle.Text = "Bill Of Sale:"
        '
        'lblNameCaption
        '
        Me.lblNameCaption.AutoSize = True
        Me.lblNameCaption.Location = New System.Drawing.Point(8, 77)
        Me.lblNameCaption.Name = "lblNameCaption"
        Me.lblNameCaption.Size = New System.Drawing.Size(38, 13)
        Me.lblNameCaption.TabIndex = 1
        Me.lblNameCaption.Text = "Name:"
        '
        'lblAddressCaption
        '
        Me.lblAddressCaption.AutoSize = True
        Me.lblAddressCaption.Location = New System.Drawing.Point(8, 142)
        Me.lblAddressCaption.Name = "lblAddressCaption"
        Me.lblAddressCaption.Size = New System.Drawing.Size(48, 13)
        Me.lblAddressCaption.TabIndex = 2
        Me.lblAddressCaption.Text = "Address:"
        '
        'lblCityCaption
        '
        Me.lblCityCaption.AutoSize = True
        Me.lblCityCaption.Location = New System.Drawing.Point(8, 203)
        Me.lblCityCaption.Name = "lblCityCaption"
        Me.lblCityCaption.Size = New System.Drawing.Size(79, 13)
        Me.lblCityCaption.TabIndex = 3
        Me.lblCityCaption.Text = "City, State  Zip:"
        '
        'lblName
        '
        Me.lblName.Location = New System.Drawing.Point(8, 92)
        Me.lblName.Name = "lblName"
        Me.lblName.Size = New System.Drawing.Size(216, 41)
        Me.lblName.TabIndex = 4
        Me.lblName.Text = "lblName"
        '
        'lblAddress
        '
        Me.lblAddress.Location = New System.Drawing.Point(8, 162)
        Me.lblAddress.Name = "lblAddress"
        Me.lblAddress.Size = New System.Drawing.Size(216, 41)
        Me.lblAddress.TabIndex = 5
        Me.lblAddress.Text = "lblAddress"
        '
        'lblCity
        '
        Me.lblCity.Location = New System.Drawing.Point(8, 223)
        Me.lblCity.Name = "lblCity"
        Me.lblCity.Size = New System.Drawing.Size(216, 41)
        Me.lblCity.TabIndex = 6
        Me.lblCity.Text = "lblCity"
        '
        'lblSaleNo
        '
        Me.lblSaleNo.Location = New System.Drawing.Point(8, 29)
        Me.lblSaleNo.Name = "lblSaleNo"
        Me.lblSaleNo.Size = New System.Drawing.Size(216, 41)
        Me.lblSaleNo.TabIndex = 7
        Me.lblSaleNo.Text = "lblSaleNo"
        '
        'lblPayDate
        '
        Me.lblPayDate.AutoSize = True
        Me.lblPayDate.Location = New System.Drawing.Point(281, 12)
        Me.lblPayDate.Name = "lblPayDate"
        Me.lblPayDate.Size = New System.Drawing.Size(33, 13)
        Me.lblPayDate.TabIndex = 8
        Me.lblPayDate.Text = "Date:"
        '
        'dtePayDate
        '
        Me.dtePayDate.CustomFormat = "MM/dd/yyyy"
        Me.dtePayDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtePayDate.Location = New System.Drawing.Point(321, 8)
        Me.dtePayDate.Name = "dtePayDate"
        Me.dtePayDate.Size = New System.Drawing.Size(78, 20)
        Me.dtePayDate.TabIndex = 9
        '
        'cmdChangeDate
        '
        Me.cmdChangeDate.Location = New System.Drawing.Point(405, 5)
        Me.cmdChangeDate.Name = "cmdChangeDate"
        Me.cmdChangeDate.Size = New System.Drawing.Size(64, 23)
        Me.cmdChangeDate.TabIndex = 10
        Me.cmdChangeDate.Text = "Chan&ge"
        Me.cmdChangeDate.UseVisualStyleBackColor = True
        '
        'txtNoPay
        '
        Me.txtNoPay.BackColor = System.Drawing.SystemColors.ButtonFace
        Me.txtNoPay.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtNoPay.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtNoPay.Location = New System.Drawing.Point(250, 36)
        Me.txtNoPay.Multiline = True
        Me.txtNoPay.Name = "txtNoPay"
        Me.txtNoPay.Size = New System.Drawing.Size(216, 91)
        Me.txtNoPay.TabIndex = 11
        Me.txtNoPay.Text = "Sale is financed." & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "No payment required." & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "Dbl-Click here to Add " & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "Additional Pay" &
    "ment."
        Me.txtNoPay.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'lblAmount
        '
        Me.lblAmount.AutoSize = True
        Me.lblAmount.Location = New System.Drawing.Point(268, 39)
        Me.lblAmount.Name = "lblAmount"
        Me.lblAmount.Size = New System.Drawing.Size(46, 13)
        Me.lblAmount.TabIndex = 12
        Me.lblAmount.Text = "Amount:"
        '
        'lblAccount
        '
        Me.lblAccount.AutoSize = True
        Me.lblAccount.Location = New System.Drawing.Point(269, 66)
        Me.lblAccount.Name = "lblAccount"
        Me.lblAccount.Size = New System.Drawing.Size(45, 13)
        Me.lblAccount.TabIndex = 13
        Me.lblAccount.Text = "How Pd"
        '
        'lblMemo
        '
        Me.lblMemo.AutoSize = True
        Me.lblMemo.Location = New System.Drawing.Point(275, 93)
        Me.lblMemo.Name = "lblMemo"
        Me.lblMemo.Size = New System.Drawing.Size(39, 13)
        Me.lblMemo.TabIndex = 14
        Me.lblMemo.Text = "Memo:"
        '
        'txtAmount
        '
        Me.txtAmount.Location = New System.Drawing.Point(321, 34)
        Me.txtAmount.Name = "txtAmount"
        Me.txtAmount.Size = New System.Drawing.Size(78, 20)
        Me.txtAmount.TabIndex = 15
        '
        'chkPayAll
        '
        Me.chkPayAll.AutoSize = True
        Me.chkPayAll.Location = New System.Drawing.Point(405, 34)
        Me.chkPayAll.Name = "chkPayAll"
        Me.chkPayAll.Size = New System.Drawing.Size(64, 17)
        Me.chkPayAll.TabIndex = 16
        Me.chkPayAll.Text = "Pay   &All"
        Me.chkPayAll.UseVisualStyleBackColor = True
        '
        'cboAccount
        '
        Me.cboAccount.FormattingEnabled = True
        Me.cboAccount.Location = New System.Drawing.Point(321, 60)
        Me.cboAccount.Name = "cboAccount"
        Me.cboAccount.Size = New System.Drawing.Size(121, 21)
        Me.cboAccount.TabIndex = 17
        '
        'Memo
        '
        Me.Memo.Location = New System.Drawing.Point(321, 87)
        Me.Memo.Name = "Memo"
        Me.Memo.Size = New System.Drawing.Size(100, 20)
        Me.Memo.TabIndex = 18
        '
        'fraControl
        '
        Me.fraControl.Controls.Add(Me.cmdCancel)
        Me.fraControl.Controls.Add(Me.cmdOk)
        Me.fraControl.Controls.Add(Me.chkEmail)
        Me.fraControl.Controls.Add(Me.chkReceipt)
        Me.fraControl.Location = New System.Drawing.Point(271, 133)
        Me.fraControl.Name = "fraControl"
        Me.fraControl.Size = New System.Drawing.Size(200, 61)
        Me.fraControl.TabIndex = 19
        Me.fraControl.TabStop = False
        '
        'cmdCancel
        '
        Me.cmdCancel.Location = New System.Drawing.Point(132, 17)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.Size = New System.Drawing.Size(51, 39)
        Me.cmdCancel.TabIndex = 3
        Me.cmdCancel.Text = "&Cancel"
        Me.cmdCancel.UseVisualStyleBackColor = True
        '
        'cmdOk
        '
        Me.cmdOk.Location = New System.Drawing.Point(75, 17)
        Me.cmdOk.Name = "cmdOk"
        Me.cmdOk.Size = New System.Drawing.Size(51, 39)
        Me.cmdOk.TabIndex = 2
        Me.cmdOk.Text = "&OK"
        Me.cmdOk.UseVisualStyleBackColor = True
        '
        'chkEmail
        '
        Me.chkEmail.AutoSize = True
        Me.chkEmail.CheckAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.chkEmail.Location = New System.Drawing.Point(18, 39)
        Me.chkEmail.Name = "chkEmail"
        Me.chkEmail.Size = New System.Drawing.Size(51, 17)
        Me.chkEmail.TabIndex = 1
        Me.chkEmail.Text = "E&mail"
        Me.chkEmail.UseVisualStyleBackColor = True
        '
        'chkReceipt
        '
        Me.chkReceipt.AutoSize = True
        Me.chkReceipt.CheckAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.chkReceipt.Location = New System.Drawing.Point(6, 16)
        Me.chkReceipt.Name = "chkReceipt"
        Me.chkReceipt.Size = New System.Drawing.Size(63, 17)
        Me.chkReceipt.TabIndex = 0
        Me.chkReceipt.Text = "&Receipt"
        Me.chkReceipt.UseVisualStyleBackColor = True
        '
        'OrdPay
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(477, 263)
        Me.ControlBox = False
        Me.Controls.Add(Me.fraControl)
        Me.Controls.Add(Me.Memo)
        Me.Controls.Add(Me.cboAccount)
        Me.Controls.Add(Me.chkPayAll)
        Me.Controls.Add(Me.txtAmount)
        Me.Controls.Add(Me.lblMemo)
        Me.Controls.Add(Me.lblAccount)
        Me.Controls.Add(Me.lblAmount)
        Me.Controls.Add(Me.txtNoPay)
        Me.Controls.Add(Me.cmdChangeDate)
        Me.Controls.Add(Me.dtePayDate)
        Me.Controls.Add(Me.lblPayDate)
        Me.Controls.Add(Me.lblSaleNo)
        Me.Controls.Add(Me.lblCity)
        Me.Controls.Add(Me.lblAddress)
        Me.Controls.Add(Me.lblName)
        Me.Controls.Add(Me.lblCityCaption)
        Me.Controls.Add(Me.lblAddressCaption)
        Me.Controls.Add(Me.lblNameCaption)
        Me.Controls.Add(Me.lblSaleTitle)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Name = "OrdPay"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Payment "
        Me.fraControl.ResumeLayout(False)
        Me.fraControl.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents lblSaleTitle As Label
    Friend WithEvents lblNameCaption As Label
    Friend WithEvents lblAddressCaption As Label
    Friend WithEvents lblCityCaption As Label
    Friend WithEvents lblName As Label
    Friend WithEvents lblAddress As Label
    Friend WithEvents lblCity As Label
    Friend WithEvents lblSaleNo As Label
    Friend WithEvents lblPayDate As Label
    Friend WithEvents dtePayDate As DateTimePicker
    Friend WithEvents cmdChangeDate As Button
    Friend WithEvents txtNoPay As TextBox
    Friend WithEvents lblAmount As Label
    Friend WithEvents lblAccount As Label
    Friend WithEvents lblMemo As Label
    Friend WithEvents txtAmount As TextBox
    Friend WithEvents chkPayAll As CheckBox
    Friend WithEvents cboAccount As ComboBox
    Friend WithEvents Memo As TextBox
    Friend WithEvents fraControl As GroupBox
    Friend WithEvents cmdCancel As Button
    Friend WithEvents cmdOk As Button
    Friend WithEvents chkEmail As CheckBox
    Friend WithEvents chkReceipt As CheckBox
    Friend WithEvents tmrLockOn As Timer
End Class
