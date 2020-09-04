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
        Me.fraMain = New System.Windows.Forms.GroupBox()
        Me.tmrLockOn = New System.Windows.Forms.Timer(Me.components)
        Me.cmdCancel = New System.Windows.Forms.Button()
        Me.cmdOk = New System.Windows.Forms.Button()
        Me.chkEmail = New System.Windows.Forms.CheckBox()
        Me.fraControl = New System.Windows.Forms.GroupBox()
        Me.chkReceipt = New System.Windows.Forms.CheckBox()
        Me.Memo = New System.Windows.Forms.TextBox()
        Me.cboAccount = New System.Windows.Forms.ComboBox()
        Me.chkPayAll = New System.Windows.Forms.CheckBox()
        Me.txtAmount = New System.Windows.Forms.TextBox()
        Me.lblMemo = New System.Windows.Forms.Label()
        Me.lblAccount = New System.Windows.Forms.Label()
        Me.lblAmount = New System.Windows.Forms.Label()
        Me.txtNoPay = New System.Windows.Forms.TextBox()
        Me.cmdChangeDate = New System.Windows.Forms.Button()
        Me.dtePayDate = New System.Windows.Forms.DateTimePicker()
        Me.lblPayDate = New System.Windows.Forms.Label()
        Me.lblSaleNo = New System.Windows.Forms.Label()
        Me.lblCity = New System.Windows.Forms.Label()
        Me.lblAddress = New System.Windows.Forms.Label()
        Me.lblName = New System.Windows.Forms.Label()
        Me.lblCityCaption = New System.Windows.Forms.Label()
        Me.lblAddressCaption = New System.Windows.Forms.Label()
        Me.lblNameCaption = New System.Windows.Forms.Label()
        Me.lblSaleTitle = New System.Windows.Forms.Label()
        Me.fraMain.SuspendLayout()
        Me.fraControl.SuspendLayout()
        Me.SuspendLayout()
        '
        'fraMain
        '
        Me.fraMain.Controls.Add(Me.fraControl)
        Me.fraMain.Controls.Add(Me.Memo)
        Me.fraMain.Controls.Add(Me.cboAccount)
        Me.fraMain.Controls.Add(Me.chkPayAll)
        Me.fraMain.Controls.Add(Me.txtAmount)
        Me.fraMain.Controls.Add(Me.lblMemo)
        Me.fraMain.Controls.Add(Me.lblAccount)
        Me.fraMain.Controls.Add(Me.lblAmount)
        Me.fraMain.Controls.Add(Me.cmdChangeDate)
        Me.fraMain.Controls.Add(Me.dtePayDate)
        Me.fraMain.Controls.Add(Me.lblPayDate)
        Me.fraMain.Controls.Add(Me.lblSaleNo)
        Me.fraMain.Controls.Add(Me.lblCity)
        Me.fraMain.Controls.Add(Me.lblAddress)
        Me.fraMain.Controls.Add(Me.lblName)
        Me.fraMain.Controls.Add(Me.lblCityCaption)
        Me.fraMain.Controls.Add(Me.lblAddressCaption)
        Me.fraMain.Controls.Add(Me.lblNameCaption)
        Me.fraMain.Controls.Add(Me.lblSaleTitle)
        Me.fraMain.Controls.Add(Me.txtNoPay)
        Me.fraMain.Location = New System.Drawing.Point(4, 2)
        Me.fraMain.Name = "fraMain"
        Me.fraMain.Size = New System.Drawing.Size(519, 266)
        Me.fraMain.TabIndex = 0
        Me.fraMain.TabStop = False
        '
        'cmdCancel
        '
        Me.cmdCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cmdCancel.Location = New System.Drawing.Point(132, 17)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.Size = New System.Drawing.Size(53, 60)
        Me.cmdCancel.TabIndex = 3
        Me.cmdCancel.Text = "&Cancel"
        Me.cmdCancel.UseVisualStyleBackColor = True
        '
        'cmdOk
        '
        Me.cmdOk.Location = New System.Drawing.Point(75, 17)
        Me.cmdOk.Name = "cmdOk"
        Me.cmdOk.Size = New System.Drawing.Size(53, 60)
        Me.cmdOk.TabIndex = 2
        Me.cmdOk.Text = "&OK"
        Me.cmdOk.UseVisualStyleBackColor = True
        '
        'chkEmail
        '
        Me.chkEmail.AutoSize = True
        Me.chkEmail.CheckAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.chkEmail.Location = New System.Drawing.Point(18, 34)
        Me.chkEmail.Name = "chkEmail"
        Me.chkEmail.Size = New System.Drawing.Size(51, 17)
        Me.chkEmail.TabIndex = 1
        Me.chkEmail.Text = "E&mail"
        Me.chkEmail.UseVisualStyleBackColor = True
        '
        'fraControl
        '
        Me.fraControl.Controls.Add(Me.cmdCancel)
        Me.fraControl.Controls.Add(Me.cmdOk)
        Me.fraControl.Controls.Add(Me.chkEmail)
        Me.fraControl.Controls.Add(Me.chkReceipt)
        Me.fraControl.Location = New System.Drawing.Point(288, 156)
        Me.fraControl.Name = "fraControl"
        Me.fraControl.Size = New System.Drawing.Size(224, 83)
        Me.fraControl.TabIndex = 39
        Me.fraControl.TabStop = False
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
        'Memo
        '
        Me.Memo.Location = New System.Drawing.Point(364, 118)
        Me.Memo.Name = "Memo"
        Me.Memo.Size = New System.Drawing.Size(100, 20)
        Me.Memo.TabIndex = 38
        '
        'cboAccount
        '
        Me.cboAccount.FormattingEnabled = True
        Me.cboAccount.Location = New System.Drawing.Point(364, 91)
        Me.cboAccount.Name = "cboAccount"
        Me.cboAccount.Size = New System.Drawing.Size(121, 21)
        Me.cboAccount.TabIndex = 37
        '
        'chkPayAll
        '
        Me.chkPayAll.AutoSize = True
        Me.chkPayAll.Location = New System.Drawing.Point(448, 65)
        Me.chkPayAll.Name = "chkPayAll"
        Me.chkPayAll.Size = New System.Drawing.Size(58, 17)
        Me.chkPayAll.TabIndex = 36
        Me.chkPayAll.Text = "Pay &All"
        Me.chkPayAll.UseVisualStyleBackColor = True
        '
        'txtAmount
        '
        Me.txtAmount.Location = New System.Drawing.Point(364, 65)
        Me.txtAmount.Name = "txtAmount"
        Me.txtAmount.Size = New System.Drawing.Size(78, 20)
        Me.txtAmount.TabIndex = 35
        '
        'lblMemo
        '
        Me.lblMemo.AutoSize = True
        Me.lblMemo.Location = New System.Drawing.Point(296, 132)
        Me.lblMemo.Name = "lblMemo"
        Me.lblMemo.Size = New System.Drawing.Size(39, 13)
        Me.lblMemo.TabIndex = 34
        Me.lblMemo.Text = "Memo:"
        '
        'lblAccount
        '
        Me.lblAccount.AutoSize = True
        Me.lblAccount.Location = New System.Drawing.Point(290, 105)
        Me.lblAccount.Name = "lblAccount"
        Me.lblAccount.Size = New System.Drawing.Size(45, 13)
        Me.lblAccount.TabIndex = 33
        Me.lblAccount.Text = "How Pd"
        '
        'lblAmount
        '
        Me.lblAmount.AutoSize = True
        Me.lblAmount.Location = New System.Drawing.Point(289, 78)
        Me.lblAmount.Name = "lblAmount"
        Me.lblAmount.Size = New System.Drawing.Size(46, 13)
        Me.lblAmount.TabIndex = 32
        Me.lblAmount.Text = "Amount:"
        '
        'txtNoPay
        '
        Me.txtNoPay.BackColor = System.Drawing.SystemColors.ButtonFace
        Me.txtNoPay.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtNoPay.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtNoPay.Location = New System.Drawing.Point(287, 59)
        Me.txtNoPay.Multiline = True
        Me.txtNoPay.Name = "txtNoPay"
        Me.txtNoPay.Size = New System.Drawing.Size(225, 91)
        Me.txtNoPay.TabIndex = 31
        Me.txtNoPay.Text = "Sale is financed." & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "No payment required." & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "Dbl-Click here to Add " & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "Additional Pay" &
    "ment."
        Me.txtNoPay.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        Me.txtNoPay.Visible = False
        '
        'cmdChangeDate
        '
        Me.cmdChangeDate.Location = New System.Drawing.Point(448, 23)
        Me.cmdChangeDate.Name = "cmdChangeDate"
        Me.cmdChangeDate.Size = New System.Drawing.Size(64, 23)
        Me.cmdChangeDate.TabIndex = 30
        Me.cmdChangeDate.Text = "Chan&ge"
        Me.cmdChangeDate.UseVisualStyleBackColor = True
        '
        'dtePayDate
        '
        Me.dtePayDate.CustomFormat = "MM/dd/yyyy"
        Me.dtePayDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtePayDate.Location = New System.Drawing.Point(364, 26)
        Me.dtePayDate.Name = "dtePayDate"
        Me.dtePayDate.Size = New System.Drawing.Size(78, 20)
        Me.dtePayDate.TabIndex = 29
        '
        'lblPayDate
        '
        Me.lblPayDate.AutoSize = True
        Me.lblPayDate.Location = New System.Drawing.Point(324, 30)
        Me.lblPayDate.Name = "lblPayDate"
        Me.lblPayDate.Size = New System.Drawing.Size(33, 13)
        Me.lblPayDate.TabIndex = 28
        Me.lblPayDate.Text = "Date:"
        '
        'lblSaleNo
        '
        Me.lblSaleNo.Location = New System.Drawing.Point(8, 36)
        Me.lblSaleNo.Name = "lblSaleNo"
        Me.lblSaleNo.Size = New System.Drawing.Size(216, 41)
        Me.lblSaleNo.TabIndex = 27
        Me.lblSaleNo.Text = "lblSaleNo"
        '
        'lblCity
        '
        Me.lblCity.Location = New System.Drawing.Point(8, 226)
        Me.lblCity.Name = "lblCity"
        Me.lblCity.Size = New System.Drawing.Size(216, 26)
        Me.lblCity.TabIndex = 26
        Me.lblCity.Text = "lblCity"
        '
        'lblAddress
        '
        Me.lblAddress.Location = New System.Drawing.Point(8, 166)
        Me.lblAddress.Name = "lblAddress"
        Me.lblAddress.Size = New System.Drawing.Size(216, 41)
        Me.lblAddress.TabIndex = 25
        Me.lblAddress.Text = "lblAddress"
        '
        'lblName
        '
        Me.lblName.Location = New System.Drawing.Point(8, 102)
        Me.lblName.Name = "lblName"
        Me.lblName.Size = New System.Drawing.Size(216, 41)
        Me.lblName.TabIndex = 24
        Me.lblName.Text = "lblName"
        '
        'lblCityCaption
        '
        Me.lblCityCaption.AutoSize = True
        Me.lblCityCaption.Location = New System.Drawing.Point(8, 206)
        Me.lblCityCaption.Name = "lblCityCaption"
        Me.lblCityCaption.Size = New System.Drawing.Size(79, 13)
        Me.lblCityCaption.TabIndex = 23
        Me.lblCityCaption.Text = "City, State  Zip:"
        '
        'lblAddressCaption
        '
        Me.lblAddressCaption.AutoSize = True
        Me.lblAddressCaption.Location = New System.Drawing.Point(8, 146)
        Me.lblAddressCaption.Name = "lblAddressCaption"
        Me.lblAddressCaption.Size = New System.Drawing.Size(48, 13)
        Me.lblAddressCaption.TabIndex = 22
        Me.lblAddressCaption.Text = "Address:"
        '
        'lblNameCaption
        '
        Me.lblNameCaption.AutoSize = True
        Me.lblNameCaption.Location = New System.Drawing.Point(8, 87)
        Me.lblNameCaption.Name = "lblNameCaption"
        Me.lblNameCaption.Size = New System.Drawing.Size(38, 13)
        Me.lblNameCaption.TabIndex = 21
        Me.lblNameCaption.Text = "Name:"
        '
        'lblSaleTitle
        '
        Me.lblSaleTitle.AutoSize = True
        Me.lblSaleTitle.Location = New System.Drawing.Point(8, 16)
        Me.lblSaleTitle.Name = "lblSaleTitle"
        Me.lblSaleTitle.Size = New System.Drawing.Size(61, 13)
        Me.lblSaleTitle.TabIndex = 20
        Me.lblSaleTitle.Text = "Bill Of Sale:"
        '
        'OrdPay
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(543, 275)
        Me.ControlBox = False
        Me.Controls.Add(Me.fraMain)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Name = "OrdPay"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Payment "
        Me.fraMain.ResumeLayout(False)
        Me.fraMain.PerformLayout()
        Me.fraControl.ResumeLayout(False)
        Me.fraControl.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents fraMain As GroupBox
    Friend WithEvents fraControl As GroupBox
    Friend WithEvents cmdCancel As Button
    Friend WithEvents cmdOk As Button
    Friend WithEvents chkEmail As CheckBox
    Friend WithEvents chkReceipt As CheckBox
    Friend WithEvents Memo As TextBox
    Friend WithEvents cboAccount As ComboBox
    Friend WithEvents chkPayAll As CheckBox
    Friend WithEvents txtAmount As TextBox
    Friend WithEvents lblMemo As Label
    Friend WithEvents lblAccount As Label
    Friend WithEvents lblAmount As Label
    Friend WithEvents cmdChangeDate As Button
    Friend WithEvents dtePayDate As DateTimePicker
    Friend WithEvents lblPayDate As Label
    Friend WithEvents lblSaleNo As Label
    Friend WithEvents lblCity As Label
    Friend WithEvents lblAddress As Label
    Friend WithEvents lblName As Label
    Friend WithEvents lblCityCaption As Label
    Friend WithEvents lblAddressCaption As Label
    Friend WithEvents lblNameCaption As Label
    Friend WithEvents lblSaleTitle As Label
    Friend WithEvents txtNoPay As TextBox
    Friend WithEvents tmrLockOn As Timer
End Class
