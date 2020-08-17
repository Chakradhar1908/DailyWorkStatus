<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class ServiceParts
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
        Me.optTagStock = New System.Windows.Forms.RadioButton()
        Me.optTagCustomer = New System.Windows.Forms.RadioButton()
        Me.fraCustomer = New System.Windows.Forms.GroupBox()
        Me.lblInvoiceNo = New System.Windows.Forms.Label()
        Me.lblSaleNo = New System.Windows.Forms.Label()
        Me.lblServiceOrderNoCaption = New System.Windows.Forms.Label()
        Me.lblServiceOrderNo = New System.Windows.Forms.Label()
        Me.lblWhatToDoWStyle = New System.Windows.Forms.Label()
        Me.txtInvoiceNo = New System.Windows.Forms.TextBox()
        Me.txtSaleNo = New System.Windows.Forms.TextBox()
        Me.txtStoreName = New System.Windows.Forms.TextBox()
        Me.dteClaimDateCaption = New System.Windows.Forms.DateTimePicker()
        Me.dteClaimDate = New System.Windows.Forms.DateTimePicker()
        Me.cboStores = New System.Windows.Forms.ComboBox()
        Me.cmdMenu = New System.Windows.Forms.Button()
        Me.txtStoreAddress = New System.Windows.Forms.TextBox()
        Me.txtStoreCity = New System.Windows.Forms.TextBox()
        Me.txtStorePhone = New System.Windows.Forms.TextBox()
        Me.cmdMoveFirst = New System.Windows.Forms.Button()
        Me.cmdMovePrevious = New System.Windows.Forms.Button()
        Me.cmdMoveNext = New System.Windows.Forms.Button()
        Me.cmdMoveLast = New System.Windows.Forms.Button()
        Me.cmdMoveSearch = New System.Windows.Forms.Button()
        Me.lblMoveRecords = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'optTagStock
        '
        Me.optTagStock.AutoSize = True
        Me.optTagStock.Location = New System.Drawing.Point(65, 30)
        Me.optTagStock.Name = "optTagStock"
        Me.optTagStock.Size = New System.Drawing.Size(90, 17)
        Me.optTagStock.TabIndex = 0
        Me.optTagStock.TabStop = True
        Me.optTagStock.Text = "RadioButton1"
        Me.optTagStock.UseVisualStyleBackColor = True
        '
        'optTagCustomer
        '
        Me.optTagCustomer.AutoSize = True
        Me.optTagCustomer.Location = New System.Drawing.Point(65, 65)
        Me.optTagCustomer.Name = "optTagCustomer"
        Me.optTagCustomer.Size = New System.Drawing.Size(90, 17)
        Me.optTagCustomer.TabIndex = 1
        Me.optTagCustomer.TabStop = True
        Me.optTagCustomer.Text = "RadioButton2"
        Me.optTagCustomer.UseVisualStyleBackColor = True
        '
        'fraCustomer
        '
        Me.fraCustomer.Location = New System.Drawing.Point(65, 109)
        Me.fraCustomer.Name = "fraCustomer"
        Me.fraCustomer.Size = New System.Drawing.Size(200, 100)
        Me.fraCustomer.TabIndex = 2
        Me.fraCustomer.TabStop = False
        Me.fraCustomer.Text = "GroupBox1"
        '
        'lblInvoiceNo
        '
        Me.lblInvoiceNo.AutoSize = True
        Me.lblInvoiceNo.Location = New System.Drawing.Point(362, 108)
        Me.lblInvoiceNo.Name = "lblInvoiceNo"
        Me.lblInvoiceNo.Size = New System.Drawing.Size(39, 13)
        Me.lblInvoiceNo.TabIndex = 3
        Me.lblInvoiceNo.Text = "Label1"
        '
        'lblSaleNo
        '
        Me.lblSaleNo.AutoSize = True
        Me.lblSaleNo.Location = New System.Drawing.Point(381, 219)
        Me.lblSaleNo.Name = "lblSaleNo"
        Me.lblSaleNo.Size = New System.Drawing.Size(39, 13)
        Me.lblSaleNo.TabIndex = 4
        Me.lblSaleNo.Text = "Label2"
        '
        'lblServiceOrderNoCaption
        '
        Me.lblServiceOrderNoCaption.AutoSize = True
        Me.lblServiceOrderNoCaption.Location = New System.Drawing.Point(381, 285)
        Me.lblServiceOrderNoCaption.Name = "lblServiceOrderNoCaption"
        Me.lblServiceOrderNoCaption.Size = New System.Drawing.Size(39, 13)
        Me.lblServiceOrderNoCaption.TabIndex = 5
        Me.lblServiceOrderNoCaption.Text = "Label3"
        '
        'lblServiceOrderNo
        '
        Me.lblServiceOrderNo.AutoSize = True
        Me.lblServiceOrderNo.Location = New System.Drawing.Point(381, 316)
        Me.lblServiceOrderNo.Name = "lblServiceOrderNo"
        Me.lblServiceOrderNo.Size = New System.Drawing.Size(39, 13)
        Me.lblServiceOrderNo.TabIndex = 6
        Me.lblServiceOrderNo.Text = "Label4"
        '
        'lblWhatToDoWStyle
        '
        Me.lblWhatToDoWStyle.AutoSize = True
        Me.lblWhatToDoWStyle.Location = New System.Drawing.Point(381, 349)
        Me.lblWhatToDoWStyle.Name = "lblWhatToDoWStyle"
        Me.lblWhatToDoWStyle.Size = New System.Drawing.Size(39, 13)
        Me.lblWhatToDoWStyle.TabIndex = 7
        Me.lblWhatToDoWStyle.Text = "Label5"
        '
        'txtInvoiceNo
        '
        Me.txtInvoiceNo.Location = New System.Drawing.Point(488, 122)
        Me.txtInvoiceNo.Name = "txtInvoiceNo"
        Me.txtInvoiceNo.Size = New System.Drawing.Size(100, 20)
        Me.txtInvoiceNo.TabIndex = 8
        '
        'txtSaleNo
        '
        Me.txtSaleNo.Location = New System.Drawing.Point(479, 166)
        Me.txtSaleNo.Name = "txtSaleNo"
        Me.txtSaleNo.Size = New System.Drawing.Size(100, 20)
        Me.txtSaleNo.TabIndex = 9
        '
        'txtStoreName
        '
        Me.txtStoreName.Location = New System.Drawing.Point(479, 202)
        Me.txtStoreName.Name = "txtStoreName"
        Me.txtStoreName.Size = New System.Drawing.Size(100, 20)
        Me.txtStoreName.TabIndex = 10
        '
        'dteClaimDateCaption
        '
        Me.dteClaimDateCaption.Location = New System.Drawing.Point(473, 248)
        Me.dteClaimDateCaption.Name = "dteClaimDateCaption"
        Me.dteClaimDateCaption.Size = New System.Drawing.Size(200, 20)
        Me.dteClaimDateCaption.TabIndex = 11
        '
        'dteClaimDate
        '
        Me.dteClaimDate.Location = New System.Drawing.Point(473, 285)
        Me.dteClaimDate.Name = "dteClaimDate"
        Me.dteClaimDate.Size = New System.Drawing.Size(200, 20)
        Me.dteClaimDate.TabIndex = 12
        '
        'cboStores
        '
        Me.cboStores.FormattingEnabled = True
        Me.cboStores.Location = New System.Drawing.Point(519, 342)
        Me.cboStores.Name = "cboStores"
        Me.cboStores.Size = New System.Drawing.Size(121, 21)
        Me.cboStores.TabIndex = 13
        '
        'cmdMenu
        '
        Me.cmdMenu.Location = New System.Drawing.Point(300, 377)
        Me.cmdMenu.Name = "cmdMenu"
        Me.cmdMenu.Size = New System.Drawing.Size(75, 23)
        Me.cmdMenu.TabIndex = 14
        Me.cmdMenu.Text = "Button1"
        Me.cmdMenu.UseVisualStyleBackColor = True
        '
        'txtStoreAddress
        '
        Me.txtStoreAddress.Location = New System.Drawing.Point(688, 122)
        Me.txtStoreAddress.Name = "txtStoreAddress"
        Me.txtStoreAddress.Size = New System.Drawing.Size(100, 20)
        Me.txtStoreAddress.TabIndex = 15
        '
        'txtStoreCity
        '
        Me.txtStoreCity.Location = New System.Drawing.Point(688, 166)
        Me.txtStoreCity.Name = "txtStoreCity"
        Me.txtStoreCity.Size = New System.Drawing.Size(100, 20)
        Me.txtStoreCity.TabIndex = 16
        '
        'txtStorePhone
        '
        Me.txtStorePhone.Location = New System.Drawing.Point(688, 192)
        Me.txtStorePhone.Name = "txtStorePhone"
        Me.txtStorePhone.Size = New System.Drawing.Size(100, 20)
        Me.txtStorePhone.TabIndex = 17
        '
        'cmdMoveFirst
        '
        Me.cmdMoveFirst.Location = New System.Drawing.Point(397, 377)
        Me.cmdMoveFirst.Name = "cmdMoveFirst"
        Me.cmdMoveFirst.Size = New System.Drawing.Size(75, 23)
        Me.cmdMoveFirst.TabIndex = 18
        Me.cmdMoveFirst.Text = "Button1"
        Me.cmdMoveFirst.UseVisualStyleBackColor = True
        '
        'cmdMovePrevious
        '
        Me.cmdMovePrevious.Location = New System.Drawing.Point(479, 377)
        Me.cmdMovePrevious.Name = "cmdMovePrevious"
        Me.cmdMovePrevious.Size = New System.Drawing.Size(75, 23)
        Me.cmdMovePrevious.TabIndex = 19
        Me.cmdMovePrevious.Text = "Button1"
        Me.cmdMovePrevious.UseVisualStyleBackColor = True
        '
        'cmdMoveNext
        '
        Me.cmdMoveNext.Location = New System.Drawing.Point(560, 377)
        Me.cmdMoveNext.Name = "cmdMoveNext"
        Me.cmdMoveNext.Size = New System.Drawing.Size(75, 23)
        Me.cmdMoveNext.TabIndex = 20
        Me.cmdMoveNext.Text = "Button1"
        Me.cmdMoveNext.UseVisualStyleBackColor = True
        '
        'cmdMoveLast
        '
        Me.cmdMoveLast.Location = New System.Drawing.Point(641, 377)
        Me.cmdMoveLast.Name = "cmdMoveLast"
        Me.cmdMoveLast.Size = New System.Drawing.Size(75, 23)
        Me.cmdMoveLast.TabIndex = 21
        Me.cmdMoveLast.Text = "Button1"
        Me.cmdMoveLast.UseVisualStyleBackColor = True
        '
        'cmdMoveSearch
        '
        Me.cmdMoveSearch.Location = New System.Drawing.Point(722, 377)
        Me.cmdMoveSearch.Name = "cmdMoveSearch"
        Me.cmdMoveSearch.Size = New System.Drawing.Size(75, 23)
        Me.cmdMoveSearch.TabIndex = 22
        Me.cmdMoveSearch.Text = "Button1"
        Me.cmdMoveSearch.UseVisualStyleBackColor = True
        '
        'lblMoveRecords
        '
        Me.lblMoveRecords.AutoSize = True
        Me.lblMoveRecords.Location = New System.Drawing.Point(395, 410)
        Me.lblMoveRecords.Name = "lblMoveRecords"
        Me.lblMoveRecords.Size = New System.Drawing.Size(39, 13)
        Me.lblMoveRecords.TabIndex = 23
        Me.lblMoveRecords.Text = "Label1"
        '
        'ServiceParts
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(800, 450)
        Me.Controls.Add(Me.lblMoveRecords)
        Me.Controls.Add(Me.cmdMoveSearch)
        Me.Controls.Add(Me.cmdMoveLast)
        Me.Controls.Add(Me.cmdMoveNext)
        Me.Controls.Add(Me.cmdMovePrevious)
        Me.Controls.Add(Me.cmdMoveFirst)
        Me.Controls.Add(Me.txtStorePhone)
        Me.Controls.Add(Me.txtStoreCity)
        Me.Controls.Add(Me.txtStoreAddress)
        Me.Controls.Add(Me.cmdMenu)
        Me.Controls.Add(Me.cboStores)
        Me.Controls.Add(Me.dteClaimDate)
        Me.Controls.Add(Me.dteClaimDateCaption)
        Me.Controls.Add(Me.txtStoreName)
        Me.Controls.Add(Me.txtSaleNo)
        Me.Controls.Add(Me.txtInvoiceNo)
        Me.Controls.Add(Me.lblWhatToDoWStyle)
        Me.Controls.Add(Me.lblServiceOrderNo)
        Me.Controls.Add(Me.lblServiceOrderNoCaption)
        Me.Controls.Add(Me.lblSaleNo)
        Me.Controls.Add(Me.lblInvoiceNo)
        Me.Controls.Add(Me.fraCustomer)
        Me.Controls.Add(Me.optTagCustomer)
        Me.Controls.Add(Me.optTagStock)
        Me.Name = "ServiceParts"
        Me.Text = "ServiceParts"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents optTagStock As RadioButton
    Friend WithEvents optTagCustomer As RadioButton
    Friend WithEvents fraCustomer As GroupBox
    Friend WithEvents lblInvoiceNo As Label
    Friend WithEvents lblSaleNo As Label
    Friend WithEvents lblServiceOrderNoCaption As Label
    Friend WithEvents lblServiceOrderNo As Label
    Friend WithEvents lblWhatToDoWStyle As Label
    Friend WithEvents txtInvoiceNo As TextBox
    Friend WithEvents txtSaleNo As TextBox
    Friend WithEvents txtStoreName As TextBox
    Friend WithEvents dteClaimDateCaption As DateTimePicker
    Friend WithEvents dteClaimDate As DateTimePicker
    Friend WithEvents cboStores As ComboBox
    Friend WithEvents cmdMenu As Button
    Friend WithEvents txtStoreAddress As TextBox
    Friend WithEvents txtStoreCity As TextBox
    Friend WithEvents txtStorePhone As TextBox
    Friend WithEvents cmdMoveFirst As Button
    Friend WithEvents cmdMovePrevious As Button
    Friend WithEvents cmdMoveNext As Button
    Friend WithEvents cmdMoveLast As Button
    Friend WithEvents cmdMoveSearch As Button
    Friend WithEvents lblMoveRecords As Label
End Class
