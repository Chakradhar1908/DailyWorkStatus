<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmCashRegisterAddress
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
        Me.txtEmail = New System.Windows.Forms.TextBox()
        Me.txtPhone2 = New System.Windows.Forms.TextBox()
        Me.txtPhone1 = New System.Windows.Forms.TextBox()
        Me.txtZip = New System.Windows.Forms.TextBox()
        Me.txtCityST = New System.Windows.Forms.TextBox()
        Me.txtAdd2 = New System.Windows.Forms.TextBox()
        Me.txtAdd1 = New System.Windows.Forms.TextBox()
        Me.txtLastName = New System.Windows.Forms.TextBox()
        Me.txtFirstName = New System.Windows.Forms.TextBox()
        Me.chkBusiness = New System.Windows.Forms.CheckBox()
        Me.lblEmail = New System.Windows.Forms.Label()
        Me.cmdShipTo = New System.Windows.Forms.Button()
        Me.lblFirstName = New System.Windows.Forms.Label()
        Me.fraCI = New System.Windows.Forms.GroupBox()
        Me.lblPhone = New System.Windows.Forms.Label()
        Me.lblZip = New System.Windows.Forms.Label()
        Me.lblCityST = New System.Windows.Forms.Label()
        Me.lblAdd1 = New System.Windows.Forms.Label()
        Me.cmdOK = New System.Windows.Forms.Button()
        Me.cmdCancel = New System.Windows.Forms.Button()
        Me.fraCI.SuspendLayout()
        Me.SuspendLayout()
        '
        'txtEmail
        '
        Me.txtEmail.Location = New System.Drawing.Point(58, 135)
        Me.txtEmail.Name = "txtEmail"
        Me.txtEmail.Size = New System.Drawing.Size(288, 20)
        Me.txtEmail.TabIndex = 22
        '
        'txtPhone2
        '
        Me.txtPhone2.Location = New System.Drawing.Point(217, 111)
        Me.txtPhone2.Name = "txtPhone2"
        Me.txtPhone2.Size = New System.Drawing.Size(129, 20)
        Me.txtPhone2.TabIndex = 21
        '
        'txtPhone1
        '
        Me.txtPhone1.Location = New System.Drawing.Point(58, 111)
        Me.txtPhone1.Name = "txtPhone1"
        Me.txtPhone1.Size = New System.Drawing.Size(141, 20)
        Me.txtPhone1.TabIndex = 20
        '
        'txtZip
        '
        Me.txtZip.Location = New System.Drawing.Point(259, 88)
        Me.txtZip.Name = "txtZip"
        Me.txtZip.Size = New System.Drawing.Size(87, 20)
        Me.txtZip.TabIndex = 19
        '
        'txtCityST
        '
        Me.txtCityST.Location = New System.Drawing.Point(58, 88)
        Me.txtCityST.Name = "txtCityST"
        Me.txtCityST.Size = New System.Drawing.Size(170, 20)
        Me.txtCityST.TabIndex = 18
        '
        'txtAdd2
        '
        Me.txtAdd2.Location = New System.Drawing.Point(58, 65)
        Me.txtAdd2.Name = "txtAdd2"
        Me.txtAdd2.Size = New System.Drawing.Size(288, 20)
        Me.txtAdd2.TabIndex = 17
        '
        'txtAdd1
        '
        Me.txtAdd1.Location = New System.Drawing.Point(58, 42)
        Me.txtAdd1.Name = "txtAdd1"
        Me.txtAdd1.Size = New System.Drawing.Size(288, 20)
        Me.txtAdd1.TabIndex = 16
        '
        'txtLastName
        '
        Me.txtLastName.Location = New System.Drawing.Point(157, 19)
        Me.txtLastName.Name = "txtLastName"
        Me.txtLastName.Size = New System.Drawing.Size(100, 20)
        Me.txtLastName.TabIndex = 15
        '
        'txtFirstName
        '
        Me.txtFirstName.Location = New System.Drawing.Point(48, 19)
        Me.txtFirstName.Name = "txtFirstName"
        Me.txtFirstName.Size = New System.Drawing.Size(100, 20)
        Me.txtFirstName.TabIndex = 14
        '
        'chkBusiness
        '
        Me.chkBusiness.AutoSize = True
        Me.chkBusiness.Location = New System.Drawing.Point(272, 21)
        Me.chkBusiness.Name = "chkBusiness"
        Me.chkBusiness.Size = New System.Drawing.Size(74, 17)
        Me.chkBusiness.TabIndex = 12
        Me.chkBusiness.Text = "&Business?"
        Me.chkBusiness.UseVisualStyleBackColor = True
        '
        'lblEmail
        '
        Me.lblEmail.AutoSize = True
        Me.lblEmail.Location = New System.Drawing.Point(9, 142)
        Me.lblEmail.Name = "lblEmail"
        Me.lblEmail.Size = New System.Drawing.Size(35, 13)
        Me.lblEmail.TabIndex = 23
        Me.lblEmail.Text = "&Email:"
        '
        'cmdShipTo
        '
        Me.cmdShipTo.Location = New System.Drawing.Point(124, 161)
        Me.cmdShipTo.Name = "cmdShipTo"
        Me.cmdShipTo.Size = New System.Drawing.Size(133, 23)
        Me.cmdShipTo.TabIndex = 24
        Me.cmdShipTo.Text = "&Shipping Address >>"
        Me.cmdShipTo.UseVisualStyleBackColor = True
        '
        'lblFirstName
        '
        Me.lblFirstName.AutoSize = True
        Me.lblFirstName.Location = New System.Drawing.Point(9, 19)
        Me.lblFirstName.Name = "lblFirstName"
        Me.lblFirstName.Size = New System.Drawing.Size(38, 13)
        Me.lblFirstName.TabIndex = 25
        Me.lblFirstName.Text = "&Name:"
        '
        'fraCI
        '
        Me.fraCI.Controls.Add(Me.lblPhone)
        Me.fraCI.Controls.Add(Me.cmdShipTo)
        Me.fraCI.Controls.Add(Me.lblZip)
        Me.fraCI.Controls.Add(Me.txtEmail)
        Me.fraCI.Controls.Add(Me.lblEmail)
        Me.fraCI.Controls.Add(Me.lblCityST)
        Me.fraCI.Controls.Add(Me.lblAdd1)
        Me.fraCI.Controls.Add(Me.txtPhone2)
        Me.fraCI.Controls.Add(Me.txtLastName)
        Me.fraCI.Controls.Add(Me.txtPhone1)
        Me.fraCI.Controls.Add(Me.lblFirstName)
        Me.fraCI.Controls.Add(Me.txtFirstName)
        Me.fraCI.Controls.Add(Me.txtZip)
        Me.fraCI.Controls.Add(Me.chkBusiness)
        Me.fraCI.Controls.Add(Me.txtAdd1)
        Me.fraCI.Controls.Add(Me.txtCityST)
        Me.fraCI.Controls.Add(Me.txtAdd2)
        Me.fraCI.Location = New System.Drawing.Point(5, 3)
        Me.fraCI.Name = "fraCI"
        Me.fraCI.Size = New System.Drawing.Size(353, 190)
        Me.fraCI.TabIndex = 26
        Me.fraCI.TabStop = False
        '
        'lblPhone
        '
        Me.lblPhone.AutoSize = True
        Me.lblPhone.Location = New System.Drawing.Point(9, 117)
        Me.lblPhone.Name = "lblPhone"
        Me.lblPhone.Size = New System.Drawing.Size(41, 13)
        Me.lblPhone.TabIndex = 29
        Me.lblPhone.Text = "&Phone:"
        '
        'lblZip
        '
        Me.lblZip.AutoSize = True
        Me.lblZip.Location = New System.Drawing.Point(234, 91)
        Me.lblZip.Name = "lblZip"
        Me.lblZip.Size = New System.Drawing.Size(25, 13)
        Me.lblZip.TabIndex = 28
        Me.lblZip.Text = "&Zip:"
        '
        'lblCityST
        '
        Me.lblCityST.AutoSize = True
        Me.lblCityST.Location = New System.Drawing.Point(9, 90)
        Me.lblCityST.Name = "lblCityST"
        Me.lblCityST.Size = New System.Drawing.Size(46, 13)
        Me.lblCityST.TabIndex = 27
        Me.lblCityST.Text = "&City/ST:"
        '
        'lblAdd1
        '
        Me.lblAdd1.AutoSize = True
        Me.lblAdd1.Location = New System.Drawing.Point(9, 49)
        Me.lblAdd1.Name = "lblAdd1"
        Me.lblAdd1.Size = New System.Drawing.Size(48, 13)
        Me.lblAdd1.TabIndex = 26
        Me.lblAdd1.Text = "&Address:"
        '
        'cmdOK
        '
        Me.cmdOK.Location = New System.Drawing.Point(9, 199)
        Me.cmdOK.Name = "cmdOK"
        Me.cmdOK.Size = New System.Drawing.Size(349, 29)
        Me.cmdOK.TabIndex = 27
        Me.cmdOK.Text = "&OK"
        Me.cmdOK.UseVisualStyleBackColor = True
        '
        'cmdCancel
        '
        Me.cmdCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cmdCancel.Location = New System.Drawing.Point(146, 199)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.Size = New System.Drawing.Size(75, 23)
        Me.cmdCancel.TabIndex = 30
        Me.cmdCancel.Text = "Cance&l"
        Me.cmdCancel.UseVisualStyleBackColor = True
        '
        'frmCashRegisterAddress
        '
        Me.AcceptButton = Me.cmdOK
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.cmdCancel
        Me.ClientSize = New System.Drawing.Size(363, 230)
        Me.Controls.Add(Me.cmdOK)
        Me.Controls.Add(Me.cmdCancel)
        Me.Controls.Add(Me.fraCI)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow
        Me.Name = "frmCashRegisterAddress"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Customer Information"
        Me.fraCI.ResumeLayout(False)
        Me.fraCI.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents txtEmail As TextBox
    Friend WithEvents txtPhone2 As TextBox
    Friend WithEvents txtPhone1 As TextBox
    Friend WithEvents txtZip As TextBox
    Friend WithEvents txtCityST As TextBox
    Friend WithEvents txtAdd2 As TextBox
    Friend WithEvents txtAdd1 As TextBox
    Friend WithEvents txtLastName As TextBox
    Friend WithEvents txtFirstName As TextBox
    Friend WithEvents chkBusiness As CheckBox
    Friend WithEvents lblEmail As Label
    Friend WithEvents cmdShipTo As Button
    Friend WithEvents lblFirstName As Label
    Friend WithEvents fraCI As GroupBox
    Friend WithEvents lblPhone As Label
    Friend WithEvents lblZip As Label
    Friend WithEvents lblCityST As Label
    Friend WithEvents lblAdd1 As Label
    Friend WithEvents cmdOK As Button
    Friend WithEvents cmdCancel As Button
End Class
