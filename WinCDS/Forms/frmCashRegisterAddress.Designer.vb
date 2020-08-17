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
        Me.CheckBox2 = New System.Windows.Forms.CheckBox()
        Me.chkBusiness = New System.Windows.Forms.CheckBox()
        Me.lblEmail = New System.Windows.Forms.Label()
        Me.cmdShipTo = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'txtEmail
        '
        Me.txtEmail.Location = New System.Drawing.Point(346, 347)
        Me.txtEmail.Name = "txtEmail"
        Me.txtEmail.Size = New System.Drawing.Size(100, 20)
        Me.txtEmail.TabIndex = 22
        '
        'txtPhone2
        '
        Me.txtPhone2.Location = New System.Drawing.Point(346, 321)
        Me.txtPhone2.Name = "txtPhone2"
        Me.txtPhone2.Size = New System.Drawing.Size(100, 20)
        Me.txtPhone2.TabIndex = 21
        '
        'txtPhone1
        '
        Me.txtPhone1.Location = New System.Drawing.Point(346, 295)
        Me.txtPhone1.Name = "txtPhone1"
        Me.txtPhone1.Size = New System.Drawing.Size(100, 20)
        Me.txtPhone1.TabIndex = 20
        '
        'txtZip
        '
        Me.txtZip.Location = New System.Drawing.Point(346, 260)
        Me.txtZip.Name = "txtZip"
        Me.txtZip.Size = New System.Drawing.Size(100, 20)
        Me.txtZip.TabIndex = 19
        '
        'txtCityST
        '
        Me.txtCityST.Location = New System.Drawing.Point(354, 234)
        Me.txtCityST.Name = "txtCityST"
        Me.txtCityST.Size = New System.Drawing.Size(100, 20)
        Me.txtCityST.TabIndex = 18
        '
        'txtAdd2
        '
        Me.txtAdd2.Location = New System.Drawing.Point(354, 208)
        Me.txtAdd2.Name = "txtAdd2"
        Me.txtAdd2.Size = New System.Drawing.Size(100, 20)
        Me.txtAdd2.TabIndex = 17
        '
        'txtAdd1
        '
        Me.txtAdd1.Location = New System.Drawing.Point(354, 182)
        Me.txtAdd1.Name = "txtAdd1"
        Me.txtAdd1.Size = New System.Drawing.Size(100, 20)
        Me.txtAdd1.TabIndex = 16
        '
        'txtLastName
        '
        Me.txtLastName.Location = New System.Drawing.Point(354, 156)
        Me.txtLastName.Name = "txtLastName"
        Me.txtLastName.Size = New System.Drawing.Size(100, 20)
        Me.txtLastName.TabIndex = 15
        '
        'txtFirstName
        '
        Me.txtFirstName.Location = New System.Drawing.Point(354, 130)
        Me.txtFirstName.Name = "txtFirstName"
        Me.txtFirstName.Size = New System.Drawing.Size(100, 20)
        Me.txtFirstName.TabIndex = 14
        '
        'CheckBox2
        '
        Me.CheckBox2.AutoSize = True
        Me.CheckBox2.Location = New System.Drawing.Point(346, 106)
        Me.CheckBox2.Name = "CheckBox2"
        Me.CheckBox2.Size = New System.Drawing.Size(81, 17)
        Me.CheckBox2.TabIndex = 13
        Me.CheckBox2.Text = "CheckBox2"
        Me.CheckBox2.UseVisualStyleBackColor = True
        '
        'chkBusiness
        '
        Me.chkBusiness.AutoSize = True
        Me.chkBusiness.Location = New System.Drawing.Point(346, 83)
        Me.chkBusiness.Name = "chkBusiness"
        Me.chkBusiness.Size = New System.Drawing.Size(81, 17)
        Me.chkBusiness.TabIndex = 12
        Me.chkBusiness.Text = "CheckBox1"
        Me.chkBusiness.UseVisualStyleBackColor = True
        '
        'lblEmail
        '
        Me.lblEmail.AutoSize = True
        Me.lblEmail.Location = New System.Drawing.Point(0, 0)
        Me.lblEmail.Name = "lblEmail"
        Me.lblEmail.Size = New System.Drawing.Size(39, 13)
        Me.lblEmail.TabIndex = 23
        Me.lblEmail.Text = "Label1"
        '
        'cmdShipTo
        '
        Me.cmdShipTo.Location = New System.Drawing.Point(3, 35)
        Me.cmdShipTo.Name = "cmdShipTo"
        Me.cmdShipTo.Size = New System.Drawing.Size(75, 23)
        Me.cmdShipTo.TabIndex = 24
        Me.cmdShipTo.Text = "Button1"
        Me.cmdShipTo.UseVisualStyleBackColor = True
        '
        'frmCashRegisterAddress
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(800, 450)
        Me.Controls.Add(Me.cmdShipTo)
        Me.Controls.Add(Me.lblEmail)
        Me.Controls.Add(Me.txtEmail)
        Me.Controls.Add(Me.txtPhone2)
        Me.Controls.Add(Me.txtPhone1)
        Me.Controls.Add(Me.txtZip)
        Me.Controls.Add(Me.txtCityST)
        Me.Controls.Add(Me.txtAdd2)
        Me.Controls.Add(Me.txtAdd1)
        Me.Controls.Add(Me.txtLastName)
        Me.Controls.Add(Me.txtFirstName)
        Me.Controls.Add(Me.CheckBox2)
        Me.Controls.Add(Me.chkBusiness)
        Me.Name = "frmCashRegisterAddress"
        Me.Text = "frmCashRegisterAddress"
        Me.ResumeLayout(False)
        Me.PerformLayout()

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
    Friend WithEvents CheckBox2 As CheckBox
    Friend WithEvents chkBusiness As CheckBox
    Friend WithEvents lblEmail As Label
    Friend WithEvents cmdShipTo As Button
End Class
