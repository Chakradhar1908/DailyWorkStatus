<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class ArApp
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
        Me.txtFirstName = New System.Windows.Forms.TextBox()
        Me.txtLastName = New System.Windows.Forms.TextBox()
        Me.txtAddress = New System.Windows.Forms.TextBox()
        Me.txtCity = New System.Windows.Forms.TextBox()
        Me.txtZip = New System.Windows.Forms.TextBox()
        Me.txtTele1 = New System.Windows.Forms.TextBox()
        Me.txtTele2 = New System.Windows.Forms.TextBox()
        Me.lblTelephone = New System.Windows.Forms.Label()
        Me.txtCoName = New System.Windows.Forms.TextBox()
        Me.txtAccount = New System.Windows.Forms.TextBox()
        Me.SuspendLayout()
        '
        'txtFirstName
        '
        Me.txtFirstName.Location = New System.Drawing.Point(128, 43)
        Me.txtFirstName.Name = "txtFirstName"
        Me.txtFirstName.Size = New System.Drawing.Size(100, 20)
        Me.txtFirstName.TabIndex = 0
        '
        'txtLastName
        '
        Me.txtLastName.Location = New System.Drawing.Point(128, 79)
        Me.txtLastName.Name = "txtLastName"
        Me.txtLastName.Size = New System.Drawing.Size(100, 20)
        Me.txtLastName.TabIndex = 1
        '
        'txtAddress
        '
        Me.txtAddress.Location = New System.Drawing.Point(128, 125)
        Me.txtAddress.Name = "txtAddress"
        Me.txtAddress.Size = New System.Drawing.Size(100, 20)
        Me.txtAddress.TabIndex = 2
        '
        'txtCity
        '
        Me.txtCity.Location = New System.Drawing.Point(128, 151)
        Me.txtCity.Name = "txtCity"
        Me.txtCity.Size = New System.Drawing.Size(100, 20)
        Me.txtCity.TabIndex = 3
        '
        'txtZip
        '
        Me.txtZip.Location = New System.Drawing.Point(128, 177)
        Me.txtZip.Name = "txtZip"
        Me.txtZip.Size = New System.Drawing.Size(100, 20)
        Me.txtZip.TabIndex = 4
        '
        'txtTele1
        '
        Me.txtTele1.Location = New System.Drawing.Point(114, 216)
        Me.txtTele1.Name = "txtTele1"
        Me.txtTele1.Size = New System.Drawing.Size(100, 20)
        Me.txtTele1.TabIndex = 5
        '
        'txtTele2
        '
        Me.txtTele2.Location = New System.Drawing.Point(114, 259)
        Me.txtTele2.Name = "txtTele2"
        Me.txtTele2.Size = New System.Drawing.Size(100, 20)
        Me.txtTele2.TabIndex = 6
        '
        'lblTelephone
        '
        Me.lblTelephone.AutoSize = True
        Me.lblTelephone.Location = New System.Drawing.Point(125, 324)
        Me.lblTelephone.Name = "lblTelephone"
        Me.lblTelephone.Size = New System.Drawing.Size(68, 13)
        Me.lblTelephone.TabIndex = 7
        Me.lblTelephone.Text = "lblTelephone"
        '
        'txtCoName
        '
        Me.txtCoName.Location = New System.Drawing.Point(114, 359)
        Me.txtCoName.Name = "txtCoName"
        Me.txtCoName.Size = New System.Drawing.Size(100, 20)
        Me.txtCoName.TabIndex = 8
        '
        'txtAccount
        '
        Me.txtAccount.Location = New System.Drawing.Point(114, 385)
        Me.txtAccount.Name = "txtAccount"
        Me.txtAccount.Size = New System.Drawing.Size(100, 20)
        Me.txtAccount.TabIndex = 9
        '
        'ArApp
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(800, 450)
        Me.Controls.Add(Me.txtAccount)
        Me.Controls.Add(Me.txtCoName)
        Me.Controls.Add(Me.lblTelephone)
        Me.Controls.Add(Me.txtTele2)
        Me.Controls.Add(Me.txtTele1)
        Me.Controls.Add(Me.txtZip)
        Me.Controls.Add(Me.txtCity)
        Me.Controls.Add(Me.txtAddress)
        Me.Controls.Add(Me.txtLastName)
        Me.Controls.Add(Me.txtFirstName)
        Me.Name = "ArApp"
        Me.Text = "ArApp"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents txtFirstName As TextBox
    Friend WithEvents txtLastName As TextBox
    Friend WithEvents txtAddress As TextBox
    Friend WithEvents txtCity As TextBox
    Friend WithEvents txtZip As TextBox
    Friend WithEvents txtTele1 As TextBox
    Friend WithEvents txtTele2 As TextBox
    Friend WithEvents lblTelephone As Label
    Friend WithEvents txtCoName As TextBox
    Friend WithEvents txtAccount As TextBox
End Class
