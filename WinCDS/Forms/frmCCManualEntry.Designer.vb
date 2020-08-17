<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmCCManualEntry
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
        Me.cmdOK = New System.Windows.Forms.Button()
        Me.cmdCancel = New System.Windows.Forms.Button()
        Me.lblZipCode = New System.Windows.Forms.Label()
        Me.txtZipCode = New System.Windows.Forms.TextBox()
        Me.txtCardHolderName = New System.Windows.Forms.TextBox()
        Me.lblCVV2 = New System.Windows.Forms.Label()
        Me.txtCVV2 = New System.Windows.Forms.TextBox()
        Me.lblExpDate = New System.Windows.Forms.Label()
        Me.txtExpDate = New System.Windows.Forms.TextBox()
        Me.fraCC = New System.Windows.Forms.GroupBox()
        Me.txtCCNumber = New System.Windows.Forms.TextBox()
        Me.lblCardHolderName = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'cmdOK
        '
        Me.cmdOK.Location = New System.Drawing.Point(12, 12)
        Me.cmdOK.Name = "cmdOK"
        Me.cmdOK.Size = New System.Drawing.Size(75, 23)
        Me.cmdOK.TabIndex = 0
        Me.cmdOK.Text = "Button1"
        Me.cmdOK.UseVisualStyleBackColor = True
        '
        'cmdCancel
        '
        Me.cmdCancel.Location = New System.Drawing.Point(12, 46)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.Size = New System.Drawing.Size(75, 23)
        Me.cmdCancel.TabIndex = 1
        Me.cmdCancel.Text = "Button1"
        Me.cmdCancel.UseVisualStyleBackColor = True
        '
        'lblZipCode
        '
        Me.lblZipCode.AutoSize = True
        Me.lblZipCode.Location = New System.Drawing.Point(21, 90)
        Me.lblZipCode.Name = "lblZipCode"
        Me.lblZipCode.Size = New System.Drawing.Size(39, 13)
        Me.lblZipCode.TabIndex = 2
        Me.lblZipCode.Text = "Label1"
        '
        'txtZipCode
        '
        Me.txtZipCode.Location = New System.Drawing.Point(24, 106)
        Me.txtZipCode.Name = "txtZipCode"
        Me.txtZipCode.Size = New System.Drawing.Size(100, 20)
        Me.txtZipCode.TabIndex = 3
        '
        'txtCardHolderName
        '
        Me.txtCardHolderName.Location = New System.Drawing.Point(24, 132)
        Me.txtCardHolderName.Name = "txtCardHolderName"
        Me.txtCardHolderName.Size = New System.Drawing.Size(100, 20)
        Me.txtCardHolderName.TabIndex = 4
        '
        'lblCVV2
        '
        Me.lblCVV2.AutoSize = True
        Me.lblCVV2.Location = New System.Drawing.Point(21, 164)
        Me.lblCVV2.Name = "lblCVV2"
        Me.lblCVV2.Size = New System.Drawing.Size(39, 13)
        Me.lblCVV2.TabIndex = 5
        Me.lblCVV2.Text = "Label1"
        '
        'txtCVV2
        '
        Me.txtCVV2.Location = New System.Drawing.Point(24, 197)
        Me.txtCVV2.Name = "txtCVV2"
        Me.txtCVV2.Size = New System.Drawing.Size(100, 20)
        Me.txtCVV2.TabIndex = 6
        '
        'lblExpDate
        '
        Me.lblExpDate.AutoSize = True
        Me.lblExpDate.Location = New System.Drawing.Point(21, 229)
        Me.lblExpDate.Name = "lblExpDate"
        Me.lblExpDate.Size = New System.Drawing.Size(39, 13)
        Me.lblExpDate.TabIndex = 7
        Me.lblExpDate.Text = "Label1"
        '
        'txtExpDate
        '
        Me.txtExpDate.Location = New System.Drawing.Point(12, 272)
        Me.txtExpDate.Name = "txtExpDate"
        Me.txtExpDate.Size = New System.Drawing.Size(100, 20)
        Me.txtExpDate.TabIndex = 8
        '
        'fraCC
        '
        Me.fraCC.Location = New System.Drawing.Point(12, 311)
        Me.fraCC.Name = "fraCC"
        Me.fraCC.Size = New System.Drawing.Size(200, 100)
        Me.fraCC.TabIndex = 9
        Me.fraCC.TabStop = False
        Me.fraCC.Text = "GroupBox1"
        '
        'txtCCNumber
        '
        Me.txtCCNumber.Location = New System.Drawing.Point(185, 28)
        Me.txtCCNumber.Name = "txtCCNumber"
        Me.txtCCNumber.Size = New System.Drawing.Size(100, 20)
        Me.txtCCNumber.TabIndex = 10
        '
        'lblCardHolderName
        '
        Me.lblCardHolderName.AutoSize = True
        Me.lblCardHolderName.Location = New System.Drawing.Point(182, 66)
        Me.lblCardHolderName.Name = "lblCardHolderName"
        Me.lblCardHolderName.Size = New System.Drawing.Size(39, 13)
        Me.lblCardHolderName.TabIndex = 11
        Me.lblCardHolderName.Text = "Label1"
        '
        'frmCCManualEntry
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(800, 450)
        Me.Controls.Add(Me.lblCardHolderName)
        Me.Controls.Add(Me.txtCCNumber)
        Me.Controls.Add(Me.fraCC)
        Me.Controls.Add(Me.txtExpDate)
        Me.Controls.Add(Me.lblExpDate)
        Me.Controls.Add(Me.txtCVV2)
        Me.Controls.Add(Me.lblCVV2)
        Me.Controls.Add(Me.txtCardHolderName)
        Me.Controls.Add(Me.txtZipCode)
        Me.Controls.Add(Me.lblZipCode)
        Me.Controls.Add(Me.cmdCancel)
        Me.Controls.Add(Me.cmdOK)
        Me.Name = "frmCCManualEntry"
        Me.Text = "frmCCManualEntry"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents cmdOK As Button
    Friend WithEvents cmdCancel As Button
    Friend WithEvents lblZipCode As Label
    Friend WithEvents txtZipCode As TextBox
    Friend WithEvents txtCardHolderName As TextBox
    Friend WithEvents lblCVV2 As Label
    Friend WithEvents txtCVV2 As TextBox
    Friend WithEvents lblExpDate As Label
    Friend WithEvents txtExpDate As TextBox
    Friend WithEvents fraCC As GroupBox
    Friend WithEvents txtCCNumber As TextBox
    Friend WithEvents lblCardHolderName As Label
End Class
