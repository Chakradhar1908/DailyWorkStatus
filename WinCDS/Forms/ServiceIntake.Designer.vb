<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class ServiceIntake
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
        Me.lblServiceOrderNumber = New System.Windows.Forms.Label()
        Me.cboImage = New System.Windows.Forms.ComboBox()
        Me.txtServiceOrderNumber = New System.Windows.Forms.Label()
        Me.txtLocation = New System.Windows.Forms.Label()
        Me.txtVendor = New System.Windows.Forms.Label()
        Me.txtMode = New System.Windows.Forms.Label()
        Me.lblLocation = New System.Windows.Forms.Label()
        Me.lblVendor = New System.Windows.Forms.Label()
        Me.lblMode = New System.Windows.Forms.Label()
        Me.fraInfo = New System.Windows.Forms.GroupBox()
        Me.optDelivery0 = New System.Windows.Forms.RadioButton()
        Me.optDelivery1 = New System.Windows.Forms.RadioButton()
        Me.lblImage = New System.Windows.Forms.Label()
        Me.imgPicture = New System.Windows.Forms.PictureBox()
        Me.cmdEditTemplate = New System.Windows.Forms.Button()
        Me.cmdPrint = New System.Windows.Forms.Button()
        Me.cmdCancel = New System.Windows.Forms.Button()
        Me.fraInfo.SuspendLayout()
        CType(Me.imgPicture, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'lblServiceOrderNumber
        '
        Me.lblServiceOrderNumber.AutoSize = True
        Me.lblServiceOrderNumber.Location = New System.Drawing.Point(6, 14)
        Me.lblServiceOrderNumber.Name = "lblServiceOrderNumber"
        Me.lblServiceOrderNumber.Size = New System.Drawing.Size(66, 13)
        Me.lblServiceOrderNumber.TabIndex = 1
        Me.lblServiceOrderNumber.Text = "Service Call:"
        '
        'cboImage
        '
        Me.cboImage.FormattingEnabled = True
        Me.cboImage.Location = New System.Drawing.Point(302, 68)
        Me.cboImage.Name = "cboImage"
        Me.cboImage.Size = New System.Drawing.Size(105, 21)
        Me.cboImage.TabIndex = 5
        Me.cboImage.Text = "cboImage"
        '
        'txtServiceOrderNumber
        '
        Me.txtServiceOrderNumber.Location = New System.Drawing.Point(87, 14)
        Me.txtServiceOrderNumber.Name = "txtServiceOrderNumber"
        Me.txtServiceOrderNumber.Size = New System.Drawing.Size(197, 20)
        Me.txtServiceOrderNumber.TabIndex = 6
        Me.txtServiceOrderNumber.Text = "txtServiceOrderNumber"
        '
        'txtLocation
        '
        Me.txtLocation.Location = New System.Drawing.Point(87, 32)
        Me.txtLocation.Name = "txtLocation"
        Me.txtLocation.Size = New System.Drawing.Size(197, 20)
        Me.txtLocation.TabIndex = 7
        Me.txtLocation.Text = "txtLocation"
        '
        'txtVendor
        '
        Me.txtVendor.Location = New System.Drawing.Point(87, 50)
        Me.txtVendor.Name = "txtVendor"
        Me.txtVendor.Size = New System.Drawing.Size(197, 20)
        Me.txtVendor.TabIndex = 8
        Me.txtVendor.Text = "txtVendor"
        '
        'txtMode
        '
        Me.txtMode.Location = New System.Drawing.Point(87, 68)
        Me.txtMode.Name = "txtMode"
        Me.txtMode.Size = New System.Drawing.Size(197, 20)
        Me.txtMode.TabIndex = 9
        Me.txtMode.Text = "txtMode"
        '
        'lblLocation
        '
        Me.lblLocation.AutoSize = True
        Me.lblLocation.Location = New System.Drawing.Point(6, 32)
        Me.lblLocation.Name = "lblLocation"
        Me.lblLocation.Size = New System.Drawing.Size(51, 13)
        Me.lblLocation.TabIndex = 10
        Me.lblLocation.Text = "Location:"
        '
        'lblVendor
        '
        Me.lblVendor.AutoSize = True
        Me.lblVendor.Location = New System.Drawing.Point(6, 50)
        Me.lblVendor.Name = "lblVendor"
        Me.lblVendor.Size = New System.Drawing.Size(44, 13)
        Me.lblVendor.TabIndex = 11
        Me.lblVendor.Text = "Vendor:"
        '
        'lblMode
        '
        Me.lblMode.AutoSize = True
        Me.lblMode.Location = New System.Drawing.Point(6, 68)
        Me.lblMode.Name = "lblMode"
        Me.lblMode.Size = New System.Drawing.Size(82, 13)
        Me.lblMode.TabIndex = 12
        Me.lblMode.Text = "Communication:"
        '
        'fraInfo
        '
        Me.fraInfo.Controls.Add(Me.cmdCancel)
        Me.fraInfo.Controls.Add(Me.cmdPrint)
        Me.fraInfo.Controls.Add(Me.cmdEditTemplate)
        Me.fraInfo.Controls.Add(Me.imgPicture)
        Me.fraInfo.Controls.Add(Me.lblImage)
        Me.fraInfo.Controls.Add(Me.cboImage)
        Me.fraInfo.Controls.Add(Me.optDelivery1)
        Me.fraInfo.Controls.Add(Me.optDelivery0)
        Me.fraInfo.Controls.Add(Me.lblLocation)
        Me.fraInfo.Controls.Add(Me.lblMode)
        Me.fraInfo.Controls.Add(Me.lblServiceOrderNumber)
        Me.fraInfo.Controls.Add(Me.lblVendor)
        Me.fraInfo.Controls.Add(Me.txtServiceOrderNumber)
        Me.fraInfo.Controls.Add(Me.txtLocation)
        Me.fraInfo.Controls.Add(Me.txtMode)
        Me.fraInfo.Controls.Add(Me.txtVendor)
        Me.fraInfo.Location = New System.Drawing.Point(6, 2)
        Me.fraInfo.Name = "fraInfo"
        Me.fraInfo.Size = New System.Drawing.Size(426, 171)
        Me.fraInfo.TabIndex = 13
        Me.fraInfo.TabStop = False
        '
        'optDelivery0
        '
        Me.optDelivery0.AutoSize = True
        Me.optDelivery0.Location = New System.Drawing.Point(315, 14)
        Me.optDelivery0.Name = "optDelivery0"
        Me.optDelivery0.Size = New System.Drawing.Size(76, 17)
        Me.optDelivery0.TabIndex = 13
        Me.optDelivery0.TabStop = True
        Me.optDelivery0.Text = "&Print Letter"
        Me.optDelivery0.UseVisualStyleBackColor = True
        '
        'optDelivery1
        '
        Me.optDelivery1.AutoSize = True
        Me.optDelivery1.Location = New System.Drawing.Point(315, 32)
        Me.optDelivery1.Name = "optDelivery1"
        Me.optDelivery1.Size = New System.Drawing.Size(80, 17)
        Me.optDelivery1.TabIndex = 14
        Me.optDelivery1.TabStop = True
        Me.optDelivery1.Text = "&Email Letter"
        Me.optDelivery1.UseVisualStyleBackColor = True
        '
        'lblImage
        '
        Me.lblImage.AutoSize = True
        Me.lblImage.Location = New System.Drawing.Point(299, 52)
        Me.lblImage.Name = "lblImage"
        Me.lblImage.Size = New System.Drawing.Size(114, 13)
        Me.lblImage.TabIndex = 15
        Me.lblImage.Text = "Print Pi&cture On Letter:"
        '
        'imgPicture
        '
        Me.imgPicture.Location = New System.Drawing.Point(302, 92)
        Me.imgPicture.Name = "imgPicture"
        Me.imgPicture.Size = New System.Drawing.Size(27, 28)
        Me.imgPicture.TabIndex = 16
        Me.imgPicture.TabStop = False
        Me.imgPicture.Visible = False
        '
        'cmdEditTemplate
        '
        Me.cmdEditTemplate.Location = New System.Drawing.Point(317, 115)
        Me.cmdEditTemplate.Name = "cmdEditTemplate"
        Me.cmdEditTemplate.Size = New System.Drawing.Size(66, 53)
        Me.cmdEditTemplate.TabIndex = 17
        Me.cmdEditTemplate.Text = "&Template"
        Me.cmdEditTemplate.UseVisualStyleBackColor = True
        '
        'cmdPrint
        '
        Me.cmdPrint.Location = New System.Drawing.Point(144, 115)
        Me.cmdPrint.Name = "cmdPrint"
        Me.cmdPrint.Size = New System.Drawing.Size(66, 53)
        Me.cmdPrint.TabIndex = 18
        Me.cmdPrint.Text = "&OK"
        Me.cmdPrint.UseVisualStyleBackColor = True
        '
        'cmdCancel
        '
        Me.cmdCancel.Location = New System.Drawing.Point(216, 115)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.Size = New System.Drawing.Size(66, 53)
        Me.cmdCancel.TabIndex = 19
        Me.cmdCancel.Text = "&Cancel"
        Me.cmdCancel.UseVisualStyleBackColor = True
        '
        'ServiceIntake
        '
        Me.AcceptButton = Me.cmdPrint
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(435, 177)
        Me.Controls.Add(Me.fraInfo)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "ServiceIntake"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Service Charge Back Intake"
        Me.fraInfo.ResumeLayout(False)
        Me.fraInfo.PerformLayout()
        CType(Me.imgPicture, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents lblServiceOrderNumber As Label
    Friend WithEvents cboImage As ComboBox
    Friend WithEvents txtServiceOrderNumber As Label
    Friend WithEvents txtLocation As Label
    Friend WithEvents txtVendor As Label
    Friend WithEvents txtMode As Label
    Friend WithEvents lblLocation As Label
    Friend WithEvents lblVendor As Label
    Friend WithEvents lblMode As Label
    Friend WithEvents fraInfo As GroupBox
    Friend WithEvents cmdCancel As Button
    Friend WithEvents cmdPrint As Button
    Friend WithEvents cmdEditTemplate As Button
    Friend WithEvents imgPicture As PictureBox
    Friend WithEvents lblImage As Label
    Friend WithEvents optDelivery1 As RadioButton
    Friend WithEvents optDelivery0 As RadioButton
End Class
