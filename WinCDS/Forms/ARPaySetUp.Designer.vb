<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class ARPaySetUp
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
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.dteDate1 = New System.Windows.Forms.DateTimePicker()
        Me.dteDate2 = New System.Windows.Forms.DateTimePicker()
        Me.lblDelDate = New System.Windows.Forms.Label()
        Me.lblFirstPaymentDue = New System.Windows.Forms.Label()
        Me.picPicture = New System.Windows.Forms.PictureBox()
        Me.fraLateChargesApplied = New System.Windows.Forms.GroupBox()
        Me.lblLateChargesApplied = New System.Windows.Forms.Label()
        Me.optLate6 = New System.Windows.Forms.RadioButton()
        Me.optLate16 = New System.Windows.Forms.RadioButton()
        Me.optLate26 = New System.Windows.Forms.RadioButton()
        Me.chkAutoARNO = New System.Windows.Forms.CheckBox()
        Me.txtArNo = New System.Windows.Forms.TextBox()
        Me.lblPrevBal = New System.Windows.Forms.Label()
        Me.lblGrossSale = New System.Windows.Forms.Label()
        Me.lblOrigDeposit = New System.Windows.Forms.Label()
        Me.lblSubTotal = New System.Windows.Forms.Label()
        Me.txtPrevBalance = New System.Windows.Forms.TextBox()
        Me.txtGrossSale = New System.Windows.Forms.TextBox()
        Me.txtOrigDeposit = New System.Windows.Forms.TextBox()
        Me.txtSubTotal = New System.Windows.Forms.TextBox()
        Me.GroupBox1.SuspendLayout()
        CType(Me.picPicture, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.fraLateChargesApplied.SuspendLayout()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.txtSubTotal)
        Me.GroupBox1.Controls.Add(Me.txtOrigDeposit)
        Me.GroupBox1.Controls.Add(Me.txtGrossSale)
        Me.GroupBox1.Controls.Add(Me.txtPrevBalance)
        Me.GroupBox1.Controls.Add(Me.lblSubTotal)
        Me.GroupBox1.Controls.Add(Me.lblOrigDeposit)
        Me.GroupBox1.Controls.Add(Me.lblGrossSale)
        Me.GroupBox1.Controls.Add(Me.lblPrevBal)
        Me.GroupBox1.Controls.Add(Me.txtArNo)
        Me.GroupBox1.Controls.Add(Me.chkAutoARNO)
        Me.GroupBox1.Controls.Add(Me.fraLateChargesApplied)
        Me.GroupBox1.Controls.Add(Me.picPicture)
        Me.GroupBox1.Controls.Add(Me.lblFirstPaymentDue)
        Me.GroupBox1.Controls.Add(Me.lblDelDate)
        Me.GroupBox1.Controls.Add(Me.dteDate2)
        Me.GroupBox1.Controls.Add(Me.dteDate1)
        Me.GroupBox1.Location = New System.Drawing.Point(12, 12)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(373, 438)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "GroupBox1"
        '
        'dteDate1
        '
        Me.dteDate1.CustomFormat = "MM/dd/yyyy"
        Me.dteDate1.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dteDate1.Location = New System.Drawing.Point(32, 28)
        Me.dteDate1.Name = "dteDate1"
        Me.dteDate1.Size = New System.Drawing.Size(78, 20)
        Me.dteDate1.TabIndex = 0
        '
        'dteDate2
        '
        Me.dteDate2.CustomFormat = "MM/dd/yyyy"
        Me.dteDate2.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dteDate2.Location = New System.Drawing.Point(32, 54)
        Me.dteDate2.Name = "dteDate2"
        Me.dteDate2.Size = New System.Drawing.Size(78, 20)
        Me.dteDate2.TabIndex = 1
        '
        'lblDelDate
        '
        Me.lblDelDate.AutoSize = True
        Me.lblDelDate.Location = New System.Drawing.Point(116, 34)
        Me.lblDelDate.Name = "lblDelDate"
        Me.lblDelDate.Size = New System.Drawing.Size(71, 13)
        Me.lblDelDate.TabIndex = 2
        Me.lblDelDate.Text = "Delivery Date"
        '
        'lblFirstPaymentDue
        '
        Me.lblFirstPaymentDue.AutoSize = True
        Me.lblFirstPaymentDue.Location = New System.Drawing.Point(116, 61)
        Me.lblFirstPaymentDue.Name = "lblFirstPaymentDue"
        Me.lblFirstPaymentDue.Size = New System.Drawing.Size(93, 13)
        Me.lblFirstPaymentDue.TabIndex = 3
        Me.lblFirstPaymentDue.Text = "First Payment Due"
        '
        'picPicture
        '
        Me.picPicture.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.picPicture.Location = New System.Drawing.Point(307, 19)
        Me.picPicture.Name = "picPicture"
        Me.picPicture.Size = New System.Drawing.Size(40, 29)
        Me.picPicture.TabIndex = 4
        Me.picPicture.TabStop = False
        '
        'fraLateChargesApplied
        '
        Me.fraLateChargesApplied.Controls.Add(Me.optLate26)
        Me.fraLateChargesApplied.Controls.Add(Me.optLate16)
        Me.fraLateChargesApplied.Controls.Add(Me.optLate6)
        Me.fraLateChargesApplied.Controls.Add(Me.lblLateChargesApplied)
        Me.fraLateChargesApplied.Location = New System.Drawing.Point(48, 104)
        Me.fraLateChargesApplied.Name = "fraLateChargesApplied"
        Me.fraLateChargesApplied.Size = New System.Drawing.Size(207, 101)
        Me.fraLateChargesApplied.TabIndex = 5
        Me.fraLateChargesApplied.TabStop = False
        '
        'lblLateChargesApplied
        '
        Me.lblLateChargesApplied.AutoSize = True
        Me.lblLateChargesApplied.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblLateChargesApplied.Location = New System.Drawing.Point(11, 16)
        Me.lblLateChargesApplied.Name = "lblLateChargesApplied"
        Me.lblLateChargesApplied.Size = New System.Drawing.Size(176, 13)
        Me.lblLateChargesApplied.TabIndex = 0
        Me.lblLateChargesApplied.Text = "Late Charges Will Be Applied:"
        '
        'optLate6
        '
        Me.optLate6.AutoSize = True
        Me.optLate6.Location = New System.Drawing.Point(11, 37)
        Me.optLate6.Name = "optLate6"
        Me.optLate6.Size = New System.Drawing.Size(164, 17)
        Me.optLate6.TabIndex = 1
        Me.optLate6.TabStop = True
        Me.optLate6.Text = "Due On The 1st.  Late on 6th"
        Me.optLate6.UseVisualStyleBackColor = True
        '
        'optLate16
        '
        Me.optLate16.AutoSize = True
        Me.optLate16.Location = New System.Drawing.Point(11, 58)
        Me.optLate16.Name = "optLate16"
        Me.optLate16.Size = New System.Drawing.Size(177, 17)
        Me.optLate16.TabIndex = 2
        Me.optLate16.TabStop = True
        Me.optLate16.Text = "Due On The 10th.  Late on 16th"
        Me.optLate16.UseVisualStyleBackColor = True
        '
        'optLate26
        '
        Me.optLate26.AutoSize = True
        Me.optLate26.Location = New System.Drawing.Point(11, 77)
        Me.optLate26.Name = "optLate26"
        Me.optLate26.Size = New System.Drawing.Size(177, 17)
        Me.optLate26.TabIndex = 3
        Me.optLate26.TabStop = True
        Me.optLate26.Text = "Due On The 20th.  Late on 26th"
        Me.optLate26.UseVisualStyleBackColor = True
        '
        'chkAutoARNO
        '
        Me.chkAutoARNO.AutoSize = True
        Me.chkAutoARNO.CheckAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.chkAutoARNO.Location = New System.Drawing.Point(50, 227)
        Me.chkAutoARNO.Name = "chkAutoARNO"
        Me.chkAutoARNO.Size = New System.Drawing.Size(91, 17)
        Me.chkAutoARNO.TabIndex = 6
        Me.chkAutoARNO.Text = "Auto A/R No:"
        Me.chkAutoARNO.UseVisualStyleBackColor = True
        '
        'txtArNo
        '
        Me.txtArNo.Location = New System.Drawing.Point(149, 225)
        Me.txtArNo.Name = "txtArNo"
        Me.txtArNo.Size = New System.Drawing.Size(100, 20)
        Me.txtArNo.TabIndex = 7
        '
        'lblPrevBal
        '
        Me.lblPrevBal.AutoSize = True
        Me.lblPrevBal.Location = New System.Drawing.Point(58, 258)
        Me.lblPrevBal.Name = "lblPrevBal"
        Me.lblPrevBal.Size = New System.Drawing.Size(93, 13)
        Me.lblPrevBal.TabIndex = 8
        Me.lblPrevBal.Text = "Previous Balance:"
        '
        'lblGrossSale
        '
        Me.lblGrossSale.AutoSize = True
        Me.lblGrossSale.Location = New System.Drawing.Point(58, 287)
        Me.lblGrossSale.Name = "lblGrossSale"
        Me.lblGrossSale.Size = New System.Drawing.Size(98, 13)
        Me.lblGrossSale.TabIndex = 9
        Me.lblGrossSale.Text = "Gross Sale W/Tax:"
        '
        'lblOrigDeposit
        '
        Me.lblOrigDeposit.AutoSize = True
        Me.lblOrigDeposit.Location = New System.Drawing.Point(60, 318)
        Me.lblOrigDeposit.Name = "lblOrigDeposit"
        Me.lblOrigDeposit.Size = New System.Drawing.Size(84, 13)
        Me.lblOrigDeposit.TabIndex = 10
        Me.lblOrigDeposit.Text = "Original Deposit:"
        '
        'lblSubTotal
        '
        Me.lblSubTotal.AutoSize = True
        Me.lblSubTotal.Location = New System.Drawing.Point(68, 348)
        Me.lblSubTotal.Name = "lblSubTotal"
        Me.lblSubTotal.Size = New System.Drawing.Size(56, 13)
        Me.lblSubTotal.TabIndex = 11
        Me.lblSubTotal.Text = "Sub Total:"
        '
        'txtPrevBalance
        '
        Me.txtPrevBalance.Location = New System.Drawing.Point(157, 255)
        Me.txtPrevBalance.Name = "txtPrevBalance"
        Me.txtPrevBalance.Size = New System.Drawing.Size(100, 20)
        Me.txtPrevBalance.TabIndex = 12
        Me.txtPrevBalance.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'txtGrossSale
        '
        Me.txtGrossSale.Location = New System.Drawing.Point(157, 284)
        Me.txtGrossSale.Name = "txtGrossSale"
        Me.txtGrossSale.Size = New System.Drawing.Size(100, 20)
        Me.txtGrossSale.TabIndex = 13
        Me.txtGrossSale.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'txtOrigDeposit
        '
        Me.txtOrigDeposit.Location = New System.Drawing.Point(150, 315)
        Me.txtOrigDeposit.Name = "txtOrigDeposit"
        Me.txtOrigDeposit.Size = New System.Drawing.Size(100, 20)
        Me.txtOrigDeposit.TabIndex = 14
        Me.txtOrigDeposit.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'txtSubTotal
        '
        Me.txtSubTotal.Location = New System.Drawing.Point(168, 350)
        Me.txtSubTotal.Name = "txtSubTotal"
        Me.txtSubTotal.Size = New System.Drawing.Size(100, 20)
        Me.txtSubTotal.TabIndex = 15
        Me.txtSubTotal.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'ARPaySetUp
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(809, 612)
        Me.Controls.Add(Me.GroupBox1)
        Me.Name = "ARPaySetUp"
        Me.Text = "ARPaySetUp"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        CType(Me.picPicture, System.ComponentModel.ISupportInitialize).EndInit()
        Me.fraLateChargesApplied.ResumeLayout(False)
        Me.fraLateChargesApplied.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents GroupBox1 As GroupBox
    Friend WithEvents fraLateChargesApplied As GroupBox
    Friend WithEvents optLate26 As RadioButton
    Friend WithEvents optLate16 As RadioButton
    Friend WithEvents optLate6 As RadioButton
    Friend WithEvents lblLateChargesApplied As Label
    Friend WithEvents picPicture As PictureBox
    Friend WithEvents lblFirstPaymentDue As Label
    Friend WithEvents lblDelDate As Label
    Friend WithEvents dteDate2 As DateTimePicker
    Friend WithEvents dteDate1 As DateTimePicker
    Friend WithEvents txtSubTotal As TextBox
    Friend WithEvents txtOrigDeposit As TextBox
    Friend WithEvents txtGrossSale As TextBox
    Friend WithEvents txtPrevBalance As TextBox
    Friend WithEvents lblSubTotal As Label
    Friend WithEvents lblOrigDeposit As Label
    Friend WithEvents lblGrossSale As Label
    Friend WithEvents lblPrevBal As Label
    Friend WithEvents txtArNo As TextBox
    Friend WithEvents chkAutoARNO As CheckBox
End Class
