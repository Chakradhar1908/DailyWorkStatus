<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class OrdCashflow
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
        Me.lblCashIn = New System.Windows.Forms.Label()
        Me.cboCashIn = New System.Windows.Forms.ComboBox()
        Me.lblCashOut = New System.Windows.Forms.Label()
        Me.cboCashOut = New System.Windows.Forms.ComboBox()
        Me.lblAuditNote = New System.Windows.Forms.Label()
        Me.txtAuditNote = New System.Windows.Forms.TextBox()
        Me.lblAccount = New System.Windows.Forms.Label()
        Me.txtAccount = New System.Windows.Forms.TextBox()
        Me.lblAmount = New System.Windows.Forms.Label()
        Me.txtAmount = New System.Windows.Forms.TextBox()
        Me.DDate = New System.Windows.Forms.DateTimePicker()
        Me.fraControls = New System.Windows.Forms.GroupBox()
        Me.cmdCancel = New System.Windows.Forms.Button()
        Me.cmdNext = New System.Windows.Forms.Button()
        Me.cmdPost = New System.Windows.Forms.Button()
        Me.pic = New System.Windows.Forms.PictureBox()
        Me.fraControls.SuspendLayout()
        CType(Me.pic, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'lblCashIn
        '
        Me.lblCashIn.AutoSize = True
        Me.lblCashIn.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblCashIn.Location = New System.Drawing.Point(33, 7)
        Me.lblCashIn.Name = "lblCashIn"
        Me.lblCashIn.Size = New System.Drawing.Size(78, 20)
        Me.lblCashIn.TabIndex = 0
        Me.lblCashIn.Text = "CASH &IN:"
        '
        'cboCashIn
        '
        Me.cboCashIn.FormattingEnabled = True
        Me.cboCashIn.Location = New System.Drawing.Point(12, 30)
        Me.cboCashIn.Name = "cboCashIn"
        Me.cboCashIn.Size = New System.Drawing.Size(121, 21)
        Me.cboCashIn.TabIndex = 1
        '
        'lblCashOut
        '
        Me.lblCashOut.AutoSize = True
        Me.lblCashOut.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblCashOut.Location = New System.Drawing.Point(269, 7)
        Me.lblCashOut.Name = "lblCashOut"
        Me.lblCashOut.Size = New System.Drawing.Size(95, 20)
        Me.lblCashOut.TabIndex = 2
        Me.lblCashOut.Text = "CASH &OUT:"
        '
        'cboCashOut
        '
        Me.cboCashOut.FormattingEnabled = True
        Me.cboCashOut.Location = New System.Drawing.Point(255, 30)
        Me.cboCashOut.Name = "cboCashOut"
        Me.cboCashOut.Size = New System.Drawing.Size(121, 21)
        Me.cboCashOut.TabIndex = 3
        '
        'lblAuditNote
        '
        Me.lblAuditNote.AutoSize = True
        Me.lblAuditNote.Location = New System.Drawing.Point(51, 151)
        Me.lblAuditNote.Name = "lblAuditNote"
        Me.lblAuditNote.Size = New System.Drawing.Size(72, 13)
        Me.lblAuditNote.TabIndex = 5
        Me.lblAuditNote.Text = "&Note on Audit"
        Me.lblAuditNote.Visible = False
        '
        'txtAuditNote
        '
        Me.txtAuditNote.Location = New System.Drawing.Point(129, 148)
        Me.txtAuditNote.Name = "txtAuditNote"
        Me.txtAuditNote.Size = New System.Drawing.Size(172, 20)
        Me.txtAuditNote.TabIndex = 6
        Me.txtAuditNote.Visible = False
        '
        'lblAccount
        '
        Me.lblAccount.AutoSize = True
        Me.lblAccount.Location = New System.Drawing.Point(17, 174)
        Me.lblAccount.Name = "lblAccount"
        Me.lblAccount.Size = New System.Drawing.Size(46, 13)
        Me.lblAccount.TabIndex = 7
        Me.lblAccount.Text = "A&cc No:"
        '
        'txtAccount
        '
        Me.txtAccount.Location = New System.Drawing.Point(60, 171)
        Me.txtAccount.Name = "txtAccount"
        Me.txtAccount.Size = New System.Drawing.Size(67, 20)
        Me.txtAccount.TabIndex = 8
        '
        'lblAmount
        '
        Me.lblAmount.AutoSize = True
        Me.lblAmount.Location = New System.Drawing.Point(142, 178)
        Me.lblAmount.Name = "lblAmount"
        Me.lblAmount.Size = New System.Drawing.Size(46, 13)
        Me.lblAmount.TabIndex = 9
        Me.lblAmount.Text = "&Amount:"
        '
        'txtAmount
        '
        Me.txtAmount.Location = New System.Drawing.Point(187, 171)
        Me.txtAmount.Name = "txtAmount"
        Me.txtAmount.Size = New System.Drawing.Size(73, 20)
        Me.txtAmount.TabIndex = 10
        '
        'DDate
        '
        Me.DDate.CustomFormat = "MM/dd/yyyy"
        Me.DDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DDate.Location = New System.Drawing.Point(287, 171)
        Me.DDate.Name = "DDate"
        Me.DDate.Size = New System.Drawing.Size(89, 20)
        Me.DDate.TabIndex = 11
        '
        'fraControls
        '
        Me.fraControls.Controls.Add(Me.cmdCancel)
        Me.fraControls.Controls.Add(Me.cmdNext)
        Me.fraControls.Controls.Add(Me.cmdPost)
        Me.fraControls.Location = New System.Drawing.Point(65, 199)
        Me.fraControls.Name = "fraControls"
        Me.fraControls.Size = New System.Drawing.Size(252, 70)
        Me.fraControls.TabIndex = 12
        Me.fraControls.TabStop = False
        '
        'cmdCancel
        '
        Me.cmdCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cmdCancel.Location = New System.Drawing.Point(168, 14)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.Size = New System.Drawing.Size(77, 49)
        Me.cmdCancel.TabIndex = 2
        Me.cmdCancel.Text = "&Cancel"
        Me.cmdCancel.UseVisualStyleBackColor = True
        '
        'cmdNext
        '
        Me.cmdNext.Location = New System.Drawing.Point(85, 14)
        Me.cmdNext.Name = "cmdNext"
        Me.cmdNext.Size = New System.Drawing.Size(77, 49)
        Me.cmdNext.TabIndex = 1
        Me.cmdNext.Text = "&Next"
        Me.cmdNext.UseVisualStyleBackColor = True
        '
        'cmdPost
        '
        Me.cmdPost.Location = New System.Drawing.Point(7, 14)
        Me.cmdPost.Name = "cmdPost"
        Me.cmdPost.Size = New System.Drawing.Size(77, 49)
        Me.cmdPost.TabIndex = 0
        Me.cmdPost.Text = "&Post"
        Me.cmdPost.UseVisualStyleBackColor = True
        '
        'pic
        '
        Me.pic.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.pic.Location = New System.Drawing.Point(139, 64)
        Me.pic.Name = "pic"
        Me.pic.Size = New System.Drawing.Size(112, 78)
        Me.pic.TabIndex = 4
        Me.pic.TabStop = False
        '
        'OrdCashflow
        '
        Me.AcceptButton = Me.cmdPost
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.cmdCancel
        Me.ClientSize = New System.Drawing.Size(382, 271)
        Me.Controls.Add(Me.fraControls)
        Me.Controls.Add(Me.DDate)
        Me.Controls.Add(Me.txtAmount)
        Me.Controls.Add(Me.lblAmount)
        Me.Controls.Add(Me.txtAccount)
        Me.Controls.Add(Me.lblAccount)
        Me.Controls.Add(Me.txtAuditNote)
        Me.Controls.Add(Me.lblAuditNote)
        Me.Controls.Add(Me.pic)
        Me.Controls.Add(Me.cboCashOut)
        Me.Controls.Add(Me.lblCashOut)
        Me.Controls.Add(Me.cboCashIn)
        Me.Controls.Add(Me.lblCashIn)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.MaximizeBox = False
        Me.Name = "OrdCashflow"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Cash Drawer"
        Me.fraControls.ResumeLayout(False)
        CType(Me.pic, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents lblCashIn As Label
    Friend WithEvents cboCashIn As ComboBox
    Friend WithEvents lblCashOut As Label
    Friend WithEvents cboCashOut As ComboBox
    Friend WithEvents pic As PictureBox
    Friend WithEvents lblAuditNote As Label
    Friend WithEvents txtAuditNote As TextBox
    Friend WithEvents lblAccount As Label
    Friend WithEvents txtAccount As TextBox
    Friend WithEvents lblAmount As Label
    Friend WithEvents txtAmount As TextBox
    Friend WithEvents DDate As DateTimePicker
    Friend WithEvents fraControls As GroupBox
    Friend WithEvents cmdCancel As Button
    Friend WithEvents cmdNext As Button
    Friend WithEvents cmdPost As Button
End Class
