<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmEmail
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
        Me.lblFromEmail = New System.Windows.Forms.Label()
        Me.lblFromName = New System.Windows.Forms.Label()
        Me.txtFromAddr = New System.Windows.Forms.TextBox()
        Me.txtFromName = New System.Windows.Forms.TextBox()
        Me.prg = New System.Windows.Forms.ProgressBar()
        Me.txt = New System.Windows.Forms.TextBox()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.cmdOK = New System.Windows.Forms.Button()
        Me.cmdMail = New System.Windows.Forms.Button()
        Me.chkReprint = New System.Windows.Forms.CheckBox()
        Me.chkPrintPO = New System.Windows.Forms.CheckBox()
        Me.fraByDate = New System.Windows.Forms.GroupBox()
        Me.dtpToDate = New System.Windows.Forms.DateTimePicker()
        Me.dtpFromDate = New System.Windows.Forms.DateTimePicker()
        Me.lblToDate = New System.Windows.Forms.Label()
        Me.lblFromDate = New System.Windows.Forms.Label()
        Me.optByPoNo = New System.Windows.Forms.RadioButton()
        Me.optByDate = New System.Windows.Forms.RadioButton()
        Me.fraResults = New System.Windows.Forms.GroupBox()
        Me.lstResults = New System.Windows.Forms.CheckedListBox()
        Me.lblStatus = New System.Windows.Forms.Label()
        Me.tmr = New System.Windows.Forms.Timer(Me.components)
        Me.fraByPoNo = New System.Windows.Forms.GroupBox()
        Me.lblToPo = New System.Windows.Forms.Label()
        Me.lblFromPO = New System.Windows.Forms.Label()
        Me.txtFromPO = New System.Windows.Forms.TextBox()
        Me.txtToPO = New System.Windows.Forms.TextBox()
        Me.GroupBox1.SuspendLayout()
        Me.fraByDate.SuspendLayout()
        Me.fraResults.SuspendLayout()
        Me.fraByPoNo.SuspendLayout()
        Me.SuspendLayout()
        '
        'lblFromEmail
        '
        Me.lblFromEmail.AutoSize = True
        Me.lblFromEmail.Location = New System.Drawing.Point(12, 6)
        Me.lblFromEmail.Name = "lblFromEmail"
        Me.lblFromEmail.Size = New System.Drawing.Size(58, 13)
        Me.lblFromEmail.TabIndex = 0
        Me.lblFromEmail.Text = "From &Addr:"
        '
        'lblFromName
        '
        Me.lblFromName.AutoSize = True
        Me.lblFromName.Location = New System.Drawing.Point(12, 29)
        Me.lblFromName.Name = "lblFromName"
        Me.lblFromName.Size = New System.Drawing.Size(64, 13)
        Me.lblFromName.TabIndex = 1
        Me.lblFromName.Text = "From &Name:"
        '
        'txtFromAddr
        '
        Me.txtFromAddr.Location = New System.Drawing.Point(82, 3)
        Me.txtFromAddr.Name = "txtFromAddr"
        Me.txtFromAddr.Size = New System.Drawing.Size(166, 20)
        Me.txtFromAddr.TabIndex = 2
        '
        'txtFromName
        '
        Me.txtFromName.Location = New System.Drawing.Point(82, 32)
        Me.txtFromName.Name = "txtFromName"
        Me.txtFromName.Size = New System.Drawing.Size(166, 20)
        Me.txtFromName.TabIndex = 3
        '
        'prg
        '
        Me.prg.Location = New System.Drawing.Point(12, 58)
        Me.prg.Name = "prg"
        Me.prg.Size = New System.Drawing.Size(233, 12)
        Me.prg.TabIndex = 4
        '
        'txt
        '
        Me.txt.Location = New System.Drawing.Point(12, 76)
        Me.txt.Multiline = True
        Me.txt.Name = "txt"
        Me.txt.ScrollBars = System.Windows.Forms.ScrollBars.Both
        Me.txt.Size = New System.Drawing.Size(233, 193)
        Me.txt.TabIndex = 5
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.fraByPoNo)
        Me.GroupBox1.Controls.Add(Me.cmdOK)
        Me.GroupBox1.Controls.Add(Me.cmdMail)
        Me.GroupBox1.Controls.Add(Me.chkReprint)
        Me.GroupBox1.Controls.Add(Me.chkPrintPO)
        Me.GroupBox1.Controls.Add(Me.optByPoNo)
        Me.GroupBox1.Controls.Add(Me.optByDate)
        Me.GroupBox1.Location = New System.Drawing.Point(253, -2)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(180, 271)
        Me.GroupBox1.TabIndex = 6
        Me.GroupBox1.TabStop = False
        '
        'cmdOK
        '
        Me.cmdOK.Location = New System.Drawing.Point(85, 208)
        Me.cmdOK.Name = "cmdOK"
        Me.cmdOK.Size = New System.Drawing.Size(75, 49)
        Me.cmdOK.TabIndex = 6
        Me.cmdOK.Text = "&Close"
        Me.cmdOK.UseVisualStyleBackColor = True
        '
        'cmdMail
        '
        Me.cmdMail.Location = New System.Drawing.Point(6, 208)
        Me.cmdMail.Name = "cmdMail"
        Me.cmdMail.Size = New System.Drawing.Size(75, 49)
        Me.cmdMail.TabIndex = 5
        Me.cmdMail.Text = "&Email POs"
        Me.cmdMail.UseVisualStyleBackColor = True
        '
        'chkReprint
        '
        Me.chkReprint.AutoSize = True
        Me.chkReprint.Location = New System.Drawing.Point(6, 185)
        Me.chkReprint.Name = "chkReprint"
        Me.chkReprint.Size = New System.Drawing.Size(102, 17)
        Me.chkReprint.TabIndex = 4
        Me.chkReprint.Text = "Re-Send Email?"
        Me.chkReprint.UseVisualStyleBackColor = True
        '
        'chkPrintPO
        '
        Me.chkPrintPO.AutoSize = True
        Me.chkPrintPO.Location = New System.Drawing.Point(6, 162)
        Me.chkPrintPO.Name = "chkPrintPO"
        Me.chkPrintPO.Size = New System.Drawing.Size(110, 17)
        Me.chkPrintPO.TabIndex = 3
        Me.chkPrintPO.Text = "Print PO As Well?"
        Me.chkPrintPO.UseVisualStyleBackColor = True
        '
        'fraByDate
        '
        Me.fraByDate.Controls.Add(Me.dtpToDate)
        Me.fraByDate.Controls.Add(Me.dtpFromDate)
        Me.fraByDate.Controls.Add(Me.lblToDate)
        Me.fraByDate.Controls.Add(Me.lblFromDate)
        Me.fraByDate.Font = New System.Drawing.Font("Arial Black", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.fraByDate.Location = New System.Drawing.Point(0, 0)
        Me.fraByDate.Name = "fraByDate"
        Me.fraByDate.Size = New System.Drawing.Size(158, 71)
        Me.fraByDate.TabIndex = 2
        Me.fraByDate.TabStop = False
        Me.fraByDate.Text = "Run POs By Date:"
        '
        'dtpToDate
        '
        Me.dtpToDate.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dtpToDate.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtpToDate.Location = New System.Drawing.Point(45, 42)
        Me.dtpToDate.Name = "dtpToDate"
        Me.dtpToDate.Size = New System.Drawing.Size(76, 20)
        Me.dtpToDate.TabIndex = 3
        '
        'dtpFromDate
        '
        Me.dtpFromDate.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dtpFromDate.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtpFromDate.Location = New System.Drawing.Point(45, 16)
        Me.dtpFromDate.Name = "dtpFromDate"
        Me.dtpFromDate.Size = New System.Drawing.Size(76, 20)
        Me.dtpFromDate.TabIndex = 2
        '
        'lblToDate
        '
        Me.lblToDate.AutoSize = True
        Me.lblToDate.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblToDate.Location = New System.Drawing.Point(12, 40)
        Me.lblToDate.Name = "lblToDate"
        Me.lblToDate.Size = New System.Drawing.Size(23, 13)
        Me.lblToDate.TabIndex = 1
        Me.lblToDate.Text = "&To:"
        '
        'lblFromDate
        '
        Me.lblFromDate.AutoSize = True
        Me.lblFromDate.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblFromDate.Location = New System.Drawing.Point(12, 16)
        Me.lblFromDate.Name = "lblFromDate"
        Me.lblFromDate.Size = New System.Drawing.Size(33, 13)
        Me.lblFromDate.TabIndex = 0
        Me.lblFromDate.Text = "&From:"
        '
        'optByPoNo
        '
        Me.optByPoNo.AutoSize = True
        Me.optByPoNo.Location = New System.Drawing.Point(85, 46)
        Me.optByPoNo.Name = "optByPoNo"
        Me.optByPoNo.Size = New System.Drawing.Size(67, 17)
        Me.optByPoNo.TabIndex = 1
        Me.optByPoNo.TabStop = True
        Me.optByPoNo.Text = "By PoNo"
        Me.optByPoNo.UseVisualStyleBackColor = True
        '
        'optByDate
        '
        Me.optByDate.AutoSize = True
        Me.optByDate.Location = New System.Drawing.Point(16, 46)
        Me.optByDate.Name = "optByDate"
        Me.optByDate.Size = New System.Drawing.Size(63, 17)
        Me.optByDate.TabIndex = 0
        Me.optByDate.TabStop = True
        Me.optByDate.Text = "By Date"
        Me.optByDate.UseVisualStyleBackColor = True
        '
        'fraResults
        '
        Me.fraResults.Controls.Add(Me.lstResults)
        Me.fraResults.Controls.Add(Me.lblStatus)
        Me.fraResults.Font = New System.Drawing.Font("Arial Black", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.fraResults.Location = New System.Drawing.Point(439, 44)
        Me.fraResults.Name = "fraResults"
        Me.fraResults.Size = New System.Drawing.Size(200, 225)
        Me.fraResults.TabIndex = 7
        Me.fraResults.TabStop = False
        Me.fraResults.Text = "Results:"
        '
        'lstResults
        '
        Me.lstResults.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lstResults.FormattingEnabled = True
        Me.lstResults.Location = New System.Drawing.Point(7, 62)
        Me.lstResults.Name = "lstResults"
        Me.lstResults.Size = New System.Drawing.Size(184, 154)
        Me.lstResults.TabIndex = 1
        '
        'lblStatus
        '
        Me.lblStatus.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblStatus.ForeColor = System.Drawing.SystemColors.Highlight
        Me.lblStatus.Location = New System.Drawing.Point(6, 32)
        Me.lblStatus.Name = "lblStatus"
        Me.lblStatus.Size = New System.Drawing.Size(152, 23)
        Me.lblStatus.TabIndex = 0
        Me.lblStatus.Text = "###"
        Me.lblStatus.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'fraByPoNo
        '
        Me.fraByPoNo.Controls.Add(Me.txtToPO)
        Me.fraByPoNo.Controls.Add(Me.txtFromPO)
        Me.fraByPoNo.Controls.Add(Me.lblFromPO)
        Me.fraByPoNo.Controls.Add(Me.lblToPo)
        Me.fraByPoNo.Controls.Add(Me.fraByDate)
        Me.fraByPoNo.Font = New System.Drawing.Font("Arial Black", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.fraByPoNo.Location = New System.Drawing.Point(13, 73)
        Me.fraByPoNo.Name = "fraByPoNo"
        Me.fraByPoNo.Size = New System.Drawing.Size(138, 67)
        Me.fraByPoNo.TabIndex = 7
        Me.fraByPoNo.TabStop = False
        Me.fraByPoNo.Text = "Run POs By PoNo:"
        '
        'lblToPo
        '
        Me.lblToPo.AutoSize = True
        Me.lblToPo.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblToPo.Location = New System.Drawing.Point(6, 19)
        Me.lblToPo.Name = "lblToPo"
        Me.lblToPo.Size = New System.Drawing.Size(33, 13)
        Me.lblToPo.TabIndex = 0
        Me.lblToPo.Text = "&From:"
        '
        'lblFromPO
        '
        Me.lblFromPO.AutoSize = True
        Me.lblFromPO.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblFromPO.Location = New System.Drawing.Point(6, 38)
        Me.lblFromPO.Name = "lblFromPO"
        Me.lblFromPO.Size = New System.Drawing.Size(23, 13)
        Me.lblFromPO.TabIndex = 1
        Me.lblFromPO.Text = "&To:"
        '
        'txtFromPO
        '
        Me.txtFromPO.Location = New System.Drawing.Point(35, 15)
        Me.txtFromPO.Name = "txtFromPO"
        Me.txtFromPO.Size = New System.Drawing.Size(70, 23)
        Me.txtFromPO.TabIndex = 2
        '
        'txtToPO
        '
        Me.txtToPO.Location = New System.Drawing.Point(35, 40)
        Me.txtToPO.Name = "txtToPO"
        Me.txtToPO.Size = New System.Drawing.Size(70, 23)
        Me.txtToPO.TabIndex = 3
        '
        'frmEmail
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(654, 279)
        Me.ControlBox = False
        Me.Controls.Add(Me.fraResults)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.txt)
        Me.Controls.Add(Me.prg)
        Me.Controls.Add(Me.txtFromName)
        Me.Controls.Add(Me.txtFromAddr)
        Me.Controls.Add(Me.lblFromName)
        Me.Controls.Add(Me.lblFromEmail)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmEmail"
        Me.Text = "Send Email"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.fraByDate.ResumeLayout(False)
        Me.fraByDate.PerformLayout()
        Me.fraResults.ResumeLayout(False)
        Me.fraByPoNo.ResumeLayout(False)
        Me.fraByPoNo.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents lblFromEmail As Label
    Friend WithEvents lblFromName As Label
    Friend WithEvents txtFromAddr As TextBox
    Friend WithEvents txtFromName As TextBox
    Friend WithEvents prg As ProgressBar
    Friend WithEvents txt As TextBox
    Friend WithEvents GroupBox1 As GroupBox
    Friend WithEvents cmdOK As Button
    Friend WithEvents cmdMail As Button
    Friend WithEvents chkReprint As CheckBox
    Friend WithEvents chkPrintPO As CheckBox
    Friend WithEvents fraByDate As GroupBox
    Friend WithEvents dtpToDate As DateTimePicker
    Friend WithEvents dtpFromDate As DateTimePicker
    Friend WithEvents lblToDate As Label
    Friend WithEvents lblFromDate As Label
    Friend WithEvents optByPoNo As RadioButton
    Friend WithEvents optByDate As RadioButton
    Friend WithEvents fraResults As GroupBox
    Friend WithEvents lblStatus As Label
    Friend WithEvents lstResults As CheckedListBox
    Friend WithEvents tmr As Timer
    Friend WithEvents fraByPoNo As GroupBox
    Friend WithEvents lblFromPO As Label
    Friend WithEvents lblToPo As Label
    Friend WithEvents txtToPO As TextBox
    Friend WithEvents txtFromPO As TextBox
End Class
