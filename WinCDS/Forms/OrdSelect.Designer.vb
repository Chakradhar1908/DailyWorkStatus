<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class OrdSelect
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
        Me.fraOpt = New System.Windows.Forms.GroupBox()
        Me.txtPayMemo = New System.Windows.Forms.TextBox()
        Me.lblPayMemo = New System.Windows.Forms.Label()
        Me.chkPayAll = New System.Windows.Forms.CheckBox()
        Me.optPayment = New System.Windows.Forms.RadioButton()
        Me.optNoTax2 = New System.Windows.Forms.RadioButton()
        Me.optTax2 = New System.Windows.Forms.RadioButton()
        Me.optNoTax = New System.Windows.Forms.RadioButton()
        Me.optTax1 = New System.Windows.Forms.RadioButton()
        Me.optCarpet = New System.Windows.Forms.RadioButton()
        Me.optStoreCredit = New System.Windows.Forms.RadioButton()
        Me.optNotes = New System.Windows.Forms.RadioButton()
        Me.optLabor = New System.Windows.Forms.RadioButton()
        Me.optDelivery = New System.Windows.Forms.RadioButton()
        Me.optStain = New System.Windows.Forms.RadioButton()
        Me.optEnterStyle = New System.Windows.Forms.RadioButton()
        Me.cmdOk = New System.Windows.Forms.Button()
        Me.cmdCancel = New System.Windows.Forms.Button()
        Me.cmdProcessSale = New System.Windows.Forms.Button()
        Me.lstOptions = New System.Windows.Forms.ListBox()
        Me.fraOpt.SuspendLayout()
        Me.SuspendLayout()
        '
        'fraOpt
        '
        Me.fraOpt.Controls.Add(Me.txtPayMemo)
        Me.fraOpt.Controls.Add(Me.lblPayMemo)
        Me.fraOpt.Controls.Add(Me.chkPayAll)
        Me.fraOpt.Controls.Add(Me.optPayment)
        Me.fraOpt.Controls.Add(Me.optNoTax2)
        Me.fraOpt.Controls.Add(Me.optTax2)
        Me.fraOpt.Controls.Add(Me.optNoTax)
        Me.fraOpt.Controls.Add(Me.optTax1)
        Me.fraOpt.Controls.Add(Me.optCarpet)
        Me.fraOpt.Controls.Add(Me.optStoreCredit)
        Me.fraOpt.Controls.Add(Me.optNotes)
        Me.fraOpt.Controls.Add(Me.optLabor)
        Me.fraOpt.Controls.Add(Me.optDelivery)
        Me.fraOpt.Controls.Add(Me.optStain)
        Me.fraOpt.Controls.Add(Me.optEnterStyle)
        Me.fraOpt.Location = New System.Drawing.Point(12, -1)
        Me.fraOpt.Name = "fraOpt"
        Me.fraOpt.Size = New System.Drawing.Size(248, 229)
        Me.fraOpt.TabIndex = 0
        Me.fraOpt.TabStop = False
        '
        'txtPayMemo
        '
        Me.txtPayMemo.Location = New System.Drawing.Point(81, 201)
        Me.txtPayMemo.Name = "txtPayMemo"
        Me.txtPayMemo.Size = New System.Drawing.Size(161, 20)
        Me.txtPayMemo.TabIndex = 14
        '
        'lblPayMemo
        '
        Me.lblPayMemo.AutoSize = True
        Me.lblPayMemo.Location = New System.Drawing.Point(12, 201)
        Me.lblPayMemo.Name = "lblPayMemo"
        Me.lblPayMemo.Size = New System.Drawing.Size(63, 13)
        Me.lblPayMemo.TabIndex = 13
        Me.lblPayMemo.Text = " Pa&y Memo:"
        '
        'chkPayAll
        '
        Me.chkPayAll.AutoSize = True
        Me.chkPayAll.Location = New System.Drawing.Point(158, 170)
        Me.chkPayAll.Name = "chkPayAll"
        Me.chkPayAll.Size = New System.Drawing.Size(58, 17)
        Me.chkPayAll.TabIndex = 12
        Me.chkPayAll.Text = "Pay &All"
        Me.chkPayAll.UseVisualStyleBackColor = True
        '
        'optPayment
        '
        Me.optPayment.AutoSize = True
        Me.optPayment.Location = New System.Drawing.Point(158, 145)
        Me.optPayment.Name = "optPayment"
        Me.optPayment.Size = New System.Drawing.Size(66, 17)
        Me.optPayment.TabIndex = 11
        Me.optPayment.TabStop = True
        Me.optPayment.Text = "&Payment"
        Me.optPayment.UseVisualStyleBackColor = True
        '
        'optNoTax2
        '
        Me.optNoTax2.AutoSize = True
        Me.optNoTax2.Location = New System.Drawing.Point(158, 95)
        Me.optNoTax2.Name = "optNoTax2"
        Me.optNoTax2.Size = New System.Drawing.Size(90, 17)
        Me.optNoTax2.TabIndex = 10
        Me.optNoTax2.TabStop = True
        Me.optNoTax2.Text = "- Variable Ta&x"
        Me.optNoTax2.UseVisualStyleBackColor = True
        '
        'optTax2
        '
        Me.optTax2.AutoSize = True
        Me.optTax2.Location = New System.Drawing.Point(158, 70)
        Me.optTax2.Name = "optTax2"
        Me.optTax2.Size = New System.Drawing.Size(84, 17)
        Me.optTax2.TabIndex = 9
        Me.optTax2.TabStop = True
        Me.optTax2.Text = "&Variable Tax"
        Me.optTax2.UseVisualStyleBackColor = True
        '
        'optNoTax
        '
        Me.optNoTax.AutoSize = True
        Me.optNoTax.Location = New System.Drawing.Point(158, 45)
        Me.optNoTax.Name = "optNoTax"
        Me.optNoTax.Size = New System.Drawing.Size(58, 17)
        Me.optNoTax.TabIndex = 8
        Me.optNoTax.TabStop = True
        Me.optNoTax.Text = "- T&ax 1"
        Me.optNoTax.UseVisualStyleBackColor = True
        '
        'optTax1
        '
        Me.optTax1.AutoSize = True
        Me.optTax1.Location = New System.Drawing.Point(158, 20)
        Me.optTax1.Name = "optTax1"
        Me.optTax1.Size = New System.Drawing.Size(52, 17)
        Me.optTax1.TabIndex = 7
        Me.optTax1.TabStop = True
        Me.optTax1.Text = "&Tax 1"
        Me.optTax1.UseVisualStyleBackColor = True
        '
        'optCarpet
        '
        Me.optCarpet.AutoSize = True
        Me.optCarpet.Location = New System.Drawing.Point(15, 170)
        Me.optCarpet.Name = "optCarpet"
        Me.optCarpet.Size = New System.Drawing.Size(56, 17)
        Me.optCarpet.TabIndex = 6
        Me.optCarpet.TabStop = True
        Me.optCarpet.Text = "Ca&rpet"
        Me.optCarpet.UseVisualStyleBackColor = True
        '
        'optStoreCredit
        '
        Me.optStoreCredit.AutoSize = True
        Me.optStoreCredit.Location = New System.Drawing.Point(15, 145)
        Me.optStoreCredit.Name = "optStoreCredit"
        Me.optStoreCredit.Size = New System.Drawing.Size(114, 17)
        Me.optStoreCredit.TabIndex = 5
        Me.optStoreCredit.TabStop = True
        Me.optStoreCredit.Text = "Purchase Gift &Card"
        Me.optStoreCredit.UseVisualStyleBackColor = True
        '
        'optNotes
        '
        Me.optNotes.AutoSize = True
        Me.optNotes.Location = New System.Drawing.Point(15, 120)
        Me.optNotes.Name = "optNotes"
        Me.optNotes.Size = New System.Drawing.Size(53, 17)
        Me.optNotes.TabIndex = 4
        Me.optNotes.TabStop = True
        Me.optNotes.Text = "&Notes"
        Me.optNotes.UseVisualStyleBackColor = True
        '
        'optLabor
        '
        Me.optLabor.AutoSize = True
        Me.optLabor.Location = New System.Drawing.Point(15, 95)
        Me.optLabor.Name = "optLabor"
        Me.optLabor.Size = New System.Drawing.Size(92, 17)
        Me.optLabor.TabIndex = 3
        Me.optLabor.TabStop = True
        Me.optLabor.Text = "&Labor  Charge"
        Me.optLabor.UseVisualStyleBackColor = True
        '
        'optDelivery
        '
        Me.optDelivery.AutoSize = True
        Me.optDelivery.Location = New System.Drawing.Point(15, 70)
        Me.optDelivery.Name = "optDelivery"
        Me.optDelivery.Size = New System.Drawing.Size(100, 17)
        Me.optDelivery.TabIndex = 2
        Me.optDelivery.TabStop = True
        Me.optDelivery.Text = "&Delivery Charge"
        Me.optDelivery.UseVisualStyleBackColor = True
        '
        'optStain
        '
        Me.optStain.AutoSize = True
        Me.optStain.Location = New System.Drawing.Point(15, 45)
        Me.optStain.Name = "optStain"
        Me.optStain.Size = New System.Drawing.Size(100, 17)
        Me.optStain.TabIndex = 1
        Me.optStain.TabStop = True
        Me.optStain.Text = "&Stain Protection"
        Me.optStain.UseVisualStyleBackColor = True
        '
        'optEnterStyle
        '
        Me.optEnterStyle.AutoSize = True
        Me.optEnterStyle.Location = New System.Drawing.Point(15, 20)
        Me.optEnterStyle.Name = "optEnterStyle"
        Me.optEnterStyle.Size = New System.Drawing.Size(116, 17)
        Me.optEnterStyle.TabIndex = 0
        Me.optEnterStyle.TabStop = True
        Me.optEnterStyle.Text = "&Enter Style Number"
        Me.optEnterStyle.UseVisualStyleBackColor = True
        '
        'cmdOk
        '
        Me.cmdOk.AutoSize = True
        Me.cmdOk.Location = New System.Drawing.Point(12, 235)
        Me.cmdOk.Name = "cmdOk"
        Me.cmdOk.Size = New System.Drawing.Size(54, 34)
        Me.cmdOk.TabIndex = 1
        Me.cmdOk.Text = "&OK"
        Me.cmdOk.UseVisualStyleBackColor = True
        '
        'cmdCancel
        '
        Me.cmdCancel.Location = New System.Drawing.Point(93, 235)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.Size = New System.Drawing.Size(75, 23)
        Me.cmdCancel.TabIndex = 2
        Me.cmdCancel.Text = "Cance&l"
        Me.cmdCancel.UseVisualStyleBackColor = True
        '
        'cmdProcessSale
        '
        Me.cmdProcessSale.Location = New System.Drawing.Point(170, 235)
        Me.cmdProcessSale.Name = "cmdProcessSale"
        Me.cmdProcessSale.Size = New System.Drawing.Size(90, 23)
        Me.cmdProcessSale.TabIndex = 3
        Me.cmdProcessSale.Text = "Process Sale"
        Me.cmdProcessSale.UseVisualStyleBackColor = True
        '
        'lstOptions
        '
        Me.lstOptions.FormattingEnabled = True
        Me.lstOptions.Location = New System.Drawing.Point(275, 11)
        Me.lstOptions.Name = "lstOptions"
        Me.lstOptions.Size = New System.Drawing.Size(94, 212)
        Me.lstOptions.TabIndex = 4
        '
        'OrdSelect
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(378, 269)
        Me.ControlBox = False
        Me.Controls.Add(Me.cmdOk)
        Me.Controls.Add(Me.lstOptions)
        Me.Controls.Add(Me.cmdProcessSale)
        Me.Controls.Add(Me.cmdCancel)
        Me.Controls.Add(Me.fraOpt)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "OrdSelect"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Click On Selection"
        Me.fraOpt.ResumeLayout(False)
        Me.fraOpt.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents fraOpt As GroupBox
    Friend WithEvents txtPayMemo As TextBox
    Friend WithEvents lblPayMemo As Label
    Friend WithEvents chkPayAll As CheckBox
    Friend WithEvents optPayment As RadioButton
    Friend WithEvents optNoTax2 As RadioButton
    Friend WithEvents optTax2 As RadioButton
    Friend WithEvents optNoTax As RadioButton
    Friend WithEvents optTax1 As RadioButton
    Friend WithEvents optCarpet As RadioButton
    Friend WithEvents optStoreCredit As RadioButton
    Friend WithEvents optNotes As RadioButton
    Friend WithEvents optLabor As RadioButton
    Friend WithEvents optDelivery As RadioButton
    Friend WithEvents optStain As RadioButton
    Friend WithEvents optEnterStyle As RadioButton
    Friend WithEvents cmdOk As Button
    Friend WithEvents cmdCancel As Button
    Friend WithEvents cmdProcessSale As Button
    Friend WithEvents lstOptions As ListBox
End Class
