<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class MailCheck
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
        Me.lblInput = New System.Windows.Forms.Label()
        Me.InputBox = New System.Windows.Forms.TextBox()
        Me.fraInputType = New System.Windows.Forms.GroupBox()
        Me.cmdCancel = New System.Windows.Forms.Button()
        Me.cmdOK = New System.Windows.Forms.Button()
        Me.optServiceCall = New System.Windows.Forms.RadioButton()
        Me.optName = New System.Windows.Forms.RadioButton()
        Me.optSaleNo = New System.Windows.Forms.RadioButton()
        Me.optTelephone = New System.Windows.Forms.RadioButton()
        Me.lstMatches = New System.Windows.Forms.ListBox()
        Me.lblMatches = New System.Windows.Forms.Label()
        Me.fraInputType.SuspendLayout()
        Me.SuspendLayout()
        '
        'lblInput
        '
        Me.lblInput.AutoSize = True
        Me.lblInput.Location = New System.Drawing.Point(12, 4)
        Me.lblInput.Name = "lblInput"
        Me.lblInput.Size = New System.Drawing.Size(41, 13)
        Me.lblInput.TabIndex = 0
        Me.lblInput.Text = "lblInput"
        '
        'InputBox
        '
        Me.InputBox.Location = New System.Drawing.Point(12, 20)
        Me.InputBox.Name = "InputBox"
        Me.InputBox.Size = New System.Drawing.Size(219, 20)
        Me.InputBox.TabIndex = 1
        '
        'fraInputType
        '
        Me.fraInputType.Controls.Add(Me.cmdCancel)
        Me.fraInputType.Controls.Add(Me.cmdOK)
        Me.fraInputType.Controls.Add(Me.optServiceCall)
        Me.fraInputType.Controls.Add(Me.optName)
        Me.fraInputType.Controls.Add(Me.optSaleNo)
        Me.fraInputType.Controls.Add(Me.optTelephone)
        Me.fraInputType.Location = New System.Drawing.Point(12, 46)
        Me.fraInputType.Name = "fraInputType"
        Me.fraInputType.Size = New System.Drawing.Size(219, 105)
        Me.fraInputType.TabIndex = 2
        Me.fraInputType.TabStop = False
        '
        'cmdCancel
        '
        Me.cmdCancel.Location = New System.Drawing.Point(159, 42)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.Size = New System.Drawing.Size(50, 52)
        Me.cmdCancel.TabIndex = 5
        Me.cmdCancel.Text = "&Cancel"
        Me.cmdCancel.UseVisualStyleBackColor = True
        '
        'cmdOK
        '
        Me.cmdOK.Location = New System.Drawing.Point(103, 42)
        Me.cmdOK.Name = "cmdOK"
        Me.cmdOK.Size = New System.Drawing.Size(50, 52)
        Me.cmdOK.TabIndex = 4
        Me.cmdOK.Text = "&OK"
        Me.cmdOK.UseVisualStyleBackColor = True
        '
        'optServiceCall
        '
        Me.optServiceCall.AutoSize = True
        Me.optServiceCall.Location = New System.Drawing.Point(103, 17)
        Me.optServiceCall.Name = "optServiceCall"
        Me.optServiceCall.Size = New System.Drawing.Size(91, 17)
        Me.optServiceCall.TabIndex = 3
        Me.optServiceCall.TabStop = True
        Me.optServiceCall.Text = "Service C&all #"
        Me.optServiceCall.UseVisualStyleBackColor = True
        '
        'optName
        '
        Me.optName.AutoSize = True
        Me.optName.Location = New System.Drawing.Point(6, 66)
        Me.optName.Name = "optName"
        Me.optName.Size = New System.Drawing.Size(53, 17)
        Me.optName.TabIndex = 2
        Me.optName.TabStop = True
        Me.optName.Text = "&Name"
        Me.optName.UseVisualStyleBackColor = True
        '
        'optSaleNo
        '
        Me.optSaleNo.AutoSize = True
        Me.optSaleNo.Location = New System.Drawing.Point(6, 42)
        Me.optSaleNo.Name = "optSaleNo"
        Me.optSaleNo.Size = New System.Drawing.Size(63, 17)
        Me.optSaleNo.TabIndex = 1
        Me.optSaleNo.TabStop = True
        Me.optSaleNo.Text = "&Sale No"
        Me.optSaleNo.UseVisualStyleBackColor = True
        '
        'optTelephone
        '
        Me.optTelephone.AutoSize = True
        Me.optTelephone.Location = New System.Drawing.Point(6, 19)
        Me.optTelephone.Name = "optTelephone"
        Me.optTelephone.Size = New System.Drawing.Size(76, 17)
        Me.optTelephone.TabIndex = 0
        Me.optTelephone.TabStop = True
        Me.optTelephone.Text = "&Telephone"
        Me.optTelephone.UseVisualStyleBackColor = True
        '
        'lstMatches
        '
        Me.lstMatches.FormattingEnabled = True
        Me.lstMatches.Location = New System.Drawing.Point(258, 20)
        Me.lstMatches.Name = "lstMatches"
        Me.lstMatches.Size = New System.Drawing.Size(266, 134)
        Me.lstMatches.TabIndex = 3
        '
        'lblMatches
        '
        Me.lblMatches.AutoSize = True
        Me.lblMatches.Location = New System.Drawing.Point(258, 4)
        Me.lblMatches.Name = "lblMatches"
        Me.lblMatches.Size = New System.Drawing.Size(58, 13)
        Me.lblMatches.TabIndex = 4
        Me.lblMatches.Text = "lblMatches"
        '
        'MailCheck
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.cmdCancel
        Me.ClientSize = New System.Drawing.Size(529, 159)
        Me.Controls.Add(Me.lblMatches)
        Me.Controls.Add(Me.lstMatches)
        Me.Controls.Add(Me.fraInputType)
        Me.Controls.Add(Me.InputBox)
        Me.Controls.Add(Me.lblInput)
        Me.KeyPreview = True
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "MailCheck"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.fraInputType.ResumeLayout(False)
        Me.fraInputType.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents lblInput As Label
    Friend WithEvents InputBox As TextBox
    Friend WithEvents fraInputType As GroupBox
    Friend WithEvents cmdCancel As Button
    Friend WithEvents cmdOK As Button
    Friend WithEvents optServiceCall As RadioButton
    Friend WithEvents optName As RadioButton
    Friend WithEvents optSaleNo As RadioButton
    Friend WithEvents optTelephone As RadioButton
    Friend WithEvents lstMatches As ListBox
    Friend WithEvents lblMatches As Label
End Class
