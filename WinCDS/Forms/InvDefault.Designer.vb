<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class InvDefault
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
        Me.fraOptions = New System.Windows.Forms.GroupBox()
        Me.optSOLawNotCarried = New System.Windows.Forms.RadioButton()
        Me.optSONotCarried = New System.Windows.Forms.RadioButton()
        Me.optEnterNotInInv = New System.Windows.Forms.RadioButton()
        Me.optReEnter = New System.Windows.Forms.RadioButton()
        Me.lstResults = New System.Windows.Forms.ListBox()
        Me.cmdApply = New System.Windows.Forms.Button()
        Me.lblStylenotindatabase = New System.Windows.Forms.Label()
        Me.fraOptions.SuspendLayout()
        Me.SuspendLayout()
        '
        'fraOptions
        '
        Me.fraOptions.Controls.Add(Me.optSOLawNotCarried)
        Me.fraOptions.Controls.Add(Me.optSONotCarried)
        Me.fraOptions.Controls.Add(Me.optEnterNotInInv)
        Me.fraOptions.Controls.Add(Me.optReEnter)
        Me.fraOptions.Location = New System.Drawing.Point(5, 3)
        Me.fraOptions.Name = "fraOptions"
        Me.fraOptions.Size = New System.Drawing.Size(160, 121)
        Me.fraOptions.TabIndex = 0
        Me.fraOptions.TabStop = False
        '
        'optSOLawNotCarried
        '
        Me.optSOLawNotCarried.AutoSize = True
        Me.optSOLawNotCarried.Location = New System.Drawing.Point(6, 88)
        Me.optSOLawNotCarried.Name = "optSOLawNotCarried"
        Me.optSOLawNotCarried.Size = New System.Drawing.Size(154, 17)
        Me.optSOLawNotCarried.TabIndex = 3
        Me.optSOLawNotCarried.TabStop = True
        Me.optSOLawNotCarried.Text = "S/O  &La-A-Way Not Carried"
        Me.optSOLawNotCarried.UseVisualStyleBackColor = True
        '
        'optSONotCarried
        '
        Me.optSONotCarried.AutoSize = True
        Me.optSONotCarried.Location = New System.Drawing.Point(6, 65)
        Me.optSONotCarried.Name = "optSONotCarried"
        Me.optSONotCarried.Size = New System.Drawing.Size(127, 17)
        Me.optSONotCarried.TabIndex = 2
        Me.optSONotCarried.TabStop = True
        Me.optSONotCarried.Text = "&S/O  Item Not Carried"
        Me.optSONotCarried.UseVisualStyleBackColor = True
        '
        'optEnterNotInInv
        '
        Me.optEnterNotInInv.AutoSize = True
        Me.optEnterNotInInv.Location = New System.Drawing.Point(6, 42)
        Me.optEnterNotInInv.Name = "optEnterNotInInv"
        Me.optEnterNotInInv.Size = New System.Drawing.Size(152, 17)
        Me.optEnterNotInInv.TabIndex = 1
        Me.optEnterNotInInv.TabStop = True
        Me.optEnterNotInInv.Text = "&Enter Item Not In Inventory"
        Me.optEnterNotInInv.UseVisualStyleBackColor = True
        '
        'optReEnter
        '
        Me.optReEnter.AutoSize = True
        Me.optReEnter.Location = New System.Drawing.Point(6, 19)
        Me.optReEnter.Name = "optReEnter"
        Me.optReEnter.Size = New System.Drawing.Size(133, 17)
        Me.optReEnter.TabIndex = 0
        Me.optReEnter.TabStop = True
        Me.optReEnter.Text = "&Re Enter Style Number"
        Me.optReEnter.UseVisualStyleBackColor = True
        '
        'lstResults
        '
        Me.lstResults.FormattingEnabled = True
        Me.lstResults.Items.AddRange(New Object() {"1", "2", "3", "4"})
        Me.lstResults.Location = New System.Drawing.Point(171, 3)
        Me.lstResults.Name = "lstResults"
        Me.lstResults.Size = New System.Drawing.Size(120, 121)
        Me.lstResults.TabIndex = 1
        '
        'cmdApply
        '
        Me.cmdApply.Location = New System.Drawing.Point(111, 130)
        Me.cmdApply.Name = "cmdApply"
        Me.cmdApply.Size = New System.Drawing.Size(75, 56)
        Me.cmdApply.TabIndex = 2
        Me.cmdApply.Text = "&OK"
        Me.cmdApply.UseVisualStyleBackColor = True
        '
        'lblStylenotindatabase
        '
        Me.lblStylenotindatabase.AutoSize = True
        Me.lblStylenotindatabase.Location = New System.Drawing.Point(31, 1)
        Me.lblStylenotindatabase.Name = "lblStylenotindatabase"
        Me.lblStylenotindatabase.Size = New System.Drawing.Size(109, 13)
        Me.lblStylenotindatabase.TabIndex = 3
        Me.lblStylenotindatabase.Text = "Style not in database."
        Me.lblStylenotindatabase.Visible = False
        '
        'InvDefault
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(297, 191)
        Me.Controls.Add(Me.lblStylenotindatabase)
        Me.Controls.Add(Me.cmdApply)
        Me.Controls.Add(Me.lstResults)
        Me.Controls.Add(Me.fraOptions)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.MaximizeBox = False
        Me.Name = "InvDefault"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Style Not In Data Base"
        Me.fraOptions.ResumeLayout(False)
        Me.fraOptions.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents fraOptions As GroupBox
    Friend WithEvents optSOLawNotCarried As RadioButton
    Friend WithEvents optSONotCarried As RadioButton
    Friend WithEvents optEnterNotInInv As RadioButton
    Friend WithEvents optReEnter As RadioButton
    Friend WithEvents lstResults As ListBox
    Friend WithEvents cmdApply As Button
    Friend WithEvents lblStylenotindatabase As Label
End Class
