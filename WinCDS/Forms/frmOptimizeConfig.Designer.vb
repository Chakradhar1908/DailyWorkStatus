<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmOptimizeConfig
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
        Me.fraConfig = New System.Windows.Forms.GroupBox()
        Me.cmdOK = New System.Windows.Forms.Button()
        Me.txtTrucks = New System.Windows.Forms.TextBox()
        Me.txtCostPerHour = New System.Windows.Forms.TextBox()
        Me.txtCostPerMile = New System.Windows.Forms.TextBox()
        Me.txtTimePerStop = New System.Windows.Forms.TextBox()
        Me.txtStartTime = New System.Windows.Forms.TextBox()
        Me.lblTrucks = New System.Windows.Forms.Label()
        Me.lblCostPerHour = New System.Windows.Forms.Label()
        Me.lblCostPerMile = New System.Windows.Forms.Label()
        Me.lblTimePerStop = New System.Windows.Forms.Label()
        Me.lblStartTime = New System.Windows.Forms.Label()
        Me.fraConfig.SuspendLayout()
        Me.SuspendLayout()
        '
        'fraConfig
        '
        Me.fraConfig.Controls.Add(Me.cmdOK)
        Me.fraConfig.Controls.Add(Me.txtTrucks)
        Me.fraConfig.Controls.Add(Me.txtCostPerHour)
        Me.fraConfig.Controls.Add(Me.txtCostPerMile)
        Me.fraConfig.Controls.Add(Me.txtTimePerStop)
        Me.fraConfig.Controls.Add(Me.txtStartTime)
        Me.fraConfig.Controls.Add(Me.lblTrucks)
        Me.fraConfig.Controls.Add(Me.lblCostPerHour)
        Me.fraConfig.Controls.Add(Me.lblCostPerMile)
        Me.fraConfig.Controls.Add(Me.lblTimePerStop)
        Me.fraConfig.Controls.Add(Me.lblStartTime)
        Me.fraConfig.Location = New System.Drawing.Point(10, 10)
        Me.fraConfig.Name = "fraConfig"
        Me.fraConfig.Size = New System.Drawing.Size(208, 198)
        Me.fraConfig.TabIndex = 0
        Me.fraConfig.TabStop = False
        Me.fraConfig.Text = "Op&tions:"
        '
        'cmdOK
        '
        Me.cmdOK.Location = New System.Drawing.Point(67, 136)
        Me.cmdOK.Name = "cmdOK"
        Me.cmdOK.Size = New System.Drawing.Size(75, 51)
        Me.cmdOK.TabIndex = 10
        Me.cmdOK.Text = "&OK"
        Me.cmdOK.UseVisualStyleBackColor = True
        '
        'txtTrucks
        '
        Me.txtTrucks.Location = New System.Drawing.Point(94, 107)
        Me.txtTrucks.Name = "txtTrucks"
        Me.txtTrucks.Size = New System.Drawing.Size(100, 20)
        Me.txtTrucks.TabIndex = 9
        Me.txtTrucks.Text = "1"
        '
        'txtCostPerHour
        '
        Me.txtCostPerHour.Location = New System.Drawing.Point(94, 87)
        Me.txtCostPerHour.Name = "txtCostPerHour"
        Me.txtCostPerHour.Size = New System.Drawing.Size(100, 20)
        Me.txtCostPerHour.TabIndex = 8
        Me.txtCostPerHour.Text = "11.00"
        '
        'txtCostPerMile
        '
        Me.txtCostPerMile.Location = New System.Drawing.Point(94, 67)
        Me.txtCostPerMile.Name = "txtCostPerMile"
        Me.txtCostPerMile.Size = New System.Drawing.Size(100, 20)
        Me.txtCostPerMile.TabIndex = 7
        Me.txtCostPerMile.Text = "0.45"
        '
        'txtTimePerStop
        '
        Me.txtTimePerStop.Location = New System.Drawing.Point(94, 47)
        Me.txtTimePerStop.Name = "txtTimePerStop"
        Me.txtTimePerStop.Size = New System.Drawing.Size(100, 20)
        Me.txtTimePerStop.TabIndex = 6
        Me.txtTimePerStop.Text = "10"
        '
        'txtStartTime
        '
        Me.txtStartTime.Location = New System.Drawing.Point(94, 27)
        Me.txtStartTime.Name = "txtStartTime"
        Me.txtStartTime.Size = New System.Drawing.Size(100, 20)
        Me.txtStartTime.TabIndex = 5
        Me.txtStartTime.Text = "7:00 AM"
        '
        'lblTrucks
        '
        Me.lblTrucks.AutoSize = True
        Me.lblTrucks.Location = New System.Drawing.Point(10, 114)
        Me.lblTrucks.Name = "lblTrucks"
        Me.lblTrucks.Size = New System.Drawing.Size(43, 13)
        Me.lblTrucks.TabIndex = 4
        Me.lblTrucks.Text = "Truc&ks:"
        '
        'lblCostPerHour
        '
        Me.lblCostPerHour.AutoSize = True
        Me.lblCostPerHour.Location = New System.Drawing.Point(10, 94)
        Me.lblCostPerHour.Name = "lblCostPerHour"
        Me.lblCostPerHour.Size = New System.Drawing.Size(76, 13)
        Me.lblCostPerHour.TabIndex = 3
        Me.lblCostPerHour.Text = "Cost Per &Hour:"
        '
        'lblCostPerMile
        '
        Me.lblCostPerMile.AutoSize = True
        Me.lblCostPerMile.Location = New System.Drawing.Point(10, 70)
        Me.lblCostPerMile.Name = "lblCostPerMile"
        Me.lblCostPerMile.Size = New System.Drawing.Size(72, 13)
        Me.lblCostPerMile.TabIndex = 2
        Me.lblCostPerMile.Text = "Cost Per Mi&le:"
        '
        'lblTimePerStop
        '
        Me.lblTimePerStop.AutoSize = True
        Me.lblTimePerStop.Location = New System.Drawing.Point(10, 47)
        Me.lblTimePerStop.Name = "lblTimePerStop"
        Me.lblTimePerStop.Size = New System.Drawing.Size(77, 13)
        Me.lblTimePerStop.TabIndex = 1
        Me.lblTimePerStop.Text = "T&ime Per Stop:"
        '
        'lblStartTime
        '
        Me.lblStartTime.AutoSize = True
        Me.lblStartTime.Location = New System.Drawing.Point(10, 27)
        Me.lblStartTime.Name = "lblStartTime"
        Me.lblStartTime.Size = New System.Drawing.Size(72, 13)
        Me.lblStartTime.TabIndex = 0
        Me.lblStartTime.Text = "&Starting Time:"
        '
        'frmOptimizeConfig
        '
        Me.AcceptButton = Me.cmdOK
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(227, 214)
        Me.Controls.Add(Me.fraConfig)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow
        Me.Name = "frmOptimizeConfig"
        Me.Text = "Optimization Configuration"
        Me.fraConfig.ResumeLayout(False)
        Me.fraConfig.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents fraConfig As GroupBox
    Friend WithEvents txtTrucks As TextBox
    Friend WithEvents txtCostPerHour As TextBox
    Friend WithEvents txtCostPerMile As TextBox
    Friend WithEvents txtTimePerStop As TextBox
    Friend WithEvents txtStartTime As TextBox
    Friend WithEvents lblTrucks As Label
    Friend WithEvents lblCostPerHour As Label
    Friend WithEvents lblCostPerMile As Label
    Friend WithEvents lblTimePerStop As Label
    Friend WithEvents lblStartTime As Label
    Friend WithEvents cmdOK As Button
End Class
