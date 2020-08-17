<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmSplash2
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
        Me.imgBackground = New System.Windows.Forms.PictureBox()
        Me.picProgress = New System.Windows.Forms.PictureBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.lblProgram3 = New System.Windows.Forms.Label()
        Me.lblProgram2 = New System.Windows.Forms.Label()
        Me.lblProgram0 = New System.Windows.Forms.Label()
        Me.lblProgram1 = New System.Windows.Forms.Label()
        Me.lblStatus = New System.Windows.Forms.Label()
        CType(Me.imgBackground, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.picProgress, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'imgBackground
        '
        Me.imgBackground.BackColor = System.Drawing.Color.Blue
        Me.imgBackground.Image = Global.WinCDS.My.Resources.Resources.Splash
        Me.imgBackground.Location = New System.Drawing.Point(0, -1)
        Me.imgBackground.Name = "imgBackground"
        Me.imgBackground.Size = New System.Drawing.Size(600, 300)
        Me.imgBackground.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize
        Me.imgBackground.TabIndex = 6
        Me.imgBackground.TabStop = False
        '
        'picProgress
        '
        Me.picProgress.BackColor = System.Drawing.Color.Blue
        Me.picProgress.Location = New System.Drawing.Point(500, 282)
        Me.picProgress.Name = "picProgress"
        Me.picProgress.Size = New System.Drawing.Size(100, 17)
        Me.picProgress.TabIndex = 1
        Me.picProgress.TabStop = False
        Me.picProgress.Visible = False
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.BackColor = System.Drawing.Color.Blue
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 18.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.Label1.Location = New System.Drawing.Point(157, 14)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(210, 29)
        Me.Label1.TabIndex = 7
        Me.Label1.Text = "DEMO VERSION"
        Me.Label1.Visible = False
        '
        'lblProgram3
        '
        Me.lblProgram3.AutoSize = True
        Me.lblProgram3.BackColor = System.Drawing.Color.Blue
        Me.lblProgram3.ForeColor = System.Drawing.Color.Transparent
        Me.lblProgram3.Location = New System.Drawing.Point(1, 218)
        Me.lblProgram3.Name = "lblProgram3"
        Me.lblProgram3.Size = New System.Drawing.Size(85, 13)
        Me.lblProgram3.TabIndex = 8
        Me.lblProgram3.Text = "IsServer? ###..."
        '
        'lblProgram2
        '
        Me.lblProgram2.AutoSize = True
        Me.lblProgram2.BackColor = System.Drawing.Color.Blue
        Me.lblProgram2.ForeColor = System.Drawing.Color.Transparent
        Me.lblProgram2.Location = New System.Drawing.Point(1, 235)
        Me.lblProgram2.Name = "lblProgram2"
        Me.lblProgram2.Size = New System.Drawing.Size(75, 13)
        Me.lblProgram2.TabIndex = 9
        Me.lblProgram2.Text = "Version ###..."
        '
        'lblProgram0
        '
        Me.lblProgram0.AutoSize = True
        Me.lblProgram0.BackColor = System.Drawing.Color.Blue
        Me.lblProgram0.ForeColor = System.Drawing.Color.Transparent
        Me.lblProgram0.Location = New System.Drawing.Point(1, 255)
        Me.lblProgram0.Name = "lblProgram0"
        Me.lblProgram0.Size = New System.Drawing.Size(78, 13)
        Me.lblProgram0.TabIndex = 10
        Me.lblProgram0.Text = "Loading ###..."
        '
        'lblProgram1
        '
        Me.lblProgram1.AutoSize = True
        Me.lblProgram1.BackColor = System.Drawing.Color.Blue
        Me.lblProgram1.ForeColor = System.Drawing.Color.Transparent
        Me.lblProgram1.Location = New System.Drawing.Point(1, 277)
        Me.lblProgram1.Name = "lblProgram1"
        Me.lblProgram1.Size = New System.Drawing.Size(84, 13)
        Me.lblProgram1.TabIndex = 11
        Me.lblProgram1.Text = "Copyright ###..."
        '
        'lblStatus
        '
        Me.lblStatus.AutoSize = True
        Me.lblStatus.BackColor = System.Drawing.Color.Blue
        Me.lblStatus.ForeColor = System.Drawing.Color.Transparent
        Me.lblStatus.Location = New System.Drawing.Point(404, 72)
        Me.lblStatus.Name = "lblStatus"
        Me.lblStatus.Size = New System.Drawing.Size(28, 13)
        Me.lblStatus.TabIndex = 12
        Me.lblStatus.Text = "###"
        '
        'frmSplash2
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(597, 298)
        Me.Controls.Add(Me.lblStatus)
        Me.Controls.Add(Me.lblProgram1)
        Me.Controls.Add(Me.lblProgram0)
        Me.Controls.Add(Me.lblProgram3)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.lblProgram2)
        Me.Controls.Add(Me.imgBackground)
        Me.Controls.Add(Me.picProgress)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Name = "frmSplash2"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "frmSplash2"
        CType(Me.imgBackground, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.picProgress, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents picProgress As PictureBox
    Friend WithEvents imgBackground As PictureBox
    Friend WithEvents Label1 As Label
    Friend WithEvents lblProgram3 As Label
    Friend WithEvents lblProgram2 As Label
    Friend WithEvents lblProgram0 As Label
    Friend WithEvents lblProgram1 As Label
    Friend WithEvents lblStatus As Label
End Class
