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
        Me.lblStatus = New System.Windows.Forms.Label()
        Me.picProgress = New System.Windows.Forms.PictureBox()
        Me.lblProgram0 = New System.Windows.Forms.Label()
        Me.lblProgram1 = New System.Windows.Forms.Label()
        Me.lblProgram2 = New System.Windows.Forms.Label()
        Me.lblProgram3 = New System.Windows.Forms.Label()
        Me.imgBackground = New System.Windows.Forms.PictureBox()
        CType(Me.picProgress, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.imgBackground, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'lblStatus
        '
        Me.lblStatus.AutoSize = True
        Me.lblStatus.Location = New System.Drawing.Point(317, 174)
        Me.lblStatus.Name = "lblStatus"
        Me.lblStatus.Size = New System.Drawing.Size(39, 13)
        Me.lblStatus.TabIndex = 0
        Me.lblStatus.Text = "Label1"
        '
        'picProgress
        '
        Me.picProgress.Location = New System.Drawing.Point(335, 324)
        Me.picProgress.Name = "picProgress"
        Me.picProgress.Size = New System.Drawing.Size(100, 50)
        Me.picProgress.TabIndex = 1
        Me.picProgress.TabStop = False
        '
        'lblProgram0
        '
        Me.lblProgram0.AutoSize = True
        Me.lblProgram0.Location = New System.Drawing.Point(381, 219)
        Me.lblProgram0.Name = "lblProgram0"
        Me.lblProgram0.Size = New System.Drawing.Size(39, 13)
        Me.lblProgram0.TabIndex = 2
        Me.lblProgram0.Text = "Label1"
        '
        'lblProgram1
        '
        Me.lblProgram1.AutoSize = True
        Me.lblProgram1.Location = New System.Drawing.Point(381, 250)
        Me.lblProgram1.Name = "lblProgram1"
        Me.lblProgram1.Size = New System.Drawing.Size(39, 13)
        Me.lblProgram1.TabIndex = 3
        Me.lblProgram1.Text = "Label1"
        '
        'lblProgram2
        '
        Me.lblProgram2.AutoSize = True
        Me.lblProgram2.Location = New System.Drawing.Point(381, 275)
        Me.lblProgram2.Name = "lblProgram2"
        Me.lblProgram2.Size = New System.Drawing.Size(39, 13)
        Me.lblProgram2.TabIndex = 4
        Me.lblProgram2.Text = "Label1"
        '
        'lblProgram3
        '
        Me.lblProgram3.AutoSize = True
        Me.lblProgram3.Location = New System.Drawing.Point(381, 288)
        Me.lblProgram3.Name = "lblProgram3"
        Me.lblProgram3.Size = New System.Drawing.Size(39, 13)
        Me.lblProgram3.TabIndex = 5
        Me.lblProgram3.Text = "Label1"
        '
        'imgBackground
        '
        Me.imgBackground.Location = New System.Drawing.Point(512, 324)
        Me.imgBackground.Name = "imgBackground"
        Me.imgBackground.Size = New System.Drawing.Size(100, 50)
        Me.imgBackground.TabIndex = 6
        Me.imgBackground.TabStop = False
        '
        'frmSplash2
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(800, 450)
        Me.Controls.Add(Me.imgBackground)
        Me.Controls.Add(Me.lblProgram3)
        Me.Controls.Add(Me.lblProgram2)
        Me.Controls.Add(Me.lblProgram1)
        Me.Controls.Add(Me.lblProgram0)
        Me.Controls.Add(Me.picProgress)
        Me.Controls.Add(Me.lblStatus)
        Me.Name = "frmSplash2"
        Me.Text = "frmSplash2"
        CType(Me.picProgress, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.imgBackground, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents lblStatus As Label
    Friend WithEvents picProgress As PictureBox
    Friend WithEvents lblProgram0 As Label
    Friend WithEvents lblProgram1 As Label
    Friend WithEvents lblProgram2 As Label
    Friend WithEvents lblProgram3 As Label
    Friend WithEvents imgBackground As PictureBox
End Class
