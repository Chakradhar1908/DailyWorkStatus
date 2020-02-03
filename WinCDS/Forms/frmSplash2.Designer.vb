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
        CType(Me.picProgress, System.ComponentModel.ISupportInitialize).BeginInit()
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
        Me.picProgress.Location = New System.Drawing.Point(345, 319)
        Me.picProgress.Name = "picProgress"
        Me.picProgress.Size = New System.Drawing.Size(100, 50)
        Me.picProgress.TabIndex = 1
        Me.picProgress.TabStop = False
        '
        'frmSplash2
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(800, 450)
        Me.Controls.Add(Me.picProgress)
        Me.Controls.Add(Me.lblStatus)
        Me.Name = "frmSplash2"
        Me.Text = "frmSplash2"
        CType(Me.picProgress, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents lblStatus As Label
    Friend WithEvents picProgress As PictureBox
End Class
