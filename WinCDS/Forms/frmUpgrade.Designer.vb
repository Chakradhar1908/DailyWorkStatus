<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmUpgrade
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
        Me.lstFiles = New System.Windows.Forms.ListBox()
        Me.prgCurrentFile = New System.Windows.Forms.ProgressBar()
        Me.prgComplete = New System.Windows.Forms.ProgressBar()
        Me.SuspendLayout()
        '
        'lstFiles
        '
        Me.lstFiles.FormattingEnabled = True
        Me.lstFiles.Location = New System.Drawing.Point(337, 176)
        Me.lstFiles.Name = "lstFiles"
        Me.lstFiles.Size = New System.Drawing.Size(120, 95)
        Me.lstFiles.TabIndex = 0
        '
        'prgCurrentFile
        '
        Me.prgCurrentFile.Location = New System.Drawing.Point(403, 325)
        Me.prgCurrentFile.Name = "prgCurrentFile"
        Me.prgCurrentFile.Size = New System.Drawing.Size(100, 23)
        Me.prgCurrentFile.TabIndex = 1
        '
        'prgComplete
        '
        Me.prgComplete.Location = New System.Drawing.Point(403, 377)
        Me.prgComplete.Name = "prgComplete"
        Me.prgComplete.Size = New System.Drawing.Size(100, 23)
        Me.prgComplete.TabIndex = 2
        '
        'frmUpgrade
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(800, 450)
        Me.Controls.Add(Me.prgComplete)
        Me.Controls.Add(Me.prgCurrentFile)
        Me.Controls.Add(Me.lstFiles)
        Me.Name = "frmUpgrade"
        Me.Text = "frmUpgrade"
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents lstFiles As ListBox
    Friend WithEvents prgCurrentFile As ProgressBar
    Friend WithEvents prgComplete As ProgressBar
End Class
