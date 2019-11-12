<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class frmProgress
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
        Me.lblCaption = New System.Windows.Forms.Label()
        Me.prg = New WinCDS.ucPBar()
        Me.SuspendLayout()
        '
        'lblCaption
        '
        Me.lblCaption.AutoSize = True
        Me.lblCaption.Location = New System.Drawing.Point(0, 0)
        Me.lblCaption.Name = "lblCaption"
        Me.lblCaption.Size = New System.Drawing.Size(73, 13)
        Me.lblCaption.TabIndex = 0
        Me.lblCaption.Text = "Please Wait..."
        '
        'prg
        '
        Me.prg.BackColorNew = System.Drawing.SystemColors.Control
        Me.prg.BorderStyle = 0
        Me.prg.FontName = "Microsoft Sans Serif"
        Me.prg.ForeColorNew = System.Drawing.SystemColors.ControlText
        Me.prg.HasCaption = False
        Me.prg.Location = New System.Drawing.Point(12, 16)
        Me.prg.Max = 0
        Me.prg.Min = 0
        Me.prg.Name = "prg"
        Me.prg.ShowDuration = False
        Me.prg.ShowRemaining = False
        Me.prg.Size = New System.Drawing.Size(150, 31)
        Me.prg.Style = 0
        Me.prg.TabIndex = 1
        Me.prg.Value = 0
        '
        'frmProgress
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(800, 450)
        Me.Controls.Add(Me.prg)
        Me.Controls.Add(Me.lblCaption)
        Me.Name = "frmProgress"
        Me.Text = "frmProgress"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents lblCaption As Label
    Friend WithEvents prg As ucPBar
End Class
