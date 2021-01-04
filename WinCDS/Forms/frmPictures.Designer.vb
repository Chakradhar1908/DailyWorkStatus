<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class frmPictures
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
        Me.lblRef = New System.Windows.Forms.Label()
        Me.lblPictures = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'lblRef
        '
        Me.lblRef.AutoSize = True
        Me.lblRef.Location = New System.Drawing.Point(80, 58)
        Me.lblRef.Name = "lblRef"
        Me.lblRef.Size = New System.Drawing.Size(39, 13)
        Me.lblRef.TabIndex = 0
        Me.lblRef.Text = "Label1"
        '
        'lblPictures
        '
        Me.lblPictures.AutoSize = True
        Me.lblPictures.Location = New System.Drawing.Point(381, 219)
        Me.lblPictures.Name = "lblPictures"
        Me.lblPictures.Size = New System.Drawing.Size(55, 13)
        Me.lblPictures.TabIndex = 1
        Me.lblPictures.Text = "lblPictures"
        '
        'frmPictures
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(800, 450)
        Me.Controls.Add(Me.lblPictures)
        Me.Controls.Add(Me.lblRef)
        Me.Name = "frmPictures"
        Me.Text = "frmPictures"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents lblRef As Label
    Friend WithEvents lblPictures As Label
End Class
