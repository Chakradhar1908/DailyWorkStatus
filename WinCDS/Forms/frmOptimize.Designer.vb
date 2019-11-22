<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmOptimize
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
        Me.picNetwork = New System.Windows.Forms.PictureBox()
        CType(Me.picNetwork, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'picNetwork
        '
        Me.picNetwork.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.picNetwork.Location = New System.Drawing.Point(0, 0)
        Me.picNetwork.Name = "picNetwork"
        Me.picNetwork.Size = New System.Drawing.Size(278, 219)
        Me.picNetwork.TabIndex = 0
        Me.picNetwork.TabStop = False
        '
        'frmOptimize
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(279, 226)
        Me.Controls.Add(Me.picNetwork)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.SizableToolWindow
        Me.Name = "frmOptimize"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "Route Optimzier"
        CType(Me.picNetwork, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents picNetwork As PictureBox
End Class
