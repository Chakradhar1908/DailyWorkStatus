<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmINet
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmINet))
        Me.inet = New AxInetCtlsObjects.AxInet()
        Me.cmdClose = New System.Windows.Forms.Button()
        CType(Me.inet, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'inet
        '
        Me.inet.Enabled = True
        Me.inet.Location = New System.Drawing.Point(12, 12)
        Me.inet.Name = "inet"
        Me.inet.OcxState = CType(resources.GetObject("inet.OcxState"), System.Windows.Forms.AxHost.State)
        Me.inet.Size = New System.Drawing.Size(38, 38)
        Me.inet.TabIndex = 0
        '
        'cmdClose
        '
        Me.cmdClose.Location = New System.Drawing.Point(46, 22)
        Me.cmdClose.Name = "cmdClose"
        Me.cmdClose.Size = New System.Drawing.Size(49, 23)
        Me.cmdClose.TabIndex = 1
        Me.cmdClose.Text = "C&lose"
        Me.cmdClose.UseVisualStyleBackColor = True
        '
        'frmINet
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(120, 57)
        Me.Controls.Add(Me.cmdClose)
        Me.Controls.Add(Me.inet)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmINet"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Communicating..."
        CType(Me.inet, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents inet As AxInetCtlsObjects.AxInet
    Friend WithEvents cmdClose As Button
End Class
