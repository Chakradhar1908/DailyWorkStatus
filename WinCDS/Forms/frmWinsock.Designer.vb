<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmWinsock
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmWinsock))
        Me.Sock = New AxMSWinsockLib.AxWinsock()
        CType(Me.Sock, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Sock
        '
        Me.Sock.Enabled = True
        Me.Sock.Location = New System.Drawing.Point(402, 215)
        Me.Sock.Name = "Sock"
        Me.Sock.OcxState = CType(resources.GetObject("Sock.OcxState"), System.Windows.Forms.AxHost.State)
        Me.Sock.Size = New System.Drawing.Size(28, 28)
        Me.Sock.TabIndex = 0
        '
        'frmWinsock
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(800, 450)
        Me.Controls.Add(Me.Sock)
        Me.Name = "frmWinsock"
        Me.Text = "frmWinsock"
        CType(Me.Sock, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents Sock As AxMSWinsockLib.AxWinsock
End Class
