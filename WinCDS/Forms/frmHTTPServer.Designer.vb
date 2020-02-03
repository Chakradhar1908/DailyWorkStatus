<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmHTTPServer
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmHTTPServer))
        Me.sck0 = New AxMSWinsockLib.AxWinsock()
        Me.lblFileProgress0 = New System.Windows.Forms.Label()
        CType(Me.sck0, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'sck0
        '
        Me.sck0.Enabled = True
        Me.sck0.Location = New System.Drawing.Point(384, 233)
        Me.sck0.Name = "sck0"
        Me.sck0.OcxState = CType(resources.GetObject("sck0.OcxState"), System.Windows.Forms.AxHost.State)
        Me.sck0.Size = New System.Drawing.Size(28, 28)
        Me.sck0.TabIndex = 0
        '
        'lblFileProgress0
        '
        Me.lblFileProgress0.AutoSize = True
        Me.lblFileProgress0.Location = New System.Drawing.Point(417, 133)
        Me.lblFileProgress0.Name = "lblFileProgress0"
        Me.lblFileProgress0.Size = New System.Drawing.Size(39, 13)
        Me.lblFileProgress0.TabIndex = 1
        Me.lblFileProgress0.Text = "Label1"
        '
        'frmHTTPServer
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(800, 450)
        Me.Controls.Add(Me.lblFileProgress0)
        Me.Controls.Add(Me.sck0)
        Me.Name = "frmHTTPServer"
        Me.Text = "frmHTTPServer"
        CType(Me.sck0, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents sck0 As AxMSWinsockLib.AxWinsock
    Friend WithEvents lblFileProgress0 As Label
End Class
