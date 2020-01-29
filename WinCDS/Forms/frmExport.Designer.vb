<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmExport
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
        Me.fraImport = New System.Windows.Forms.GroupBox()
        Me.fraType = New System.Windows.Forms.GroupBox()
        Me.picSize = New System.Windows.Forms.PictureBox()
        CType(Me.picSize, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'fraImport
        '
        Me.fraImport.Location = New System.Drawing.Point(334, 60)
        Me.fraImport.Name = "fraImport"
        Me.fraImport.Size = New System.Drawing.Size(200, 100)
        Me.fraImport.TabIndex = 0
        Me.fraImport.TabStop = False
        Me.fraImport.Text = "GroupBox1"
        '
        'fraType
        '
        Me.fraType.Location = New System.Drawing.Point(351, 187)
        Me.fraType.Name = "fraType"
        Me.fraType.Size = New System.Drawing.Size(77, 46)
        Me.fraType.TabIndex = 1
        Me.fraType.TabStop = False
        Me.fraType.Text = "GroupBox1"
        '
        'picSize
        '
        Me.picSize.Location = New System.Drawing.Point(356, 257)
        Me.picSize.Name = "picSize"
        Me.picSize.Size = New System.Drawing.Size(100, 50)
        Me.picSize.TabIndex = 2
        Me.picSize.TabStop = False
        '
        'frmExport
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(800, 450)
        Me.Controls.Add(Me.picSize)
        Me.Controls.Add(Me.fraType)
        Me.Controls.Add(Me.fraImport)
        Me.Name = "frmExport"
        Me.Text = "frmExport"
        CType(Me.picSize, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents fraImport As GroupBox
    Friend WithEvents fraType As GroupBox
    Friend WithEvents picSize As PictureBox
End Class
