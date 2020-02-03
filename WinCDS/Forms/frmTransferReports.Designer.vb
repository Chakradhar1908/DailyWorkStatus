<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmTransferReports
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
        Me.cmdPrint0 = New System.Windows.Forms.Button()
        Me.cmdPrint1 = New System.Windows.Forms.Button()
        Me.cmdCancel = New System.Windows.Forms.Button()
        Me.fraPending = New System.Windows.Forms.GroupBox()
        Me.fraPrevious = New System.Windows.Forms.GroupBox()
        Me.SuspendLayout()
        '
        'cmdPrint0
        '
        Me.cmdPrint0.Location = New System.Drawing.Point(128, 16)
        Me.cmdPrint0.Name = "cmdPrint0"
        Me.cmdPrint0.Size = New System.Drawing.Size(75, 23)
        Me.cmdPrint0.TabIndex = 0
        Me.cmdPrint0.Text = "Button1"
        Me.cmdPrint0.UseVisualStyleBackColor = True
        '
        'cmdPrint1
        '
        Me.cmdPrint1.Location = New System.Drawing.Point(128, 65)
        Me.cmdPrint1.Name = "cmdPrint1"
        Me.cmdPrint1.Size = New System.Drawing.Size(75, 23)
        Me.cmdPrint1.TabIndex = 1
        Me.cmdPrint1.Text = "Button2"
        Me.cmdPrint1.UseVisualStyleBackColor = True
        '
        'cmdCancel
        '
        Me.cmdCancel.Location = New System.Drawing.Point(131, 113)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.Size = New System.Drawing.Size(75, 23)
        Me.cmdCancel.TabIndex = 2
        Me.cmdCancel.Text = "Button3"
        Me.cmdCancel.UseVisualStyleBackColor = True
        '
        'fraPending
        '
        Me.fraPending.Location = New System.Drawing.Point(128, 159)
        Me.fraPending.Name = "fraPending"
        Me.fraPending.Size = New System.Drawing.Size(200, 100)
        Me.fraPending.TabIndex = 3
        Me.fraPending.TabStop = False
        Me.fraPending.Text = "GroupBox1"
        '
        'fraPrevious
        '
        Me.fraPrevious.Location = New System.Drawing.Point(128, 276)
        Me.fraPrevious.Name = "fraPrevious"
        Me.fraPrevious.Size = New System.Drawing.Size(200, 100)
        Me.fraPrevious.TabIndex = 4
        Me.fraPrevious.TabStop = False
        Me.fraPrevious.Text = "GroupBox2"
        '
        'frmTransferReports
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(800, 450)
        Me.Controls.Add(Me.fraPrevious)
        Me.Controls.Add(Me.fraPending)
        Me.Controls.Add(Me.cmdCancel)
        Me.Controls.Add(Me.cmdPrint1)
        Me.Controls.Add(Me.cmdPrint0)
        Me.Name = "frmTransferReports"
        Me.Text = "frmTransferReports"
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents cmdPrint0 As Button
    Friend WithEvents cmdPrint1 As Button
    Friend WithEvents cmdCancel As Button
    Friend WithEvents fraPending As GroupBox
    Friend WithEvents fraPrevious As GroupBox
End Class
