<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmPrintPreviewDocument
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
        Me.picPicture = New System.Windows.Forms.PictureBox()
        Me.cmdNavigate7 = New System.Windows.Forms.Button()
        Me.fraNavigate = New System.Windows.Forms.GroupBox()
        Me.lblHelp = New System.Windows.Forms.Label()
        CType(Me.picPicture, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'picPicture
        '
        Me.picPicture.Location = New System.Drawing.Point(0, 0)
        Me.picPicture.Name = "picPicture"
        Me.picPicture.Size = New System.Drawing.Size(100, 50)
        Me.picPicture.TabIndex = 0
        Me.picPicture.TabStop = False
        '
        'cmdNavigate7
        '
        Me.cmdNavigate7.Location = New System.Drawing.Point(53, 131)
        Me.cmdNavigate7.Name = "cmdNavigate7"
        Me.cmdNavigate7.Size = New System.Drawing.Size(75, 23)
        Me.cmdNavigate7.TabIndex = 1
        Me.cmdNavigate7.Text = "Goto"
        Me.cmdNavigate7.UseVisualStyleBackColor = True
        '
        'fraNavigate
        '
        Me.fraNavigate.Location = New System.Drawing.Point(119, 210)
        Me.fraNavigate.Name = "fraNavigate"
        Me.fraNavigate.Size = New System.Drawing.Size(200, 100)
        Me.fraNavigate.TabIndex = 2
        Me.fraNavigate.TabStop = False
        '
        'lblHelp
        '
        Me.lblHelp.AutoSize = True
        Me.lblHelp.Location = New System.Drawing.Point(547, 254)
        Me.lblHelp.Name = "lblHelp"
        Me.lblHelp.Size = New System.Drawing.Size(39, 13)
        Me.lblHelp.TabIndex = 3
        Me.lblHelp.Text = "Label1"
        '
        'frmPrintPreviewDocument
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(800, 450)
        Me.Controls.Add(Me.lblHelp)
        Me.Controls.Add(Me.fraNavigate)
        Me.Controls.Add(Me.cmdNavigate7)
        Me.Controls.Add(Me.picPicture)
        Me.Name = "frmPrintPreviewDocument"
        Me.Text = "frmPrintPreviewDocument"
        CType(Me.picPicture, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents picPicture As PictureBox
    Friend WithEvents cmdNavigate7 As Button
    Friend WithEvents fraNavigate As GroupBox
    Friend WithEvents lblHelp As Label
End Class
