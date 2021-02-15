<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class frmPrintPreviewDocument
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
        Me.picPicture = New System.Windows.Forms.PictureBox()
        Me.fraNavigate = New System.Windows.Forms.GroupBox()
        Me.btnClose = New System.Windows.Forms.Button()
        Me.btnPrint = New System.Windows.Forms.Button()
        Me.lblHelp = New System.Windows.Forms.Label()
        Me.PrintDialog1 = New System.Windows.Forms.PrintDialog()
        Me.PrintDocument1 = New System.Drawing.Printing.PrintDocument()
        CType(Me.picPicture, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.fraNavigate.SuspendLayout()
        Me.SuspendLayout()
        '
        'picPicture
        '
        Me.picPicture.Location = New System.Drawing.Point(1, 12)
        Me.picPicture.Name = "picPicture"
        Me.picPicture.Size = New System.Drawing.Size(268, 88)
        Me.picPicture.TabIndex = 0
        Me.picPicture.TabStop = False
        '
        'fraNavigate
        '
        Me.fraNavigate.Controls.Add(Me.btnClose)
        Me.fraNavigate.Controls.Add(Me.btnPrint)
        Me.fraNavigate.Location = New System.Drawing.Point(69, 156)
        Me.fraNavigate.Name = "fraNavigate"
        Me.fraNavigate.Size = New System.Drawing.Size(173, 68)
        Me.fraNavigate.TabIndex = 2
        Me.fraNavigate.TabStop = False
        '
        'btnClose
        '
        Me.btnClose.Location = New System.Drawing.Point(98, 10)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(69, 54)
        Me.btnClose.TabIndex = 5
        Me.btnClose.Text = "Close"
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'btnPrint
        '
        Me.btnPrint.Location = New System.Drawing.Point(6, 10)
        Me.btnPrint.Name = "btnPrint"
        Me.btnPrint.Size = New System.Drawing.Size(69, 54)
        Me.btnPrint.TabIndex = 4
        Me.btnPrint.Text = "Print"
        Me.btnPrint.UseVisualStyleBackColor = True
        '
        'lblHelp
        '
        Me.lblHelp.AutoSize = True
        Me.lblHelp.Location = New System.Drawing.Point(24, 188)
        Me.lblHelp.Name = "lblHelp"
        Me.lblHelp.Size = New System.Drawing.Size(39, 13)
        Me.lblHelp.TabIndex = 3
        Me.lblHelp.Text = "Label1"
        '
        'PrintDialog1
        '
        Me.PrintDialog1.UseEXDialog = True
        '
        'PrintDocument1
        '
        '
        'frmPrintPreviewDocument
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(322, 236)
        Me.ControlBox = False
        Me.Controls.Add(Me.picPicture)
        Me.Controls.Add(Me.lblHelp)
        Me.Controls.Add(Me.fraNavigate)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmPrintPreviewDocument"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "Up"
        CType(Me.picPicture, System.ComponentModel.ISupportInitialize).EndInit()
        Me.fraNavigate.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents picPicture As PictureBox
    Friend WithEvents fraNavigate As GroupBox
    Friend WithEvents lblHelp As Label
    Friend WithEvents btnPrint As Button
    Friend WithEvents btnClose As Button
    Friend WithEvents PrintDialog1 As PrintDialog
    Friend WithEvents PrintDocument1 As Printing.PrintDocument
End Class
