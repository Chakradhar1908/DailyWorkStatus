<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmMsg2
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
        Me.components = New System.ComponentModel.Container()
        Me.tmrMax = New System.Windows.Forms.Timer(Me.components)
        Me.picIcon = New System.Windows.Forms.PictureBox()
        Me.imlStyles = New System.Windows.Forms.ImageList(Me.components)
        Me.lblMessage = New System.Windows.Forms.Label()
        Me.cmdButton1 = New System.Windows.Forms.Button()
        Me.cmdButton2 = New System.Windows.Forms.Button()
        Me.cmdButton3 = New System.Windows.Forms.Button()
        Me.cmdButton4 = New System.Windows.Forms.Button()
        Me.txtConfirm = New System.Windows.Forms.TextBox()
        Me.picButtons = New System.Windows.Forms.PictureBox()
        CType(Me.picIcon, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.picButtons, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'tmrMax
        '
        Me.tmrMax.Interval = 1000
        '
        'picIcon
        '
        Me.picIcon.Location = New System.Drawing.Point(215, 342)
        Me.picIcon.Name = "picIcon"
        Me.picIcon.Size = New System.Drawing.Size(100, 50)
        Me.picIcon.TabIndex = 0
        Me.picIcon.TabStop = False
        '
        'imlStyles
        '
        Me.imlStyles.ColorDepth = System.Windows.Forms.ColorDepth.Depth8Bit
        Me.imlStyles.ImageSize = New System.Drawing.Size(16, 16)
        Me.imlStyles.TransparentColor = System.Drawing.Color.Transparent
        '
        'lblMessage
        '
        Me.lblMessage.AutoSize = True
        Me.lblMessage.Location = New System.Drawing.Point(24, 9)
        Me.lblMessage.Name = "lblMessage"
        Me.lblMessage.Size = New System.Drawing.Size(39, 13)
        Me.lblMessage.TabIndex = 1
        Me.lblMessage.Text = "Label1"
        '
        'cmdButton1
        '
        Me.cmdButton1.Location = New System.Drawing.Point(27, 79)
        Me.cmdButton1.Name = "cmdButton1"
        Me.cmdButton1.Size = New System.Drawing.Size(65, 23)
        Me.cmdButton1.TabIndex = 2
        Me.cmdButton1.Text = "Button1"
        Me.cmdButton1.UseVisualStyleBackColor = True
        '
        'cmdButton2
        '
        Me.cmdButton2.Location = New System.Drawing.Point(100, 79)
        Me.cmdButton2.Name = "cmdButton2"
        Me.cmdButton2.Size = New System.Drawing.Size(65, 23)
        Me.cmdButton2.TabIndex = 3
        Me.cmdButton2.Text = "Button2"
        Me.cmdButton2.UseVisualStyleBackColor = True
        '
        'cmdButton3
        '
        Me.cmdButton3.Location = New System.Drawing.Point(173, 79)
        Me.cmdButton3.Name = "cmdButton3"
        Me.cmdButton3.Size = New System.Drawing.Size(65, 23)
        Me.cmdButton3.TabIndex = 4
        Me.cmdButton3.Text = "Button3"
        Me.cmdButton3.UseVisualStyleBackColor = True
        '
        'cmdButton4
        '
        Me.cmdButton4.Location = New System.Drawing.Point(246, 79)
        Me.cmdButton4.Name = "cmdButton4"
        Me.cmdButton4.Size = New System.Drawing.Size(65, 23)
        Me.cmdButton4.TabIndex = 5
        Me.cmdButton4.Text = "Button4"
        Me.cmdButton4.UseVisualStyleBackColor = True
        '
        'txtConfirm
        '
        Me.txtConfirm.Location = New System.Drawing.Point(12, 39)
        Me.txtConfirm.Name = "txtConfirm"
        Me.txtConfirm.Size = New System.Drawing.Size(100, 20)
        Me.txtConfirm.TabIndex = 6
        '
        'picButtons
        '
        Me.picButtons.Location = New System.Drawing.Point(18, 76)
        Me.picButtons.Name = "picButtons"
        Me.picButtons.Size = New System.Drawing.Size(353, 34)
        Me.picButtons.TabIndex = 7
        Me.picButtons.TabStop = False
        '
        'FrmMsg2
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(800, 450)
        Me.Controls.Add(Me.txtConfirm)
        Me.Controls.Add(Me.cmdButton4)
        Me.Controls.Add(Me.cmdButton3)
        Me.Controls.Add(Me.cmdButton2)
        Me.Controls.Add(Me.cmdButton1)
        Me.Controls.Add(Me.lblMessage)
        Me.Controls.Add(Me.picIcon)
        Me.Controls.Add(Me.picButtons)
        Me.Name = "FrmMsg2"
        Me.Text = "frmMsg2"
        CType(Me.picIcon, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.picButtons, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents tmrMax As Timer
    Friend WithEvents picIcon As PictureBox
    Friend WithEvents imlStyles As ImageList
    Friend WithEvents lblMessage As Label
    Friend WithEvents cmdButton1 As Button
    Friend WithEvents cmdButton2 As Button
    Friend WithEvents cmdButton3 As Button
    Friend WithEvents cmdButton4 As Button
    Friend WithEvents txtConfirm As TextBox
    Friend WithEvents picButtons As PictureBox
End Class
