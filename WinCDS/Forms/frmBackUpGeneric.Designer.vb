<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmBackUpGeneric
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
        Me.lblDriveSelect = New System.Windows.Forms.Label()
        Me.chkNewFolder = New System.Windows.Forms.CheckBox()
        Me.txtNewFolder = New System.Windows.Forms.TextBox()
        Me.lvwFiles = New System.Windows.Forms.ListView()
        Me.cmdStart = New System.Windows.Forms.Button()
        Me.cmdCancel = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'lblDriveSelect
        '
        Me.lblDriveSelect.AutoSize = True
        Me.lblDriveSelect.Location = New System.Drawing.Point(0, 0)
        Me.lblDriveSelect.Name = "lblDriveSelect"
        Me.lblDriveSelect.Size = New System.Drawing.Size(39, 13)
        Me.lblDriveSelect.TabIndex = 0
        Me.lblDriveSelect.Text = "Label1"
        '
        'chkNewFolder
        '
        Me.chkNewFolder.AutoSize = True
        Me.chkNewFolder.Location = New System.Drawing.Point(354, 104)
        Me.chkNewFolder.Name = "chkNewFolder"
        Me.chkNewFolder.Size = New System.Drawing.Size(81, 17)
        Me.chkNewFolder.TabIndex = 1
        Me.chkNewFolder.Text = "CheckBox1"
        Me.chkNewFolder.UseVisualStyleBackColor = True
        '
        'txtNewFolder
        '
        Me.txtNewFolder.Location = New System.Drawing.Point(363, 162)
        Me.txtNewFolder.Name = "txtNewFolder"
        Me.txtNewFolder.Size = New System.Drawing.Size(100, 20)
        Me.txtNewFolder.TabIndex = 2
        '
        'lvwFiles
        '
        Me.lvwFiles.HideSelection = False
        Me.lvwFiles.Location = New System.Drawing.Point(290, 219)
        Me.lvwFiles.Name = "lvwFiles"
        Me.lvwFiles.Size = New System.Drawing.Size(121, 97)
        Me.lvwFiles.TabIndex = 3
        Me.lvwFiles.UseCompatibleStateImageBehavior = False
        '
        'cmdStart
        '
        Me.cmdStart.Location = New System.Drawing.Point(312, 357)
        Me.cmdStart.Name = "cmdStart"
        Me.cmdStart.Size = New System.Drawing.Size(75, 23)
        Me.cmdStart.TabIndex = 4
        Me.cmdStart.Text = "Button1"
        Me.cmdStart.UseVisualStyleBackColor = True
        '
        'cmdCancel
        '
        Me.cmdCancel.Location = New System.Drawing.Point(428, 369)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.Size = New System.Drawing.Size(75, 23)
        Me.cmdCancel.TabIndex = 5
        Me.cmdCancel.Text = "Button1"
        Me.cmdCancel.UseVisualStyleBackColor = True
        '
        'frmBackUpGeneric
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(800, 450)
        Me.Controls.Add(Me.cmdCancel)
        Me.Controls.Add(Me.cmdStart)
        Me.Controls.Add(Me.lvwFiles)
        Me.Controls.Add(Me.txtNewFolder)
        Me.Controls.Add(Me.chkNewFolder)
        Me.Controls.Add(Me.lblDriveSelect)
        Me.Name = "frmBackUpGeneric"
        Me.Text = "frmBackUpGeneric"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents lblDriveSelect As Label
    Friend WithEvents chkNewFolder As CheckBox
    Friend WithEvents txtNewFolder As TextBox
    Friend WithEvents lvwFiles As ListView
    Friend WithEvents cmdStart As Button
    Friend WithEvents cmdCancel As Button
End Class
