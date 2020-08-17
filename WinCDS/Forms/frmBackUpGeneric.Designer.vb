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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmBackUpGeneric))
        Me.lblDriveSelect = New System.Windows.Forms.Label()
        Me.chkNewFolder = New System.Windows.Forms.CheckBox()
        Me.txtNewFolder = New System.Windows.Forms.TextBox()
        Me.lvwFiles = New System.Windows.Forms.ListView()
        Me.cmdStart = New System.Windows.Forms.Button()
        Me.cmdCancel = New System.Windows.Forms.Button()
        Me.NewZipBackup = New AxVJCZIPLib.AxVjcZip()
        Me.fraExtra = New System.Windows.Forms.GroupBox()
        Me.prgExtra = New System.Windows.Forms.ProgressBar()
        Me.sb = New AxMSComctlLib.AxStatusBar()
        CType(Me.NewZipBackup, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.sb, System.ComponentModel.ISupportInitialize).BeginInit()
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
        'NewZipBackup
        '
        Me.NewZipBackup.Enabled = True
        Me.NewZipBackup.Location = New System.Drawing.Point(558, 259)
        Me.NewZipBackup.Name = "NewZipBackup"
        Me.NewZipBackup.OcxState = CType(resources.GetObject("NewZipBackup.OcxState"), System.Windows.Forms.AxHost.State)
        Me.NewZipBackup.Size = New System.Drawing.Size(100, 50)
        Me.NewZipBackup.TabIndex = 7
        '
        'fraExtra
        '
        Me.fraExtra.Location = New System.Drawing.Point(605, 140)
        Me.fraExtra.Name = "fraExtra"
        Me.fraExtra.Size = New System.Drawing.Size(200, 100)
        Me.fraExtra.TabIndex = 8
        Me.fraExtra.TabStop = False
        Me.fraExtra.Text = "GroupBox1"
        '
        'prgExtra
        '
        Me.prgExtra.Location = New System.Drawing.Point(581, 364)
        Me.prgExtra.Name = "prgExtra"
        Me.prgExtra.Size = New System.Drawing.Size(100, 23)
        Me.prgExtra.TabIndex = 9
        '
        'sb
        '
        Me.sb.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.sb.Location = New System.Drawing.Point(0, 425)
        Me.sb.Name = "sb"
        Me.sb.OcxState = CType(resources.GetObject("sb.OcxState"), System.Windows.Forms.AxHost.State)
        Me.sb.Size = New System.Drawing.Size(800, 25)
        Me.sb.TabIndex = 10
        '
        'frmBackUpGeneric
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(800, 450)
        Me.Controls.Add(Me.sb)
        Me.Controls.Add(Me.prgExtra)
        Me.Controls.Add(Me.fraExtra)
        Me.Controls.Add(Me.NewZipBackup)
        Me.Controls.Add(Me.cmdCancel)
        Me.Controls.Add(Me.cmdStart)
        Me.Controls.Add(Me.lvwFiles)
        Me.Controls.Add(Me.txtNewFolder)
        Me.Controls.Add(Me.chkNewFolder)
        Me.Controls.Add(Me.lblDriveSelect)
        Me.Name = "frmBackUpGeneric"
        Me.Text = "frmBackUpGeneric"
        CType(Me.NewZipBackup, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.sb, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents lblDriveSelect As Label
    Friend WithEvents chkNewFolder As CheckBox
    Friend WithEvents txtNewFolder As TextBox
    Friend WithEvents lvwFiles As ListView
    Friend WithEvents cmdStart As Button
    Friend WithEvents cmdCancel As Button
    'Friend WithEvents sb As AxComctlLib.AxStatusBar
    Friend WithEvents NewZipBackup As AxVJCZIPLib.AxVjcZip
    Friend WithEvents fraExtra As GroupBox
    Friend WithEvents prgExtra As ProgressBar
    Friend WithEvents sb As AxMSComctlLib.AxStatusBar
End Class
