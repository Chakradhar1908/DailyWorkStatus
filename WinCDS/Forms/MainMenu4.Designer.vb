<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class MainMenu4
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(MainMenu4))
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.imlMiniButtons = New System.Windows.Forms.ImageList(Me.components)
        Me.imlStandardButtons = New System.Windows.Forms.ImageList(Me.components)
        Me.rtbStorePolicy = New WinCDS.RichTextBoxNew()
        Me.rtbn = New WinCDS.RichTextBoxNew()
        Me.txtPassword = New System.Windows.Forms.TextBox()
        Me.cmdEnterPassword = New System.Windows.Forms.Button()
        Me.MSComm1 = New AxMSCommLib.AxMSComm()
        CType(Me.MSComm1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.Location = New System.Drawing.Point(0, 0)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(200, 100)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "GroupBox1"
        '
        'imlMiniButtons
        '
        Me.imlMiniButtons.ColorDepth = System.Windows.Forms.ColorDepth.Depth8Bit
        Me.imlMiniButtons.ImageSize = New System.Drawing.Size(16, 16)
        Me.imlMiniButtons.TransparentColor = System.Drawing.Color.Transparent
        '
        'imlStandardButtons
        '
        Me.imlStandardButtons.ColorDepth = System.Windows.Forms.ColorDepth.Depth8Bit
        Me.imlStandardButtons.ImageSize = New System.Drawing.Size(16, 16)
        Me.imlStandardButtons.TransparentColor = System.Drawing.Color.Transparent
        '
        'rtbStorePolicy
        '
        Me.rtbStorePolicy.Location = New System.Drawing.Point(91, 165)
        Me.rtbStorePolicy.Name = "rtbStorePolicy"
        Me.rtbStorePolicy.Size = New System.Drawing.Size(150, 150)
        Me.rtbStorePolicy.TabIndex = 1
        '
        'rtbn
        '
        Me.rtbn.Location = New System.Drawing.Point(228, 211)
        Me.rtbn.Name = "rtbn"
        Me.rtbn.Size = New System.Drawing.Size(150, 150)
        Me.rtbn.TabIndex = 2
        '
        'txtPassword
        '
        Me.txtPassword.Location = New System.Drawing.Point(464, 265)
        Me.txtPassword.Name = "txtPassword"
        Me.txtPassword.Size = New System.Drawing.Size(100, 20)
        Me.txtPassword.TabIndex = 3
        '
        'cmdEnterPassword
        '
        Me.cmdEnterPassword.Location = New System.Drawing.Point(464, 323)
        Me.cmdEnterPassword.Name = "cmdEnterPassword"
        Me.cmdEnterPassword.Size = New System.Drawing.Size(75, 23)
        Me.cmdEnterPassword.TabIndex = 4
        Me.cmdEnterPassword.Text = "Password"
        Me.cmdEnterPassword.UseVisualStyleBackColor = True
        '
        'MSComm1
        '
        Me.MSComm1.Enabled = True
        Me.MSComm1.Location = New System.Drawing.Point(427, 369)
        Me.MSComm1.Name = "MSComm1"
        Me.MSComm1.OcxState = CType(resources.GetObject("MSComm1.OcxState"), System.Windows.Forms.AxHost.State)
        Me.MSComm1.Size = New System.Drawing.Size(38, 38)
        Me.MSComm1.TabIndex = 5
        '
        'MainMenu4
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(800, 450)
        Me.Controls.Add(Me.MSComm1)
        Me.Controls.Add(Me.cmdEnterPassword)
        Me.Controls.Add(Me.txtPassword)
        Me.Controls.Add(Me.rtbn)
        Me.Controls.Add(Me.rtbStorePolicy)
        Me.Controls.Add(Me.GroupBox1)
        Me.Name = "MainMenu4"
        Me.Text = "MainMenu4"
        CType(Me.MSComm1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents GroupBox1 As GroupBox
    Friend WithEvents imlMiniButtons As ImageList
    Friend WithEvents imlStandardButtons As ImageList
    Friend WithEvents rtbStorePolicy As RichTextBoxNew
    Friend WithEvents rtbn As RichTextBoxNew
    Friend WithEvents txtPassword As TextBox
    Friend WithEvents cmdEnterPassword As Button
    Friend WithEvents MSComm1 As AxMSCommLib.AxMSComm
End Class
