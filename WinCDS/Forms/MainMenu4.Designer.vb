<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class MainMenu4
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
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(MainMenu4))
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.imlMiniButtons = New System.Windows.Forms.ImageList(Me.components)
        Me.imlStandardButtons = New System.Windows.Forms.ImageList(Me.components)
        Me.txtPassword = New System.Windows.Forms.TextBox()
        Me.cmdEnterPassword = New System.Windows.Forms.Button()
        Me.MSComm1 = New AxMSCommLib.AxMSComm()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.rtbn = New WinCDS.RichTextBoxNew()
        Me.rtbStorePolicy = New WinCDS.RichTextBoxNew()
        Me.imlSmallButtons = New System.Windows.Forms.ImageList(Me.components)
        Me.datPicture = New Microsoft.VisualBasic.Compatibility.VB6.ADODC()
        Me.imgPicture = New System.Windows.Forms.PictureBox()
        Me.rtb = New System.Windows.Forms.RichTextBox()
        CType(Me.MSComm1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.imgPicture, System.ComponentModel.ISupportInitialize).BeginInit()
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
        Me.imlStandardButtons.ImageStream = CType(resources.GetObject("imlStandardButtons.ImageStream"), System.Windows.Forms.ImageListStreamer)
        Me.imlStandardButtons.TransparentColor = System.Drawing.Color.Transparent
        Me.imlStandardButtons.Images.SetKeyName(0, "StandardButton-Add.gif")
        Me.imlStandardButtons.Images.SetKeyName(1, "StandardButton-ForwardMenu.gif")
        Me.imlStandardButtons.Images.SetKeyName(2, "StandardButton-OK.gif")
        Me.imlStandardButtons.Images.SetKeyName(3, "StandardButton-Cancel.gif")
        Me.imlStandardButtons.Images.SetKeyName(4, "StandardButton-Back.gif")
        Me.imlStandardButtons.Images.SetKeyName(5, "StandardButton-Foward.gif")
        Me.imlStandardButtons.Images.SetKeyName(6, "StandardButton-Next.gif")
        Me.imlStandardButtons.Images.SetKeyName(7, "StandardButton-Previous.gif")
        Me.imlStandardButtons.Images.SetKeyName(8, "StandardButton-BackMenu.gif")
        Me.imlStandardButtons.Images.SetKeyName(9, "StandardButton-Reload.gif")
        Me.imlStandardButtons.Images.SetKeyName(10, "StandardButton-Delete.gif")
        Me.imlStandardButtons.Images.SetKeyName(11, "StandardButton-Minus.gif")
        Me.imlStandardButtons.Images.SetKeyName(12, "StandardButton-Refresh.gif")
        Me.imlStandardButtons.Images.SetKeyName(13, "StandardButton-Down.gif")
        Me.imlStandardButtons.Images.SetKeyName(14, "StandardButton-Left.gif")
        Me.imlStandardButtons.Images.SetKeyName(15, "StandardButton-Right.gif")
        Me.imlStandardButtons.Images.SetKeyName(16, "StandardButton-Up.gif")
        Me.imlStandardButtons.Images.SetKeyName(17, "poorder.gif")
        Me.imlStandardButtons.Images.SetKeyName(18, "StandardButton-Calendar.gif")
        Me.imlStandardButtons.Images.SetKeyName(19, "StandardButton-Print.gif")
        Me.imlStandardButtons.Images.SetKeyName(20, "StandardButton-Preview.gif")
        Me.imlStandardButtons.Images.SetKeyName(21, "StandardButton-Compass.gif")
        Me.imlStandardButtons.Images.SetKeyName(22, "StandardButton-Clear.gif")
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
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(357, 50)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(75, 23)
        Me.Button1.TabIndex = 6
        Me.Button1.Text = "Button1"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'rtbn
        '
        Me.rtbn.Location = New System.Drawing.Point(228, 211)
        Me.rtbn.Name = "rtbn"
        Me.rtbn.Size = New System.Drawing.Size(150, 150)
        Me.rtbn.TabIndex = 2
        '
        'rtbStorePolicy
        '
        Me.rtbStorePolicy.Location = New System.Drawing.Point(91, 165)
        Me.rtbStorePolicy.Name = "rtbStorePolicy"
        Me.rtbStorePolicy.Size = New System.Drawing.Size(150, 150)
        Me.rtbStorePolicy.TabIndex = 1
        '
        'imlSmallButtons
        '
        Me.imlSmallButtons.ImageStream = CType(resources.GetObject("imlSmallButtons.ImageStream"), System.Windows.Forms.ImageListStreamer)
        Me.imlSmallButtons.TransparentColor = System.Drawing.Color.Transparent
        Me.imlSmallButtons.Images.SetKeyName(0, "disc.gif")
        Me.imlSmallButtons.Images.SetKeyName(1, "email.gif")
        Me.imlSmallButtons.Images.SetKeyName(2, "nodisc.gif")
        Me.imlSmallButtons.Images.SetKeyName(3, "print.gif")
        Me.imlSmallButtons.Images.SetKeyName(4, "StandardButton-DownMenu.gif")
        Me.imlSmallButtons.Images.SetKeyName(5, "StandardButton-UpMenu.gif")
        Me.imlSmallButtons.Images.SetKeyName(6, "StandardButton-Web3.gif")
        '
        'datPicture
        '
        Me.datPicture.BackColor = System.Drawing.SystemColors.Control
        Me.datPicture.CommandTimeout = 0
        Me.datPicture.CommandType = ADODB.CommandTypeEnum.adCmdUnknown
        Me.datPicture.ConnectionString = Nothing
        Me.datPicture.CursorType = ADODB.CursorTypeEnum.adOpenStatic
        Me.datPicture.Location = New System.Drawing.Point(91, 381)
        Me.datPicture.LockType = ADODB.LockTypeEnum.adLockOptimistic
        Me.datPicture.Name = "datPicture"
        Me.datPicture.Size = New System.Drawing.Size(83, 26)
        Me.datPicture.TabIndex = 7
        Me.datPicture.Text = "Adodc1"
        '
        'imgPicture
        '
        Me.imgPicture.Location = New System.Drawing.Point(201, 381)
        Me.imgPicture.Name = "imgPicture"
        Me.imgPicture.Size = New System.Drawing.Size(100, 50)
        Me.imgPicture.TabIndex = 8
        Me.imgPicture.TabStop = False
        '
        'rtb
        '
        Me.rtb.Location = New System.Drawing.Point(595, 323)
        Me.rtb.Name = "rtb"
        Me.rtb.Size = New System.Drawing.Size(100, 96)
        Me.rtb.TabIndex = 9
        Me.rtb.Text = ""
        '
        'MainMenu4
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(800, 450)
        Me.Controls.Add(Me.rtb)
        Me.Controls.Add(Me.imgPicture)
        Me.Controls.Add(Me.datPicture)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.MSComm1)
        Me.Controls.Add(Me.cmdEnterPassword)
        Me.Controls.Add(Me.txtPassword)
        Me.Controls.Add(Me.rtbn)
        Me.Controls.Add(Me.rtbStorePolicy)
        Me.Controls.Add(Me.GroupBox1)
        Me.Name = "MainMenu4"
        Me.Text = "MainMenu4"
        CType(Me.MSComm1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.imgPicture, System.ComponentModel.ISupportInitialize).EndInit()
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
    Friend WithEvents Button1 As Button
    Friend WithEvents imlSmallButtons As ImageList
    Friend WithEvents datPicture As Compatibility.VB6.ADODC
    Friend WithEvents imgPicture As PictureBox
    Friend WithEvents rtb As RichTextBox
End Class
