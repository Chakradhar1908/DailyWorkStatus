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
        Me.imlMiniButtons = New System.Windows.Forms.ImageList(Me.components)
        Me.imlStandardButtons = New System.Windows.Forms.ImageList(Me.components)
        Me.txtPassword = New System.Windows.Forms.TextBox()
        Me.cmdEnterPassword = New System.Windows.Forms.Button()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.imlSmallButtons = New System.Windows.Forms.ImageList(Me.components)
        Me.rtb = New System.Windows.Forms.RichTextBox()
        Me.fraSupplies = New System.Windows.Forms.GroupBox()
        Me.picAlpha = New System.Windows.Forms.PictureBox()
        Me.cdgFile = New AxMSComDlg.AxCommonDialog()
        Me.MSComm1 = New AxMSCommLib.AxMSComm()
        Me.imgPicture = New System.Windows.Forms.PictureBox()
        Me.datPicture = New Microsoft.VisualBasic.Compatibility.VB6.ADODC()
        Me.rtbStorePolicy = New WinCDS.RichTextBoxNew()
        Me.rtbn = New WinCDS.RichTextBoxNew()
        Me.imlMM = New System.Windows.Forms.ImageList(Me.components)
        Me.lblLastBackup = New System.Windows.Forms.Label()
        Me.lblDEMO = New System.Windows.Forms.Label()
        Me.lblIDE = New System.Windows.Forms.Label()
        Me.lblWinCDS = New System.Windows.Forms.Label()
        Me.lblMenuCaption = New System.Windows.Forms.Label()
        Me.lblMenuItem = New System.Windows.Forms.Label()
        Me.imgStoreLogoBorder = New System.Windows.Forms.Label()
        Me.bvb5 = New System.Windows.Forms.PictureBox()
        Me.bvb4 = New System.Windows.Forms.PictureBox()
        Me.bvb3 = New System.Windows.Forms.PictureBox()
        Me.bvb2 = New System.Windows.Forms.PictureBox()
        Me.bvb1 = New System.Windows.Forms.PictureBox()
        Me.bvb0 = New System.Windows.Forms.PictureBox()
        Me.cmdLogout = New System.Windows.Forms.PictureBox()
        Me.imgStoreLogo = New System.Windows.Forms.PictureBox()
        Me.imgHR = New System.Windows.Forms.PictureBox()
        Me.imgSelected = New System.Windows.Forms.PictureBox()
        Me.imgSubSelected = New System.Windows.Forms.PictureBox()
        Me.imgMenuItem = New System.Windows.Forms.PictureBox()
        Me.lblHR = New System.Windows.Forms.Label()
        Me.ttpMainMenu = New System.Windows.Forms.ToolTip(Me.components)
        Me.lblDevMode = New System.Windows.Forms.Label()
        Me.lblCDSComputer = New System.Windows.Forms.Label()
        Me.lblRDP = New System.Windows.Forms.Label()
        Me.lblELEVATE = New System.Windows.Forms.Label()
        Me.lblBETA = New System.Windows.Forms.Label()
        Me.txtYearEnd = New System.Windows.Forms.TextBox()
        Me.txtInfo = New System.Windows.Forms.TextBox()
        Me.lblStore0 = New System.Windows.Forms.Label()
        Me.lblStore1 = New System.Windows.Forms.Label()
        Me.lblStore2 = New System.Windows.Forms.Label()
        Me.KeyCatch = New System.Windows.Forms.TextBox()
        Me.imgInfo = New System.Windows.Forms.PictureBox()
        Me.imgBackground = New System.Windows.Forms.PictureBox()
        Me.tmrPulse = New System.Windows.Forms.Timer(Me.components)
        Me.tmrMaintain = New System.Windows.Forms.Timer(Me.components)
        Me.fraSupplies.SuspendLayout()
        CType(Me.picAlpha, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.cdgFile, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.MSComm1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.imgPicture, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.bvb5, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.bvb4, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.bvb3, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.bvb2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.bvb1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.bvb0, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.cmdLogout, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.imgStoreLogo, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.imgHR, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.imgSelected, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.imgSubSelected, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.imgMenuItem, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.imgInfo, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.imgBackground, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
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
        Me.imlStandardButtons.Images.SetKeyName(23, "calendar.bmp")
        '
        'txtPassword
        '
        Me.txtPassword.Location = New System.Drawing.Point(12, 19)
        Me.txtPassword.Name = "txtPassword"
        Me.txtPassword.Size = New System.Drawing.Size(128, 20)
        Me.txtPassword.TabIndex = 3
        Me.txtPassword.Visible = False
        '
        'cmdEnterPassword
        '
        Me.cmdEnterPassword.Location = New System.Drawing.Point(35, 41)
        Me.cmdEnterPassword.Name = "cmdEnterPassword"
        Me.cmdEnterPassword.Size = New System.Drawing.Size(75, 23)
        Me.cmdEnterPassword.TabIndex = 4
        Me.cmdEnterPassword.Text = "Password"
        Me.cmdEnterPassword.UseVisualStyleBackColor = True
        Me.cmdEnterPassword.Visible = False
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(690, 526)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(75, 23)
        Me.Button1.TabIndex = 6
        Me.Button1.Text = "New Sale"
        Me.Button1.UseVisualStyleBackColor = True
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
        'rtb
        '
        Me.rtb.Location = New System.Drawing.Point(88, 19)
        Me.rtb.Name = "rtb"
        Me.rtb.Size = New System.Drawing.Size(40, 38)
        Me.rtb.TabIndex = 9
        Me.rtb.Text = ""
        Me.rtb.Visible = False
        '
        'fraSupplies
        '
        Me.fraSupplies.Controls.Add(Me.picAlpha)
        Me.fraSupplies.Controls.Add(Me.cdgFile)
        Me.fraSupplies.Controls.Add(Me.MSComm1)
        Me.fraSupplies.Controls.Add(Me.imgPicture)
        Me.fraSupplies.Controls.Add(Me.rtb)
        Me.fraSupplies.Controls.Add(Me.datPicture)
        Me.fraSupplies.Controls.Add(Me.rtbStorePolicy)
        Me.fraSupplies.Controls.Add(Me.rtbn)
        Me.fraSupplies.Location = New System.Drawing.Point(552, 95)
        Me.fraSupplies.Name = "fraSupplies"
        Me.fraSupplies.Size = New System.Drawing.Size(236, 106)
        Me.fraSupplies.TabIndex = 14
        Me.fraSupplies.TabStop = False
        Me.fraSupplies.Text = "SUPPLIES - NOT FOR DISPLAY"
        Me.fraSupplies.Visible = False
        '
        'picAlpha
        '
        Me.picAlpha.Location = New System.Drawing.Point(149, 65)
        Me.picAlpha.Name = "picAlpha"
        Me.picAlpha.Size = New System.Drawing.Size(33, 25)
        Me.picAlpha.TabIndex = 16
        Me.picAlpha.TabStop = False
        Me.picAlpha.Visible = False
        '
        'cdgFile
        '
        Me.cdgFile.Enabled = True
        Me.cdgFile.Location = New System.Drawing.Point(50, 19)
        Me.cdgFile.Name = "cdgFile"
        Me.cdgFile.OcxState = CType(resources.GetObject("cdgFile.OcxState"), System.Windows.Forms.AxHost.State)
        Me.cdgFile.Size = New System.Drawing.Size(32, 32)
        Me.cdgFile.TabIndex = 15
        '
        'MSComm1
        '
        Me.MSComm1.Enabled = True
        Me.MSComm1.Location = New System.Drawing.Point(6, 19)
        Me.MSComm1.Name = "MSComm1"
        Me.MSComm1.OcxState = CType(resources.GetObject("MSComm1.OcxState"), System.Windows.Forms.AxHost.State)
        Me.MSComm1.Size = New System.Drawing.Size(38, 38)
        Me.MSComm1.TabIndex = 5
        Me.MSComm1.Visible = False
        '
        'imgPicture
        '
        Me.imgPicture.Location = New System.Drawing.Point(88, 63)
        Me.imgPicture.Name = "imgPicture"
        Me.imgPicture.Size = New System.Drawing.Size(40, 27)
        Me.imgPicture.TabIndex = 8
        Me.imgPicture.TabStop = False
        Me.imgPicture.Visible = False
        '
        'datPicture
        '
        Me.datPicture.BackColor = System.Drawing.SystemColors.Control
        Me.datPicture.CommandTimeout = 0
        Me.datPicture.CommandType = ADODB.CommandTypeEnum.adCmdUnknown
        Me.datPicture.ConnectionString = Nothing
        Me.datPicture.CursorType = ADODB.CursorTypeEnum.adOpenStatic
        Me.datPicture.Location = New System.Drawing.Point(8, 72)
        Me.datPicture.LockType = ADODB.LockTypeEnum.adLockOptimistic
        Me.datPicture.Name = "datPicture"
        Me.datPicture.Size = New System.Drawing.Size(74, 19)
        Me.datPicture.TabIndex = 7
        Me.datPicture.Text = "Adodc1"
        Me.datPicture.Visible = False
        '
        'rtbStorePolicy
        '
        Me.rtbStorePolicy.Location = New System.Drawing.Point(134, 19)
        Me.rtbStorePolicy.Name = "rtbStorePolicy"
        Me.rtbStorePolicy.Size = New System.Drawing.Size(40, 38)
        Me.rtbStorePolicy.TabIndex = 1
        Me.rtbStorePolicy.Visible = False
        '
        'rtbn
        '
        Me.rtbn.Location = New System.Drawing.Point(180, 19)
        Me.rtbn.Name = "rtbn"
        Me.rtbn.Size = New System.Drawing.Size(40, 38)
        Me.rtbn.TabIndex = 2
        Me.rtbn.Visible = False
        '
        'imlMM
        '
        Me.imlMM.ColorDepth = System.Windows.Forms.ColorDepth.Depth8Bit
        Me.imlMM.ImageSize = New System.Drawing.Size(16, 16)
        Me.imlMM.TransparentColor = System.Drawing.Color.Transparent
        '
        'lblLastBackup
        '
        Me.lblLastBackup.AutoSize = True
        Me.lblLastBackup.BackColor = System.Drawing.SystemColors.Control
        Me.lblLastBackup.Font = New System.Drawing.Font("Arial Black", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblLastBackup.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblLastBackup.Location = New System.Drawing.Point(399, 4)
        Me.lblLastBackup.Name = "lblLastBackup"
        Me.lblLastBackup.Size = New System.Drawing.Size(131, 17)
        Me.lblLastBackup.TabIndex = 15
        Me.lblLastBackup.Text = "### LAST BACKUP"
        '
        'lblDEMO
        '
        Me.lblDEMO.AutoSize = True
        Me.lblDEMO.Font = New System.Drawing.Font("Lucida Console", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblDEMO.Location = New System.Drawing.Point(579, 4)
        Me.lblDEMO.Name = "lblDEMO"
        Me.lblDEMO.Size = New System.Drawing.Size(37, 11)
        Me.lblDEMO.TabIndex = 16
        Me.lblDEMO.Text = "DEMO"
        '
        'lblIDE
        '
        Me.lblIDE.AutoSize = True
        Me.lblIDE.Font = New System.Drawing.Font("Lucida Console", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblIDE.Location = New System.Drawing.Point(620, 4)
        Me.lblIDE.Name = "lblIDE"
        Me.lblIDE.Size = New System.Drawing.Size(37, 11)
        Me.lblIDE.TabIndex = 17
        Me.lblIDE.Text = "IDE "
        '
        'lblWinCDS
        '
        Me.lblWinCDS.AutoSize = True
        Me.lblWinCDS.Font = New System.Drawing.Font("Arial", 24.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblWinCDS.Location = New System.Drawing.Point(580, 26)
        Me.lblWinCDS.Name = "lblWinCDS"
        Me.lblWinCDS.Size = New System.Drawing.Size(217, 36)
        Me.lblWinCDS.TabIndex = 18
        Me.lblWinCDS.Text = "WinCDS PRO"
        Me.lblWinCDS.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'lblMenuCaption
        '
        Me.lblMenuCaption.AutoSize = True
        Me.lblMenuCaption.Font = New System.Drawing.Font("Arial", 24.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblMenuCaption.Location = New System.Drawing.Point(238, 26)
        Me.lblMenuCaption.Name = "lblMenuCaption"
        Me.lblMenuCaption.Size = New System.Drawing.Size(251, 36)
        Me.lblMenuCaption.TabIndex = 19
        Me.lblMenuCaption.Text = "### Menu Name"
        '
        'lblMenuItem
        '
        Me.lblMenuItem.AutoSize = True
        Me.lblMenuItem.Font = New System.Drawing.Font("Arial Narrow", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblMenuItem.Location = New System.Drawing.Point(308, 213)
        Me.lblMenuItem.Name = "lblMenuItem"
        Me.lblMenuItem.Size = New System.Drawing.Size(109, 20)
        Me.lblMenuItem.TabIndex = 23
        Me.lblMenuItem.Text = "Menu Item Label"
        '
        'imgStoreLogoBorder
        '
        Me.imgStoreLogoBorder.AutoSize = True
        Me.imgStoreLogoBorder.Location = New System.Drawing.Point(339, 296)
        Me.imgStoreLogoBorder.Name = "imgStoreLogoBorder"
        Me.imgStoreLogoBorder.Size = New System.Drawing.Size(103, 13)
        Me.imgStoreLogoBorder.TabIndex = 27
        Me.imgStoreLogoBorder.Text = "imgStoreLogoBorder"
        '
        'bvb5
        '
        Me.bvb5.Image = Global.WinCDS.My.Resources.Resources.mInstallment_U
        Me.bvb5.Location = New System.Drawing.Point(10, 469)
        Me.bvb5.Name = "bvb5"
        Me.bvb5.Size = New System.Drawing.Size(160, 80)
        Me.bvb5.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize
        Me.bvb5.TabIndex = 35
        Me.bvb5.TabStop = False
        '
        'bvb4
        '
        Me.bvb4.Image = Global.WinCDS.My.Resources.Resources.mMailing_U
        Me.bvb4.Location = New System.Drawing.Point(10, 389)
        Me.bvb4.Name = "bvb4"
        Me.bvb4.Size = New System.Drawing.Size(160, 80)
        Me.bvb4.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize
        Me.bvb4.TabIndex = 34
        Me.bvb4.TabStop = False
        '
        'bvb3
        '
        Me.bvb3.Image = Global.WinCDS.My.Resources.Resources.mAccounting_U
        Me.bvb3.Location = New System.Drawing.Point(10, 309)
        Me.bvb3.Name = "bvb3"
        Me.bvb3.Size = New System.Drawing.Size(160, 80)
        Me.bvb3.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize
        Me.bvb3.TabIndex = 33
        Me.bvb3.TabStop = False
        '
        'bvb2
        '
        Me.bvb2.Image = Global.WinCDS.My.Resources.Resources.mInventory_U
        Me.bvb2.Location = New System.Drawing.Point(10, 229)
        Me.bvb2.Name = "bvb2"
        Me.bvb2.Size = New System.Drawing.Size(160, 80)
        Me.bvb2.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize
        Me.bvb2.TabIndex = 32
        Me.bvb2.TabStop = False
        '
        'bvb1
        '
        Me.bvb1.Image = Global.WinCDS.My.Resources.Resources.mOrder_U
        Me.bvb1.Location = New System.Drawing.Point(10, 149)
        Me.bvb1.Name = "bvb1"
        Me.bvb1.Size = New System.Drawing.Size(160, 80)
        Me.bvb1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize
        Me.bvb1.TabIndex = 31
        Me.bvb1.TabStop = False
        '
        'bvb0
        '
        Me.bvb0.BackColor = System.Drawing.SystemColors.Control
        Me.bvb0.Image = Global.WinCDS.My.Resources.Resources.mFile_U
        Me.bvb0.Location = New System.Drawing.Point(10, 69)
        Me.bvb0.Name = "bvb0"
        Me.bvb0.Size = New System.Drawing.Size(160, 80)
        Me.bvb0.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize
        Me.bvb0.TabIndex = 30
        Me.bvb0.TabStop = False
        '
        'cmdLogout
        '
        Me.cmdLogout.Location = New System.Drawing.Point(26, 4)
        Me.cmdLogout.Name = "cmdLogout"
        Me.cmdLogout.Size = New System.Drawing.Size(100, 50)
        Me.cmdLogout.TabIndex = 29
        Me.cmdLogout.TabStop = False
        '
        'imgStoreLogo
        '
        Me.imgStoreLogo.Location = New System.Drawing.Point(550, 328)
        Me.imgStoreLogo.Name = "imgStoreLogo"
        Me.imgStoreLogo.Size = New System.Drawing.Size(136, 84)
        Me.imgStoreLogo.TabIndex = 28
        Me.imgStoreLogo.TabStop = False
        '
        'imgHR
        '
        Me.imgHR.Location = New System.Drawing.Point(311, 350)
        Me.imgHR.Name = "imgHR"
        Me.imgHR.Size = New System.Drawing.Size(174, 36)
        Me.imgHR.TabIndex = 26
        Me.imgHR.TabStop = False
        '
        'imgSelected
        '
        Me.imgSelected.Location = New System.Drawing.Point(216, 180)
        Me.imgSelected.Name = "imgSelected"
        Me.imgSelected.Size = New System.Drawing.Size(64, 85)
        Me.imgSelected.TabIndex = 25
        Me.imgSelected.TabStop = False
        '
        'imgSubSelected
        '
        Me.imgSubSelected.Location = New System.Drawing.Point(286, 237)
        Me.imgSubSelected.Name = "imgSubSelected"
        Me.imgSubSelected.Size = New System.Drawing.Size(266, 28)
        Me.imgSubSelected.TabIndex = 24
        Me.imgSubSelected.TabStop = False
        '
        'imgMenuItem
        '
        Me.imgMenuItem.Location = New System.Drawing.Point(294, 91)
        Me.imgMenuItem.Name = "imgMenuItem"
        Me.imgMenuItem.Size = New System.Drawing.Size(103, 93)
        Me.imgMenuItem.TabIndex = 22
        Me.imgMenuItem.TabStop = False
        '
        'lblHR
        '
        Me.lblHR.AutoSize = True
        Me.lblHR.Font = New System.Drawing.Font("Arial Narrow", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblHR.Location = New System.Drawing.Point(328, 358)
        Me.lblHR.Name = "lblHR"
        Me.lblHR.Size = New System.Drawing.Size(88, 20)
        Me.lblHR.TabIndex = 36
        Me.lblHR.Text = "### HR Label"
        '
        'lblDevMode
        '
        Me.lblDevMode.AutoSize = True
        Me.lblDevMode.Font = New System.Drawing.Font("Lucida Console", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblDevMode.Location = New System.Drawing.Point(654, 4)
        Me.lblDevMode.Name = "lblDevMode"
        Me.lblDevMode.Size = New System.Drawing.Size(37, 11)
        Me.lblDevMode.TabIndex = 37
        Me.lblDevMode.Text = "DEV "
        '
        'lblCDSComputer
        '
        Me.lblCDSComputer.AutoSize = True
        Me.lblCDSComputer.Font = New System.Drawing.Font("Lucida Console", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblCDSComputer.Location = New System.Drawing.Point(687, 4)
        Me.lblCDSComputer.Name = "lblCDSComputer"
        Me.lblCDSComputer.Size = New System.Drawing.Size(37, 11)
        Me.lblCDSComputer.TabIndex = 38
        Me.lblCDSComputer.Text = "CDS "
        '
        'lblRDP
        '
        Me.lblRDP.AutoSize = True
        Me.lblRDP.Font = New System.Drawing.Font("Lucida Console", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblRDP.Location = New System.Drawing.Point(717, 4)
        Me.lblRDP.Name = "lblRDP"
        Me.lblRDP.Size = New System.Drawing.Size(37, 11)
        Me.lblRDP.TabIndex = 39
        Me.lblRDP.Text = "RDP "
        '
        'lblELEVATE
        '
        Me.lblELEVATE.AutoSize = True
        Me.lblELEVATE.Font = New System.Drawing.Font("Lucida Console", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblELEVATE.Location = New System.Drawing.Point(749, 4)
        Me.lblELEVATE.Name = "lblELEVATE"
        Me.lblELEVATE.Size = New System.Drawing.Size(37, 11)
        Me.lblELEVATE.TabIndex = 40
        Me.lblELEVATE.Text = "ELEV"
        '
        'lblBETA
        '
        Me.lblBETA.AutoSize = True
        Me.lblBETA.Font = New System.Drawing.Font("Lucida Console", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblBETA.Location = New System.Drawing.Point(782, 4)
        Me.lblBETA.Name = "lblBETA"
        Me.lblBETA.Size = New System.Drawing.Size(37, 11)
        Me.lblBETA.TabIndex = 41
        Me.lblBETA.Text = "BETA"
        '
        'txtYearEnd
        '
        Me.txtYearEnd.Location = New System.Drawing.Point(700, 237)
        Me.txtYearEnd.Name = "txtYearEnd"
        Me.txtYearEnd.Size = New System.Drawing.Size(100, 20)
        Me.txtYearEnd.TabIndex = 42
        '
        'txtInfo
        '
        Me.txtInfo.Location = New System.Drawing.Point(700, 263)
        Me.txtInfo.Name = "txtInfo"
        Me.txtInfo.Size = New System.Drawing.Size(100, 20)
        Me.txtInfo.TabIndex = 43
        Me.txtInfo.Text = "### INFO BOX"
        '
        'lblStore0
        '
        Me.lblStore0.AutoSize = True
        Me.lblStore0.Location = New System.Drawing.Point(220, 534)
        Me.lblStore0.Name = "lblStore0"
        Me.lblStore0.Size = New System.Drawing.Size(28, 13)
        Me.lblStore0.TabIndex = 44
        Me.lblStore0.Text = "###"
        '
        'lblStore1
        '
        Me.lblStore1.AutoSize = True
        Me.lblStore1.Location = New System.Drawing.Point(308, 534)
        Me.lblStore1.Name = "lblStore1"
        Me.lblStore1.Size = New System.Drawing.Size(28, 13)
        Me.lblStore1.TabIndex = 45
        Me.lblStore1.Text = "###"
        '
        'lblStore2
        '
        Me.lblStore2.AutoSize = True
        Me.lblStore2.Location = New System.Drawing.Point(414, 536)
        Me.lblStore2.Name = "lblStore2"
        Me.lblStore2.Size = New System.Drawing.Size(28, 13)
        Me.lblStore2.TabIndex = 46
        Me.lblStore2.Text = "###"
        '
        'KeyCatch
        '
        Me.KeyCatch.Location = New System.Drawing.Point(700, 293)
        Me.KeyCatch.Name = "KeyCatch"
        Me.KeyCatch.Size = New System.Drawing.Size(100, 20)
        Me.KeyCatch.TabIndex = 47
        '
        'imgInfo
        '
        Me.imgInfo.Location = New System.Drawing.Point(495, 445)
        Me.imgInfo.Name = "imgInfo"
        Me.imgInfo.Size = New System.Drawing.Size(120, 69)
        Me.imgInfo.TabIndex = 48
        Me.imgInfo.TabStop = False
        '
        'imgBackground
        '
        Me.imgBackground.Location = New System.Drawing.Point(661, 451)
        Me.imgBackground.Name = "imgBackground"
        Me.imgBackground.Size = New System.Drawing.Size(103, 62)
        Me.imgBackground.TabIndex = 49
        Me.imgBackground.TabStop = False
        '
        'tmrPulse
        '
        '
        'tmrMaintain
        '
        '
        'MainMenu4
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(800, 563)
        Me.Controls.Add(Me.imgBackground)
        Me.Controls.Add(Me.imgInfo)
        Me.Controls.Add(Me.KeyCatch)
        Me.Controls.Add(Me.lblStore2)
        Me.Controls.Add(Me.lblStore1)
        Me.Controls.Add(Me.lblStore0)
        Me.Controls.Add(Me.txtInfo)
        Me.Controls.Add(Me.txtYearEnd)
        Me.Controls.Add(Me.lblBETA)
        Me.Controls.Add(Me.lblELEVATE)
        Me.Controls.Add(Me.lblRDP)
        Me.Controls.Add(Me.lblCDSComputer)
        Me.Controls.Add(Me.lblDevMode)
        Me.Controls.Add(Me.lblHR)
        Me.Controls.Add(Me.bvb5)
        Me.Controls.Add(Me.bvb4)
        Me.Controls.Add(Me.bvb3)
        Me.Controls.Add(Me.bvb2)
        Me.Controls.Add(Me.bvb1)
        Me.Controls.Add(Me.bvb0)
        Me.Controls.Add(Me.cmdEnterPassword)
        Me.Controls.Add(Me.txtPassword)
        Me.Controls.Add(Me.cmdLogout)
        Me.Controls.Add(Me.imgStoreLogo)
        Me.Controls.Add(Me.imgStoreLogoBorder)
        Me.Controls.Add(Me.imgHR)
        Me.Controls.Add(Me.imgSelected)
        Me.Controls.Add(Me.imgSubSelected)
        Me.Controls.Add(Me.lblMenuItem)
        Me.Controls.Add(Me.imgMenuItem)
        Me.Controls.Add(Me.lblIDE)
        Me.Controls.Add(Me.lblWinCDS)
        Me.Controls.Add(Me.lblDEMO)
        Me.Controls.Add(Me.fraSupplies)
        Me.Controls.Add(Me.lblLastBackup)
        Me.Controls.Add(Me.lblMenuCaption)
        Me.Controls.Add(Me.Button1)
        Me.Name = "MainMenu4"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "WinCDS 2016"
        Me.fraSupplies.ResumeLayout(False)
        CType(Me.picAlpha, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.cdgFile, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.MSComm1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.imgPicture, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.bvb5, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.bvb4, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.bvb3, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.bvb2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.bvb1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.bvb0, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.cmdLogout, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.imgStoreLogo, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.imgHR, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.imgSelected, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.imgSubSelected, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.imgMenuItem, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.imgInfo, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.imgBackground, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
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
    Friend WithEvents fraSupplies As GroupBox
    Friend WithEvents cdgFile As AxMSComDlg.AxCommonDialog
    Friend WithEvents imlMM As ImageList
    Friend WithEvents picAlpha As PictureBox
    Friend WithEvents lblLastBackup As Label
    Friend WithEvents lblDEMO As Label
    Friend WithEvents lblIDE As Label
    Friend WithEvents lblWinCDS As Label
    Friend WithEvents lblMenuCaption As Label
    Friend WithEvents imgMenuItem As PictureBox
    Friend WithEvents lblMenuItem As Label
    Friend WithEvents imgSubSelected As PictureBox
    Friend WithEvents imgSelected As PictureBox
    Friend WithEvents imgHR As PictureBox
    Friend WithEvents imgStoreLogoBorder As Label
    Friend WithEvents imgStoreLogo As PictureBox
    Friend WithEvents cmdLogout As PictureBox
    Friend WithEvents bvb0 As PictureBox
    Friend WithEvents bvb1 As PictureBox
    Friend WithEvents bvb2 As PictureBox
    Friend WithEvents bvb3 As PictureBox
    Friend WithEvents bvb4 As PictureBox
    Friend WithEvents bvb5 As PictureBox
    Friend WithEvents lblHR As Label
    Friend WithEvents ttpMainMenu As ToolTip
    Friend WithEvents lblDevMode As Label
    Friend WithEvents lblCDSComputer As Label
    Friend WithEvents lblRDP As Label
    Friend WithEvents lblELEVATE As Label
    Friend WithEvents lblBETA As Label
    Friend WithEvents txtYearEnd As TextBox
    Friend WithEvents txtInfo As TextBox
    Friend WithEvents lblStore0 As Label
    Friend WithEvents lblStore1 As Label
    Friend WithEvents lblStore2 As Label
    Friend WithEvents KeyCatch As TextBox
    Friend WithEvents imgInfo As PictureBox
    Friend WithEvents imgBackground As PictureBox
    Friend WithEvents tmrPulse As Timer
    Friend WithEvents tmrMaintain As Timer
End Class
