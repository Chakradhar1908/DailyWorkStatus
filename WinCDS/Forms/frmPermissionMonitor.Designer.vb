<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmPermissionMonitor
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
        Me.optPerm = New System.Windows.Forms.RadioButton()
        Me.optStatus = New System.Windows.Forms.RadioButton()
        Me.optLog = New System.Windows.Forms.RadioButton()
        Me.optExtra = New System.Windows.Forms.RadioButton()
        Me.optFormList = New System.Windows.Forms.RadioButton()
        Me.optMemory = New System.Windows.Forms.RadioButton()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.txtPO = New System.Windows.Forms.TextBox()
        Me.txtMail = New System.Windows.Forms.TextBox()
        Me.fraStatus = New System.Windows.Forms.GroupBox()
        Me.txtReports = New System.Windows.Forms.TextBox()
        Me.txtAr = New System.Windows.Forms.TextBox()
        Me.txtInvent = New System.Windows.Forms.TextBox()
        Me.txtOrder = New System.Windows.Forms.TextBox()
        Me.lblPO = New System.Windows.Forms.Label()
        Me.lblMail = New System.Windows.Forms.Label()
        Me.lblReports = New System.Windows.Forms.Label()
        Me.lblAr = New System.Windows.Forms.Label()
        Me.lblInvent = New System.Windows.Forms.Label()
        Me.lblOrder = New System.Windows.Forms.Label()
        Me.cmdOK = New System.Windows.Forms.Button()
        Me.fraFormList = New System.Windows.Forms.GroupBox()
        Me.fraMemory = New System.Windows.Forms.GroupBox()
        Me.fraExtras = New System.Windows.Forms.GroupBox()
        Me.txtExtras = New System.Windows.Forms.TextBox()
        Me.txtMemory = New System.Windows.Forms.TextBox()
        Me.lstForms = New System.Windows.Forms.ListBox()
        Me.tmrFormList = New System.Windows.Forms.Timer(Me.components)
        Me.fraPerm = New System.Windows.Forms.GroupBox()
        Me.txtExpiry = New System.Windows.Forms.TextBox()
        Me.txtLastZone = New System.Windows.Forms.TextBox()
        Me.txtManagerGroups = New System.Windows.Forms.TextBox()
        Me.txtManagerName = New System.Windows.Forms.TextBox()
        Me.txtGroups = New System.Windows.Forms.TextBox()
        Me.txtUser = New System.Windows.Forms.TextBox()
        Me.lblExpiry = New System.Windows.Forms.Label()
        Me.lblLastZone = New System.Windows.Forms.Label()
        Me.lblManagerGroups = New System.Windows.Forms.Label()
        Me.lblManagerName = New System.Windows.Forms.Label()
        Me.lblGroups = New System.Windows.Forms.Label()
        Me.lblUser = New System.Windows.Forms.Label()
        Me.tmrExtra = New System.Windows.Forms.Timer(Me.components)
        Me.fraLog = New System.Windows.Forms.GroupBox()
        Me.lblLogType = New System.Windows.Forms.Label()
        Me.cmbLogType = New System.Windows.Forms.ComboBox()
        Me.lblLogLvl = New System.Windows.Forms.Label()
        Me.txtLogLvl = New System.Windows.Forms.TextBox()
        Me.lblBot = New System.Windows.Forms.Label()
        Me.lblTop = New System.Windows.Forms.Label()
        Me.chkLogTS = New System.Windows.Forms.CheckBox()
        Me.lblMaxLogLines = New System.Windows.Forms.Label()
        Me.txtMaxLogLines = New System.Windows.Forms.TextBox()
        Me.txtLog = New System.Windows.Forms.TextBox()
        Me.tmrMemory = New System.Windows.Forms.Timer(Me.components)
        Me.tmrPerm = New System.Windows.Forms.Timer(Me.components)
        Me.fraStatus.SuspendLayout()
        Me.fraFormList.SuspendLayout()
        Me.fraMemory.SuspendLayout()
        Me.fraExtras.SuspendLayout()
        Me.fraPerm.SuspendLayout()
        Me.fraLog.SuspendLayout()
        Me.SuspendLayout()
        '
        'optPerm
        '
        Me.optPerm.AutoSize = True
        Me.optPerm.Checked = True
        Me.optPerm.Location = New System.Drawing.Point(12, 8)
        Me.optPerm.Name = "optPerm"
        Me.optPerm.Size = New System.Drawing.Size(32, 17)
        Me.optPerm.TabIndex = 0
        Me.optPerm.TabStop = True
        Me.optPerm.Text = "&P"
        Me.ToolTip1.SetToolTip(Me.optPerm, "Permissions")
        Me.optPerm.UseVisualStyleBackColor = True
        '
        'optStatus
        '
        Me.optStatus.AutoSize = True
        Me.optStatus.Location = New System.Drawing.Point(44, 8)
        Me.optStatus.Name = "optStatus"
        Me.optStatus.Size = New System.Drawing.Size(32, 17)
        Me.optStatus.TabIndex = 1
        Me.optStatus.Text = "&S"
        Me.ToolTip1.SetToolTip(Me.optStatus, "Status Modes")
        Me.optStatus.UseVisualStyleBackColor = True
        '
        'optLog
        '
        Me.optLog.AutoSize = True
        Me.optLog.Location = New System.Drawing.Point(76, 8)
        Me.optLog.Name = "optLog"
        Me.optLog.Size = New System.Drawing.Size(31, 17)
        Me.optLog.TabIndex = 2
        Me.optLog.Text = "&L"
        Me.ToolTip1.SetToolTip(Me.optLog, "Log")
        Me.optLog.UseVisualStyleBackColor = True
        '
        'optExtra
        '
        Me.optExtra.AutoSize = True
        Me.optExtra.Location = New System.Drawing.Point(107, 8)
        Me.optExtra.Name = "optExtra"
        Me.optExtra.Size = New System.Drawing.Size(32, 17)
        Me.optExtra.TabIndex = 3
        Me.optExtra.Text = "&X"
        Me.ToolTip1.SetToolTip(Me.optExtra, "DeveloperEX")
        Me.optExtra.UseVisualStyleBackColor = True
        '
        'optFormList
        '
        Me.optFormList.AutoSize = True
        Me.optFormList.Location = New System.Drawing.Point(139, 8)
        Me.optFormList.Name = "optFormList"
        Me.optFormList.Size = New System.Drawing.Size(31, 17)
        Me.optFormList.TabIndex = 4
        Me.optFormList.Text = "&F"
        Me.ToolTip1.SetToolTip(Me.optFormList, "DeveloperEX")
        Me.optFormList.UseVisualStyleBackColor = True
        '
        'optMemory
        '
        Me.optMemory.AutoSize = True
        Me.optMemory.Location = New System.Drawing.Point(170, 8)
        Me.optMemory.Name = "optMemory"
        Me.optMemory.Size = New System.Drawing.Size(34, 17)
        Me.optMemory.TabIndex = 5
        Me.optMemory.Text = "&M"
        Me.ToolTip1.SetToolTip(Me.optMemory, "DeveloperEX")
        Me.optMemory.UseVisualStyleBackColor = True
        '
        'txtPO
        '
        Me.txtPO.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.txtPO.Location = New System.Drawing.Point(87, 162)
        Me.txtPO.Name = "txtPO"
        Me.txtPO.Size = New System.Drawing.Size(100, 22)
        Me.txtPO.TabIndex = 23
        Me.ToolTip1.SetToolTip(Me.txtPO, "MainMenu.PurchaseOrder")
        '
        'txtMail
        '
        Me.txtMail.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.txtMail.Location = New System.Drawing.Point(87, 134)
        Me.txtMail.Name = "txtMail"
        Me.txtMail.Size = New System.Drawing.Size(100, 22)
        Me.txtMail.TabIndex = 22
        Me.ToolTip1.SetToolTip(Me.txtMail, "MainMenu.Mail")
        '
        'fraStatus
        '
        Me.fraStatus.Controls.Add(Me.txtPO)
        Me.fraStatus.Controls.Add(Me.txtMail)
        Me.fraStatus.Controls.Add(Me.txtReports)
        Me.fraStatus.Controls.Add(Me.txtAr)
        Me.fraStatus.Controls.Add(Me.txtInvent)
        Me.fraStatus.Controls.Add(Me.txtOrder)
        Me.fraStatus.Controls.Add(Me.lblPO)
        Me.fraStatus.Controls.Add(Me.lblMail)
        Me.fraStatus.Controls.Add(Me.lblReports)
        Me.fraStatus.Controls.Add(Me.lblAr)
        Me.fraStatus.Controls.Add(Me.lblInvent)
        Me.fraStatus.Controls.Add(Me.lblOrder)
        Me.fraStatus.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.fraStatus.Location = New System.Drawing.Point(12, 31)
        Me.fraStatus.Name = "fraStatus"
        Me.fraStatus.Size = New System.Drawing.Size(200, 194)
        Me.fraStatus.TabIndex = 6
        Me.fraStatus.TabStop = False
        Me.fraStatus.Text = "Program Status:"
        '
        'txtReports
        '
        Me.txtReports.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.txtReports.Location = New System.Drawing.Point(87, 106)
        Me.txtReports.Name = "txtReports"
        Me.txtReports.Size = New System.Drawing.Size(100, 22)
        Me.txtReports.TabIndex = 21
        '
        'txtAr
        '
        Me.txtAr.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.txtAr.Location = New System.Drawing.Point(87, 78)
        Me.txtAr.Name = "txtAr"
        Me.txtAr.Size = New System.Drawing.Size(100, 22)
        Me.txtAr.TabIndex = 20
        '
        'txtInvent
        '
        Me.txtInvent.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.txtInvent.Location = New System.Drawing.Point(87, 50)
        Me.txtInvent.Name = "txtInvent"
        Me.txtInvent.Size = New System.Drawing.Size(100, 22)
        Me.txtInvent.TabIndex = 19
        '
        'txtOrder
        '
        Me.txtOrder.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.txtOrder.Location = New System.Drawing.Point(87, 22)
        Me.txtOrder.Name = "txtOrder"
        Me.txtOrder.Size = New System.Drawing.Size(100, 22)
        Me.txtOrder.TabIndex = 18
        '
        'lblPO
        '
        Me.lblPO.AutoSize = True
        Me.lblPO.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblPO.Location = New System.Drawing.Point(9, 166)
        Me.lblPO.Name = "lblPO"
        Me.lblPO.Size = New System.Drawing.Size(25, 13)
        Me.lblPO.TabIndex = 17
        Me.lblPO.Text = "&PO:"
        '
        'lblMail
        '
        Me.lblMail.AutoSize = True
        Me.lblMail.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblMail.Location = New System.Drawing.Point(9, 137)
        Me.lblMail.Name = "lblMail"
        Me.lblMail.Size = New System.Drawing.Size(29, 13)
        Me.lblMail.TabIndex = 16
        Me.lblMail.Text = "&Mail:"
        '
        'lblReports
        '
        Me.lblReports.AutoSize = True
        Me.lblReports.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblReports.Location = New System.Drawing.Point(9, 110)
        Me.lblReports.Name = "lblReports"
        Me.lblReports.Size = New System.Drawing.Size(47, 13)
        Me.lblReports.TabIndex = 15
        Me.lblReports.Text = "&Reports:"
        '
        'lblAr
        '
        Me.lblAr.AutoSize = True
        Me.lblAr.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblAr.Location = New System.Drawing.Point(9, 84)
        Me.lblAr.Name = "lblAr"
        Me.lblAr.Size = New System.Drawing.Size(50, 13)
        Me.lblAr.TabIndex = 14
        Me.lblAr.Text = "&ArSelect:"
        '
        'lblInvent
        '
        Me.lblInvent.AutoSize = True
        Me.lblInvent.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblInvent.Location = New System.Drawing.Point(9, 54)
        Me.lblInvent.Name = "lblInvent"
        Me.lblInvent.Size = New System.Drawing.Size(37, 13)
        Me.lblInvent.TabIndex = 13
        Me.lblInvent.Text = "&Inven:"
        '
        'lblOrder
        '
        Me.lblOrder.AutoSize = True
        Me.lblOrder.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblOrder.Location = New System.Drawing.Point(9, 27)
        Me.lblOrder.Name = "lblOrder"
        Me.lblOrder.Size = New System.Drawing.Size(36, 13)
        Me.lblOrder.TabIndex = 12
        Me.lblOrder.Text = "&Order:"
        '
        'cmdOK
        '
        Me.cmdOK.Location = New System.Drawing.Point(74, 231)
        Me.cmdOK.Name = "cmdOK"
        Me.cmdOK.Size = New System.Drawing.Size(75, 23)
        Me.cmdOK.TabIndex = 7
        Me.cmdOK.Text = "&Close"
        Me.cmdOK.UseVisualStyleBackColor = True
        '
        'fraFormList
        '
        Me.fraFormList.Controls.Add(Me.lstForms)
        Me.fraFormList.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.fraFormList.Location = New System.Drawing.Point(12, 31)
        Me.fraFormList.Name = "fraFormList"
        Me.fraFormList.Size = New System.Drawing.Size(200, 149)
        Me.fraFormList.TabIndex = 8
        Me.fraFormList.TabStop = False
        Me.fraFormList.Text = "Form List:"
        '
        'fraMemory
        '
        Me.fraMemory.Controls.Add(Me.txtMemory)
        Me.fraMemory.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.fraMemory.Location = New System.Drawing.Point(12, 31)
        Me.fraMemory.Name = "fraMemory"
        Me.fraMemory.Size = New System.Drawing.Size(200, 157)
        Me.fraMemory.TabIndex = 9
        Me.fraMemory.TabStop = False
        Me.fraMemory.Text = "Memory Summary:"
        '
        'fraExtras
        '
        Me.fraExtras.Controls.Add(Me.txtExtras)
        Me.fraExtras.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.fraExtras.Location = New System.Drawing.Point(12, 34)
        Me.fraExtras.Name = "fraExtras"
        Me.fraExtras.Size = New System.Drawing.Size(200, 144)
        Me.fraExtras.TabIndex = 9
        Me.fraExtras.TabStop = False
        Me.fraExtras.Text = "Developer Extras:"
        '
        'txtExtras
        '
        Me.txtExtras.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.txtExtras.Location = New System.Drawing.Point(6, 21)
        Me.txtExtras.Multiline = True
        Me.txtExtras.Name = "txtExtras"
        Me.txtExtras.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.txtExtras.Size = New System.Drawing.Size(188, 117)
        Me.txtExtras.TabIndex = 0
        '
        'txtMemory
        '
        Me.txtMemory.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtMemory.Location = New System.Drawing.Point(6, 19)
        Me.txtMemory.Multiline = True
        Me.txtMemory.Name = "txtMemory"
        Me.txtMemory.Size = New System.Drawing.Size(188, 128)
        Me.txtMemory.TabIndex = 0
        '
        'lstForms
        '
        Me.lstForms.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lstForms.FormattingEnabled = True
        Me.lstForms.Location = New System.Drawing.Point(11, 19)
        Me.lstForms.Name = "lstForms"
        Me.lstForms.Size = New System.Drawing.Size(179, 121)
        Me.lstForms.TabIndex = 0
        '
        'fraPerm
        '
        Me.fraPerm.Controls.Add(Me.txtExpiry)
        Me.fraPerm.Controls.Add(Me.txtLastZone)
        Me.fraPerm.Controls.Add(Me.txtManagerGroups)
        Me.fraPerm.Controls.Add(Me.txtManagerName)
        Me.fraPerm.Controls.Add(Me.txtGroups)
        Me.fraPerm.Controls.Add(Me.txtUser)
        Me.fraPerm.Controls.Add(Me.lblExpiry)
        Me.fraPerm.Controls.Add(Me.lblLastZone)
        Me.fraPerm.Controls.Add(Me.lblManagerGroups)
        Me.fraPerm.Controls.Add(Me.lblManagerName)
        Me.fraPerm.Controls.Add(Me.lblGroups)
        Me.fraPerm.Controls.Add(Me.lblUser)
        Me.fraPerm.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.fraPerm.Location = New System.Drawing.Point(14, 31)
        Me.fraPerm.Name = "fraPerm"
        Me.fraPerm.Size = New System.Drawing.Size(196, 152)
        Me.fraPerm.TabIndex = 9
        Me.fraPerm.TabStop = False
        Me.fraPerm.Text = "&Current Security Levels:"
        '
        'txtExpiry
        '
        Me.txtExpiry.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.txtExpiry.Location = New System.Drawing.Point(86, 145)
        Me.txtExpiry.Name = "txtExpiry"
        Me.txtExpiry.Size = New System.Drawing.Size(100, 22)
        Me.txtExpiry.TabIndex = 11
        '
        'txtLastZone
        '
        Me.txtLastZone.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.txtLastZone.Location = New System.Drawing.Point(86, 119)
        Me.txtLastZone.Name = "txtLastZone"
        Me.txtLastZone.Size = New System.Drawing.Size(100, 22)
        Me.txtLastZone.TabIndex = 10
        '
        'txtManagerGroups
        '
        Me.txtManagerGroups.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.txtManagerGroups.Location = New System.Drawing.Point(86, 94)
        Me.txtManagerGroups.Name = "txtManagerGroups"
        Me.txtManagerGroups.Size = New System.Drawing.Size(100, 22)
        Me.txtManagerGroups.TabIndex = 9
        '
        'txtManagerName
        '
        Me.txtManagerName.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.txtManagerName.Location = New System.Drawing.Point(86, 72)
        Me.txtManagerName.Name = "txtManagerName"
        Me.txtManagerName.Size = New System.Drawing.Size(100, 22)
        Me.txtManagerName.TabIndex = 8
        '
        'txtGroups
        '
        Me.txtGroups.Location = New System.Drawing.Point(86, 44)
        Me.txtGroups.Name = "txtGroups"
        Me.txtGroups.Size = New System.Drawing.Size(100, 22)
        Me.txtGroups.TabIndex = 7
        '
        'txtUser
        '
        Me.txtUser.Location = New System.Drawing.Point(86, 19)
        Me.txtUser.Name = "txtUser"
        Me.txtUser.Size = New System.Drawing.Size(100, 22)
        Me.txtUser.TabIndex = 6
        '
        'lblExpiry
        '
        Me.lblExpiry.AutoSize = True
        Me.lblExpiry.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblExpiry.Location = New System.Drawing.Point(10, 149)
        Me.lblExpiry.Name = "lblExpiry"
        Me.lblExpiry.Size = New System.Drawing.Size(38, 13)
        Me.lblExpiry.TabIndex = 5
        Me.lblExpiry.Text = "Expir&y:"
        '
        'lblLastZone
        '
        Me.lblLastZone.AutoSize = True
        Me.lblLastZone.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblLastZone.Location = New System.Drawing.Point(10, 127)
        Me.lblLastZone.Name = "lblLastZone"
        Me.lblLastZone.Size = New System.Drawing.Size(58, 13)
        Me.lblLastZone.TabIndex = 4
        Me.lblLastZone.Text = "Last Zon&e:"
        '
        'lblManagerGroups
        '
        Me.lblManagerGroups.AutoSize = True
        Me.lblManagerGroups.BackColor = System.Drawing.SystemColors.Control
        Me.lblManagerGroups.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblManagerGroups.ForeColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.lblManagerGroups.Location = New System.Drawing.Point(10, 99)
        Me.lblManagerGroups.Name = "lblManagerGroups"
        Me.lblManagerGroups.Size = New System.Drawing.Size(71, 13)
        Me.lblManagerGroups.TabIndex = 3
        Me.lblManagerGroups.Text = "Mngr Gro&ups:"
        '
        'lblManagerName
        '
        Me.lblManagerName.AutoSize = True
        Me.lblManagerName.BackColor = System.Drawing.SystemColors.Control
        Me.lblManagerName.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblManagerName.ForeColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.lblManagerName.Location = New System.Drawing.Point(10, 75)
        Me.lblManagerName.Name = "lblManagerName"
        Me.lblManagerName.Size = New System.Drawing.Size(52, 13)
        Me.lblManagerName.TabIndex = 2
        Me.lblManagerName.Text = "Mana&ger:"
        '
        'lblGroups
        '
        Me.lblGroups.AutoSize = True
        Me.lblGroups.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblGroups.Location = New System.Drawing.Point(10, 50)
        Me.lblGroups.Name = "lblGroups"
        Me.lblGroups.Size = New System.Drawing.Size(44, 13)
        Me.lblGroups.TabIndex = 1
        Me.lblGroups.Text = "Grou&ps:"
        '
        'lblUser
        '
        Me.lblUser.AutoSize = True
        Me.lblUser.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblUser.Location = New System.Drawing.Point(10, 27)
        Me.lblUser.Name = "lblUser"
        Me.lblUser.Size = New System.Drawing.Size(63, 13)
        Me.lblUser.TabIndex = 0
        Me.lblUser.Text = "User Na&me:"
        '
        'fraLog
        '
        Me.fraLog.Controls.Add(Me.txtLog)
        Me.fraLog.Controls.Add(Me.txtMaxLogLines)
        Me.fraLog.Controls.Add(Me.lblMaxLogLines)
        Me.fraLog.Controls.Add(Me.chkLogTS)
        Me.fraLog.Controls.Add(Me.lblTop)
        Me.fraLog.Controls.Add(Me.lblBot)
        Me.fraLog.Controls.Add(Me.txtLogLvl)
        Me.fraLog.Controls.Add(Me.lblLogLvl)
        Me.fraLog.Controls.Add(Me.cmbLogType)
        Me.fraLog.Controls.Add(Me.lblLogType)
        Me.fraLog.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.fraLog.Location = New System.Drawing.Point(12, 35)
        Me.fraLog.Name = "fraLog"
        Me.fraLog.Size = New System.Drawing.Size(198, 165)
        Me.fraLog.TabIndex = 9
        Me.fraLog.TabStop = False
        Me.fraLog.Text = "&Active Log"
        '
        'lblLogType
        '
        Me.lblLogType.AutoSize = True
        Me.lblLogType.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblLogType.Location = New System.Drawing.Point(6, 20)
        Me.lblLogType.Name = "lblLogType"
        Me.lblLogType.Size = New System.Drawing.Size(34, 13)
        Me.lblLogType.TabIndex = 0
        Me.lblLogType.Text = "T&ype:"
        '
        'cmbLogType
        '
        Me.cmbLogType.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmbLogType.FormattingEnabled = True
        Me.cmbLogType.Items.AddRange(New Object() {"All", "General"})
        Me.cmbLogType.Location = New System.Drawing.Point(38, 17)
        Me.cmbLogType.Name = "cmbLogType"
        Me.cmbLogType.Size = New System.Drawing.Size(150, 21)
        Me.cmbLogType.TabIndex = 1
        Me.cmbLogType.Text = "All"
        '
        'lblLogLvl
        '
        Me.lblLogLvl.AutoSize = True
        Me.lblLogLvl.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblLogLvl.Location = New System.Drawing.Point(6, 43)
        Me.lblLogLvl.Name = "lblLogLvl"
        Me.lblLogLvl.Size = New System.Drawing.Size(36, 13)
        Me.lblLogLvl.TabIndex = 2
        Me.lblLogLvl.Text = "Le&vel:"
        '
        'txtLogLvl
        '
        Me.txtLogLvl.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtLogLvl.Location = New System.Drawing.Point(38, 42)
        Me.txtLogLvl.Name = "txtLogLvl"
        Me.txtLogLvl.Size = New System.Drawing.Size(18, 20)
        Me.txtLogLvl.TabIndex = 3
        Me.txtLogLvl.Text = "9"
        '
        'lblBot
        '
        Me.lblBot.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.lblBot.Location = New System.Drawing.Point(59, 40)
        Me.lblBot.Name = "lblBot"
        Me.lblBot.Size = New System.Drawing.Size(14, 10)
        Me.lblBot.TabIndex = 4
        '
        'lblTop
        '
        Me.lblTop.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.lblTop.Location = New System.Drawing.Point(75, 40)
        Me.lblTop.Name = "lblTop"
        Me.lblTop.Size = New System.Drawing.Size(14, 10)
        Me.lblTop.TabIndex = 5
        '
        'chkLogTS
        '
        Me.chkLogTS.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkLogTS.Location = New System.Drawing.Point(60, 49)
        Me.chkLogTS.Name = "chkLogTS"
        Me.chkLogTS.Size = New System.Drawing.Size(30, 20)
        Me.chkLogTS.TabIndex = 6
        Me.chkLogTS.Text = "T"
        Me.chkLogTS.UseVisualStyleBackColor = True
        '
        'lblMaxLogLines
        '
        Me.lblMaxLogLines.AutoSize = True
        Me.lblMaxLogLines.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblMaxLogLines.Location = New System.Drawing.Point(137, 42)
        Me.lblMaxLogLines.Name = "lblMaxLogLines"
        Me.lblMaxLogLines.Size = New System.Drawing.Size(19, 13)
        Me.lblMaxLogLines.TabIndex = 7
        Me.lblMaxLogLines.Text = "M:"
        '
        'txtMaxLogLines
        '
        Me.txtMaxLogLines.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtMaxLogLines.Location = New System.Drawing.Point(156, 39)
        Me.txtMaxLogLines.Name = "txtMaxLogLines"
        Me.txtMaxLogLines.Size = New System.Drawing.Size(31, 20)
        Me.txtMaxLogLines.TabIndex = 8
        Me.txtMaxLogLines.Text = "0000"
        '
        'txtLog
        '
        Me.txtLog.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtLog.Location = New System.Drawing.Point(6, 69)
        Me.txtLog.Multiline = True
        Me.txtLog.Name = "txtLog"
        Me.txtLog.ScrollBars = System.Windows.Forms.ScrollBars.Both
        Me.txtLog.Size = New System.Drawing.Size(184, 90)
        Me.txtLog.TabIndex = 9
        '
        'frmPermissionMonitor
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(219, 259)
        Me.Controls.Add(Me.fraStatus)
        Me.Controls.Add(Me.fraLog)
        Me.Controls.Add(Me.fraPerm)
        Me.Controls.Add(Me.fraMemory)
        Me.Controls.Add(Me.fraExtras)
        Me.Controls.Add(Me.fraFormList)
        Me.Controls.Add(Me.cmdOK)
        Me.Controls.Add(Me.optMemory)
        Me.Controls.Add(Me.optFormList)
        Me.Controls.Add(Me.optExtra)
        Me.Controls.Add(Me.optLog)
        Me.Controls.Add(Me.optStatus)
        Me.Controls.Add(Me.optPerm)
        Me.Name = "frmPermissionMonitor"
        Me.Text = "Permission Mon"
        Me.fraStatus.ResumeLayout(False)
        Me.fraStatus.PerformLayout()
        Me.fraFormList.ResumeLayout(False)
        Me.fraMemory.ResumeLayout(False)
        Me.fraMemory.PerformLayout()
        Me.fraExtras.ResumeLayout(False)
        Me.fraExtras.PerformLayout()
        Me.fraPerm.ResumeLayout(False)
        Me.fraPerm.PerformLayout()
        Me.fraLog.ResumeLayout(False)
        Me.fraLog.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents optPerm As RadioButton
    Friend WithEvents optStatus As RadioButton
    Friend WithEvents optLog As RadioButton
    Friend WithEvents optExtra As RadioButton
    Friend WithEvents optFormList As RadioButton
    Friend WithEvents optMemory As RadioButton
    Friend WithEvents ToolTip1 As ToolTip
    Friend WithEvents fraStatus As GroupBox
    Friend WithEvents txtPO As TextBox
    Friend WithEvents txtMail As TextBox
    Friend WithEvents txtReports As TextBox
    Friend WithEvents txtAr As TextBox
    Friend WithEvents txtInvent As TextBox
    Friend WithEvents txtOrder As TextBox
    Friend WithEvents lblPO As Label
    Friend WithEvents lblMail As Label
    Friend WithEvents lblReports As Label
    Friend WithEvents lblAr As Label
    Friend WithEvents lblInvent As Label
    Friend WithEvents lblOrder As Label
    Friend WithEvents cmdOK As Button
    Friend WithEvents fraFormList As GroupBox
    Friend WithEvents lstForms As ListBox
    Friend WithEvents tmrFormList As Timer
    Friend WithEvents fraMemory As GroupBox
    Friend WithEvents fraExtras As GroupBox
    Friend WithEvents txtExtras As TextBox
    Friend WithEvents txtMemory As TextBox
    Friend WithEvents fraPerm As GroupBox
    Friend WithEvents txtExpiry As TextBox
    Friend WithEvents txtLastZone As TextBox
    Friend WithEvents txtManagerGroups As TextBox
    Friend WithEvents txtManagerName As TextBox
    Friend WithEvents txtGroups As TextBox
    Friend WithEvents txtUser As TextBox
    Friend WithEvents lblExpiry As Label
    Friend WithEvents lblLastZone As Label
    Friend WithEvents lblManagerGroups As Label
    Friend WithEvents lblManagerName As Label
    Friend WithEvents lblGroups As Label
    Friend WithEvents lblUser As Label
    Friend WithEvents tmrExtra As Timer
    Friend WithEvents fraLog As GroupBox
    Friend WithEvents txtLog As TextBox
    Friend WithEvents txtMaxLogLines As TextBox
    Friend WithEvents lblMaxLogLines As Label
    Friend WithEvents chkLogTS As CheckBox
    Friend WithEvents lblTop As Label
    Friend WithEvents lblBot As Label
    Friend WithEvents txtLogLvl As TextBox
    Friend WithEvents lblLogLvl As Label
    Friend WithEvents cmbLogType As ComboBox
    Friend WithEvents lblLogType As Label
    Friend WithEvents tmrMemory As Timer
    Friend WithEvents tmrPerm As Timer
End Class
