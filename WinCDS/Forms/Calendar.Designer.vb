<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Calendar
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Calendar))
        Me.fraButtons = New System.Windows.Forms.GroupBox()
        Me.chkMultiple = New System.Windows.Forms.CheckBox()
        Me.cmdDDT = New System.Windows.Forms.Button()
        Me.cmdMap = New System.Windows.Forms.Button()
        Me.cmdInstr = New System.Windows.Forms.Button()
        Me.cmdPrint = New System.Windows.Forms.Button()
        Me.cmdManifest = New System.Windows.Forms.Button()
        Me.cmdMenu = New System.Windows.Forms.Button()
        Me.lblDayLabel = New System.Windows.Forms.Label()
        Me.txtDayLabel = New System.Windows.Forms.TextBox()
        Me.cmdApply = New System.Windows.Forms.Button()
        Me.grid = New AxMSFlexGridLib.AxMSFlexGrid()
        Me.lblCubes = New System.Windows.Forms.Label()
        Me.fraButtons.SuspendLayout()
        CType(Me.grid, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'fraButtons
        '
        Me.fraButtons.Controls.Add(Me.chkMultiple)
        Me.fraButtons.Controls.Add(Me.cmdDDT)
        Me.fraButtons.Controls.Add(Me.cmdMap)
        Me.fraButtons.Controls.Add(Me.cmdInstr)
        Me.fraButtons.Controls.Add(Me.cmdPrint)
        Me.fraButtons.Controls.Add(Me.cmdManifest)
        Me.fraButtons.Controls.Add(Me.cmdMenu)
        Me.fraButtons.Location = New System.Drawing.Point(12, 366)
        Me.fraButtons.Name = "fraButtons"
        Me.fraButtons.Size = New System.Drawing.Size(649, 86)
        Me.fraButtons.TabIndex = 2
        Me.fraButtons.TabStop = False
        '
        'chkMultiple
        '
        Me.chkMultiple.Location = New System.Drawing.Point(570, 16)
        Me.chkMultiple.Name = "chkMultiple"
        Me.chkMultiple.Size = New System.Drawing.Size(68, 48)
        Me.chkMultiple.TabIndex = 6
        Me.chkMultiple.Text = "Route Multiple Stores"
        Me.chkMultiple.UseVisualStyleBackColor = True
        '
        'cmdDDT
        '
        Me.cmdDDT.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cmdDDT.Location = New System.Drawing.Point(471, 16)
        Me.cmdDDT.Name = "cmdDDT"
        Me.cmdDDT.Size = New System.Drawing.Size(93, 62)
        Me.cmdDDT.TabIndex = 5
        Me.cmdDDT.Text = "&Dispatch"
        Me.cmdDDT.UseVisualStyleBackColor = True
        '
        'cmdMap
        '
        Me.cmdMap.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cmdMap.Location = New System.Drawing.Point(378, 16)
        Me.cmdMap.Name = "cmdMap"
        Me.cmdMap.Size = New System.Drawing.Size(93, 62)
        Me.cmdMap.TabIndex = 4
        Me.cmdMap.Text = "M&aps"
        Me.cmdMap.UseVisualStyleBackColor = True
        '
        'cmdInstr
        '
        Me.cmdInstr.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cmdInstr.Location = New System.Drawing.Point(285, 16)
        Me.cmdInstr.Name = "cmdInstr"
        Me.cmdInstr.Size = New System.Drawing.Size(93, 62)
        Me.cmdInstr.TabIndex = 3
        Me.cmdInstr.Text = "&View Sp Instr"
        Me.cmdInstr.UseVisualStyleBackColor = True
        '
        'cmdPrint
        '
        Me.cmdPrint.Location = New System.Drawing.Point(192, 16)
        Me.cmdPrint.Name = "cmdPrint"
        Me.cmdPrint.Size = New System.Drawing.Size(93, 62)
        Me.cmdPrint.TabIndex = 2
        Me.cmdPrint.Text = "&Print"
        Me.cmdPrint.UseVisualStyleBackColor = True
        '
        'cmdManifest
        '
        Me.cmdManifest.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cmdManifest.Location = New System.Drawing.Point(99, 16)
        Me.cmdManifest.Name = "cmdManifest"
        Me.cmdManifest.Size = New System.Drawing.Size(93, 62)
        Me.cmdManifest.TabIndex = 1
        Me.cmdManifest.Text = "Print Mani&fest"
        Me.cmdManifest.UseVisualStyleBackColor = True
        '
        'cmdMenu
        '
        Me.cmdMenu.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cmdMenu.Location = New System.Drawing.Point(6, 16)
        Me.cmdMenu.Name = "cmdMenu"
        Me.cmdMenu.Size = New System.Drawing.Size(93, 62)
        Me.cmdMenu.TabIndex = 0
        Me.cmdMenu.Text = "&Menu"
        Me.cmdMenu.UseVisualStyleBackColor = True
        '
        'lblDayLabel
        '
        Me.lblDayLabel.AutoSize = True
        Me.lblDayLabel.Location = New System.Drawing.Point(667, 375)
        Me.lblDayLabel.Name = "lblDayLabel"
        Me.lblDayLabel.Size = New System.Drawing.Size(245, 13)
        Me.lblDayLabel.TabIndex = 3
        Me.lblDayLabel.Text = "Click column heading (date) to enter delivery zone:"
        '
        'txtDayLabel
        '
        Me.txtDayLabel.Location = New System.Drawing.Point(670, 391)
        Me.txtDayLabel.Name = "txtDayLabel"
        Me.txtDayLabel.Size = New System.Drawing.Size(247, 20)
        Me.txtDayLabel.TabIndex = 4
        '
        'cmdApply
        '
        Me.cmdApply.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cmdApply.Location = New System.Drawing.Point(920, 388)
        Me.cmdApply.Name = "cmdApply"
        Me.cmdApply.Size = New System.Drawing.Size(46, 23)
        Me.cmdApply.TabIndex = 5
        Me.cmdApply.Text = "Apply"
        Me.cmdApply.UseVisualStyleBackColor = True
        '
        'grid
        '
        Me.grid.Location = New System.Drawing.Point(12, 12)
        Me.grid.Name = "grid"
        Me.grid.OcxState = CType(resources.GetObject("grid.OcxState"), System.Windows.Forms.AxHost.State)
        Me.grid.Size = New System.Drawing.Size(958, 348)
        Me.grid.TabIndex = 1
        '
        'lblCubes
        '
        Me.lblCubes.Location = New System.Drawing.Point(599, 425)
        Me.lblCubes.Name = "lblCubes"
        Me.lblCubes.Size = New System.Drawing.Size(117, 19)
        Me.lblCubes.TabIndex = 7
        '
        'Calendar
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.cmdMenu
        Me.ClientSize = New System.Drawing.Size(982, 462)
        Me.Controls.Add(Me.cmdApply)
        Me.Controls.Add(Me.txtDayLabel)
        Me.Controls.Add(Me.lblDayLabel)
        Me.Controls.Add(Me.fraButtons)
        Me.Controls.Add(Me.grid)
        Me.Controls.Add(Me.lblCubes)
        Me.Name = "Calendar"
        Me.Text = "Delivery Calendar - Next 31 Days"
        Me.fraButtons.ResumeLayout(False)
        CType(Me.grid, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents grid As AxMSFlexGridLib.AxMSFlexGrid
    Friend WithEvents fraButtons As GroupBox
    Friend WithEvents chkMultiple As CheckBox
    Friend WithEvents cmdDDT As Button
    Friend WithEvents cmdMap As Button
    Friend WithEvents cmdInstr As Button
    Friend WithEvents cmdPrint As Button
    Friend WithEvents cmdManifest As Button
    Friend WithEvents cmdMenu As Button
    Friend WithEvents lblDayLabel As Label
    Friend WithEvents txtDayLabel As TextBox
    Friend WithEvents cmdApply As Button
    Friend WithEvents lblCubes As Label
End Class
