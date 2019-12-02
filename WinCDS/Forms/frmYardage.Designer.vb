<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmYardage
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmYardage))
        Me.fraCalc = New System.Windows.Forms.GroupBox()
        Me.cmdCancel = New System.Windows.Forms.Button()
        Me.cmdClear = New System.Windows.Forms.Button()
        Me.cmdOK = New System.Windows.Forms.Button()
        Me.txtWIn = New System.Windows.Forms.TextBox()
        Me.txtLIn = New System.Windows.Forms.TextBox()
        Me.txtWFt = New System.Windows.Forms.TextBox()
        Me.txtLFt = New System.Windows.Forms.TextBox()
        Me.lblIn = New System.Windows.Forms.Label()
        Me.lblFt = New System.Windows.Forms.Label()
        Me.lblWidth = New System.Windows.Forms.Label()
        Me.lblLength = New System.Windows.Forms.Label()
        Me.OptSqYd = New System.Windows.Forms.RadioButton()
        Me.txtSqYd = New System.Windows.Forms.TextBox()
        Me.optSqFt = New System.Windows.Forms.RadioButton()
        Me.txtSqFt = New System.Windows.Forms.TextBox()
        Me.updWIn = New AxMSComCtl2.AxUpDown()
        Me.updLIn = New AxMSComCtl2.AxUpDown()
        Me.updWFt = New AxMSComCtl2.AxUpDown()
        Me.updLFt = New AxMSComCtl2.AxUpDown()
        Me.fraCalc.SuspendLayout()
        CType(Me.updWIn, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.updLIn, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.updWFt, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.updLFt, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'fraCalc
        '
        Me.fraCalc.Controls.Add(Me.updWIn)
        Me.fraCalc.Controls.Add(Me.updLIn)
        Me.fraCalc.Controls.Add(Me.updWFt)
        Me.fraCalc.Controls.Add(Me.updLFt)
        Me.fraCalc.Controls.Add(Me.cmdCancel)
        Me.fraCalc.Controls.Add(Me.cmdClear)
        Me.fraCalc.Controls.Add(Me.cmdOK)
        Me.fraCalc.Controls.Add(Me.txtWIn)
        Me.fraCalc.Controls.Add(Me.txtLIn)
        Me.fraCalc.Controls.Add(Me.txtWFt)
        Me.fraCalc.Controls.Add(Me.txtLFt)
        Me.fraCalc.Controls.Add(Me.lblIn)
        Me.fraCalc.Controls.Add(Me.lblFt)
        Me.fraCalc.Controls.Add(Me.lblWidth)
        Me.fraCalc.Controls.Add(Me.lblLength)
        Me.fraCalc.Controls.Add(Me.OptSqYd)
        Me.fraCalc.Controls.Add(Me.txtSqYd)
        Me.fraCalc.Controls.Add(Me.optSqFt)
        Me.fraCalc.Controls.Add(Me.txtSqFt)
        Me.fraCalc.Location = New System.Drawing.Point(8, 3)
        Me.fraCalc.Name = "fraCalc"
        Me.fraCalc.Size = New System.Drawing.Size(295, 183)
        Me.fraCalc.TabIndex = 0
        Me.fraCalc.TabStop = False
        Me.fraCalc.Text = "Yardage/Footage Calculator:"
        '
        'cmdCancel
        '
        Me.cmdCancel.Location = New System.Drawing.Point(203, 123)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.Size = New System.Drawing.Size(59, 54)
        Me.cmdCancel.TabIndex = 14
        Me.cmdCancel.Text = "Cancel"
        Me.cmdCancel.UseVisualStyleBackColor = True
        '
        'cmdClear
        '
        Me.cmdClear.Location = New System.Drawing.Point(120, 123)
        Me.cmdClear.Name = "cmdClear"
        Me.cmdClear.Size = New System.Drawing.Size(59, 54)
        Me.cmdClear.TabIndex = 13
        Me.cmdClear.Text = "Clear"
        Me.cmdClear.UseVisualStyleBackColor = True
        '
        'cmdOK
        '
        Me.cmdOK.Location = New System.Drawing.Point(37, 123)
        Me.cmdOK.Name = "cmdOK"
        Me.cmdOK.Size = New System.Drawing.Size(59, 54)
        Me.cmdOK.TabIndex = 12
        Me.cmdOK.Text = "Apply"
        Me.cmdOK.UseVisualStyleBackColor = True
        '
        'txtWIn
        '
        Me.txtWIn.Location = New System.Drawing.Point(159, 92)
        Me.txtWIn.Name = "txtWIn"
        Me.txtWIn.Size = New System.Drawing.Size(44, 20)
        Me.txtWIn.TabIndex = 11
        '
        'txtLIn
        '
        Me.txtLIn.Location = New System.Drawing.Point(159, 65)
        Me.txtLIn.Name = "txtLIn"
        Me.txtLIn.Size = New System.Drawing.Size(44, 20)
        Me.txtLIn.TabIndex = 10
        '
        'txtWFt
        '
        Me.txtWFt.Location = New System.Drawing.Point(88, 92)
        Me.txtWFt.Name = "txtWFt"
        Me.txtWFt.Size = New System.Drawing.Size(44, 20)
        Me.txtWFt.TabIndex = 9
        '
        'txtLFt
        '
        Me.txtLFt.Location = New System.Drawing.Point(88, 65)
        Me.txtLFt.Name = "txtLFt"
        Me.txtLFt.Size = New System.Drawing.Size(44, 20)
        Me.txtLFt.TabIndex = 8
        '
        'lblIn
        '
        Me.lblIn.AutoSize = True
        Me.lblIn.Location = New System.Drawing.Point(156, 52)
        Me.lblIn.Name = "lblIn"
        Me.lblIn.Size = New System.Drawing.Size(39, 13)
        Me.lblIn.TabIndex = 7
        Me.lblIn.Text = "Inches"
        '
        'lblFt
        '
        Me.lblFt.AutoSize = True
        Me.lblFt.Location = New System.Drawing.Point(85, 52)
        Me.lblFt.Name = "lblFt"
        Me.lblFt.Size = New System.Drawing.Size(28, 13)
        Me.lblFt.TabIndex = 6
        Me.lblFt.Text = "Feet"
        '
        'lblWidth
        '
        Me.lblWidth.AutoSize = True
        Me.lblWidth.Location = New System.Drawing.Point(47, 95)
        Me.lblWidth.Name = "lblWidth"
        Me.lblWidth.Size = New System.Drawing.Size(35, 13)
        Me.lblWidth.TabIndex = 5
        Me.lblWidth.Text = "Width"
        '
        'lblLength
        '
        Me.lblLength.AutoSize = True
        Me.lblLength.Location = New System.Drawing.Point(42, 65)
        Me.lblLength.Name = "lblLength"
        Me.lblLength.Size = New System.Drawing.Size(40, 13)
        Me.lblLength.TabIndex = 4
        Me.lblLength.Text = "Length"
        '
        'OptSqYd
        '
        Me.OptSqYd.AutoSize = True
        Me.OptSqYd.Location = New System.Drawing.Point(233, 20)
        Me.OptSqYd.Name = "OptSqYd"
        Me.OptSqYd.Size = New System.Drawing.Size(59, 17)
        Me.OptSqYd.TabIndex = 3
        Me.OptSqYd.TabStop = True
        Me.OptSqYd.Text = "Sq Yds"
        Me.OptSqYd.UseVisualStyleBackColor = True
        '
        'txtSqYd
        '
        Me.txtSqYd.Location = New System.Drawing.Point(157, 20)
        Me.txtSqYd.Name = "txtSqYd"
        Me.txtSqYd.Size = New System.Drawing.Size(72, 20)
        Me.txtSqYd.TabIndex = 2
        '
        'optSqFt
        '
        Me.optSqFt.AutoSize = True
        Me.optSqFt.Location = New System.Drawing.Point(87, 20)
        Me.optSqFt.Name = "optSqFt"
        Me.optSqFt.Size = New System.Drawing.Size(62, 17)
        Me.optSqFt.TabIndex = 1
        Me.optSqFt.TabStop = True
        Me.optSqFt.Text = "Sq Feet"
        Me.optSqFt.UseVisualStyleBackColor = True
        '
        'txtSqFt
        '
        Me.txtSqFt.Location = New System.Drawing.Point(10, 20)
        Me.txtSqFt.Name = "txtSqFt"
        Me.txtSqFt.Size = New System.Drawing.Size(72, 20)
        Me.txtSqFt.TabIndex = 0
        '
        'updWIn
        '
        Me.updWIn.Location = New System.Drawing.Point(202, 92)
        Me.updWIn.Name = "updWIn"
        Me.updWIn.OcxState = CType(resources.GetObject("updWIn.OcxState"), System.Windows.Forms.AxHost.State)
        Me.updWIn.Size = New System.Drawing.Size(17, 20)
        Me.updWIn.TabIndex = 18
        '
        'updLIn
        '
        Me.updLIn.Location = New System.Drawing.Point(202, 65)
        Me.updLIn.Name = "updLIn"
        Me.updLIn.OcxState = CType(resources.GetObject("updLIn.OcxState"), System.Windows.Forms.AxHost.State)
        Me.updLIn.Size = New System.Drawing.Size(17, 20)
        Me.updLIn.TabIndex = 17
        '
        'updWFt
        '
        Me.updWFt.Location = New System.Drawing.Point(131, 92)
        Me.updWFt.Name = "updWFt"
        Me.updWFt.OcxState = CType(resources.GetObject("updWFt.OcxState"), System.Windows.Forms.AxHost.State)
        Me.updWFt.Size = New System.Drawing.Size(17, 20)
        Me.updWFt.TabIndex = 16
        '
        'updLFt
        '
        Me.updLFt.Location = New System.Drawing.Point(131, 65)
        Me.updLFt.Name = "updLFt"
        Me.updLFt.OcxState = CType(resources.GetObject("updLFt.OcxState"), System.Windows.Forms.AxHost.State)
        Me.updLFt.Size = New System.Drawing.Size(17, 20)
        Me.updLFt.TabIndex = 15
        '
        'frmYardage
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(308, 190)
        Me.Controls.Add(Me.fraCalc)
        Me.Name = "frmYardage"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Yardage/Footage Caclulator"
        Me.fraCalc.ResumeLayout(False)
        Me.fraCalc.PerformLayout()
        CType(Me.updWIn, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.updLIn, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.updWFt, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.updLFt, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents fraCalc As GroupBox
    Friend WithEvents cmdCancel As Button
    Friend WithEvents cmdClear As Button
    Friend WithEvents cmdOK As Button
    Friend WithEvents txtWIn As TextBox
    Friend WithEvents txtLIn As TextBox
    Friend WithEvents txtWFt As TextBox
    Friend WithEvents txtLFt As TextBox
    Friend WithEvents lblIn As Label
    Friend WithEvents lblFt As Label
    Friend WithEvents lblWidth As Label
    Friend WithEvents lblLength As Label
    Friend WithEvents OptSqYd As RadioButton
    Friend WithEvents txtSqYd As TextBox
    Friend WithEvents optSqFt As RadioButton
    Friend WithEvents txtSqFt As TextBox
    Friend WithEvents updWIn As AxMSComCtl2.AxUpDown
    Friend WithEvents updLIn As AxMSComCtl2.AxUpDown
    Friend WithEvents updWFt As AxMSComCtl2.AxUpDown
    Friend WithEvents updLFt As AxMSComCtl2.AxUpDown
End Class
