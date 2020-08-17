<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Practice
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Practice))
        Me.pbarMargin = New System.Windows.Forms.ProgressBar()
        Me.cmdConvertOld = New System.Windows.Forms.Button()
        Me.cmdWinCDSOnly = New System.Windows.Forms.Button()
        Me.lblLoc = New System.Windows.Forms.Label()
        Me.txtLoc = New System.Windows.Forms.TextBox()
        Me.fraStartupCrash = New System.Windows.Forms.GroupBox()
        Me.updLoc = New AxMSComCtl2.AxUpDown()
        CType(Me.updLoc, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'pbarMargin
        '
        Me.pbarMargin.Location = New System.Drawing.Point(469, 172)
        Me.pbarMargin.Name = "pbarMargin"
        Me.pbarMargin.Size = New System.Drawing.Size(100, 23)
        Me.pbarMargin.TabIndex = 0
        '
        'cmdConvertOld
        '
        Me.cmdConvertOld.Location = New System.Drawing.Point(418, 244)
        Me.cmdConvertOld.Name = "cmdConvertOld"
        Me.cmdConvertOld.Size = New System.Drawing.Size(75, 23)
        Me.cmdConvertOld.TabIndex = 1
        Me.cmdConvertOld.Text = "Button1"
        Me.cmdConvertOld.UseVisualStyleBackColor = True
        '
        'cmdWinCDSOnly
        '
        Me.cmdWinCDSOnly.Location = New System.Drawing.Point(396, 296)
        Me.cmdWinCDSOnly.Name = "cmdWinCDSOnly"
        Me.cmdWinCDSOnly.Size = New System.Drawing.Size(75, 23)
        Me.cmdWinCDSOnly.TabIndex = 2
        Me.cmdWinCDSOnly.Text = "Button2"
        Me.cmdWinCDSOnly.UseVisualStyleBackColor = True
        '
        'lblLoc
        '
        Me.lblLoc.AutoSize = True
        Me.lblLoc.Location = New System.Drawing.Point(401, 356)
        Me.lblLoc.Name = "lblLoc"
        Me.lblLoc.Size = New System.Drawing.Size(39, 13)
        Me.lblLoc.TabIndex = 3
        Me.lblLoc.Text = "Label1"
        '
        'txtLoc
        '
        Me.txtLoc.Location = New System.Drawing.Point(340, 390)
        Me.txtLoc.Name = "txtLoc"
        Me.txtLoc.Size = New System.Drawing.Size(100, 20)
        Me.txtLoc.TabIndex = 4
        '
        'fraStartupCrash
        '
        Me.fraStartupCrash.Location = New System.Drawing.Point(535, 269)
        Me.fraStartupCrash.Name = "fraStartupCrash"
        Me.fraStartupCrash.Size = New System.Drawing.Size(200, 100)
        Me.fraStartupCrash.TabIndex = 5
        Me.fraStartupCrash.TabStop = False
        Me.fraStartupCrash.Text = "GroupBox1"
        '
        'updLoc
        '
        Me.updLoc.Location = New System.Drawing.Point(371, 71)
        Me.updLoc.Name = "updLoc"
        Me.updLoc.OcxState = CType(resources.GetObject("updLoc.OcxState"), System.Windows.Forms.AxHost.State)
        Me.updLoc.Size = New System.Drawing.Size(17, 50)
        Me.updLoc.TabIndex = 6
        '
        'Practice
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(800, 450)
        Me.Controls.Add(Me.updLoc)
        Me.Controls.Add(Me.fraStartupCrash)
        Me.Controls.Add(Me.txtLoc)
        Me.Controls.Add(Me.lblLoc)
        Me.Controls.Add(Me.cmdWinCDSOnly)
        Me.Controls.Add(Me.cmdConvertOld)
        Me.Controls.Add(Me.pbarMargin)
        Me.Name = "Practice"
        Me.Text = "Practice"
        CType(Me.updLoc, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents pbarMargin As ProgressBar
    Friend WithEvents cmdConvertOld As Button
    Friend WithEvents cmdWinCDSOnly As Button
    Friend WithEvents lblLoc As Label
    Friend WithEvents txtLoc As TextBox
    Friend WithEvents fraStartupCrash As GroupBox
    Friend WithEvents updLoc As AxMSComCtl2.AxUpDown
    'Friend WithEvents updLoc As AxComCtl2.AxUpDown
End Class
