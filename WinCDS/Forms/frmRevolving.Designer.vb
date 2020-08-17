<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmRevolving
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
        Me.lstLog = New System.Windows.Forms.ListBox()
        Me.fraDev = New System.Windows.Forms.GroupBox()
        Me.chkDevVerbose = New System.Windows.Forms.CheckBox()
        Me.chkDevDebugLite = New System.Windows.Forms.CheckBox()
        Me.chkDevDebugFull = New System.Windows.Forms.CheckBox()
        Me.SuspendLayout()
        '
        'lstLog
        '
        Me.lstLog.FormattingEnabled = True
        Me.lstLog.Location = New System.Drawing.Point(0, 0)
        Me.lstLog.Name = "lstLog"
        Me.lstLog.Size = New System.Drawing.Size(120, 95)
        Me.lstLog.TabIndex = 0
        '
        'fraDev
        '
        Me.fraDev.Location = New System.Drawing.Point(12, 131)
        Me.fraDev.Name = "fraDev"
        Me.fraDev.Size = New System.Drawing.Size(200, 100)
        Me.fraDev.TabIndex = 1
        Me.fraDev.TabStop = False
        Me.fraDev.Text = "fraDev"
        '
        'chkDevVerbose
        '
        Me.chkDevVerbose.AutoSize = True
        Me.chkDevVerbose.Location = New System.Drawing.Point(12, 251)
        Me.chkDevVerbose.Name = "chkDevVerbose"
        Me.chkDevVerbose.Size = New System.Drawing.Size(103, 17)
        Me.chkDevVerbose.TabIndex = 0
        Me.chkDevVerbose.Text = "chkDevVerbose"
        Me.chkDevVerbose.UseVisualStyleBackColor = True
        '
        'chkDevDebugLite
        '
        Me.chkDevDebugLite.AutoSize = True
        Me.chkDevDebugLite.Location = New System.Drawing.Point(12, 274)
        Me.chkDevDebugLite.Name = "chkDevDebugLite"
        Me.chkDevDebugLite.Size = New System.Drawing.Size(113, 17)
        Me.chkDevDebugLite.TabIndex = 2
        Me.chkDevDebugLite.Text = "chkDevDebugLite"
        Me.chkDevDebugLite.UseVisualStyleBackColor = True
        '
        'chkDevDebugFull
        '
        Me.chkDevDebugFull.AutoSize = True
        Me.chkDevDebugFull.Location = New System.Drawing.Point(12, 297)
        Me.chkDevDebugFull.Name = "chkDevDebugFull"
        Me.chkDevDebugFull.Size = New System.Drawing.Size(112, 17)
        Me.chkDevDebugFull.TabIndex = 3
        Me.chkDevDebugFull.Text = "chkDevDebugFull"
        Me.chkDevDebugFull.UseVisualStyleBackColor = True
        '
        'frmRevolving
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(800, 450)
        Me.Controls.Add(Me.chkDevDebugFull)
        Me.Controls.Add(Me.chkDevDebugLite)
        Me.Controls.Add(Me.chkDevVerbose)
        Me.Controls.Add(Me.fraDev)
        Me.Controls.Add(Me.lstLog)
        Me.Name = "frmRevolving"
        Me.Text = "frmRevolving"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents lstLog As ListBox
    Friend WithEvents fraDev As GroupBox
    Friend WithEvents chkDevVerbose As CheckBox
    Friend WithEvents chkDevDebugLite As CheckBox
    Friend WithEvents chkDevDebugFull As CheckBox
End Class
