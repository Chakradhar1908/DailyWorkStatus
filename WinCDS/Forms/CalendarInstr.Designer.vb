<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class CalendarInstr
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
        Me.fraDate = New System.Windows.Forms.GroupBox()
        Me.tvwList = New System.Windows.Forms.TreeView()
        Me.cmdExit = New System.Windows.Forms.Button()
        Me.fraDate.SuspendLayout()
        Me.SuspendLayout()
        '
        'fraDate
        '
        Me.fraDate.Controls.Add(Me.tvwList)
        Me.fraDate.Location = New System.Drawing.Point(12, 12)
        Me.fraDate.Name = "fraDate"
        Me.fraDate.Size = New System.Drawing.Size(396, 209)
        Me.fraDate.TabIndex = 0
        Me.fraDate.TabStop = False
        Me.fraDate.Text = "#"
        '
        'tvwList
        '
        Me.tvwList.Location = New System.Drawing.Point(10, 16)
        Me.tvwList.Name = "tvwList"
        Me.tvwList.Size = New System.Drawing.Size(369, 184)
        Me.tvwList.TabIndex = 0
        '
        'cmdExit
        '
        Me.cmdExit.Location = New System.Drawing.Point(123, 227)
        Me.cmdExit.Name = "cmdExit"
        Me.cmdExit.Size = New System.Drawing.Size(75, 39)
        Me.cmdExit.TabIndex = 1
        Me.cmdExit.Text = "&Close"
        Me.cmdExit.UseVisualStyleBackColor = True
        '
        'CalendarInstr
        '
        Me.AcceptButton = Me.cmdExit
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(800, 450)
        Me.Controls.Add(Me.cmdExit)
        Me.Controls.Add(Me.fraDate)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.SizableToolWindow
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "CalendarInstr"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Special Instructions"
        Me.fraDate.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents fraDate As GroupBox
    Friend WithEvents tvwList As TreeView
    Friend WithEvents cmdExit As Button
End Class
