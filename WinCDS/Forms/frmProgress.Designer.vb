<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmProgress
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
        Me.lblCaption = New System.Windows.Forms.Label()
        Me.prg = New System.Windows.Forms.Label()
        Me.ShapeContainer1 = New Microsoft.VisualBasic.PowerPacks.ShapeContainer()
        Me.linBorder = New Microsoft.VisualBasic.PowerPacks.LineShape()
        Me.ProgressBar1 = New System.Windows.Forms.ProgressBar()
        Me.btnProgress = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'lblCaption
        '
        Me.lblCaption.AutoSize = True
        Me.lblCaption.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblCaption.Location = New System.Drawing.Point(12, 9)
        Me.lblCaption.Name = "lblCaption"
        Me.lblCaption.Size = New System.Drawing.Size(84, 15)
        Me.lblCaption.TabIndex = 0
        Me.lblCaption.Text = "Please Wait..."
        '
        'prg
        '
        Me.prg.AutoSize = True
        Me.prg.Location = New System.Drawing.Point(24, 37)
        Me.prg.Name = "prg"
        Me.prg.Size = New System.Drawing.Size(0, 13)
        Me.prg.TabIndex = 1
        '
        'ShapeContainer1
        '
        Me.ShapeContainer1.Location = New System.Drawing.Point(0, 0)
        Me.ShapeContainer1.Margin = New System.Windows.Forms.Padding(0)
        Me.ShapeContainer1.Name = "ShapeContainer1"
        Me.ShapeContainer1.Shapes.AddRange(New Microsoft.VisualBasic.PowerPacks.Shape() {Me.linBorder})
        Me.ShapeContainer1.Size = New System.Drawing.Size(273, 88)
        Me.ShapeContainer1.TabIndex = 2
        Me.ShapeContainer1.TabStop = False
        '
        'linBorder
        '
        Me.linBorder.Name = "linBorder"
        Me.linBorder.Visible = False
        Me.linBorder.X1 = 11
        Me.linBorder.X2 = 11
        Me.linBorder.Y1 = 3
        Me.linBorder.Y2 = 40
        '
        'ProgressBar1
        '
        Me.ProgressBar1.Location = New System.Drawing.Point(12, 37)
        Me.ProgressBar1.Name = "ProgressBar1"
        Me.ProgressBar1.Size = New System.Drawing.Size(258, 13)
        Me.ProgressBar1.Style = System.Windows.Forms.ProgressBarStyle.Marquee
        Me.ProgressBar1.TabIndex = 3
        '
        'btnProgress
        '
        Me.btnProgress.Location = New System.Drawing.Point(-100, 62)
        Me.btnProgress.Name = "btnProgress"
        Me.btnProgress.Size = New System.Drawing.Size(75, 23)
        Me.btnProgress.TabIndex = 4
        Me.btnProgress.Text = "Progress"
        Me.btnProgress.UseVisualStyleBackColor = True
        '
        'frmProgress
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(273, 88)
        Me.Controls.Add(Me.btnProgress)
        Me.Controls.Add(Me.ProgressBar1)
        Me.Controls.Add(Me.prg)
        Me.Controls.Add(Me.lblCaption)
        Me.Controls.Add(Me.ShapeContainer1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Name = "frmProgress"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "frmProgress"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents lblCaption As Label
    Friend WithEvents prg As Label
    Friend WithEvents ShapeContainer1 As PowerPacks.ShapeContainer
    Friend WithEvents linBorder As PowerPacks.LineShape
    Friend WithEvents ProgressBar1 As ProgressBar
    Friend WithEvents btnProgress As Button
End Class
