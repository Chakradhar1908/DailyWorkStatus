<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmOptimizeRoute
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
        Me.lvw = New System.Windows.Forms.ListView()
        Me.cmdOK = New System.Windows.Forms.Button()
        Me.cmdContinue = New System.Windows.Forms.Button()
        Me.lblTCost = New System.Windows.Forms.Label()
        Me.txtTCost = New System.Windows.Forms.Label()
        Me.cmdUD1 = New System.Windows.Forms.Label()
        Me.cmdUD0 = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'lvw
        '
        Me.lvw.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lvw.HideSelection = False
        Me.lvw.Location = New System.Drawing.Point(8, 6)
        Me.lvw.Name = "lvw"
        Me.lvw.Size = New System.Drawing.Size(552, 127)
        Me.lvw.Sorting = System.Windows.Forms.SortOrder.Ascending
        Me.lvw.TabIndex = 0
        Me.lvw.UseCompatibleStateImageBehavior = False
        Me.lvw.View = System.Windows.Forms.View.Tile
        '
        'cmdOK
        '
        Me.cmdOK.Location = New System.Drawing.Point(8, 140)
        Me.cmdOK.Name = "cmdOK"
        Me.cmdOK.Size = New System.Drawing.Size(75, 49)
        Me.cmdOK.TabIndex = 1
        Me.cmdOK.Text = "C&lose"
        Me.cmdOK.UseVisualStyleBackColor = True
        '
        'cmdContinue
        '
        Me.cmdContinue.Location = New System.Drawing.Point(87, 140)
        Me.cmdContinue.Name = "cmdContinue"
        Me.cmdContinue.Size = New System.Drawing.Size(75, 49)
        Me.cmdContinue.TabIndex = 2
        Me.cmdContinue.Text = "C&ontinue"
        Me.cmdContinue.UseVisualStyleBackColor = True
        '
        'lblTCost
        '
        Me.lblTCost.AutoSize = True
        Me.lblTCost.Location = New System.Drawing.Point(376, 143)
        Me.lblTCost.Name = "lblTCost"
        Me.lblTCost.Size = New System.Drawing.Size(58, 13)
        Me.lblTCost.TabIndex = 3
        Me.lblTCost.Text = "Total Cost:"
        '
        'txtTCost
        '
        Me.txtTCost.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.txtTCost.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtTCost.Location = New System.Drawing.Point(460, 140)
        Me.txtTCost.Name = "txtTCost"
        Me.txtTCost.Size = New System.Drawing.Size(100, 23)
        Me.txtTCost.TabIndex = 4
        Me.txtTCost.Text = "0.00"
        Me.txtTCost.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'cmdUD1
        '
        Me.cmdUD1.Image = Global.WinCDS.My.Resources.Resources.small_down
        Me.cmdUD1.Location = New System.Drawing.Point(566, 74)
        Me.cmdUD1.Name = "cmdUD1"
        Me.cmdUD1.Size = New System.Drawing.Size(24, 23)
        Me.cmdUD1.TabIndex = 8
        '
        'cmdUD0
        '
        Me.cmdUD0.Image = Global.WinCDS.My.Resources.Resources.small_up
        Me.cmdUD0.Location = New System.Drawing.Point(566, 45)
        Me.cmdUD0.Name = "cmdUD0"
        Me.cmdUD0.Size = New System.Drawing.Size(23, 23)
        Me.cmdUD0.TabIndex = 9
        '
        'frmOptimizeRoute
        '
        Me.AcceptButton = Me.cmdOK
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(665, 195)
        Me.Controls.Add(Me.cmdUD0)
        Me.Controls.Add(Me.cmdUD1)
        Me.Controls.Add(Me.txtTCost)
        Me.Controls.Add(Me.lblTCost)
        Me.Controls.Add(Me.cmdContinue)
        Me.Controls.Add(Me.cmdOK)
        Me.Controls.Add(Me.lvw)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow
        Me.Name = "frmOptimizeRoute"
        Me.Text = "Route"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents lvw As ListView
    Friend WithEvents cmdOK As Button
    Friend WithEvents cmdContinue As Button
    Friend WithEvents lblTCost As Label
    Friend WithEvents txtTCost As Label
    Friend WithEvents cmdUD1 As Label
    Friend WithEvents cmdUD0 As Label
End Class
