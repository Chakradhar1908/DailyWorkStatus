<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class frmOptimizeRoute
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
        Me.cmdOK = New System.Windows.Forms.Button()
        Me.cmdContinue = New System.Windows.Forms.Button()
        Me.lblTCost = New System.Windows.Forms.Label()
        Me.txtTCost = New System.Windows.Forms.Label()
        Me.lvw = New System.Windows.Forms.ListView()
        Me.ColumnHeader1 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader2 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader3 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader4 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader5 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader6 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader7 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader8 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader9 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader10 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader11 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader12 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader13 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader14 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.cmdUD0 = New System.Windows.Forms.Button()
        Me.cmdUD1 = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'cmdOK
        '
        Me.cmdOK.Location = New System.Drawing.Point(8, 192)
        Me.cmdOK.Name = "cmdOK"
        Me.cmdOK.Size = New System.Drawing.Size(86, 53)
        Me.cmdOK.TabIndex = 1
        Me.cmdOK.Text = "C&lose"
        Me.cmdOK.UseVisualStyleBackColor = True
        '
        'cmdContinue
        '
        Me.cmdContinue.Location = New System.Drawing.Point(100, 192)
        Me.cmdContinue.Name = "cmdContinue"
        Me.cmdContinue.Size = New System.Drawing.Size(86, 53)
        Me.cmdContinue.TabIndex = 2
        Me.cmdContinue.Text = "C&ontinue"
        Me.cmdContinue.UseVisualStyleBackColor = True
        '
        'lblTCost
        '
        Me.lblTCost.AutoSize = True
        Me.lblTCost.Location = New System.Drawing.Point(396, 192)
        Me.lblTCost.Name = "lblTCost"
        Me.lblTCost.Size = New System.Drawing.Size(58, 13)
        Me.lblTCost.TabIndex = 3
        Me.lblTCost.Text = "Total Cost:"
        '
        'txtTCost
        '
        Me.txtTCost.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.txtTCost.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtTCost.Location = New System.Drawing.Point(460, 191)
        Me.txtTCost.Name = "txtTCost"
        Me.txtTCost.Size = New System.Drawing.Size(100, 22)
        Me.txtTCost.TabIndex = 4
        Me.txtTCost.Text = "0.00"
        Me.txtTCost.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'lvw
        '
        Me.lvw.BackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.lvw.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lvw.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.ColumnHeader1, Me.ColumnHeader2, Me.ColumnHeader3, Me.ColumnHeader4, Me.ColumnHeader5, Me.ColumnHeader6, Me.ColumnHeader7, Me.ColumnHeader8, Me.ColumnHeader9, Me.ColumnHeader10, Me.ColumnHeader11, Me.ColumnHeader12, Me.ColumnHeader13, Me.ColumnHeader14})
        Me.lvw.FullRowSelect = True
        Me.lvw.GridLines = True
        Me.lvw.HideSelection = False
        Me.lvw.LabelEdit = True
        Me.lvw.Location = New System.Drawing.Point(10, 8)
        Me.lvw.Name = "lvw"
        Me.lvw.Size = New System.Drawing.Size(550, 177)
        Me.lvw.Sorting = System.Windows.Forms.SortOrder.Ascending
        Me.lvw.TabIndex = 10
        Me.lvw.UseCompatibleStateImageBehavior = False
        Me.lvw.View = System.Windows.Forms.View.Details
        '
        'ColumnHeader1
        '
        Me.ColumnHeader1.Text = "ID"
        Me.ColumnHeader1.Width = 0
        '
        'ColumnHeader2
        '
        Me.ColumnHeader2.Text = "StopName"
        Me.ColumnHeader2.Width = 120
        '
        'ColumnHeader3
        '
        Me.ColumnHeader3.Text = "X"
        Me.ColumnHeader3.Width = 0
        '
        'ColumnHeader4
        '
        Me.ColumnHeader4.Text = "Y"
        Me.ColumnHeader4.Width = 0
        '
        'ColumnHeader5
        '
        Me.ColumnHeader5.Text = "WFr"
        Me.ColumnHeader5.Width = 40
        '
        'ColumnHeader6
        '
        Me.ColumnHeader6.Text = "WTo"
        Me.ColumnHeader6.Width = 40
        '
        'ColumnHeader7
        '
        Me.ColumnHeader7.Text = "Dist"
        Me.ColumnHeader7.Width = 40
        '
        'ColumnHeader8
        '
        Me.ColumnHeader8.Text = "Delay"
        Me.ColumnHeader8.Width = 40
        '
        'ColumnHeader9
        '
        Me.ColumnHeader9.Text = "Arrival"
        Me.ColumnHeader9.Width = 50
        '
        'ColumnHeader10
        '
        Me.ColumnHeader10.Text = "Stop Duration"
        Me.ColumnHeader10.Width = 75
        '
        'ColumnHeader11
        '
        Me.ColumnHeader11.Text = "Departure"
        '
        'ColumnHeader12
        '
        Me.ColumnHeader12.Text = "City"
        Me.ColumnHeader12.Width = 130
        '
        'ColumnHeader13
        '
        Me.ColumnHeader13.Text = "State"
        Me.ColumnHeader13.Width = 130
        '
        'ColumnHeader14
        '
        Me.ColumnHeader14.Text = "Zip"
        '
        'cmdUD0
        '
        Me.cmdUD0.Location = New System.Drawing.Point(566, 78)
        Me.cmdUD0.Name = "cmdUD0"
        Me.cmdUD0.Size = New System.Drawing.Size(20, 23)
        Me.cmdUD0.TabIndex = 11
        Me.cmdUD0.UseVisualStyleBackColor = True
        '
        'cmdUD1
        '
        Me.cmdUD1.Location = New System.Drawing.Point(566, 110)
        Me.cmdUD1.Name = "cmdUD1"
        Me.cmdUD1.Size = New System.Drawing.Size(21, 23)
        Me.cmdUD1.TabIndex = 12
        Me.cmdUD1.UseVisualStyleBackColor = True
        '
        'frmOptimizeRoute
        '
        Me.AcceptButton = Me.cmdOK
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(596, 249)
        Me.Controls.Add(Me.cmdUD1)
        Me.Controls.Add(Me.cmdUD0)
        Me.Controls.Add(Me.lvw)
        Me.Controls.Add(Me.txtTCost)
        Me.Controls.Add(Me.lblTCost)
        Me.Controls.Add(Me.cmdContinue)
        Me.Controls.Add(Me.cmdOK)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow
        Me.Name = "frmOptimizeRoute"
        Me.Text = "Route"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents cmdOK As Button
    Friend WithEvents cmdContinue As Button
    Friend WithEvents lblTCost As Label
    Friend WithEvents txtTCost As Label
    Friend WithEvents lvw As ListView
    Friend WithEvents ColumnHeader1 As ColumnHeader
    Friend WithEvents ColumnHeader2 As ColumnHeader
    Friend WithEvents ColumnHeader3 As ColumnHeader
    Friend WithEvents ColumnHeader4 As ColumnHeader
    Friend WithEvents ColumnHeader5 As ColumnHeader
    Friend WithEvents ColumnHeader6 As ColumnHeader
    Friend WithEvents ColumnHeader7 As ColumnHeader
    Friend WithEvents ColumnHeader8 As ColumnHeader
    Friend WithEvents ColumnHeader9 As ColumnHeader
    Friend WithEvents ColumnHeader10 As ColumnHeader
    Friend WithEvents ColumnHeader11 As ColumnHeader
    Friend WithEvents ColumnHeader12 As ColumnHeader
    Friend WithEvents ColumnHeader13 As ColumnHeader
    Friend WithEvents ColumnHeader14 As ColumnHeader
    Friend WithEvents cmdUD0 As Button
    Friend WithEvents cmdUD1 As Button
End Class
