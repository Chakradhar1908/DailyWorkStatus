<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class OrdConvertLAW
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
        Me.lblStyle = New System.Windows.Forms.Label()
        Me.fra = New System.Windows.Forms.GroupBox()
        Me.optTW = New System.Windows.Forms.RadioButton()
        Me.optPO = New System.Windows.Forms.RadioButton()
        Me.optSO = New System.Windows.Forms.RadioButton()
        Me.optST = New System.Windows.Forms.RadioButton()
        Me.optNn = New System.Windows.Forms.RadioButton()
        Me.ShapeContainer1 = New Microsoft.VisualBasic.PowerPacks.ShapeContainer()
        Me.LineShape1 = New Microsoft.VisualBasic.PowerPacks.LineShape()
        Me.lblAvailable = New System.Windows.Forms.Label()
        Me.txtAvailable = New System.Windows.Forms.TextBox()
        Me.lblOnOrder = New System.Windows.Forms.Label()
        Me.txtOnOrder = New System.Windows.Forms.TextBox()
        Me.LineShape2 = New Microsoft.VisualBasic.PowerPacks.LineShape()
        Me.cmdViewPO = New System.Windows.Forms.Button()
        Me.lblLoc = New System.Windows.Forms.Label()
        Me.txtLoc = New System.Windows.Forms.TextBox()
        Me.LineShape3 = New Microsoft.VisualBasic.PowerPacks.LineShape()
        Me.cmdOK = New System.Windows.Forms.Button()
        Me.cmdCancel = New System.Windows.Forms.Button()
        Me.fra.SuspendLayout()
        Me.SuspendLayout()
        '
        'lblStyle
        '
        Me.lblStyle.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblStyle.Font = New System.Drawing.Font("Arial", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblStyle.Location = New System.Drawing.Point(12, 9)
        Me.lblStyle.Name = "lblStyle"
        Me.lblStyle.Size = New System.Drawing.Size(223, 23)
        Me.lblStyle.TabIndex = 0
        Me.lblStyle.Text = "##"
        Me.lblStyle.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'fra
        '
        Me.fra.Controls.Add(Me.cmdCancel)
        Me.fra.Controls.Add(Me.cmdOK)
        Me.fra.Controls.Add(Me.txtLoc)
        Me.fra.Controls.Add(Me.lblLoc)
        Me.fra.Controls.Add(Me.cmdViewPO)
        Me.fra.Controls.Add(Me.txtOnOrder)
        Me.fra.Controls.Add(Me.lblOnOrder)
        Me.fra.Controls.Add(Me.txtAvailable)
        Me.fra.Controls.Add(Me.lblAvailable)
        Me.fra.Controls.Add(Me.optTW)
        Me.fra.Controls.Add(Me.optPO)
        Me.fra.Controls.Add(Me.optSO)
        Me.fra.Controls.Add(Me.optST)
        Me.fra.Controls.Add(Me.optNn)
        Me.fra.Controls.Add(Me.ShapeContainer1)
        Me.fra.Location = New System.Drawing.Point(11, 34)
        Me.fra.Name = "fra"
        Me.fra.Size = New System.Drawing.Size(224, 233)
        Me.fra.TabIndex = 1
        Me.fra.TabStop = False
        '
        'optTW
        '
        Me.optTW.AutoSize = True
        Me.optTW.Location = New System.Drawing.Point(15, 111)
        Me.optTW.Name = "optTW"
        Me.optTW.Size = New System.Drawing.Size(127, 17)
        Me.optTW.TabIndex = 4
        Me.optTW.TabStop = True
        Me.optTW.Text = "Convert To &DELTTW"
        Me.optTW.UseVisualStyleBackColor = True
        '
        'optPO
        '
        Me.optPO.AutoSize = True
        Me.optPO.Location = New System.Drawing.Point(15, 88)
        Me.optPO.Name = "optPO"
        Me.optPO.Size = New System.Drawing.Size(101, 17)
        Me.optPO.TabIndex = 3
        Me.optPO.TabStop = True
        Me.optPO.Text = "Convert To &P/O"
        Me.optPO.UseVisualStyleBackColor = True
        '
        'optSO
        '
        Me.optSO.AutoSize = True
        Me.optSO.Location = New System.Drawing.Point(15, 65)
        Me.optSO.Name = "optSO"
        Me.optSO.Size = New System.Drawing.Size(101, 17)
        Me.optSO.TabIndex = 2
        Me.optSO.TabStop = True
        Me.optSO.Text = "Convert To &S/O"
        Me.optSO.UseVisualStyleBackColor = True
        '
        'optST
        '
        Me.optST.AutoSize = True
        Me.optST.Location = New System.Drawing.Point(15, 42)
        Me.optST.Name = "optST"
        Me.optST.Size = New System.Drawing.Size(109, 17)
        Me.optST.TabIndex = 1
        Me.optST.TabStop = True
        Me.optST.Text = "Convert To S&tock"
        Me.optST.UseVisualStyleBackColor = True
        '
        'optNn
        '
        Me.optNn.AutoSize = True
        Me.optNn.Location = New System.Drawing.Point(15, 19)
        Me.optNn.Name = "optNn"
        Me.optNn.Size = New System.Drawing.Size(99, 17)
        Me.optNn.TabIndex = 0
        Me.optNn.TabStop = True
        Me.optNn.Text = "Do Not Chan&ge"
        Me.optNn.UseVisualStyleBackColor = True
        '
        'ShapeContainer1
        '
        Me.ShapeContainer1.Location = New System.Drawing.Point(3, 16)
        Me.ShapeContainer1.Margin = New System.Windows.Forms.Padding(0)
        Me.ShapeContainer1.Name = "ShapeContainer1"
        Me.ShapeContainer1.Shapes.AddRange(New Microsoft.VisualBasic.PowerPacks.Shape() {Me.LineShape3, Me.LineShape2, Me.LineShape1})
        Me.ShapeContainer1.Size = New System.Drawing.Size(218, 214)
        Me.ShapeContainer1.TabIndex = 5
        Me.ShapeContainer1.TabStop = False
        '
        'LineShape1
        '
        Me.LineShape1.BorderColor = System.Drawing.SystemColors.ScrollBar
        Me.LineShape1.Name = "LineShape1"
        Me.LineShape1.X1 = 140
        Me.LineShape1.X2 = 141
        Me.LineShape1.Y1 = -9
        Me.LineShape1.Y2 = 124
        '
        'lblAvailable
        '
        Me.lblAvailable.AutoSize = True
        Me.lblAvailable.Location = New System.Drawing.Point(157, 46)
        Me.lblAvailable.Name = "lblAvailable"
        Me.lblAvailable.Size = New System.Drawing.Size(53, 13)
        Me.lblAvailable.TabIndex = 6
        Me.lblAvailable.Text = "Available:"
        '
        'txtAvailable
        '
        Me.txtAvailable.BackColor = System.Drawing.SystemColors.Control
        Me.txtAvailable.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.txtAvailable.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtAvailable.Location = New System.Drawing.Point(154, 69)
        Me.txtAvailable.Name = "txtAvailable"
        Me.txtAvailable.Size = New System.Drawing.Size(56, 15)
        Me.txtAvailable.TabIndex = 7
        Me.txtAvailable.Text = "###"
        Me.txtAvailable.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'lblOnOrder
        '
        Me.lblOnOrder.AutoSize = True
        Me.lblOnOrder.Location = New System.Drawing.Point(157, 92)
        Me.lblOnOrder.Name = "lblOnOrder"
        Me.lblOnOrder.Size = New System.Drawing.Size(53, 13)
        Me.lblOnOrder.TabIndex = 8
        Me.lblOnOrder.Text = "On Order:"
        '
        'txtOnOrder
        '
        Me.txtOnOrder.BackColor = System.Drawing.SystemColors.Control
        Me.txtOnOrder.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.txtOnOrder.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtOnOrder.Location = New System.Drawing.Point(149, 111)
        Me.txtOnOrder.Name = "txtOnOrder"
        Me.txtOnOrder.Size = New System.Drawing.Size(66, 15)
        Me.txtOnOrder.TabIndex = 9
        Me.txtOnOrder.Text = "###"
        Me.txtOnOrder.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'LineShape2
        '
        Me.LineShape2.BorderColor = System.Drawing.SystemColors.ScrollBar
        Me.LineShape2.Name = "LineShape2"
        Me.LineShape2.X1 = 2
        Me.LineShape2.X2 = 247
        Me.LineShape2.Y1 = 125
        Me.LineShape2.Y2 = 124
        '
        'cmdViewPO
        '
        Me.cmdViewPO.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.cmdViewPO.Location = New System.Drawing.Point(10, 146)
        Me.cmdViewPO.Name = "cmdViewPO"
        Me.cmdViewPO.Size = New System.Drawing.Size(61, 23)
        Me.cmdViewPO.TabIndex = 10
        Me.cmdViewPO.Text = "View P&O"
        Me.cmdViewPO.UseVisualStyleBackColor = False
        '
        'lblLoc
        '
        Me.lblLoc.AutoSize = True
        Me.lblLoc.Location = New System.Drawing.Point(78, 146)
        Me.lblLoc.Name = "lblLoc"
        Me.lblLoc.Size = New System.Drawing.Size(51, 13)
        Me.lblLoc.TabIndex = 11
        Me.lblLoc.Text = "Locatio&n:"
        '
        'txtLoc
        '
        Me.txtLoc.Location = New System.Drawing.Point(135, 146)
        Me.txtLoc.Name = "txtLoc"
        Me.txtLoc.Size = New System.Drawing.Size(33, 20)
        Me.txtLoc.TabIndex = 12
        '
        'LineShape3
        '
        Me.LineShape3.BorderColor = System.Drawing.SystemColors.ScrollBar
        Me.LineShape3.Name = "LineShape3"
        Me.LineShape3.X1 = 0
        Me.LineShape3.X2 = 254
        Me.LineShape3.Y1 = 157
        Me.LineShape3.Y2 = 154
        '
        'cmdOK
        '
        Me.cmdOK.Location = New System.Drawing.Point(34, 183)
        Me.cmdOK.Name = "cmdOK"
        Me.cmdOK.Size = New System.Drawing.Size(75, 42)
        Me.cmdOK.TabIndex = 13
        Me.cmdOK.Text = "&OK"
        Me.cmdOK.UseVisualStyleBackColor = True
        '
        'cmdCancel
        '
        Me.cmdCancel.Location = New System.Drawing.Point(115, 183)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.Size = New System.Drawing.Size(75, 42)
        Me.cmdCancel.TabIndex = 14
        Me.cmdCancel.Text = "&Cancel"
        Me.cmdCancel.UseVisualStyleBackColor = True
        '
        'OrdConvertLAW
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(242, 275)
        Me.Controls.Add(Me.fra)
        Me.Controls.Add(Me.lblStyle)
        Me.Name = "OrdConvertLAW"
        Me.Text = "Convert LAW To..."
        Me.fra.ResumeLayout(False)
        Me.fra.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents lblStyle As Label
    Friend WithEvents fra As GroupBox
    Friend WithEvents optTW As RadioButton
    Friend WithEvents optPO As RadioButton
    Friend WithEvents optSO As RadioButton
    Friend WithEvents optST As RadioButton
    Friend WithEvents optNn As RadioButton
    Friend WithEvents cmdCancel As Button
    Friend WithEvents cmdOK As Button
    Friend WithEvents txtLoc As TextBox
    Friend WithEvents lblLoc As Label
    Friend WithEvents cmdViewPO As Button
    Friend WithEvents txtOnOrder As TextBox
    Friend WithEvents lblOnOrder As Label
    Friend WithEvents txtAvailable As TextBox
    Friend WithEvents lblAvailable As Label
    Friend WithEvents ShapeContainer1 As PowerPacks.ShapeContainer
    Friend WithEvents LineShape3 As PowerPacks.LineShape
    Friend WithEvents LineShape2 As PowerPacks.LineShape
    Friend WithEvents LineShape1 As PowerPacks.LineShape
End Class
