<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class RTBCompose
    Inherits System.Windows.Forms.UserControl

    'UserControl overrides dispose to clean up the component list.
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
        Me.lblFontColor = New System.Windows.Forms.Label()
        Me.cmdSend = New System.Windows.Forms.Button()
        Me.cmdPicture = New System.Windows.Forms.Button()
        Me.cmdRight = New System.Windows.Forms.Button()
        Me.cmdCenter = New System.Windows.Forms.Button()
        Me.cmdLeft = New System.Windows.Forms.Button()
        Me.cmdFontColor = New System.Windows.Forms.Button()
        Me.chkStrikeThru = New System.Windows.Forms.CheckBox()
        Me.chkUnderline = New System.Windows.Forms.CheckBox()
        Me.chkItalic = New System.Windows.Forms.CheckBox()
        Me.chkBold = New System.Windows.Forms.CheckBox()
        Me.cmdDelete = New System.Windows.Forms.Button()
        Me.cmdPaste = New System.Windows.Forms.Button()
        Me.cmdCopy = New System.Windows.Forms.Button()
        Me.cmdCut = New System.Windows.Forms.Button()
        Me.RTB = New System.Windows.Forms.RichTextBox()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.SuspendLayout()
        '
        'lblFontColor
        '
        Me.lblFontColor.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblFontColor.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblFontColor.Location = New System.Drawing.Point(228, 3)
        Me.lblFontColor.Name = "lblFontColor"
        Me.lblFontColor.Size = New System.Drawing.Size(127, 23)
        Me.lblFontColor.TabIndex = 15
        '
        'cmdSend
        '
        Me.cmdSend.Image = Global.WinCDS.My.Resources.Resources.MAIL
        Me.cmdSend.Location = New System.Drawing.Point(495, 3)
        Me.cmdSend.Name = "cmdSend"
        Me.cmdSend.Size = New System.Drawing.Size(28, 23)
        Me.cmdSend.TabIndex = 14
        Me.ToolTip1.SetToolTip(Me.cmdSend, "Send")
        Me.cmdSend.UseVisualStyleBackColor = True
        '
        'cmdPicture
        '
        Me.cmdPicture.Image = Global.WinCDS.My.Resources.Resources.DRAWING
        Me.cmdPicture.Location = New System.Drawing.Point(467, 3)
        Me.cmdPicture.Name = "cmdPicture"
        Me.cmdPicture.Size = New System.Drawing.Size(28, 23)
        Me.cmdPicture.TabIndex = 13
        Me.ToolTip1.SetToolTip(Me.cmdPicture, "Insert picture")
        Me.cmdPicture.UseVisualStyleBackColor = True
        '
        'cmdRight
        '
        Me.cmdRight.Image = Global.WinCDS.My.Resources.Resources.RT
        Me.cmdRight.Location = New System.Drawing.Point(439, 3)
        Me.cmdRight.Name = "cmdRight"
        Me.cmdRight.Size = New System.Drawing.Size(28, 23)
        Me.cmdRight.TabIndex = 12
        Me.ToolTip1.SetToolTip(Me.cmdRight, "Right justify")
        Me.cmdRight.UseVisualStyleBackColor = True
        '
        'cmdCenter
        '
        Me.cmdCenter.Image = Global.WinCDS.My.Resources.Resources.CNT
        Me.cmdCenter.Location = New System.Drawing.Point(411, 3)
        Me.cmdCenter.Name = "cmdCenter"
        Me.cmdCenter.Size = New System.Drawing.Size(28, 23)
        Me.cmdCenter.TabIndex = 11
        Me.ToolTip1.SetToolTip(Me.cmdCenter, "Center")
        Me.cmdCenter.UseVisualStyleBackColor = True
        '
        'cmdLeft
        '
        Me.cmdLeft.Image = Global.WinCDS.My.Resources.Resources.LFT
        Me.cmdLeft.Location = New System.Drawing.Point(383, 3)
        Me.cmdLeft.Name = "cmdLeft"
        Me.cmdLeft.Size = New System.Drawing.Size(28, 23)
        Me.cmdLeft.TabIndex = 10
        Me.ToolTip1.SetToolTip(Me.cmdLeft, "Left justify")
        Me.cmdLeft.UseVisualStyleBackColor = True
        '
        'cmdFontColor
        '
        Me.cmdFontColor.Image = Global.WinCDS.My.Resources.Resources.FONTCOLOR
        Me.cmdFontColor.Location = New System.Drawing.Point(355, 3)
        Me.cmdFontColor.Name = "cmdFontColor"
        Me.cmdFontColor.Size = New System.Drawing.Size(28, 23)
        Me.cmdFontColor.TabIndex = 9
        Me.ToolTip1.SetToolTip(Me.cmdFontColor, "Font and color")
        Me.cmdFontColor.UseVisualStyleBackColor = True
        '
        'chkStrikeThru
        '
        Me.chkStrikeThru.Appearance = System.Windows.Forms.Appearance.Button
        Me.chkStrikeThru.Image = Global.WinCDS.My.Resources.Resources.STRIKTHR
        Me.chkStrikeThru.Location = New System.Drawing.Point(200, 3)
        Me.chkStrikeThru.Name = "chkStrikeThru"
        Me.chkStrikeThru.Size = New System.Drawing.Size(28, 23)
        Me.chkStrikeThru.TabIndex = 8
        Me.ToolTip1.SetToolTip(Me.chkStrikeThru, "Strikethrough")
        Me.chkStrikeThru.UseVisualStyleBackColor = True
        '
        'chkUnderline
        '
        Me.chkUnderline.Appearance = System.Windows.Forms.Appearance.Button
        Me.chkUnderline.Image = Global.WinCDS.My.Resources.Resources.UNDRLN
        Me.chkUnderline.Location = New System.Drawing.Point(172, 3)
        Me.chkUnderline.Name = "chkUnderline"
        Me.chkUnderline.Size = New System.Drawing.Size(28, 23)
        Me.chkUnderline.TabIndex = 7
        Me.ToolTip1.SetToolTip(Me.chkUnderline, "Underline")
        Me.chkUnderline.UseVisualStyleBackColor = True
        '
        'chkItalic
        '
        Me.chkItalic.Appearance = System.Windows.Forms.Appearance.Button
        Me.chkItalic.Image = Global.WinCDS.My.Resources.Resources.ITL
        Me.chkItalic.Location = New System.Drawing.Point(144, 3)
        Me.chkItalic.Name = "chkItalic"
        Me.chkItalic.Size = New System.Drawing.Size(28, 23)
        Me.chkItalic.TabIndex = 6
        Me.ToolTip1.SetToolTip(Me.chkItalic, "Italic")
        Me.chkItalic.UseVisualStyleBackColor = True
        '
        'chkBold
        '
        Me.chkBold.Appearance = System.Windows.Forms.Appearance.Button
        Me.chkBold.Image = Global.WinCDS.My.Resources.Resources.BLD
        Me.chkBold.Location = New System.Drawing.Point(116, 3)
        Me.chkBold.Name = "chkBold"
        Me.chkBold.Size = New System.Drawing.Size(28, 23)
        Me.chkBold.TabIndex = 5
        Me.ToolTip1.SetToolTip(Me.chkBold, "Bold")
        Me.chkBold.UseVisualStyleBackColor = True
        '
        'cmdDelete
        '
        Me.cmdDelete.Image = Global.WinCDS.My.Resources.Resources.DELETE
        Me.cmdDelete.Location = New System.Drawing.Point(88, 3)
        Me.cmdDelete.Name = "cmdDelete"
        Me.cmdDelete.Size = New System.Drawing.Size(28, 23)
        Me.cmdDelete.TabIndex = 4
        Me.ToolTip1.SetToolTip(Me.cmdDelete, "Delete")
        Me.cmdDelete.UseVisualStyleBackColor = True
        '
        'cmdPaste
        '
        Me.cmdPaste.Image = Global.WinCDS.My.Resources.Resources.PASTE
        Me.cmdPaste.Location = New System.Drawing.Point(60, 3)
        Me.cmdPaste.Name = "cmdPaste"
        Me.cmdPaste.Size = New System.Drawing.Size(28, 23)
        Me.cmdPaste.TabIndex = 3
        Me.ToolTip1.SetToolTip(Me.cmdPaste, "Paste")
        Me.cmdPaste.UseVisualStyleBackColor = True
        '
        'cmdCopy
        '
        Me.cmdCopy.Image = Global.WinCDS.My.Resources.Resources.COPY
        Me.cmdCopy.Location = New System.Drawing.Point(32, 3)
        Me.cmdCopy.Name = "cmdCopy"
        Me.cmdCopy.Size = New System.Drawing.Size(28, 23)
        Me.cmdCopy.TabIndex = 2
        Me.ToolTip1.SetToolTip(Me.cmdCopy, "Copy")
        Me.cmdCopy.UseVisualStyleBackColor = True
        '
        'cmdCut
        '
        Me.cmdCut.Image = Global.WinCDS.My.Resources.Resources.CUT
        Me.cmdCut.Location = New System.Drawing.Point(4, 3)
        Me.cmdCut.Name = "cmdCut"
        Me.cmdCut.Size = New System.Drawing.Size(28, 23)
        Me.cmdCut.TabIndex = 1
        Me.ToolTip1.SetToolTip(Me.cmdCut, "Cut")
        Me.cmdCut.UseVisualStyleBackColor = True
        '
        'RTB
        '
        Me.RTB.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.RTB.Location = New System.Drawing.Point(4, 31)
        Me.RTB.Name = "RTB"
        Me.RTB.ScrollBars = System.Windows.Forms.RichTextBoxScrollBars.Vertical
        Me.RTB.Size = New System.Drawing.Size(504, 134)
        Me.RTB.TabIndex = 0
        Me.RTB.Text = ""
        '
        'RTBCompose
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Controls.Add(Me.RTB)
        Me.Controls.Add(Me.cmdSend)
        Me.Controls.Add(Me.cmdPicture)
        Me.Controls.Add(Me.cmdRight)
        Me.Controls.Add(Me.cmdCenter)
        Me.Controls.Add(Me.cmdLeft)
        Me.Controls.Add(Me.cmdFontColor)
        Me.Controls.Add(Me.lblFontColor)
        Me.Controls.Add(Me.chkStrikeThru)
        Me.Controls.Add(Me.chkUnderline)
        Me.Controls.Add(Me.chkItalic)
        Me.Controls.Add(Me.chkBold)
        Me.Controls.Add(Me.cmdDelete)
        Me.Controls.Add(Me.cmdPaste)
        Me.Controls.Add(Me.cmdCopy)
        Me.Controls.Add(Me.cmdCut)
        Me.Name = "RTBCompose"
        Me.Size = New System.Drawing.Size(526, 172)
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents cmdCut As Button
    Friend WithEvents cmdCopy As Button
    Friend WithEvents cmdPaste As Button
    Friend WithEvents cmdDelete As Button
    Friend WithEvents chkBold As CheckBox
    Friend WithEvents chkItalic As CheckBox
    Friend WithEvents chkUnderline As CheckBox
    Friend WithEvents chkStrikeThru As CheckBox
    Friend WithEvents lblFontColor As Label
    Friend WithEvents cmdFontColor As Button
    Friend WithEvents cmdLeft As Button
    Friend WithEvents cmdCenter As Button
    Friend WithEvents cmdRight As Button
    Friend WithEvents cmdPicture As Button
    Friend WithEvents cmdSend As Button
    Friend WithEvents RTB As RichTextBox
    Friend WithEvents ToolTip1 As ToolTip
End Class
