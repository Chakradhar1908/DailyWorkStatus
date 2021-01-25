<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class frmPictures
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
        Me.components = New System.ComponentModel.Container()
        Me.lblRef = New System.Windows.Forms.Label()
        Me.lblPictures = New System.Windows.Forms.Label()
        Me.fraPicture = New System.Windows.Forms.GroupBox()
        Me.cmdMoveLast = New System.Windows.Forms.Button()
        Me.cmdMoveNext = New System.Windows.Forms.Button()
        Me.cmdMovePrevious = New System.Windows.Forms.Button()
        Me.cmdMoveFirst = New System.Windows.Forms.Button()
        Me.txtCaption = New System.Windows.Forms.TextBox()
        Me.txtPictureRef = New System.Windows.Forms.TextBox()
        Me.lblPictureRef = New System.Windows.Forms.Label()
        Me.txtPictureType = New System.Windows.Forms.TextBox()
        Me.lblPictureType = New System.Windows.Forms.Label()
        Me.txtPictureID = New System.Windows.Forms.TextBox()
        Me.lblPictureID = New System.Windows.Forms.Label()
        Me.lblCaption = New System.Windows.Forms.Label()
        Me.imgPicture = New System.Windows.Forms.PictureBox()
        Me.cmdDelete = New System.Windows.Forms.Button()
        Me.cmdAdd = New System.Windows.Forms.Button()
        Me.cmdPrint = New System.Windows.Forms.Button()
        Me.cmdOK = New System.Windows.Forms.Button()
        Me.imgReveal = New System.Windows.Forms.PictureBox()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.fraPicture.SuspendLayout()
        CType(Me.imgPicture, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.imgReveal, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'lblRef
        '
        Me.lblRef.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.lblRef.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblRef.Font = New System.Drawing.Font("Courier New", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblRef.Location = New System.Drawing.Point(16, 16)
        Me.lblRef.Name = "lblRef"
        Me.lblRef.Size = New System.Drawing.Size(281, 28)
        Me.lblRef.TabIndex = 0
        Me.lblRef.Text = "###"
        '
        'lblPictures
        '
        Me.lblPictures.Location = New System.Drawing.Point(64, 262)
        Me.lblPictures.Name = "lblPictures"
        Me.lblPictures.Size = New System.Drawing.Size(181, 20)
        Me.lblPictures.TabIndex = 1
        '
        'fraPicture
        '
        Me.fraPicture.Controls.Add(Me.cmdMoveLast)
        Me.fraPicture.Controls.Add(Me.cmdMoveNext)
        Me.fraPicture.Controls.Add(Me.cmdMovePrevious)
        Me.fraPicture.Controls.Add(Me.cmdMoveFirst)
        Me.fraPicture.Controls.Add(Me.txtCaption)
        Me.fraPicture.Controls.Add(Me.lblPictures)
        Me.fraPicture.Controls.Add(Me.txtPictureRef)
        Me.fraPicture.Controls.Add(Me.lblPictureRef)
        Me.fraPicture.Controls.Add(Me.txtPictureType)
        Me.fraPicture.Controls.Add(Me.lblPictureType)
        Me.fraPicture.Controls.Add(Me.txtPictureID)
        Me.fraPicture.Controls.Add(Me.lblPictureID)
        Me.fraPicture.Controls.Add(Me.lblCaption)
        Me.fraPicture.Controls.Add(Me.imgPicture)
        Me.fraPicture.Controls.Add(Me.lblRef)
        Me.fraPicture.Location = New System.Drawing.Point(9, 3)
        Me.fraPicture.Name = "fraPicture"
        Me.fraPicture.Size = New System.Drawing.Size(312, 295)
        Me.fraPicture.TabIndex = 2
        Me.fraPicture.TabStop = False
        '
        'cmdMoveLast
        '
        Me.cmdMoveLast.Location = New System.Drawing.Point(274, 256)
        Me.cmdMoveLast.Name = "cmdMoveLast"
        Me.cmdMoveLast.Size = New System.Drawing.Size(25, 32)
        Me.cmdMoveLast.TabIndex = 14
        Me.ToolTip1.SetToolTip(Me.cmdMoveLast, " Move To The Last Record ")
        Me.cmdMoveLast.UseVisualStyleBackColor = True
        '
        'cmdMoveNext
        '
        Me.cmdMoveNext.Location = New System.Drawing.Point(251, 256)
        Me.cmdMoveNext.Name = "cmdMoveNext"
        Me.cmdMoveNext.Size = New System.Drawing.Size(25, 32)
        Me.cmdMoveNext.TabIndex = 13
        Me.ToolTip1.SetToolTip(Me.cmdMoveNext, " Move Forward 1 Record ")
        Me.cmdMoveNext.UseVisualStyleBackColor = True
        '
        'cmdMovePrevious
        '
        Me.cmdMovePrevious.Location = New System.Drawing.Point(35, 256)
        Me.cmdMovePrevious.Name = "cmdMovePrevious"
        Me.cmdMovePrevious.Size = New System.Drawing.Size(25, 32)
        Me.cmdMovePrevious.TabIndex = 12
        Me.ToolTip1.SetToolTip(Me.cmdMovePrevious, " Move Back 1 Record ")
        Me.cmdMovePrevious.UseVisualStyleBackColor = True
        '
        'cmdMoveFirst
        '
        Me.cmdMoveFirst.Location = New System.Drawing.Point(13, 256)
        Me.cmdMoveFirst.Name = "cmdMoveFirst"
        Me.cmdMoveFirst.Size = New System.Drawing.Size(25, 32)
        Me.cmdMoveFirst.TabIndex = 11
        Me.ToolTip1.SetToolTip(Me.cmdMoveFirst, " Move To The First Record ")
        Me.cmdMoveFirst.UseVisualStyleBackColor = True
        '
        'txtCaption
        '
        Me.txtCaption.Location = New System.Drawing.Point(16, 199)
        Me.txtCaption.Multiline = True
        Me.txtCaption.Name = "txtCaption"
        Me.txtCaption.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.txtCaption.Size = New System.Drawing.Size(281, 51)
        Me.txtCaption.TabIndex = 9
        '
        'txtPictureRef
        '
        Me.txtPictureRef.Location = New System.Drawing.Point(211, 178)
        Me.txtPictureRef.Name = "txtPictureRef"
        Me.txtPictureRef.Size = New System.Drawing.Size(86, 20)
        Me.txtPictureRef.TabIndex = 8
        Me.ToolTip1.SetToolTip(Me.txtPictureRef, "PictureRef")
        Me.txtPictureRef.Visible = False
        '
        'lblPictureRef
        '
        Me.lblPictureRef.AutoSize = True
        Me.lblPictureRef.Location = New System.Drawing.Point(187, 181)
        Me.lblPictureRef.Name = "lblPictureRef"
        Me.lblPictureRef.Size = New System.Drawing.Size(27, 13)
        Me.lblPictureRef.TabIndex = 7
        Me.lblPictureRef.Text = "Ref:"
        Me.lblPictureRef.Visible = False
        '
        'txtPictureType
        '
        Me.txtPictureType.Location = New System.Drawing.Point(149, 181)
        Me.txtPictureType.Name = "txtPictureType"
        Me.txtPictureType.Size = New System.Drawing.Size(32, 20)
        Me.txtPictureType.TabIndex = 6
        Me.ToolTip1.SetToolTip(Me.txtPictureType, "PictureType")
        Me.txtPictureType.Visible = False
        '
        'lblPictureType
        '
        Me.lblPictureType.AutoSize = True
        Me.lblPictureType.Location = New System.Drawing.Point(116, 181)
        Me.lblPictureType.Name = "lblPictureType"
        Me.lblPictureType.Size = New System.Drawing.Size(34, 13)
        Me.lblPictureType.TabIndex = 5
        Me.lblPictureType.Text = "Type:"
        Me.lblPictureType.Visible = False
        '
        'txtPictureID
        '
        Me.txtPictureID.Location = New System.Drawing.Point(83, 178)
        Me.txtPictureID.Name = "txtPictureID"
        Me.txtPictureID.Size = New System.Drawing.Size(28, 20)
        Me.txtPictureID.TabIndex = 4
        Me.ToolTip1.SetToolTip(Me.txtPictureID, "PictureID")
        Me.txtPictureID.Visible = False
        '
        'lblPictureID
        '
        Me.lblPictureID.AutoSize = True
        Me.lblPictureID.Location = New System.Drawing.Point(65, 181)
        Me.lblPictureID.Name = "lblPictureID"
        Me.lblPictureID.Size = New System.Drawing.Size(21, 13)
        Me.lblPictureID.TabIndex = 3
        Me.lblPictureID.Text = "ID:"
        Me.lblPictureID.Visible = False
        '
        'lblCaption
        '
        Me.lblCaption.AutoSize = True
        Me.lblCaption.Location = New System.Drawing.Point(16, 181)
        Me.lblCaption.Name = "lblCaption"
        Me.lblCaption.Size = New System.Drawing.Size(46, 13)
        Me.lblCaption.TabIndex = 2
        Me.lblCaption.Text = "Capt&ion:"
        '
        'imgPicture
        '
        Me.imgPicture.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.imgPicture.Location = New System.Drawing.Point(16, 51)
        Me.imgPicture.Name = "imgPicture"
        Me.imgPicture.Size = New System.Drawing.Size(281, 125)
        Me.imgPicture.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.imgPicture.TabIndex = 1
        Me.imgPicture.TabStop = False
        '
        'cmdDelete
        '
        Me.cmdDelete.Location = New System.Drawing.Point(12, 301)
        Me.cmdDelete.Name = "cmdDelete"
        Me.cmdDelete.Size = New System.Drawing.Size(71, 51)
        Me.cmdDelete.TabIndex = 3
        Me.cmdDelete.Text = "&Delete"
        Me.cmdDelete.UseVisualStyleBackColor = True
        '
        'cmdAdd
        '
        Me.cmdAdd.Location = New System.Drawing.Point(91, 301)
        Me.cmdAdd.Name = "cmdAdd"
        Me.cmdAdd.Size = New System.Drawing.Size(71, 51)
        Me.cmdAdd.TabIndex = 4
        Me.cmdAdd.Text = "&Add"
        Me.cmdAdd.UseVisualStyleBackColor = True
        '
        'cmdPrint
        '
        Me.cmdPrint.Location = New System.Drawing.Point(170, 301)
        Me.cmdPrint.Name = "cmdPrint"
        Me.cmdPrint.Size = New System.Drawing.Size(71, 51)
        Me.cmdPrint.TabIndex = 5
        Me.cmdPrint.Text = "&Print"
        Me.cmdPrint.UseVisualStyleBackColor = True
        '
        'cmdOK
        '
        Me.cmdOK.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cmdOK.Location = New System.Drawing.Point(249, 301)
        Me.cmdOK.Name = "cmdOK"
        Me.cmdOK.Size = New System.Drawing.Size(71, 51)
        Me.cmdOK.TabIndex = 6
        Me.cmdOK.Text = "&OK"
        Me.cmdOK.UseVisualStyleBackColor = True
        '
        'imgReveal
        '
        Me.imgReveal.Location = New System.Drawing.Point(0, -3)
        Me.imgReveal.Name = "imgReveal"
        Me.imgReveal.Size = New System.Drawing.Size(10, 10)
        Me.imgReveal.TabIndex = 7
        Me.imgReveal.TabStop = False
        '
        'frmPictures
        '
        Me.AcceptButton = Me.cmdOK
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.cmdOK
        Me.ClientSize = New System.Drawing.Size(329, 357)
        Me.Controls.Add(Me.imgReveal)
        Me.Controls.Add(Me.cmdOK)
        Me.Controls.Add(Me.cmdPrint)
        Me.Controls.Add(Me.cmdAdd)
        Me.Controls.Add(Me.cmdDelete)
        Me.Controls.Add(Me.fraPicture)
        Me.MinimizeBox = False
        Me.Name = "frmPictures"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.fraPicture.ResumeLayout(False)
        Me.fraPicture.PerformLayout()
        CType(Me.imgPicture, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.imgReveal, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents lblRef As Label
    Friend WithEvents lblPictures As Label
    Friend WithEvents fraPicture As GroupBox
    Friend WithEvents txtCaption As TextBox
    Friend WithEvents txtPictureRef As TextBox
    Friend WithEvents lblPictureRef As Label
    Friend WithEvents txtPictureType As TextBox
    Friend WithEvents lblPictureType As Label
    Friend WithEvents txtPictureID As TextBox
    Friend WithEvents lblPictureID As Label
    Friend WithEvents lblCaption As Label
    Friend WithEvents imgPicture As PictureBox
    Friend WithEvents cmdMoveLast As Button
    Friend WithEvents cmdMoveNext As Button
    Friend WithEvents cmdMovePrevious As Button
    Friend WithEvents cmdMoveFirst As Button
    Friend WithEvents cmdDelete As Button
    Friend WithEvents cmdAdd As Button
    Friend WithEvents cmdPrint As Button
    Friend WithEvents cmdOK As Button
    Friend WithEvents imgReveal As PictureBox
    Friend WithEvents ToolTip1 As ToolTip
End Class
