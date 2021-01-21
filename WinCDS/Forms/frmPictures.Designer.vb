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
        Me.lblRef = New System.Windows.Forms.Label()
        Me.lblPictures = New System.Windows.Forms.Label()
        Me.fraPicture = New System.Windows.Forms.GroupBox()
        Me.imgPicture = New System.Windows.Forms.PictureBox()
        Me.lblCaption = New System.Windows.Forms.Label()
        Me.lblPictureID = New System.Windows.Forms.Label()
        Me.txtPictureID = New System.Windows.Forms.TextBox()
        Me.lblPictureType = New System.Windows.Forms.Label()
        Me.txtPictureType = New System.Windows.Forms.TextBox()
        Me.lblPictureRef = New System.Windows.Forms.Label()
        Me.txtPictureRef = New System.Windows.Forms.TextBox()
        Me.txtCaption = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.cmdMoveFirst = New System.Windows.Forms.Button()
        Me.cmdMovePrevious = New System.Windows.Forms.Button()
        Me.cmdMoveNext = New System.Windows.Forms.Button()
        Me.cmdMoveLast = New System.Windows.Forms.Button()
        Me.cmdDelete = New System.Windows.Forms.Button()
        Me.cmdAdd = New System.Windows.Forms.Button()
        Me.cmdPrint = New System.Windows.Forms.Button()
        Me.cmdOK = New System.Windows.Forms.Button()
        Me.fraPicture.SuspendLayout()
        CType(Me.imgPicture, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'lblRef
        '
        Me.lblRef.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.lblRef.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblRef.Font = New System.Drawing.Font("Courier New", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblRef.Location = New System.Drawing.Point(16, 16)
        Me.lblRef.Name = "lblRef"
        Me.lblRef.Size = New System.Drawing.Size(312, 28)
        Me.lblRef.TabIndex = 0
        Me.lblRef.Text = "###"
        '
        'lblPictures
        '
        Me.lblPictures.AutoSize = True
        Me.lblPictures.Location = New System.Drawing.Point(712, 342)
        Me.lblPictures.Name = "lblPictures"
        Me.lblPictures.Size = New System.Drawing.Size(55, 13)
        Me.lblPictures.TabIndex = 1
        Me.lblPictures.Text = "lblPictures"
        '
        'fraPicture
        '
        Me.fraPicture.Controls.Add(Me.cmdMoveLast)
        Me.fraPicture.Controls.Add(Me.cmdMoveNext)
        Me.fraPicture.Controls.Add(Me.cmdMovePrevious)
        Me.fraPicture.Controls.Add(Me.cmdMoveFirst)
        Me.fraPicture.Controls.Add(Me.Label1)
        Me.fraPicture.Controls.Add(Me.txtCaption)
        Me.fraPicture.Controls.Add(Me.txtPictureRef)
        Me.fraPicture.Controls.Add(Me.lblPictureRef)
        Me.fraPicture.Controls.Add(Me.txtPictureType)
        Me.fraPicture.Controls.Add(Me.lblPictureType)
        Me.fraPicture.Controls.Add(Me.txtPictureID)
        Me.fraPicture.Controls.Add(Me.lblPictureID)
        Me.fraPicture.Controls.Add(Me.lblCaption)
        Me.fraPicture.Controls.Add(Me.imgPicture)
        Me.fraPicture.Controls.Add(Me.lblRef)
        Me.fraPicture.Location = New System.Drawing.Point(107, 36)
        Me.fraPicture.Name = "fraPicture"
        Me.fraPicture.Size = New System.Drawing.Size(394, 319)
        Me.fraPicture.TabIndex = 2
        Me.fraPicture.TabStop = False
        '
        'imgPicture
        '
        Me.imgPicture.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.imgPicture.Location = New System.Drawing.Point(16, 61)
        Me.imgPicture.Name = "imgPicture"
        Me.imgPicture.Size = New System.Drawing.Size(281, 125)
        Me.imgPicture.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.imgPicture.TabIndex = 1
        Me.imgPicture.TabStop = False
        '
        'lblCaption
        '
        Me.lblCaption.AutoSize = True
        Me.lblCaption.Location = New System.Drawing.Point(13, 200)
        Me.lblCaption.Name = "lblCaption"
        Me.lblCaption.Size = New System.Drawing.Size(46, 13)
        Me.lblCaption.TabIndex = 2
        Me.lblCaption.Text = "Capt&ion:"
        '
        'lblPictureID
        '
        Me.lblPictureID.AutoSize = True
        Me.lblPictureID.Location = New System.Drawing.Point(65, 200)
        Me.lblPictureID.Name = "lblPictureID"
        Me.lblPictureID.Size = New System.Drawing.Size(21, 13)
        Me.lblPictureID.TabIndex = 3
        Me.lblPictureID.Text = "ID:"
        '
        'txtPictureID
        '
        Me.txtPictureID.Location = New System.Drawing.Point(83, 197)
        Me.txtPictureID.Name = "txtPictureID"
        Me.txtPictureID.Size = New System.Drawing.Size(28, 20)
        Me.txtPictureID.TabIndex = 4
        '
        'lblPictureType
        '
        Me.lblPictureType.AutoSize = True
        Me.lblPictureType.Location = New System.Drawing.Point(116, 200)
        Me.lblPictureType.Name = "lblPictureType"
        Me.lblPictureType.Size = New System.Drawing.Size(34, 13)
        Me.lblPictureType.TabIndex = 5
        Me.lblPictureType.Text = "Type:"
        '
        'txtPictureType
        '
        Me.txtPictureType.Location = New System.Drawing.Point(149, 200)
        Me.txtPictureType.Name = "txtPictureType"
        Me.txtPictureType.Size = New System.Drawing.Size(32, 20)
        Me.txtPictureType.TabIndex = 6
        '
        'lblPictureRef
        '
        Me.lblPictureRef.AutoSize = True
        Me.lblPictureRef.Location = New System.Drawing.Point(187, 200)
        Me.lblPictureRef.Name = "lblPictureRef"
        Me.lblPictureRef.Size = New System.Drawing.Size(27, 13)
        Me.lblPictureRef.TabIndex = 7
        Me.lblPictureRef.Text = "Ref:"
        '
        'txtPictureRef
        '
        Me.txtPictureRef.Location = New System.Drawing.Point(211, 197)
        Me.txtPictureRef.Name = "txtPictureRef"
        Me.txtPictureRef.Size = New System.Drawing.Size(86, 20)
        Me.txtPictureRef.TabIndex = 8
        '
        'txtCaption
        '
        Me.txtCaption.Location = New System.Drawing.Point(16, 223)
        Me.txtCaption.Multiline = True
        Me.txtCaption.Name = "txtCaption"
        Me.txtCaption.Size = New System.Drawing.Size(281, 51)
        Me.txtCaption.TabIndex = 9
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.SystemColors.ButtonHighlight
        Me.Label1.Location = New System.Drawing.Point(65, 289)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(215, 17)
        Me.Label1.TabIndex = 10
        Me.Label1.Text = "Datacontrol"
        '
        'cmdMoveFirst
        '
        Me.cmdMoveFirst.Location = New System.Drawing.Point(6, 283)
        Me.cmdMoveFirst.Name = "cmdMoveFirst"
        Me.cmdMoveFirst.Size = New System.Drawing.Size(32, 30)
        Me.cmdMoveFirst.TabIndex = 11
        Me.cmdMoveFirst.Text = "Button1"
        Me.cmdMoveFirst.UseVisualStyleBackColor = True
        '
        'cmdMovePrevious
        '
        Me.cmdMovePrevious.Location = New System.Drawing.Point(37, 283)
        Me.cmdMovePrevious.Name = "cmdMovePrevious"
        Me.cmdMovePrevious.Size = New System.Drawing.Size(22, 23)
        Me.cmdMovePrevious.TabIndex = 12
        Me.cmdMovePrevious.Text = "Button2"
        Me.cmdMovePrevious.UseVisualStyleBackColor = True
        '
        'cmdMoveNext
        '
        Me.cmdMoveNext.Location = New System.Drawing.Point(286, 284)
        Me.cmdMoveNext.Name = "cmdMoveNext"
        Me.cmdMoveNext.Size = New System.Drawing.Size(22, 23)
        Me.cmdMoveNext.TabIndex = 13
        Me.cmdMoveNext.Text = "Button3"
        Me.cmdMoveNext.UseVisualStyleBackColor = True
        '
        'cmdMoveLast
        '
        Me.cmdMoveLast.Location = New System.Drawing.Point(314, 260)
        Me.cmdMoveLast.Name = "cmdMoveLast"
        Me.cmdMoveLast.Size = New System.Drawing.Size(53, 46)
        Me.cmdMoveLast.TabIndex = 14
        Me.cmdMoveLast.Text = "Button4"
        Me.cmdMoveLast.UseVisualStyleBackColor = True
        '
        'cmdDelete
        '
        Me.cmdDelete.Location = New System.Drawing.Point(123, 374)
        Me.cmdDelete.Name = "cmdDelete"
        Me.cmdDelete.Size = New System.Drawing.Size(75, 48)
        Me.cmdDelete.TabIndex = 3
        Me.cmdDelete.Text = "&Delete"
        Me.cmdDelete.UseVisualStyleBackColor = True
        '
        'cmdAdd
        '
        Me.cmdAdd.Location = New System.Drawing.Point(213, 374)
        Me.cmdAdd.Name = "cmdAdd"
        Me.cmdAdd.Size = New System.Drawing.Size(75, 48)
        Me.cmdAdd.TabIndex = 4
        Me.cmdAdd.Text = "&Add"
        Me.cmdAdd.UseVisualStyleBackColor = True
        '
        'cmdPrint
        '
        Me.cmdPrint.Location = New System.Drawing.Point(297, 374)
        Me.cmdPrint.Name = "cmdPrint"
        Me.cmdPrint.Size = New System.Drawing.Size(75, 48)
        Me.cmdPrint.TabIndex = 5
        Me.cmdPrint.Text = "&Print"
        Me.cmdPrint.UseVisualStyleBackColor = True
        '
        'cmdOK
        '
        Me.cmdOK.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cmdOK.Location = New System.Drawing.Point(378, 374)
        Me.cmdOK.Name = "cmdOK"
        Me.cmdOK.Size = New System.Drawing.Size(75, 48)
        Me.cmdOK.TabIndex = 6
        Me.cmdOK.Text = "&OK"
        Me.cmdOK.UseVisualStyleBackColor = True
        '
        'frmPictures
        '
        Me.AcceptButton = Me.cmdOK
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.cmdOK
        Me.ClientSize = New System.Drawing.Size(800, 450)
        Me.Controls.Add(Me.cmdOK)
        Me.Controls.Add(Me.cmdPrint)
        Me.Controls.Add(Me.cmdAdd)
        Me.Controls.Add(Me.cmdDelete)
        Me.Controls.Add(Me.fraPicture)
        Me.Controls.Add(Me.lblPictures)
        Me.MinimizeBox = False
        Me.Name = "frmPictures"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.fraPicture.ResumeLayout(False)
        Me.fraPicture.PerformLayout()
        CType(Me.imgPicture, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

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
    Friend WithEvents Label1 As Label
    Friend WithEvents cmdDelete As Button
    Friend WithEvents cmdAdd As Button
    Friend WithEvents cmdPrint As Button
    Friend WithEvents cmdOK As Button
End Class
