<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class MessageCenter
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(MessageCenter))
        Me.txt = New System.Windows.Forms.TextBox()
        Me.img = New System.Windows.Forms.PictureBox()
        Me.iml = New System.Windows.Forms.ImageList(Me.components)
        Me.oX = New System.Windows.Forms.PictureBox()
        Me.X = New System.Windows.Forms.PictureBox()
        CType(Me.img, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.oX, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.X, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'txt
        '
        Me.txt.Location = New System.Drawing.Point(25, 13)
        Me.txt.Multiline = True
        Me.txt.Name = "txt"
        Me.txt.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.txt.Size = New System.Drawing.Size(237, 47)
        Me.txt.TabIndex = 0
        '
        'img
        '
        Me.img.BackColor = System.Drawing.SystemColors.Control
        Me.img.Location = New System.Drawing.Point(25, 93)
        Me.img.Name = "img"
        Me.img.Size = New System.Drawing.Size(237, 50)
        Me.img.TabIndex = 1
        Me.img.TabStop = False
        '
        'iml
        '
        Me.iml.ImageStream = CType(resources.GetObject("iml.ImageStream"), System.Windows.Forms.ImageListStreamer)
        Me.iml.TransparentColor = System.Drawing.Color.Transparent
        Me.iml.Images.SetKeyName(0, "CC.jpg")
        Me.iml.Images.SetKeyName(1, "CCAd2.bmp")
        Me.iml.Images.SetKeyName(2, "CloudAd.bmp")
        Me.iml.Images.SetKeyName(3, "DTAd.jpg")
        Me.iml.Images.SetKeyName(4, "x.bmp")
        Me.iml.Images.SetKeyName(5, "TEST.bmp")
        '
        'oX
        '
        Me.oX.BackColor = System.Drawing.SystemColors.Desktop
        Me.oX.Image = Global.WinCDS.My.Resources.Resources._STOP
        Me.oX.Location = New System.Drawing.Point(282, 13)
        Me.oX.Name = "oX"
        Me.oX.Size = New System.Drawing.Size(24, 23)
        Me.oX.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.oX.TabIndex = 2
        Me.oX.TabStop = False
        '
        'X
        '
        Me.X.BackColor = System.Drawing.SystemColors.Desktop
        Me.X.Image = Global.WinCDS.My.Resources.Resources._STOP
        Me.X.Location = New System.Drawing.Point(282, 51)
        Me.X.Name = "X"
        Me.X.Size = New System.Drawing.Size(24, 23)
        Me.X.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.X.TabIndex = 3
        Me.X.TabStop = False
        '
        'MessageCenter
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.Controls.Add(Me.X)
        Me.Controls.Add(Me.oX)
        Me.Controls.Add(Me.img)
        Me.Controls.Add(Me.txt)
        Me.Name = "MessageCenter"
        Me.Size = New System.Drawing.Size(318, 188)
        CType(Me.img, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.oX, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.X, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents txt As TextBox
    Friend WithEvents img As PictureBox
    Friend WithEvents iml As ImageList
    Friend WithEvents oX As PictureBox
    Friend WithEvents X As PictureBox
End Class
