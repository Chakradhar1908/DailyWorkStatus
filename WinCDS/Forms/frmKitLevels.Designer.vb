<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class frmKitLevels
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
        Me.txtItemQuan = New System.Windows.Forms.TextBox()
        Me.lblStyle = New System.Windows.Forms.Label()
        Me.txtKitQuantity = New System.Windows.Forms.TextBox()
        Me.cmdStatus = New System.Windows.Forms.Button()
        Me.lblItem = New System.Windows.Forms.Label()
        Me.lblItemLoc = New System.Windows.Forms.Label()
        Me.lblOnOrd = New System.Windows.Forms.Label()
        Me.lblItemAvail = New System.Windows.Forms.Label()
        Me.cmdItemLoc = New System.Windows.Forms.Button()
        Me.cmdItemStatus = New System.Windows.Forms.Button()
        Me.lblItemNum = New System.Windows.Forms.Label()
        Me.fraItems = New System.Windows.Forms.GroupBox()
        Me.lblOnOrdCaption = New System.Windows.Forms.Label()
        Me.lblItemAvailCaption = New System.Windows.Forms.Label()
        Me.lblItemQuanCaption = New System.Windows.Forms.Label()
        Me.lblItemCaption = New System.Windows.Forms.Label()
        Me.lblItemNumCaption = New System.Windows.Forms.Label()
        Me.lblItemLocCaption = New System.Windows.Forms.Label()
        Me.fraControls = New System.Windows.Forms.GroupBox()
        Me.cmdCancel = New System.Windows.Forms.Button()
        Me.cmdOK = New System.Windows.Forms.Button()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.lblKitQuantityCaption = New System.Windows.Forms.Label()
        Me.tmrReload = New System.Windows.Forms.Timer(Me.components)
        Me.fraItems.SuspendLayout()
        Me.fraControls.SuspendLayout()
        Me.SuspendLayout()
        '
        'txtItemQuan
        '
        Me.txtItemQuan.BackColor = System.Drawing.Color.White
        Me.txtItemQuan.Font = New System.Drawing.Font("Lucida Console", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtItemQuan.Location = New System.Drawing.Point(119, 31)
        Me.txtItemQuan.Name = "txtItemQuan"
        Me.txtItemQuan.Size = New System.Drawing.Size(44, 18)
        Me.txtItemQuan.TabIndex = 1
        Me.txtItemQuan.Text = "0"
        Me.txtItemQuan.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'lblStyle
        '
        Me.lblStyle.BackColor = System.Drawing.SystemColors.Window
        Me.lblStyle.Font = New System.Drawing.Font("Lucida Console", 20.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblStyle.Location = New System.Drawing.Point(10, 9)
        Me.lblStyle.Name = "lblStyle"
        Me.lblStyle.Size = New System.Drawing.Size(384, 40)
        Me.lblStyle.TabIndex = 2
        Me.lblStyle.Text = "###"
        Me.lblStyle.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'txtKitQuantity
        '
        Me.txtKitQuantity.Location = New System.Drawing.Point(150, 58)
        Me.txtKitQuantity.Name = "txtKitQuantity"
        Me.txtKitQuantity.Size = New System.Drawing.Size(52, 20)
        Me.txtKitQuantity.TabIndex = 3
        '
        'cmdStatus
        '
        Me.cmdStatus.Location = New System.Drawing.Point(208, 58)
        Me.cmdStatus.Name = "cmdStatus"
        Me.cmdStatus.Size = New System.Drawing.Size(52, 20)
        Me.cmdStatus.TabIndex = 4
        Me.cmdStatus.Text = "ST"
        Me.cmdStatus.UseVisualStyleBackColor = True
        '
        'lblItem
        '
        Me.lblItem.Font = New System.Drawing.Font("Lucida Console", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblItem.Location = New System.Drawing.Point(40, 31)
        Me.lblItem.Name = "lblItem"
        Me.lblItem.Size = New System.Drawing.Size(68, 11)
        Me.lblItem.TabIndex = 5
        Me.lblItem.Text = "### 1 ###"
        Me.lblItem.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'lblItemLoc
        '
        Me.lblItemLoc.Font = New System.Drawing.Font("Lucida Console", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblItemLoc.Location = New System.Drawing.Point(197, 31)
        Me.lblItemLoc.Name = "lblItemLoc"
        Me.lblItemLoc.Size = New System.Drawing.Size(12, 11)
        Me.lblItemLoc.TabIndex = 6
        Me.lblItemLoc.Text = "0"
        Me.lblItemLoc.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'lblOnOrd
        '
        Me.lblOnOrd.Font = New System.Drawing.Font("Lucida Console", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblOnOrd.Location = New System.Drawing.Point(286, 31)
        Me.lblOnOrd.Name = "lblOnOrd"
        Me.lblOnOrd.Size = New System.Drawing.Size(12, 11)
        Me.lblOnOrd.TabIndex = 7
        Me.lblOnOrd.Text = "0"
        Me.lblOnOrd.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'lblItemAvail
        '
        Me.lblItemAvail.Font = New System.Drawing.Font("Lucida Console", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblItemAvail.Location = New System.Drawing.Point(248, 31)
        Me.lblItemAvail.Name = "lblItemAvail"
        Me.lblItemAvail.Size = New System.Drawing.Size(12, 11)
        Me.lblItemAvail.TabIndex = 8
        Me.lblItemAvail.Text = "0"
        Me.lblItemAvail.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'cmdItemLoc
        '
        Me.cmdItemLoc.Font = New System.Drawing.Font("Lucida Console", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdItemLoc.Location = New System.Drawing.Point(307, 28)
        Me.cmdItemLoc.Name = "cmdItemLoc"
        Me.cmdItemLoc.Size = New System.Drawing.Size(30, 20)
        Me.cmdItemLoc.TabIndex = 9
        Me.cmdItemLoc.Text = "L1"
        Me.cmdItemLoc.UseVisualStyleBackColor = True
        '
        'cmdItemStatus
        '
        Me.cmdItemStatus.Font = New System.Drawing.Font("Lucida Console", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdItemStatus.Location = New System.Drawing.Point(342, 28)
        Me.cmdItemStatus.Name = "cmdItemStatus"
        Me.cmdItemStatus.Size = New System.Drawing.Size(30, 20)
        Me.cmdItemStatus.TabIndex = 10
        Me.cmdItemStatus.Text = "ST"
        Me.cmdItemStatus.UseVisualStyleBackColor = True
        '
        'lblItemNum
        '
        Me.lblItemNum.Font = New System.Drawing.Font("Lucida Console", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblItemNum.Location = New System.Drawing.Point(19, 31)
        Me.lblItemNum.Name = "lblItemNum"
        Me.lblItemNum.Size = New System.Drawing.Size(12, 11)
        Me.lblItemNum.TabIndex = 11
        Me.lblItemNum.Text = "1"
        '
        'fraItems
        '
        Me.fraItems.Controls.Add(Me.lblOnOrdCaption)
        Me.fraItems.Controls.Add(Me.lblItemAvailCaption)
        Me.fraItems.Controls.Add(Me.lblItemQuanCaption)
        Me.fraItems.Controls.Add(Me.lblItemCaption)
        Me.fraItems.Controls.Add(Me.lblItemNumCaption)
        Me.fraItems.Controls.Add(Me.lblItemLocCaption)
        Me.fraItems.Controls.Add(Me.cmdItemStatus)
        Me.fraItems.Controls.Add(Me.lblItemNum)
        Me.fraItems.Controls.Add(Me.cmdItemLoc)
        Me.fraItems.Controls.Add(Me.lblItem)
        Me.fraItems.Controls.Add(Me.lblItemAvail)
        Me.fraItems.Controls.Add(Me.txtItemQuan)
        Me.fraItems.Controls.Add(Me.lblOnOrd)
        Me.fraItems.Controls.Add(Me.lblItemLoc)
        Me.fraItems.Location = New System.Drawing.Point(13, 84)
        Me.fraItems.Name = "fraItems"
        Me.fraItems.Size = New System.Drawing.Size(381, 92)
        Me.fraItems.TabIndex = 12
        Me.fraItems.TabStop = False
        '
        'lblOnOrdCaption
        '
        Me.lblOnOrdCaption.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblOnOrdCaption.Location = New System.Drawing.Point(271, 16)
        Me.lblOnOrdCaption.Name = "lblOnOrdCaption"
        Me.lblOnOrdCaption.Size = New System.Drawing.Size(27, 13)
        Me.lblOnOrdCaption.TabIndex = 16
        Me.lblOnOrdCaption.Text = "Ord"
        Me.lblOnOrdCaption.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'lblItemAvailCaption
        '
        Me.lblItemAvailCaption.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblItemAvailCaption.Location = New System.Drawing.Point(225, 16)
        Me.lblItemAvailCaption.Name = "lblItemAvailCaption"
        Me.lblItemAvailCaption.Size = New System.Drawing.Size(35, 13)
        Me.lblItemAvailCaption.TabIndex = 15
        Me.lblItemAvailCaption.Text = "Avail"
        Me.lblItemAvailCaption.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'lblItemQuanCaption
        '
        Me.lblItemQuanCaption.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblItemQuanCaption.Location = New System.Drawing.Point(126, 16)
        Me.lblItemQuanCaption.Name = "lblItemQuanCaption"
        Me.lblItemQuanCaption.Size = New System.Drawing.Size(37, 13)
        Me.lblItemQuanCaption.TabIndex = 2
        Me.lblItemQuanCaption.Text = "Quan"
        Me.lblItemQuanCaption.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'lblItemCaption
        '
        Me.lblItemCaption.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblItemCaption.Location = New System.Drawing.Point(45, 16)
        Me.lblItemCaption.Name = "lblItemCaption"
        Me.lblItemCaption.Size = New System.Drawing.Size(63, 13)
        Me.lblItemCaption.TabIndex = 1
        Me.lblItemCaption.Text = "Item Style"
        Me.lblItemCaption.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'lblItemNumCaption
        '
        Me.lblItemNumCaption.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblItemNumCaption.Location = New System.Drawing.Point(17, 16)
        Me.lblItemNumCaption.Name = "lblItemNumCaption"
        Me.lblItemNumCaption.Size = New System.Drawing.Size(15, 13)
        Me.lblItemNumCaption.TabIndex = 0
        Me.lblItemNumCaption.Text = "#"
        '
        'lblItemLocCaption
        '
        Me.lblItemLocCaption.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblItemLocCaption.Location = New System.Drawing.Point(186, 16)
        Me.lblItemLocCaption.Name = "lblItemLocCaption"
        Me.lblItemLocCaption.Size = New System.Drawing.Size(23, 13)
        Me.lblItemLocCaption.TabIndex = 14
        Me.lblItemLocCaption.Text = "ST"
        Me.lblItemLocCaption.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'fraControls
        '
        Me.fraControls.Controls.Add(Me.cmdCancel)
        Me.fraControls.Controls.Add(Me.cmdOK)
        Me.fraControls.Location = New System.Drawing.Point(124, 182)
        Me.fraControls.Name = "fraControls"
        Me.fraControls.Size = New System.Drawing.Size(154, 56)
        Me.fraControls.TabIndex = 13
        Me.fraControls.TabStop = False
        '
        'cmdCancel
        '
        Me.cmdCancel.Location = New System.Drawing.Point(84, 13)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.Size = New System.Drawing.Size(64, 38)
        Me.cmdCancel.TabIndex = 19
        Me.cmdCancel.Text = "Cancel"
        Me.cmdCancel.UseVisualStyleBackColor = True
        '
        'cmdOK
        '
        Me.cmdOK.Location = New System.Drawing.Point(9, 13)
        Me.cmdOK.Name = "cmdOK"
        Me.cmdOK.Size = New System.Drawing.Size(64, 38)
        Me.cmdOK.TabIndex = 18
        Me.cmdOK.Text = "OK"
        Me.cmdOK.UseVisualStyleBackColor = True
        '
        'lblKitQuantityCaption
        '
        Me.lblKitQuantityCaption.AutoSize = True
        Me.lblKitQuantityCaption.Location = New System.Drawing.Point(100, 61)
        Me.lblKitQuantityCaption.Name = "lblKitQuantityCaption"
        Me.lblKitQuantityCaption.Size = New System.Drawing.Size(49, 13)
        Me.lblKitQuantityCaption.TabIndex = 18
        Me.lblKitQuantityCaption.Text = "Quantity:"
        '
        'frmKitLevels
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(403, 250)
        Me.Controls.Add(Me.lblKitQuantityCaption)
        Me.Controls.Add(Me.fraControls)
        Me.Controls.Add(Me.fraItems)
        Me.Controls.Add(Me.cmdStatus)
        Me.Controls.Add(Me.txtKitQuantity)
        Me.Controls.Add(Me.lblStyle)
        Me.Name = "frmKitLevels"
        Me.Text = "Kit Stock Levels"
        Me.fraItems.ResumeLayout(False)
        Me.fraItems.PerformLayout()
        Me.fraControls.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents txtItemQuan As TextBox
    Friend WithEvents lblStyle As Label
    Friend WithEvents txtKitQuantity As TextBox
    Friend WithEvents cmdStatus As Button
    Friend WithEvents lblItem As Label
    Friend WithEvents lblItemLoc As Label
    Friend WithEvents lblOnOrd As Label
    Friend WithEvents lblItemAvail As Label
    Friend WithEvents cmdItemLoc As Button
    Friend WithEvents cmdItemStatus As Button
    Friend WithEvents lblItemNum As Label
    Friend WithEvents fraItems As GroupBox
    Friend WithEvents fraControls As GroupBox
    Friend WithEvents lblItemLocCaption As Label
    Friend WithEvents ToolTip1 As ToolTip
    Friend WithEvents lblOnOrdCaption As Label
    Friend WithEvents lblItemAvailCaption As Label
    Friend WithEvents lblItemQuanCaption As Label
    Friend WithEvents lblItemCaption As Label
    Friend WithEvents lblItemNumCaption As Label
    Friend WithEvents cmdCancel As Button
    Friend WithEvents cmdOK As Button
    Friend WithEvents lblKitQuantityCaption As Label
    Friend WithEvents tmrReload As Timer
End Class
