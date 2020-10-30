<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class SelectPrinter
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(SelectPrinter))
        Me.lblDisplayStyle = New System.Windows.Forms.Label()
        Me.fra = New System.Windows.Forms.GroupBox()
        Me.lblCustomCaption = New System.Windows.Forms.Label()
        Me.cboCustomTagTemplate = New System.Windows.Forms.ComboBox()
        Me.lblExtraRecLbl = New System.Windows.Forms.Label()
        Me.cmdRecLbl = New System.Windows.Forms.Button()
        Me.lblOrientation = New System.Windows.Forms.Label()
        Me.cboTagJustify = New System.Windows.Forms.ComboBox()
        Me.lblQty = New System.Windows.Forms.Label()
        Me.txtCopies = New System.Windows.Forms.TextBox()
        Me.updQuantity = New AxMSComCtl2.AxUpDown()
        Me.cmdCustom = New System.Windows.Forms.Button()
        Me.cmdLarge = New System.Windows.Forms.Button()
        Me.cmdMedium = New System.Windows.Forms.Button()
        Me.cmdSmall = New System.Windows.Forms.Button()
        Me.cmdDYMO = New System.Windows.Forms.Button()
        Me.cmdDone = New System.Windows.Forms.Button()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.pic = New System.Windows.Forms.PictureBox()
        Me.PrSel = New WinCDS.PrinterSelector()
        Me.fra.SuspendLayout()
        CType(Me.updQuantity, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.pic, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'lblDisplayStyle
        '
        Me.lblDisplayStyle.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.lblDisplayStyle.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblDisplayStyle.Font = New System.Drawing.Font("Arial", 15.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblDisplayStyle.Location = New System.Drawing.Point(6, 3)
        Me.lblDisplayStyle.Name = "lblDisplayStyle"
        Me.lblDisplayStyle.Size = New System.Drawing.Size(329, 29)
        Me.lblDisplayStyle.TabIndex = 0
        Me.lblDisplayStyle.Text = "##"
        Me.lblDisplayStyle.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'fra
        '
        Me.fra.Controls.Add(Me.PrSel)
        Me.fra.Location = New System.Drawing.Point(6, 37)
        Me.fra.Name = "fra"
        Me.fra.Size = New System.Drawing.Size(329, 135)
        Me.fra.TabIndex = 1
        Me.fra.TabStop = False
        Me.fra.Text = "Select Pr&inter"
        '
        'lblCustomCaption
        '
        Me.lblCustomCaption.AutoSize = True
        Me.lblCustomCaption.Location = New System.Drawing.Point(6, 184)
        Me.lblCustomCaption.Name = "lblCustomCaption"
        Me.lblCustomCaption.Size = New System.Drawing.Size(67, 13)
        Me.lblCustomCaption.TabIndex = 2
        Me.lblCustomCaption.Text = "Custom Tag:"
        '
        'cboCustomTagTemplate
        '
        Me.cboCustomTagTemplate.FormattingEnabled = True
        Me.cboCustomTagTemplate.Items.AddRange(New Object() {"Center", "Left", "Right"})
        Me.cboCustomTagTemplate.Location = New System.Drawing.Point(6, 200)
        Me.cboCustomTagTemplate.Name = "cboCustomTagTemplate"
        Me.cboCustomTagTemplate.Size = New System.Drawing.Size(100, 21)
        Me.cboCustomTagTemplate.TabIndex = 3
        '
        'lblExtraRecLbl
        '
        Me.lblExtraRecLbl.AutoSize = True
        Me.lblExtraRecLbl.Location = New System.Drawing.Point(112, 184)
        Me.lblExtraRecLbl.Name = "lblExtraRecLbl"
        Me.lblExtraRecLbl.Size = New System.Drawing.Size(74, 13)
        Me.lblExtraRecLbl.TabIndex = 4
        Me.lblExtraRecLbl.Text = "Extra Rec Lbl:"
        '
        'cmdRecLbl
        '
        Me.cmdRecLbl.Location = New System.Drawing.Point(115, 200)
        Me.cmdRecLbl.Name = "cmdRecLbl"
        Me.cmdRecLbl.Size = New System.Drawing.Size(75, 23)
        Me.cmdRecLbl.TabIndex = 5
        Me.cmdRecLbl.Text = "&Rec Label"
        Me.ToolTip1.SetToolTip(Me.cmdRecLbl, "Print an extra receiving label.")
        Me.cmdRecLbl.UseVisualStyleBackColor = True
        '
        'lblOrientation
        '
        Me.lblOrientation.AutoSize = True
        Me.lblOrientation.Location = New System.Drawing.Point(193, 184)
        Me.lblOrientation.Name = "lblOrientation"
        Me.lblOrientation.Size = New System.Drawing.Size(83, 13)
        Me.lblOrientation.TabIndex = 6
        Me.lblOrientation.Text = "Tag Orientation:"
        '
        'cboTagJustify
        '
        Me.cboTagJustify.FormattingEnabled = True
        Me.cboTagJustify.Items.AddRange(New Object() {"Center", "Left", "Right"})
        Me.cboTagJustify.Location = New System.Drawing.Point(196, 200)
        Me.cboTagJustify.Name = "cboTagJustify"
        Me.cboTagJustify.Size = New System.Drawing.Size(71, 21)
        Me.cboTagJustify.TabIndex = 7
        '
        'lblQty
        '
        Me.lblQty.AutoSize = True
        Me.lblQty.Location = New System.Drawing.Point(279, 184)
        Me.lblQty.Name = "lblQty"
        Me.lblQty.Size = New System.Drawing.Size(26, 13)
        Me.lblQty.TabIndex = 8
        Me.lblQty.Text = "Qty:"
        '
        'txtCopies
        '
        Me.txtCopies.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.txtCopies.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtCopies.Location = New System.Drawing.Point(282, 200)
        Me.txtCopies.Name = "txtCopies"
        Me.txtCopies.Size = New System.Drawing.Size(33, 20)
        Me.txtCopies.TabIndex = 9
        Me.txtCopies.Text = "1"
        '
        'updQuantity
        '
        Me.updQuantity.Location = New System.Drawing.Point(318, 199)
        Me.updQuantity.Name = "updQuantity"
        Me.updQuantity.OcxState = CType(resources.GetObject("updQuantity.OcxState"), System.Windows.Forms.AxHost.State)
        Me.updQuantity.Size = New System.Drawing.Size(17, 21)
        Me.updQuantity.TabIndex = 10
        '
        'cmdCustom
        '
        Me.cmdCustom.Location = New System.Drawing.Point(5, 227)
        Me.cmdCustom.Name = "cmdCustom"
        Me.cmdCustom.Size = New System.Drawing.Size(51, 28)
        Me.cmdCustom.TabIndex = 11
        Me.cmdCustom.Text = "C&ustom"
        Me.ToolTip1.SetToolTip(Me.cmdCustom, " Click On to print a Price Tag ")
        Me.cmdCustom.UseVisualStyleBackColor = True
        '
        'cmdLarge
        '
        Me.cmdLarge.Location = New System.Drawing.Point(58, 227)
        Me.cmdLarge.Name = "cmdLarge"
        Me.cmdLarge.Size = New System.Drawing.Size(51, 28)
        Me.cmdLarge.TabIndex = 12
        Me.cmdLarge.Text = "&Large"
        Me.ToolTip1.SetToolTip(Me.cmdLarge, " Click On to print a Price Tag ")
        Me.cmdLarge.UseVisualStyleBackColor = True
        '
        'cmdMedium
        '
        Me.cmdMedium.Location = New System.Drawing.Point(111, 227)
        Me.cmdMedium.Name = "cmdMedium"
        Me.cmdMedium.Size = New System.Drawing.Size(68, 28)
        Me.cmdMedium.TabIndex = 13
        Me.cmdMedium.Text = "&Medium"
        Me.cmdMedium.UseVisualStyleBackColor = True
        '
        'cmdSmall
        '
        Me.cmdSmall.Location = New System.Drawing.Point(181, 227)
        Me.cmdSmall.Name = "cmdSmall"
        Me.cmdSmall.Size = New System.Drawing.Size(51, 28)
        Me.cmdSmall.TabIndex = 14
        Me.cmdSmall.Text = "&Small"
        Me.ToolTip1.SetToolTip(Me.cmdSmall, " Avery #5260, Office Max 86101 ")
        Me.cmdSmall.UseVisualStyleBackColor = True
        '
        'cmdDYMO
        '
        Me.cmdDYMO.Location = New System.Drawing.Point(234, 227)
        Me.cmdDYMO.Name = "cmdDYMO"
        Me.cmdDYMO.Size = New System.Drawing.Size(51, 28)
        Me.cmdDYMO.TabIndex = 15
        Me.cmdDYMO.Text = "D&YMO"
        Me.ToolTip1.SetToolTip(Me.cmdDYMO, "Print using DYMO 30323 Shipping labels.")
        Me.cmdDYMO.UseVisualStyleBackColor = True
        '
        'cmdDone
        '
        Me.cmdDone.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cmdDone.Location = New System.Drawing.Point(287, 227)
        Me.cmdDone.Name = "cmdDone"
        Me.cmdDone.Size = New System.Drawing.Size(51, 28)
        Me.cmdDone.TabIndex = 16
        Me.cmdDone.Text = "&Done"
        Me.cmdDone.UseVisualStyleBackColor = True
        '
        'pic
        '
        Me.pic.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.pic.Location = New System.Drawing.Point(-353, 68)
        Me.pic.Name = "pic"
        Me.pic.Size = New System.Drawing.Size(43, 50)
        Me.pic.TabIndex = 17
        Me.pic.TabStop = False
        Me.pic.Visible = False
        '
        'PrSel
        '
        Me.PrSel.AllowDYMO = True
        Me.PrSel.AutoSelect = True
        Me.PrSel.Location = New System.Drawing.Point(1, 17)
        Me.PrSel.Name = "PrSel"
        Me.PrSel.Size = New System.Drawing.Size(309, 108)
        Me.PrSel.TabIndex = 0
        '
        'SelectPrinter
        '
        Me.ClientSize = New System.Drawing.Size(342, 258)
        Me.Controls.Add(Me.pic)
        Me.Controls.Add(Me.cmdDone)
        Me.Controls.Add(Me.cmdDYMO)
        Me.Controls.Add(Me.cmdSmall)
        Me.Controls.Add(Me.cmdMedium)
        Me.Controls.Add(Me.cmdLarge)
        Me.Controls.Add(Me.cmdCustom)
        Me.Controls.Add(Me.updQuantity)
        Me.Controls.Add(Me.txtCopies)
        Me.Controls.Add(Me.lblQty)
        Me.Controls.Add(Me.cboTagJustify)
        Me.Controls.Add(Me.lblOrientation)
        Me.Controls.Add(Me.cmdRecLbl)
        Me.Controls.Add(Me.lblExtraRecLbl)
        Me.Controls.Add(Me.cboCustomTagTemplate)
        Me.Controls.Add(Me.lblCustomCaption)
        Me.Controls.Add(Me.fra)
        Me.Controls.Add(Me.lblDisplayStyle)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Name = "SelectPrinter"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Select Printer And Size"
        Me.fra.ResumeLayout(False)
        CType(Me.updQuantity, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.pic, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents lblDisplayStyle As Label
    Friend WithEvents fra As GroupBox
    Friend WithEvents PrSel As PrinterSelector
    Friend WithEvents lblCustomCaption As Label
    Friend WithEvents cboCustomTagTemplate As ComboBox
    Friend WithEvents lblExtraRecLbl As Label
    Friend WithEvents cmdRecLbl As Button
    Friend WithEvents lblOrientation As Label
    Friend WithEvents cboTagJustify As ComboBox
    Friend WithEvents lblQty As Label
    Friend WithEvents txtCopies As TextBox
    Friend WithEvents updQuantity As AxMSComCtl2.AxUpDown
    Friend WithEvents cmdCustom As Button
    Friend WithEvents cmdLarge As Button
    Friend WithEvents cmdMedium As Button
    Friend WithEvents cmdSmall As Button
    Friend WithEvents cmdDYMO As Button
    Friend WithEvents cmdDone As Button
    Friend WithEvents ToolTip1 As ToolTip
    Friend WithEvents pic As PictureBox
End Class
