<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class PackagePrice
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
        Me.lstItems = New System.Windows.Forms.ListBox()
        Me.txtTotLanded = New System.Windows.Forms.TextBox()
        Me.txtTotOnSale = New System.Windows.Forms.TextBox()
        Me.txtTotList = New System.Windows.Forms.TextBox()
        Me.txtGM = New System.Windows.Forms.TextBox()
        Me.txtPackagePrice = New System.Windows.Forms.TextBox()
        Me.txtOrigGM = New System.Windows.Forms.TextBox()
        Me.UGridIO1 = New WinCDS.UGridIO()
        Me.SuspendLayout()
        '
        'lstItems
        '
        Me.lstItems.FormattingEnabled = True
        Me.lstItems.Location = New System.Drawing.Point(0, 0)
        Me.lstItems.Name = "lstItems"
        Me.lstItems.Size = New System.Drawing.Size(120, 95)
        Me.lstItems.TabIndex = 0
        '
        'txtTotLanded
        '
        Me.txtTotLanded.Location = New System.Drawing.Point(20, 121)
        Me.txtTotLanded.Name = "txtTotLanded"
        Me.txtTotLanded.Size = New System.Drawing.Size(100, 20)
        Me.txtTotLanded.TabIndex = 1
        '
        'txtTotOnSale
        '
        Me.txtTotOnSale.Location = New System.Drawing.Point(20, 147)
        Me.txtTotOnSale.Name = "txtTotOnSale"
        Me.txtTotOnSale.Size = New System.Drawing.Size(100, 20)
        Me.txtTotOnSale.TabIndex = 2
        '
        'txtTotList
        '
        Me.txtTotList.Location = New System.Drawing.Point(20, 173)
        Me.txtTotList.Name = "txtTotList"
        Me.txtTotList.Size = New System.Drawing.Size(100, 20)
        Me.txtTotList.TabIndex = 3
        '
        'txtGM
        '
        Me.txtGM.Location = New System.Drawing.Point(20, 208)
        Me.txtGM.Name = "txtGM"
        Me.txtGM.Size = New System.Drawing.Size(100, 20)
        Me.txtGM.TabIndex = 4
        '
        'txtPackagePrice
        '
        Me.txtPackagePrice.Location = New System.Drawing.Point(20, 243)
        Me.txtPackagePrice.Name = "txtPackagePrice"
        Me.txtPackagePrice.Size = New System.Drawing.Size(100, 20)
        Me.txtPackagePrice.TabIndex = 5
        '
        'txtOrigGM
        '
        Me.txtOrigGM.Location = New System.Drawing.Point(20, 285)
        Me.txtOrigGM.Name = "txtOrigGM"
        Me.txtOrigGM.Size = New System.Drawing.Size(100, 20)
        Me.txtOrigGM.TabIndex = 6
        '
        'UGridIO1
        '
        Me.UGridIO1.Activated = False
        Me.UGridIO1.Col = 0
        Me.UGridIO1.firstrow = 0
        Me.UGridIO1.Loading = False
        Me.UGridIO1.Location = New System.Drawing.Point(56, 338)
        Me.UGridIO1.MaxCols = 2
        Me.UGridIO1.MaxRows = 10
        Me.UGridIO1.Name = "UGridIO1"
        Me.UGridIO1.Row = 0
        Me.UGridIO1.Size = New System.Drawing.Size(211, 70)
        Me.UGridIO1.TabIndex = 7
        '
        'PackagePrice
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(800, 450)
        Me.Controls.Add(Me.UGridIO1)
        Me.Controls.Add(Me.txtOrigGM)
        Me.Controls.Add(Me.txtPackagePrice)
        Me.Controls.Add(Me.txtGM)
        Me.Controls.Add(Me.txtTotList)
        Me.Controls.Add(Me.txtTotOnSale)
        Me.Controls.Add(Me.txtTotLanded)
        Me.Controls.Add(Me.lstItems)
        Me.Name = "PackagePrice"
        Me.Text = "PackagePrice"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents lstItems As ListBox
    Friend WithEvents txtTotLanded As TextBox
    Friend WithEvents txtTotOnSale As TextBox
    Friend WithEvents txtTotList As TextBox
    Friend WithEvents txtGM As TextBox
    Friend WithEvents txtPackagePrice As TextBox
    Friend WithEvents txtOrigGM As TextBox
    Friend WithEvents UGridIO1 As UGridIO
End Class
