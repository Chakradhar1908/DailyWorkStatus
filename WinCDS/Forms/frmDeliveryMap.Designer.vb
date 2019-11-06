<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmDeliveryMap
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmDeliveryMap))
        Me.mapDelivery = New AxMapPoint.AxMappointControl()
        Me.cmdSplit = New System.Windows.Forms.Button()
        Me.lvwAllStops = New System.Windows.Forms.ListView()
        Me.lvwThisTruck = New System.Windows.Forms.ListView()
        CType(Me.mapDelivery, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'mapDelivery
        '
        Me.mapDelivery.Enabled = True
        Me.mapDelivery.Location = New System.Drawing.Point(82, 181)
        Me.mapDelivery.Name = "mapDelivery"
        Me.mapDelivery.OcxState = CType(resources.GetObject("mapDelivery.OcxState"), System.Windows.Forms.AxHost.State)
        Me.mapDelivery.Size = New System.Drawing.Size(192, 192)
        Me.mapDelivery.TabIndex = 0
        '
        'cmdSplit
        '
        Me.cmdSplit.Location = New System.Drawing.Point(357, 321)
        Me.cmdSplit.Name = "cmdSplit"
        Me.cmdSplit.Size = New System.Drawing.Size(75, 23)
        Me.cmdSplit.TabIndex = 1
        Me.cmdSplit.Text = "&Split Loads"
        Me.cmdSplit.UseVisualStyleBackColor = True
        '
        'lvwAllStops
        '
        Me.lvwAllStops.HideSelection = False
        Me.lvwAllStops.Location = New System.Drawing.Point(196, 43)
        Me.lvwAllStops.Name = "lvwAllStops"
        Me.lvwAllStops.Size = New System.Drawing.Size(121, 97)
        Me.lvwAllStops.TabIndex = 2
        Me.lvwAllStops.UseCompatibleStateImageBehavior = False
        '
        'lvwThisTruck
        '
        Me.lvwThisTruck.HideSelection = False
        Me.lvwThisTruck.Location = New System.Drawing.Point(398, 62)
        Me.lvwThisTruck.Name = "lvwThisTruck"
        Me.lvwThisTruck.Size = New System.Drawing.Size(121, 97)
        Me.lvwThisTruck.TabIndex = 3
        Me.lvwThisTruck.UseCompatibleStateImageBehavior = False
        '
        'frmDeliveryMap
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(800, 450)
        Me.Controls.Add(Me.lvwThisTruck)
        Me.Controls.Add(Me.lvwAllStops)
        Me.Controls.Add(Me.cmdSplit)
        Me.Controls.Add(Me.mapDelivery)
        Me.Name = "frmDeliveryMap"
        Me.Text = "frmDeliveryMap"
        CType(Me.mapDelivery, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents mapDelivery As AxMapPoint.AxMappointControl
    Friend WithEvents cmdSplit As Button
    Friend WithEvents lvwAllStops As ListView
    Friend WithEvents lvwThisTruck As ListView
End Class
