<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class frmDeliveryMap
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmDeliveryMap))
        Me.mapDelivery = New AxMapPoint.AxMappointControl()
        Me.cmdSplit = New System.Windows.Forms.Button()
        Me.lvwAllStops = New System.Windows.Forms.ListView()
        Me.lvwThisTruck = New System.Windows.Forms.ListView()
        Me.imlStops = New System.Windows.Forms.ImageList(Me.components)
        Me.cmdAddAll = New System.Windows.Forms.Button()
        Me.cmdDone = New System.Windows.Forms.Button()
        Me.cmdPrint = New System.Windows.Forms.Button()
        Me.cmdRemoveAll = New System.Windows.Forms.Button()
        Me.cmdShow = New System.Windows.Forms.Button()
        Me.cmdAdjust = New System.Windows.Forms.Button()
        Me.cmdManifest = New System.Windows.Forms.Button()
        Me.cmbPrintType = New System.Windows.Forms.ComboBox()
        Me.cmdConfigure = New System.Windows.Forms.Button()
        Me.cmdCancel = New System.Windows.Forms.Button()
        Me.cmdDetails = New System.Windows.Forms.Button()
        Me.lblAllStopsCubes = New System.Windows.Forms.Label()
        Me.lblCurrentTruckCubes = New System.Windows.Forms.Label()
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
        Me.cmdSplit.Location = New System.Drawing.Point(645, 312)
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
        Me.lvwAllStops.SmallImageList = Me.imlStops
        Me.lvwAllStops.TabIndex = 2
        Me.lvwAllStops.UseCompatibleStateImageBehavior = False
        Me.lvwAllStops.View = System.Windows.Forms.View.SmallIcon
        '
        'lvwThisTruck
        '
        Me.lvwThisTruck.HideSelection = False
        Me.lvwThisTruck.Location = New System.Drawing.Point(398, 62)
        Me.lvwThisTruck.Name = "lvwThisTruck"
        Me.lvwThisTruck.Size = New System.Drawing.Size(121, 97)
        Me.lvwThisTruck.SmallImageList = Me.imlStops
        Me.lvwThisTruck.TabIndex = 3
        Me.lvwThisTruck.UseCompatibleStateImageBehavior = False
        '
        'imlStops
        '
        Me.imlStops.ImageStream = CType(resources.GetObject("imlStops.ImageStream"), System.Windows.Forms.ImageListStreamer)
        Me.imlStops.TransparentColor = System.Drawing.Color.Transparent
        Me.imlStops.Images.SetKeyName(0, "Find 2.ico")
        Me.imlStops.Images.SetKeyName(1, "Home 2.ico")
        '
        'cmdAddAll
        '
        Me.cmdAddAll.Location = New System.Drawing.Point(645, 167)
        Me.cmdAddAll.Name = "cmdAddAll"
        Me.cmdAddAll.Size = New System.Drawing.Size(75, 23)
        Me.cmdAddAll.TabIndex = 4
        Me.cmdAddAll.Text = ">>>"
        Me.cmdAddAll.UseVisualStyleBackColor = True
        '
        'cmdDone
        '
        Me.cmdDone.Location = New System.Drawing.Point(645, 196)
        Me.cmdDone.Name = "cmdDone"
        Me.cmdDone.Size = New System.Drawing.Size(75, 23)
        Me.cmdDone.TabIndex = 5
        Me.cmdDone.Text = "&Done"
        Me.cmdDone.UseVisualStyleBackColor = True
        '
        'cmdPrint
        '
        Me.cmdPrint.Location = New System.Drawing.Point(645, 225)
        Me.cmdPrint.Name = "cmdPrint"
        Me.cmdPrint.Size = New System.Drawing.Size(75, 23)
        Me.cmdPrint.TabIndex = 6
        Me.cmdPrint.Text = "&Print"
        Me.cmdPrint.UseVisualStyleBackColor = True
        '
        'cmdRemoveAll
        '
        Me.cmdRemoveAll.Location = New System.Drawing.Point(645, 254)
        Me.cmdRemoveAll.Name = "cmdRemoveAll"
        Me.cmdRemoveAll.Size = New System.Drawing.Size(75, 23)
        Me.cmdRemoveAll.TabIndex = 7
        Me.cmdRemoveAll.Text = "<<<"
        Me.cmdRemoveAll.UseVisualStyleBackColor = True
        '
        'cmdShow
        '
        Me.cmdShow.Location = New System.Drawing.Point(645, 283)
        Me.cmdShow.Name = "cmdShow"
        Me.cmdShow.Size = New System.Drawing.Size(75, 23)
        Me.cmdShow.TabIndex = 8
        Me.cmdShow.Text = "Locate Sto&ps on Map"
        Me.cmdShow.UseVisualStyleBackColor = True
        '
        'cmdAdjust
        '
        Me.cmdAdjust.Location = New System.Drawing.Point(645, 350)
        Me.cmdAdjust.Name = "cmdAdjust"
        Me.cmdAdjust.Size = New System.Drawing.Size(75, 23)
        Me.cmdAdjust.TabIndex = 9
        Me.cmdAdjust.Text = "Ad&just Route"
        Me.cmdAdjust.UseVisualStyleBackColor = True
        '
        'cmdManifest
        '
        Me.cmdManifest.Location = New System.Drawing.Point(645, 389)
        Me.cmdManifest.Name = "cmdManifest"
        Me.cmdManifest.Size = New System.Drawing.Size(75, 23)
        Me.cmdManifest.TabIndex = 10
        Me.cmdManifest.Text = "P&rint Manifest"
        Me.cmdManifest.UseVisualStyleBackColor = True
        '
        'cmbPrintType
        '
        Me.cmbPrintType.FormattingEnabled = True
        Me.cmbPrintType.Items.AddRange(New Object() {"Strips", "Full", "Turns", "Dirs", "Map"})
        Me.cmbPrintType.Location = New System.Drawing.Point(645, 418)
        Me.cmbPrintType.Name = "cmbPrintType"
        Me.cmbPrintType.Size = New System.Drawing.Size(121, 21)
        Me.cmbPrintType.TabIndex = 11
        Me.cmbPrintType.Text = "cmbPrintType"
        '
        'cmdConfigure
        '
        Me.cmdConfigure.Location = New System.Drawing.Point(726, 167)
        Me.cmdConfigure.Name = "cmdConfigure"
        Me.cmdConfigure.Size = New System.Drawing.Size(75, 41)
        Me.cmdConfigure.TabIndex = 12
        Me.cmdConfigure.Text = "Con&figure Optimization"
        Me.cmdConfigure.UseVisualStyleBackColor = True
        '
        'cmdCancel
        '
        Me.cmdCancel.Location = New System.Drawing.Point(726, 214)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.Size = New System.Drawing.Size(75, 23)
        Me.cmdCancel.TabIndex = 13
        Me.cmdCancel.Text = "&Cancel Process"
        Me.cmdCancel.UseVisualStyleBackColor = True
        '
        'cmdDetails
        '
        Me.cmdDetails.Location = New System.Drawing.Point(726, 254)
        Me.cmdDetails.Name = "cmdDetails"
        Me.cmdDetails.Size = New System.Drawing.Size(75, 23)
        Me.cmdDetails.TabIndex = 14
        Me.cmdDetails.Text = "Deta&ils"
        Me.cmdDetails.UseVisualStyleBackColor = True
        '
        'lblAllStopsCubes
        '
        Me.lblAllStopsCubes.AutoSize = True
        Me.lblAllStopsCubes.Location = New System.Drawing.Point(362, 409)
        Me.lblAllStopsCubes.Name = "lblAllStopsCubes"
        Me.lblAllStopsCubes.Size = New System.Drawing.Size(91, 13)
        Me.lblAllStopsCubes.TabIndex = 15
        Me.lblAllStopsCubes.Text = "Total Cubes: 0.00"
        '
        'lblCurrentTruckCubes
        '
        Me.lblCurrentTruckCubes.AutoSize = True
        Me.lblCurrentTruckCubes.Location = New System.Drawing.Point(369, 436)
        Me.lblCurrentTruckCubes.Name = "lblCurrentTruckCubes"
        Me.lblCurrentTruckCubes.Size = New System.Drawing.Size(91, 13)
        Me.lblCurrentTruckCubes.TabIndex = 16
        Me.lblCurrentTruckCubes.Text = "Total Cubes: 0.00"
        '
        'frmDeliveryMap
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(800, 450)
        Me.Controls.Add(Me.lblCurrentTruckCubes)
        Me.Controls.Add(Me.lblAllStopsCubes)
        Me.Controls.Add(Me.cmdDetails)
        Me.Controls.Add(Me.cmdCancel)
        Me.Controls.Add(Me.cmdConfigure)
        Me.Controls.Add(Me.cmbPrintType)
        Me.Controls.Add(Me.cmdManifest)
        Me.Controls.Add(Me.cmdAdjust)
        Me.Controls.Add(Me.cmdShow)
        Me.Controls.Add(Me.cmdRemoveAll)
        Me.Controls.Add(Me.cmdPrint)
        Me.Controls.Add(Me.cmdDone)
        Me.Controls.Add(Me.cmdAddAll)
        Me.Controls.Add(Me.lvwThisTruck)
        Me.Controls.Add(Me.lvwAllStops)
        Me.Controls.Add(Me.cmdSplit)
        Me.Controls.Add(Me.mapDelivery)
        Me.Name = "frmDeliveryMap"
        Me.Text = "frmDeliveryMap"
        CType(Me.mapDelivery, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents mapDelivery As AxMapPoint.AxMappointControl
    Friend WithEvents cmdSplit As Button
    Friend WithEvents lvwAllStops As ListView
    Friend WithEvents lvwThisTruck As ListView
    Friend WithEvents imlStops As ImageList
    Friend WithEvents cmdAddAll As Button
    Friend WithEvents cmdDone As Button
    Friend WithEvents cmdPrint As Button
    Friend WithEvents cmdRemoveAll As Button
    Friend WithEvents cmdShow As Button
    Friend WithEvents cmdAdjust As Button
    Friend WithEvents cmdManifest As Button
    Friend WithEvents cmbPrintType As ComboBox
    Friend WithEvents cmdConfigure As Button
    Friend WithEvents cmdCancel As Button
    Friend WithEvents cmdDetails As Button
    Friend WithEvents lblAllStopsCubes As Label
    Friend WithEvents lblCurrentTruckCubes As Label
End Class
