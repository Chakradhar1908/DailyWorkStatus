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
        Me.cmdSplit = New System.Windows.Forms.Button()
        Me.lvwAllStops = New System.Windows.Forms.ListView()
        Me.imlStops = New System.Windows.Forms.ImageList(Me.components)
        Me.lvwThisTruck = New System.Windows.Forms.ListView()
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
        Me.lblInstructions = New System.Windows.Forms.Label()
        Me.fraSplitLoads = New System.Windows.Forms.GroupBox()
        Me.mapStops = New AxMapPoint.AxMappointControl()
        Me.lblmapStops = New System.Windows.Forms.Label()
        Me.lblThisTruck = New System.Windows.Forms.Label()
        Me.lblAllStops = New System.Windows.Forms.Label()
        Me.fraMapContainer = New System.Windows.Forms.GroupBox()
        Me.mapDelivery = New AxMapPoint.AxMappointControl()
        Me.fraSplitLoads.SuspendLayout()
        CType(Me.mapStops, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.fraMapContainer.SuspendLayout()
        CType(Me.mapDelivery, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'cmdSplit
        '
        Me.cmdSplit.Location = New System.Drawing.Point(601, 143)
        Me.cmdSplit.Name = "cmdSplit"
        Me.cmdSplit.Size = New System.Drawing.Size(75, 33)
        Me.cmdSplit.TabIndex = 1
        Me.cmdSplit.Text = "&Split Loads"
        Me.cmdSplit.UseVisualStyleBackColor = True
        '
        'lvwAllStops
        '
        Me.lvwAllStops.HideSelection = False
        Me.lvwAllStops.Location = New System.Drawing.Point(8, 76)
        Me.lvwAllStops.Name = "lvwAllStops"
        Me.lvwAllStops.Size = New System.Drawing.Size(233, 163)
        Me.lvwAllStops.SmallImageList = Me.imlStops
        Me.lvwAllStops.TabIndex = 2
        Me.lvwAllStops.UseCompatibleStateImageBehavior = False
        Me.lvwAllStops.View = System.Windows.Forms.View.SmallIcon
        '
        'imlStops
        '
        Me.imlStops.ImageStream = CType(resources.GetObject("imlStops.ImageStream"), System.Windows.Forms.ImageListStreamer)
        Me.imlStops.TransparentColor = System.Drawing.Color.Transparent
        Me.imlStops.Images.SetKeyName(0, "Find 2.ico")
        Me.imlStops.Images.SetKeyName(1, "Home 2.ico")
        '
        'lvwThisTruck
        '
        Me.lvwThisTruck.HideSelection = False
        Me.lvwThisTruck.Location = New System.Drawing.Point(324, 80)
        Me.lvwThisTruck.Name = "lvwThisTruck"
        Me.lvwThisTruck.Size = New System.Drawing.Size(237, 159)
        Me.lvwThisTruck.SmallImageList = Me.imlStops
        Me.lvwThisTruck.TabIndex = 3
        Me.lvwThisTruck.UseCompatibleStateImageBehavior = False
        '
        'cmdAddAll
        '
        Me.cmdAddAll.Location = New System.Drawing.Point(257, 157)
        Me.cmdAddAll.Name = "cmdAddAll"
        Me.cmdAddAll.Size = New System.Drawing.Size(57, 23)
        Me.cmdAddAll.TabIndex = 4
        Me.cmdAddAll.Text = ">>>"
        Me.cmdAddAll.UseVisualStyleBackColor = True
        '
        'cmdDone
        '
        Me.cmdDone.Location = New System.Drawing.Point(601, 263)
        Me.cmdDone.Name = "cmdDone"
        Me.cmdDone.Size = New System.Drawing.Size(75, 39)
        Me.cmdDone.TabIndex = 5
        Me.cmdDone.Text = "&Done"
        Me.cmdDone.UseVisualStyleBackColor = True
        '
        'cmdPrint
        '
        Me.cmdPrint.Location = New System.Drawing.Point(601, 220)
        Me.cmdPrint.Name = "cmdPrint"
        Me.cmdPrint.Size = New System.Drawing.Size(75, 37)
        Me.cmdPrint.TabIndex = 6
        Me.cmdPrint.Text = "&Print"
        Me.cmdPrint.UseVisualStyleBackColor = True
        '
        'cmdRemoveAll
        '
        Me.cmdRemoveAll.Location = New System.Drawing.Point(257, 124)
        Me.cmdRemoveAll.Name = "cmdRemoveAll"
        Me.cmdRemoveAll.Size = New System.Drawing.Size(57, 27)
        Me.cmdRemoveAll.TabIndex = 7
        Me.cmdRemoveAll.Text = "<<<"
        Me.cmdRemoveAll.UseVisualStyleBackColor = True
        '
        'cmdShow
        '
        Me.cmdShow.Location = New System.Drawing.Point(257, 187)
        Me.cmdShow.Name = "cmdShow"
        Me.cmdShow.Size = New System.Drawing.Size(57, 52)
        Me.cmdShow.TabIndex = 8
        Me.cmdShow.Text = "Locate Sto&ps on Map"
        Me.cmdShow.UseVisualStyleBackColor = True
        '
        'cmdAdjust
        '
        Me.cmdAdjust.Location = New System.Drawing.Point(598, 54)
        Me.cmdAdjust.Name = "cmdAdjust"
        Me.cmdAdjust.Size = New System.Drawing.Size(78, 36)
        Me.cmdAdjust.TabIndex = 9
        Me.cmdAdjust.Text = "Ad&just Route"
        Me.cmdAdjust.UseVisualStyleBackColor = True
        '
        'cmdManifest
        '
        Me.cmdManifest.Location = New System.Drawing.Point(258, 73)
        Me.cmdManifest.Name = "cmdManifest"
        Me.cmdManifest.Size = New System.Drawing.Size(60, 34)
        Me.cmdManifest.TabIndex = 10
        Me.cmdManifest.Text = "P&rint Manifest"
        Me.cmdManifest.UseVisualStyleBackColor = True
        '
        'cmbPrintType
        '
        Me.cmbPrintType.FormattingEnabled = True
        Me.cmbPrintType.Items.AddRange(New Object() {"Strips", "Full", "Turns", "Dirs", "Map"})
        Me.cmbPrintType.Location = New System.Drawing.Point(601, 187)
        Me.cmbPrintType.Name = "cmbPrintType"
        Me.cmbPrintType.Size = New System.Drawing.Size(75, 21)
        Me.cmbPrintType.TabIndex = 11
        Me.cmbPrintType.Text = "cmbPrintType"
        '
        'cmdConfigure
        '
        Me.cmdConfigure.Location = New System.Drawing.Point(601, 98)
        Me.cmdConfigure.Name = "cmdConfigure"
        Me.cmdConfigure.Size = New System.Drawing.Size(75, 41)
        Me.cmdConfigure.TabIndex = 12
        Me.cmdConfigure.Text = "Con&figure Optimization"
        Me.cmdConfigure.UseVisualStyleBackColor = True
        '
        'cmdCancel
        '
        Me.cmdCancel.Location = New System.Drawing.Point(598, 98)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.Size = New System.Drawing.Size(75, 39)
        Me.cmdCancel.TabIndex = 13
        Me.cmdCancel.Text = "&Cancel Process"
        Me.cmdCancel.UseVisualStyleBackColor = True
        '
        'cmdDetails
        '
        Me.cmdDetails.Location = New System.Drawing.Point(257, 33)
        Me.cmdDetails.Name = "cmdDetails"
        Me.cmdDetails.Size = New System.Drawing.Size(61, 25)
        Me.cmdDetails.TabIndex = 14
        Me.cmdDetails.Text = "Deta&ils"
        Me.cmdDetails.UseVisualStyleBackColor = True
        '
        'lblAllStopsCubes
        '
        Me.lblAllStopsCubes.AutoSize = True
        Me.lblAllStopsCubes.Location = New System.Drawing.Point(87, 58)
        Me.lblAllStopsCubes.Name = "lblAllStopsCubes"
        Me.lblAllStopsCubes.Size = New System.Drawing.Size(91, 13)
        Me.lblAllStopsCubes.TabIndex = 15
        Me.lblAllStopsCubes.Text = "Total Cubes: 0.00"
        '
        'lblCurrentTruckCubes
        '
        Me.lblCurrentTruckCubes.AutoSize = True
        Me.lblCurrentTruckCubes.Location = New System.Drawing.Point(394, 65)
        Me.lblCurrentTruckCubes.Name = "lblCurrentTruckCubes"
        Me.lblCurrentTruckCubes.Size = New System.Drawing.Size(91, 13)
        Me.lblCurrentTruckCubes.TabIndex = 16
        Me.lblCurrentTruckCubes.Text = "Total Cubes: 0.00"
        '
        'lblInstructions
        '
        Me.lblInstructions.AutoSize = True
        Me.lblInstructions.Location = New System.Drawing.Point(6, 16)
        Me.lblInstructions.Name = "lblInstructions"
        Me.lblInstructions.Size = New System.Drawing.Size(326, 13)
        Me.lblInstructions.TabIndex = 17
        Me.lblInstructions.Text = "To move an individual stop, simply drag it from one side to the other."
        '
        'fraSplitLoads
        '
        Me.fraSplitLoads.Controls.Add(Me.mapStops)
        Me.fraSplitLoads.Controls.Add(Me.lblmapStops)
        Me.fraSplitLoads.Controls.Add(Me.lblThisTruck)
        Me.fraSplitLoads.Controls.Add(Me.lblCurrentTruckCubes)
        Me.fraSplitLoads.Controls.Add(Me.lblAllStops)
        Me.fraSplitLoads.Controls.Add(Me.lblInstructions)
        Me.fraSplitLoads.Controls.Add(Me.cmdDetails)
        Me.fraSplitLoads.Controls.Add(Me.lblAllStopsCubes)
        Me.fraSplitLoads.Controls.Add(Me.lvwAllStops)
        Me.fraSplitLoads.Controls.Add(Me.lvwThisTruck)
        Me.fraSplitLoads.Controls.Add(Me.cmdManifest)
        Me.fraSplitLoads.Controls.Add(Me.cmdShow)
        Me.fraSplitLoads.Controls.Add(Me.cmdRemoveAll)
        Me.fraSplitLoads.Controls.Add(Me.cmdAddAll)
        Me.fraSplitLoads.Location = New System.Drawing.Point(12, 5)
        Me.fraSplitLoads.Name = "fraSplitLoads"
        Me.fraSplitLoads.Size = New System.Drawing.Size(580, 367)
        Me.fraSplitLoads.TabIndex = 18
        Me.fraSplitLoads.TabStop = False
        '
        'mapStops
        '
        Me.mapStops.Enabled = True
        Me.mapStops.Location = New System.Drawing.Point(9, 293)
        Me.mapStops.Name = "mapStops"
        Me.mapStops.OcxState = CType(resources.GetObject("mapStops.OcxState"), System.Windows.Forms.AxHost.State)
        Me.mapStops.Size = New System.Drawing.Size(56, 56)
        Me.mapStops.TabIndex = 21
        '
        'lblmapStops
        '
        Me.lblmapStops.Location = New System.Drawing.Point(6, 248)
        Me.lblmapStops.Name = "lblmapStops"
        Me.lblmapStops.Size = New System.Drawing.Size(529, 42)
        Me.lblmapStops.TabIndex = 20
        Me.lblmapStops.Text = resources.GetString("lblmapStops.Text")
        '
        'lblThisTruck
        '
        Me.lblThisTruck.AutoSize = True
        Me.lblThisTruck.Font = New System.Drawing.Font("Arial", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblThisTruck.ForeColor = System.Drawing.SystemColors.Highlight
        Me.lblThisTruck.Location = New System.Drawing.Point(376, 43)
        Me.lblThisTruck.Name = "lblThisTruck"
        Me.lblThisTruck.Size = New System.Drawing.Size(140, 22)
        Me.lblThisTruck.TabIndex = 19
        Me.lblThisTruck.Text = "Current Truck"
        '
        'lblAllStops
        '
        Me.lblAllStops.AutoSize = True
        Me.lblAllStops.Font = New System.Drawing.Font("Arial", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblAllStops.ForeColor = System.Drawing.SystemColors.Highlight
        Me.lblAllStops.Location = New System.Drawing.Point(86, 36)
        Me.lblAllStops.Name = "lblAllStops"
        Me.lblAllStops.Size = New System.Drawing.Size(92, 22)
        Me.lblAllStops.TabIndex = 18
        Me.lblAllStops.Text = "All Stops"
        '
        'fraMapContainer
        '
        Me.fraMapContainer.Controls.Add(Me.mapDelivery)
        Me.fraMapContainer.Location = New System.Drawing.Point(682, 8)
        Me.fraMapContainer.Name = "fraMapContainer"
        Me.fraMapContainer.Size = New System.Drawing.Size(68, 83)
        Me.fraMapContainer.TabIndex = 19
        Me.fraMapContainer.TabStop = False
        '
        'mapDelivery
        '
        Me.mapDelivery.Enabled = True
        Me.mapDelivery.Location = New System.Drawing.Point(6, 19)
        Me.mapDelivery.Name = "mapDelivery"
        Me.mapDelivery.OcxState = CType(resources.GetObject("mapDelivery.OcxState"), System.Windows.Forms.AxHost.State)
        Me.mapDelivery.Size = New System.Drawing.Size(53, 49)
        Me.mapDelivery.TabIndex = 0
        '
        'frmDeliveryMap
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(753, 390)
        Me.Controls.Add(Me.fraMapContainer)
        Me.Controls.Add(Me.fraSplitLoads)
        Me.Controls.Add(Me.cmdCancel)
        Me.Controls.Add(Me.cmdConfigure)
        Me.Controls.Add(Me.cmbPrintType)
        Me.Controls.Add(Me.cmdAdjust)
        Me.Controls.Add(Me.cmdPrint)
        Me.Controls.Add(Me.cmdDone)
        Me.Controls.Add(Me.cmdSplit)
        Me.Name = "frmDeliveryMap"
        Me.Text = "Delivery Map"
        Me.fraSplitLoads.ResumeLayout(False)
        Me.fraSplitLoads.PerformLayout()
        CType(Me.mapStops, System.ComponentModel.ISupportInitialize).EndInit()
        Me.fraMapContainer.ResumeLayout(False)
        CType(Me.mapDelivery, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
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
    Friend WithEvents lblInstructions As Label
    Friend WithEvents fraSplitLoads As GroupBox
    Friend WithEvents lblThisTruck As Label
    Friend WithEvents lblAllStops As Label
    Friend WithEvents mapStops As AxMapPoint.AxMappointControl
    Friend WithEvents lblmapStops As Label
    Friend WithEvents fraMapContainer As GroupBox
    Friend WithEvents mapDelivery As AxMapPoint.AxMappointControl
End Class
