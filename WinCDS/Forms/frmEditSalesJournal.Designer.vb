<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmEditSalesJournal
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmEditSalesJournal))
        Me.cmdContinueAudit = New System.Windows.Forms.Button()
        Me.cmdCancelAudit = New System.Windows.Forms.Button()
        Me.grdSalesJournal = New AxMSFlexGridLib.AxMSFlexGrid()
        Me.fraEditControls = New System.Windows.Forms.GroupBox()
        Me.lblIndex = New System.Windows.Forms.Label()
        Me.dteTransDate = New System.Windows.Forms.DateTimePicker()
        Me.cmdApply = New System.Windows.Forms.Button()
        Me.cmdCancel = New System.Windows.Forms.Button()
        Me.cmdRefresh = New System.Windows.Forms.Button()
        Me.lblTransDate = New System.Windows.Forms.Label()
        CType(Me.grdSalesJournal, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.fraEditControls.SuspendLayout()
        Me.SuspendLayout()
        '
        'cmdContinueAudit
        '
        Me.cmdContinueAudit.Location = New System.Drawing.Point(281, 291)
        Me.cmdContinueAudit.Name = "cmdContinueAudit"
        Me.cmdContinueAudit.Size = New System.Drawing.Size(86, 60)
        Me.cmdContinueAudit.TabIndex = 0
        Me.cmdContinueAudit.Text = "C&ontinue Audit Report"
        Me.cmdContinueAudit.UseVisualStyleBackColor = True
        '
        'cmdCancelAudit
        '
        Me.cmdCancelAudit.Location = New System.Drawing.Point(373, 291)
        Me.cmdCancelAudit.Name = "cmdCancelAudit"
        Me.cmdCancelAudit.Size = New System.Drawing.Size(98, 60)
        Me.cmdCancelAudit.TabIndex = 1
        Me.cmdCancelAudit.Text = "Cancel Audit Re&port"
        Me.cmdCancelAudit.UseVisualStyleBackColor = True
        '
        'grdSalesJournal
        '
        Me.grdSalesJournal.Location = New System.Drawing.Point(12, 12)
        Me.grdSalesJournal.Name = "grdSalesJournal"
        Me.grdSalesJournal.OcxState = CType(resources.GetObject("grdSalesJournal.OcxState"), System.Windows.Forms.AxHost.State)
        Me.grdSalesJournal.Size = New System.Drawing.Size(481, 129)
        Me.grdSalesJournal.TabIndex = 2
        '
        'fraEditControls
        '
        Me.fraEditControls.Controls.Add(Me.lblTransDate)
        Me.fraEditControls.Controls.Add(Me.lblIndex)
        Me.fraEditControls.Controls.Add(Me.dteTransDate)
        Me.fraEditControls.Location = New System.Drawing.Point(96, 156)
        Me.fraEditControls.Name = "fraEditControls"
        Me.fraEditControls.Size = New System.Drawing.Size(316, 129)
        Me.fraEditControls.TabIndex = 3
        Me.fraEditControls.TabStop = False
        '
        'lblIndex
        '
        Me.lblIndex.Location = New System.Drawing.Point(39, 17)
        Me.lblIndex.Name = "lblIndex"
        Me.lblIndex.Size = New System.Drawing.Size(249, 57)
        Me.lblIndex.TabIndex = 0
        Me.lblIndex.Text = "Label1"
        '
        'dteTransDate
        '
        Me.dteTransDate.CustomFormat = "MM/dd/yyyy"
        Me.dteTransDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dteTransDate.Location = New System.Drawing.Point(118, 89)
        Me.dteTransDate.Name = "dteTransDate"
        Me.dteTransDate.Size = New System.Drawing.Size(92, 20)
        Me.dteTransDate.TabIndex = 4
        '
        'cmdApply
        '
        Me.cmdApply.Location = New System.Drawing.Point(119, 291)
        Me.cmdApply.Name = "cmdApply"
        Me.cmdApply.Size = New System.Drawing.Size(75, 60)
        Me.cmdApply.TabIndex = 5
        Me.cmdApply.Text = "&Apply"
        Me.cmdApply.UseVisualStyleBackColor = True
        '
        'cmdCancel
        '
        Me.cmdCancel.Location = New System.Drawing.Point(200, 291)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.Size = New System.Drawing.Size(75, 60)
        Me.cmdCancel.TabIndex = 6
        Me.cmdCancel.Text = "&Cancel"
        Me.cmdCancel.UseVisualStyleBackColor = True
        '
        'cmdRefresh
        '
        Me.cmdRefresh.Location = New System.Drawing.Point(418, 168)
        Me.cmdRefresh.Name = "cmdRefresh"
        Me.cmdRefresh.Size = New System.Drawing.Size(75, 62)
        Me.cmdRefresh.TabIndex = 7
        Me.cmdRefresh.Text = "&Refresh"
        Me.cmdRefresh.UseVisualStyleBackColor = True
        '
        'lblTransDate
        '
        Me.lblTransDate.AutoSize = True
        Me.lblTransDate.Location = New System.Drawing.Point(49, 96)
        Me.lblTransDate.Name = "lblTransDate"
        Me.lblTransDate.Size = New System.Drawing.Size(63, 13)
        Me.lblTransDate.TabIndex = 1
        Me.lblTransDate.Text = "Trans Date:"
        '
        'frmEditSalesJournal
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(502, 359)
        Me.Controls.Add(Me.cmdRefresh)
        Me.Controls.Add(Me.cmdCancel)
        Me.Controls.Add(Me.cmdApply)
        Me.Controls.Add(Me.fraEditControls)
        Me.Controls.Add(Me.grdSalesJournal)
        Me.Controls.Add(Me.cmdCancelAudit)
        Me.Controls.Add(Me.cmdContinueAudit)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.MaximizeBox = False
        Me.Name = "frmEditSalesJournal"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Sales Journal Editor"
        CType(Me.grdSalesJournal, System.ComponentModel.ISupportInitialize).EndInit()
        Me.fraEditControls.ResumeLayout(False)
        Me.fraEditControls.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents cmdContinueAudit As Button
    Friend WithEvents cmdCancelAudit As Button
    Friend WithEvents grdSalesJournal As AxMSFlexGridLib.AxMSFlexGrid
    Friend WithEvents fraEditControls As GroupBox
    Friend WithEvents lblIndex As Label
    Friend WithEvents dteTransDate As DateTimePicker
    Friend WithEvents cmdApply As Button
    Friend WithEvents cmdCancel As Button
    Friend WithEvents cmdRefresh As Button
    Friend WithEvents lblTransDate As Label
End Class
