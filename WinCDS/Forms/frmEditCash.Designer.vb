<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmEditCash
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmEditCash))
        Me.cmdContinueAudit = New System.Windows.Forms.Button()
        Me.cmdCancelAudit = New System.Windows.Forms.Button()
        Me.grdCashJournal = New AxMSFlexGridLib.AxMSFlexGrid()
        Me.fraEditControls = New System.Windows.Forms.GroupBox()
        Me.lblIndex = New System.Windows.Forms.Label()
        Me.dteTransDate = New System.Windows.Forms.DateTimePicker()
        Me.cmdApply = New System.Windows.Forms.Button()
        Me.cmdCancel = New System.Windows.Forms.Button()
        Me.cmdRefresh = New System.Windows.Forms.Button()
        Me.lblTransDate = New System.Windows.Forms.Label()
        CType(Me.grdCashJournal, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.fraEditControls.SuspendLayout()
        Me.SuspendLayout()
        '
        'cmdContinueAudit
        '
        Me.cmdContinueAudit.Location = New System.Drawing.Point(263, 306)
        Me.cmdContinueAudit.Name = "cmdContinueAudit"
        Me.cmdContinueAudit.Size = New System.Drawing.Size(84, 54)
        Me.cmdContinueAudit.TabIndex = 0
        Me.cmdContinueAudit.Text = "C&ontinue Audit Report"
        Me.cmdContinueAudit.UseVisualStyleBackColor = True
        '
        'cmdCancelAudit
        '
        Me.cmdCancelAudit.Location = New System.Drawing.Point(348, 306)
        Me.cmdCancelAudit.Name = "cmdCancelAudit"
        Me.cmdCancelAudit.Size = New System.Drawing.Size(75, 54)
        Me.cmdCancelAudit.TabIndex = 1
        Me.cmdCancelAudit.Text = "Cancel Audit Re&port"
        Me.cmdCancelAudit.UseVisualStyleBackColor = True
        '
        'grdCashJournal
        '
        Me.grdCashJournal.Location = New System.Drawing.Point(12, 12)
        Me.grdCashJournal.Name = "grdCashJournal"
        Me.grdCashJournal.OcxState = CType(resources.GetObject("grdCashJournal.OcxState"), System.Windows.Forms.AxHost.State)
        Me.grdCashJournal.Size = New System.Drawing.Size(486, 163)
        Me.grdCashJournal.TabIndex = 2
        '
        'fraEditControls
        '
        Me.fraEditControls.Controls.Add(Me.lblTransDate)
        Me.fraEditControls.Controls.Add(Me.lblIndex)
        Me.fraEditControls.Controls.Add(Me.dteTransDate)
        Me.fraEditControls.Location = New System.Drawing.Point(114, 181)
        Me.fraEditControls.Name = "fraEditControls"
        Me.fraEditControls.Size = New System.Drawing.Size(294, 119)
        Me.fraEditControls.TabIndex = 3
        Me.fraEditControls.TabStop = False
        '
        'lblIndex
        '
        Me.lblIndex.Location = New System.Drawing.Point(23, 29)
        Me.lblIndex.Name = "lblIndex"
        Me.lblIndex.Size = New System.Drawing.Size(235, 55)
        Me.lblIndex.TabIndex = 4
        Me.lblIndex.Text = "Label1"
        '
        'dteTransDate
        '
        Me.dteTransDate.CustomFormat = "MM/dd/yyyy"
        Me.dteTransDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dteTransDate.Location = New System.Drawing.Point(92, 89)
        Me.dteTransDate.Name = "dteTransDate"
        Me.dteTransDate.Size = New System.Drawing.Size(109, 20)
        Me.dteTransDate.TabIndex = 5
        '
        'cmdApply
        '
        Me.cmdApply.Location = New System.Drawing.Point(112, 306)
        Me.cmdApply.Name = "cmdApply"
        Me.cmdApply.Size = New System.Drawing.Size(75, 54)
        Me.cmdApply.TabIndex = 6
        Me.cmdApply.Text = "&Apply"
        Me.cmdApply.UseVisualStyleBackColor = True
        '
        'cmdCancel
        '
        Me.cmdCancel.Location = New System.Drawing.Point(188, 306)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.Size = New System.Drawing.Size(75, 54)
        Me.cmdCancel.TabIndex = 7
        Me.cmdCancel.Text = "&Cancel"
        Me.cmdCancel.UseVisualStyleBackColor = True
        '
        'cmdRefresh
        '
        Me.cmdRefresh.Location = New System.Drawing.Point(414, 181)
        Me.cmdRefresh.Name = "cmdRefresh"
        Me.cmdRefresh.Size = New System.Drawing.Size(75, 60)
        Me.cmdRefresh.TabIndex = 8
        Me.cmdRefresh.Text = "&Refresh"
        Me.cmdRefresh.UseVisualStyleBackColor = True
        '
        'lblTransDate
        '
        Me.lblTransDate.AutoSize = True
        Me.lblTransDate.Location = New System.Drawing.Point(23, 95)
        Me.lblTransDate.Name = "lblTransDate"
        Me.lblTransDate.Size = New System.Drawing.Size(63, 13)
        Me.lblTransDate.TabIndex = 5
        Me.lblTransDate.Text = "Trans Date:"
        '
        'frmEditCash
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(502, 362)
        Me.Controls.Add(Me.cmdRefresh)
        Me.Controls.Add(Me.cmdCancel)
        Me.Controls.Add(Me.cmdApply)
        Me.Controls.Add(Me.fraEditControls)
        Me.Controls.Add(Me.grdCashJournal)
        Me.Controls.Add(Me.cmdCancelAudit)
        Me.Controls.Add(Me.cmdContinueAudit)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.MaximizeBox = False
        Me.Name = "frmEditCash"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Cash Journal Editor"
        CType(Me.grdCashJournal, System.ComponentModel.ISupportInitialize).EndInit()
        Me.fraEditControls.ResumeLayout(False)
        Me.fraEditControls.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents cmdContinueAudit As Button
    Friend WithEvents cmdCancelAudit As Button
    Friend WithEvents grdCashJournal As AxMSFlexGridLib.AxMSFlexGrid
    Friend WithEvents fraEditControls As GroupBox
    Friend WithEvents lblIndex As Label
    Friend WithEvents dteTransDate As DateTimePicker
    Friend WithEvents cmdApply As Button
    Friend WithEvents cmdCancel As Button
    Friend WithEvents cmdRefresh As Button
    Friend WithEvents lblTransDate As Label
End Class
