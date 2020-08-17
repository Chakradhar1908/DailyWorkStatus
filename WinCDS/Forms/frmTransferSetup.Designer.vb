<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmTransferSetup
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmTransferSetup))
        Me.cmdTransfer0 = New System.Windows.Forms.Button()
        Me.cmdTransfer1 = New System.Windows.Forms.Button()
        Me.cmdTransfer2 = New System.Windows.Forms.Button()
        Me.cmdTransfer3 = New System.Windows.Forms.Button()
        Me.cmdGo = New System.Windows.Forms.Button()
        Me.updTr = New AxMSComCtl2.AxUpDown()
        Me.txtTransferNo = New System.Windows.Forms.TextBox()
        Me.grd = New AxMSFlexGridLib.AxMSFlexGrid()
        Me.txtInput = New System.Windows.Forms.TextBox()
        Me.cmdOK = New System.Windows.Forms.Button()
        Me.fraView = New System.Windows.Forms.GroupBox()
        Me.lblSchedule = New System.Windows.Forms.Label()
        Me.lblNote = New System.Windows.Forms.Label()
        Me.dtpSchedule = New System.Windows.Forms.DateTimePicker()
        Me.txtNote = New System.Windows.Forms.TextBox()
        Me.chkCompleted = New System.Windows.Forms.CheckBox()
        Me.txtSetupDate = New System.Windows.Forms.TextBox()
        Me.txtVendor = New System.Windows.Forms.TextBox()
        Me.txtDisplayNote = New System.Windows.Forms.TextBox()
        Me.txtStatus = New System.Windows.Forms.TextBox()
        Me.txtLineStatus = New System.Windows.Forms.TextBox()
        CType(Me.updTr, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.grd, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.fraView.SuspendLayout()
        Me.SuspendLayout()
        '
        'cmdTransfer0
        '
        Me.cmdTransfer0.Location = New System.Drawing.Point(323, 33)
        Me.cmdTransfer0.Name = "cmdTransfer0"
        Me.cmdTransfer0.Size = New System.Drawing.Size(75, 23)
        Me.cmdTransfer0.TabIndex = 0
        Me.cmdTransfer0.Text = "T&ransfer"
        Me.cmdTransfer0.UseVisualStyleBackColor = True
        '
        'cmdTransfer1
        '
        Me.cmdTransfer1.Location = New System.Drawing.Point(404, 33)
        Me.cmdTransfer1.Name = "cmdTransfer1"
        Me.cmdTransfer1.Size = New System.Drawing.Size(75, 23)
        Me.cmdTransfer1.TabIndex = 1
        Me.cmdTransfer1.Text = "&Void Line"
        Me.cmdTransfer1.UseVisualStyleBackColor = True
        '
        'cmdTransfer2
        '
        Me.cmdTransfer2.Location = New System.Drawing.Point(485, 33)
        Me.cmdTransfer2.Name = "cmdTransfer2"
        Me.cmdTransfer2.Size = New System.Drawing.Size(75, 23)
        Me.cmdTransfer2.TabIndex = 2
        Me.cmdTransfer2.Text = "Transfer A&ll"
        Me.cmdTransfer2.UseVisualStyleBackColor = True
        '
        'cmdTransfer3
        '
        Me.cmdTransfer3.Location = New System.Drawing.Point(585, 33)
        Me.cmdTransfer3.Name = "cmdTransfer3"
        Me.cmdTransfer3.Size = New System.Drawing.Size(75, 23)
        Me.cmdTransfer3.TabIndex = 3
        Me.cmdTransfer3.Text = "Vo&id All"
        Me.cmdTransfer3.UseVisualStyleBackColor = True
        '
        'cmdGo
        '
        Me.cmdGo.Location = New System.Drawing.Point(173, 10)
        Me.cmdGo.Name = "cmdGo"
        Me.cmdGo.Size = New System.Drawing.Size(75, 23)
        Me.cmdGo.TabIndex = 4
        Me.cmdGo.Text = "G&o"
        Me.cmdGo.UseVisualStyleBackColor = True
        '
        'updTr
        '
        Me.updTr.Location = New System.Drawing.Point(136, 2)
        Me.updTr.Name = "updTr"
        Me.updTr.OcxState = CType(resources.GetObject("updTr.OcxState"), System.Windows.Forms.AxHost.State)
        Me.updTr.Size = New System.Drawing.Size(17, 50)
        Me.updTr.TabIndex = 5
        '
        'txtTransferNo
        '
        Me.txtTransferNo.Location = New System.Drawing.Point(17, 12)
        Me.txtTransferNo.Name = "txtTransferNo"
        Me.txtTransferNo.Size = New System.Drawing.Size(100, 20)
        Me.txtTransferNo.TabIndex = 6
        '
        'grd
        '
        Me.grd.Location = New System.Drawing.Point(12, 131)
        Me.grd.Name = "grd"
        Me.grd.OcxState = CType(resources.GetObject("grd.OcxState"), System.Windows.Forms.AxHost.State)
        Me.grd.Size = New System.Drawing.Size(701, 160)
        Me.grd.TabIndex = 7
        '
        'txtInput
        '
        Me.txtInput.Location = New System.Drawing.Point(566, 317)
        Me.txtInput.Name = "txtInput"
        Me.txtInput.Size = New System.Drawing.Size(100, 20)
        Me.txtInput.TabIndex = 8
        '
        'cmdOK
        '
        Me.cmdOK.Location = New System.Drawing.Point(215, 297)
        Me.cmdOK.Name = "cmdOK"
        Me.cmdOK.Size = New System.Drawing.Size(75, 23)
        Me.cmdOK.TabIndex = 9
        Me.cmdOK.Text = "&OK"
        Me.cmdOK.UseVisualStyleBackColor = True
        '
        'fraView
        '
        Me.fraView.Controls.Add(Me.txtLineStatus)
        Me.fraView.Controls.Add(Me.txtStatus)
        Me.fraView.Controls.Add(Me.txtDisplayNote)
        Me.fraView.Controls.Add(Me.txtVendor)
        Me.fraView.Controls.Add(Me.txtSetupDate)
        Me.fraView.Controls.Add(Me.txtTransferNo)
        Me.fraView.Controls.Add(Me.updTr)
        Me.fraView.Controls.Add(Me.cmdGo)
        Me.fraView.Controls.Add(Me.cmdTransfer3)
        Me.fraView.Controls.Add(Me.cmdTransfer2)
        Me.fraView.Controls.Add(Me.cmdTransfer1)
        Me.fraView.Controls.Add(Me.cmdTransfer0)
        Me.fraView.Location = New System.Drawing.Point(42, 12)
        Me.fraView.Name = "fraView"
        Me.fraView.Size = New System.Drawing.Size(703, 84)
        Me.fraView.TabIndex = 10
        Me.fraView.TabStop = False
        Me.fraView.Text = "GroupBox1"
        '
        'lblSchedule
        '
        Me.lblSchedule.AutoSize = True
        Me.lblSchedule.Location = New System.Drawing.Point(52, 340)
        Me.lblSchedule.Name = "lblSchedule"
        Me.lblSchedule.Size = New System.Drawing.Size(87, 13)
        Me.lblSchedule.TabIndex = 11
        Me.lblSchedule.Text = "Scheduled Date:"
        '
        'lblNote
        '
        Me.lblNote.AutoSize = True
        Me.lblNote.Location = New System.Drawing.Point(62, 371)
        Me.lblNote.Name = "lblNote"
        Me.lblNote.Size = New System.Drawing.Size(33, 13)
        Me.lblNote.TabIndex = 12
        Me.lblNote.Text = "Note:"
        '
        'dtpSchedule
        '
        Me.dtpSchedule.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtpSchedule.Location = New System.Drawing.Point(145, 340)
        Me.dtpSchedule.Name = "dtpSchedule"
        Me.dtpSchedule.Size = New System.Drawing.Size(94, 20)
        Me.dtpSchedule.TabIndex = 13
        '
        'txtNote
        '
        Me.txtNote.Location = New System.Drawing.Point(113, 372)
        Me.txtNote.Name = "txtNote"
        Me.txtNote.Size = New System.Drawing.Size(152, 20)
        Me.txtNote.TabIndex = 14
        '
        'chkCompleted
        '
        Me.chkCompleted.AutoSize = True
        Me.chkCompleted.Location = New System.Drawing.Point(254, 339)
        Me.chkCompleted.Name = "chkCompleted"
        Me.chkCompleted.Size = New System.Drawing.Size(82, 17)
        Me.chkCompleted.TabIndex = 15
        Me.chkCompleted.Text = "Completed?"
        Me.chkCompleted.UseVisualStyleBackColor = True
        '
        'txtSetupDate
        '
        Me.txtSetupDate.Location = New System.Drawing.Point(277, 17)
        Me.txtSetupDate.Name = "txtSetupDate"
        Me.txtSetupDate.Size = New System.Drawing.Size(100, 20)
        Me.txtSetupDate.TabIndex = 7
        '
        'txtVendor
        '
        Me.txtVendor.Location = New System.Drawing.Point(595, 16)
        Me.txtVendor.Name = "txtVendor"
        Me.txtVendor.Size = New System.Drawing.Size(100, 20)
        Me.txtVendor.TabIndex = 8
        '
        'txtDisplayNote
        '
        Me.txtDisplayNote.Location = New System.Drawing.Point(194, 58)
        Me.txtDisplayNote.Name = "txtDisplayNote"
        Me.txtDisplayNote.Size = New System.Drawing.Size(100, 20)
        Me.txtDisplayNote.TabIndex = 9
        '
        'txtStatus
        '
        Me.txtStatus.Location = New System.Drawing.Point(460, 21)
        Me.txtStatus.Name = "txtStatus"
        Me.txtStatus.Size = New System.Drawing.Size(100, 20)
        Me.txtStatus.TabIndex = 10
        '
        'txtLineStatus
        '
        Me.txtLineStatus.Location = New System.Drawing.Point(181, 35)
        Me.txtLineStatus.Name = "txtLineStatus"
        Me.txtLineStatus.Size = New System.Drawing.Size(90, 20)
        Me.txtLineStatus.TabIndex = 11
        '
        'frmTransferSetup
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(800, 450)
        Me.Controls.Add(Me.chkCompleted)
        Me.Controls.Add(Me.txtNote)
        Me.Controls.Add(Me.dtpSchedule)
        Me.Controls.Add(Me.lblNote)
        Me.Controls.Add(Me.lblSchedule)
        Me.Controls.Add(Me.fraView)
        Me.Controls.Add(Me.cmdOK)
        Me.Controls.Add(Me.txtInput)
        Me.Controls.Add(Me.grd)
        Me.Name = "frmTransferSetup"
        Me.Text = "frmTransferSetup"
        CType(Me.updTr, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.grd, System.ComponentModel.ISupportInitialize).EndInit()
        Me.fraView.ResumeLayout(False)
        Me.fraView.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents cmdTransfer0 As Button
    Friend WithEvents cmdTransfer1 As Button
    Friend WithEvents cmdTransfer2 As Button
    Friend WithEvents cmdTransfer3 As Button
    Friend WithEvents cmdGo As Button
    Friend WithEvents updTr As AxMSComCtl2.AxUpDown
    Friend WithEvents txtTransferNo As TextBox
    Friend WithEvents grd As AxMSFlexGridLib.AxMSFlexGrid
    Friend WithEvents txtInput As TextBox
    Friend WithEvents cmdOK As Button
    Friend WithEvents fraView As GroupBox
    Friend WithEvents lblSchedule As Label
    Friend WithEvents lblNote As Label
    Friend WithEvents dtpSchedule As DateTimePicker
    Friend WithEvents txtNote As TextBox
    Friend WithEvents chkCompleted As CheckBox
    Friend WithEvents txtSetupDate As TextBox
    Friend WithEvents txtVendor As TextBox
    Friend WithEvents txtDisplayNote As TextBox
    Friend WithEvents txtStatus As TextBox
    Friend WithEvents txtLineStatus As TextBox
End Class
