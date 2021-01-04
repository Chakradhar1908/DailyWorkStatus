<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class ServiceIntake
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
        Me.txtServiceOrderNumber = New System.Windows.Forms.TextBox()
        Me.lblServiceOrderNumber = New System.Windows.Forms.Label()
        Me.txtLocation = New System.Windows.Forms.TextBox()
        Me.txtVendor = New System.Windows.Forms.TextBox()
        Me.txtMode = New System.Windows.Forms.TextBox()
        Me.cboImage = New System.Windows.Forms.ComboBox()
        Me.SuspendLayout()
        '
        'txtServiceOrderNumber
        '
        Me.txtServiceOrderNumber.Location = New System.Drawing.Point(90, 57)
        Me.txtServiceOrderNumber.Name = "txtServiceOrderNumber"
        Me.txtServiceOrderNumber.Size = New System.Drawing.Size(100, 20)
        Me.txtServiceOrderNumber.TabIndex = 0
        '
        'lblServiceOrderNumber
        '
        Me.lblServiceOrderNumber.AutoSize = True
        Me.lblServiceOrderNumber.Location = New System.Drawing.Point(21, 64)
        Me.lblServiceOrderNumber.Name = "lblServiceOrderNumber"
        Me.lblServiceOrderNumber.Size = New System.Drawing.Size(39, 13)
        Me.lblServiceOrderNumber.TabIndex = 1
        Me.lblServiceOrderNumber.Text = "Label1"
        '
        'txtLocation
        '
        Me.txtLocation.Location = New System.Drawing.Point(90, 96)
        Me.txtLocation.Name = "txtLocation"
        Me.txtLocation.Size = New System.Drawing.Size(100, 20)
        Me.txtLocation.TabIndex = 2
        '
        'txtVendor
        '
        Me.txtVendor.Location = New System.Drawing.Point(90, 135)
        Me.txtVendor.Name = "txtVendor"
        Me.txtVendor.Size = New System.Drawing.Size(100, 20)
        Me.txtVendor.TabIndex = 3
        '
        'txtMode
        '
        Me.txtMode.Location = New System.Drawing.Point(90, 174)
        Me.txtMode.Name = "txtMode"
        Me.txtMode.Size = New System.Drawing.Size(100, 20)
        Me.txtMode.TabIndex = 4
        '
        'cboImage
        '
        Me.cboImage.FormattingEnabled = True
        Me.cboImage.Location = New System.Drawing.Point(90, 200)
        Me.cboImage.Name = "cboImage"
        Me.cboImage.Size = New System.Drawing.Size(121, 21)
        Me.cboImage.TabIndex = 5
        '
        'ServiceIntake
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(800, 450)
        Me.Controls.Add(Me.cboImage)
        Me.Controls.Add(Me.txtMode)
        Me.Controls.Add(Me.txtVendor)
        Me.Controls.Add(Me.txtLocation)
        Me.Controls.Add(Me.lblServiceOrderNumber)
        Me.Controls.Add(Me.txtServiceOrderNumber)
        Me.Name = "ServiceIntake"
        Me.Text = "ServiceIntake"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents txtServiceOrderNumber As TextBox
    Friend WithEvents lblServiceOrderNumber As Label
    Friend WithEvents txtLocation As TextBox
    Friend WithEvents txtVendor As TextBox
    Friend WithEvents txtMode As TextBox
    Friend WithEvents cboImage As ComboBox
End Class
