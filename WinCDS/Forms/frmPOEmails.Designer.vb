<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmPOEmails
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
        Me.cmdEditTemplate = New System.Windows.Forms.Button()
        Me.dtpRunAsDate = New System.Windows.Forms.DateTimePicker()
        Me.fraSelect = New System.Windows.Forms.GroupBox()
        Me.lstSelect = New System.Windows.Forms.ListBox()
        Me.SuspendLayout()
        '
        'cmdEditTemplate
        '
        Me.cmdEditTemplate.Location = New System.Drawing.Point(370, 92)
        Me.cmdEditTemplate.Name = "cmdEditTemplate"
        Me.cmdEditTemplate.Size = New System.Drawing.Size(75, 23)
        Me.cmdEditTemplate.TabIndex = 0
        Me.cmdEditTemplate.Text = "Button1"
        Me.cmdEditTemplate.UseVisualStyleBackColor = True
        '
        'dtpRunAsDate
        '
        Me.dtpRunAsDate.Location = New System.Drawing.Point(347, 25)
        Me.dtpRunAsDate.Name = "dtpRunAsDate"
        Me.dtpRunAsDate.Size = New System.Drawing.Size(200, 20)
        Me.dtpRunAsDate.TabIndex = 1
        '
        'fraSelect
        '
        Me.fraSelect.Location = New System.Drawing.Point(63, 25)
        Me.fraSelect.Name = "fraSelect"
        Me.fraSelect.Size = New System.Drawing.Size(200, 100)
        Me.fraSelect.TabIndex = 2
        Me.fraSelect.TabStop = False
        Me.fraSelect.Text = "GroupBox1"
        '
        'lstSelect
        '
        Me.lstSelect.FormattingEnabled = True
        Me.lstSelect.Location = New System.Drawing.Point(347, 162)
        Me.lstSelect.Name = "lstSelect"
        Me.lstSelect.Size = New System.Drawing.Size(120, 95)
        Me.lstSelect.TabIndex = 3
        '
        'frmPOEmails
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(800, 450)
        Me.Controls.Add(Me.lstSelect)
        Me.Controls.Add(Me.fraSelect)
        Me.Controls.Add(Me.dtpRunAsDate)
        Me.Controls.Add(Me.cmdEditTemplate)
        Me.Name = "frmPOEmails"
        Me.Text = "frmPOEmails"
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents cmdEditTemplate As Button
    Friend WithEvents dtpRunAsDate As DateTimePicker
    Friend WithEvents fraSelect As GroupBox
    Friend WithEvents lstSelect As ListBox
End Class
