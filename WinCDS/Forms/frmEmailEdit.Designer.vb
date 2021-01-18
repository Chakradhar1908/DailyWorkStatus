<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class frmEmailEdit
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
        Me.fraEditTemplate = New System.Windows.Forms.GroupBox()
        Me.cmdOK = New System.Windows.Forms.Button()
        Me.cmdCancel = New System.Windows.Forms.Button()
        Me.RTBEditTemplate = New WinCDS.RTBCompose()
        Me.fraEditTemplate.SuspendLayout()
        Me.SuspendLayout()
        '
        'fraEditTemplate
        '
        Me.fraEditTemplate.Controls.Add(Me.cmdCancel)
        Me.fraEditTemplate.Controls.Add(Me.cmdOK)
        Me.fraEditTemplate.Controls.Add(Me.RTBEditTemplate)
        Me.fraEditTemplate.Location = New System.Drawing.Point(6, 4)
        Me.fraEditTemplate.Name = "fraEditTemplate"
        Me.fraEditTemplate.Size = New System.Drawing.Size(482, 316)
        Me.fraEditTemplate.TabIndex = 0
        Me.fraEditTemplate.TabStop = False
        Me.fraEditTemplate.Text = "Edit Template:"
        '
        'cmdOK
        '
        Me.cmdOK.Location = New System.Drawing.Point(169, 264)
        Me.cmdOK.Name = "cmdOK"
        Me.cmdOK.Size = New System.Drawing.Size(64, 46)
        Me.cmdOK.TabIndex = 1
        Me.cmdOK.Text = "&Save"
        Me.cmdOK.UseVisualStyleBackColor = True
        '
        'cmdCancel
        '
        Me.cmdCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cmdCancel.Location = New System.Drawing.Point(249, 264)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.Size = New System.Drawing.Size(64, 46)
        Me.cmdCancel.TabIndex = 2
        Me.cmdCancel.Text = "&Cancel"
        Me.cmdCancel.UseVisualStyleBackColor = True
        '
        'RTBEditTemplate
        '
        Me.RTBEditTemplate.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.RTBEditTemplate.IsDirty = False
        Me.RTBEditTemplate.Location = New System.Drawing.Point(6, 19)
        Me.RTBEditTemplate.Margins = 4
        Me.RTBEditTemplate.Name = "RTBEditTemplate"
        Me.RTBEditTemplate.SendButton = False
        Me.RTBEditTemplate.Size = New System.Drawing.Size(469, 239)
        Me.RTBEditTemplate.TabIndex = 0
        Me.RTBEditTemplate.TextRTF = "{\rtf1\ansi\ansicpg1252\deff0\deflang1033{\fonttbl{\f0\fnil\fcharset0 Microsoft S" &
    "ans Serif;}}" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "\viewkind4\uc1\pard\f0\fs17\par" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "}" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10)
        '
        'frmEmailEdit
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.cmdCancel
        Me.ClientSize = New System.Drawing.Size(492, 321)
        Me.Controls.Add(Me.fraEditTemplate)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmEmailEdit"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Edit Email Template"
        Me.fraEditTemplate.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents fraEditTemplate As GroupBox
    Friend WithEvents cmdCancel As Button
    Friend WithEvents cmdOK As Button
    Friend WithEvents RTBEditTemplate As RTBCompose
End Class
