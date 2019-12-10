<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmNotes
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
        Me.Notes_Frame = New System.Windows.Forms.GroupBox()
        Me.txtOldNotes = New System.Windows.Forms.TextBox()
        Me.txtNewNotes = New System.Windows.Forms.TextBox()
        Me.lblOldNotes = New System.Windows.Forms.Label()
        Me.lblNewNotes = New System.Windows.Forms.Label()
        Me.cmdSave = New System.Windows.Forms.Button()
        Me.cmdPrint = New System.Windows.Forms.Button()
        Me.cmdCancel = New System.Windows.Forms.Button()
        Me.Notes_Frame.SuspendLayout()
        Me.SuspendLayout()
        '
        'Notes_Frame
        '
        Me.Notes_Frame.Controls.Add(Me.cmdCancel)
        Me.Notes_Frame.Controls.Add(Me.cmdPrint)
        Me.Notes_Frame.Controls.Add(Me.cmdSave)
        Me.Notes_Frame.Controls.Add(Me.lblNewNotes)
        Me.Notes_Frame.Controls.Add(Me.lblOldNotes)
        Me.Notes_Frame.Controls.Add(Me.txtNewNotes)
        Me.Notes_Frame.Controls.Add(Me.txtOldNotes)
        Me.Notes_Frame.Location = New System.Drawing.Point(8, 6)
        Me.Notes_Frame.Name = "Notes_Frame"
        Me.Notes_Frame.Size = New System.Drawing.Size(601, 210)
        Me.Notes_Frame.TabIndex = 1
        Me.Notes_Frame.TabStop = False
        Me.Notes_Frame.Text = " Not&es:"
        '
        'txtOldNotes
        '
        Me.txtOldNotes.BackColor = System.Drawing.SystemColors.ButtonFace
        Me.txtOldNotes.Location = New System.Drawing.Point(12, 19)
        Me.txtOldNotes.Multiline = True
        Me.txtOldNotes.Name = "txtOldNotes"
        Me.txtOldNotes.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.txtOldNotes.Size = New System.Drawing.Size(522, 82)
        Me.txtOldNotes.TabIndex = 0
        '
        'txtNewNotes
        '
        Me.txtNewNotes.Location = New System.Drawing.Point(12, 104)
        Me.txtNewNotes.Multiline = True
        Me.txtNewNotes.Name = "txtNewNotes"
        Me.txtNewNotes.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.txtNewNotes.Size = New System.Drawing.Size(522, 100)
        Me.txtNewNotes.TabIndex = 1
        '
        'lblOldNotes
        '
        Me.lblOldNotes.AutoSize = True
        Me.lblOldNotes.Location = New System.Drawing.Point(541, 22)
        Me.lblOldNotes.Name = "lblOldNotes"
        Me.lblOldNotes.Size = New System.Drawing.Size(54, 13)
        Me.lblOldNotes.TabIndex = 2
        Me.lblOldNotes.Text = "Ol&d Notes"
        '
        'lblNewNotes
        '
        Me.lblNewNotes.AutoSize = True
        Me.lblNewNotes.Location = New System.Drawing.Point(541, 191)
        Me.lblNewNotes.Name = "lblNewNotes"
        Me.lblNewNotes.Size = New System.Drawing.Size(60, 13)
        Me.lblNewNotes.TabIndex = 3
        Me.lblNewNotes.Text = "Ne&w Notes"
        '
        'cmdSave
        '
        Me.cmdSave.Location = New System.Drawing.Point(542, 44)
        Me.cmdSave.Name = "cmdSave"
        Me.cmdSave.Size = New System.Drawing.Size(53, 49)
        Me.cmdSave.TabIndex = 4
        Me.cmdSave.Text = "&Save"
        Me.cmdSave.UseVisualStyleBackColor = True
        '
        'cmdPrint
        '
        Me.cmdPrint.Location = New System.Drawing.Point(542, 89)
        Me.cmdPrint.Name = "cmdPrint"
        Me.cmdPrint.Size = New System.Drawing.Size(53, 49)
        Me.cmdPrint.TabIndex = 5
        Me.cmdPrint.Text = "&Print"
        Me.cmdPrint.UseVisualStyleBackColor = True
        '
        'cmdCancel
        '
        Me.cmdCancel.Location = New System.Drawing.Point(542, 137)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.Size = New System.Drawing.Size(53, 49)
        Me.cmdCancel.TabIndex = 6
        Me.cmdCancel.Text = "&Cancel"
        Me.cmdCancel.UseVisualStyleBackColor = True
        '
        'frmNotes
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(622, 219)
        Me.ControlBox = False
        Me.Controls.Add(Me.Notes_Frame)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow
        Me.Name = "frmNotes"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Notes"
        Me.Notes_Frame.ResumeLayout(False)
        Me.Notes_Frame.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents Notes_Frame As GroupBox
    Friend WithEvents cmdCancel As Button
    Friend WithEvents cmdPrint As Button
    Friend WithEvents cmdSave As Button
    Friend WithEvents lblNewNotes As Label
    Friend WithEvents lblOldNotes As Label
    Friend WithEvents txtNewNotes As TextBox
    Friend WithEvents txtOldNotes As TextBox
End Class
