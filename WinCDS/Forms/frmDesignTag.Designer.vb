<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmDesignTag
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
        Me.cmbPageAlign = New System.Windows.Forms.ComboBox()
        Me.imgPrintHelper = New System.Windows.Forms.PictureBox()
        CType(Me.imgPrintHelper, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'cmbPageAlign
        '
        Me.cmbPageAlign.FormattingEnabled = True
        Me.cmbPageAlign.Location = New System.Drawing.Point(377, 147)
        Me.cmbPageAlign.Name = "cmbPageAlign"
        Me.cmbPageAlign.Size = New System.Drawing.Size(121, 21)
        Me.cmbPageAlign.TabIndex = 0
        '
        'imgPrintHelper
        '
        Me.imgPrintHelper.Location = New System.Drawing.Point(404, 276)
        Me.imgPrintHelper.Name = "imgPrintHelper"
        Me.imgPrintHelper.Size = New System.Drawing.Size(100, 50)
        Me.imgPrintHelper.TabIndex = 1
        Me.imgPrintHelper.TabStop = False
        '
        'frmDesignTag
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(800, 450)
        Me.Controls.Add(Me.imgPrintHelper)
        Me.Controls.Add(Me.cmbPageAlign)
        Me.Name = "frmDesignTag"
        Me.Text = "frmDesignTag"
        CType(Me.imgPrintHelper, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents cmbPageAlign As ComboBox
    Friend WithEvents imgPrintHelper As PictureBox
End Class
