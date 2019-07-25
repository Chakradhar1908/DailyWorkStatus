<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class InvKitStock
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
        Me.fra = New System.Windows.Forms.GroupBox()
        Me.UGridIO1 = New WinCDS.UGridIO()
        Me.fra.SuspendLayout()
        Me.SuspendLayout()
        '
        'fra
        '
        Me.fra.Controls.Add(Me.UGridIO1)
        Me.fra.Location = New System.Drawing.Point(0, 0)
        Me.fra.Name = "fra"
        Me.fra.Size = New System.Drawing.Size(721, 266)
        Me.fra.TabIndex = 0
        Me.fra.TabStop = False
        Me.fra.Text = "GroupBox1"
        '
        'UGridIO1
        '
        Me.UGridIO1.Activated = False
        Me.UGridIO1.Col = 0
        Me.UGridIO1.firstrow = 0
        Me.UGridIO1.Loading = False
        Me.UGridIO1.Location = New System.Drawing.Point(20, 31)
        Me.UGridIO1.MaxCols = 2
        Me.UGridIO1.MaxRows = 10
        Me.UGridIO1.Name = "UGridIO1"
        Me.UGridIO1.Row = 0
        Me.UGridIO1.Size = New System.Drawing.Size(626, 184)
        Me.UGridIO1.TabIndex = 0
        '
        'InvKitStock
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(800, 450)
        Me.Controls.Add(Me.fra)
        Me.Name = "InvKitStock"
        Me.Text = "InvKitStock"
        Me.fra.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents fra As GroupBox
    Friend WithEvents UGridIO1 As UGridIO
End Class
