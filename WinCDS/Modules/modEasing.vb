Imports System.Drawing
Module modEasing
    Public Sub DimAllForms()
        Dim F As Object, FF As Form
        'If IsIDE Or IsDevelopment Then Exit Sub
        'For Each F In Forms
        For Each F In My.Application.OpenForms
            FF = F
            If Left(FF.Name, 8) = "MainMenu" Then GoTo SkipForm
            If FF.Name = "frmPermissionMonitor" Then GoTo SkipForm
            If F.Visible Then Easing_DimForm(FF)
SkipForm:
        Next
    End Sub

    Public Sub UnDimAllForms()
        Dim F As Object, FF As Form
        'If IsIDE Or IsDevelopment Then Exit Sub
        For Each F In My.Application.OpenForms
            FF = F
            If Left(FF.Name, 8) = "MainMenu" Then GoTo SkipForm
            Easing_UnDimForm(FF)
SkipForm:
        Next
    End Sub

    Public Sub Easing_UnDimForm(ByRef F As Form)
        On Error Resume Next
        Dim C As Object
        For Each C In F.Controls
            If Left(C.Name, 9) = "picDimmer" Then
                ' Some events, such as ComboBox_Click(), do not allow unloading...  Hide it no matter what..
                C.Visible = False
                F.Controls.Remove(C)
            End If
        Next
    End Sub

    Public Sub Easing_DimForm(ByRef F As Form, Optional ByVal Rate As Integer = 192)
        On Error Resume Next
        'Dim X As PictureBox
        Dim X As New PictureBox
        'X = F.Controls.Add("vb.picturebox", "picDimmer" & Second(Now) & "_" & Random(1000))
        X.Name = "picDimmer" & Second(Now) & "_" & Random(1000)
        F.Controls.Add(X)

        Dim PictureBoxHeight As Integer, PictureBoxLeft As Integer, PictureBoxWidth As Integer, PictureBoxTop As Integer
        Dim FormInsideHeight As Integer, FormInsideWidth As Integer
        Dim FormAutoRedrawValue As Boolean

        'X.AutoRedraw = True
        X.Refresh()
        X.BackColor = Color.Blue
        'X.Move 0, 0, F.Width, F.Height
        X.Location = New Point(0, 0)
        X.Size = New Size(F.Width, F.Height)
        'X.Cls
        X.Image = New Bitmap(X.ClientSize.Width, X.ClientSize.Height)

        'X.Picture = CaptureForm(F)
        X.Image = CaptureForm(F)
        'X.FillColor = RGB(255, 255, 255)
        X.BackColor = Color.FromArgb(255, 255, 255)   '-> backcolor is replacement for fillcolor in vb.net

        'X.FillStyle = vbSolid
        'X.DrawMode = vbMaskPen  Drawmode is not available in vb.net. Its functionality available as default.

        Rate = FitRange(0, Rate, 255)
        'X.Line(0, 0)-(X.Width, X.Height), RGB(Rate, Rate, Rate), BF


        X.Visible = True
        'X.Move -(F.Width - F.ScaleWidth), -(F.Height - F.ScaleHeight)
        X.Location = New Point(-(F.Width - F.ClientSize.Width), -(F.Height - F.ClientSize.Height))
        'X.ZOrder 0
        X.BringToFront()
    End Sub

End Module
