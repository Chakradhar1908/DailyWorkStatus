Imports Microsoft.VisualBasic.PowerPacks.Printing.Compatibility.VB6
Public Class PrinterSelector
    Public Event Click()
    Private mSelectedPrinter As String
    Private mAllowDYMO As Boolean
    Private mAutoSelect As Boolean
    Private Const d_AllowDYMO As Boolean = False
    Private Const d_AutoSelect As Boolean = False
    Dim optCount As Integer

    Private Sub optPrinter0_CheckedChanged(sender As Object, e As EventArgs) Handles optPrinter0.CheckedChanged
        ' Change the selected printer...
        'mSelectedPrinter = optPrinter(Index).Caption
        'If AutoSelect Then SetPrinter mSelectedPrinter
        'RaiseEvent Click()

        mSelectedPrinter = CType(sender, RadioButton).Text
        If AutoSelect Then SetPrinter(mSelectedPrinter)
        RaiseEvent Click()
    End Sub

    Public Property AutoSelect() As Boolean
        Get
            AutoSelect = mAutoSelect
        End Get
        Set(value As Boolean)
            mAutoSelect = value
            Refresh()
        End Set
    End Property

    Public Property AllowDYMO() As Boolean
        Get
            AllowDYMO = mAllowDYMO
        End Get
        Set(value As Boolean)
            mAllowDYMO = value
            Refresh()
        End Set
    End Property

    Private Sub optPrinter0_MouseDown(sender As Object, e As MouseEventArgs) Handles optPrinter0.MouseDown
        If Not IsDevelopment() Then Exit Sub
        'If Button = vbRightButton Then PrinterStat optPrinter(Index).Caption
        'If e.Button = vbRightButton Then PrinterStat optPrinter(Index).Caption
        If e.Button = MouseButtons.Right Then PrinterStat(CType(sender, RadioButton).Text)
    End Sub

    Private Sub Scroller_Scroll(sender As Object, e As ScrollEventArgs) Handles Scroller.Scroll
        Dim I As Integer
        Dim optCount As Integer

        'For I = 1 To optPrinter.Count - 1
        '    optPrinter(I).Top = (I - Scroller.Value) * optPrinter(0).Height
        '    '    Debug.Print I, optPrinter(I).Top
        'Next
        '  Debug.Print "-", UserControl.Height

        For Each C As Control In Me.Controls
            If Mid(C.Name, 1, 10) = "optPrinter" Then
                optCount = optCount + 1
            End If
        Next

        For I = 1 To optCount
            For Each C As Control In Me.Controls
                If C.Name = "optPrinter" & I Then
                    C.Top = (I - Scroller.Value) * optPrinter0.Height
                    Exit For
                End If
            Next
        Next
    End Sub

    Private Sub Scroller_ValueChanged(sender As Object, e As EventArgs) Handles Scroller.ValueChanged
        Dim I As Integer

        'For I = 1 To optPrinter.Count - 1
        '    optPrinter(I).Top = (I - Scroller.Value) * optPrinter(0).Height
        '    '    Debug.Print I, optPrinter(I).Top
        'Next
        '  Debug.Print "-", UserControl.Height
        optCount = 0
        For Each C As Control In Me.Controls
            If Mid(C.Name, 1, 10) = "optPrinter" Then
                optCount = optCount + 1
            End If
        Next

        For I = 1 To optCount - 1
            For Each C As Control In Me.Controls
                If C.Name = "optPrinter" & I Then
                    C.Top = (I - Scroller.Value) * optPrinter0.Height
                    Exit For
                End If
            Next
        Next
    End Sub

    Public Sub Refresh()
        ' Load the list of printers.
        Dim I As Integer
        Dim El As Printer, DN As String, Show As Boolean
        'Dim R As RadioButton

        'If optPrinter.Count > 1 Then ' clear them
        '    For I = 1 To optPrinter.Count - 1
        '        Unload optPrinter(I)
        '    Next
        'End If
        optCount = 0
        For Each C As Control In Me.Controls
            If Mid(C.Name, 1, 10) = "optPrinter" Then
                optCount = optCount + 1
            End If
        Next
        If optCount > 1 Then
            For I = 1 To optCount - 1
                For Each C As Control In Me.Controls
                    If C.Name = "optPrinter" & I Then
                        'C.Hide()
                        Me.Controls.Remove(C)
                        Exit For
                    End If
                Next
                'R.Name = "optPrinter" & I
                'Me.Controls.Remove(R)
            Next
        End If

        optCount = 1
        For Each El In Printers
            DN = UCase(El.DeviceName)
            Show = True
            ' PROHIBIT CERTAIN DEVICES FROM THIS LIST!
            If DN = "FAX" Or DN Like "FAX*" Or DN Like "*FAX" Or DN Like "*FAX*" Then Show = False
            If DN Like "LPT" Or DN Like "LPT*" Or DN Like "*LPT" Or DN Like "*LPT*" Then Show = False
            If DN Like "MS" Or DN Like "MS*" Or DN Like "*MS" Or DN Like "*MS*" Then
                If Not DN Like "SAMSUNG" And Not DN Like "*SAMSUNG" And Not DN Like "SAMSUNG*" And Not DN Like "*SAMSUNG*" Then
                    Show = False
                End If
            End If
            If DN Like "VB5" Or DN Like "VB5*" Or DN Like "*VB5" Or DN Like "*VB5*" Then Show = False
            If DN Like "IMAGE WRITE" Or DN Like "IMAGE WRITE*" Or DN Like "*IMAGE WRITE" Or DN Like "*IMAGE WRITE*" Then Show = False
            If DN Like "INTUIT" Or DN Like "INTUIT*" Or DN Like "*INTUIT" Or DN Like "*INTUIT*" Then Show = False

            If Not AllowDYMO Then
                If DN Like "DYMO" Or DN Like "DYMO*" Or DN Like "*DYMO" Or DN Like "*DYMO*" Then Show = False
            End If

            If IsDevelopment() Then
                'If LCase(DN) Like "*redirect*" Then Show = False
                If LCase(DN) Like "*(copy*" Then Show = False
                If DN Like "*XPS*" Then Show = False
                If DN Like "QUICKBOOK*" Then Show = False
                If DN Like "*INVENTORY2*" Then Show = False
                If LCase(DN) Like "*onenote*" Then Show = False
                If LCase(DN) Like "*utepdf*" Then Show = False
            End If

            If Show Then
                'Load optPrinter(optPrinter.Count)
                'For Each C As Control In Me.Controls
                '    If C.Name = "optPrinter" & optCount Then
                '        C.Show()
                '        Exit For
                '    End If
                'Next

                'With optPrinter(optPrinter.Count - 1)
                '    .Caption = El.DeviceName
                '    .Move 0, (optPrinter.Count - 2) * optPrinter(0).Height, UserControl.Width - IIf(Scroller.Visible, Scroller.Width, 0), .Height
                '    .Visible = True
                '    If .Top + .Height > UserControl.Height Then Scroller.Visible = True
                '    ' Check it, if it's default..
                'End With
                'For Each C As Control In Me.Controls
                '    If C.Name = "optPrinter" & optCount - 1 Then
                '        C.Text = El.DeviceName
                '        C.Location = New Point(0, (optCount - 2) * optPrinter0.Height)
                '        C.Size = New Size(Me.Width - IIf(Scroller.Visible, Scroller.Width, 0), C.Height)
                '        C.Visible = True
                '        If C.Top + C.Height > Me.Height Then Scroller.Visible = True
                '    End If
                'Next

                Dim optP As RadioButton
                optP = New RadioButton
                optP.Name = "optPrinter" & optCount
                optP.Text = El.DeviceName
                optCount = optCount + 1
                optP.Location = New Point(0, (optCount - 2) * optPrinter0.Height)
                If optP.Top + optP.Height > Me.Height Then Scroller.Visible = True
                Me.Controls.Add(optP)
            End If
        Next

        AdjustScroller()
    End Sub

    Private Sub AdjustScroller()
        Dim sMax As Integer, sPage As Integer

        'sPage = CInt(UserControl.Height / optPrinter(0).Height)
        sPage = CInt(Me.Height / optPrinter0.Height)
        'If UserControl.Height <> sPage * optPrinter(0).Height Then UserControl.Height = sPage * optPrinter(0).Height
        If Height <> sPage * optPrinter0.Height Then Me.Height = sPage * optPrinter0.Height
        'sMax = optPrinter.Count - sPage
        sMax = optCount - sPage
        Scroller.LargeChange = sPage
        Scroller.SmallChange = 1

        If sMax > 1 Then
            Scroller.Minimum = 1
            Scroller.Maximum = sMax
            Scroller.Visible = True
        Else
            Scroller.Minimum = 1
            Scroller.Maximum = 1
            Scroller.Visible = False
        End If
        'Scroller.Move UserControl.Width - Scroller.Width, 0, Scroller.Width, sPage * optPrinter(0).Height
        Scroller.Location = New Point(Me.Width - Scroller.Width, 0)
        Scroller.Size = New Size(Scroller.Width, sPage * optPrinter0.Height)
    End Sub

    Public Function GetSelectedPrinter() As Printer
        ' Return the printer object corresponding to the selected entry.
        Dim El As Printer
        On Error GoTo NoPr
        For Each El In Printers
            If El.DeviceName = mSelectedPrinter Then
                GetSelectedPrinter = El
                Exit Function
            End If
        Next
NoPr:
    End Function

    Public Function GetSelectedPrinterName() As String
        ' Return the printer object corresponding to the selected entry.
        Dim El As Printer
        For Each El In Printers
            If El.DeviceName = mSelectedPrinter Then
                GetSelectedPrinterName = El.DeviceName
                Exit Function
            End If
        Next
NoPr:
    End Function

    Public Sub SetSelectedPrinter(Optional ByVal vData As String = "-")
        Dim El As Printer, I As Integer
        If vData = "-" Then vData = mSelectedPrinter
        If vData = "" Then mSelectedPrinter = ""

        For Each El In Printers
            If El.DeviceName = vData Then mSelectedPrinter = vData
        Next

        'For I = 1 To optPrinter.Count - 1
        '    optPrinter(I).Value = (optPrinter(I).Caption = vData)
        'Next
        '  Err.Raise -11377, , "Invalid printer."   ' Bad printer.
        optCount = 0
        For Each C As Control In Me.Controls
            If Mid(C.Name, 1, 10) = "optPrinter" Then
                optCount = optCount + 1
            End If
        Next
        For I = 1 To optCount - 1
            For Each C As RadioButton In Me.Controls
                If C.Name = "optPrinter" & I Then
                    C.Checked = (C.Text = vData)
                End If
            Next
        Next
    End Sub

    Private Sub PrinterSelector_Resize(sender As Object, e As EventArgs) Handles MyBase.Resize
        AdjustScroller()
        Dim I As Integer
        'For I = 1 To optPrinter.Count - 1
        '    optPrinter(I).Width = UserControl.Width - IIf(Scroller.Visible, Scroller.Width, 0)
        'Next
        For Each C As Control In Me.Controls
            If Mid(C.Name, 1, 10) = "optPrinter" Then
                C.Width = Me.Width - IIf(Scroller.Visible, Scroller.Width, 0)
            End If
        Next
    End Sub

    Public Sub New()
        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        AllowDYMO = d_AllowDYMO
        AutoSelect = d_AutoSelect
        Refresh()
    End Sub
End Class
