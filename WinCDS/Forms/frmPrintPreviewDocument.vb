Imports Microsoft.VisualBasic.PowerPacks.Printing.Compatibility.VB6
Public Class frmPrintPreviewDocument
    Public CallingForm As Object
    Public ReportName As String

    Public CurrentPage As Integer 'Used while viewing
    Public TotalPages As Integer
    Private SkipFormActivate As Boolean
    Private SkipKey As Boolean

    Private DoLandscape As Boolean
    Private NotifiedOverrun As Boolean

    Public Sub NewPage()
        Application.DoEvents()
        If OutputObject.CurrentY = 0 And PageNumber <> 0 Then
            Exit Sub 'If page is blank, do not add a page
        End If
        'picPicture.Visible = False

        If IsDevelopment() Then
            If OutputObject.CurrentY > Printer.ScaleHeight Then
                If Not NotifiedOverrun Then
                    'MsgBox "DEVELOPER NOTIFICATION:" & vbCrLf & "Printer overrun", , "Developer Notice", , , 3
                    MessageBox.Show("DEVELOPER NOTIFICATION:" & vbCrLf & "Printer overrun", "Developer Notice")
                    NotifiedOverrun = True
                End If
            End If
        End If

        PrintPageOverflowIndicator()  ' draws nice dotted lines at the page size so we know if it would overflow
        If PageNumber <> 0 Then SavePage(PageNumber) 'Save current page to temp file on disk
        'Set picPicture = Nothing
        'picPicture.Cls 'Clear the picturebox
        picPicture.Image = Nothing
        'Set OutputObject = Nothing
        PageNumber = PageNumber + 1 'Increment current page
        TotalPages = TotalPages + 1 'Increment total page count
        'Set OutputObject = picPicture
        'Load OutputObject
        'OutputObject.Visible = True
        'fraNavigate.ZOrder 0
        Application.DoEvents()
    End Sub

    Private Function SavePage(Optional ByVal N As Integer = 0) As Boolean
        If N <> 0 Then
            'SavePicture(picPicture.Image, PageFile(N))
            picPicture.Image.Save(PageFile(N))
        End If
        SavePage = True
    End Function

    Private ReadOnly Property PageFile(Optional ByVal N As Integer = 0) As String
        Get
            If N = 0 Then N = CurrentPage
            PageFile = GetTempDir() & "PP" & Format(N, "000") & ".tmp"
        End Get
    End Property

    Public Sub DataEnd()
        'If EndOfDocumentEnabled Then Exit Sub
        DoLandscape = Printer.Orientation = vbPRORLandscape
        'If picPicture.CurrentY = 0 And PageNumber <= 1 Then 'Nothing was printed
        If picPicture.Location.Y = 0 And PageNumber <= 1 Then
            PageNumber = 1
            TotalPages = 1
            'MousePointer = vbDefault
            Me.Cursor = Cursors.Default
            SkipFormActivate = False
            'Unload frmPrintPreviewMain
            frmPrintPreviewMain.Close()
            'Unload Me
            Me.Close()
            Exit Sub
            'ElseIf picPicture.CurrentY = 0 And PageNumber > 1 Then 'Page is blank
        ElseIf picPicture.Location.Y = 0 And PageNumber > 1 Then 'Page is blank
            'Unload picPicture(PageNumber)
            PageNumber = PageNumber - 1
            TotalPages = TotalPages - 1
        Else 'Page is not blank, so save first
            PrintPageOverflowIndicator()  ' draws nice dotted lines at the page size so we know if it would overflow
            SavePage(PageNumber)
            'MousePointer = 0
            Me.Cursor = Cursors.Default
        End If
        'picPicture.Cls 'Clear the picturebox
        picPicture.Image = Nothing
        'picPicture(PageNumber).Visible = False
        CurrentPage = 1
        LoadPage()
        'picPicture(CurrentPage).Visible = True
        'cmdNavigate(7).Enabled = (picPicture.Count - 1 > 1) 'Enable Goto button only if number of pages is greater than one
        'cmdNavigate(7).Enabled = (TotalPages > 1) 'Enable Goto button only if number of pages is greater than one
        cmdNavigate7.Enabled = (TotalPages > 1) 'Enable Goto button only if number of pages is greater than one
        'frmPrintPreviewMain.Caption = "Print Preview: " & ReportName & ", page " & CurrentPage & " of " & picPicture.Count - 1
        frmPrintPreviewMain.Text = "Print Preview: " & ReportName & ", page " & CurrentPage & " of " & TotalPages
        Show() 'Show PrintPreview module
        'MousePointer = vbDefault
        Me.Cursor = Cursors.Default
        Exit Sub
ErrorHandler:
        MessageBox.Show(Err.Description, Err.Number.ToString, MessageBoxButtons.OK, MessageBoxIcon.Warning)
        Debug.Assert(False) 'Pause only if in debug mode
    End Sub

    Private Function LoadPage(Optional ByVal N As Integer = 0) As Boolean
        On Error GoTo LoadFailed
        If N = 0 Then N = CurrentPage
        'picPicture.Cls
        picPicture.Image = Nothing
        'picPicture.Picture = LoadPictureStd(PageFile(N))
        picPicture.Image = LoadPictureStd(PageFile(N))
        LoadPage = True
        Exit Function

LoadFailed:
        Select Case Err.Number
            Case 53
                MessageBox.Show("Failed to find Print Preview Page #" & CurrentPage & vbCrLf & PageFile(N), "Temp Dir Access Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Case Else
                MessageBox.Show("Error loading Print Preview Page." & vbCrLf & PageFile(N))
        End Select
    End Function

    Private Sub frmPrintPreviewDocument_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ActiveLog("frmPrintPreviewDocument::Form_Load()", 9)
        SkipFormActivate = False
        PageNumber = 1
        TotalPages = 1
        Me.MdiParent = frmPrintPreviewMain
    End Sub

    Private Sub frmPrintPreviewDocument_Resize(sender As Object, e As EventArgs) Handles Me.Resize
        If Not Visible Then Exit Sub
        ActiveLog("frmPrintPreviewDocument::Form_Resize() -- Scalewidth=" & Me.ClientSize.Width & ", Scaleheight=" & Me.ClientSize.Height, 8)
        'Move 0, 0, frmPrintPreviewMain.ScaleWidth, frmPrintPreviewMain.ScaleHeight
        Me.Location = New Point(0, 0)
        Me.Size = New Size(frmPrintPreviewMain.ClientSize.Width, frmPrintPreviewMain.ClientSize.Height)
        'picPicture.Move 0 - 30, 0 - 30, ScaleWidth + 30, 15840 + 1440 + 30 '1440=1 Inch
        picPicture.Location = New Point(0 - 30, 0 - 30)
        picPicture.Size = New Size(Me.ClientSize.Width + 30, 15840 + 1440 + 30)

        On Error Resume Next ' Printer.Scalewidth can fail...
        'fraNavigate.Move ScaleWidth - fraNavigate.Width, ScaleHeight - fraNavigate.Height
        fraNavigate.Location = New Point(Me.ClientSize.Width - fraNavigate.Width, Me.ClientSize.Height - fraNavigate.Height)
        '  fraNavigate.Move Printer.ScaleWidth - fraNavigate.Width, ScaleHeight - fraNavigate.Height
        'fraNavigate.Move Printer.ScaleWidth, ScaleHeight - fraNavigate.Height
        fraNavigate.Location = New Point(Printer.ScaleWidth, Me.ClientSize.Height - fraNavigate.Height)
        lblHelp.Top = Me.ClientSize.Height - lblHelp.Height
    End Sub

    Private Sub frmPrintPreviewDocument_Activated(sender As Object, e As EventArgs) Handles Me.Activated
        If SkipFormActivate Then Exit Sub
        Application.DoEvents()
        'Move 0, 0, frmPrintPreviewMain.ScaleWidth, frmPrintPreviewMain.ScaleHeight
        Me.Location = New Point(0, 0)
        Me.Size = New Size(frmPrintPreviewMain.ClientSize.Width, frmPrintPreviewMain.ClientSize.Height)
        'MousePointer = vbHourglass
        Me.Cursor = Cursors.WaitCursor
        SkipFormActivate = True
        PageNumber = 1
        TotalPages = 1
    End Sub

End Class