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

End Class