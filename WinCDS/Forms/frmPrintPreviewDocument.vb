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
    Dim LoadBitmap As Boolean
    'Dim b As Bitmap
    Dim g As Graphics
    Public SrnoArray() As String
    Public LastNameArray() As String
    Public DateofClaimArray() As String
    Public TelephoneArray() As String
    Public ItemLineArray() As String

    Public Sub NewPage()
        'Application.DoEvents()

        If TypeOf OutputObject Is PictureBox Then
            'If OutputObject.Location.Y = 0 And PageNumber <> 0 Then
            If TopValue = 0 And PageNumber <> 0 Then
                Exit Sub 'If page is blank, do not add a page
            End If
        Else
            If OutputObject.CurrentY = 0 And PageNumber <> 0 Then
                Exit Sub 'If page is blank, do not add a page
            End If
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

        If TypeOf OutputObject Is PictureBox Then
        Else
            PrintPageOverflowIndicator()  ' draws nice dotted lines at the page size so we know if it would overflow
        End If

        If TypeOf OutputObject Is PictureBox Then
        Else
            If PageNumber <> 0 Then SavePage(PageNumber) 'Save current page to temp file on disk
        End If

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
        'Application.DoEvents()
    End Sub

    Private Function SavePage(Optional ByVal N As Integer = 0) As Boolean
        If N <> 0 Then
            'SavePicture(picPicture.Image, PageFile(N))
            'picPicture.Image.Save(PageFile(N), Imaging.ImageFormat.Png)
            picPicture.Image.Save(PageFile(N))
        End If
        SavePage = True
    End Function

    Private ReadOnly Property PageFile(Optional ByVal N As Integer = 0) As String
        Get
            If N = 0 Then N = CurrentPage
            PageFile = GetTempDir() & "PP" & Format(N, "000") & ".tmp"
            'PageFile = GetTempDir() & "PP" & Format(N, "000") & ".png"
        End Get
    End Property

    Public Sub DataEnd()
        'If EndOfDocumentEnabled Then Exit Sub
        If TypeOf OutputObject Is PictureBox Then
            'frmPrintPreviewMain.Text = "Print Preview: " & ReportName & ", page " & CurrentPage & " of " & TotalPages
            'Show() 'Show PrintPreview module
            'MousePointer = vbDefault
            Me.Cursor = Cursors.Default
            Exit Sub
        End If
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
        ServiceReports.ExecutePaint = False
        LoadBitmap = True
        SetButtonImage(btnPrint, 19)
    End Sub

    Private Sub frmPrintPreviewDocument_Resize(sender As Object, e As EventArgs) Handles Me.Resize
        If Not Visible Then Exit Sub
        ActiveLog("frmPrintPreviewDocument::Form_Resize() -- Scalewidth=" & Me.ClientSize.Width & ", Scaleheight=" & Me.ClientSize.Height, 8)
        'Move 0, 0, frmPrintPreviewMain.ScaleWidth, frmPrintPreviewMain.ScaleHeight
        Me.Location = New Point(0, 0)
        'Me.Size = New Size(frmPrintPreviewMain.ClientSize.Width, frmPrintPreviewMain.ClientSize.Height)
        'Me.Size = New Size(frmPrintPreviewMain.Width, frmPrintPreviewMain.Height)
        Me.Size = New Size(frmPrintPreviewMain.Width, frmPrintPreviewMain.Height)
        'picPicture.Move 0 - 30, 0 - 30, ScaleWidth + 30, 15840 + 1440 + 30 '1440=1 Inch
        'picPicture.Location = New Point(0 - 30, 0 - 30)
        picPicture.Location = New Point(0, 0)
        'picPicture.Size = New Size(Me.ClientSize.Width + 30, 15840 + 1440 + 30)
        picPicture.Size = New Size(Me.Width, Me.Height)

        'pnlPicture.Location = New Point(0, 0)
        'pnlPicture.Size = New Size(Me.Width, Me.Height)

        On Error Resume Next ' Printer.Scalewidth can fail...
        'fraNavigate.Move ScaleWidth - fraNavigate.Width, ScaleHeight - fraNavigate.Height
        'fraNavigate.Location = New Point(Me.ClientSize.Width - fraNavigate.Width, Me.ClientSize.Height - fraNavigate.Height)
        '  fraNavigate.Move Printer.ScaleWidth - fraNavigate.Width, ScaleHeight - fraNavigate.Height
        'fraNavigate.Move Printer.ScaleWidth, ScaleHeight - fraNavigate.Height
        fraNavigate.Location = New Point(500, 500)
        lblHelp.Top = Me.ClientSize.Height - lblHelp.Height
    End Sub

    Private Sub frmPrintPreviewDocument_Activated(sender As Object, e As EventArgs) Handles Me.Activated
        If SkipFormActivate Then Exit Sub
        Application.DoEvents()
        'Move 0, 0, frmPrintPreviewMain.ScaleWidth, frmPrintPreviewMain.ScaleHeight
        Me.Location = New Point(0, 0)
        'Me.Size = New Size(frmPrintPreviewMain.ClientSize.Width, frmPrintPreviewMain.ClientSize.Height)
        Me.Size = New Size(frmPrintPreviewMain.Width, frmPrintPreviewMain.Height)
        'MousePointer = vbHourglass
        Me.Cursor = Cursors.WaitCursor
        SkipFormActivate = True
        PageNumber = 1
        TotalPages = 1
    End Sub

    Public Sub picPicture_Paint(sender As Object, e As PaintEventArgs) Handles picPicture.Paint
        Dim MyBrush As New SolidBrush(Color.Black)

        'If NoPaint = True Then Exit Sub
        'If LoadBitmap = True Then
        'Dim b As Bitmap = New Bitmap(picPicture.Width, picPicture.Height)
        '    '    'Dim g As Graphics = Graphics.FromImage(b)
        'e.Graphics.FillRectangle(Brushes.White, New Rectangle(0, 0, b.Width, b.Height))
        '    '    'e.Graphics.DrawString(PrintText, New Font("Arial", 10), MyBrush, L, T)
        '    LoadBitmap = False
        'End If
        'Dim Y As Integer
        'OutputObject.FontName = "Arial"
        'OutputObject.FontSize = 18
        'PrintCentered(ReportTitle, 100, True)

        If ServiceReports.ExecutePaint = True Then
            e.Graphics.DrawString(ServiceReports.ReportTitle, New Font("Arial", 18), MyBrush, picPicture.Width / 6, 10)
            'Button1_Click(Button1, New EventArgs)
            e.Graphics.DrawString("Time: " & Format(Now, "h:mm:ss tt"), New Font("Arial", 8), MyBrush, 1, 10)
            If OutputToPrinter Then PageNumber = OutputObject.Page
            'e.Graphics.DrawString("Page: " & PageNumber, New Font("Arial", 8), MyBrush, 1000, 10)
            e.Graphics.DrawString(StoreSettings.Name & "    " & StoreSettings.Address & "    " & StoreSettings.City, New Font("Arial", 8), MyBrush, picPicture.Width / 6, 35)

            Select Case ServiceReports.Mode
                Case "SCR"
                    'OutputObject.CurrentX = 0
                    'OutputObject.CurrentY = 700
                    'PrintToTab(, "ServiceNo", 0)
                    'PrintToTab(, "ServiceNo", 0,,, 700, 0, True)
                    'PrintToPosition2(, "ServiceNo", ,,, 700, 0)
                    e.Graphics.DrawString("ServiceNo", New Font("Arial", 9, FontStyle.Bold), MyBrush, 0, 50)
                    'PrintToTab(, "DateOfClaim", 20)
                    'PrintToTab(, "DateOfClaim", 20,,, 700, 20, True)
                    'PrintToPosition2(, "DateOfClaim", ,,, 700, 1500)
                    e.Graphics.DrawString("DateOfClaim", New Font("Arial", 9, FontStyle.Bold), MyBrush, 100, 50)
                    'PrintToTab(, "Last", 40)
                    'PrintToTab(, "Last", 40,,, 700, 40, True)
                    'PrintToPosition2(, "Last", ,,, 700, 3000)
                    e.Graphics.DrawString("Last", New Font("Arial", 9, FontStyle.Bold), MyBrush, 200, 50)
                    'PrintToTab(, "Telephone", 60)
                    'PrintToTab(, "Telephone", 60,,, 700, 60, True)
                    'PrintToPosition2(, "Telephone", ,,, 700, 4800)
                    e.Graphics.DrawString("Telephone", New Font("Arial", 9, FontStyle.Bold), MyBrush, 300, 50)
                    '    Case "SPR"
                    '        PrintAligned("PartsOrderNo", , 10, Y, True)
                    '        PrintAligned("Status", , 1300, Y, True)
                    '        PrintAligned("ServiceNo", , 2000, Y, True)
                    '        PrintAligned("Vendor", , 3000, Y, True)
                    '        PrintAligned("DateOfClaim", , 5900, Y, True)
                    '        PrintAligned("Repair Cost", , 7200, Y, True)
                    '        PrintAligned("Paid", , 8400, Y, True)
                    '    Case "SBR"
                    '        PrintAligned("Vendor", , 10, Y, True)
                    '        PrintAligned("Date", , 3200, Y, True)
                    '        PrintAligned("Repair Cost", , 4500, Y, True)
                    '        PrintAligned("Type", , 6000, Y, True)
                    '        PrintAligned("PartsOrderNo", , 7500, Y, True)
                    '        PrintAligned("Status", , 8800, Y, True)
                    '        PrintAligned("Service No", , 9500, Y, True)
            End Select

            'e.Graphics.DrawString(PrintText, New Font("Arial", 10), MyBrush, L, T)
            'e.Graphics.DrawString(PServiceNo, New Font("Arial", 10), MyBrush, 0, 65)
            'e.Graphics.DrawString(PDateOfClaim, New Font("Arial", 10), MyBrush, 100, 65)
            'e.Graphics.DrawString(PLast, New Font("Arial", 10), MyBrush, 200, 65)
            'e.Graphics.DrawString(PTele, New Font("Arial", 10), MyBrush, 300, 65)
            'e.Graphics.DrawString(PServiceNo, New Font("Arial", 10), MyBrush, 0, TopValue)
            'e.Graphics.DrawString(DateFormat(PDateOfClaim), New Font("Arial", 10), MyBrush, 100, TopValue)
            'e.Graphics.DrawString(Microsoft.VisualBasic.Left(PLast, 20), New Font("Arial", 10), MyBrush, 200, TopValue)
            'e.Graphics.DrawString(DressAni(CleanAni(PTele, 0)), New Font("Arial", 10), MyBrush, 300, TopValue)
            'e.Graphics.DrawString(ItemLine, New Font("Lucida Console", 10), MyBrush, 0, TopValue + 15)
            '--------
            Dim R As Integer = 65, i As Integer, j As Integer
            For i = 0 To SrnoArray.Count - 1
                e.Graphics.DrawString(SrnoArray(i), New Font("Arial", 8), MyBrush, 0, R)
                e.Graphics.DrawString(DateFormat(DateofClaimArray(i)), New Font("Arial", 8), MyBrush, 100, R)
                e.Graphics.DrawString(Microsoft.VisualBasic.Left(LastNameArray(i), 20), New Font("Arial", 8), MyBrush, 200, R)
                e.Graphics.DrawString(DressAni(CleanAni(TelephoneArray(i), 0)), New Font("Arial", 8), MyBrush, 300, R)

                e.Graphics.DrawString(ItemLineArray(i), New Font("Lucida Console", 8), MyBrush, 0, R + 15)
                'If i = 25 Then
                '    fraNavigate.Location = New Point(picPicture.Width / 2, R)
                'End If
                R = R + 30
            Next
            Dim Pages As Decimal
            Pages = i / 20
            Pages = Math.Ceiling(Pages)
            e.Graphics.DrawString("Page: " & 1, New Font("Arial", 8), MyBrush, 700, 10)
            frmPrintPreviewMain.Text = "Print Preview: " & ReportName & ", page " & CurrentPage + 1 & " of " & Pages

            For j = 0 To R Step 4
                e.Graphics.DrawString("'", New Font("Arial", 8), MyBrush, (picPicture.Width / 2) + 80, j)
            Next

            'fraNavigate.Location = New Point(800, 10)

            'picPicture.Image = b
            'b.Save("D:\pp.png", Imaging.ImageFormat.Png)
        End If
        'OutputObject.FontSize = 8
        ''PrintAligned("Time: " & Format(Now, "h:mm:ss am/pm"), , 10, 100)
        'PrintAligned("Time: " & Format(Now, "h:mm:ss tt"), , 10, 100)

        'If OutputToPrinter Then PageNumber = OutputObject.Page
        'PrintAligned("Page: " & PageNumber, , 10100, 100)

        'PrintCentered(StoreSettings.Name & "    " & StoreSettings.Address & "    " & StoreSettings.City, 500)

        'OutputObject.FontSize = 9
        'OutputObject.FontBold = True
        'Y = OutputObject.CurrentY
        'Select Case Mode
        '    Case "SCR"
        '        OutputObject.CurrentX = 0
        '        OutputObject.CurrentY = 700
        '        'PrintToTab(, "ServiceNo", 0)
        '        'PrintToTab(, "ServiceNo", 0,,, 700, 0, True)
        '        PrintToPosition2(, "ServiceNo", ,,, 700, 0)
        '        'PrintToTab(, "DateOfClaim", 20)
        '        'PrintToTab(, "DateOfClaim", 20,,, 700, 20, True)
        '        PrintToPosition2(, "DateOfClaim", ,,, 700, 1500)
        '        'PrintToTab(, "Last", 40)
        '        'PrintToTab(, "Last", 40,,, 700, 40, True)
        '        PrintToPosition2(, "Last", ,,, 700, 3000)
        '        'PrintToTab(, "Telephone", 60)
        '        'PrintToTab(, "Telephone", 60,,, 700, 60, True)
        '        PrintToPosition2(, "Telephone", ,,, 700, 4800)
        '        OutputObject.FontBold = False
        '    Case "SPR"
        '        PrintAligned("PartsOrderNo", , 10, Y, True)
        '        PrintAligned("Status", , 1300, Y, True)
        '        PrintAligned("ServiceNo", , 2000, Y, True)
        '        PrintAligned("Vendor", , 3000, Y, True)
        '        PrintAligned("DateOfClaim", , 5900, Y, True)
        '        PrintAligned("Repair Cost", , 7200, Y, True)
        '        PrintAligned("Paid", , 8400, Y, True)
        '    Case "SBR"
        '        PrintAligned("Vendor", , 10, Y, True)
        '        PrintAligned("Date", , 3200, Y, True)
        '        PrintAligned("Repair Cost", , 4500, Y, True)
        '        PrintAligned("Type", , 6000, Y, True)
        '        PrintAligned("PartsOrderNo", , 7500, Y, True)
        '        PrintAligned("Status", , 8800, Y, True)
        '        PrintAligned("Service No", , 9500, Y, True)
        'End Select
        'OutputObject.FontBold = False

        'picPicture.Image = b
        'b.Save("D:\pp.jpg", Imaging.ImageFormat.Jpeg)
    End Sub

    Dim ServiceNo As String
    Dim Last As String
    Dim Tele As String
    Dim DateOfClaim As String
    Dim Tv As Integer
    Dim Il As String

    Public Property PServiceNo As String
        Get
            PServiceNo = ServiceNo
            Return PServiceNo
        End Get
        Set(value As String)
            ServiceNo = value
        End Set
    End Property

    Public Property PLast As String
        Get
            PLast = Last
            Return PLast
        End Get
        Set(value As String)
            Last = value
        End Set
    End Property

    Public Property PTele As String
        Get
            PTele = Tele
            Return PTele
        End Get
        Set(value As String)
            Tele = value
        End Set
    End Property

    Public Property PDateOfClaim As String
        Get
            PDateOfClaim = DateOfClaim
            Return PDateOfClaim
        End Get
        Set(value As String)
            DateOfClaim = value
        End Set
    End Property

    Public Property TopValue As Integer
        Get
            TopValue = Tv
            Return TopValue
        End Get
        Set(value As Integer)
            Tv = value
        End Set
    End Property

    Public Property ItemLine As String
        Get
            ItemLine = Il
            Return Il
        End Get
        Set(value As String)
            Il = value
        End Set
    End Property

End Class